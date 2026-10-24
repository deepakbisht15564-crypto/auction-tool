import streamlit as st
import os
import re
import json
import zipfile
import io
import time
import fitz
import pandas as pd
from PIL import Image
from google import genai
from google.genai import types

st.set_page_config(
    page_title="Bulk E-Auction Notice Processor",
    layout="wide"
)
st.title("📄 Bulk E-Auction PDF Processor")
st.write("Clean Box Cutting + 9-Column Multi-Sheet Excel.")

api_key = st.sidebar.text_input(
    "Enter Gemini API Key",
    type="password"
)

uploaded_files = st.file_uploader(
    "Upload Auction PDFs (Bulk)",
    type=["pdf"],
    accept_multiple_files=True
)
start_btn = st.button("Start Processing")

def clean_val(text):
    if not text:
        return ""
    c = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', str(text))
    return " ".join(c.split()).strip()

def clean_prop(text):
    if not text:
        return ""
    c = clean_val(text)
    return re.sub(r'[^a-zA-Z0-9\s.,]', ' ', c).strip()

def get_notice_box(page):
    anchors = [
        "sale notice", "e-auction", "sarfaesi",
        "arcil", "reserve price", "authorised officer"
    ]
    hits = []
    for a in anchors:
        r = page.search_for(a)
        if r:
            hits.extend(r)
    if hits:
        min_x = min(h.x0 for h in hits)
        min_y = min(h.y0 for h in hits)
        max_x = max(h.x1 for h in hits)
        max_y = max(h.y1 for h in hits)
        return fitz.Rect(
            max(0, min_x - 15),
            max(0, min_y - 25),
            min(page.rect.width, max_x + 15),
            min(page.rect.height, max_y + 35)
        )
    return fitz.Rect(
        page.rect.width * 0.30,
        page.rect.height * 0.45,
        page.rect.width,
        page.rect.height
    )

def extract_with_ai(client, page, box):
    pix = page.get_pixmap(clip=box, dpi=200)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    
    prompt = """
    Extract all Bank / ARC E-Auction Sale Notices from this image.
    Return ONLY a JSON list of objects with these exact keys:
    - "bank_or_arc_name": Exact Bank or ARC name (e.g. Arcil, SBI, etc.)
    - "account_prospect_no": Account / Loan / NPA number
    - "borrower_name": ONLY Primary Borrower Name. Exclude Co-Borrowers/Guarantors.
    - "property_description": Property details & schedule
    - "reserve_price": Reserve Price with currency
    - "emd_amount": EMD Amount with currency
    - "inspection_date_time": Inspection Date & Time
    - "emd_last_date_time": EMD Last Date & Time
    - "auction_date_time": Auction Date & Time
    """
    
    for attempt in range(3):
        try:
            resp = client.models.generate_content(
                model="gemini-3.8-flash",
                contents=[prompt, img],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.0
                )
            )
            raw = resp.text.strip()
            if raw.startswith("```json"):
                raw = raw[7:-3].strip()
            elif raw.startswith("```"):
                raw = raw[3:-3].strip()
            data = json.loads(raw)
            if isinstance(data, dict) and "notices" in data:
                return data["notices"]
            elif isinstance(data, list):
                return data
            return [data]
        except Exception:
            time.sleep(2)
    return []

def process_pdf(file_bytes, filename, client):
    base = os.path.splitext(filename)[0]
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    
    t_idx = 0
    for idx, page in enumerate(doc):
        txt = page.get_text().lower()
        if any(w in txt for w in ["sale", "auction", "sarfaesi", "arcil"]):
            t_idx = idx
            break
            
    page = doc[t_idx]
    box = get_notice_box(page)
    
    # 1. High-resolution Vector Cutting PDF
    cut_doc = fitz.open()
    cut_doc.insert_pdf(doc, from_page=t_idx, to_page=t_idx)
    cut_doc[0].set_cropbox(box)
    pdf_bytes = cut_doc.write()
    cut_doc.close()
    
    # 2. Vision Extraction
    records = extract_with_ai(client, page, box)
    doc.close()
    
    # 3. Dedicated Excel Generation (Bank Tabs + 9 Columns)
    bank_groups = {}
    master_recs = []
    
    for idx, r in enumerate(records, start=1):
        bank = clean_val(r.get("bank_or_arc_name", "OTHERS")).upper()
        sheet = re.sub(r'[\\/*?:\[\]]', '', bank)[:30].strip() or "Notice Data"
        
        row_dict = {
            "S.No.": idx,
            "Account / Prospect No.": clean_val(r.get("account_prospect_no", "")),
            "Borrower(s) Name": clean_val(r.get("borrower_name", "")),
            "Property Description": clean_prop(r.get("property_description", "")),
            "Reserve Price": clean_val(r.get("reserve_price", "")),
            "EMD Amount": clean_val(r.get("emd_amount", "")),
            "Inspection Date & Time": clean_val(r.get("inspection_date_time", "")),
            "EMD Last Date & Time": clean_val(r.get("emd_last_date_time", "")),
            "Auction Date & Time": clean_val(r.get("auction_date_time", ""))
        }
        bank_groups.setdefault(sheet, []).append(row_dict)
        
        m_row = dict(row_dict)
        m_row["Source_File"] = filename
        m_row["Bank_or_ARC"] = bank
        master_recs.append(m_row)
        
    excel_bytes = None
    if bank_groups:
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as writer:
            for s_name, rows in bank_groups.items():
                pd.DataFrame(rows).to_excel(writer, sheet_name=s_name, index=False)
        excel_bytes = out.getvalue()
        
    return base, pdf_bytes, excel_bytes, master_recs

if start_btn:
    if not api_key:
        st.error("Please enter Gemini API Key in the sidebar!")
    elif not uploaded_files:
        st.warning("Please upload auction PDF files!")
    else:
        client = genai.Client(api_key=api_key)
        prog = st.progress(0)
        status = st.empty()
        
        zip_buf = io.BytesIO()
        total = len(uploaded_files)
        all_master = []
        
        with zipfile.ZipFile(zip_buf, "w") as zf:
            for i, f in enumerate(uploaded_files):
                status.text(f"Processing ({i+1}/{total}): {f.name}")
                base, c_pdf, x_data, m_rows = process_pdf(f.getvalue(), f.name, client)
                
                if c_pdf:
                    zf.writestr(f"{base}/{base}_Cutting.pdf", c_pdf)
                if x_data:
                    zf.writestr(f"{base}/{base}_Data.xlsx", x_data)
                if m_rows:
                    all_master.extend(m_rows)
                    
                prog.progress((i + 1) / total)
                
            if all_master:
                m_df = pd.DataFrame(all_master)
                zf.writestr(
                    "CRM_All_Accounts_Master.csv",
                    m_df.to_csv(index=False).encode('utf-8')
                )
                
        status.success("All Done!")
        st.download_button(
            label="📥 Download All Deliverables (ZIP)",
            data=zip_buf.getvalue(),
            file_name="Processed_Auctions.zip",
            mime="application/zip"
        )
        
