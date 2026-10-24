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
st.title("Bulk E-Auction PDF Processor")
st.write("Full Original Page PDF + Bank-Named Multi-Sheet Excel.")

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

def extract_with_ai(client, pil_image):
    prompt = """
    Look at this newspaper page and extract all details from the Bank / ARC E-Auction / SARFAESI Sale Notice.
    Return ONLY a strict JSON list of objects with these exact keys:
    - "bank_or_arc_name": Exact Bank or ARC name (e.g. Arcil, State Bank of India, Indian Bank, Canara Bank, etc.)
    - "account_prospect_no": Account / Loan / NPA number
    - "borrower_name": ONLY Primary Borrower Name. Strictly exclude Co-Borrowers, Guarantors, and Mortgagors.
    - "property_description": Property details & schedule
    - "reserve_price": Reserve Price with currency (e.g. Rs. 25,00,000)
    - "emd_amount": EMD Amount with currency
    - "inspection_date_time": Inspection Date & Time
    - "emd_last_date_time": EMD Last Date & Time
    - "auction_date_time": Auction Date & Time
    """

    for attempt in range(3):
        try:
            resp = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=[prompt, pil_image],
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

    # Find the exact page containing sale notice
    t_idx = 0
    for idx, page in enumerate(doc):
        txt = page.get_text().lower()
        if any(w in txt for w in ["sale notice", "e-auction", "sarfaesi", "arcil", "reserve price"]):
            t_idx = idx
            break

    # 1. Pura Original HD Page (Zero Crop, No Cutting, Zero Pixel Loss)
    full_doc = fitz.open()
    full_doc.insert_pdf(doc, from_page=t_idx, to_page=t_idx)
    pdf_bytes = full_doc.write()
    full_doc.close()

    # 2. Render Page for AI Vision
    target_page = doc[t_idx]
    pix = target_page.get_pixmap(dpi=200)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    doc.close()

    # 3. AI Vision Extraction
    records = extract_with_ai(client, img)

    # 4. Excel Generation: Sheet Rename to Bank Name
    bank_groups = {}
    master_recs = []

    for idx, r in enumerate(records, start=1):
        raw_bank = clean_val(r.get("bank_or_arc_name", "BANK NOTICE")).upper()
        # Excel sheet name limit 30 chars and no special chars
        clean_sheet = re.sub(r'[\\/*?:\[\]]', '', raw_bank)[:30].strip() or "NOTICE DATA"

        row_dict = {
            "S.No.": idx,
            "Bank / ARC": raw_bank,
            "Account / Prospect No.": clean_val(r.get("account_prospect_no", "")),
            "Borrower(s) Name": clean_val(r.get("borrower_name", "")),
            "Property Description": clean_prop(r.get("property_description", "")),
            "Reserve Price": clean_val(r.get("reserve_price", "")),
            "EMD Amount": clean_val(r.get("emd_amount", "")),
            "Inspection Date & Time": clean_val(r.get("inspection_date_time", "")),
            "EMD Last Date & Time": clean_val(r.get("emd_last_date_time", "")),
            "Auction Date & Time": clean_val(r.get("auction_date_time", ""))
        }
        bank_groups.setdefault(clean_sheet, []).append(row_dict)

        m_row = dict(row_dict)
        m_row["Source_File"] = filename
        master_recs.append(m_row)

    excel_bytes = None
    if bank_groups:
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as writer:
            for sheet_name, rows in bank_groups.items():
                pd.DataFrame(rows).to_excel(writer, sheet_name=sheet_name, index=False)
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
                base, p_pdf, x_data, m_rows = process_pdf(f.getvalue(), f.name, client)

                if p_pdf:
                    zf.writestr(f"{base}/{base}_FullPage.pdf", p_pdf)
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
            label="Download All Deliverables (ZIP)",
            data=zip_buf.getvalue(),
            file_name="Processed_Auctions.zip",
            mime="application/zip"
        )
