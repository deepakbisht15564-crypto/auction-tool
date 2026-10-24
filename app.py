import streamlit as st
import os
import re
import json
import zipfile
import io
import fitz  # PyMuPDF
import pandas as pd
import google.generativeai as genai
from PIL import Image

st.set_page_config(page_title="Bulk E-Auction Notice Processor", layout="wide")
st.title("📄 Bulk E-Auction PDF Processor")
st.write("AI Vision Extraction — Works 100% on Newspaper Tables & Scanned PDFs.")

# Sidebar API Key
api_key = st.sidebar.text_input("Enter Gemini API Key (starts with AIza)", type="password")

uploaded_files = st.file_uploader("Upload Auction PDFs (Bulk)", type=["pdf"], accept_multiple_files=True)
start_btn = st.button("Start Processing")

def clean_val(text):
    if not text:
        return ""
    c = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', str(text))
    return " ".join(c.split()).strip()

def extract_via_gemini(pil_image, filename):
    prompt = """
You are an expert E-Auction document analyzer. 
Look at this newspaper page and extract all details from the E-Auction / SARFAESI Sale Notice into a strict JSON list.

Extract these exact keys:
- "bank_or_arc_name": Exact Bank or ARC name printed on the notice (e.g. Arcil, SBI, Indian Bank, Canara Bank, etc.).
- "borrower_name": ONLY the primary Borrower's name. Strictly exclude Co-Borrowers, Guarantors, Mortgagors, or Partners.
- "property_description": Brief description and address of the property.
- "reserve_price": Reserve Price with currency (e.g. Rs. 25,00,000).
- "emd_amount": EMD Amount with currency.
- "emd_submit_date": Last date for EMD submission (DD/MM/YYYY).
- "emd_submit_time": Time for EMD submission (e.g. 05:00 PM).
- "auction_date": Date of Auction (DD/MM/YYYY).
- "auction_start_time": Auction start time (e.g. 11:00 AM).
- "auction_end_date": Auction end date (DD/MM/YYYY).
- "auction_end_time": Auction end time (e.g. 12:00 PM).

Return ONLY the raw JSON array. Example:
[{"bank_or_arc_name": "...", "borrower_name": "...", "property_description": "...", "reserve_price": "...", "emd_amount": "...", "emd_submit_date": "...", "emd_submit_time": "...", "auction_date": "...", "auction_start_time": "...", "auction_end_date": "...", "auction_end_time": "..."}]
"""
    model = genai.GenerativeModel("gemini-1.5-flash")
    response = model.generate_content([prompt, pil_image])
    text_out = response.text.strip()
    
    if text_out.startswith("```json"):
        text_out = text_out[7:-3].strip()
    elif text_out.startswith("```"):
        text_out = text_out[3:-3].strip()

    data = json.loads(text_out)
    if isinstance(data, dict):
        data = [data]
    return data

def process_single_pdf(file_bytes, filename):
    base_name = os.path.splitext(filename)[0]
    doc = fitz.open(stream=file_bytes, filetype="pdf")

    # Target the notice page
    target_idx = 0
    for idx, page in enumerate(doc):
        txt = page.get_text().lower()
        if any(w in txt for w in ["sale", "notice", "auction", "arcil", "sarfaesi"]):
            target_idx = idx
            break

    # 1. Full Page HD PDF Deliverable
    full_doc = fitz.open()
    full_doc.insert_pdf(doc, from_page=target_idx, to_page=target_idx)
    pdf_bytes = full_doc.write()
    full_doc.close()

    # 2. Render Page to High-Res Image for AI Vision
    target_page = doc[target_idx]
    pix = target_page.get_pixmap(dpi=200)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    doc.close()

    # 3. Extract exact details
    rows = []
    bank_name = "NOTICE"
    try:
        items = extract_via_gemini(img, filename)
        for item in items:
            bank_name = clean_val(item.get("bank_or_arc_name", "BANK"))
            r = {
                "Notice File": filename,
                "Bank / ARC": bank_name,
                "Borrower Name": clean_val(item.get("borrower_name", "")),
                "Property Description": clean_val(item.get("property_description", "")),
                "Reserve Price": clean_val(item.get("reserve_price", "")),
                "EMD Amount": clean_val(item.get("emd_amount", "")),
                "EMD Submit Date": clean_val(item.get("emd_submit_date", "")),
                "EMD Submit Time": clean_val(item.get("emd_submit_time", "")),
                "Auction Date": clean_val(item.get("auction_date", "")),
                "Auction Start Time": clean_val(item.get("auction_start_time", "")),
                "Auction End Date": clean_val(item.get("auction_end_date", "")),
                "Auction End Time": clean_val(item.get("auction_end_time", ""))
            }
            rows.append(r)
    except Exception as e:
        st.error(f"Error on {filename}: {str(e)}")

    # 4. Generate Excel
    excel_bytes = None
    if rows:
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as writer:
            sheet_name = re.sub(r'[\\/*?:\[\]]', '', bank_name)[:30] or "Sheet1"
            pd.DataFrame(rows).to_excel(writer, sheet_name=sheet_name, index=False)
        excel_bytes = out.getvalue()

    return base_name, pdf_bytes, excel_bytes, rows

if start_btn:
    if not api_key:
        st.error("Please enter your Gemini API Key in the sidebar!")
    elif not uploaded_files:
        st.warning("Pehle koi PDF upload kijiye!")
    else:
        genai.configure(api_key=api_key)
        progress_bar = st.progress(0)
        status_text = st.empty()

        zip_buffer = io.BytesIO()
        total = len(uploaded_files)
        master_rows = []

        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for idx, file in enumerate(uploaded_files):
                status_text.text(f"Processing ({idx+1}/{total}): {file.name}")
                base, notice_pdf, excel_data, rows = process_single_pdf(file.getvalue(), file.name)

                if notice_pdf:
                    zip_file.writestr(f"{base}/{base}_NoticePage.pdf", notice_pdf)
                if excel_data:
                    zip_file.writestr(f"{base}/{base}_Data.xlsx", excel_data)
                if rows:
                    master_rows.extend(rows)

                progress_bar.progress((idx + 1) / total)

            if master_rows:
                master_df = pd.DataFrame(master_rows)
                zip_file.writestr("CRM_All_Accounts_Master.csv", master_df.to_csv(index=False).encode('utf-8'))

        status_text.success("All Done!")
        st.download_button(
            label="📥 Download All Deliverables (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Processed_Auctions.zip",
            mime="application/zip"
        )
