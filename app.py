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
st.write("AI Vision Extraction + Auto-Model Fallback.")

api_key = st.sidebar.text_input("Enter Gemini API Key (starts with AIza)", type="password")

uploaded_files = st.file_uploader("Upload Auction PDFs (Bulk)", type=["pdf"], accept_multiple_files=True)
start_btn = st.button("Start Processing")

def clean_val(text):
    if not text:
        return ""
    c = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', str(text))
    return " ".join(c.split()).strip()

def get_best_available_model():
    # Priority order of supported vision models
    preferred = [
        "gemini-1.5-flash-latest",
        "gemini-1.5-flash",
        "gemini-1.5-pro",
        "gemini-1.5-pro-latest",
        "gemini-pro-vision"
    ]
    try:
        available_models = [
            m.name.replace("models/", "") 
            for m in genai.list_models() 
            if 'generateContent' in m.supported_generation_methods
        ]
        for pref in preferred:
            if pref in available_models:
                return pref
        if available_models:
            return available_models[0]
    except Exception:
        pass
    return "gemini-1.5-flash-latest"

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
    model_name = get_best_available_model()
    model = genai.GenerativeModel(model_name)
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

def local_fallback_extraction(full_text, filename):
    lines = [clean_val(l) for l in full_text.split("\n") if clean_val(l)]
    bank_name = "ARCIL" if "arcil" in full_text.lower() else "BANK NOTICE"
    for line in lines[:25]:
        u = line.upper()
        if any(b in u for b in ["ARCIL", "SBI", "CANARA", "PNB", "BANK OF BARODA", "INDIAN BANK", "HDFC", "ICICI", "AXIS", "KOTAK"]):
            bank_name = line.strip().upper()
            break

    # Borrower from filename
    clean_fn = filename.replace("_Cutting", "").replace(".pdf", "")
    tokens = [w for w in clean_fn.split() if not any(c.isdigit() for c in w)]
    filter_words = ["THE", "NEW", "INDIAN", "EXPRESS", "TIMES", "DAINIK", "NOTICE", "E-AUCTION", "SALE"]
    clean_tokens = [w for w in tokens if w.upper() not in filter_words]
    borrower_name = " ".join(clean_tokens).replace("MrMrs.", "").replace("Mr.", "").replace("Mrs.", "").strip()

    amounts = re.findall(r'(?:Rs\.?|₹|INR)\s*([\d,]+(?:\.\d+)?)', full_text, re.I)
    reserve_price = f"Rs. {clean_val(amounts[0])}" if len(amounts) >= 1 else ""
    emd_amount = f"Rs. {clean_val(amounts[1])}" if len(amounts) >= 2 else ""

    dates = re.findall(r'\b\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4}\b', full_text)
    times = re.findall(r'\b\d{1,2}[:\.]\d{2}\s*(?:AM|PM|a\.m\.|p\.m\.|hrs)?\b', full_text, re.I)

    return [{
        "Notice File": filename,
        "Bank / ARC": bank_name,
        "Borrower Name": borrower_name,
        "Property Description": clean_val(full_text[:300]),
        "Reserve Price": reserve_price,
        "EMD Amount": emd_amount,
        "EMD Submit Date": dates[0] if len(dates) > 0 else "",
        "EMD Submit Time": clean_val(times[0]) if len(times) > 0 else "05:00 PM",
        "Auction Date": dates[1] if len(dates) > 1 else (dates[0] if dates else ""),
        "Auction Start Time": clean_val(times[1]) if len(times) > 1 else "11:00 AM",
        "Auction End Date": dates[1] if len(dates) > 1 else (dates[0] if dates else ""),
        "Auction End Time": clean_val(times[2]) if len(times) > 2 else "12:00 PM"
    }]

def process_single_pdf(file_bytes, filename):
    base_name = os.path.splitext(filename)[0]
    doc = fitz.open(stream=file_bytes, filetype="pdf")

    target_idx = 0
    full_text = ""
    for idx, page in enumerate(doc):
        txt = page.get_text()
        full_text += txt + "\n"
        if any(w in txt.lower() for w in ["sale", "notice", "auction", "arcil", "sarfaesi"]):
            target_idx = idx
            break

    # Full Page HD PDF Deliverable
    full_doc = fitz.open()
    full_doc.insert_pdf(doc, from_page=target_idx, to_page=target_idx)
    pdf_bytes = full_doc.write()
    full_doc.close()

    # Image conversion for Gemini
    target_page = doc[target_idx]
    pix = target_page.get_pixmap(dpi=200)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    doc.close()

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
        # Fallback agar API kisi vajah se fail ho
        rows = local_fallback_extraction(full_text, filename)
        bank_name = rows[0]["Bank / ARC"]

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
