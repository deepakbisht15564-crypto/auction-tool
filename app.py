import streamlit as st
import os
import re
import json
import zipfile
import io
import fitz
import pandas as pd
from groq import Groq
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Bulk E-Auction Notice Processor",
    layout="wide"
)
st.title("Bulk E-Auction PDF Processor")
st.write("ExtractFox-Grade Engine: Multi-Bank Isolation + Clean Corporate Excel.")

api_key = st.sidebar.text_input(
    "Enter Groq API Key (starts with gsk_)",
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
    return re.sub(r'[^a-zA-Z0-9\s.,\-\/\(\)&:;]', ' ', c).strip()

def format_borrower_name(raw_name):
    if not raw_name or raw_name.upper() in ["N/A", "UNKNOWN", "NONE"]:
        return "N/A"
    prefixes = [
        r'\bmr[.]?\s*mrs[.]?', r'\bmrmrs[.]?', r'\bmr[.]?', r'\bmrs[.]?',
        r'\bms[.]?', r'\bshri[.]?', r'\bsmt[.]?', r'\bm/s[.]?', r'\bdr[.]?'
    ]
    cleaned = raw_name
    for p in prefixes:
        cleaned = re.sub(p, '', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'[^a-zA-Z\s.]', ' ', cleaned)
    words = cleaned.split()
    formatted = []
    for w in words:
        if len(w) == 1:
            formatted.append(w.upper())
        else:
            formatted.append(w.capitalize())
    return " ".join(formatted).strip() or "N/A"

def clean_amount(val_str):
    if not val_str:
        return ""
    cleaned = re.sub(r'(?i)(rs\.?|inr|₹|\/-)', '', str(val_str))
    cleaned = re.sub(r'[^\d,.]', '', cleaned).strip('., ')
    if cleaned in ["2023", "2024", "2025", "2026", "2027", "2028"]:
        return ""
    return cleaned

def extract_notices_with_groq(groq_client, page_text):
    prompt = f"""
You are an expert financial auditor extracting Bank & ARC E-Auction Sale Notices from a newspaper broadsheet.

CRITICAL RULES:
1. Ignore all general news, politics, world news, advertisements, matrimony, and classifieds.
2. A single newspaper page may contain notices from MULTIPLE DIFFERENT BANKS/ARCs (e.g. Arcil, ICICI Bank, SBI, PNB, Canara Bank, Axis Bank, HDFC, etc.).
3. Identify every distinct Bank/ARC notice on this page.
4. For each notice, extract all properties/accounts into a list of JSON objects.
5. Strictly map each account to its own correct Bank/ARC name. DO NOT mix data between banks.
6. Borrower Name: Primary Borrower ONLY in Title Case. Strip prefixes (Mr, Mrs, Shri, Smt). Exclude Co-Borrowers/Guarantors.
7. Property Description: Full property schedule as-it-is.
8. Reserve Price & EMD: Pure numerical amounts without 'Rs.' or currency symbols. DO NOT capture years like 2026 as amounts.

Return ONLY a pure valid JSON array of objects with these exact keys:
[
  {{
    "bank_or_arc_name": "Exact Bank or ARC Name",
    "borrower_name": "Primary Borrower Name",
    "property_description": "Full property description",
    "reserve_price": "Numeric Amount",
    "emd_amount": "Numeric Amount",
    "inspection_date_time": "Date & Time or As per notice terms",
    "emd_last_date_time": "EMD submission deadline",
    "auction_date_time": "Auction Date and Time"
  }}
]

Text from broadsheet page:
{page_text[:35000]}
"""
    models_to_try = ["openai/gpt-oss-20b", "openai/gpt-oss-120b"]
    for model_name in models_to_try:
        try:
            res = groq_client.chat.completions.create(
                messages=[
                    {
                        "role": "system",
                        "content": "You return ONLY a pure JSON array of objects with zero explanations or markdown tags."
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                model=model_name,
                temperature=0.0
            )
            raw = res.choices[0].message.content.strip()
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
            continue
    return []

def style_excel_worksheet(ws):
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    cell_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    ws.row_dimensions[1].height = 32
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = cell_border

    for row_idx in range(2, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 65
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = Font(name="Calibri", size=10)
            cell.alignment = center_align
            cell.border = cell_border

    col_widths = {
        "A": 9.0,   # S.No.
        "B": 24.0,  # Bank Name
        "C": 30.0,  # Borrower(s) Name
        "D": 55.0,  # Property Description
        "E": 18.0,  # Reserve Price
        "F": 18.0,  # EMD Amount
        "G": 22.0,  # Inspection Date & Time
        "H": 22.0,  # EMD Last Date & Time
        "I": 22.0   # Auction Date & Time
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

def process_single_pdf(file_bytes, filename, groq_client):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    t_idx = 0
    for idx in range(len(doc)):
        txt = doc[idx].get_text().lower()
        if any(w in txt for w in ["sale notice", "e-auction", "sarfaesi", "arcil", "reserve price"]):
            t_idx = idx
            break

    target_page = doc[t_idx]

    full_rect = target_page.rect
    target_page.set_cropbox(full_rect)
    target_page.set_mediabox(full_rect)

    full_doc = fitz.open()
    full_doc.insert_pdf(doc, from_page=t_idx, to_page=t_idx)
    full_doc[0].set_cropbox(full_rect)
    full_doc[0].set_mediabox(full_rect)
    pdf_bytes = full_doc.write()
    full_doc.close()

    page_text = target_page.get_text()
    doc.close()

    extracted = extract_notices_with_groq(groq_client, page_text)

    bank_sheets = {}
    master_records = []

    for item in extracted:
        raw_bank = clean_val(item.get("bank_or_arc_name", "BANK NOTICE")).upper()
        clean_sheet = re.sub(r'[\\/*?:\[\]]', '', raw_bank)[:25].strip() or "NOTICE DATA"

        row_dict = {
            "S.No.": len(bank_sheets.get(clean_sheet, [])) + 1,
            "Bank Name": raw_bank,
            "Borrower(s) Name": format_borrower_name(item.get("borrower_name", "")),
            "Property Description": clean_prop(item.get("property_description", "")),
            "Reserve Price": clean_amount(item.get("reserve_price", "")),
            "EMD Amount": clean_amount(item.get("emd_amount", "")),
            "Inspection Date & Time": clean_val(item.get("inspection_date_time", "As per notice terms")),
            "EMD Last Date & Time": clean_val(item.get("emd_last_date_time", "")),
            "Auction Date & Time": clean_val(item.get("auction_date_time", ""))
        }

        bank_sheets.setdefault(clean_sheet, []).append(row_dict)

        m_row = dict(row_dict)
        m_row["Source_File"] = filename
        master_records.append(m_row)

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        for sheet_title, rows in bank_sheets.items():
            pd.DataFrame(rows).to_excel(writer, sheet_name=sheet_title, index=False)
            ws = writer.sheets[sheet_title]
            style_excel_worksheet(ws)

    excel_bytes = out.getvalue()
    return pdf_bytes, excel_bytes, master_records

if start_btn:
    if not api_key:
        st.error("Please enter your Groq API Key (starts with gsk_) in the sidebar!")
    elif not uploaded_files:
        st.warning("Please upload auction PDF files!")
    else:
        groq_client = Groq(api_key=api_key)
        prog = st.progress(0)
        status = st.empty()

        zip_buf = io.BytesIO()
        total = len(uploaded_files)
        all_master = []

        with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for i, f in enumerate(uploaded_files):
                status.text(f"Processing ({i+1}/{total}): {f.name}")
                base = os.path.splitext(f.name)[0]

                p_pdf, x_data, m_rows = process_single_pdf(f.getvalue(), f.name, groq_client)

                if p_pdf:
                    zf.writestr(f"{base}_Complete_Page.pdf", p_pdf)
                if x_data:
                    zf.writestr(f"{base}_Data.xlsx", x_data)
                if m_rows:
                    all_master.extend(m_rows)

                prog.progress((i + 1) / total)

            if all_master:
                m_df = pd.DataFrame(all_master)
                zf.writestr(
                    "CRM_All_Accounts_Master.csv",
                    m_df.to_csv(index=False).encode('utf-8')
                )

        zip_buf.seek(0)
        final_zip = zip_buf.getvalue()

        status.success("All Done!")
        st.download_button(
            label="Download All Deliverables (ZIP)",
            data=final_zip,
            file_name="Processed_Auctions.zip",
            mime="application/zip"
        )
