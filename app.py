import streamlit as st
import os
import re
import json
import zipfile
import io
import base64
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
st.write("ExtractFox-Grade Vision Engine: Multi-Row Table Extraction.")

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

def get_optimized_image_b64(page):
    """
    Renders notice at optimal resolution (< 1.5MB) to avoid 413 Payload Too Large error
    """
    pix_w, pix_h = page.rect.width, page.rect.height
    clip_box = fitz.Rect(0, pix_h * 0.20, pix_w, pix_h * 0.98)
    
    # 130 DPI keeps image crisp while payload stays well below Groq 4MB limit
    pix = page.get_pixmap(dpi=130, clip=clip_box)
    img_bytes = pix.tobytes("jpeg")
    return base64.b64encode(img_bytes).decode("utf-8")

def extract_notices_with_groq_vision(groq_client, page):
    b64_img = get_optimized_image_b64(page)

    prompt = """
You are an expert legal data extraction engine. Analyze this newspaper auction notice table image.

CRITICAL INSTRUCTIONS:
1. Locate the multi-row table of auction properties/accounts (e.g. Arcil, Cholamandalam, etc.).
2. EXTRACT EVERY ROW SEPARATELY. Do not skip any borrower or property.
3. For EVERY account row, output a JSON object:
   - bank_name: Name of Bank or ARC (e.g. Arcil, Cholamandalam).
   - borrower_name: Primary borrower name in Title Case (no Mr/Mrs/Shri). Strip co-borrowers.
   - property_description: Complete schedule of property (Survey No, Natham, Village, Area, Boundaries North/South/East/West).
   - reserve_price: Pure numbers only (e.g. 1151688). No Rs symbol.
   - emd_amount: Pure numbers only (e.g. 115169). No year numbers like 2026.
   - inspection_date_time: Inspection date/time or 'As per notice terms'.
   - emd_last_date_time: Last date & time for EMD.
   - auction_date_time: Date and time of auction.

Return ONLY a pure valid JSON array of objects.
"""
    try:
        res = groq_client.chat.completions.create(
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/jpeg;base64,{b64_img}"
                            }
                        }
                    ]
                }
            ],
            model="llama-3.2-11b-vision-preview",
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
    except Exception as e:
        st.error(f"Vision API Warning: {str(e)}")
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
    target_page = doc[0]

    full_rect = target_page.rect
    target_page.set_cropbox(full_rect)
    target_page.set_mediabox(full_rect)

    full_doc = fitz.open()
    full_doc.insert_pdf(doc, from_page=0, to_page=0)
    full_doc[0].set_cropbox(full_rect)
    full_doc[0].set_mediabox(full_rect)
    pdf_bytes = full_doc.write()
    full_doc.close()

    extracted = extract_notices_with_groq_vision(groq_client, target_page)
    doc.close()

    bank_sheets = {}
    master_records = []

    for item in extracted:
        raw_bank = clean_val(item.get("bank_name", "ARCIL")).upper()
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

    # Bulletproof fallback so bank_sheets is never empty
    if not bank_sheets:
        clean_fn = filename.replace("_Cutting", "").replace(".pdf", "").replace("-", " ")
        tokens = [w for w in clean_fn.split() if not any(c.isdigit() for c in w)]
        filter_words = ["THE", "NEW", "INDIAN", "EXPRESS", "TIMES", "DAINIK", "NOTICE", "E-AUCTION", "SALE", "AUCTION"]
