import streamlit as st
import os
import re
import zipfile
import io
import fitz  # PyMuPDF
import pandas as pd

st.set_page_config(page_title="Bulk E-Auction Notice Processor", layout="wide")
st.title("📄 Bulk E-Auction PDF Processor")
st.write("100% Offline / No-API Engine — Clean Box Cuttings + Field-Wise Excel Extraction.")

def clean_val(text: str) -> str:
    if not text:
        return ""
    cleaned = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', str(text))
    return re.sub(r'[\r\n\t]+', ' ', cleaned).strip()

def get_notice_crop_box(page):
    keywords = ["arcil", "asset reconstruction", "e-auction", "sale notice", "sarfaesi", "reserve price", "emd"]
    found_rects = []
    
    for kw in keywords:
        rects = page.search_for(kw)
        if rects:
            found_rects.extend(rects)

    if found_rects:
        bbox = found_rects[0]
        for r in found_rects[1:]:
            bbox |= r
        
        # Exact auction notice box crop (cuts out left side classifieds/matrimonial)
        return fitz.Rect(
            max(0, bbox.x0 - 15),
            max(0, bbox.y0 - 20),
            min(page.rect.width, bbox.x1 + 15),
            min(page.rect.height, bbox.y1 + 40)
        )
    else:
        # Fallback to bottom-right quadrant
        return fitz.Rect(page.rect.width * 0.38, page.rect.height * 0.52, page.rect.width, page.rect.height)

def parse_auction_fields(full_text, filename):
    text_lower = full_text.lower()
    
    # 1. Bank / ARC Name
    bank_name = "OTHERS"
    if "arcil" in text_lower or "asset reconstruction" in text_lower:
        bank_name = "ARCIL"
    elif "state bank" in text_lower or "sbi" in text_lower:
        bank_name = "STATE BANK OF INDIA"
    elif "hdfc" in text_lower:
        bank_name = "HDFC BANK"
    elif "icici" in text_lower:
        bank_name = "ICICI BANK"
    elif "canara" in text_lower:
        bank_name = "CANARA BANK"
    elif "punjab national" in text_lower or "pnb" in text_lower:
        bank_name = "PNB"

    # 2. Borrower Name (Strictly only the primary borrower)
    borrower_name = ""
    b_match = re.search(r'(?:borrower|name\s*of\s*borrower)s?[:\s\-]+([A-Za-z\.\s]+?)(?=\s*(?:co-borrower|guarantor|legal|mortgagor|\/|\n|rs\.|₹|property))', full_text, re.IGNORECASE)
    if b_match:
        borrower_name = clean_val(b_match.group(1))
    else:
        # Fallback extraction from filename if notice has table format
        name_part = re.search(r'(?:MrMrs\.|Mr\.|Mrs\.)?([A-Z\s]+?)(?=\s+THE\s+NEW|\s+INDIAN|\.pdf)', filename, re.IGNORECASE)
        borrower_name = clean_val(name_part.group(1)) if name_part else ""

    # 3. Reserve Price
    rp_match = re.search(r'(?:reserve\s*price|r\.?p\.?)[^\d₹Rs]*([₹Rs\.]*\s*[\d,]+(?:\.\d+)?)', full_text, re.IGNORECASE)
    reserve_price = clean_val(rp_match.group(1)) if rp_match else ""

    # 4. Dates & Times
    dates = re.findall(r'(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4})', full_text)
    times = re.findall(r'(\d{1,2}[:\.]\d{2}\s*(?:am|pm|a\.m\.|p\.m\.|hours|hrs)?)', full_text, re.IGNORECASE)

    emd_submit_date = dates[0] if len(dates) > 0 else ""
    auction_date = dates[1] if len(dates) > 1 else (dates[0] if dates else "")
    auction_end_date = auction_date

    emd_submit_time = clean_val(times[0]) if len(times) > 0 else ""
    auction_start_time = clean_val(times[1]) if len(times) > 1 else ""
    auction_end_time = clean_val(times[2]) if len(times) > 2 else ""

    # 5. Property Description
    prop_match = re.search(r'(?:description\s*of\s*(?:the\s*)?property|schedule\s*of\s*property|property\s*details)[^\n]*\n([\s\S]{50,400})', full_text, re.IGNORECASE)
    prop_desc = clean_val(prop_match.group(1)) if prop_match else clean_val(full_text[:300])

    row = {
        "Notice File": filename,
        "Bank / ARC": bank_name,
        "Borrower Name": borrower_name,
        "Property Description": prop_desc,
        "Reserve Price": reserve_price,
        "EMD Submit Date": emd_submit_date,
        "EMD Submit Time": emd_submit_time,
        "Auction Date": auction_date,
        "Auction Start Time": auction_start_time,
        "Auction End Date": auction_end_date,
        "Auction End Time": auction_end_time
    }
    return bank_name, [row]

def process_single_pdf(file_bytes, filename):
    base_name = os.path.splitext(filename)[0]
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    
    # 1. Exact Cutting
    cut_doc = fitz.open()
    full_text = ""

    for page_num in range(len(doc)):
        page = doc[page_num]
        txt = page.get_text()
        full_text += txt + "\n"
        
        if any(kw in txt.lower() for kw in ["auction", "sale notice", "sarfaesi", "arcil", "reserve price"]):
            crop_box = get_notice_crop_box(page)
            temp_doc = fitz.open()
            temp_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
            p = temp_doc[0]
            p.set_cropbox(crop_box)
            cut_doc.insert_pdf(temp_doc)
            temp_doc.close()

    if len(cut_doc) == 0 and len(doc) > 0:
        crop_box = get
