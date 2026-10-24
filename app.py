import streamlit as st
import os
import re
import zipfile
import io
import fitz  # PyMuPDF
import pandas as pd

st.set_page_config(page_title="Bulk E-Auction Notice Processor", layout="wide")
st.title("📄 Bulk E-Auction PDF Processor")
st.write("Vector Quality Box Cuttings + Complete Field-Wise Excel Extraction.")

def clean_val(text):
    if not text:
        return ""
    cleaned = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', str(text))
    return re.sub(r'[\r\n\t]+', ' ', cleaned).strip()

def get_precise_notice_box(page):
    keywords = ["sale notice", "e-auction", "sarfaesi", "arcil", "reserve price", "earnest money", "description of property"]
    rects = []
    for kw in keywords:
        matches = page.search_for(kw)
        if matches:
            rects.extend(matches)
            
    if rects:
        min_x = min([r.x0 for r in rects])
        min_y = min([r.y0 for r in rects])
        max_x = max([r.x1 for r in rects])
        max_y = max([r.y1 for r in rects])
        
        return fitz.Rect(
            max(0, min_x - 30),
            max(0, min_y - 45),
            min(page.rect.width, max_x + 30),
            min(page.rect.height, max_y + 45)
        )
    return fitz.Rect(page.rect.width * 0.30, page.rect.height * 0.45, page.rect.width, page.rect.height)

def parse_notice_data(full_text, filename):
    lines = [clean_val(l) for l in full_text.split('\n') if clean_val(l)]
    
    # 1. Exact Bank / ARC Name Detection
    bank_name = ""
    for line in lines[:25]:
        upper_line = line.upper()
        if "ARCIL" in upper_line or "ASSET RECONSTRUCTION" in upper_line:
            bank_name = "ARCIL"
            break
        elif "STATE BANK" in upper_line or "SBI" in upper_line:
            bank_name = "STATE BANK OF INDIA"
            break
        elif "HDFC" in upper_line:
            bank_name = "HDFC BANK"
            break
        elif "ICICI" in upper_line:
            bank_name = "ICICI BANK"
            break
        elif "CANARA" in upper_line:
            bank_name = "CANARA BANK"
            break
        elif "PUNJAB NATIONAL" in upper_line or "PNB" in upper_line:
            bank_name = "PUNJAB NATIONAL BANK"
            break
        elif "BANK OF BARODA" in upper_line or "BOB" in upper_line:
            bank_name = "BANK OF BARODA"
            break
        elif "INDIAN BANK" in upper_line:
            bank_name = "INDIAN BANK"
            break
        elif "AXIS" in upper_line:
            bank_name = "AXIS BANK"
            break
        elif "KOTAK" in upper_line:
            bank_name = "KOTAK MAHINDRA BANK"
            break
        elif "BANK" in upper_line:
            bank_name = line.split("BRANCH")[0].split("NOTICE")[0].strip()
            if len(bank_name) > 3:
                break
    
    if not bank_name:
        bank_name = "ARCIL" if "arcil" in full_text.lower() else "BANK NOTICE"
    bank_name = bank_name.upper()

    # 2. Borrower Name (Strictly only the primary borrower)
    borrower_name = ""
    name_match = re.search(r'(?:MrMrs\.|Mr\.|Mrs\.)?([A-Z\s]+?)(?=\s+THE\s+NEW|\s+INDIAN|\.pdf)', filename, re.IGNORECASE
