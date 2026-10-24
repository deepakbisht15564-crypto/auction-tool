import streamlit as st
import os
import re
import zipfile
import io
import fitz  # PyMuPDF
import pandas as pd

st.set_page_config(page_title="Bulk E-Auction Notice Processor", layout="wide")
st.title("📄 Bulk E-Auction PDF Processor")
st.write("Clean Box Cuttings + Field-Wise Excel Extraction.")

def clean_val(text):
    if not text:
        return ""
    cleaned = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', str(text))
    return re.sub(r'[\r\n\t]+', ' ', cleaned).strip()

def get_precise_notice_box(page):
    keywords = ["sale notice", "e-auction", "sarfaesi", "arcil", "reserve price", "earnest money", "description"]
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
        u = line.upper()
        if "ARCIL" in u or "ASSET RECONSTRUCTION" in u:
            bank_name = "ARCIL"
            break
        elif "STATE BANK" in u or "SBI" in u:
            bank_name = "STATE BANK OF INDIA"
            break
        elif "HDFC" in u:
            bank_name = "HDFC BANK"
            break
        elif "ICICI" in u:
            bank_name = "ICICI BANK"
            break
        elif "CANARA" in u:
            bank_name = "CANARA BANK"
            break
        elif "PUNJAB NATIONAL" in u or "PNB" in u:
            bank_name = "PUNJAB NATIONAL BANK"
            break
        elif "BANK OF BARODA" in u or "BOB" in u:
            bank_name = "BANK OF BARODA"
            break
        elif "INDIAN BANK" in u:
            bank_name = "INDIAN BANK"
            break
