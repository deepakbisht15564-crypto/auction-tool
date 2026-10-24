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

def clean_val(text: str) -> str:
    if not text:
        return ""
    cleaned = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', str(text))
    return re.sub(r'[\r\n\t]+', ' ', cleaned).strip()

def get_precise_notice_box(page):
    # Search for all elements belonging to auction sale notice
    target_terms = [
        "sale notice", "e-auction", "sarfaesi", "asset reconstruction", 
        "arcil", "reserve price", "earnest money", "authorised officer", 
        "description of property", "last date of submission"
    ]
    
    rects = []
    for term in target_terms:
        matches = page.search_for(term)
        if matches:
            rects.extend(matches)
            
    if rects:
        # Combined bounding box of all notice content
        min_x = min([r.x0 for r in rects])
        min_y = min([r.y0 for r in rects])
        max_x = max([r.x1 for r in rects])
        max_y = max([r.y1 for r in rects])
        
        # Expand slightly to encompass borders, headers, and signatures
        crop_box = fitz.Rect(
            max(0, min_x - 30),
            max(0, min_y - 45),
            min(page.rect.width, max_x + 30),
            min(page.rect.height, max_y + 45)
        )
        return crop_box
    else:
        # Fallback to bottom-right newspaper region
        return fitz.Rect(page.rect.width * 0.30, page.rect.height * 0.45, page.rect.width, page.rect.height)

def parse_notice_data(full_text, filename):
    lines = [clean_val(l) for l in full_text.split('\n') if clean_val(l)]
    
    # 1. Exact Bank / ARC Name Detection
    bank_name = ""
    for line in lines[:30]:
        # Match registered institutions
        if re.search(r'(Asset\s*Reconstruction\s*Company\s*\(India\)\s*Ltd\.?|Arcil|State\s*Bank\s*of\s*India|HDFC\s*Bank|ICICI\s*Bank|Canara\s*Bank|Punjab\s*National\s*Bank|Bank\s*of\s*Baroda|Union\s*Bank|Indian\s*Bank|Axis\s*Bank|Kotak\s*Mahindra)', line, re.IGNORECASE):
            m = re.search(r'([A-Za-z\s\(\)]+(?:Bank|Arcil|Asset\s*Reconstruction(?:\s*Company(?:\s*\(India\)\s*Ltd\.?)?)?|Finance|Housing))', line, re.IGNORE
