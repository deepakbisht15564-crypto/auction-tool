import streamlit as st
import os
import re
import zipfile
import io
import fitz
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Bulk E-Auction Notice Processor",
    layout="wide"
)
st.title("Bulk E-Auction PDF Processor")
st.write("Full Page PDF + Formatted Auto-Fit Styled Excel.")

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

def extract_bank_name(lines, full_text):
    for line in lines[:40]:
        u = line.upper()
        if "ARCIL" in u or "ASSET RECONSTRUCTION" in u:
            return "ARCIL"
        elif "STATE BANK OF INDIA" in u or "SBI" in u:
            return "STATE BANK OF INDIA"
        elif "CANARA BANK" in u:
            return "CANARA BANK"
        elif "PUNJAB NATIONAL BANK" in u or "PNB" in u:
            return "PUNJAB NATIONAL BANK"
        elif "BANK OF BARODA" in u or "BOB" in u:
            return "BANK OF BARODA"
        elif "INDIAN BANK" in u:
            return "INDIAN BANK"
        elif "UNION BANK" in u:
            return "UNION BANK OF INDIA"
        elif "HDFC BANK" in u:
            return "HDFC BANK"
        elif "ICICI BANK" in u:
            return "ICICI BANK"
        elif "AXIS BANK" in u:
            return "AXIS BANK"
        elif "KOTAK MAHINDRA" in u:
            return "KOTAK MAHINDRA BANK"
        elif "BANK" in u:
            candidate = u.split("BRANCH")[0].split("NOTICE")[0].strip()
            if len(candidate) > 4:
                return candidate
    return "ARCIL" if "arcil" in full_text.lower() else "BANK NOTICE"

def extract_borrower(filename, lines):
    clean_fn = filename.replace("_Cutting", "").replace(".pdf", "")
    tokens = [w for w in clean_fn.split() if not any(c.isdigit() for c in w)]
    filter_words = ["THE", "NEW", "INDIAN", "EXPRESS", "TIMES", "DAINIK", "NOTICE", "E-AUCTION", "SALE"]
    clean_tokens = [w for w in tokens if w.upper() not in filter_words]
    if clean_tokens:
        b_name = " ".join(clean_tokens)
        return b_name.replace("MrMrs.", "").replace("Mr.", "").replace("Mrs.", "").strip()

    for idx, line in enumerate(lines):
        if "borrower"
        
