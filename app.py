import streamlit as st
import os
import re
import json
import zipfile
import io
import fitz  # PyMuPDF
import pandas as pd
from pydantic import BaseModel, Field
from google import genai
from google.genai import types

st.set_page_config(page_title="Bulk E-Auction Notice Processor", layout="wide")
st.title("📄 Bulk E-Auction PDF Processor")
st.write("Upload up to 100 PDFs to get Clean Vector Cuttings + Multi-Sheet Bank Excels.")

api_key = st.sidebar.text_input("Enter Gemini API Key", type="password")

class AuctionNotice(BaseModel):
    bank_or_arc_name: str = Field(description="Name of Bank or ARC, e.g. SBI, HDFC, Arcil")
    s_no: str = Field(description="Serial number")
    account_prospect_no: str = Field(description="Account or Prospect number")
    borrower_name: str = Field(description="Borrower(s) name")
    property_description: str = Field(description="Property description (clean text, only '.' and ',' allowed)")
    reserve_price: str = Field(description="Reserve price")
    emd_amount: str = Field(description="EMD amount")
    inspection_date_time: str = Field(description="Inspection Date & Time")
    emd_last_date_time: str = Field(description="EMD Last Date & Time")
    auction_date_time: str = Field(description="Auction Date & Time")

class AuctionList(BaseModel):
    notices: list[AuctionNotice]

def clean_text(text: str) -> str:
    return re.sub(r'[^a-zA-Z0-9\s.,]', ' ', str(text))

def process_single_pdf(file_bytes, filename, client):
    base_name = os.path.splitext(filename)[0]
    
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    cut_doc = fitz.open()
    keywords = ["e-auction", "sale notice", "sarfaesi", "reserve price", "emd", "auction", "possession"]
    full_text = ""

    for page_num in range(len(doc)):
        page = doc[page_num]
        txt = page.get_text()
        full_text += txt + "\n"
        if any(kw in txt.lower() for kw in keywords):
            cut_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)

    if len(cut_doc) == 0 and len(doc) > 0:
        cut_doc.insert_pdf(doc)

    cut_pdf_bytes = cut_doc.write()
    cut_doc.close()
    doc.close()

    excel_bytes = None
    if full_text.strip():
        prompt = "Extract all E-Auction Sale Notices published by any Bank or ARC. Keep property description clean (only '.' and ',' allowed)."
        try:
            response = client.models.generate_content(
                model="gemini-1.5-flash",
                contents=[prompt, full_text[:40000]],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=AuctionList,
                    temperature=0.1
                )
            )
            data = json.loads(response.text).get("notices", [])
            if data:
                bank_groups = {}
                for item in data:
                    bank = item.get("bank_or_arc_name", "Others") or "Others"
                    bank = str(bank).strip().upper()
                    sheet_name = re.sub(r'[\\/*?:\[\]]', '', bank)[:30] or "Sheet1"
                    row = {
                        "S.No.": item.get("s_no"),
                        "Account / Prospect No.": item.get("account_prospect_no"),
                        "Borrower(s) Name": item.get("borrower_name"),
                        "Property Description": clean_text(item.get("property_description", "")),
                        "Reserve Price": item.get("reserve_price"),
                        "EMD Amount": item.get("emd_amount"),
                        "Inspection Date & Time": item.get("inspection_date_time"),
                        "EMD Last Date & Time": item.get("emd_last_date_time"),
                        "Auction Date & Time": item.get("auction_date_time")
                    }
                    bank_groups.setdefault(sheet_name, []).append(row)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for sheet, rows in bank_groups.items():
                        pd.
    
