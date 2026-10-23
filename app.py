import streamlit as st
import os
import re
import json
import zipfile
import io
import fitz  # PyMuPDF
import pandas as pd
from google import genai
from google.genai import types

st.set_page_config(page_title="Bulk E-Auction Notice Processor", layout="wide")
st.title("📄 Bulk E-Auction PDF Processor")
st.write("Upload up to 100 PDFs to get Clean Vector Cuttings + Multi-Sheet Bank Excels.")

api_key = st.sidebar.text_input("Enter Gemini API Key", type="password")

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
        prompt = f"""
        You are a financial data extraction expert.
        Extract all E-Auction Sale Notices from the following text into a clean JSON list.
        Each notice in the list must have these exact keys:
        - "bank_or_arc_name": Bank or ARC name (e.g. SBI, HDFC, Arcil)
        - "s_no": Serial number
        - "account_prospect_no": Account or Prospect number
        - "borrower_name": Borrower(s) name
        - "property_description": Property details (clean text)
        - "reserve_price": Reserve Price
        - "emd_amount": EMD Amount
        - "inspection_date_time": Inspection Date & Time
        - "emd_last_date_time": EMD Last Date & Time
        - "auction_date_time": Auction Date & Time

        Return ONLY a JSON array of objects: [{{"bank_or_arc_name": "...", ...}}]

        Text:
        {full_text[:45000]}
        """
        
        try:
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.1
                )
            )
            raw_json = response.text.strip()
            # Clean possible markdown wrap
            if raw_json.startswith("```json"):
                raw_json = raw_json[7:-3].strip()
            elif raw_json.startswith("```"):
                raw_json = raw_json[3:-3].strip()

            data = json.loads(raw_json)
            if isinstance(data, dict) and "notices" in data:
                data = data["notices"]
            elif not isinstance(data, list):
                data = [data]

            if data:
                bank_groups = {}
                for item in data:
                    bank = item.get("bank_or_arc_name", "Others") or "Others"
                    bank = str(bank).strip().upper()
                    sheet_name = re.sub(r'[\\/*?:\[\]]', '', bank)[:30] or "Sheet1"
                    row = {
                        "S.No.": item.get("s_no", ""),
                        "Account / Prospect No.": item.get("account_prospect_no", ""),
                        "Borrower(s) Name": item.get("borrower_name", ""),
                        "Property Description": clean_text(item.get("property_description", "")),
                        "Reserve Price": item.get("reserve_price", ""),
                        "EMD Amount": item.get("emd_amount", ""),
                        "Inspection Date & Time": item.get("inspection_date_time", ""),
                        "EMD Last Date & Time": item.get("emd_last_date_time", ""),
                        "Auction Date & Time": item.get("auction_date_time", "")
                    }
                    bank_groups.setdefault(sheet_name, []).append(row)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for sheet, rows in bank_groups.items():
                        pd.DataFrame(rows).to_excel(writer, sheet_name=sheet, index=False)
                excel_bytes = output.getvalue()
        except Exception as e:
            st.error(f"API Error on {filename}: {str(e)}")

    return base_name, cut_pdf_bytes, excel_bytes

uploaded_files = st.file_uploader("Upload Auction PDFs (Bulk)", type=["pdf"], accept_multiple_files=True)

if uploaded_files and st.button("Start Processing"):
    if not api_key:
        st.error("Please enter your Gemini API Key in the sidebar!")
    else:
        client = genai.Client(api_key=api_key)
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        zip_buffer = io.BytesIO()
        total = len(uploaded_files)
        
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for idx, file in enumerate(uploaded_files):
                status_text.text(f"Processing ({idx+1}/{total}): {file.name}")
                file_data = file.getvalue()
                base, cut_pdf, excel_data = process_single_pdf(file_data, file.name, client)
                
                if cut_pdf:
                    zip_file.writestr(f"{base}/{base}_Cutting.pdf", cut_pdf)
                if excel_data:
                    zip_file.writestr(f"{base}/{base}_Data.xlsx", excel_data)
                
                progress_bar.progress((idx + 1) / total)

        status_text.success("All PDFs Processed Successfully!")
        st.download_button(
            label="📥 Download All Deliverables (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Processed_Auctions.zip",
            mime="application/zip"
)
        
