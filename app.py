import streamlit as st
import os
import re
import json
import zipfile
import io
import time
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

def get_notice_crop_rect(page):
    keywords = ["e-auction", "sale notice", "sarfaesi", "reserve price", "emd", "possession notice", "arcil"]
    blocks = page.get_text("blocks")
    matching_rects = []
    
    for b in blocks:
        text = b[4].lower()
        if any(kw in text for kw in keywords):
            matching_rects.append(fitz.Rect(b[:4]))
            
    if not matching_rects:
        return None
        
    # Sirf auction notice wale blocks ka area calculate karna
    combined_rect = matching_rects[0]
    for r in matching_rects[1:]:
        combined_rect |= r
        
    # Thoda margin (padding) add karna taaki border na kate
    margin = 25
    crop_rect = fitz.Rect(
        max(0, combined_rect.x0 - margin),
        max(0, combined_rect.y0 - margin),
        min(page.rect.width, combined_rect.x1 + margin),
        min(page.rect.height, combined_rect.y1 + margin)
    )
    return crop_rect

def process_single_pdf(file_bytes, filename, client):
    base_name = os.path.splitext(filename)[0]
    
    # 1. Asli Cutting (Poore page ke bajaye sirf Notice crop hoga)
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    cut_doc = fitz.open()
    full_text = ""

    for page_num in range(len(doc)):
        page = doc[page_num]
        txt = page.get_text()
        full_text += txt + "\n"
        
        crop_box = get_notice_crop_rect(page)
        if crop_box:
            temp_doc = fitz.open()
            temp_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
            temp_page = temp_doc[0]
            temp_page.set_cropbox(crop_box)
            cut_doc.insert_pdf(temp_doc)
            temp_doc.close()

    if len(cut_doc) == 0 and len(doc) > 0:
        cut_doc.insert_pdf(doc)

    cut_pdf_bytes = cut_doc.write()
    cut_doc.close()
    doc.close()

    excel_bytes = None
    all_rows = []

    # 2. AI Extraction (Auto retry agar 503 error aaye)
    if full_text.strip():
        prompt = f"""
Extract all E-Auction Sale Notices from the text into a clean JSON array.
Keys for each object:
- "bank_or_arc_name"
- "s_no"
- "account_prospect_no"
- "borrower_name"
- "property_description"
- "reserve_price"
- "emd_amount"
- "inspection_date_time"
- "emd_last_date_time"
- "auction_date_time"

Return ONLY valid raw JSON list of objects: [{{"bank_or_arc_name": "...", ...}}]

Text:
{full_text[:35000]}
"""
        raw_json = None
        last_err = ""

        # Agar server busy ho toh 3 baar apne aap retry karega
        for attempt in range(3):
            try:
                response = client.models.generate_content(
                    model="gemini-3.8-flash",
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        temperature=0.1
                    )
                )
                if response and response.text:
                    raw_json = response.text.strip()
                    break
            except Exception as e:
                last_err = str(e)
                time.sleep(3)

        if raw_json:
            try:
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
                            "Notice File": filename,
                            "Bank / ARC": bank,
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
                        all_rows.append(row)

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        for sheet, rows in bank_groups.items():
                            pd.DataFrame(rows).to_excel(writer, sheet_name=sheet, index=False)
                    excel_bytes = output.getvalue()
            except Exception as e:
                st.error(f"Excel formatting error: {str(e)}")
        else:
            st.error(f"Google server busy: {last_err}")

    return base_name, cut_pdf_bytes, excel_bytes, all_rows

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
        master_rows = []
        
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for idx, file in enumerate(uploaded_files):
                status_text.text(f"Processing ({idx+1}/{total}): {file.name}")
                file_data = file.getvalue()
                base, cut_pdf, excel_data, rows = process_single_pdf(file_data, file.name, client)
                
                if cut_pdf:
                    zip_file.writestr(f"{base}/{base}_Cutting.pdf", cut_pdf)
                if excel_data:
                    zip_file.writestr(f"{base}/{base}_Data.xlsx", excel_data)
                if rows:
                    master_rows.extend(rows)
                
                progress_bar.progress((idx + 1) / total)

            if master_rows:
                master_df = pd.DataFrame(master_rows)
                zip_file.writestr("CRM_All_Accounts_Master.csv", master_df.to_csv(index=False).encode('utf-8'))

        status_text.success("Processing Complete!")
        st.download_button(
            label="📥 Download All Deliverables (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Processed_Auctions.zip",
            mime="application/zip"
        )
