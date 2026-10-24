import streamlit as st
import os
import re
import json
import zipfile
import io
import fitz  # PyMuPDF
import pandas as pd
from groq import Groq

st.set_page_config(page_title="Bulk E-Auction Notice Processor", layout="wide")
st.title("📄 Bulk E-Auction PDF Processor")
st.write("Clean Box Cuttings + Field-Wise Excel Extraction.")

api_key = st.sidebar.text_input("Enter Groq API Key (starts with gsk_)", type="password")

def clean_for_excel(text: str) -> str:
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
        # Calculate strict bounding box of the auction notice
        bbox = found_rects[0]
        for r in found_rects[1:]:
            bbox |= r
        
        # Crop strictly around the notice box (removes left side classifieds/matrimonial)
        return fitz.Rect(
            max(0, bbox.x0 - 25),
            max(0, bbox.y0 - 25),
            min(page.rect.width, bbox.x1 + 25),
            min(page.rect.height, bbox.y1 + 40)
        )
    else:
        # Fallback to bottom-right quadrant where commercial notices usually sit
        return fitz.Rect(page.rect.width * 0.35, page.rect.height * 0.50, page.rect.width, page.rect.height)

def process_single_pdf(file_bytes, filename, groq_client):
    base_name = os.path.splitext(filename)[0]
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    
    # 1. Exact Notice Box Cutting
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
        crop_box = get_notice_crop_box(doc[0])
        temp_doc = fitz.open()
        temp_doc.insert_pdf(doc, from_page=0, to_page=0)
        temp_doc[0].set_cropbox(crop_box)
        cut_doc.insert_pdf(temp_doc)
        temp_doc.close()

    cut_pdf_bytes = cut_doc.write()
    cut_doc.close()

    # 2. Field-Wise AI Extraction via Groq
    excel_bytes = None
    all_rows = []

    if full_text.strip():
        prompt = f"""
You are a financial E-Auction data extraction expert.
Extract all Bank / ARC E-Auction Sale Notices from the text into a clean JSON array.
Strict Extraction Rules:
- "bank_or_arc_name": Exact Bank or ARC name (e.g. Arcil, SBI, Canara Bank).
- "borrower_name": ONLY the primary Borrower's name. Strictly exclude Co-Borrowers, Guarantors, Legal Heirs, Mortgagors, or Partners.
- "property_description": Concise property location and details.
- "reserve_price": Reserve Price amount with currency symbol if present.
- "emd_submit_date": Last date for EMD submission (DD/MM/YYYY).
- "emd_submit_time": Time for EMD submission deadline (e.g. 05:00 PM).
- "auction_date": Date of auction (DD/MM/YYYY).
- "auction_start_time": Auction starting time (e.g. 11:00 AM).
- "auction_end_date": Auction ending date (DD/MM/YYYY).
- "auction_end_time": Auction ending time (e.g. 12:00 PM).

Return ONLY a valid JSON list of objects:
[{{"bank_or_arc_name": "...", "borrower_name": "...", "property_description": "...", "reserve_price": "...", "emd_submit_date": "...", "emd_submit_time": "...", "auction_date": "...", "auction_start_time": "...", "auction_end_date": "...", "auction_end_time": "..."}}]

Text:
{full_text[:30000]}
"""
        models_to_try = ["openai/gpt-oss-20b", "openai/gpt-oss-120b"]
        raw_json = None
        last_err = ""

        for model_name in models_to_try:
            try:
                chat_completion = groq_client.chat.completions.create(
                    messages=[
                        {"role": "system", "content": "You output only clean, valid JSON without explanations or markdown."},
                        {"role": "user", "content": prompt}
                    ],
                    model=model_name,
                    temperature=0.1
                )
                raw_json = chat_completion.choices[0].message.content.strip()
                if raw_json:
                    break
            except Exception as e:
                last_err = str(e)
                continue

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
                            "Borrower Name": clean_for_excel(item.get("borrower_name", "")),
                            "Property Description": clean_for_excel(item.get("property_description", "")),
                            "Reserve Price": clean_for_excel(item.get("reserve_price", "")),
                            "EMD Submit Date": clean_for_excel(item.get("emd_submit_date", "")),
                            "EMD Submit Time": clean_for_excel(item.get("emd_submit_time", "")),
                            "Auction Date": clean_for_excel(item.get("auction_date", "")),
                            "Auction Start Time": clean_for_excel(item.get("auction_start_time", "")),
                            "Auction End Date": clean_for_excel(item.get("auction_end_date", "")),
                            "Auction End Time": clean_for_excel(item.get("auction_end_time", ""))
                        }
                        bank_groups.setdefault(sheet_name, []).append(row)
                        all_rows.append(row)

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        for sheet, rows in bank_groups.items():
                            pd.DataFrame(rows).to_excel(writer, sheet_name=sheet, index=False)
                    excel_bytes = output.getvalue()
            except Exception as e:
                st.error(f"Excel generation error on {filename}: {str(e)}")
        else:
            st.error(f"AI Extraction error: {last_err}")

    doc.close()
    return base_name, cut_pdf_bytes, excel_bytes, all_rows

uploaded_files = st.file_uploader("Upload Auction PDFs (Bulk)", type=["pdf"], accept_multiple_files=True)

if uploaded_files and st.button("Start Processing"):
    if not api_key:
        st.error("Please enter your Groq API Key in the sidebar!")
    else:
        groq_client = Groq(api_key=api_key)
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        zip_buffer = io.BytesIO()
        total = len(uploaded_files)
        master_rows = []
        
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for idx, file in enumerate(uploaded_files):
                status_text.text(f"Processing ({idx+1}/{total}): {file.name}")
                file_data = file.getvalue()
                base, cut_pdf, excel_data, rows = process_single_pdf(file_data, file.name, groq_client)
                
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

        status_text.success("All Done!")
        st.download_button(
            label="📥 Download All Deliverables (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="Processed_Auctions.zip",
            mime="application/zip"
        )
