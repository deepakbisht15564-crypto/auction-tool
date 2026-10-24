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
        
        return fitz.Rect(
            max(0, bbox.x0 - 15),
            max(0, bbox.y0 - 20),
            min(page.rect.width, bbox.x1 + 15),
            min(page.rect.height, bbox.y1 + 40)
        )
    else:
        return fitz.Rect(page.rect.width * 0.38, page.rect.height * 0.52, page.rect.width, page.rect.height)

def parse_auction_fields(full_text, filename):
    text_lower = full_text.lower()
    
    # Bank name
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

    # Primary Borrower only
    borrower_name = ""
    b_match = re.search(r'(?:borrower|name\s*of\s*borrower)s?[:\s\-]+([A-Za-z\.\s]+?)(?=\s*(?:co-borrower|guarantor|legal|mortgagor|\/|\n|rs\.|₹|property))', full_text, re.IGNORECASE)
    if b_match:
        borrower_name = clean_val(b_match.group(1))
    else:
        name_part = re.search(r'(?:MrMrs\.|Mr\.|Mrs\.)?([A-Z\s]+?)(?=\s+THE\s+NEW|\s+INDIAN|\.pdf)', filename, re.IGNORECASE)
        borrower_name = clean_val(name_part.group(1)) if name_part else ""

    # Reserve Price
    rp_match = re.search(r'(?:reserve\s*price|r\.?p\.?)[^\d₹Rs]*([₹Rs\.]*\s*[\d,]+(?:\.\d+)?)', full_text, re.IGNORECASE)
    reserve_price = clean_val(rp_match.group(1)) if rp_match else ""

    # Dates & Times
    dates = re.findall(r'(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4})', full_text)
    times = re.findall(r'(\d{1,2}[:\.]\d{2}\s*(?:am|pm|a\.m\.|p\.m\.|hours|hrs)?)', full_text, re.IGNORECASE)

    emd_submit_date = dates[0] if len(dates) > 0 else ""
    auction_date = dates[1] if len(dates) > 1 else (dates[0] if dates else "")
    auction_end_date = auction_date

    emd_submit_time = clean_val(times[0]) if len(times) > 0 else ""
    auction_start_time = clean_val(times[1]) if len(times) > 1 else ""
    auction_end_time = clean_val(times[2]) if len(times) > 2 else ""

    # Property Description
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
            temp_doc[0].set_cropbox(crop_box)
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
    doc.close()

    bank_name, rows = parse_auction_fields(full_text, filename)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sheet = re.sub(r'[\\/*?:\[\]]', '', bank_name)[:30] or "Sheet1"
        pd.DataFrame(rows).to_excel(writer, sheet_name=sheet, index=False)
    excel_bytes = output.getvalue()

    return base_name, cut_pdf_bytes, excel_bytes, rows

uploaded_files = st.file_uploader("Upload Auction PDFs (Bulk)", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Start Processing"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        zip_buffer = io.BytesIO()
        total = len(uploaded_files)
        master_rows = []
        
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for idx, file in enumerate(uploaded_files):
                status_text.text(f"Processing ({idx+1}/{total}): {file.name}")
                file_data = file.getvalue()
                base, cut_pdf, excel_data, rows = process_single_pdf(file_data, file.name)
                
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
