import streamlit as st
import os
import re
import zipfile
import io
import fitz  # PyMuPDF
import pandas as pd

st.set_page_config(page_title="Bulk E-Auction Notice Processor", layout="wide")
st.title("📄 Bulk E-Auction PDF Processor")
st.write("100% Local Engine — No API Keys, No Quota Limits, No Server Errors.")

def clean_text(text: str) -> str:
    return re.sub(r'[\r\n\t]+', ' ', str(text)).strip()

def extract_notice_details(full_text, filename):
    # 1. Bank / ARC Identification
    bank_name = "OTHERS"
    text_lower = full_text.lower()
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

    # 2. Reserve Price
    rp_match = re.search(r'(?:reserve\s*price|r\.?p\.?)[^\d₹Rs]*([₹Rs\.]*\s*[\d,]+(?:\.\d+)?)', full_text, re.IGNORECASE)
    reserve_price = rp_match.group(1).strip() if rp_match else ""

    # 3. EMD Amount
    emd_match = re.search(r'(?:emd\s*amount|emd)[^\d₹Rs]*([₹Rs\.]*\s*[\d,]+(?:\.\d+)?)', full_text, re.IGNORECASE)
    emd_amount = emd_match.group(1).strip() if emd_match else ""

    # 4. Account / NPA / Prospect No
    acc_match = re.search(r'(?:account|ac|prospect|loan|npa)\s*(?:no|number)?[:\s\.\-]*([A-Z0-9\/\-]+)', full_text, re.IGNORECASE)
    account_no = acc_match.group(1).strip() if acc_match else ""

    # 5. Borrower Name
    borrower_match = re.search(r'(?:borrower|guarantor|name\s*of\s*borrower)s?[:\s\-]+([A-Za-z0-9\s,\.\/]+?)(?=\n|property|description|reserve|outstanding)', full_text, re.IGNORECASE)
    borrower_name = clean_text(borrower_match.group(1)) if borrower_match else ""

    # 6. Auction Date & Time
    auction_match = re.search(r'(?:auction\s*date|date\s*of\s*auction|date\s*and\s*time)[^\d]*(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4}[^\n]*)', full_text, re.IGNORECASE)
    auction_date = clean_text(auction_match.group(1)) if auction_match else ""

    # 7. EMD Last Date
    emd_date_match = re.search(r'(?:last\s*date\s*of\s*submission|last\s*date\s*for\s*submission|emd\s*last\s*date)[^\d]*(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4})', full_text, re.IGNORECASE)
    emd_date = emd_date_match.group(1).strip() if emd_date_match else ""

    # 8. Property Description
    prop_match = re.search(r'(?:description\s*of\s*property|property\s*details|schedule)[^\n]*\n([\s\S]{80,500})', full_text, re.IGNORECASE)
    property_desc = clean_text(prop_match.group(1)) if prop_match else clean_text(full_text[:350])

    row = {
        "Notice File": filename,
        "Bank / ARC": bank_name,
        "S.No.": "1",
        "Account / Prospect No.": account_no,
        "Borrower(s) Name": borrower_name,
        "Property Description": property_desc,
        "Reserve Price": reserve_price,
        "EMD Amount": emd_amount,
        "Inspection Date & Time": "",
        "EMD Last Date & Time": emd_date,
        "Auction Date & Time": auction_date
    }
    return bank_name, [row]

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
        
        # Crop exactly around notice keywords with margin
        return fitz.Rect(
            max(0, bbox.x0 - 20),
            max(0, bbox.y0 - 25),
            min(page.rect.width, bbox.x1 + 20),
            min(page.rect.height, bbox.y1 + 100)
        )
    else:
        # Fallback to notice section (bottom half)
        return fitz.Rect(0, page.rect.height * 0.45, page.rect.width, page.rect.height)

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

    # Local Excel Creation
    bank_name, rows = extract_notice_details(full_text, filename)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sheet = re.sub(r'[\\/*?:\[\]]', '', bank_name)[:30] or "Sheet1"
        pd.DataFrame(rows).to_excel(writer, sheet_name=sheet, index=False)
    excel_bytes = output.getvalue()

    return base_name, cut_pdf_bytes, excel_bytes, rows

uploaded_files = st.file_uploader("Upload Auction PDFs (Bulk)", type=["pdf"], accept_multiple_files=True)

if uploaded_files and st.button("Start Processing"):
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
