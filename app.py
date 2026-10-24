import streamlit as st
import os
import re
import zipfile
import io
import fitz  # PyMuPDF
import pandas as pd

st.set_page_config(page_title="Bulk E-Auction Notice Processor", layout="wide")
st.title("📄 Bulk E-Auction PDF Processor")
st.write("Dynamic Bank Detection + Exact Notice Box Cuttings.")

def clean_val(text: str) -> str:
    if not text:
        return ""
    cleaned = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', str(text))
    return re.sub(r'[\r\n\t]+', ' ', cleaned).strip()

def get_notice_crop_box(page):
    # Anchor words that mark the START of an auction notice
    header_keywords = [
        "sale notice", "e-auction notice", "possession notice", 
        "demand notice", "appendix iv", "sarfaesi", "asset reconstruction", 
        "auction sale", "public notice for sale"
    ]
    
    start_rects = []
    for kw in header_keywords:
        r = page.search_for(kw)
        if r:
            start_rects.extend(r)
            
    # Bottom keywords that define the content bounds
    content_keywords = ["reserve price", "emd", "date of auction", "terms and conditions", "authorised officer", "arcil"]
    content_rects = []
    for kw in content_keywords:
        r = page.search_for(kw)
        if r:
            content_rects.extend(r)

    if start_rects:
        # Lowest y0 among the found headers (start point of notice box)
        y_top = min([r.y0 for r in start_rects]) - 35
        x_left = min([r.x0 for r in start_rects]) - 25
        
        if content_rects:
            y_bottom = max([r.y1 for r in content_rects]) + 40
            x_right = max([r.x1 for r in content_rects]) + 25
        else:
            y_bottom = page.rect.height
            x_right = page.rect.width

        # Strictly contain inside the page
        return fitz.Rect(
            max(0, x_left),
            max(0, y_top),
            min(page.rect.width, x_right),
            min(page.rect.height, y_bottom)
        )
    else:
        # Fallback to bottom area where commercial notices are published
        return fitz.Rect(page.rect.width * 0.30, page.rect.height * 0.50, page.rect.width, page.rect.height)

def parse_auction_fields(full_text, filename):
    lines = [clean_val(l) for l in full_text.split('\n') if clean_val(l)]
    
    # 1. Dynamic Bank / ARC Name Extraction (No hardcoded bank list)
    bank_name = ""
    for line in lines[:25]:
        if re.search(r'(?:bank|arcil|asset\s*reconstruction|finance|housing|co-operative|gramin)', line, re.IGNORECASE):
            # Clean up unwanted words if any
            bank_name = re.sub(r'(branch|e-auction|sale notice|notice|appendix).*', '', line, flags=re.IGNORECASE).strip()
            if len(bank_name) > 3:
                break
    
    if not bank_name:
        b_match = re.search(r'([A-Za-z\s]+(?:Bank|Reconstruction|Finance|ARC|Capital))', full_text, re.IGNORECASE)
        bank_name = clean_val(b_match.group(1)) if b_match else "Commercial Bank / ARC"

    bank_name = bank_name.upper()

    # 2. Borrower Name (Strictly primary borrower only)
    borrower_name = ""
    b_match = re.search(r'(?:name\s*of\s*borrower|borrower)s?[:\s\-]+([A-Za-z\.\s]+?)(?=\s*(?:co-borrower|guarantor|legal|mortgagor|\/|\n|rs\.|₹|schedule|property))', full_text, re.IGNORECASE)
    if b_match:
        borrower_name = clean_val(b_match.group(1))
    else:
        # Extract name directly from standard format file title
        name_part = re.search(r'(?:MrMrs\.|Mr\.|Mrs\.)?([A-Z\s]+?)(?=\s+THE\s+NEW|\s+INDIAN|\.pdf)', filename, re.IGNORECASE)
        borrower_name = clean_val(name_part.group(1)) if name_part else ""

    # 3. Reserve Price
    rp_match = re.search(r'(?:reserve\s*price|r\.?p\.?)[^\d₹Rs]*([₹Rs\.]*\s*[\d,]+(?:\.\d+)?)', full_text, re.IGNORECASE)
    reserve_price = clean_val(rp_match.group(1)) if rp_match else ""

    # 4. Dates & Times Extraction
    dates = re.findall(r'(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4})', full_text)
    times = re.findall(r'(\d{1,2}[:\.]\d{2}\s*(?:am|pm|a\.m\.|p\.m\.|hours|hrs)?)', full_text, re.IGNORECASE)

    emd_submit_date = dates[0] if len(dates) > 0 else ""
    auction_date = dates[1] if len(dates) > 1 else (dates[0] if dates else "")
    auction_end_date = auction_date

    emd_submit_time = clean_val(times[0]) if len(times) > 0 else ""
    auction_start_time = clean_val(times[1]) if len(times) > 1 else ""
    auction_end_time = clean_val(times[2]) if len(times) > 2 else ""

    # 5. Property Description
    prop_match = re.search(r'(?:description\s*of\s*(?:the\s*)?property|schedule\s*of\s*(?:immovable\s*)?property|property\s*details)[^\n]*\n([\s\S]{50,450})', full_text, re.IGNORECASE)
    prop_desc = clean_val(prop_match.group(1)) if prop_match else clean_val(full_text[:350])

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
    doc.close()

    # Dynamic Bank Excel Generation
    bank_name, rows = parse_auction_fields(full_text, filename)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sheet = re.sub(r'[\\/*?:\[\]]', '', bank_name)[:30] or "Notice Data"
        pd.DataFrame(rows).to_excel(writer, sheet_name=sheet, index=False)
    excel_bytes = output.getvalue()

    return base_name, cut_pdf_bytes, excel_bytes, rows

# --- Streamlit UI ---
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
