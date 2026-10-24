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
        elif "AXIS" in u:
            bank_name = "AXIS BANK"
            break
        elif "KOTAK" in u:
            bank_name = "KOTAK MAHINDRA BANK"
            break
        elif "BANK" in u:
            parts = u.split("BRANCH")[0].split("NOTICE")[0]
            bank_name = parts.strip()
            if len(bank_name) > 3:
                break
    
    if not bank_name:
        bank_name = "ARCIL" if "arcil" in full_text.lower() else "BANK NOTICE"
    bank_name = bank_name.upper()

    borrower_name = ""
    clean_fn = filename.replace("_Cutting", "").replace(".pdf", "")
    fn_parts = clean_fn.split()
    name_tokens = [w for w in fn_parts if not re.search(r'(\d|NEW|INDIAN|EXPRESS|TIMES|DAINIK)', w, re.IGNORECASE)]
    if name_tokens:
        borrower_name = " ".join(name_tokens).replace("MrMrs.", "").replace("Mr.", "").replace("Mrs.", "").strip()
    
    if not borrower_name:
        b_match = re.search(r'borrower[s]?[:\s\-]+([A-Za-z\s\.]+)', full_text, re.IGNORECASE)
        if b_match:
            borrower_name = clean_val(b_match.group(1).split("\n")[0])

    rp_matches = re.findall(r'(?:Reserve\s*Price|RP)[^\d₹Rs]*([₹Rs\.]*\s*[\d,]+(?:\.\d+)?)', full_text, re.IGNORECASE)
    reserve_price = clean_val(rp_matches[0]) if rp_matches else ""
    if not reserve_price:
        cur_matches = re.findall(r'(?:Rs\.?|₹)\s*[\d,]+(?:\.\d+)?', full_text)
        reserve_price = clean_val(cur_matches[0]) if cur_matches else ""

    emd_matches = re.findall(r'(?:EMD|Earnest\s*Money)[^\d₹Rs]*([₹Rs\.]*\s*[\d,]+(?:\.\d+)?)', full_text, re.IGNORECASE)
    emd_amount = clean_val(emd_matches[0]) if emd_matches else ""

    all_dates = re.findall(r'(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4})', full_text)
    all_times = re.findall(r'(\d{1,2}[:\.]\d{2}\s*(?:AM|PM|a\.m\.|p\.m\.|hours|hrs)?)', full_text, re.IGNORECASE)

    emd_submit_date = all_dates[0] if len(all_dates) > 0 else ""
    auction_date = all_dates[1] if len(all_dates) > 1 else emd_submit_date
    auction_end_date = auction_date

    emd_submit_time = clean_val(all_times[0]) if len(all_times) > 0 else "05:00 PM"
    auction_start_time = clean_val(all_times[1]) if len(all_times) > 1 else "11:00 AM"
    auction_end_time = clean_val(all_times[2]) if len(all_times) > 2 else "12:00 PM"

    prop_match = re.search(r'(?:description\s*of\s*(?:the\s*)?property|schedule\s*of\s*property)[^\n]*\n([\s\S]{50,450})', full_text, re.IGNORECASE)
    prop_desc = clean_val(prop_match.group(1)) if prop_match else clean_val(full_text[:350])

    row = {
        "Notice File": filename,
        "Bank / ARC": bank_name,
        "Borrower Name": borrower_name,
        "Property Description": prop_desc,
        "Reserve Price": reserve_price,
        "EMD Amount": emd_amount,
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
            crop_box = get_precise_notice_box(page)
            temp_doc = fitz.open()
            temp_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
            temp_doc[0].set_cropbox(crop_box)
            cut_doc.insert_pdf(temp_doc)
            temp_doc.close()

    if len(cut_doc) == 0 and len(doc) > 0:
        crop_box = get_precise_notice_box(doc[0])
        temp_doc = fitz.open()
        temp_doc.insert_pdf(doc, from_page=0, to_page=0)
        temp_doc[0].set_cropbox(crop_box)
        cut_doc.insert_pdf(temp_doc)
        temp_doc.close()

    cut_pdf_bytes = cut_doc.write()
    cut_doc.close()
    doc.close()

    bank_name, rows = parse_notice_data(full_text, filename)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        clean_sheet_name = re.sub(r'[\\/*?:\[\]]', '', bank_name)[:30] or "Notice Data"
        pd.DataFrame(rows).to_excel(writer, sheet_name=clean_sheet_name, index=False)
    excel_bytes = output.getvalue()

    return base_name, cut_pdf_bytes, excel_bytes, rows

# --- UI Layout (Directly Rendered) ---
uploaded_files = st.file_uploader("Upload Auction PDFs (Bulk)", type=["pdf"], accept_multiple_files=True)
start_btn = st.button("Start Processing")

if start_btn:
    if not uploaded_files:
        st.warning("Pehle koi PDF file upload kijiye!")
    else:
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
