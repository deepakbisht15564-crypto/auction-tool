import streamlit as st
import os
import re
import zipfile
import io
import fitz  # PyMuPDF
import pandas as pd

st.set_page_config(page_title="Bulk E-Auction Notice Processor", layout="wide")
st.title("Bulk E-Auction PDF Processor")
st.write("Full Page High-Res Output + Structured Excel Data Extraction.")

uploaded_files = st.file_uploader("Upload Auction PDFs (Bulk)", type=["pdf"], accept_multiple_files=True)
start_btn = st.button("Start Processing")

def clean_val(text):
    if not text:
        return ""
    val = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', str(text))
    return " ".join(val.split()).strip()

def parse_notice_data(page, filename):
    full_text = page.get_text()
    lines = [clean_val(l) for l in full_text.split("\n") if clean_val(l)]

    # 1. Exact Bank / ARC Name Detection
    bank_name = ""
    for line in lines[:30]:
        u = line.upper()
        if "ARCIL" in u or "ASSET RECONSTRUCTION" in u:
            bank_name = "ARCIL"
            break
        elif "STATE BANK" in u or "SBI" in u:
            bank_name = "STATE BANK OF INDIA"
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
        elif "HDFC" in u:
            bank_name = "HDFC BANK"
            break
        elif "ICICI" in u:
            bank_name = "ICICI BANK"
            break
        elif "AXIS" in u:
            bank_name = "AXIS BANK"
            break
        elif "KOTAK" in u:
            bank_name = "KOTAK MAHINDRA BANK"
            break
        elif "BANK" in u:
            candidate = u.split("BRANCH")[0].split("NOTICE")[0].strip()
            if len(candidate) > 4:
                bank_name = candidate
                break

    if not bank_name:
        bank_name = "ARCIL" if "arcil" in full_text.lower() else "BANK NOTICE"
    bank_name = bank_name.upper()

    # 2. Borrower Name (Strictly only the primary borrower)
    borrower_name = ""
    clean_fn = filename.replace("_Cutting", "").replace(".pdf", "")
    tokens = [w for w in clean_fn.split() if not any(c.isdigit() for c in w)]
    filter_words = ["THE", "NEW", "INDIAN", "EXPRESS", "TIMES", "DAINIK", "NOTICE", "E-AUCTION", "SALE"]
    clean_tokens = [w for w in tokens if w.upper() not in filter_words]
    if clean_tokens:
        borrower_name = " ".join(clean_tokens).replace("MrMrs.", "").replace("Mr.", "").replace("Mrs.", "").strip()

    if not borrower_name:
        for line in lines:
            if "borrower" in line.lower():
                val = line.split(":")[-1]
                val = val.split("/")[0].split("Co-")[0].strip()
                if len(val) > 2:
                    borrower_name = clean_val(val)
                    break

    # 3. Reserve Price & EMD Amounts
    amounts = re.findall(r'(?:Rs\.?|₹|INR)\s*([\d,]+(?:\.\d+)?)', full_text, re.I)
    if not amounts:
        amounts = re.findall(r'\b\d{1,3}(?:,\d{2,3})+(?:\.\d+)?\b', full_text)

    reserve_price = ""
    emd_amount = ""
    if len(amounts) >= 2:
        reserve_price = f"Rs. {clean_val(amounts[0])}"
        emd_amount = f"Rs. {clean_val(amounts[1])}"
    elif len(amounts) == 1:
        reserve_price = f"Rs. {clean_val(amounts[0])}"

    # 4. Dates & Times
    dates = re.findall(r'\b\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4}\b', full_text)
    times = re.findall(r'\b\d{1,2}[:\.]\d{2}\s*(?:AM|PM|a\.m\.|p\.m\.|hrs|hours)?\b', full_text, re.I)

    emd_submit_date = dates[0] if len(dates) > 0 else ""
    auction_date = dates[1] if len(dates) > 1 else emd_submit_date
    auction_end_date = auction_date

    emd_submit_time = clean_val(times[0]) if len(times) > 0 else "05:00 PM"
    auction_start_time = clean_val(times[1]) if len(times) > 1 else "11:00 AM"
    auction_end_time = clean_val(times[2]) if len(times) > 2 else "12:00 PM"

    # 5. Property Description
    prop_desc = ""
    for idx, line in enumerate(lines):
        if any(kw in line.lower() for kw in ["property", "schedule", "immovable", "piece and parcel"]):
            prop_desc = " ".join(lines[idx:idx+5])
            break
    if not prop_desc:
        mid = len(lines) // 2
        prop_desc = " ".join(lines[mid:mid+4])

    row = {
        "Notice File": filename,
        "Bank / ARC": bank_name,
        "Borrower Name": borrower_name,
        "Property Description": clean_val(prop_desc),
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

    full_page_doc = fitz.open()
    target_page_idx = 0

    # Locate the page with the sale notice
    for page_num in range(len(doc)):
        txt = doc[page_num].get_text().lower()
        if any(kw in txt for kw in ["auction", "sale notice", "sarfaesi", "arcil", "reserve price"]):
            target_page_idx = page_num
            break

    # Extract the full original vector page (no clipping / no text distortion)
    full_page_doc.insert_pdf(doc, from_page=target_page_idx, to_page=target_page_idx)
    target_page = doc[target_page_idx]

    pdf_bytes = full_page_doc.write()
    full_page_doc.close()

    # Extract Excel Data
    bank_name, rows = parse_notice_data(target_page, filename)
    doc.close()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        s_name = re.sub(r'[\\/*?:\[\]]', '', bank_name)[:30] or "Notice Data"
        pd.DataFrame(rows).to_excel(writer, sheet_name=s_name, index=False)
    excel_bytes = output.getvalue()

    return base_name, pdf_bytes, excel_bytes, rows

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
                base, notice_pdf, excel_data, rows = process_single_pdf(file_data, file.name)

                if notice_pdf:
                    zip_file.writestr(f"{base}/{base}_NoticePage.pdf", notice_pdf)
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
