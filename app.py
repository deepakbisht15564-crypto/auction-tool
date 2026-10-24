import streamlit as st
import os
import re
import zipfile
import io
import fitz
import pandas as pd

st.set_page_config(
    page_title="Bulk E-Auction Notice Processor",
    layout="wide"
)
st.title("Bulk E-Auction PDF Processor")
st.write("Full Uncut Page PDF + Bank-Named Excel Sheets.")

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
    for line in lines[:35]:
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

    for line in lines:
        if "borrower" in line.lower():
            val = line.split(":")[-1]
            val = val.split("/")[0].split("Co-")[0].strip()
            if len(val) > 2:
                return clean_val(val)
    return "N/A"

def parse_notice_document(doc, filename):
    t_idx = 0
    for idx in range(len(doc)):
        txt = doc[idx].get_text().lower()
        if any(w in txt for w in ["sale notice", "e-auction", "sarfaesi", "arcil", "reserve price"]):
            t_idx = idx
            break

    target_page = doc[t_idx]

    full_rect = target_page.rect
    target_page.set_cropbox(full_rect)
    target_page.set_mediabox(full_rect)

    full_doc = fitz.open()
    full_doc.insert_pdf(doc, from_page=t_idx, to_page=t_idx)
    full_doc[0].set_cropbox(full_rect)
    full_doc[0].set_mediabox(full_rect)
    pdf_bytes = full_doc.write()
    full_doc.close()

    full_text = target_page.get_text()
    lines = [clean_val(l) for l in full_text.split("\n") if clean_val(l)]

    bank_name = extract_bank_name(lines, full_text).upper()
    borrower = extract_borrower(filename, lines)

    amounts = re.findall(r'(?:Rs\.?|₹|INR)\s*([\d,]+(?:\.\d+)?)', full_text, re.I)
    if not amounts:
        amounts = re.findall(r'\b\d{1,3}(?:,\d{2,3})+(?:\.\d+)?\b', full_text)

    reserve_price = f"Rs. {clean_val(amounts[0])}" if len(amounts) >= 1 else ""
    emd_amount = f"Rs. {clean_val(amounts[1])}" if len(amounts) >= 2 else ""

    dates = re.findall(r'\b\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4}\b', full_text)
    times = re.findall(r'\b\d{1,2}[:\.]\d{2}\s*(?:AM|PM|a\.m\.|p\.m\.|hrs|hours)?\b', full_text, re.I)

    emd_last_date = dates[0] if len(dates) > 0 else ""
    auction_date = dates[1] if len(dates) > 1 else emd_last_date

    emd_last_time = clean_val(times[0]) if len(times) > 0 else "05:00 PM"
    auction_time = clean_val(times[1]) if len(times) > 1 else "11:00 AM - 12:00 PM"

    prop_desc = ""
    for idx, line in enumerate(lines):
        if any(kw in line.lower() for kw in ["property", "schedule", "immovable", "piece and parcel"]):
            prop_desc = " ".join(lines[idx:idx+4])
            break
    if not prop_desc:
        mid = len(lines) // 2
        prop_desc = " ".join(lines[mid:mid+4]) if lines else "Full notice details available on attached PDF."

    acct_match = re.findall(r'\b[A-Z0-9]{8,20}\b', filename)
    acct_no = acct_match[0] if acct_match else ""

    row_data = {
        "S.No.": 1,
        "Bank / ARC": bank_name,
        "Account / Prospect No.": acct_no,
        "Borrower(s) Name": borrower,
        "Property Description": clean_prop(prop_desc),
        "Reserve Price": reserve_price,
        "EMD Amount": emd_amount,
        "Inspection Date & Time": "As per notice terms",
        "EMD Last Date & Time": f"{emd_last_date} {emd_last_time}".strip(),
        "Auction Date & Time": f"{auction_date} {auction_time}".strip()
    }

    clean_sheet = re.sub(r'[\\/*?:\[\]]', '', bank_name)[:30].strip() or "NOTICE DATA"
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        pd.DataFrame([row_data]).to_excel(writer, sheet_name=clean_sheet, index=False)
    excel_bytes = out.getvalue()

    master_row = dict(row_data)
    master_row["Source_File"] = filename

    return pdf_bytes, excel_bytes, [master_row]

if start_btn:
    if not uploaded_files:
        st.warning("Please upload auction PDF files!")
    else:
        prog = st.progress(0)
        status = st.empty()

        zip_buf = io.BytesIO()
        total = len(uploaded_files)
        all_master = []

        with zipfile.ZipFile(zip_buf, "w") as zf:
            for i, f in enumerate(uploaded_files):
                status.text(f"Processing ({i+1}/{total}): {f.name}")
                base = os.path.splitext(f.name)[0]
                doc = fitz.open(stream=f.getvalue(), filetype="pdf")

                p_pdf, x_data, m_rows = parse_notice_document(doc, f.name)
                doc.close()

                if p_pdf:
                    zf.writestr(f"{base}/{base}_Complete_Page.pdf", p_pdf)
                if x_data:
                    zf.writestr(f"{base}/{base}_Data.xlsx", x_data)
                if m_rows:
                    all_master.extend(m_rows)

                prog.progress((i + 1) / total)

            if all_master:
                m_df = pd.DataFrame(all_master)
                zf.writestr(
                    "CRM_All_Accounts_Master.csv",
                    m_df.to_csv(index=False).encode('utf-8')
                )

        status.success("All Done!")
        st.download_button(
            label="📥 Download All Deliverables (ZIP)",
            data=zip_buf.getvalue(),
            file_name="Processed_Auctions.zip",
            mime="application/zip"
        )
        
