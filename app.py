import streamlit as st
import os
import re
import zipfile
import io
import fitz
import pandas as pd

st.set_page_config(page_title="Bulk E-Auction Notice Processor", layout="wide")
st.title("Bulk E-Auction PDF Processor")
st.write("Clean Box Cuttings + Field-Wise Excel Extraction.")

uploaded_files = st.file_uploader("Upload Auction PDFs (Bulk)", type=["pdf"], accept_multiple_files=True)
start_btn = st.button("Start Processing")

def clean_val(text):
    if not text:
        return ""
    val = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', str(text))
    return " ".join(val.split()).strip()

def get_precise_notice_box(page):
    # Notice identify karne wale primary anchors
    top_anchors = ["arcil", "asset reconstruction", "sale notice", "e-auction notice", "sarfaesi"]
    bottom_anchors = ["authorised officer", "terms and conditions", "arcil.co.in", "possession", "reserve price"]

    top_rects = []
    for t in top_anchors:
        found = page.search_for(t)
        if found:
            top_rects.extend(found)

    bottom_rects = []
    for b in bottom_anchors:
        found = page.search_for(b)
        if found:
            bottom_rects.extend(found)

    if top_rects:
        # Notice box ka starting top coordinate
        y_start = min(r.y0 for r in top_rects) - 25
        x_start = min(r.x0 for r in top_rects) - 15

        if bottom_rects:
            # Notice box ke bottom/right coordinates
            valid_bottoms = [r for r in bottom_rects if r.y1 > y_start]
            if valid_bottoms:
                y_end = max(r.y1 for r in valid_bottoms) + 30
                x_end = max(r.x1 for r in valid_bottoms) + 20
            else:
                y_end = page.rect.height
                x_end = page.rect.width
        else:
            y_end = page.rect.height
            x_end = page.rect.width

        # Strictly notice boundary (prevents left/top ad clippings)
        return fitz.Rect(
            max(0, x_start),
            max(0, y_start),
            min(page.rect.width, max(x_end, x_start + 250)),
            min(page.rect.height, y_end)
        )

    # Fallback agar text coordinates match na ho
    return fitz.Rect(page.rect.width * 0.35, page.rect.height * 0.45, page.rect.width, page.rect.height)

def parse_notice_data(full_text, filename):
    lines = [clean_val(l) for l in full_text.split("\n")]
    lines = [l for l in lines if l]

    # 1. Exact Bank / ARC Name Detection
    bank_name = ""
    for line in lines[:25]:
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

    # 2. Borrower Name (Only primary borrower)
    borrower_name = ""
    clean_fn = filename.replace("_Cutting", "").replace(".pdf", "")
    tokens = [w for w in clean_fn.split() if not any(c.isdigit() for c in w)]
    filter_words = ["THE", "NEW", "INDIAN", "EXPRESS", "TIMES", "DAINIK", "NOTICE", "E-AUCTION"]
    clean_tokens = [w for w in tokens if w.upper() not in filter_words]
    if clean_tokens:
        b_raw = " ".join(clean_tokens)
        borrower_name = b_raw.replace("MrMrs.", "").replace("Mr.", "").replace("Mrs.", "").strip()

    if not borrower_name:
        for idx, line in enumerate(lines):
            if "borrower" in line.lower():
                val = line.split(":")[-1]
                val = val.split("/")[0].split("Co-")[0].strip()
                if len(val) > 2:
                    borrower_name = clean_val(val)
                    break

    # 3. Currency / Reserve Price & EMD
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

    cut_doc = fitz.open()
    full_text = ""

    for page_num in range(len(doc)):
        page = doc[page_num]
        txt = page.get_text()
        full_text += txt + "\n"

        check_kw = ["auction", "sale notice", "sarfaesi", "arcil"]
        if any(kw in txt.lower() for kw in check_kw):
            crop_box = get_precise_notice_box(page)
            temp = fitz.open()
            temp.insert_pdf(doc, from_page=page_num, to_page=page_num)
            temp[0].set_cropbox(crop_box)
            cut_doc.insert_pdf(temp)
            temp.close()

    if len(cut_doc) == 0 and len(doc) > 0:
        crop_box = get_precise_notice_box(doc[0])
        temp = fitz.open()
        temp.insert_pdf(doc, from_page=0, to_page=0)
        temp[0].set_cropbox(crop_box)
        cut_doc.insert_pdf(temp)
        temp.close()

    cut_pdf_bytes = cut_doc.write()
    cut_doc.close()
    doc.close()

    bank_name, rows = parse_notice_data(full_text, filename)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        s_name = re.sub(r'[\\/*?:\[\]]', '', bank_name)[:30] or "Notice Data"
        pd.DataFrame(rows).to_excel(writer, sheet_name=s_name, index=False)
    excel_bytes = out.getvalue()

    return base_name, cut_pdf_bytes, excel_bytes, rows

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
