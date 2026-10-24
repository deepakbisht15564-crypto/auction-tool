import streamlit as st
import os
import re
import zipfile
import io
import fitz
import pandas as pd

st.set_page_config(page_title="Bulk E-Auction Notice Processor", layout="wide")
st.title("Bulk E-Auction PDF Processor")
st.write("Box Cuttings + Field-Wise Excel Extraction.")

uploaded_files = st.file_uploader("Upload Auction PDFs (Bulk)", type=["pdf"], accept_multiple_files=True)
start_btn = st.button("Start Processing")

def clean_val(text):
    if not text:
        return ""
    val = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', str(text))
    return " ".join(val.split()).strip()

def get_precise_notice_box(page):
    terms = ["sale notice", "e-auction", "sarfaesi", "arcil", "reserve price", "earnest money"]
    rects = []
    for t in terms:
        m = page.search_for(t)
        if m:
            rects.extend(m)

    if rects:
        rx = rects[0].x0
        ry = rects[0].y0

        box_list = []
        for d in page.get_drawings():
            r = d["rect"]
            if r.width > 140 and r.height > 90:
                if (r.x0 <= rx <= r.x1) and (r.y0 - 25 <= ry <= r.y1):
                    box_list.append(r)

        if box_list:
            best = max(box_list, key=lambda b: b.width * b.height)
            return fitz.Rect(
                max(0, best.x0 - 2),
                max(0, best.y0 - 2),
                min(page.rect.width, best.x1 + 2),
                min(page.rect.height, best.y1 + 2)
            )

        blks = page.get_text("blocks")
        nb = []
        for b in blks:
            bx0, by0, bx1, by1 = b[0], b[1], b[2], b[3]
            if (abs(bx0 - rx) < 90 or (bx0 <= rx <= bx1)) and by1 >= (ry - 35):
                nb.append(fitz.Rect(bx0, by0, bx1, by1))

        if nb:
            return fitz.Rect(
                max(0, min(r.x0 for r in nb) - 10),
                max(0, min(r.y0 for r in nb) - 15),
                min(page.rect.width, max(r.x1 for r in nb) + 10),
                min(page.rect.height, max(r.y1 for r in nb) + 20)
            )

    return fitz.Rect(
        page.rect.width * 0.35,
        page.rect.height * 0.50,
        page.rect.width,
        page.rect.height
    )

def parse_notice_data(full_text, filename):
    lines = [clean_val(l) for l in full_text.split("\n")]
    lines = [l for l in lines if l]

    # 1. Bank / ARC
    bank_name = ""
    for line in lines[:30]:
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
            bank_name = u.split("BRANCH")[0].split("NOTICE")[0].strip()
            if len(bank_name) > 3:
                break

    if not bank_name:
        bank_name = "ARCIL" if "arcil" in full_text.lower() else "BANK NOTICE"
    bank_name = bank_name.upper()

    # 2. Borrower Name
    borrower_name = ""
    cfn = filename.replace("_Cutting", "").replace(".pdf", "")
    tokens = [w for w in cfn.split() if not any(c.isdigit() for c in w)]
    filter_words = ["THE", "NEW", "INDIAN", "EXPRESS", "TIMES", "DAINIK"]
    clean_tokens = [w for w in tokens if w.upper() not in filter_words]
    if clean_tokens:
        b_raw = " ".join(clean_tokens)
        borrower_name = b_raw.replace("MrMrs.", "").replace("Mr.", "").replace("Mrs.", "").strip()

    if not borrower_name:
        for idx, line in enumerate(lines):
            if "borrower" in line.lower():
                borrower_name = clean_val(line.split(":")[-1])
                break

    # 3. Reserve Price
    reserve_price = ""
    for line in lines:
        if "reserve" in line.lower() and ("rs" in line.lower() or "₹" in line):
            reserve_price = clean_val(line.split(":")[-1])
            break
    if not reserve_price:
        amounts = re.findall(r'(?:Rs\.?|₹)\s*[\d,]+(?:\.\d+)?', full_text)
        if amounts:
            reserve_price = clean_val(amounts[0])

    # 4. EMD Amount
    emd_amount = ""
    for line in lines:
        if "emd" in line.lower() and ("rs" in line.lower() or "₹" in line):
            emd_amount = clean_val(line.split(":")[-1])
            break

    # 5. Dates & Times
    all_d = re.findall(r'\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4}', full_text)
    all_t = re.findall(r'\d{1,2}[:\.]\d{2}\s*(?:AM|PM|hrs|hours)?', full_text, re.I)

    emd_sub_d = all_d[0] if len(all_d) > 0 else ""
    auc_d = all_d[1] if len(all_d) > 1 else emd_sub_d

    emd_sub_t = clean_val(all_t[0]) if len(all_t) > 0 else "05:00 PM"
    auc_st_t = clean_val(all_t[1]) if len(all_t) > 1 else "11:00 AM"
    auc_end_t = clean_val(all_t[2]) if len(all_t) > 2 else "12:00 PM"

    # 6. Property Description
    prop_desc = ""
    for idx, line in enumerate(lines):
        if "property" in line.lower() or "schedule" in line.lower():
            prop_desc = " ".join(lines[idx:idx+4])
            break
    if not prop_desc:
        prop_desc = " ".join(lines[:6])

    row = {
        "Notice File": filename,
        "Bank / ARC": bank_name,
        "Borrower Name": borrower_name,
        "Property Description": clean_val(prop_desc),
        "Reserve Price": reserve_price,
        "EMD Amount": emd_amount,
        "EMD Submit Date": emd_sub_d,
        "EMD Submit Time": emd_sub_t,
        "Auction Date": auc_d,
        "Auction Start Time": auc_st_t,
        "Auction End Date": auc_d,
        "Auction End Time": auc_end_t
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
