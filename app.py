import streamlit as st
import os
import re
import zipfile
import io
import fitz
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Bulk E-Auction Notice Processor",
    layout="wide"
)
st.title("Bulk E-Auction PDF Processor")
st.write("Full Page PDF + Formatted Auto-Fit Styled Excel.")

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
    # Only alphanumeric, spaces, commas, and periods
    return re.sub(r'[^a-zA-Z0-9\s.,]', ' ', c).strip()

def clean_amount(val_str):
    if not val_str:
        return ""
    cleaned = re.sub(r'(?i)(rs\.?|inr|₹)', '', val_str)
    return clean_val(cleaned)

def extract_bank_name(lines, full_text):
    for line in lines[:45]:
        u = line.upper()
        if "ARCIL" in u or "ASSET RECONSTRUCTION" in u:
            return "ARCIL"
        elif "STATE BANK" in u or "SBI" in u:
            return "STATE BANK OF INDIA"
        elif "CANARA" in u:
            return "CANARA BANK"
        elif "PUNJAB NATIONAL" in u or "PNB" in u:
            return "PUNJAB NATIONAL BANK"
        elif "BANK OF BARODA" in u or "BOB" in u:
            return "BANK OF BARODA"
        elif "INDIAN BANK" in u:
            return "INDIAN BANK"
        elif "UNION BANK" in u:
            return "UNION BANK OF INDIA"
        elif "HDFC" in u:
            return "HDFC BANK"
        elif "ICICI" in u:
            return "ICICI BANK"
        elif "AXIS" in u:
            return "AXIS BANK"
        elif "KOTAK" in u:
            return "KOTAK MAHINDRA BANK"
        elif "BANK" in u:
            cand = u.split("BRANCH")[0].split("NOTICE")[0].strip()
            if len(cand) > 4:
                return cand
    if "arcil" in full_text.lower():
        return "ARCIL"
    return "BANK NOTICE"

def extract_borrower(filename, lines):
    clean_fn = filename.replace("_Cutting", "").replace(".pdf", "")
    tokens = [w for w in clean_fn.split() if not any(c.isdigit() for c in w)]
    filter_words = ["THE", "NEW", "INDIAN", "EXPRESS", "TIMES", "DAINIK", "NOTICE", "E-AUCTION", "SALE"]
    clean_tokens = [w for w in tokens if w.upper() not in filter_words]
    if clean_tokens:
        b_name = " ".join(clean_tokens)
        return b_name.replace("MrMrs.", "").replace("Mr.", "").replace("Mrs.", "").strip()

    for idx, line in enumerate(lines):
        line_low = line.lower()
        if "borrower" in line_low:
            parts = line.split(":")
            val = parts[-1].split("/")[0].split("Co-")[0].strip()
            if len(val) > 2:
                return clean_val(val)
            elif idx + 1 < len(lines):
                return clean_val(lines[idx + 1])
    return "N/A"

def style_excel_worksheet(ws):
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    prop_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    thin_side = Side(border_style="thin", color="000000")
    all_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # Header Row
    ws.row_dimensions[1].height = 28
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = all_border

    # Data Rows
    for row_idx in range(2, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 55
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = Font(name="Calibri", size=10)
            cell.border = all_border
            # Column 4 is Property Description
            if col_idx == 4:
                cell.alignment = prop_align
            else:
                cell.alignment = center_align

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        adjusted_width = min(max(max_len + 4, 15), 50)
        ws.column_dimensions[col_letter].width = adjusted_width

def parse_notice_document(doc, filename):
    t_idx = 0
    for idx in range(len(doc)):
        txt = doc[idx].get_text().lower()
        if any(w in txt for w in ["sale notice", "e-auction", "sarfaesi", "arcil", "reserve price"]):
            t_idx = idx
            break

    target_page = doc[t_idx]

    # Full unclipped broadsheet page export
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

    # Reserve Price & EMD
    amounts = re.findall(r'(?:Rs\.?|₹|INR)\s*([\d,]+(?:\.\d+)?)', full_text, re.I)
    if not amounts:
        amounts = re.findall(r'\b\d{1,3}(?:,\d{2,3})+(?:\.\d+)?\b', full_text)

    raw_res = amounts[0] if len(amounts) >= 1 else ""
    raw_emd = amounts[1] if len(amounts) >= 2 else ""

    reserve_price = clean_amount(raw_res)
    emd_amount = clean_amount(raw_emd)

    dates = re.findall(r'\b\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4}\b', full_text)
    times = re.findall(r'\b\d{1,2}[:\.]\d{2}\s*(?:AM|PM|a\.m\.|p\.m\.|hrs|hours)?\b', full_text, re.I)

    emd_last_date = dates[0] if len(dates) > 0 else ""
    auction_date = dates[1] if len(dates) > 1 else emd_last_date

    emd_last_time = clean_val(times[0]) if len(times) > 0 else "05:00 PM"
    auction_time = clean_val(times[1]) if len(times) > 1 else "11:00 AM - 12:00 PM"

    # Property Description
    prop_desc = ""
    for idx, line in enumerate(lines):
        line_l = line.lower()
        if any(kw in line_l for kw in ["property", "schedule", "immovable", "piece and parcel", "boundaries"]):
            prop_desc = " ".join(lines[idx:idx+6])
            break
    if not prop_desc:
        mid = len(lines) // 2
        prop_desc = " ".join(lines[mid:mid+5]) if lines else "Details as per attached notice."

    # 9 Required Columns without Account Number
    row_data = {
        "S.No.": 1,
        "Bank Name": bank_name,
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
        ws = writer.sheets[clean_sheet]
        style_excel_worksheet(ws)

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
