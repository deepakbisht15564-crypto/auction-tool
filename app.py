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
st.write("Multi-Notice Aware Engine: Auto Detects Each Bank Notice & Creates Separate Excel Sheets.")

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
    return re.sub(r'[^a-zA-Z0-9\s.,\-\/\(\)&:;]', ' ', c).strip()

def clean_amount(val_str):
    if not val_str:
        return ""
    cleaned = re.sub(r'(?i)(rs\.?|inr|₹|\/-)', '', str(val_str))
    cleaned = re.sub(r'[^\d,.]', '', cleaned).strip('., ')
    return cleaned

def identify_bank_name_from_box(box_text):
    lines = [clean_val(l) for l in box_text.split("\n") if clean_val(l)]
    for line in lines[:12]:
        u = line.upper()
        if "ARCIL" in u or "ASSET RECONSTRUCTION" in u:
            return "ARCIL"
        elif "ICICI" in u:
            return "ICICI BANK"
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
        elif "AXIS" in u:
            return "AXIS BANK"
        elif "KOTAK" in u:
            return "KOTAK MAHINDRA BANK"
        elif "BANK" in u or "FINANCE" in u:
            cand = u.split("BRANCH")[0].split("NOTICE")[0].split("E-AUCTION")[0].strip()
            if len(cand) > 4:
                return cand[:25]
    u_full = box_text.upper()
    if "ICICI" in u_full:
        return "ICICI BANK"
    elif "ARCIL" in u_full:
        return "ARCIL"
    return "BANK NOTICE"

def split_page_into_notices(page):
    """
    Akhbar ke panne par jitne alag-alag notices hain unke alag-alag bboxes banata hai
    taaki ek notice ka data doosre me mix na ho.
    """
    anchor_keywords = [
        "arcil", "icici bank", "asset reconstruction", "state bank", "canara bank",
        "punjab national", "bank of baroda", "indian bank", "sale notice", "e-auction notice"
    ]
    
    found_headers = []
    for kw in anchor_keywords:
        rects = page.search_for(kw)
        for r in rects:
            # Check karein agar header top 85% area me ho
            if r.y0 < page.rect.height * 0.85:
                # Merge nearby header duplicates
                is_duplicate = False
                for existing in found_headers:
                    if abs(existing.y0 - r.y0) < 40 and abs(existing.x0 - r.x0) < 80:
                        is_duplicate = True
                        break
                if not is_duplicate:
                    found_headers.append(r)
    
    # Header ko vertical position ke mutabiq sort karein
    found_headers.sort(key=lambda r: (r.y0, r.x0))
    
    # Agar 1 se zyada headers hain toh alag-alag boxes banayein
    notice_boxes = []
    if len(found_headers) > 1:
        # Group vertically
        y_splits = [h.y0 for h in found_headers]
        y_splits.append(page.rect.height)
        
        for idx in range(len(found_headers)):
            top = max(0, found_headers[idx].y0 - 25)
            bottom = min(page.rect.height, y_splits[idx + 1] - 10)
            # Minimum readable notice height
            if bottom - top > 80:
                notice_boxes.append(fitz.Rect(0, top, page.rect.width, bottom))
    else:
        notice_boxes.append(page.rect)
        
    return notice_boxes

def parse_notice_box(page, box_rect, default_filename):
    box_text = page.get_text("text", clip=box_rect)
    bank_name = identify_bank_name_from_box(box_text)

    # Dates extraction inside this box
    dates = re.findall(r'\b\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4}\b', box_text)
    times = re.findall(r'\b\d{1,2}[:\.]\d{2}\s*(?:AM|PM|am|pm|hrs|hours)?\b', box_text)

    emd_match = re.search(r'(?is)(?:last\s*date|submission\s*of\s*emd|emd\s*last)[^\n\.\;]*?(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4})', box_text)
    auc_match = re.search(r'(?is)(?:date\s*of\s*auction|auction\s*date|e-auction\s*date)[^\n\.\;]*?(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4})', box_text)

    emd_d = emd_match.group(1) if emd_match else (dates[0] if len(dates) > 0 else "")
    auc_d = auc_match.group(1) if auc_match else (dates[1] if len(dates) > 1 else emd_d)
    emd_t = times[0] if len(times) > 0 else "05:00 PM"
    auc_t = times[1] if len(times) > 1 else "11:00 AM - 13:00 PM"

    insp_match = re.search(r'(?is)(?:inspection|inspect)[^\n\.\;]*?(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4})', box_text)
    insp_d = insp_match.group(1) if insp_match else "As per notice terms"

    # PyMuPDF Table Extraction inside this specific notice
    extracted_rows = []
    try:
        tabs = page.find_tables(clip=box_rect)
        if tabs.tables:
            for tab in tabs.tables:
                grid = tab.extract()
                if not grid or len(grid) < 2:
                    continue

                c_borrower, c_prop, c_rp, c_emd, c_auc = -1, -1, -1, -1, -1
                for ci, col_txt in enumerate(grid[0]):
                    h = str(col_txt or "").lower()
                    if any(w in h for w in ["borrower", "guarantor", "party", "name"]):
                        c_borrower = ci
                    elif any(w in h for w in ["property", "description", "asset", "schedule"]):
                        c_prop = ci
                    elif any(w in h for w in ["reserve", "price", "rp"]):
                        c_rp = ci
                    elif any(w in h for w in ["emd", "earnest"]):
                        c_emd = ci
                    elif any(w in h for w in ["auction", "date"]):
                        c_auc = ci

                for r in grid[1:]:
                    b_raw = str(r[c_borrower]) if c_borrower != -1 and c_borrower < len(r) and r[c_borrower] else ""
                    p_raw = str(r[c_prop]) if c_prop != -1 and c_prop < len(r) and r[c_prop] else ""
                    rp_raw = str(r[c_rp]) if c_rp != -1 and c_rp < len(r) and r[c_rp] else ""
                    emd_raw = str(r[c_emd]) if c_emd != -1 and c_emd < len(r) and r[c_emd] else ""
                    a_raw = str(r[c_auc]) if c_auc != -1 and c_auc < len(r) and r[c_auc] else ""

                    if p_raw or rp_raw or b_raw:
                        extracted_rows.append({
                            "borrower": clean_val(b_raw),
                            "property": clean_prop(p_raw),
                            "rp": clean_amount(rp_raw),
                            "emd": clean_amount(emd_raw),
                            "auc": clean_val(a_raw) if a_raw else f"{auc_d} {auc_t}".strip()
                        })
    except Exception:
        pass

    # Fallback if no table grid in this box
    if not extracted_rows:
        b_m = re.search(r'(?i)(?:borrower|name\s*of\s*borrower)[\s:\-]+([^\n,]+)', box_text)
        borrower_name = clean_val(b_m.group(1)) if b_m else ""
        if not borrower_name or len(borrower_name) < 3:
            clean_fn = default_filename.replace("_Cutting", "").replace(".pdf", "").replace("-", " ")
            tokens = [w for w in clean_fn.split() if not any(c.isdigit() for c in w)]
            filter_words = ["THE", "NEW", "INDIAN", "EXPRESS", "TIMES", "DAINIK", "NOTICE", "E-AUCTION", "SALE", "AUCTION"]
            clean_tokens = [w for w in tokens if w.upper() not in filter_words]
            borrower_name = " ".join(clean_tokens).replace("MrMrs.", "").replace("Mr.", "").replace("Mrs.", "").strip()

        prop_m = re.search(
            r'(?is)(?:Description\s+of\s+(?:the\s+)?Property|Schedule\s*[-:])\s*[:\-\n](.*?)(?=(?:Reserve\s*Price|EMD|Earnest\s*Money|Terms\s*and\s*Conditions)|$)',
            box_text
        )
        prop_text = clean_prop(prop_m.group(1)) if prop_m else clean_prop(box_text[:350])

        rp_m = re.search(r'(?is)Reserve\s*Price[^\d\n]*[:\-\s]*(?:Rs\.?|₹|INR)?\s*([\d,]+(?:\.\d+)?)', box_text)
        emd_m = re.search(r'(?is)(?:EMD\s*Amount|EMD|Earnest\s*Money)[^\d\n]*[:\-\s]*(?:Rs\.?|₹|INR)?\s*([\d,]+(?:\.\d+)?)', box_text)
        all_nums = re.findall(r'(?:Rs\.?|₹|INR)?\s*([\d,]{4,}(?:\.\d+)?)', box_text)

        res_price = clean_amount(rp_m.group(1)) if rp_m else (clean_amount(all_nums[0]) if all_nums else "")
        emd_val = clean_amount(emd_m.group(1)) if emd_m else (clean_amount(all_nums[1]) if len(all_nums) > 1 else "")

        extracted_rows.append({
            "borrower": borrower_name or "Borrower As per Notice",
            "property": prop_text,
            "rp": res_price,
            "emd": emd_val,
            "auc": f"{auc_d} {auc_t}".strip()
        })

    # Rows list create karein
    box_rows = []
    for s_idx, r in enumerate(extracted_rows, start=1):
        box_rows.append({
            "S.No.": s_idx,
            "Bank Name": bank_name,
            "Borrower(s) Name": r["borrower"],
            "Property Description": r["property"],
            "Reserve Price": r["rp"],
            "EMD Amount": r["emd"],
            "Inspection Date & Time": insp_d,
            "EMD Last Date & Time": f"{emd_d} {emd_t}".strip(),
            "Auction Date & Time": r.get("auc") or f"{auc_d} {auc_t}".strip()
        })

    return bank_name, box_rows

def style_like_clean_template(ws):
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_top_align = Alignment(horizontal="left", vertical="top", wrap_text=True)
    
    cell_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    ws.row_dimensions[1].height = 32
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = cell_border

    for row_idx in range(2, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 65
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = Font(name="Calibri", size=10)
            cell.border = cell_border
            if col_idx in [3, 4]:
                cell.alignment = left_top_align
            else:
                cell.alignment = center_align

    col_widths = {
        "A": 9.0,   # S.No.
        "B": 24.0,  # Bank Name
        "C": 32.0,  # Borrower(s) Name
        "D": 50.0,  # Property Description
        "E": 18.0,  # Reserve Price
        "F": 18.0,  # EMD Amount
        "G": 22.0,  # Inspection Date & Time
        "H": 22.0,  # EMD Last Date & Time
        "I": 22.0   # Auction Date & Time
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

def process_single_pdf(file_bytes, filename):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    t_idx = 0
    for idx in range(len(doc)):
        txt = doc[idx].get_text().lower()
        if any(w in txt for w in ["sale notice", "e-auction", "sarfaesi", "arcil", "reserve price"]):
            t_idx = idx
            break

    target_page = doc[t_idx]

    # Full broadsheet page export (vector)
    full_rect = target_page.rect
    target_page.set_cropbox(full_rect)
    target_page.set_mediabox(full_rect)

    full_doc = fitz.open()
    full_doc.insert_pdf(doc, from_page=t_idx, to_page=t_idx)
    full_doc[0].set_cropbox(full_rect)
    full_doc[0].set_mediabox(full_rect)
    pdf_bytes = full_doc.write()
    full_doc.close()

    # Detect independent notice boxes on page
    notice_boxes = split_page_into_notices(target_page)
    
    bank_sheets = {}
    master_records = []

    for b_idx, box_rect in enumerate(notice_boxes):
        bank_name, b_rows = parse_notice_box(target_page, box_rect, filename)
        if b_rows:
            # Sheet name unique ensure karein
            clean_bank_sheet = re.sub(r'[\\/*?:\[\]]', '', bank_name)[:25].strip() or "NOTICE"
            if clean_bank_sheet in bank_sheets:
                clean_bank_sheet = f"{clean_bank_sheet}_{b_idx+1}"

            bank_sheets[clean_bank_sheet] = b_rows
            for r in b_rows:
                mr = dict(r)
                mr["Source_File"] = filename
                master_records.append(mr)

    doc.close()

    # Excel with separate sheet per Bank Notice
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        for sheet_title, rows in bank_sheets.items():
            pd.DataFrame(rows).to_excel(writer, sheet_name=sheet_title, index=False)
            ws = writer.sheets[sheet_title]
            style_like_clean_template(ws)

    excel_bytes = out.getvalue()
    return pdf_bytes, excel_bytes, master_records

if start_btn:
    if not uploaded_files:
        st.warning("Please upload auction PDF files!")
    else:
        prog = st.progress(0)
        status = st.empty()

        zip_buf = io.BytesIO()
        total = len(uploaded_files)
        all_master = []

        with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for i, f in enumerate(uploaded_files):
                status.text(f"Processing ({i+1}/{total}): {f.name}")
                base = os.path.splitext(f.name)[0]

                p_pdf, x_data, m_rows = process_single_pdf(f.getvalue(), f.name)

                if p_pdf:
                    zf.writestr(f"{base}_Complete_Page.pdf", p_pdf)
                if x_data:
                    zf.writestr(f"{base}_Data.xlsx", x_data)
                if m_rows:
                    all_master.extend(m_rows)

                prog.progress((i + 1) / total)

            if all_master:
                m_df = pd.DataFrame(all_master)
                zf.writestr(
                    "CRM_All_Accounts_Master.csv",
                    m_df.to_csv(index=False).encode('utf-8')
                )

        zip_buf.seek(0)
        final_zip = zip_buf.getvalue()

        status.success("All Done!")
        st.download_button(
            label="📥 Download All Deliverables (ZIP)",
            data=final_zip,
            file_name="Processed_Auctions.zip",
            mime="application/zip"
        )
