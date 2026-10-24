import streamlit as st
import os
import re
import zipfile
import io
import fitz  # PyMuPDF with table support
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Bulk E-Auction Notice Processor",
    layout="wide"
)
st.title("Bulk E-Auction PDF Processor")
st.write("Pure Python Engine: Full Vector Page + Formatted Multi-Sheet Excel.")

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
    # Strictly allow alphanumeric, spaces, commas, periods, hyphens, and slashes
    return re.sub(r'[^a-zA-Z0-9\s.,\-\/]', ' ', c).strip()

def clean_amount(val_str):
    if not val_str:
        return ""
    # Strip currency words/symbols and keep pure numbers/commas
    cleaned = re.sub(r'(?i)(rs\.?|inr|₹|\/-)', '', str(val_str))
    cleaned = re.sub(r'[^\d,.]', '', cleaned).strip('., ')
    return cleaned

def find_notice_bounding_box(page):
    anchors = [
        "arcil", "asset reconstruction", "sale notice",
        "e-auction notice", "sarfaesi", "authorised officer"
    ]
    hits = []
    for a in anchors:
        r = page.search_for(a)
        if r:
            hits.extend(r)

    if hits:
        top_y = max(0, min(h.y0 for h in hits) - 30)
        left_x = max(0, min(h.x0 for h in hits) - 20)

        bottom_terms = ["authorised officer", "terms & conditions", "terms and conditions", "statutory 30"]
        bot_hits = []
        for b in bottom_terms:
            br = page.search_for(b)
            if br:
                bot_hits.extend([p for p in br if p.y1 > top_y])

        bottom_y = min(page.rect.height, max(p.y1 for p in bot_hits) + 40) if bot_hits else page.rect.height
        right_x = min(page.rect.width, max(h.x1 for h in hits + bot_hits) + 30) if bot_hits else page.rect.width

        return fitz.Rect(left_x, top_y, max(right_x, left_x + 320), bottom_y)

    return page.rect

def extract_bank_from_header(notice_text):
    lines = [clean_val(l) for l in notice_text.split("\n") if clean_val(l)]
    for line in lines[:15]:
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
        elif "BANK" in u or "FINANCE" in u:
            cand = u.split("BRANCH")[0].split("NOTICE")[0].split("E-AUCTION")[0].strip()
            if len(cand) > 4:
                return cand[:30]

    return "ARCIL" if "arcil" in notice_text.lower() else "BANK NOTICE"

def extract_dates_from_notice(notice_text):
    dates = re.findall(r'\b\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4}\b', notice_text)
    times = re.findall(r'\b\d{1,2}[:\.]\d{2}\s*(?:AM|PM|am|pm|hrs)?\b', notice_text)

    emd_match = re.search(r'(?is)(?:last\s*date|submission\s*of\s*emd|emd\s*last)[^\n\.\;]*?(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4})', notice_text)
    auc_match = re.search(r'(?is)(?:date\s*of\s*auction|auction\s*date|e-auction\s*date)[^\n\.\;]*?(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4})', notice_text)

    emd_d = emd_match.group(1) if emd_match else (dates[0] if len(dates) > 0 else "")
    auc_d = auc_match.group(1) if auc_match else (dates[1] if len(dates) > 1 else emd_d)

    emd_t = times[0] if len(times) > 0 else "05:00 PM"
    auc_t = times[1] if len(times) > 1 else "11:00 AM - 12:00 PM"

    insp_match = re.search(r'(?is)(?:inspection|inspect)[^\n\.\;]*?(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4})', notice_text)
    insp_d = insp_match.group(1) if insp_match else "As per notice terms"

    return f"{emd_d} {emd_t}".strip(), f"{auc_d} {auc_t}".strip(), insp_d

def parse_notice_document(doc, filename):
    t_idx = 0
    for idx in range(len(doc)):
        txt = doc[idx].get_text().lower()
        if any(w in txt for w in ["sale notice", "e-auction", "sarfaesi", "arcil", "reserve price"]):
            t_idx = idx
            break

    target_page = doc[t_idx]

    # Full unclipped broadsheet page export (vector preservation)
    full_rect = target_page.rect
    target_page.set_cropbox(full_rect)
    target_page.set_mediabox(full_rect)

    full_doc = fitz.open()
    full_doc.insert_pdf(doc, from_page=t_idx, to_page=t_idx)
    full_doc[0].set_cropbox(full_rect)
    full_doc[0].set_mediabox(full_rect)
    pdf_bytes = full_doc.write()
    full_doc.close()

    # Scope strictly inside the notice box (eliminates broadsheet news)
    notice_box = find_notice_bounding_box(target_page)
    notice_text = target_page.get_text("text", clip=notice_box)

    bank_name = extract_bank_from_header(notice_text)
    emd_dt, auc_dt, insp_d = extract_dates_from_notice(notice_text)

    # 1. Attempt structured table extraction via PyMuPDF find_tables
    table_rows = []
    try:
        tabs = target_page.find_tables(clip=notice_box)
        if tabs.tables:
            for tab in tabs.tables:
                grid = tab.extract()
                if not grid or len(grid) < 2:
                    continue

                header_line = " ".join([str(c or "").lower() for c in grid[0]])
                if not any(k in header_line for k in ["borrower", "property", "asset", "reserve", "emd", "price"]):
                    continue

                c_borrower, c_prop, c_rp, c_emd, c_auc = -1, -1, -1, -1, -1
                for ci, col_txt in enumerate(grid[0]):
                    h = str(col_txt or "").lower()
                    if any(w in h for w in ["borrower", "guarantor", "party"]):
                        c_borrower = ci
                    elif any(w in h for w in ["property", "description", "asset", "schedule"]):
                        c_prop = ci
                    elif any(w in h for w in ["reserve", "price", "rp"]):
                        c_rp = ci
                    elif any(w in h for w in ["emd", "earnest"]):
                        c_emd = ci
                    elif any(w in h for w in ["auction", "date"]):
                        c_auc = ci

                for row_idx, r in enumerate(grid[1:], start=1):
                    b_raw = str(r[c_borrower]) if c_borrower != -1 and c_borrower < len(r) and r[c_borrower] else ""
                    p_raw = str(r[c_prop]) if c_prop != -1 and c_prop < len(r) and r[c_prop] else ""
                    rp_raw = str(r[c_rp]) if c_rp != -1 and c_rp < len(r) and r[c_rp] else ""
                    emd_raw = str(r[c_emd]) if c_emd != -1 and c_emd < len(r) and r[c_emd] else ""
                    a_raw = str(r[c_auc]) if c_auc != -1 and c_auc < len(r) and r[c_auc] else ""

                    if p_raw or rp_raw or b_raw:
                        table_rows.append({
                            "borrower": clean_val(b_raw),
                            "property": clean_prop(p_raw),
                            "rp": clean_amount(rp_raw),
                            "emd": clean_amount(emd_raw),
                            "auc": clean_val(a_raw) if a_raw else auc_dt
                        })
    except Exception:
        pass

    # 2. Fallback scoped text parser if grid table was not cleanly extracted
    if not table_rows:
        # Borrower from filename or scoped notice text
        clean_fn = filename.replace("_Cutting", "").replace(".pdf", "").replace("-", " ")
        tokens = [w for w in clean_fn.split() if not any(c.isdigit() for c in w)]
        filter_words = ["THE", "NEW", "INDIAN", "EXPRESS", "TIMES", "DAINIK", "NOTICE", "E-AUCTION", "SALE", "AUCTION"]
        clean_tokens = [w for w in tokens if w.upper() not in filter_words]
        borrower_name = " ".join(clean_tokens).replace("MrMrs.", "").replace("Mr.", "").replace("Mrs.", "").strip()

        if not borrower_name or len(borrower_name) < 3:
            b_m = re.search(r'(?i)(?:borrower|name\s*of\s*borrower)[\s:\-]+([^\n,]+)', notice_text)
            borrower_name = clean_val(b_m.group(1)) if b_m else "Borrower As per Notice"

        # Property Description strictly between notice markers
        prop_m = re.search(
            r'(?is)(?:Description\s+of\s+(?:the\s+)?Property|Schedule\s*[-:])\s*[:\-\n](.*?)(?=(?:Reserve\s*Price|EMD|Earnest\s*Money|Terms\s*and\s*Conditions)|$)',
            notice_text
        )
        prop_text = clean_prop(prop_m.group(1)) if prop_m else clean_prop(notice_text[:350])

        # Financials from scoped text
        rp_m = re.search(r'(?is)Reserve\s*Price[^\d\n]*[:\-\s]*(?:Rs\.?|₹|INR)?\s*([\d,]+(?:\.\d+)?)', notice_text)
        emd_m = re.search(r'(?is)(?:EMD\s*Amount|EMD|Earnest\s*Money)[^\d\n]*[:\-\s]*(?:Rs\.?|₹|INR)?\s*([\d,]+(?:\.\d+)?)', notice_text)

        all_nums = re.findall(r'(?:Rs\.?|₹|INR)?\s*([\d,]{4,}(?:\.\d+)?)', notice_text)
        res_price = clean_amount(rp_m.group(1)) if rp_m else (clean_amount(all_nums[0]) if all_nums else "")
        emd_val = clean_amount(emd_m.group(1)) if emd_m else (clean_amount(all_nums[1]) if len(all_nums) > 1 else "")

        table_rows.append({
            "borrower": borrower_name,
            "property": prop_text,
            "rp": res_price,
            "emd": emd_val,
            "auc": auc_dt
        })

    # Build final row payload
    final_sheet_rows = []
    for s_idx, row in enumerate(table_rows, start=1):
        final_sheet_rows.append({
            "S.No.": s_idx,
            "Bank Name": bank_name,
            "Borrower(s) Name": row["borrower"],
            "Property Description": row["property"],
            "Reserve Price": row["rp"],
            "EMD Amount": row["emd"],
            "Inspection Date & Time": insp_d,
            "EMD Last Date & Time": emd_dt,
            "Auction Date & Time": row.get("auc") or auc_dt
        })

    clean_sheet = re.sub(r'[\\/*?:\[\]]', '', bank_name)[:30].strip() or "NOTICE DATA"
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        pd.DataFrame(final_sheet_rows).to_excel(writer, sheet_name=clean_sheet, index=False)
        ws = writer.sheets[clean_sheet]
        style_excel_worksheet(ws)

    excel_bytes = out.getvalue()
    
    master_recs = []
    for r in final_sheet_rows:
        mr = dict(r)
        mr["Source_File"] = filename
        master_recs.append(mr)

    return pdf_bytes, excel_bytes, master_recs

def style_excel_worksheet(ws):
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000")
    )

    # Header Row
    ws.row_dimensions[1].height = 30
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border

    # Data Rows (All cells center aligned + All Borders)
    for row_idx in range(2, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 65
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = Font(name="Calibri", size=10)
            cell.alignment = center_align
            cell.border = thin_border

    # Auto-adjust column widths
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_len = len(str(cell.value or ""))
            if val_len > max_len:
                max_len = val_len
        # Clean visible width constraint
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 16), 55)

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
                doc = fitz.open(stream=f.getvalue(), filetype="pdf")

                p_pdf, x_data, m_rows = parse_notice_document(doc, f.name)
                doc.close()

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
    
