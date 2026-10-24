import io, re, os, tempfile
from datetime import datetime
import streamlit as st
import pandas as pd
from PIL import Image, ImageOps, ImageEnhance, ImageFilter
import pytesseract
from pdf2image import convert_from_bytes
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Bank E-Auction Notice Extractor", page_icon="📄", layout="wide")
st.title("📄 Bank E-Auction Notice Extractor")
st.caption("Upload newspaper cuttings, photos or PDFs → OCR → review → Excel export. Free, local-first starter version.")

with st.expander("Important: accuracy and privacy", expanded=True):
    st.write(
        "OCR can misread small, blurred or low-resolution print. This app keeps the extracted text and original file name "
        "for checking; always compare names, amounts and dates against the notice before relying on them. "
        "The app runs on your computer when started locally. Files are processed by this app; do not expose it to the public internet."
    )

with st.sidebar:
    st.header("OCR settings")
    lang = st.selectbox("Notice language", ["English", "English + Hindi"], index=0)
    ocr_lang = "eng+hin" if lang == "English + Hindi" else "eng"
    psm = st.selectbox("OCR layout mode", ["Automatic (recommended)", "Single column", "Single block"], index=0)
    psm_num = {"Automatic (recommended)": 3, "Single column": 4, "Single block": 6}[psm]
    scale = st.selectbox("Image enlargement", ["2× (recommended)", "3×", "No enlargement"], index=0)
    scale_factor = {"2× (recommended)": 2, "3×": 3, "No enlargement": 1}[scale]

st.markdown("### 1. Upload notices")
uploads = st.file_uploader("Choose image(s) or PDF(s)", type=["png","jpg","jpeg","tif","tiff","bmp","webp","pdf"], accept_multiple_files=True)

def prep_image(im, factor=2):
    im = ImageOps.exif_transpose(im).convert("RGB")
    if factor > 1:
        im = im.resize((im.width*factor, im.height*factor), Image.Resampling.LANCZOS)
    gray = ImageOps.grayscale(im)
    gray = ImageEnhance.Contrast(gray).enhance(1.8)
    gray = gray.filter(ImageFilter.SHARPEN)
    return gray

def get_pages(upload, factor):
    data = upload.getvalue()
    name = upload.name.lower()
    if name.endswith(".pdf"):
        try:
            pages = convert_from_bytes(data, dpi=300, fmt="png")
        except Exception as e:
            raise RuntimeError("PDF read nahi hua. Windows par Poppler install karna zaroori ho sakta hai. Details: " + str(e))
        return [prep_image(p, factor) for p in pages]
    return [prep_image(Image.open(io.BytesIO(data)), factor)]

def ocr_pages(pages, lang, psm_num):
    chunks = []
    for i, page in enumerate(pages, 1):
        config = f"--oem 3 --psm {psm_num}"
        txt = pytesseract.image_to_string(page, lang=lang, config=config)
        chunks.append(f"--- PAGE {i} ---\n{txt}")
    return "\n\n".join(chunks)

def find_first(patterns, text):
    for pat in patterns:
        m = re.search(pat, text, re.I | re.M)
        if m:
            return re.sub(r"\s+", " ", m.group(1)).strip(" :.-\t")
    return ""

def money_candidates(text):
    pats = [
        r"(?:reserve\s*price|reserved\s*price|upset\s*price|reserve\s*price\s*\(rs\.?\))\s*[:\-]?\s*(?:Rs\.?|INR|₹)?\s*([0-9][0-9,]*(?:\.\d{1,2})?)",
        r"(?:EMD|earnest\s*money\s*deposit)\s*(?:amount)?\s*[:\-]?\s*(?:Rs\.?|INR|₹)?\s*([0-9][0-9,]*(?:\.\d{1,2})?)",
    ]
    reserve = find_first([pats[0]], text)
    emd = find_first([pats[1]], text)
    return reserve, emd

def extract_fields(text, filename):
    # Conservative label-based extraction; review every result before use.
    borrower = find_first([
        r"(?:name\s+of\s+the\s+borrower|borrower(?:\s+name)?|name\s+of\s+borrower)\s*[:\-]\s*([^\n\r]+)",
        r"(?:borrower)\s*[:\-]\s*([^\n\r]+)"
    ], text)
    # Do not use co-borrower/coborrower label as borrower.
    if borrower and re.search(r"co[\s\-]?borrower", borrower, re.I):
        borrower = ""
    prop = find_first([
        r"(?:description\s+of\s+(?:the\s+)?property|property\s+description|details\s+of\s+(?:the\s+)?property|description\s+of\s+secured\s+asset)\s*[:\-]?\s*([\s\S]{1,900}?)(?=\n\s*(?:reserve\s*price|emd|earnest|date\s+of\s+auction|auction\s+date|inspection|terms\s+and\s+conditions)\b|$)",
    ], text)
    bank = find_first([
        r"((?:[A-Z][A-Za-z&'.()\- ]{2,70}\s+Bank)(?:\s+Ltd\.?|\s+Limited)?)",
    ], text)
    branch = find_first([r"(?:branch|branch office)\s*[:\-]\s*([^\n\r]+)"], text)
    reserve, emd = money_candidates(text)
    emd_date = find_first([
        r"(?:last\s+date\s+for\s+(?:submission|deposit)\s+of\s+EMD|EMD\s+(?:submission|deposit)\s+date|date\s+for\s+deposit\s+of\s+EMD)\s*[:\-]?\s*([^\n\r]+)"
    ], text)
    emd_time = find_first([
        r"(?:EMD\s+(?:submission|deposit)\s+time|time\s+for\s+deposit\s+of\s+EMD|last\s+time\s+for\s+(?:submission|deposit)\s+of\s+EMD)\s*[:\-]?\s*([^\n\r]+)"
    ], text)
    auction_date = find_first([
        r"(?:date\s+(?:of\s+)?(?:e[\-\s]?auction|auction)|auction\s+date|e[\-\s]?auction\s+date)\s*[:\-]?\s*([^\n\r]+)"
    ], text)
    auction_time = find_first([
        r"(?:time\s+(?:of\s+)?(?:e[\-\s]?auction|auction)|auction\s+time|e[\-\s]?auction\s+time|auction\s+will\s+commence)\s*[:\-]?\s*([^\n\r]+)"
    ], text)
    inspection = find_first([
        r"(?:date\s+and\s+time\s+of\s+inspection|inspection\s+(?:date|schedule|time))\s*[:\-]?\s*([^\n\r]+)"
    ], text)
    notice_no = find_first([
        r"(?:notice\s*(?:no\.?|number)|reference\s*(?:no\.?|number)|loan\s*(?:account)?\s*no\.?)\s*[:\-]?\s*([A-Z0-9/\-]+)"
    ], text)
    return {
        "Source File": filename,
        "Bank Name": bank,
        "Branch": branch,
        "Borrower Name (NOT Co-borrower)": borrower,
        "Property Description (OCR text)": prop,
        "Reserve Price": reserve,
        "EMD Amount": emd,
        "EMD Date (as printed)": emd_date,
        "EMD Time (as printed)": emd_time,
        "Auction Date (as printed)": auction_date,
        "Auction Time (as printed)": auction_time,
        "Inspection Details (if stated)": inspection,
        "Notice / Reference No.": notice_no,
        "Additional Notice Details / Review Notes": "",
        "Verification Status": "Review required",
    }

if uploads:
    if st.button("🔎 Read notices", type="primary"):
        records, full_texts, errors = [], {}, []
        prog = st.progress(0)
        for idx, up in enumerate(uploads):
            try:
                pages = get_pages(up, scale_factor)
                txt = ocr_pages(pages, ocr_lang, psm_num)
                full_texts[up.name] = txt
                rec = extract_fields(txt, up.name)
                rec["Additional Notice Details / Review Notes"] = "Check full OCR text; add any notice details not captured in fields."
                records.append(rec)
            except Exception as e:
                errors.append(f"{up.name}: {e}")
            prog.progress((idx+1)/len(uploads))
        st.session_state["records"] = records
        st.session_state["full_texts"] = full_texts
        st.session_state["errors"] = errors

if "errors" in st.session_state and st.session_state["errors"]:
    st.error("Some files could not be processed:\n\n" + "\n".join(st.session_state["errors"]))

if "records" in st.session_state and st.session_state["records"]:
    st.markdown("### 2. Review and correct extracted data")
    st.warning("Borrower name, property description, money values, dates and times must be checked against the original notice. OCR extraction is a first pass, not a guarantee.")
    df = pd.DataFrame(st.session_state["records"])
    edited = st.data_editor(df, use_container_width=True, num_rows="dynamic", key="edit_table")
    st.markdown("### 3. Check original OCR text")
    for fname, txt in st.session_state["full_texts"].items():
        with st.expander(f"Extracted OCR text — {fname}"):
            st.text_area("OCR text", txt, height=280, key="ocr_"+re.sub(r'[^A-Za-z0-9]', '_', fname))
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        edited.to_excel(writer, index=False, sheet_name="Auction Notices")
        raw = pd.DataFrame([{"Source File": k, "Full OCR Text": v} for k,v in st.session_state["full_texts"].items()])
        raw.to_excel(writer, index=False, sheet_name="Full OCR Text")
        wb = writer.book
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            for col in ws.columns:
                max_len = min(max((len(str(c.value or "")) for c in col[:80]), default=10) + 2, 55)
                ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len, 14)
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
    st.download_button(
        "⬇️ Download Excel (.xlsx)",
        data=out.getvalue(),
        file_name=f"bank_eauction_notices_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
else:
    st.info("Upload one or more notices above, then click **Read notices**.")

st.markdown("---")
st.markdown("**Borrower-only rule:** the app tries to read a value attached to a Borrower label and rejects a value that appears to be labelled Co-borrower. Because bank formats vary, verify borrower names manually.")
