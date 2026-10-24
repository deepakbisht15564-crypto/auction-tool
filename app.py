import os
import re
import io
import time
import json
import zipfile
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
import fitz  # PyMuPDF
import cv2
import numpy as np
import pandas as pd
from PIL import Image
from pydantic import BaseModel, Field
from google import genai
from google.genai import types

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

# -------------------------------------------------------------------------
# SCHEMA DEFINITIONS
# -------------------------------------------------------------------------
class AuctionAccount(BaseModel):
    bank_or_arc_name: str = Field(description="Name of Bank or ARC e.g. Arcil, SBI, HDFC, Canara Bank")
    account_prospect_no: str = Field(description="Account number, loan number, or prospect reference")
    borrower_name: str = Field(description="Primary Borrower Name ONLY. Exclude co-borrowers, guarantors, mortgagors.")
    property_description: str = Field(description="Detailed property schedule and description")
    reserve_price: str = Field(description="Reserve price with currency")
    emd_amount: str = Field(description="EMD amount with currency")
    inspection_date_time: str = Field(description="Inspection Date and Time window")
    emd_last_date_time: str = Field(description="EMD submission deadline date and time")
    auction_date_time: str = Field(description="Auction bidding date and time window")

class AuctionNoticeExtraction(BaseModel):
    notices: list[AuctionAccount]

# -------------------------------------------------------------------------
# DATA SANITIZATION & UTILITIES
# -------------------------------------------------------------------------
def sanitize_for_excel(text: str) -> str:
    """Removes XML 1.0 illegal control characters and strips redundant whitespace."""
    if not text:
        return ""
    cleaned = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', str(text))
    return re.sub(r'\s+', ' ', cleaned).strip()

def sanitize_property_description(text: str) -> str:
    """Strictly allows only alphanumeric characters, spaces, full stops, and commas."""
    if not text:
        return ""
    cleaned = sanitize_for_excel(text)
    return re.sub(r'[^a-zA-Z0-9\s.,]', ' ', cleaned).strip()

# -------------------------------------------------------------------------
# COMPUTER VISION & SPATIAL CROPPING
# -------------------------------------------------------------------------
def detect_notice_bounding_box(page: fitz.Page) -> fitz.Rect:
    """
    Combines morphological line detection with keyword spatial anchors
    to isolate the exact auction notice bounding box.
    """
    pix = page.get_pixmap(dpi=150)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
    if pix.n >= 4:
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    _, thresh = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)
    
    scale_w = max(10, pix.w // 35)
    scale_h = max(10, pix.h // 35)
    kernel_h = cv2.getStructuringElement(cv2.MORPH_RECT, (scale_w, 1))
    kernel_v = cv2.getStructuringElement(cv2.MORPH_RECT, (1, scale_h))
    
    line_h = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel_h)
    line_v = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel_v)
    table_mask = cv2.add(line_h, line_v)
    
    contours, _ = cv2.findContours(table_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    scale_factor_x = page.rect.width / pix.w
    scale_factor_y = page.rect.height / pix.h
    
    boxes = []
    min_area = (pix.w * pix.h) * 0.04
    for c in contours:
        x, y, w, h = cv2.boundingRect(c)
        if w * h >= min_area:
            boxes.append(fitz.Rect(
                x * scale_factor_x,
                y * scale_factor_y,
                (x + w) * scale_factor_x,
                (y + h) * scale_factor_y
            ))
            
    anchors = ["sale notice", "e-auction", "sarfaesi", "arcil", "reserve price", "authorised officer"]
    anchor_points = []
    for a in anchors:
        hits = page.search_for(a)
        if hits:
            anchor_points.extend(hits)
            
    if anchor_points and boxes:
        best_box = None
        max_hits = -1
        for b in boxes:
            hits = sum(1 for p in anchor_points if b.contains(p))
            if hits > max_hits:
                max_hits = hits
                best_box = b
        if best_box and max_hits > 0:
            margin = 5
            return fitz.Rect(
                max(0, best_box.x0 - margin),
                max(0, best_box.y0 - margin),
                min(page.rect.width, best_box.x1 + margin),
                min(page.rect.height, best_box.y1 + margin)
            )

    if anchor_points:
        return fitz.Rect(
            max(0, min(r.x0 for r in anchor_points) - 15),
            max(0, min(r.y0 for r in anchor_points) - 25),
            min(page.rect.width, max(r.x1 for r in anchor_points) + 15),
            min(page.rect.height, max(r.y1 for r in anchor_points) + 35)
        )
        
    return page.rect

def generate_vector_cutting(doc: fitz.Document, target_page_idx: int, crop_box: fitz.Rect) -> bytes:
    """Creates a standalone vector PDF containing the strictly cropped notice."""
    new_doc = fitz.open()
    new_doc.insert_pdf(doc, from_page=target_page_idx, to_page=target_page_idx)
    new_doc[0].set_cropbox(crop_box)
    pdf_bytes = new_doc.write()
    new_doc.close()
    return pdf_bytes

# -------------------------------------------------------------------------
# VISION-LANGUAGE MODEL EXTRACTION
# -------------------------------------------------------------------------
def extract_structured_data(client: genai.Client, page: fitz.Page, crop_box: fitz.Rect) -> list[dict]:
    """Renders the cropped region at high resolution and extracts structured accounts."""
    pix = page.get_pixmap(clip=crop_box, dpi=200)
    image = Image.open(io.BytesIO(pix.tobytes("png")))
    
    prompt = """
    Extract all E-Auction Sale Notices published by any Bank or ARC from this image.
    Strict Instructions:
    1. Identify all auction rows/accounts accurately.
    2. Primary Borrower Name ONLY. Exclude Co-Borrowers, Guarantors, and Mortgagors.
    3. Property Description must contain full physical schedule details.
    4. Cleanly map Reserve Price, EMD, and all Dates/Times.
    """
    
    last_exception = None
    for attempt in range(4):
        try:
            response = client.models.generate_content(
                model="gemini-3.8-flash",
                contents=[prompt, image],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=AuctionNoticeExtraction,
                    temperature=0.0
                )
            )
            data = json.loads(response.text)
            return data.get("notices", [])
        except Exception as e:
            last_exception = e
            wait_time = (2 ** attempt) + 1
            logging.warning(f"Inference attempt {attempt + 1} failed: {e}. Retrying in {wait_time}s...")
            time.sleep(wait_time)
            
    logging.error(f"Inference failed after 4 attempts: {last_exception}")
    return []

# -------------------------------------------------------------------------
# WORKBOOK & EXCEL SYNTHESIS
# -------------------------------------------------------------------------
def generate_excel_deliverable(records: list[dict]) -> bytes:
    """Partitions extracted notices into Bank/ARC-specific sheets with standard 9 columns."""
    bank_groups = {}
    for idx, r in enumerate(records, start=1):
        bank = sanitize_for_excel(r.get("bank_or_arc_name", "UNKNOWN")).upper()
        sheet_name = re.sub(r'[\\/*?:\[\]]', '', bank)[:30].strip() or "General"
        
        row_payload = {
            "S.No.": idx,
            "Account / Prospect No.": sanitize_for_excel(r.get("account_prospect_no", "")),
            "Borrower(s) Name": sanitize_for_excel(r.get("borrower_name", "")),
            "Property Description": sanitize_property_description(r.get("property_description", "")),
            "Reserve Price": sanitize_for_excel(r.get("reserve_price", "")),
            "EMD Amount": sanitize_for_excel(r.get("emd_amount", "")),
            "Inspection Date & Time": sanitize_for_excel(r.get("inspection_date_time", "")),
            "EMD Last Date & Time": sanitize_for_excel(r.get("emd_last_date_time", "")),
            "Auction Date & Time": sanitize_for_excel(r.get("auction_date_time", ""))
        }
        bank_groups.setdefault(sheet_name, []).append(row_payload)
        
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        for sheet, rows in bank_groups.items():
            df = pd.DataFrame(rows)
            df.to_excel(writer, sheet_name=sheet, index=False)
            
    return out.getvalue()

# -------------------------------------------------------------------------
# PIPELINE EXECUTION ENGINE
# -------------------------------------------------------------------------
def process_single_pdf(file_path: str, client: genai.Client) -> tuple[str, bytes, bytes, list[dict]]:
    """Isolated thread worker handling a single PDF end-to-end."""
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    doc = fitz.open(file_path)
    
    try:
        target_page_idx = 0
        for idx in range(len(doc)):
            text = doc[idx].get_text().lower()
            if any(k in text for k in ["auction", "sale notice", "sarfaesi", "arcil"]):
                target_page_idx = idx
                break
                
        page = doc[target_page_idx]
        crop_box = detect_notice_bounding_box(page)
        
        # 1. Vector PDF Cutting
        cut_pdf_bytes = generate_vector_cutting(doc, target_page_idx, crop_box)
        
        # 2. Vision Inference
        extracted_notices = extract_structured_data(client, page, crop_box)
        
        # 3. Dedicated Excel Generation
        excel_bytes = generate_excel_deliverable(extracted_notices)
        
        # Aggregate records for Master CSV export
        master_records = []
        for n in extracted_notices:
            rec = {
                "Source_PDF": base_name,
                "Bank_or_ARC": sanitize_for_excel(n.get("bank_or_arc_name", "")),
                "Account_No": sanitize_for_excel(n.get("account_prospect_no", "")),
                "Borrower_Name": sanitize_for_excel(n.get("borrower_name", "")),
                "Property_Description": sanitize_property_description(n.get("property_description", "")),
                "Reserve_Price": sanitize_for_excel(n.get("reserve_price", "")),
                "EMD_Amount": sanitize_for_excel(n.get("emd_amount", "")),
                "Inspection_Date_Time": sanitize_for_excel(n.get("inspection_date_time", "")),
                "EMD_Last_Date_Time": sanitize_for_excel(n.get("emd_last_date_time", "")),
                "Auction_Date_Time": sanitize_for_excel(n.get("auction_date_time", ""))
            }
            master_records.append(rec)
            
        return base_name, cut_pdf_bytes, excel_bytes, master_records
    finally:
        doc.close()

def execute_bulk_processing(input_dir: str, output_zip_path: str, api_key: str, max_workers: int = 8):
    """Orchestrates multi-threaded bulk processing across all PDFs in the input directory."""
    client = genai.Client(api_key=api_key)
    pdf_files = [os.path.join(input_dir, f) for f in os.listdir(input_dir) if f.lower().endswith(".pdf")]
    
    logging.info(f"Initiating bulk pipeline. Found {len(pdf_files)} PDF documents.")
    global_master_ledger = []
    
    with zipfile.ZipFile(output_zip_path, "w", compression=zipfile.ZIP_DEFLATED) as master_zip:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_pdf = {executor.submit(process_single_pdf, pdf, client): pdf for pdf in pdf_files}
            
            for future in as_completed(future_to_pdf):
                pdf_path = future_to_pdf[future]
                try:
                    base_name, cut_pdf, excel_data, master_recs = future.result()
                    
                    # Store segregated files in distinct folders per PDF name
                    master_zip.writestr(f"{base_name}/{base_name}_Cutting.pdf", cut_pdf)
                    master_zip.writestr(f"{base_name}/{base_name}_Data.xlsx", excel_data)
                    global_master_ledger.extend(master_recs)
                    
                    logging.info(f"Successfully processed: {base_name}")
                except Exception as e:
                    logging.error(f"Failed processing {pdf_path}: {e}")
                    
        # Append Global CRM Consolidated Master Sheet
        if global_master_ledger:
            master_df = pd.DataFrame(global_master_ledger)
            master_zip.writestr("CRM_All_Accounts_Master.csv", master_df.to_csv(index=False).encode('utf-8'))
            
    logging.info(f"Bulk execution finalized. Deliverables packaged to {output_zip_path}")

if __name__ == "__main__":
    API_KEY = os.environ.get("GEMINI_API_KEY", "")
    INPUT_FOLDER = "./input_auctions"
    OUTPUT_PACKAGE = "./Bulk_Auction_Deliverables.zip"
    
    if API_KEY:
        execute_bulk_processing(INPUT_FOLDER, OUTPUT_PACKAGE, API_KEY, max_workers=8)
    else:
        print("Please provide a valid GEMINI_API_KEY environment variable.")
        
