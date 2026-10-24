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
st.title("📄 Bulk E-Auction PDF Processor")
st.write("Box Cuttings + Field-Wise Excel Extraction.")

def clean_val(text):
    if not text:
        return ""
    c = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', str(text))
    return re.sub(r'\s+', ' ', c).strip()

def get_precise_notice_box(page):
    terms = [
        "sale notice", "e-auction", "sarfaesi",
        "arcil", "reserve price", "earnest money"
    ]
    rects = []
    for t in terms:
        m = page.search_for(t)
        if m:
            rects.extend(m)

    if rects:
        rx = rects[0].x0
        ry = rects[0].y0

        # Border drawing check
        box_list = []
        for d in page.get_drawings():
            r = d["rect"]
            w_ok = r.width > 140
            h_ok = r.height > 90
            in_x = (r.x0 <= rx <= r.x1)
            in_y = (r.y0 - 25 <= ry <= r.y1)
            if w_ok and h_ok and in_x and in_y:
                box_list.append(r)

        if box_list:
            best = max(box_list, key=lambda b: b.width * b.height)
            return fitz.Rect(
                max(0, best.x0 - 2),
                max(0, best.y0 - 2),
                min(page.rect.width, best.x1 + 2),
                min(page.rect.height, best.y1 + 2)
            )

        # Fallback to text blocks
        blks = page.get_text("blocks")
        nb = []
        for b in blks:
            bx0, by0, bx1, by1 = b[:4]
            same_col = abs(bx0 - rx) < 90 or (bx0 <= rx <= bx1)
            below_top = by1 >= (ry - 35)
            if same_col and below_top:
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
    lines = [clean_val(l) for l in full_text.split('\n')]
    lines = [l for l in lines if l]

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
            p = u.split("BRANCH")[0].split("NOTICE")[0]
            bank_name = p.strip()
            if len(bank_name) > 3:
                break

    if not bank_name:
        bank_name = "ARCIL" if "arcil" in full_text.lower() else "BANK NOTICE"
    bank_name = bank_name.upper()

    borrower_name = ""
    cfn = filename.replace("_Cutting", "").replace(".pdf", "")
    tokens = [
        w for w in cfn.split()
        if not re.search(r'(\d|NEW|INDIAN|EXPRESS|TIMES)', w, re.I)
    ]
    if tokens:
        b_raw = " ".join(tokens)
        borrower_name = b_raw.replace("MrMrs.", "").replace("Mr.", "").strip()

    if not borrower_name:
        bm = re.search(r'borrower[s]?[:\s\-]+([A-Za-z\s\.]+)', full_text, re.I)
        if bm:
            borrower_name = clean_val(bm.group(1).split("\n")[0])

    rpm = re.findall(
        r'(?:Reserve\s*Price|RP)[^\d₹Rs]*([₹Rs\.]*\s*[\d,]+(?:\.\d+)?)',
        full_text,
        re.I
    )
    reserve_price = clean_val(rpm[0]) if rpm else ""
    if not reserve_price:
        cm = re.findall(r'(?:Rs\.?|₹)\s*[\d,]+(?:\.\d+)?', full_text)
        reserve_price = clean_val(cm[0]) if cm else ""

    emdm = re.findall(
        r'(?:EMD|Earnest\s*Money)[^\d₹Rs]*([₹Rs\.]*\s*[\d,]+(?:\.\d+)?)',
        full_text,
        re.I
    )
    emd_amount = clean_val(emdm[0]) if emdm else ""

    all_d = re.findall(r'(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4})', full_text)
