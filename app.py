import streamlit as st
import os
import re
import zipfile
import io
import fitz  # PyMuPDF
import pandas as pd

st.set_page_config(page_title="Bulk E-Auction Notice Processor", layout="wide")
st.title("📄 Bulk E-Auction PDF Processor")
st.write("Precise Box Cuttings + Field-Wise Excel Extraction.")

def clean_val(text):
    if not text:
        return ""
    cleaned = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', str(text))
    return re.sub(r'[\r\n\t]+', ' ', cleaned).strip()

def get_precise_notice_box(page):
    # Anchor words specifically indicating an auction notice
    core_anchors = ["sale notice", "e-auction notice", "sarfaesi", "arcil", "asset reconstruction", "possession notice"]
    
    anchor_rects = []
    for term in core_anchors:
        matches = page.search_for(term)
        if matches:
            anchor_rects.extend(matches)
            
    # Agar anchor text milta hai to usi section ke boundaries check karein
    if anchor_rects:
        # Notice ka anchor point
        ref_x = anchor_rects[0].x0
        ref_y = anchor_rects[0].y0
        
        # 1. Newspaper me notice border (Rectangles / Drawings) khojein
        drawings = page.get_drawings()
        candidate_boxes = []
        for d in drawings:
            r = d["rect"]
            # Check if this box contains our auction notice anchor
            if r.width > 150 and r.height > 100:
                if r.x0 <= ref_x <= r.x1 and r.y0 - 20 <= ref_
