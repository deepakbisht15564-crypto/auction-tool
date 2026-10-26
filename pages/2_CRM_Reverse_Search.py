import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
from playwright.sync_api import sync_playwright
import re
import io
from rapidfuzz import fuzz

st.set_page_config(page_title="CRM Reverse Search", layout="wide")
st.title("CRM Intelligent Reverse Search & Data Enrichment")

uploaded_file = st.file_uploader("Upload Excel File (Address / Account Name)", type=["xlsx", "xls"])

def clean_tokens(text):
    if not text or str(text).lower() == 'nan':
        return set()
    cleaned = re.sub(r'[^a-zA-Z0-9\s]', ' ', str(text)).lower()
    return set(cleaned.split())

def calculate_match_score(excel_row, crm_candidate):
    total_score = 0.0

    # 1. Shop / Unit / Floor Match (40% Weight)
    ex_tokens = clean_tokens(excel_row.get("Address", ""))
    crm_tokens = clean_tokens(crm_candidate.get("crm_address", ""))
    if ex_tokens and crm_tokens:
        overlap = ex_tokens.intersection(crm_tokens)
        ratio = len(overlap) / max(len(ex_tokens), 1)
        total_score += 40.0 * min(ratio * 1.8, 1.0)

    # 2. Account Name Match (30% Weight)
    name_ex = str(excel_row.get("Account_Name", "")).strip()
    name_crm = str(crm_candidate.get("crm_name", "")).strip()
    if name_ex and name_crm and name_ex.lower() != 'nan':
        total_score += 30.0 * (fuzz.token_sort_ratio(name_ex, name_crm) / 100.0)

    # 3. Locality / Mall Match (20% Weight)
    addr_ex = str(excel_row.get("Address", "")).strip()
    addr_crm = str(crm_candidate.get("crm_address", "")).strip()
    if addr_ex and addr_crm and addr_ex.lower() != 'nan':
        total_score += 20.0 * (fuzz.partial_ratio(addr_ex.lower(), addr_crm.lower()) / 100.0)

    # 4. RP / Amount Match (10% Weight)
    rp_ex = re.sub(r'[^\d]', '', str(excel_row.get("RP", "")))
    rp_crm = re.sub(r'[^\d]', '', str(crm_candidate.get("crm_rp", "")))
    if rp_ex and rp_crm and rp_ex == rp_crm and rp_ex != "":
        total_score += 10.0

    return round(min(total_score, 100.0), 1)

def get_color_for_score(score):
    if score >= 90: return "27AE60"
    elif score >= 80: return "2ECC71"
    elif score >= 70: return "82E0AA"
    elif score >= 60: return "D4EFDF"
    elif score >= 50: return "F1C40F"
    elif score >= 40: return "F39C12"
    elif score >= 30: return "E67E22"
    elif score >= 20: return "EDBB99"
    elif score >= 10: return "E74C3C"
    else: return "C0392B"

def extract_search_terms(row):
    terms = []
    name = str(row.get("Account_Name", "")).strip()
    if name and name.lower() != 'nan':
        terms.append(name)

    addr = str(row.get("Address", "")).strip()
    if addr and addr.lower() != 'nan':
        terms.append(addr[:45])
        parts = [p.strip() for p in re.split(r'[,.\-/]', addr) if len(p.strip()) > 3]
        if len(parts) >= 2:
            terms.append(f"{parts[0]} {parts[1]}")
        for part in parts:
            if any(k in part.lower() for k in ["mall", "plaza", "sector", "sec", "road", "vihar", "complex"]):
                terms.append(part)

    seen = set()
    return [t for t in terms if not (t in seen or seen.add(t))]

def run_reverse_search(df, progress_bar, status_box):
    df["Found_Account_ID"] = ""
    df["Found_Account_Name"] = ""
    df["Found_CRM_Address"] = ""
    df["Found_RP"] = ""
    df["Match_Score"] = 0.0

    proxy_config = None
    if "webshare" in st.secrets:
        proxy_config = {
            "server": st.secrets["webshare"]["proxy_server"],
            "username": st.secrets["webshare"]["proxy_username"],
            "password": st.secrets["webshare"]["proxy_password"]
        }

    with sync_playwright() as p:
        launch_args = {"headless": True, "args": ["--no-sandbox", "--disable-dev-shm-usage"]}
        if proxy_config:
            launch_args["proxy"] = proxy_config

        browser = p.chromium.launch(**launch_args)
        page = browser.new_page()

        status_box.info("CRM connect ho raha hai...")
        page.goto("https://crm.auctionbaba.com", wait_until="networkidle")

        total = len(df)
        for idx in range(total):
            row_dict = df.iloc[idx].to_dict()
            search_queries = extract_search_terms(row_dict)

            best_candidate = None
            max_score = 0.0

            for q in search_queries:
                try:
                    search_box = page.locator("#crm-search-box") # CRM search input selector
                    search_box.fill("")
                    search_box.fill(q)
                    page.keyboard.press("Enter")
                    page.wait_for_timeout(1200)

                    rows = page.locator("table tbody tr")
                    count = rows.count()
                    if count == 0:
                        continue

                    for r in range(min(count, 8)):
                        tr = rows.nth(r)
                        cand = {
                            "crm_id": tr.locator("td:nth-child(1)").inner_text().strip(),
                            "crm_name": tr.locator("td:nth-child(2)").inner_text().strip(),
                            "crm_address": tr.locator("td:nth-child(3)").inner_text().strip(),
                            "crm_rp": tr.locator("td:nth-child(4)").inner_text().strip(),
                        }
                        score = calculate_match_score(row_dict, cand)
                        if score > max_score:
                            max_score = score
                            best_candidate = cand

                    if max_score >= 85.0:
                        break
                except Exception:
                    continue

            if best_candidate:
                df.at[idx, "Found_Account_ID"] = best_candidate["crm_id"]
                df.at[idx, "Found_Account_Name"] = best_candidate["crm_name"]
                df.at[idx, "Found_CRM_Address"] = best_candidate["crm_address"]
                df.at[idx, "Found_RP"] = best_candidate["crm_rp"]
                df.at[idx, "Match_Score"] = max_score

            progress_bar.progress((idx + 1) / total)
            status_box.text(f"Processed: {idx + 1}/{total} | Score: {max_score}%")

        browser.close()
    return df

def generate_styled_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    wb = openpyxl.load_workbook(output)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]
    score_col = headers.index("Match_Score") + 1

    for r in range(2, ws.max_row + 1):
        try:
            score = float(ws.cell(row=r, column=score_col).value)
        except:
            score = 0.0

        hex_c = get_color_for_score(score)
        fill = PatternFill(start_color=hex_c, end_color=hex_c, fill_type="solid")
        font = Font(color="FFFFFF" if score >= 80 or score < 30 else "000000", bold=True)

        cell = ws.cell(row=r, column=score_col)
        cell.fill = fill
        cell.font = font

    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file)
    st.write(f"Total Rows: **{len(df_raw)}**")
    if st.button("Start Auto Search & Enrichment", type="primary"):
        prog = st.progress(0)
        status = st.empty()
        enriched_df = run_reverse_search(df_raw, prog, status)
        styled_excel = generate_styled_excel(enriched_df)
        st.success("Search Complete!")
        st.download_button("Download Enriched Excel", styled_excel, "CRM_Search_Enriched.xlsx")
