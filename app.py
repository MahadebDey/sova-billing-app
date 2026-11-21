import streamlit as st
import pandas as pd
from datetime import datetime
import os
import math
import re
import json
import base64
import requests
import streamlit.components.v1 as components
from reportlab.lib.pagesizes import A5
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm, inch

st.set_page_config(page_title="SOVAA JEWELLERS - Billing & Sync", layout="wide", page_icon="💎")

st.markdown("""<style>input[type="text"] { text-transform: uppercase !important; }</style>""", unsafe_allow_html=True)

HALLMARK_REG_NO = "HM/C-5800000504"
FIRM_GSTIN = "20AUJPD1127G1ZE"
FIRM_PHONE = "7717740697"
GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwyLXOi4yCglsCueHm_nVQ9wYjPUI0eKsPlxYxYGqFnwIvcML66W5JWfRCwtqj8WuhU/exec"

PDF_DIR = "Invoices_PDF"
os.makedirs(PDF_DIR, exist_ok=True)
RATES_FILE = "daily_rates.json"
GST_FILING_FILE = "GSTR1_Filing_Report.xlsx"

def load_saved_rates():
    defs = {"gold_22k": 72000.0, "gold_18k": 59000.0, "silver": 880.0}
    if os.path.exists(RATES_FILE):
        try:
            with open(RATES_FILE, "r") as f:
                d = json.load(f)
                return {"gold_22k": float(d.get("gold_22k", 72000.0)), "gold_18k": float(d.get("gold_18k", 59000.0)), "silver": float(d.get("silver", 880.0))}
        except Exception:
            return defs
    return defs

def save_rates_to_file(g22, g18, sil):
    try:
        with open(RATES_FILE, "w") as f:
            json.dump({"gold_22k": float(g22), "gold_18k": float(g18), "silver": float(sil)}, f)
    except Exception:
        pass

if "daily_gold_rate" not in st.session_state:
    rates = load_saved_rates()
    st.session_state.daily_gold_rate = rates["gold_22k"]
    st.session_state.daily_gold_18k_rate = rates["gold_18k"]
    st.session_state.daily_silver_rate = rates["silver"]

for k, v in [("c_name", ""), ("c_mob", ""), ("c_addr", "MANGO, JAMSHEDPUR"), ("c_aadhaar", ""), ("c_pan", ""), ("c_gstin", ""), ("items_list", []), ("form_reset_count", 0), ("last_looked_up_mob", ""), ("editing_item_idx", None), ("lock_bill_date", False), ("locked_date_val", datetime.today().date()), ("purity_selector", "22K (916)"), ("last_selected_preset", "-- Select Common Item --")]:
    if k not in st.session_state:
        st.session_state[k] = v

COMMON_ITEMS = ["-- Select Common Item --", "GOLD LOCKET", "SILVER LOCKET", "GOLD NATHIYA", "SILVER BABY BALA", "SILVER BABY PAYAL", "SILVER RING", "SILVER BICHHIYA", "GOLD CHAIN", "GOLD RING", "GOLD TOPS / EARING", "GOLD NOSE PIN", "GOLD PENDANT", "SILVER PAYAL", "SILVER CHAIN", "SILVER BRACELET", "SILVER COIN", "SILVER MURTI / IDOL", "➕ Custom / Other Item"]

def lookup_customer_by_mobile(mob_str):
    clean_mob = re.sub(r'\D', '', str(mob_str))
    if len(clean_mob) < 10:
        return None
    for db_file in ["Sales_Database.xlsx", "Estimate_Database.xlsx"]:
        if os.path.exists(db_file):
            try:
                df = pd.read_excel(db_file)
                if not df.empty and "Mobile No" in df.columns:
                    df["clean_m"] = df["Mobile No"].astype(str).str.replace(r'\D', '', regex=True)
                    matched = df[df["clean_m"].str.endswith(clean_mob[-10:])]
                    if not matched.empty:
                        last_rec = matched.iloc[-1]
                        raw_name = str(last_rec.get("Customer Name", ""))
                        raw_addr = str(last_rec.get("Address", "MANGO, JAMSHEDPUR"))
                        pan_val = ""
                        pan_match = re.search(r'\(PAN:\s*([A-Z0-9]+)\)', raw_name)
                        if pan_match:
                            pan_val = pan_match.group(1)
                            raw_name = re.sub(r'\s*\(PAN:\s*[A-Z0-9]+\)', '', raw_name).strip()
                        aadhaar_val = ""
                        aadh_match = re.search(r'\(Aadhaar:\s*([0-9X\s-]+)\)', raw_addr, re.IGNORECASE)
                        if aadh_match:
                            aadhaar_val = aadh_match.group(1).strip()
                            raw_addr = re.sub(r'\s*\(Aadhaar:\s*[0-9X\s-]+\)', '', raw_addr, flags=re.IGNORECASE).strip()
                        gst_val = str(last_rec.get("Party GSTIN", ""))
                        if gst_val in ["NA", "nan", "None"]:
                            gst_val = ""
                        return {"name": raw_name.upper(), "address": raw_addr.upper(), "pan": pan_val.upper(), "aadhaar": aadhaar_val, "gstin": gst_val.upper()}
            except Exception:
                pass
    return None

def get_current_fy_string():
    now = datetime.now()
    year = now.year
    s_yr, e_yr = (year - 1, year) if now.month < 4 else (year, year + 1)
    return f"{str(s_yr)[-2:]}-{str(e_yr)[-2:]}"

def get_next_number(doc_type):
    prefix = "SV" if doc_type == "Tax Invoice (GST)" else "EST"
    fy = get_current_fy_string()
    max_val = 0
    try:
        res = requests.get(GOOGLE_SCRIPT_URL, timeout=3)
        if res.status_code == 200:
            c_data = res.json()
            if c_data.get("status") == "success":
                c_num = c_data.get("last_gst" if doc_type == "Tax Invoice (GST)" else "last_est", 0)
                if c_num and int(c_num) > max_val:
                    max_val = int(c_num)
    except Exception:
        pass
    db = "Sales_Database.xlsx" if doc_type == "Tax Invoice (GST)" else "Estimate_Database.xlsx"
    if os.path.exists(db):
        try:
            df = pd.read_excel(db)
            if not df.empty and "Document No" in df.columns:
                for val in df["Document No"].dropna().astype(str):
                    v = val.strip()
                    if v.startswith(prefix) and fy in v:
                        m = re.findall(r'(\d+)$', v)
                        if m and int(m[0]) > max_val:
                            max_val = int(m[0])
        except Exception:
            pass
    return f"{prefix}-{fy}/{str(max_val + 1).zfill(3)}"

def save_to_database(row_dict, doc_type, gstr1_row=None):
    db = "Sales_Database.xlsx" if doc_type == "Tax Invoice (GST)" else "Estimate_Database.xlsx"
    try:
        if os.path.exists(db):
            df = pd.read_excel(db)
            df = df[df["Document No"].astype(str).str.strip() != str(row_dict["Document No"]).strip()]
            df = pd.concat([df, pd.DataFrame([row_dict])], ignore_index=True)
        else:
            df = pd.DataFrame([row_dict])
        df["_sort_dt"] = pd.to_datetime(df["Date"], format="%d %b %Y", errors="coerce")
        df = df.sort_values(by=["_sort_dt", "Document No"], na_position="last").drop(columns=["_sort_dt"])
        df.to_excel(db, index=False)
    except Exception:
        pass
    if doc_type == "Tax Invoice (GST)" and gstr1_row:
        try:
            if os.path.exists(GST_FILING_FILE):
                gdf = pd.read_excel(GST_FILING_FILE)
                gdf = gdf[gdf["Invoice No"].astype(str).str.strip() != str(gstr1_row["Invoice No"]).strip()]
                gdf = pd.concat([gdf, pd.DataFrame([gstr1_row])], ignore_index=True)
            else:
                gdf = pd.DataFrame([gstr1_row])
            gdf["_sort_dt"] = pd.to_datetime(gdf["Invoice Date"], format="%d-%m-%Y", errors="coerce")
            gdf = gdf.sort_values(by=["_sort_dt", "Invoice No"], na_position="last").drop(columns=["_sort_dt"])
            gdf.to_excel(GST_FILING_FILE, index=False)
        except Exception:
            pass

def delete_bill_record(doc_no_to_delete):
    doc_clean = str(doc_no_to_delete).strip()
    deleted_any = False
    for db in ["Sales_Database.xlsx", "Estimate_Database.xlsx"]:
        if os.path.exists(db):
            try:
                df = pd.read_excel(db)
                if not df.empty and "Document No" in df.columns:
                    new_df = df[df["Document No"].astype(str).str.strip() != doc_clean]
                    if len(new_df) < len(df):
                        new_df.to_excel(db, index=False)
                        deleted_any = True
            except Exception:
                pass
    if os.path.exists(GST_FILING_FILE):
        try:
            gdf = pd.read_excel(GST_FILING_FILE)
            if not gdf.empty and "Invoice No" in gdf.columns:
                new_gdf = gdf[gdf["Invoice No"].astype(str).str.strip() != doc_clean]
                if len(new_gdf) < len(gdf):
                    new_gdf.to_excel(GST_FILING_FILE, index=False)
                    deleted_any = True
        except Exception:
            pass
    p_file = os.path.join(PDF_DIR, f"{doc_clean.replace('/', '_')}.pdf")
    if os.path.exists(p_file):
        try:
            os.remove(p_file)
            deleted_any = True
        except Exception:
            pass
    try:
        requests.post(GOOGLE_SCRIPT_URL, json={"action": "delete", "doc_no": doc_clean}, timeout=5)
    except Exception:
        pass
    return deleted_any

def sync_to_google_sheet(payload):
    try:
        res = requests.post(GOOGLE_SCRIPT_URL, json=payload, headers={"Content-Type": "application/json"}, timeout=12)
        if res.status_code == 200:
            d = res.json()
            if d.get("status") == "success":
                return True, d.get("pdf_url"), "Success"
            return False, None, d.get("message", "Script Error")
        return False, None, f"HTTP Error: {res.status_code}"
    except Exception as e:
        return False, None, str(e)

def generate_a5_pdf(doc_type, meta_info, items, pdf_path):
    doc = SimpleDocTemplate(pdf_path, pagesize=A5, leftMargin=6*mm, rightMargin=6*mm, topMargin=1.25*inch, bottomMargin=4*mm)
    story = []
    styles = getSampleStyleSheet()
    num_items = len(items)
    if num_items <= 2:
        item_font, item_leading, item_pad, dynamic_bottom_gap, terms_font, terms_leading = 6.6, 8.0, 2.2, max(10*mm, (4 - num_items)*11.0*mm), 5.2, 6.4
    elif num_items <= 4:
        item_font, item_leading, item_pad, dynamic_bottom_gap, terms_font, terms_leading = 5.8, 7.0, 1.4, max(4*mm, (5 - num_items)*6.0*mm), 4.8, 6.0
    else:
        item_font, item_leading, item_pad, dynamic_bottom_gap, terms_font, terms_leading = 5.0, 6.2, 0.8, 2*mm, 4.4, 5.5

    title_style = ParagraphStyle('DocTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=12, alignment=1)
    cell_meta = ParagraphStyle('CellMeta', parent=styles['Normal'], fontName='Helvetica', fontSize=6.2, leading=7.5) 
