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

# Page Configuration
st.set_page_config(page_title="SOVAA JEWELLERS - Billing & Sync", layout="wide", page_icon="💎")

# CSS: Input boxes mein typing ko uppercase (capital) dikhane ke liye
st.markdown("""
<style>
input[type="text"] {
    text-transform: uppercase !important;
}
</style>
""", unsafe_allow_html=True)

# Business & Legal Details
HALLMARK_REG_NO = "HM/C-5800000504"
FIRM_GSTIN = "20AUJPD1127G1ZE"
FIRM_PHONE = "7717740697"
GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwyLXOi4yCglsCueHm_nVQ9wYjPUI0eKsPlxYxYGqFnwIvcML66W5JWfRCwtqj8WuhU/exec"

# Directory
PDF_DIR = "Invoices_PDF"
os.makedirs(PDF_DIR, exist_ok=True)

# Rates Storage File
RATES_FILE = "daily_rates.json"

def load_saved_rates():
    default_vals = {"gold_22k": 72000.0, "gold_18k": 59000.0, "silver": 880.0}
    if os.path.exists(RATES_FILE):
        try:
            with open(RATES_FILE, "r") as f:
                data = json.load(f)
                return {
                    "gold_22k": float(data.get("gold_22k", 72000.0)),
                    "gold_18k": float(data.get("gold_18k", 59000.0)),
                    "silver": float(data.get("silver", 880.0))
                }
        except Exception:
            return default_vals
    return default_vals

def save_rates_to_file(g22, g18, sil):
    try:
        with open(RATES_FILE, "w") as f:
            json.dump({"gold_22k": float(g22), "gold_18k": float(g18), "silver": float(sil)}, f)
    except Exception:
        pass

# Initialize Session States Safely
if "daily_gold_rate" not in st.session_state:
    rates = load_saved_rates()
    st.session_state.daily_gold_rate = rates["gold_22k"]
    st.session_state.daily_gold_18k_rate = rates["gold_18k"]
    st.session_state.daily_silver_rate = rates["silver"]

if "c_name" not in st.session_state:
    st.session_state.c_name = ""
if "c_mob" not in st.session_state:
    st.session_state.c_mob = ""
if "c_addr" not in st.session_state:
    st.session_state.c_addr = "MANGO, JAMSHEDPUR"
if "c_id" not in st.session_state:
    st.session_state.c_id = ""
if "c_pan" not in st.session_state:
    st.session_state.c_pan = ""
if "c_gstin" not in st.session_state:
    st.session_state.c_gstin = ""
if "items_list" not in st.session_state:
    st.session_state.items_list = []
if "form_reset_count" not in st.session_state:
    st.session_state.form_reset_count = 0

COMMON_ITEMS = [
    "-- Select Common Item --",
    "GOLD LOCKET",
    "SILVER LOCKET",
    "GOLD NATHIYA",
    "SILVER BABY BALA",
    "SILVER BABY PAYAL",
    "SILVER KI RING",
    "SILVER BICHHIYA",
    "GOLD CHAIN",
    "GOLD RING",
    "GOLD EARRING / JHUMKA",
    "GOLD NECKLACE",
    "GOLD BRACELET",
    "GOLD BANGLE / BALA",
    "GOLD NOSE PIN",
    "GOLD PENDANT",
    "GOLD COIN",
    "SILVER PAYAL / PAJEB",
    "SILVER CHAIN",
    "SILVER BRACELET",
    "SILVER COIN",
    "SILVER UTENSIL / BARTAN",
    "SILVER MURTI / IDOL",
    "➕ Custom / Other Item"
]

def get_current_fy_string():
    now = datetime.now()
    year = now.year
    if now.month < 4:
        start_yr = year - 1
        end_yr = year
    else:
        start_yr = year
        end_yr = year + 1
    return f"{str(start_yr)[-2:]}-{str(end_yr)[-2:]}"

def get_next_number(doc_type):
    prefix = "SV" if doc_type == "Tax Invoice (GST)" else "EST"
    current_fy = get_current_fy_string()
    max_val = 0

    try:
        res = requests.get(GOOGLE_SCRIPT_URL, timeout=3)
        if res.status_code == 200:
            cloud_data = res.json()
            if cloud_data.get("status") == "success":
                cloud_num = cloud_data.get("last_gst" if doc_type == "Tax Invoice (GST)" else "last_est", 0)
                if cloud_num and int(cloud_num) > max_val:
                    max_val = int(cloud_num)
    except Exception:
        pass

    db_file = "Sales_Database.xlsx" if doc_type == "Tax Invoice (GST)" else "Estimate_Database.xlsx"
    num_col = "Document No"
    if os.path.exists(db_file):
        try:
            df = pd.read_excel(db_file)
            if not df.empty and num_col in df.columns:
                for val in df[num_col].dropna().astype(str):
                    val_clean = val.strip()
                    if val_clean.startswith(prefix) and current_fy in val_clean:
                        matches = re.findall(r'(\d+)$', val_clean)
                        if matches:
                            num = int(matches[0])
                            if num > max_val:
                                max_val = num
        except Exception:
            pass

    next_seq = max_val + 1
    return f"{prefix}-{current_fy}/{str(next_seq).zfill(3)}"

def save_to_database(row_dict, doc_type):
    db_file = "Sales_Database.xlsx" if doc_type == "Tax Invoice (GST)" else "Estimate_Database.xlsx"
    try:
        if os.path.exists(db_file):
            df = pd.read_excel(db_file)
            df = pd.concat([df, pd.DataFrame([row_dict])], ignore_index=True)
        else:
            df = pd.DataFrame([row_dict])
        df.to_excel(db_file, index=False)
    except Exception:
        pass

def sync_to_google_sheet(payload):
    try:
        headers = {"Content-Type": "application/json"}
        response = requests.post(GOOGLE_SCRIPT_URL, json=payload, headers=headers, timeout=20)
        if response.status_code == 200:
            res_data = response.json()
            if res_data.get("status") == "success":
                return True, res_data.get("pdf_url"), "Success"
            else:
                return False, None, res_data.get("message", "Script Error")
        return False, None, f"HTTP Error: {response.status_code}"
    except Exception as e:
        return False, None, f"Network Error: {str(e)}"

def generate_a5_pdf(doc_type, meta_info, items, pdf_path):
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A5,
        leftMargin=6 * mm,
        rightMargin=6 * mm,
        topMargin=1.2 * inch,
        bottomMargin=4 * mm
    )
    story = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'DocTitle', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=10, leading=12, alignment=1, textColor=colors.HexColor('#222222')
    )
    cell_style = ParagraphStyle('CellText', parent=styles['Normal'], fontName='Helvetica', fontSize=6.2, leading=7.5, textColor=colors.HexColor('#222222'))
    cell_bold = ParagraphStyle('CellBold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=6.2, leading=7.5, textColor=colors.HexColor('#111111'))
    cell_header = ParagraphStyle('CellHeader', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=6.2, leading=7.5, textColor=colors.HexColor('#111111'), alignment=1)
    cell_right = ParagraphStyle('CellRight', parent=styles['Normal'], fontName='Helvetica', fontSize=6.2, leading=7.5, textColor=colors.HexColor('#222222'), alignment=2)
    cell_right_bold = ParagraphStyle('CellRightBold', parent=styles['Normal'], fontName
