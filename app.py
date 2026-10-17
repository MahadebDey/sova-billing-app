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
    cell_right_bold = ParagraphStyle('CellRightBold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=6.2, leading=7.5, textColor=colors.HexColor('#111111'), alignment=2)
    terms_style = ParagraphStyle('TermsText', parent=styles['Normal'], fontName='Helvetica', fontSize=5.0, leading=6.2, textColor=colors.HexColor('#333333'))
    sign_style = ParagraphStyle('SignText', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=5.5, leading=7.0, textColor=colors.HexColor('#111111'))

    is_gst = (doc_type == "Tax Invoice (GST)")
    title_text = "<u>TAX INVOICE</u>" if is_gst else "<u>ESTIMATE</u>"
    story.append(Paragraph(title_text, title_style))
    story.append(Spacer(1, 1.2 * mm))

    cust_mob_display = meta_info['customer_mob'] if meta_info['customer_mob'] else "NA"
    cust_pan_display = meta_info.get('customer_pan') if meta_info.get('customer_pan') else "NA"
    cust_id_display = meta_info.get('customer_id') if meta_info.get('customer_id') else "NA"

    if is_gst:
        meta_data = [
            [Paragraph(f"<b>Date:</b> {meta_info['date']}", cell_style), Paragraph(f"<b>GSTIN:</b> {FIRM_GSTIN}", cell_style)],
            [Paragraph(f"<b>Invoice No:</b> {meta_info['doc_no']}", cell_style), Paragraph(f"<b>Hallmark Reg:</b> {HALLMARK_REG_NO}", cell_style)],
            [Paragraph(f"<b>Customer Name:</b> {meta_info['customer_name']}", cell_style), Paragraph(f"<b>Payment Mode:</b> {meta_info.get('payment_mode', 'UPI / Cash')}", cell_style)],
            [Paragraph(f"<b>Customer Mob:</b> {cust_mob_display}", cell_style), Paragraph(f"<b>Govt ID / Aadhaar:</b> {cust_id_display}", cell_style)],
            [Paragraph(f"<b>Customer PAN:</b> {cust_pan_display}", cell_style), Paragraph("<b>State & Code:</b> 20 - JHARKHAND", cell_style)],
            [Paragraph(f"<b>Address:</b> {meta_info['customer_address']}", cell_style), Paragraph(f"<b>Party GSTIN:</b> {meta_info['customer_gstin'] or 'NA'}", cell_style)]
        ]
    else:
        meta_data = [
            [Paragraph(f"<b>Date:</b> {meta_info['date']}", cell_style), Paragraph(f"<b>Hallmark Reg:</b> {HALLMARK_REG_NO}", cell_style)],
            [Paragraph(f"<b>Estimate No:</b> {meta_info['doc_no']}", cell_style), Paragraph(f"<b>Payment Mode:</b> {meta_info.get('payment_mode', 'UPI / Cash')}", cell_style)],
            [Paragraph(f"<b>Customer Name:</b> {meta_info['customer_name']}", cell_style), Paragraph(f"<b>Customer Mob:</b> {cust_mob_display}", cell_style)],
            [Paragraph(f"<b>Govt ID / Aadhaar:</b> {cust_id_display}", cell_style), Paragraph(f"<b>Customer PAN:</b> {cust_pan_display}", cell_style)],
            [Paragraph(f"<b>Address:</b> {meta_info['customer_address']}", cell_style), Paragraph("<b>State:</b> JHARKHAND", cell_style)]
        ]

    meta_table = Table(meta_data, colWidths=[68 * mm, 68 * mm])
    meta_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#444444')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 0.8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0.8),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 1.2 * mm))

    if is_gst:
        item_rows = [[
            Paragraph("Sr", cell_header), Paragraph("Description & HUID", cell_header),
            Paragraph("Gross Wt", cell_header), Paragraph("Net Wt", cell_header),
            Paragraph("HSN", cell_header), Paragraph("Purity", cell_header),
            Paragraph("Rate/10g", cell_header), Paragraph("Making", cell_header),
            Paragraph("Total", cell_header)
        ]]
        col_widths = [6 * mm, 32 * mm, 13 * mm, 13 * mm, 10 * mm, 14 * mm, 18 * mm, 13 * mm, 17 * mm]
    else:
        item_rows = [[
            Paragraph("Sr", cell_header), Paragraph("Description & HUID", cell_header),
            Paragraph("Gross Wt", cell_header), Paragraph("Net Wt", cell_header),
            Paragraph("Purity", cell_header), Paragraph("Rate/10g", cell_header),
            Paragraph("Making", cell_header), Paragraph("Total Amount", cell_header)
        ]]
        col_widths = [6 * mm, 36 * mm, 15 * mm, 15 * mm, 15 * mm, 19 * mm, 13 * mm, 17 * mm]

    for idx, itm in enumerate(items, 1):
        desc_huid_text = f"<b>{itm['desc']}</b>"
        if itm.get('huid') and itm['huid'] != "NA":
            desc_huid_text += f"<br/><font size='5.2' color='#444444'>HUID: {itm['huid']}</font>"

        if is_gst:
            item_rows.append([
                Paragraph(str(idx), cell_style), Paragraph(desc_huid_text, cell_style),
                Paragraph(f"{itm['gross_wt']:.2f} g", cell_style), Paragraph(f"{itm['net_wt']:.2f} g", cell_style),
                Paragraph(str(itm['hsn']), cell_style), Paragraph(str(itm['purity']), cell_style),
                Paragraph(f"Rs. {itm['rate']:,.2f}", cell_style), Paragraph(str(itm['making_display']), cell_style),
                Paragraph(f"Rs. {itm['total']:,.2f}", cell_style)
            ])
        else:
            item_rows.append([
                Paragraph(str(idx), cell_style), Paragraph(desc_huid_text, cell_style),
                Paragraph(f"{itm['gross_wt']:.2f} g", cell_style), Paragraph(f"{itm['net_wt']:.2f} g", cell_style),
                Paragraph(str(itm['purity']), cell_style), Paragraph(f"Rs. {itm['rate']:,.2f}", cell_style),
                Paragraph(str(itm['making_display']), cell_style), Paragraph(f"Rs. {itm['total']:,.2f}", cell_style)
            ])

    item_table = Table(item_rows, colWidths=col_widths)
    item_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#eeeeee')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#555555')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 1.2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.2),
    ]))
    story.append(item_table)

    dynamic_blank_space = max(4 * mm, (3 - len(items)) * 4.5 * mm)
    story.append(Spacer(1, dynamic_blank_space))

    # Terms Block with Customer Legal Binding Declaration
    left_block = Paragraph(
        "<b>Terms & Conditions:</b><br/>"
        "1. Gold purity certified as per BIS Hallmarking guidelines.<br/>"
        "2. Breakage/stone responsibility ceases once item delivered.<br/>"
        "3. Making charges & GST are strictly non-refundable.<br/>"
        "4. Old metal exchange confirms lawful self-acquired property.<br/>"
        "5. Disputes subject to Jamshedpur jurisdiction only.",
        terms_style
    )
    
    old_exchange = meta_info.get('old_exchange', 0.0)
    exchange_label = meta_info.get('exchange_label', 'Less Old Exchange')
    final_balance_label = meta_info.get('final_balance_label', 'Net Payable')

    if is_gst:
        summary_rows = [
            [left_block, Paragraph("<b>Subtotal:</b>", cell_style), Paragraph(f"Rs. {meta_info['subtotal']:,.2f}", cell_right)],
            ["", Paragraph("<b>CGST (1.5%):</b>", cell_style), Paragraph(f"Rs. {meta_info['cgst']:,.2f}", cell_right)],
            ["", Paragraph("<b>SGST (1.5%):</b>", cell_style), Paragraph(f"Rs. {meta_info['sgst']:,.2f}", cell_right)],
            ["", Paragraph("<b>Gross Total:</b>", cell_style), Paragraph(f"Rs. {meta_info['gross']:,.2f}", cell_right)],
        ]
        if old_exchange > 0:
            summary_rows.append(["", Paragraph(f"<b>{exchange_label}:</b>", cell_style), Paragraph(f"- Rs. {old_exchange:,.2f}", cell_right)])
        
        summary_rows.append(["", Paragraph("<b>Round Off:</b>", cell_style), Paragraph(f"Rs. {meta_info['round_off']:+.2f}", cell_right)])
        summary_rows.append(["", Paragraph(f"<b>{final_balance_label}:</b>", cell_bold), Paragraph(f"<b>Rs. {meta_info['net_payable']:,.2f}</b>", cell_right_bold)])
    else:
        summary_rows = [
            [left_block, Paragraph("<b>Total Value:</b>", cell_style), Paragraph(f"Rs. {meta_info['subtotal']:,.2f}", cell_right)],
        ]
        if old_exchange > 0:
            summary_rows.append(["", Paragraph(f"<b>{exchange_label}:</b>", cell_style), Paragraph(f"- Rs. {old_exchange:,.2f}", cell_right)])
        
        summary_rows.append(["", Paragraph("<b>Round Off:</b>", cell_style), Paragraph(f"Rs. {meta_info['round_off']:+.2f}", cell_right)])
        summary_rows.append(["", Paragraph(f"<b>{final_balance_label}:</b>", cell_bold), Paragraph(f"<b>Rs. {meta_info['net_payable']:,.2f}</b>", cell_right_bold)])

    summary_table = Table(summary_rows, colWidths=[68 * mm, 36 * mm, 32 * mm])
    summary_table.setStyle(TableStyle([
        ('SPAN', (0,0), (0, len(summary_rows)-1)),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('BOX', (1,0), (-1,-1), 0.5, colors.HexColor('#444444')),
        ('INNERGRID', (1,0), (-1,-1), 0.4, colors.HexColor('#777777')),
        ('TOPPADDING', (0,0), (-1,-1), 0.8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0.8),
    ]))
    story.append(summary_table)

    story.append(Spacer(1, 3.5 * mm))

    # Dual Signature Layout for Strict Legal Binding
    sig_data = [
        [
            Paragraph("Customer Signature / अंगूठा<br/><br/><br/>______________________<br/><font size='4.5' color='#555555'>(Declared lawful owner of exchange metal)</font>", sign_style),
            Paragraph("For SOVAA JEWELLERS<br/><br/><br/>______________________<br/><font size='4.5' color='#555555'>(Authorised Signatory)</font>", sign_style)
        ]
    ]
    sig_table = Table(sig_data, colWidths=[68 * mm, 68 * mm])
    sig_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (0,0), 'LEFT'),
        ('ALIGN', (1,0), (1,0), 'RIGHT'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 0.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0.5),
    ]))
    story.append(sig_table)

    doc.build(story)

def render_mobile_print_button(pdf_base64, doc_no):
    print_code = f"""
    <html>
    <head>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
      @media print {{
        html, body {{
          filter: grayscale(100%) !important;
          -webkit-filter: grayscale(100%) !important;
          color: #000000 !important;
        }}
      }}
      .print-btn {{
        background-color: #222222;
        color: white;
        padding: 12px 20px;
        font-size: 16px;
        font-weight: bold;
        border: none;
        border-radius: 8px;
        cursor: pointer;
        width: 100%;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
      }}
      .print-btn:active {{
        background-color: #000000;
      }}
    </style>
    </head>
    <body style="margin:0; padding:4px 0;">
      <button class="print-btn" onclick="printInvoice()">🖨️ Quick Print ({doc_no}) to Epson L3350</button>
      <script>
        function printInvoice() {{
