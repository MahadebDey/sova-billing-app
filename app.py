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

st.markdown("""
<style>
input[type="text"] {
    text-transform: uppercase !important;
}
</style>
""", unsafe_allow_html=True)

# Business Details
HALLMARK_REG_NO = "HM/C-5800000504"
FIRM_GSTIN = "20AUJPD1127G1ZE"
FIRM_PHONE = "7717740697"
GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwyLXOi4yCglsCueHm_nVQ9wYjPUI0eKsPlxYxYGqFnwIvcML66W5JWfRCwtqj8WuhU/exec"

# Setup Directory & Rates Storage
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
                return {
                    "gold_22k": float(d.get("gold_22k", 72000.0)),
                    "gold_18k": float(d.get("gold_18k", 59000.0)),
                    "silver": float(d.get("silver", 880.0))
                }
        except Exception:
            return defs
    return defs

def save_rates_to_file(g22, g18, sil):
    try:
        with open(RATES_FILE, "w") as f:
            json.dump({"gold_22k": float(g22), "gold_18k": float(g18), "silver": float(sil)}, f)
    except Exception:
        pass

# Initialize Session States
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
if "c_aadhaar" not in st.session_state:
    st.session_state.c_aadhaar = ""
if "c_pan" not in st.session_state:
    st.session_state.c_pan = ""
if "c_gstin" not in st.session_state:
    st.session_state.c_gstin = ""
if "items_list" not in st.session_state:
    st.session_state.items_list = []
if "form_reset_count" not in st.session_state:
    st.session_state.form_reset_count = 0
if "last_looked_up_mob" not in st.session_state:
    st.session_state.last_looked_up_mob = ""
if "editing_item_idx" not in st.session_state:
    st.session_state.editing_item_idx = None
if "lock_bill_date" not in st.session_state:
    st.session_state.lock_bill_date = False
if "locked_date_val" not in st.session_state:
    st.session_state.locked_date_val = datetime.today().date()
if "purity_selector" not in st.session_state:
    st.session_state.purity_selector = "22K (916)"
if "last_selected_preset" not in st.session_state:
    st.session_state.last_selected_preset = "-- Select Common Item --"

COMMON_ITEMS = [
    "-- Select Common Item --", "GOLD LOCKET", "SILVER LOCKET", "GOLD NATHIYA",
    "SILVER BABY BALA", "SILVER BABY PAYAL", "SILVER RING", "SILVER BICHHIYA",
    "GOLD CHAIN", "GOLD RING", "GOLD TOPS / EARING", "GOLD NOSE PIN", "GOLD PENDANT",
    "SILVER PAYAL", "SILVER CHAIN", "SILVER BRACELET",
    "SILVER COIN", "SILVER MURTI / IDOL", "➕ Custom / Other Item"
]

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

                        return {
                            "name": raw_name.upper(),
                            "address": raw_addr.upper(),
                            "pan": pan_val.upper(),
                            "aadhaar": aadhaar_val,
                            "gstin": gst_val.upper()
                        }
            except Exception:
                pass
    return None

def get_current_fy_string():
    now = datetime.now()
    year = now.year
    if now.month < 4:
        s_yr, e_yr = year - 1, year
    else:
        s_yr, e_yr = year, year + 1
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

    safe_p_name = doc_clean.replace('/', '_')
    p_file = os.path.join(PDF_DIR, f"{safe_p_name}.pdf")
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
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A5,
        leftMargin=6 * mm,
        rightMargin=6 * mm,
        topMargin=1.25 * inch,
        bottomMargin=4 * mm
    )
    story = []
    styles = getSampleStyleSheet()

    num_items = len(items)
    if num_items <= 2:
        item_font = 6.6
        item_leading = 8.0
        item_pad = 2.2
        dynamic_bottom_gap = max(10 * mm, (4 - num_items) * 11.0 * mm)
        terms_font = 5.2
        terms_leading = 6.4
    elif num_items <= 4:
        item_font = 5.8
        item_leading = 7.0
        item_pad = 1.4
        dynamic_bottom_gap = max(4 * mm, (5 - num_items) * 6.0 * mm)
        terms_font = 4.8
        terms_leading = 6.0
    else:
        item_font = 5.0
        item_leading = 6.2
        item_pad = 0.8
        dynamic_bottom_gap = 2 * mm
        terms_font = 4.4
        terms_leading = 5.5

    title_style = ParagraphStyle('DocTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=12, alignment=1)
    cell_meta = ParagraphStyle('CellMeta', parent=styles['Normal'], fontName='Helvetica', fontSize=6.2, leading=7.5)
    
    cell_style = ParagraphStyle('CellText', parent=styles['Normal'], fontName='Helvetica', fontSize=item_font, leading=item_leading)
    cell_header = ParagraphStyle('CellHeader', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=item_font, leading=item_leading, alignment=1, textColor=colors.HexColor('#000000'))
    
    cell_right = ParagraphStyle('CellRight', parent=styles['Normal'], fontName='Helvetica', fontSize=5.8, leading=7.0, alignment=2)
    cell_right_bold = ParagraphStyle('CellRightBold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=6.2, leading=7.4, alignment=2)
    terms_style = ParagraphStyle('TermsText', parent=styles['Normal'], fontName='Helvetica', fontSize=terms_font, leading=terms_leading)
    sign_style = ParagraphStyle('SignText', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=5.8, leading=7.2)

    is_gst = (doc_type == "Tax Invoice (GST)")
    story.append(Paragraph("<u>TAX INVOICE</u>" if is_gst else "<u>ESTIMATE</u>", title_style))
    story.append(Spacer(1, 1.2 * mm))

    c_mob = meta_info['customer_mob'] or "NA"
    c_pan = meta_info.get('customer_pan')
    c_aadhaar = meta_info.get('customer_aadhaar')

    cust_name_pan = f"<b>Customer Name:</b> {meta_info['customer_name']}"
    if c_pan and c_pan != "NA":
        cust_name_pan += f" (PAN: {c_pan})"

    cust_addr_aadhaar = f"<b>Address:</b> {meta_info['customer_address']}"
    if c_aadhaar and c_aadhaar != "NA":
        cust_addr_aadhaar += f" (Aadhaar: {c_aadhaar})"

    if is_gst:
        meta_data = [
            [Paragraph(f"<b>Date:</b> {meta_info['date']}", cell_meta), Paragraph(f"<b>GSTIN:</b> {FIRM_GSTIN}", cell_meta)],
            [Paragraph(f"<b>Invoice No:</b> {meta_info['doc_no']}", cell_meta), Paragraph(f"<b>Hallmark Reg:</b> {HALLMARK_REG_NO}", cell_meta)],
            [Paragraph(cust_name_pan, cell_meta), Paragraph(f"<b>Payment Mode:</b> {meta_info.get('payment_mode', 'UPI/Cash')}", cell_meta)],
            [Paragraph(f"<b>Customer Mob:</b> {c_mob}", cell_meta), Paragraph("<b>State:</b> 20 - JHARKHAND", cell_meta)],
            [Paragraph(cust_addr_aadhaar, cell_meta), Paragraph(f"<b>Party GSTIN:</b> {meta_info['customer_gstin'] or 'NA'}", cell_meta)]
        ]
    else:
        meta_data = [
            [Paragraph(f"<b>Date:</b> {meta_info['date']}", cell_meta), Paragraph(f"<b>Hallmark Reg:</b> {HALLMARK_REG_NO}", cell_meta)],
            [Paragraph(f"<b>Estimate No:</b> {meta_info['doc_no']}", cell_meta), Paragraph(f"<b>Payment Mode:</b> {meta_info.get('payment_mode', 'UPI/Cash')}", cell_meta)],
            [Paragraph(cust_name_pan, cell_meta), Paragraph(f"<b>Customer Mob:</b> {c_mob}", cell_meta)],
            [Paragraph(cust_addr_aadhaar, cell_meta), Paragraph("<b>State:</b> JHARKHAND", cell_meta)]
        ]

    meta_table = Table(meta_data, colWidths=[70 * mm, 66 * mm])
    meta_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#444')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 0.7),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0.7),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 1.2 * mm))

    # Items Table
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
            desc_huid_text += f"<br/><font size='{item_font - 0.8}' color='#333'>HUID: {itm['huid']}</font>"

        if is_gst:
            item_rows.append([
                Paragraph(str(idx), cell_style), Paragraph(desc_huid_text, cell_style),
                Paragraph(f"{itm['gross_wt']:.2f}g", cell_style), Paragraph(f"{itm['net_wt']:.2f}g", cell_style),
                Paragraph(str(itm['hsn']), cell_style), Paragraph(str(itm['purity']), cell_style),
                Paragraph(f"Rs.{itm['rate']:,.2f}", cell_style), Paragraph(str(itm['making_display']), cell_style),
                Paragraph(f"Rs.{itm['total']:,.2f}", cell_style)
            ])
        else:
            item_rows.append([
                Paragraph(str(idx), cell_style), Paragraph(desc_huid_text, cell_style),
                Paragraph(f"{itm['gross_wt']:.2f}g", cell_style), Paragraph(f"{itm['net_wt']:.2f}g", cell_style),
                Paragraph(str(itm['purity']), cell_style), Paragraph(f"Rs.{itm['rate']:,.2f}", cell_style),
                Paragraph(str(itm['making_display']), cell_style), Paragraph(f"Rs.{itm['total']:,.2f}", cell_style)
            ])

    item_table = Table(item_rows, colWidths=col_widths)
    item_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#EAEAEA')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#555')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), item_pad),
        ('BOTTOMPADDING', (0,0), (-1,-1), item_pad),
    ]))
    story.append(item_table)

    story.append(Spacer(1, dynamic_bottom_gap))

    # Real Sub-Table for Signatures (Left: Customer, Middle: 12.7mm Half-inch gap, Right: SOVAA Jewellers)
    sig_subtable = Table([
        [
            Paragraph("Customer Signature<br/><br/><br/>", sign_style),
            "",
            Paragraph("For SOVAA JEWELLERS<br/><br/><br/>(Authorised Signatory)", sign_style)
        ]
    ], colWidths=[27 * mm, 12.7 * mm, 29 * mm])
    sig_subtable.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
    ]))

    terms_paragraph = Paragraph(
        "<b>Terms & Conditions:</b><br/>"
        "1. Breakage/stone responsibility ceases once item delivered.<br/>"
        "2. Making charges & GST are strictly non-refundable.<br/>"
        "3. Old metal exchange confirms lawful self-acquired property.<br/>"
        "4. Disputes subject to Jamshedpur jurisdiction.",
        terms_style
    )

    left_block = [
        terms_paragraph,
        Spacer(1, 2.5 * mm),
        sig_subtable
    ]

    old_exchange = meta_info.get('old_exchange', 0.0)
    exchange_label = meta_info.get('exchange_label', 'Less Old Exchange')
    final_balance_label = meta_info.get('final_balance_label', 'Net Payable')

    if is_gst:
        summary_rows = [
            [left_block, Paragraph("<b>Subtotal:</b>", cell_right), Paragraph(f"Rs.{meta_info['subtotal']:,.2f}", cell_right)],
            ["", Paragraph("<b>CGST (1.5%):</b>", cell_right), Paragraph(f"Rs.{meta_info['cgst']:,.2f}", cell_right)],
            ["", Paragraph("<b>SGST (1.5%):</b>", cell_right), Paragraph(f"Rs.{meta_info['sgst']:,.2f}", cell_right)],
            ["", Paragraph("<b>Gross Total:</b>", cell_right), Paragraph(f"Rs.{meta_info['gross']:,.2f}", cell_right)],
        ]
        if old_exchange > 0:
            summary_rows.append(["", Paragraph(f"<b>{exchange_label}:</b>", cell_right), Paragraph(f"- Rs.{old_exchange:,.2f}", cell_right)])
        summary_rows.append(["", Paragraph("<b>Round Off:</b>", cell_right), Paragraph(f"Rs.{meta_info['round_off']:+.2f}", cell_right)])
        summary_rows.append(["", Paragraph(f"<b>{final_balance_label}:</b>", cell_right_bold), Paragraph(f"<b>Rs.{meta_info['net_payable']:,.2f}</b>", cell_right_bold)])
    else:
        summary_rows = [
            [left_block, Paragraph("<b>Total Value:</b>", cell_right), Paragraph(f"Rs.{meta_info['subtotal']:,.2f}", cell_right)],
        ]
        if old_exchange > 0:
            summary_rows.append(["", Paragraph(f"<b>{exchange_label}:</b>", cell_right), Paragraph(f"- Rs.{old_exchange:,.2f}", cell_right)])
        summary_rows.append(["", Paragraph("<b>Round Off:</b>", cell_right), Paragraph(f"Rs.{meta_info['round_off']:+.2f}", cell_right)])
        summary_rows.append(["", Paragraph(f"<b>{final_balance_label}:</b>", cell_right_bold), Paragraph(f"<b>Rs.{meta_info['net_payable']:,.2f}</b>", cell_right_bold)])

    summary_table = Table(summary_rows, colWidths=[70 * mm, 34 * mm, 32 * mm])
    summary_table.setStyle(TableStyle([
        ('SPAN', (0,0), (0, len(summary_rows)-1)),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('BOX', (1,0), (-1,-1), 0.5, colors.HexColor('#444')),
        ('INNERGRID', (1,0), (-1,-1), 0.4, colors.HexColor('#777')),
        ('TOPPADDING', (0,0), (-1,-1), 0.8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0.8),
    ]))
    story.append(summary_table)

    doc.build(story)

def render_mobile_print_button(pdf_base64, doc_no):
    html = f"""<div style="margin:4px 0;"><button style="background:#222;color:#fff;padding:12px;font-size:16px;font-weight:bold;border:none;border-radius:8px;width:100%;cursor:pointer;" onclick="pDoc()">🖨️ Quick Print ({doc_no}) to Epson L3350</button><script>function pDoc(){{var b='{pdf_base64}';var c=atob(b);var n=new Array(c.length);for(var i=0;i<c.length;i++){{n[i]=c.charCodeAt(i);}}var blob=new Blob([new Uint8Array(n)],{{type:'application/pdf'}});var u=URL.createObjectURL(blob);var w=window.open(u);if(w){{w.focus();w.print();}}else{{window.location.href=u;}}}}</script></div>"""
    components.html(html, height=60)

# --- HEADER SECTION ---
st.title("💎 SOVAA JEWELLERS")
st.write(f"**Hallmark Reg No:** `{HALLMARK_REG_NO}` | **GSTIN:** `{FIRM_GSTIN}`")

# --- DAILY MARKET RATES PANEL ---
with st.expander("📈 **Today's Market Rates (Aaj ka Gold & Silver Bhav)**", expanded=False):
    r_col1, r_col2, r_col3 = st.columns(3)
    with r_col1:
        new_g22 = st.number_input(
            "🪙 22K Gold Rate (Per 10g) ₹", 
            min_value=0.0, step=100.0, format="%.2f", 
            value=float(st.session_state.daily_gold_rate)
        )
    with r_col2:
        new_g18 = st.number_input(
            "🪙 18K Gold Rate (Per 10g) ₹", 
            min_value=0.0, step=100.0, format="%.2f", 
            value=float(st.session_state.daily_gold_18k_rate)
        )
    with r_col3:
        new_sil = st.number_input(
            "🥈 Silver Rate (Per 10g) ₹", 
            min_value=0.0, step=10.0, format="%.2f", 
            value=float(st.session_state.daily_silver_rate)
        )
    
    if st.button("💾 Save Rates (Bhav Save Karein)"):
        st.session_state.daily_gold_rate = new_g22
        st.session_state.daily_gold_18k_rate = new_g18
        st.session_state.daily_silver_rate = new_sil
        save_rates_to_file(new_g22, new_g18, new_sil)
        st.success("✅ Rates save ho gaye!")
        st.rerun()

tab_billing, tab_gst_report, tab_search = st.tabs([
    "📝 New Bill / Estimate",
    "📊 GST Filing (GSTR-1 Ready)",
    "🔍 Search Customer / History"
])

with tab_billing:
    mode = st.radio(
        "Select Document Type",
        ["Tax Invoice (GST)", "Estimate (Without GST)"],
        index=1,
        horizontal=True
    )

    if "current_mode" not in st.session_state or st.session_state.current_mode != mode:
        st.session_state.current_mode = mode
        st.session_state.active_doc_number = get_next_number(mode)

    if "active_doc_number" not in st.session_state:
        st.session_state.active_doc_number = get_next_number(mode)

    c_hdr1, c_hdr2 = st.columns([4, 1])
    with c_hdr1:
        st.markdown("##### 👤 Customer Information")
    with c_hdr2:
        if st.button("🧹 Clear Form / New Customer"):
            st.session_state.c_name = ""
            st.session_state.c_mob = ""
            st.session_state.c_addr = "MANGO, JAMSHEDPUR"
            st.session_state.c_aadhaar = ""
            st.session_state.c_pan = ""
            st.session_state.c_gstin = ""
            st.session_state.last_looked_up_mob = ""
            st.session_state.items_list = []
            st.session_state.editing_item_idx = None
            st.session_state.form_reset_count += 1
            st.session_state.active_doc_number = get_next_number(mode)
            st.rerun()

    r_key = st.session_state.form_reset_count
    mode_tag = "gst" if mode == "Tax Invoice (GST)" else "est"

    col1, col2 = st.columns([1.5, 1.5])
    with col2:
        # Date Input with Lock Option for Back-Date Entries
        d_col1, d_col2 = st.columns([2, 1.3])
        with d_col1:
            default_dt = st.session_state.locked_date_val if st.session_state.lock_bill_date else datetime.today().date()
            doc_date = st.date_input("Date", value=default_dt, key=f"cdate_{r_key}")
            if st.session_state.lock_bill_date:
                st.session_state.locked_date_val = doc_date
        with d_col2:
            st.write("")
            st.write("")
            is_locked = st.checkbox("🔒 Lock Date", value=st.session_state.lock_bill_date, help="Back-date entries karte samay date ko lock rakhein")
            if is_locked != st.session_state.lock_bill_date:
                st.session_state.lock_bill_date = is_locked
                st.session_state.locked_date_val = doc_date
                st.rerun()

        cust_mob = st.text_input("Mobile No", value=st.session_state.c_mob, key=f"cmob_{r_key}", placeholder="10 Digit Number")

        clean_entered_mob = re.sub(r'\D', '', str(cust_mob))
        if len(clean_entered_mob) >= 10 and clean_entered_mob != st.session_state.last_looked_up_mob:
            st.session_state.last_looked_up_mob = clean_entered_mob
            st.session_state.c_mob = clean_entered_mob
            found_cust = lookup_customer_by_mobile(clean_entered_mob)
            if found_cust:
                st.session_state.c_name = found_cust["name"]
                st.session_state.c_addr = found_cust["address"]
                st.session_state.c_pan = found_cust["pan"]
                st.session_state.c_aadhaar = found_cust["aadhaar"]
                st.session_state.c_gstin = found_cust["gstin"]
                st.session_state.form_reset_count += 1
                st.toast(f"⚡ Purana Record Mil Gaya: {found_cust['name']}")
                st.rerun()

        payment_mode_selected = st.selectbox("Payment Mode", ["UPI / QR (PhonePe/GPay)", "Cash", "Bank Transfer / Card", "Multiple Modes"], key=f"pmode_{r_key}")
        if mode == "Tax Invoice (GST)":
            cust_gstin_in = st.text_input("Party GSTIN (Optional)", value=st.session_state.c_gstin, key=f"cgst_{r_key}", placeholder="OPTIONAL")
        else:
            cust_gstin_in = ""

    with col1:
        nc1, nc2 = st.columns([3, 1])
        with nc1:
            doc_number = st.text_input(
                "Invoice / Estimate No", 
                value=st.session_state.active_doc_number, 
                key=f"doc_box_{mode_tag}_{r_key}"
            )
        with nc2:
            st.write("")
            st.write("")
            if st.button("🔄 Auto", help="Latest serial number fetch karein"):
                st.session_state.active_doc_number = get_next_number(mode)
                st.rerun()
                
        cust_name_in = st.text_input("Customer Name", value=st.session_state.c_name, key=f"cname_{r_key}", placeholder="e.g. SOMNATH DAS")
        cust_pan_in = st.text_input("Customer PAN (Name ke sath print hoga)", value=st.session_state.c_pan, key=f"cpan_{r_key}", placeholder="e.g. ABCDE1234F")
        cust_address_in = st.text_input("Address", value=st.session_state.c_addr, key=f"caddr_{r_key}")
        cust_aadhaar_in = st.text_input("Aadhaar No (Address ke sath print hoga)", value=st.session_state.c_aadhaar, key=f"caadhaar_{r_key}", placeholder="e.g. 12 Digit Aadhaar")

    cust_name = cust_name_in.strip().upper()
    cust_pan = cust_pan_in.strip().upper()
    cust_address = cust_address_in.strip().upper()
    cust_aadhaar = cust_aadhaar_in.strip().upper()
    cust_gstin = cust_gstin_in.strip().upper()

    st.markdown("---")
    st.subheader("🛒 Item Details")

    selected_preset = st.selectbox("⚡ Quick Select Item (Optional)", COMMON_ITEMS)

    purity_options = ["22K (916)", "18K (750)", "SILVER", "24K (999)"]

    # Automatically switch purity to SILVER or 22K when Common Item changes
    if selected_preset != st.session_state.last_selected_preset:
        st.session_state.last_selected_preset = selected_preset
        if "SILVER" in selected_preset:
            st.session_state.purity_selector = "SILVER"
        elif "GOLD" in selected_preset:
            st.session_state.purity_selector = "22K (916)"

    p_col1, p_col2 = st.columns([1.5, 2.5])
    with p_col1:
        purity_choice = st.selectbox("Purity", purity_options, key="purity_selector")
    with p_col2:
        st.write("")
        st.caption(f"🛡️ Hallmarking: {purity_choice} under {HALLMARK_REG_NO}")

    if purity_choice == "18K (750)":
        auto_rate_val = st.session_state.daily_gold_18k_rate
    elif purity_choice == "SILVER":
        auto_rate_val = st.session_state.daily_silver_rate
    else:
        auto_rate_val = st.session_state.daily_gold_rate

    with st.form("item_entry_form", clear_on_submit=True):
        ic1, ic2, ic3, ic4, ic5, ic6, ic7 = st.columns([2.2, 1.2, 1.1, 1.1, 1.3, 1.1, 0.8])
        
        default_text = ""
        if selected_preset not in ["-- Select Common Item --", "➕ Custom / Other Item"]:
            default_text = selected_preset

        with ic1:
            item_desc = st.text_input("Item Description", value=default_text, placeholder="e.g. GOLD LOCKET")
        with ic2:
            item_huid = st.text_input("HUID Code (Optional)", placeholder="e.g. AB1234")
        with ic3:
            gross_wt = st.number_input("Gross Wt (g)", min_value=0.0, step=0.01, format="%.2f", value=None, placeholder="0.00")
        with ic4:
            net_wt = st.number_input("Net Wt (g)", min_value=0.0, step=0.01, format="%.2f", value=None, placeholder="0.00")
        with ic5:
            rate_10g = st.number_input("Rate (Per 10g)", min_value=0.0, step=10.0, format="%.2f", value=auto_rate_val)
        with ic6:
            making_val = st.number_input("Making Charge", min_value=0.0, step=1.0, value=None, placeholder="Making")
        with ic7:
            making_type = st.selectbox("Unit", ["%", "₹"])

        next_item_no = len(st.session_state.items_list) + 1
        submitted = st.form_submit_button(f"➕ Add Item #{next_item_no} ({purity_choice})", use_container_width=True)
        if submitted:
            if item_desc and net_wt is not None and net_wt > 0 and rate_10g is not None and rate_10g > 0:
                val = making_val if making_val is not None else 0.0
                g_wt = gross_wt if (gross_wt is not None and gross_wt > 0) else net_wt
                metal_cost = (net_wt * rate_10g) / 10.0
                
                if making_type == "%":
                    making_amt = metal_cost * (val / 100.0)
                    display_str = f"{val}%"
                else:
                    making_amt = val
                    display_str = f"Rs. {val:,.0f}"

                total_item_amt = metal_cost + making_amt
                st.session_state.items_list.append({
                    "desc": item_desc.strip().upper(),
                    "huid": item_huid.strip().upper() if item_huid else "NA",
                    "gross_wt": g_wt,
                    "net_wt": net_wt,
                    "hsn": "7113",
                    "purity": purity_choice,
                    "rate": rate_10g,
                    "making_val": val,
                    "making_unit": making_type,
                    "making_display": display_str,
                    "total": round(total_item_amt, 2)
                })
                st.rerun()
            else:
                st.warning("Kripya Item Description, Net Weight aur Rate sahi se bharein.")

    if not st.session_state.items_list:
        st.info("💡 Bill banane ke liye pehle Item Description aur Wazan daal kar **'➕ Add Item'** par click karein.")
    else:
        df_items = pd.DataFrame(st.session_state.items_list)
        display_cols = ["desc", "huid", "gross_wt", "net_wt", "purity", "rate", "making_display", "total"]
        if mode == "Tax Invoice (GST)":
            display_cols.insert(4, "hsn")
        st.table(df_items[display_cols])

        # --- INDIVIDUAL ITEM EDIT & DELETE SECTION ---
        with st.expander("✏️ Edit / Delete Any Added Item (किसी आइटम को सुधारें या हटाएँ)", expanded=False):
            for idx, itm in enumerate(st.session_state.items_list):
                ec1, ec2, ec3, ec4 = st.columns([3.5, 2.0, 1.0, 1.0])
                with ec1:
                    st.markdown(f"**#{idx+1} {itm['desc']}** (HUID: {itm.get('huid','NA')})")
                with ec2:
                    st.caption(f"Net: {itm['net_wt']}g | Rate: ₹{itm['rate']:,.0f} | Making: {itm['making_display']} | Total: ₹{itm['total']:,.2f}")
                with ec3:
                    if st.button("✏️ Edit", key=f"btn_edit_itm_{idx}"):
                        st.session_state.editing_item_idx = idx
                        st.rerun()
                with ec4:
                    if st.button("❌ Del", key=f"btn_del_itm_{idx}"):
                        st.session_state.items_list.pop(idx)
                        if st.session_state.editing_item_idx == idx:
                            st.session_state.editing_item_idx = None
                        st.rerun()

            if st.session_state.editing_item_idx is not None and st.session_state.editing_item_idx < len(st.session_state.items_list):
                e_idx = st.session_state.editing_item_idx
                e_itm = st.session_state.items_list[e_idx]
                st.markdown("---")
                st.markdown(f"##### 🔧 Editing Item #{e_idx + 1}: {e_itm['desc']}")
                with st.form(f"edit_item_form_{e_idx}"):
                    f1, f2, f3, f4 = st.columns([2, 1.2, 1.2, 1.2])
                    with f1:
                        new_desc = st.text_input("Description", value=e_itm["desc"])
                    with f2:
                        new_huid = st.text_input("HUID", value="" if e_itm.get("huid") == "NA" else e_itm.get("huid", ""))
                    with f3:
                        new_gwt = st.number_input("Gross Wt (g)", min_value=0.0, step=0.01, format="%.2f", value=float(e_itm["gross_wt"]))
                    with f4:
                        new_nwt = st.number_input("Net Wt (g)", min_value=0.0, step=0.01, format="%.2f", value=float(e_itm["net_wt"]))

                    f5, f6, f7, f8 = st.columns([1.5, 1.5, 1.2, 1.0])
                    with f5:
                        p_idx = purity_options.index(e_itm["purity"]) if e_itm["purity"] in purity_options else 0
                        new_pur = st.selectbox("Purity", purity_options, index=p_idx)
                    with f6:
                        new_rate = st.number_input("Rate (Per 10g)", min_value=0.0, step=10.0, format="%.2f", value=float(e_itm["rate"]))
                    with f7:
                        cur_mval = float(e_itm.get("making_val", 0.0))
                        new_mval = st.number_input("Making Charge", min_value=0.0, step=1.0, value=cur_mval)
                    with f8:
                        cur_munit = e_itm.get("making_unit", "%" if "%" in str(e_itm["making_display"]) else "₹")
                        new_munit = st.selectbox("Unit", ["%", "₹"], index=0 if cur_munit == "%" else 1)

                    b_save, b_cancel = st.columns(2)
                    with b_save:
                        if st.form_submit_button("💾 Update Item", type="primary", use_container_width=True):
                            if new_desc and new_nwt > 0 and new_rate > 0:
                                m_cost = (new_nwt * new_rate) / 10.0
                                if new_munit == "%":
                                    m_amt = m_cost * (new_mval / 100.0)
                                    d_str = f"{new_mval}%"
                                else:
                                    m_amt = new_mval
                                    d_str = f"Rs. {new_mval:,.0f}"
                                st.session_state.items_list[e_idx] = {
                                    "desc": new_desc.strip().upper(),
                                    "huid": new_huid.strip().upper() if new_huid else "NA",
                                    "gross_wt": new_gwt if new_gwt > 0 else new_nwt,
                                    "net_wt": new_nwt,
                                    "hsn": "7113",
                                    "purity": new_pur,
                                    "rate": new_rate,
                                    "making_val": new_mval,
                                    "making_unit": new_munit,
                                    "making_display": d_str,
                                    "total": round(m_cost + m_amt, 2)
                                }
                                st.session_state.editing_item_idx = None
                                st.rerun()
                    with b_cancel:
                        if st.form_submit_button("❌ Cancel", use_container_width=True):
                            st.session_state.editing_item_idx = None
                            st.rerun()

        # --- REVERSE MAKING CHARGE CALCULATOR ---
        with st.expander("🧮 Reverse Making Charge Calculator (कस्टमर के फाइनल अमाउंट से मेकिंग निकालें)", expanded=False):
            st.caption("यदि कस्टमर कुल अमाउंट से कम रुपए दे रहा है, तो यहाँ कस्टमर द्वारा दिया जा रहा फाइनल अमाउंट डालें—मेकिंग चार्ज अपने आप एडजस्ट हो जाएगा:")
            rc_col1, rc_col2, rc_col3 = st.columns([2, 1.5, 1.5])
            with rc_col1:
                item_labels = [f"#{i+1}
