import streamlit as st
import pandas as pd
from datetime import datetime
import os
import math
import re
import base64
import requests
from reportlab.lib.pagesizes import A5
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm, inch

# Page Configuration
st.set_page_config(page_title="SOVAA JEWELLERS - Billing & Sync", layout="wide", page_icon="💎")

# Google Apps Script Web App URL
GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbx9zxPUIPhyCk46GLaYeBlBdcF--BbQStDSi-EQAGdcmqj-E6ahzfVmPH4KqEfT0WatCQ/exec"

# --- PASSWORD AUTHENTICATION ---
APP_PASSWORD = "sovaa"

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

def check_password():
    if st.session_state.get("password_input") == APP_PASSWORD:
        st.session_state.authenticated = True
        del st.session_state["password_input"]
    else:
        st.error("❌ Galat Password! Kripya sahi password dalein.")

if not st.session_state.authenticated:
    st.markdown("<h2 style='text-align: center;'>🔒 SOVAA JEWELLERS</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>App access karne ke liye password enter karein</p>", unsafe_allow_html=True)
    
    col_l, col_m, col_r = st.columns([1, 1.2, 1])
    with col_m:
        st.text_input("Password", type="password", key="password_input", on_change=check_password)
        st.button("🔓 Unlock App", on_click=check_password, use_container_width=True)
    st.stop()

# --- APP SETUP ---
PDF_DIR = "Invoices_PDF"
os.makedirs(PDF_DIR, exist_ok=True)

# Form Session States
if "cust_name_input" not in st.session_state:
    st.session_state.cust_name_input = ""
if "cust_mob_input" not in st.session_state:
    st.session_state.cust_mob_input = ""
if "cust_addr_input" not in st.session_state:
    st.session_state.cust_addr_input = "Mango, Jamshedpur"
if "cust_gstin_input" not in st.session_state:
    st.session_state.cust_gstin_input = ""
if "items_list" not in st.session_state:
    st.session_state.items_list = []

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

def get_next_number(doc_type):
    db_file = "Sales_Database.xlsx" if doc_type == "Tax Invoice (GST)" else "Estimate_Database.xlsx"
    prefix = "SV" if doc_type == "Tax Invoice (GST)" else "EST"
    num_col = "Document No"
    cur_year = datetime.now().year
    
    max_val = 0
    if os.path.exists(db_file):
        try:
            df = pd.read_excel(db_file)
            if not df.empty and num_col in df.columns:
                for val in df[num_col].dropna().astype(str):
                    matches = re.findall(r'(\d+)$', val.strip())
                    if matches:
                        num = int(matches[0])
                        if num > max_val:
                            max_val = num
        except Exception:
            pass
    
    next_seq = max_val + 1
    return f"{prefix}-{cur_year}-{str(next_seq).zfill(3)}"

def save_to_database(row_dict, doc_type):
    db_file = "Sales_Database.xlsx" if doc_type == "Tax Invoice (GST)" else "Estimate_Database.xlsx"
    if os.path.exists(db_file):
        try:
            df = pd.read_excel(db_file)
            df = pd.concat([df, pd.DataFrame([row_dict])], ignore_index=True)
        except Exception:
            df = pd.DataFrame([row_dict])
    else:
        df = pd.DataFrame([row_dict])
    df.to_excel(db_file, index=False)

def sync_to_google_sheet(payload):
    try:
        response = requests.post(GOOGLE_SCRIPT_URL, json=payload, timeout=15)
        if response.status_code == 200:
            res_data = response.json()
            if res_data.get("status") == "success":
                return True, res_data.get("pdf_url")
        return False, None
    except Exception:
        return False, None

def generate_a5_pdf(doc_type, meta_info, items, pdf_path):
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A5,
        leftMargin=7*mm,
        rightMargin=7*mm,
        topMargin=1.7*inch,
        bottomMargin=7*mm
    )
    story = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'DocTitle', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=10, leading=12, alignment=1
    )
    cell_style = ParagraphStyle('CellText', parent=styles['Normal'], fontName='Helvetica', fontSize=7, leading=8.5)
    cell_bold = ParagraphStyle('CellBold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7, leading=8.5)
    cell_header = ParagraphStyle('CellHeader', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=6.5, leading=8, textColor=colors.whitesmoke, alignment=1)

    is_gst = (doc_type == "Tax Invoice (GST)")
    title_text = "<u>TAX INVOICE</u>" if is_gst else "<u>ESTIMATE</u>"
    story.append(Paragraph(title_text, title_style))
    story.append(Spacer(1, 1.5*mm))

    cust_mob_display = meta_info['customer_mob'] if meta_info['customer_mob'] else "NA"

    if is_gst:
        meta_data = [
            [Paragraph(f"<b>Date:</b> {meta_info['date']}", cell_style), Paragraph("<b>GSTIN:</b> 20AUJPD1127G1ZE", cell_style)],
            [Paragraph(f"<b>Invoice No:</b> {meta_info['doc_no']}", cell_style), Paragraph("<b>Mobile No:</b> 7717740697", cell_style)],
            [Paragraph(f"<b>Customer Name:</b> {meta_info['customer_name']}", cell_style), Paragraph("<b>GST Type:</b> CGST + SGST (1.5% Each)", cell_style)],
            [Paragraph(f"<b>Customer Mob:</b> {cust_mob_display}", cell_style), Paragraph("<b>State & Code:</b> 20 - JHARKHAND", cell_style)],
            [Paragraph(f"<b>Address:</b> {meta_info['customer_address']}", cell_style), Paragraph(f"<b>Party GSTIN:</b> {meta_info['customer_gstin'] or 'NA'}", cell_style)]
        ]
    else:
        meta_data = [
            [Paragraph(f"<b>Date:</b> {meta_info['date']}", cell_style), Paragraph("<b>Mobile No:</b> 7717740697", cell_style)],
            [Paragraph(f"<b>Estimate No:</b> {meta_info['doc_no']}", cell_style), Paragraph("<b>State:</b> JHARKHAND", cell_style)],
            [Paragraph(f"<b>Customer Name:</b> {meta_info['customer_name']}", cell_style), Paragraph(f"<b>Customer Mob:</b> {cust_mob_display}", cell_style)],
            [Paragraph(f"<b>Address:</b> {meta_info['customer_address']}", cell_style), Paragraph("", cell_style)]
        ]

    meta_table = Table(meta_data, colWidths=[67*mm, 67*mm])
    meta_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#0b2f64')),
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f7f9fc')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 1.2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.2),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 1.5*mm))

    # Perfect balanced columns for both GST & Estimate
    if is_gst:
        item_rows = [[
            Paragraph("Sr", cell_header), Paragraph("Description", cell_header),
            Paragraph("Gross Wt", cell_header), Paragraph("Net Wt", cell_header),
            Paragraph("HSN", cell_header), Paragraph("Purity", cell_header),
            Paragraph("Rate/10g", cell_header), Paragraph("Making", cell_header),
            Paragraph("Total", cell_header)
        ]]
        col_widths = [6*mm, 29*mm, 14*mm, 14*mm, 10*mm, 14*mm, 18*mm, 14*mm, 15*mm]
    else:
        item_rows = [[
            Paragraph("Sr", cell_header), Paragraph("Description", cell_header),
            Paragraph("Gross Wt", cell_header), Paragraph("Net Wt", cell_header),
            Paragraph("Purity", cell_header), Paragraph("Rate/10g", cell_header),
            Paragraph("Making", cell_header), Paragraph("Total Amount", cell_header)
        ]]
        col_widths = [6*mm, 34*mm, 15*mm, 15*mm, 15*mm, 19*mm, 14*mm, 16*mm]

    for idx, itm in enumerate(items, 1):
        if is_gst:
            item_rows.append([
                Paragraph(str(idx), cell_style), Paragraph(str(itm['desc']), cell_style),
                Paragraph(f"{itm['gross_wt']:.2f} g", cell_style), Paragraph(f"{itm['net_wt']:.2f} g", cell_style),
                Paragraph(str(itm['hsn']), cell_style), Paragraph(str(itm['purity']), cell_style),
                Paragraph(f"Rs. {itm['rate']:,.2f}", cell_style), Paragraph(str(itm['making_display']), cell_style),
                Paragraph(f"Rs. {itm['total']:,.2f}", cell_style)
            ])
        else:
            item_rows.append([
                Paragraph(str(idx), cell_style), Paragraph(str(itm['desc']), cell_style),
                Paragraph(f"{itm['gross_wt']:.2f} g", cell_style), Paragraph(f"{itm['net_wt']:.2f} g", cell_style),
                Paragraph(str(itm['purity']), cell_style), Paragraph(f"Rs. {itm['rate']:,.2f}", cell_style),
                Paragraph(str(itm['making_display']), cell_style), Paragraph(f"Rs. {itm['total']:,.2f}", cell_style)
            ])

    item_table = Table(item_rows, colWidths=col_widths)
    item_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0b2f64')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#b0bec5')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 1.8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.8),
    ]))
    story.append(item_table)
    story.append(Spacer(1, 1.5*mm))

    # Identical Left Block for both GST and Estimate
    cond_line = "1. We are not responsible for any breakage/damage.<br/>" if is_gst else "1. Estimation only. Rates subject to daily market change.<br/>"
    left_block = Paragraph(
        "<b>Terms & Conditions:</b><br/>"
        f"{cond_line}"
        "2. All disputes subject to Jamshedpur jurisdiction.<br/><br/>"
        "<b>For SOVAA JEWELLERS</b><br/><br/><br/>"
        "(Authorised Signatory)",
        cell_style
    )
    
    old_exchange = meta_info.get('old_exchange', 0.0)
    exchange_label = meta_info.get('exchange_label', 'Less Old Exchange')

    if is_gst:
        summary_rows = [
            [left_block, Paragraph("<b>Subtotal:</b>", cell_style), Paragraph(f"Rs. {meta_info['subtotal']:,.2f}", cell_style)],
            ["", Paragraph("<b>CGST (1.5%):</b>", cell_style), Paragraph(f"Rs. {meta_info['cgst']:,.2f}", cell_style)],
            ["", Paragraph("<b>SGST (1.5%):</b>", cell_style), Paragraph(f"Rs. {meta_info['sgst']:,.2f}", cell_style)],
            ["", Paragraph("<b>Gross Total:</b>", cell_style), Paragraph(f"Rs. {meta_info['gross']:,.2f}", cell_style)],
        ]
        if old_exchange > 0:
            summary_rows.append(["", Paragraph(f"<b>{exchange_label}:</b>", cell_style), Paragraph(f"- Rs. {old_exchange:,.2f}", cell_style)])
        
        summary_rows.append(["", Paragraph("<b>Round Off:</b>", cell_style), Paragraph(f"Rs. {meta_info['round_off']:+.2f}", cell_style)])
        summary_rows.append(["", Paragraph("<b>Net Payable:</b>", cell_bold), Paragraph(f"<b>Rs. {meta_info['net_payable']:,.2f}</b>", cell_bold)])
    else:
        summary_rows = [
            [left_block, Paragraph("<b>Total Value:</b>", cell_style), Paragraph(f"Rs. {meta_info['subtotal']:,.2f}", cell_style)],
        ]
        if old_exchange > 0:
            summary_rows.append(["", Paragraph(f"<b>{exchange_label}:</b>", cell_style), Paragraph(f"- Rs. {old_exchange:,.2f}", cell_style)])
        
        summary_rows.append(["", Paragraph("<b>Round Off:</b>", cell_style), Paragraph(f"Rs. {meta_info['round_off']:+.2f}", cell_style)])
        summary_rows.append(["", Paragraph("<b>Net Estimated:</b>", cell_bold), Paragraph(f"<b>Rs. {meta_info['net_payable']:,.2f}</b>", cell_bold)])

    summary_table = Table(summary_rows, colWidths=[68*mm, 35*mm, 31*mm])
    summary_table.setStyle(TableStyle([
        ('SPAN', (0,0), (0, len(summary_rows)-1)),
        ('VALIGN', (0,0), (0,0), 'TOP'),
        ('BOX', (1,0), (-1,-1), 0.5, colors.HexColor('#0b2f64')),
        ('INNERGRID', (1,0), (-1,-1), 0.5, colors.HexColor('#cfd8dc')),
        ('BACKGROUND', (1, len(summary_rows)-1), (2, len(summary_rows)-1), colors.HexColor('#e8f5e9')),
        ('TOPPADDING', (0,0), (-1,-1), 1.2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.2),
    ]))
    story.append(summary_table)

    doc.build(story)

# --- HEADER & LOGOUT ---
head_col1, head_col2 = st.columns([4, 1])
with head_col1:
    st.title("💎 SOVAA JEWELLERS - System")
with head_col2:
    st.write("")
    if st.button("🔒 Logout"):
        st.session_state.authenticated = False
        st.rerun()

tab_billing, tab_search = st.tabs(["📝 New Bill / Estimate", "🔍 Search Customer / History"])

# ==========================================
# TAB 1: NEW BILLING
# ==========================================
with tab_billing:
    mode = st.radio("Select Document Type", ["Tax Invoice (GST)", "Estimate (Without GST)"], horizontal=True)

    if "last_mode" not in st.session_state:
        st.session_state.last_mode = mode
    if mode != st.session_state.last_mode:
        st.session_state.last_mode = mode
        st.session_state.current_doc_no = get_next_number(mode)

    if "current_doc_no" not in st.session_state:
        st.session_state.current_doc_no = get_next_number(mode)

    c_hdr1, c_hdr2 = st.columns([4, 1])
    with c_hdr1:
        st.markdown("##### 👤 Customer Information")
    with c_hdr2:
        if st.button("🧹 New / Clear Customer Form"):
            st.session_state.cust_name_input = ""
            st.session_state.cust_mob_input = ""
            st.session_state.cust_addr_input = "Mango, Jamshedpur"
            st.session_state.cust_gstin_input = ""
            st.rerun()

    col1, col2 = st.columns([1.5, 1.5])
    with col1:
        nc1, nc2 = st.columns([3, 1])
        with nc1:
            doc_number = st.text_input("Invoice / Estimate No", value=st.session_state.current_doc_no)
        with nc2:
            st.write("")
            st.write("")
            if st.button("🔄 Auto", help="Latest serial number fetch karein"):
                st.session_state.current_doc_no = get_next_number(mode)
                st.rerun()
                
        cust_name = st.text_input("Customer Name", key="cust_name_input", placeholder="e.g. Somnath Das")
        cust_address = st.text_input("Address", key="cust_addr_input")
    with col2:
        doc_date = st.date_input("Date", value=datetime.today())
        cust_mob = st.text_input("Mobile No", key="cust_mob_input", placeholder="10 Digit Number")
        if mode == "Tax Invoice (GST)":
            cust_gstin = st.text_input("Party GSTIN (Optional)", key="cust_gstin_input", placeholder="Optional")
        else:
            cust_gstin = ""

    st.markdown("---")
    st.subheader("🛒 Item Details")

    selected_preset = st.selectbox("⚡ Quick Select Item (Optional)", COMMON_ITEMS)

    purity_options = ["22K (916)", "SILVER", "18K (750)", "24K (999)"]
    default_purity_idx = 0
    if "SILVER" in selected_preset:
        default_purity_idx = 1
    elif "GOLD" in selected_preset:
        default_purity_idx = 0

    with st.form("item_entry_form", clear_on_submit=True):
        ic1, ic2, ic3, ic4, ic5, ic6, ic7 = st.columns([2.2, 1.1, 1.1, 1.3, 1.1, 0.8, 1.1])
        
        default_text = ""
        if selected_preset not in ["-- Select Common Item --", "➕ Custom / Other Item"]:
            default_text = selected_preset

        with ic1:
            item_desc = st.text_input("Item Description", value=default_text, placeholder="e.g. GOLD LOCKET")
        with ic2:
            gross_wt = st.number_input("Gross Wt (g)", min_value=0.0, step=0.01, format="%.2f", value=None, placeholder="0.00")
        with ic3:
            net_wt = st.number_input("Net Wt (g)", min_value=0.0, step=0.01, format="%.2f", value=None, placeholder="0.00")
        with ic4:
            rate_10g = st.number_input("Rate (Per 10g)", min_value=0.0, step=10.0, format="%.2f", value=None, placeholder="Rate")
        with ic5:
            making_val = st.number_input("Making Charge", min_value=0.0, step=1.0, value=None, placeholder="Making")
        with ic6:
            making_type = st.selectbox("Unit", ["%", "₹"])
        with ic7:
            purity = st.selectbox("Purity", purity_options, index=default_purity_idx)

        next_item_no = len(st.session_state.items_list) + 1
        submitted = st.form_submit_button(f"➕ Add Item #{next_item_no}", use_container_width=True)
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
                    "desc": item_desc.upper(),
                    "gross_wt": g_wt,
                    "net_wt": net_wt,
                    "hsn": "7113",
                    "purity": purity,
                    "rate": rate_10g,
                    "making_display": display_str,
                    "total": round(total_item_amt, 2)
                })
                st.rerun()
            else:
                st.warning("Kripya Item Description, Net Weight aur Rate sahi se bharein.")

    if st.session_state.items_list:
        df_items = pd.DataFrame(st.session_state.items_list)
        display_cols = ["desc", "gross_wt", "net_wt", "purity", "rate", "making_display", "total"]
        if mode == "Tax Invoice (GST)":
            display_cols.insert(3, "hsn")
        st.table(df_items[display_cols])

        if st.button("🗑️ Clear All Items"):
            st.session_state.items_list = []
            st.rerun()

        subtotal = sum(i["total"] for i in st.session_state.items_list)
        if mode == "Tax Invoice (GST)":
            cgst = round(subtotal * 0.015, 2)
            sgst = round(subtotal * 0.015, 2)
            gross_total = subtotal + cgst + sgst
        else:
            cgst = 0.0
            sgst = 0.0
            gross_total = subtotal

        st.markdown("---")
        sc1, sc2 = st.columns([1.5, 1.5])
        
        with sc1:
            st.subheader("🔄 Old Metal Exchange")
            exchange_metal_type = st.radio(
                "Exchange Type", 
                ["None", "Old Gold", "Old Silver", "Both (Gold & Silver)"], 
                horizontal=True
            )
            
            old_val_input = 0.0
            exchange_display_label = "Less Old Exchange"

            if exchange_metal_type == "Old Gold":
                old_val_input = st.number_input("Purane Sone ka Value (₹)", min_value=0.0, step=100.0, format="%.2f", value=0.0)
                exchange_display_label = "Less Old Gold Exchange"
            elif exchange_metal_type == "Old Silver":
                old_val_input = st.number_input("Purani Chandi ka Value (₹)", min_value=0.0, step=100.0, format="%.2f", value=0.0)
                exchange_display_label = "Less Old Silver Exchange"
            elif exchange_metal_type == "Both (Gold & Silver)":
                c_g, c_s = st.columns(2)
                with c_g:
                    gold_val = st.number_input("Old Gold Value (₹)", min_value=0.0, step=100.0, format="%.2f", value=0.0)
                with c_s:
                    silver_val = st.number_input("Old Silver Value (₹)", min_value=0.0, step=100.0, format="%.2f", value=0.0)
                old_val_input = gold_val + silver_val
                exchange_display_label = "Less Old Gold & Silver Exchange"

        after_exchange = max(0.0, gross_total - old_val_input)
        net_payable = int(math.floor(after_exchange / 10.0) * 10)
        round_off = round(net_payable - after_exchange, 2)

        with sc2:
            st.subheader("📊 Bill Summary")
            if mode == "Tax Invoice (GST)":
                st.markdown(f"**Total Taxable:** Rs. {subtotal:,.2f}")
                st.markdown(f"**CGST (1.5%):** Rs. {cgst:,.2f}")
                st.markdown(f"**SGST (1.5%):** Rs. {sgst:,.2f}")
                st.markdown(f"**Gross Bill Total:** Rs. {gross_total:,.2f}")
            else:
                st.markdown(f"**Total Item Value:** Rs. {subtotal:,.2f}")
            
            if old_val_input > 0:
                st.markdown(f"**{exchange_display_label}:** - Rs. {old_val_input:,.2f}")
            
            st.markdown(f"**Round Off:** Rs. {round_off:,.2f}")
            st.subheader(f"💰 Net Payable: Rs. {net_payable:,.2f}")

        btn_text = f"💾 Generate & Save {mode}"
        if st.button(btn_text, type="primary", use_container_width=True):
            if not cust_name:
                st.error("Customer Name zaroori hai!")
            else:
                formatted_date = doc_date.strftime("%d %b %Y")
                pdf_filename = f"{doc_number}.pdf"
                pdf_filepath = os.path.join(PDF_DIR, pdf_filename)

                meta_info = {
                    "doc_no": doc_number,
                    "date": formatted_date,
                    "customer_name": cust_name.upper(),
                    "customer_mob": cust_mob,
                    "customer_address": cust_address,
                    "customer_gstin": cust_gstin,
                    "subtotal": subtotal,
                    "cgst": cgst,
                    "sgst": sgst,
                    "gross": gross_total,
                    "old_exchange": old_val_input,
                    "exchange_label": exchange_display_label,
                    "round_off": round_off,
                    "net_payable": net_payable
                }

                generate_a5_pdf(mode, meta_info, st.session_state.items_list, pdf_filepath)

                with open(pdf_filepath, "rb") as f:
                    pdf_base64 = base64.b64encode(f.read()).decode('utf-8')

                payload = {
                    "doc_no": doc_number,
                    "date": formatted_date,
                    "type": mode,
                    "customer_name": cust_name.upper(),
                    "customer_mob": str(cust_mob).strip() if cust_mob else "NA",
                    "customer_address": cust_address,
                    "gross": gross_total,
                    "old_exchange": old_val_input,
                    "net_payable": net_payable,
                    "cgst": cgst,
                    "sgst": sgst,
                    "pdf_base64": pdf_base64
                }

                items_summary_str = ", ".join([f"{it['desc']} ({it['net_wt']}g)" for it in st.session_state.items_list])
                db_row = {
                    "Document No": doc_number,
                    "Date": formatted_date,
                    "Customer Name": cust_name.upper(),
                    "Mobile No": str(cust_mob).strip() if cust_mob else "NA",
                    "Address": cust_address,
                    "Party GSTIN": cust_gstin if cust_gstin else "NA",
                    "Type": mode,
                    "Items": items_summary_str,
                    "Gross Amount": gross_total,
                    "Old Exchange": old_val_input,
                    "Net Payable": net_payable,
                    "SGST": sgst,
                    "CGST": cgst
                }
                save_to_database(db_row, mode)

                with st.spinner("☁️ Google Sheet & Drive par sync ho raha hai..."):
                    sync_ok, drive_pdf_url = sync_to_google_sheet(payload)

                # Reset customer info & items automatically for next fresh customer
                st.session_state.cust_name_input = ""
                st.session_state.cust_mob_input = ""
                st.session_state.cust_addr_input = "Mango, Jamshedpur"
                st.session_state.cust_gstin_input = ""
                st.session_state.items_list = []
                st.session_state.current_doc_no = get_next_number(mode)
                
                if sync_ok:
                    st.success(f"✅ {mode} ({doc_number}) safalta-purvak ban gaya aur Google Sheet & Drive par sync ho gaya!")
                    if drive_pdf_url:
                        st.markdown(f"🔗 [Google Drive par PDF dekhein]({drive_pdf_url})")
                else:
                    st.warning(f"⚠️ {mode} ({doc_number}) locally save ho gaya hai, par Google sync mein samay laga.")

                with open(pdf_filepath, "rb") as f:
                    st.download_button(
                        label=f"📄 Download & Print A5 PDF ({doc_number})",
                        data=f,
                        file_name=pdf_filename,
                        mime="application/pdf",
                        use_container_width=True
                    )

# ==========================================
# TAB 2: SEARCH & AUTO-FILL CUSTOMER
# ==========================================
with tab_search:
    st.subheader("🔍 Customer Bill & History Search")
    st.write("Customer ka **Naam** ya **Mobile No** search karein. Naya bill banane ke liye **'⚡ Auto-Fill in New Bill'** par click karein:")

    records = []
    for f_name in ["Sales_Database.xlsx", "Estimate_Database.xlsx"]:
        if os.path.exists(f_name):
            try:
                temp_df = pd.read_excel(f_name)
                if not temp_df.empty:
                    records.append(temp_df)
            except Exception:
                pass

    if records:
        all_data = pd.concat(records, ignore_index=True)
        search_query = st.text_input("🔎 Search by Name or Mobile No", placeholder="e.g. Raju ya 77177...").strip()

        if search_query:
            mask = (
                all_data["Customer Name"].astype(str).str.contains(search_query, case=False, na=False) |
                all_data["Mobile No"].astype(str).str.contains(search_query, case=False, na=False)
            )
            filtered_df = all_data[mask]

            if not filtered_df.empty:
                st.success(f"🎯 Total **{len(filtered_df)}** records mile:")

                for idx, row in filtered_df.iterrows():
                    d_no = str(row.get("Document No", ""))
                    c_name = str(row.get("Customer Name", ""))
                    c_mob = str(row.get("Mobile No", ""))
                    c_addr = str(row.get("Address", "Mango, Jamshedpur"))
                    c_gst = str(row.get("Party GSTIN", ""))
                    if c_gst == "NA":
                        c_gst = ""
                    p_file = os.path.join(PDF_DIR, f"{d_no}.pdf")

                    card = st.container()
                    with card:
                        r_col1, r_col2, r_col3, r_col4 = st.columns([2, 2.5, 1.8, 1.2])
                        with r_col1:
                            st.markdown(f"**{d_no}** ({row.get('Type', '')})")
                            st.caption(f"📅 {row.get('Date', '')}")
                        with r_col2:
                            st.markdown(f"👤 **{c_name}** | 📞 {c_mob}")
                            st.caption(f"📦 Items: {row.get('Items', 'NA')}")
                        with r_col3:
                            st.markdown(f"💰 **₹{row.get('Net Payable', 0):,.2f}**")
                            if st.button("⚡ Auto-Fill in New Bill", key=f"btn_s_fill_{idx}_{d_no}"):
                                st.session_state.cust_name_input = c_name
                                st.session_state.cust_mob_input = "" if c_mob == "NA" else c_mob
                                st.session_state.cust_addr_input = c_addr
                                st.session_state.cust_gstin_input = c_gst
                                st.success(f"✅ {c_name} ki details load ho gayi hain! Upar 'New Bill / Estimate' tab par click karein.")
                        with r_col4:
                            if os.path.exists(p_file):
                                with open(p_file, "rb") as pf:
                                    st.download_button(
                                        label="⬇️ PDF",
                                        data=pf,
                                        file_name=f"{d_no}.pdf",
                                        mime="application/pdf",
                                        key=f"btn_s_dl_{idx}_{d_no}"
                                    )
                        st.divider()
            else:
                st.warning(f"'{search_query}' ke naam ya number se koi record nahi mila.")
        else:
            st.info("💡 Upar search box mein customer ka naam ya mobile number likhein.")
            st.markdown("##### 🕒 Recent 5 Transactions:")
            recent_df = all_data.tail(5).iloc[::-1]
            for idx, row in recent_df.iterrows():
                d_no = str(row.get("Document No", ""))
                c_name = str(row.get("Customer Name", ""))
                c_mob = str(row.get("Mobile No", ""))
                c_addr = str(row.get("Address", "Mango, Jamshedpur"))
                c_gst = str(row.get("Party GSTIN", ""))
                if c_gst == "NA":
                    c_gst = ""
                
                rc1, rc2, rc3 = st.columns([2.5, 2.5, 2])
                with rc1:
                    st.markdown(f"**{d_no}** - {c_name}")
                with rc2:
                    st.markdown(f"📞 {c_mob} | 💰 ₹{row.get('Net Payable', 0):,.2f}")
                with rc3:
                    if st.button("⚡ Use Details", key=f"btn_r_fill_{idx}_{d_no}"):
                        st.session_state.cust_name_input = c_name
                        st.session_state.cust_mob_input = "" if c_mob == "NA" else c_mob
                        st.session_state.cust_addr_input = c_addr
                        st.session_state.cust_gstin_input = c_gst
                        st.success(f"✅ {c_name} ki details load ho gayi hain! Upar 'New Bill / Estimate' tab par click karein.")
                st.divider()
    else:
        st.info("Abhi tak koi transaction record nahi hua hai.")
