import streamlit as st
import pandas as pd
from datetime import datetime
import os
import math
from reportlab.lib.pagesizes import A5
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm, inch

# Page Configuration
st.set_page_config(page_title="SOVAA JEWELLERS - Billing & Estimate", layout="wide", page_icon="💎")

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

# --- APP START AFTER LOGIN ---
PDF_DIR = "Invoices_PDF"
os.makedirs(PDF_DIR, exist_ok=True)

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
    num_col = "Invoice No" if doc_type == "Tax Invoice (GST)" else "Estimate No"
    
    if os.path.exists(db_file):
        try:
            df = pd.read_excel(db_file)
            if not df.empty and num_col in df.columns:
                last_num = str(df[num_col].iloc[-1])
                if "-" in last_num:
                    parts = last_num.split("-")
                    next_val = int(parts[-1]) + 1
                    return f"{prefix}-{datetime.now().year}-{str(next_val).zfill(3)}"
        except Exception:
            pass
    return f"{prefix}-{datetime.now().year}-001"

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
        col_widths = [6*mm, 35*mm, 15*mm, 15*mm, 15*mm, 19*mm, 13*mm, 16*mm]

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

# --- TOP HEADER & LOGOUT ---
head_col1, head_col2 = st.columns([4, 1])
with head_col1:
    st.title("💎 SOVAA JEWELLERS - System")
with head_col2:
    st.write("")
    if st.button("🔒 Logout"):
        st.session_state.authenticated = False
        st.rerun()

mode = st.radio("Select Document Type", ["Tax Invoice (GST)", "Estimate (Without GST)"], horizontal=True)

col1, col2 = st.columns([1.5, 1.5])
with col1:
    doc_number = st.text_input("Invoice / Estimate No", value=get_next_number(mode))
    cust_name = st.text_input("Customer Name", placeholder="e.g. Somnath Das")
    cust_address = st.text_input("Address", value="Mango, Jamshedpur")
with col2:
    doc_date = st.date_input("Date", value=datetime.today())
    cust_mob = st.text_input("Mobile No", placeholder="10 Digit Number")
    if mode == "Tax Invoice (GST)":
        cust_gstin = st.text_input("Party GSTIN (Optional)", placeholder="Optional")
    else:
        cust_gstin = ""

st.markdown("---")
st.subheader("🛒 Item Details")

if "items_list" not in st.session_state:
    st.session_state.items_list = []

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

            db_row = {
                "Document No": doc_number,
                "Date": formatted_date,
                "Customer Name": cust_name.upper(),
                "Mobile No": cust_mob,
                "Type": mode,
                "Total Amount": gross_total,
                "Exchange Type": exchange_metal_type,
                "Old Exchange": old_val_input,
                "Net Payable": net_payable,
                "SGST": sgst,
                "CGST": cgst
            }
            save_to_database(db_row, mode)

            st.success(f"✅ {mode} ({doc_number}) safalta-purvak generate aur save ho gaya!")

            with open(pdf_filepath, "rb") as f:
                st.download_button(
                    label=f"📄 Download & Print A5 PDF ({doc_number})",
                    data=f,
                    file_name=pdf_filename,
                    mime="application/pdf",
                    use_container_width=True
                )
