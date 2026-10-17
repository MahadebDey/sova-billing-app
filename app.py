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

st.markdown("<style>input[type='text']{text-transform:uppercase!important;}</style>", unsafe_allow_html=True)

HALLMARK_REG_NO = "HM/C-5800000504"
FIRM_GSTIN = "20AUJPD1127G1ZE"
FIRM_PHONE = "7717740697"
GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwyLXOi4yCglsCueHm_nVQ9wYjPUI0eKsPlxYxYGqFnwIvcML66W5JWfRCwtqj8WuhU/exec"

PDF_DIR = "Invoices_PDF"
os.makedirs(PDF_DIR, exist_ok=True)
RATES_FILE = "daily_rates.json"

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

if "daily_gold_rate" not in st.session_state:
    rts = load_saved_rates()
    st.session_state.daily_gold_rate = rts["gold_22k"]
    st.session_state.daily_gold_18k_rate = rts["gold_18k"]
    st.session_state.daily_silver_rate = rts["silver"]

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
    "-- Select Common Item --", "GOLD LOCKET", "SILVER LOCKET", "GOLD NATHIYA",
    "SILVER BABY BALA", "SILVER BABY PAYAL", "SILVER KI RING", "SILVER BICHHIYA",
    "GOLD CHAIN", "GOLD RING", "GOLD EARRING / JHUMKA", "GOLD NECKLACE",
    "GOLD BRACELET", "GOLD BANGLE / BALA", "GOLD NOSE PIN", "GOLD PENDANT",
    "GOLD COIN", "SILVER PAYAL / PAJEB", "SILVER CHAIN", "SILVER BRACELET",
    "SILVER COIN", "SILVER UTENSIL / BARTAN", "SILVER MURTI / IDOL", "➕ Custom / Other Item"
]

def get_current_fy_string():
    now = datetime.now()
    yr = now.year
    if now.month < 4:
        s_yr, e_yr = yr - 1, yr
    else:
        s_yr, e_yr = yr, yr + 1
    return f"{str(s_yr)[-2:]}-{str(e_yr)[-2:]}"

def get_next_number(doc_type):
    pfx = "SV" if doc_type == "Tax Invoice (GST)" else "EST"
    fy = get_current_fy_string()
    m_val = 0
    try:
        r = requests.get(GOOGLE_SCRIPT_URL, timeout=3)
        if r.status_code == 200:
            c_data = r.json()
            if c_data.get("status") == "success":
                c_num = c_data.get("last_gst" if doc_type == "Tax Invoice (GST)" else "last_est", 0)
                if c_num and int(c_num) > m_val:
                    m_val = int(c_num)
    except Exception:
        pass

    db = "Sales_Database.xlsx" if doc_type == "Tax Invoice (GST)" else "Estimate_Database.xlsx"
    if os.path.exists(db):
        try:
            df = pd.read_excel(db)
            if not df.empty and "Document No" in df.columns:
                for val in df["Document No"].dropna().astype(str):
                    v = val.strip()
                    if v.startswith(pfx) and fy in v:
                        m = re.findall(r'(\d+)$', v)
                        if m and int(m[0]) > m_val:
                            m_val = int(m[0])
        except Exception:
            pass

    return f"{pfx}-{fy}/{str(m_val + 1).zfill(3)}"

def save_to_database(row_dict, doc_type):
    db = "Sales_Database.xlsx" if doc_type == "Tax Invoice (GST)" else "Estimate_Database.xlsx"
    try:
        if os.path.exists(db):
            df = pd.read_excel(db)
            df = pd.concat([df, pd.DataFrame([row_dict])], ignore_index=True)
        else:
            df = pd.DataFrame([row_dict])
        df.to_excel(db, index=False)
    except Exception:
        pass

def sync_to_google_sheet(payload):
    try:
        res = requests.post(GOOGLE_SCRIPT_URL, json=payload, headers={"Content-Type": "application/json"}, timeout=20)
        if res.status_code == 200:
            d = res.json()
            if d.get("status") == "success":
                return True, d.get("pdf_url"), "Success"
            return False, None, d.get("message", "Script Error")
        return False, None, f"HTTP {res.status_code}"
    except Exception as e:
        return False, None, str(e)

def generate_a5_pdf(doc_type, meta_info, items, pdf_path):
    doc = SimpleDocTemplate(pdf_path, pagesize=A5, leftMargin=6*mm, rightMargin=6*mm, topMargin=1.2*inch, bottomMargin=4*mm)
    story = []
    st_norm = getSampleStyleSheet()['Normal']

    t_style = ParagraphStyle('T', parent=st_norm, fontName='Helvetica-Bold', fontSize=10, leading=12, alignment=1)
    c_txt = ParagraphStyle('C', parent=st_norm, fontName='Helvetica', fontSize=6.2, leading=7.5)
    c_bld = ParagraphStyle('CB', parent=st_norm, fontName='Helvetica-Bold', fontSize=6.2, leading=7.5)
    c_hdr = ParagraphStyle('CH', parent=st_norm, fontName='Helvetica-Bold', fontSize=6.2, leading=7.5, alignment=1)
    c_rt = ParagraphStyle('CR', parent=st_norm, fontName='Helvetica', fontSize=6.2, leading=7.5, alignment=2)
    c_rt_b = ParagraphStyle('CRB', parent=st_norm, fontName='Helvetica-Bold', fontSize=6.2, leading=7.5, alignment=2)
    t_terms = ParagraphStyle('TT', parent=st_norm, fontName='Helvetica', fontSize=5.0, leading=6.2)
    s_style = ParagraphStyle('SS', parent=st_norm, fontName='Helvetica-Bold', fontSize=5.5, leading=7.0)

    is_gst = (doc_type == "Tax Invoice (GST)")
    story.append(Paragraph("<u>TAX INVOICE</u>" if is_gst else "<u>ESTIMATE</u>", t_style))
    story.append(Spacer(1, 1.2*mm))

    c_mob = meta_info['customer_mob'] or "NA"
    c_pan = meta_info.get('customer_pan') or "NA"
    c_id = meta_info.get('customer_id') or "NA"

    if is_gst:
        meta_data = [
            [Paragraph(f"<b>Date:</b> {meta_info['date']}", c_txt), Paragraph(f"<b>GSTIN:</b> {FIRM_GSTIN}", c_txt)],
            [Paragraph(f"<b>Invoice No:</b> {meta_info['doc_no']}", c_txt), Paragraph(f"<b>Hallmark Reg:</b> {HALLMARK_REG_NO}", c_txt)],
            [Paragraph(f"<b>Customer Name:</b> {meta_info['customer_name']}", c_txt), Paragraph(f"<b>Payment Mode:</b> {meta_info.get('payment_mode', 'UPI/Cash')}", c_txt)],
            [Paragraph(f"<b>Customer Mob:</b> {c_mob}", c_txt), Paragraph(f"<b>Govt ID / Aadhaar:</b> {c_id}", c_txt)],
            [Paragraph(f"<b>Customer PAN:</b> {c_pan}", c_txt), Paragraph("<b>State:</b> 20 - JHARKHAND", c_txt)],
            [Paragraph(f"<b>Address:</b> {meta_info['customer_address']}", c_txt), Paragraph(f"<b>Party GSTIN:</b> {meta_info['customer_gstin'] or 'NA'}", c_txt)]
        ]
    else:
        meta_data = [
            [Paragraph(f"<b>Date:</b> {meta_info['date']}", c_txt), Paragraph(f"<b>Hallmark Reg:</b> {HALLMARK_REG_NO}", c_txt)],
            [Paragraph(f"<b>Estimate No:</b> {meta_info['doc_no']}", c_txt), Paragraph(f"<b>Payment Mode:</b> {meta_info.get('payment_mode', 'UPI/Cash')}", c_txt)],
            [Paragraph(f"<b>Customer Name:</b> {meta_info['customer_name']}", c_txt), Paragraph(f"<b>Customer Mob:</b> {c_mob}", c_txt)],
            [Paragraph(f"<b>Govt ID / Aadhaar:</b> {c_id}", c_txt), Paragraph(f"<b>Customer PAN:</b> {c_pan}", c_txt)],
            [Paragraph(f"<b>Address:</b> {meta_info['customer_address']}", c_txt), Paragraph("<b>State:</b> JHARKHAND", c_txt)]
        ]

    m_tbl = Table(meta_data, colWidths=[68*mm, 68*mm])
    m_tbl.setStyle(TableStyle([('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#444')), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('TOPPADDING', (0,0), (-1,-1), 0.8), ('BOTTOMPADDING', (0,0), (-1,-1), 0.8)]))
    story.append(m_tbl)
    story.append(Spacer(1, 1.2*mm))

    if is_gst:
        item_rows = [[Paragraph("Sr", c_hdr), Paragraph("Description & HUID", c_hdr), Paragraph("Gross Wt", c_hdr), Paragraph("Net Wt", c_hdr), Paragraph("HSN", c_hdr), Paragraph("Purity", c_hdr), Paragraph("Rate/10g", c_hdr), Paragraph("Making", c_hdr), Paragraph("Total", c_hdr)]]
        c_w = [6*mm, 32*mm, 13*mm, 13*mm, 10*mm, 14*mm, 18*mm, 13*mm, 17*mm]
    else:
        item_rows = [[Paragraph("Sr", c_hdr), Paragraph("Description & HUID", c_hdr), Paragraph("Gross Wt", c_hdr), Paragraph("Net Wt", c_hdr), Paragraph("Purity", c_hdr), Paragraph("Rate/10g", c_hdr), Paragraph("Making", c_hdr), Paragraph("Total Amount", c_hdr)]]
        c_w = [6*mm, 36*mm, 15*mm, 15*mm, 15*mm, 19*mm, 13*mm, 17*mm]

    for idx, itm in enumerate(items, 1):
        d_txt = f"<b>{itm['desc']}</b>"
        if itm.get('huid') and itm['huid'] != "NA":
            d_txt += f"<br/><font size='5.2' color='#444'>HUID: {itm['huid']}</font>"
        if is_gst:
            item_rows.append([Paragraph(str(idx), c_txt), Paragraph(d_txt, c_txt), Paragraph(f"{itm['gross_wt']:.2f}g", c_txt), Paragraph(f"{itm['net_wt']:.2f}g", c_txt), Paragraph(str(itm['hsn']), c_txt), Paragraph(str(itm['purity']), c_txt), Paragraph(f"Rs.{itm['rate']:,.2f}", c_txt), Paragraph(str(itm['making_display']), c_txt), Paragraph(f"Rs.{itm['total']:,.2f}", c_txt)])
        else:
            item_rows.append([Paragraph(str(idx), c_txt), Paragraph(d_txt, c_txt), Paragraph(f"{itm['gross_wt']:.2f}g", c_txt), Paragraph(f"{itm['net_wt']:.2f}g", c_txt), Paragraph(str(itm['purity']), c_txt), Paragraph(f"Rs.{itm['rate']:,.2f}", c_txt), Paragraph(str(itm['making_display']), c_txt), Paragraph(f"Rs.{itm['total']:,.2f}", c_txt)])

    i_tbl = Table(item_rows, colWidths=c_w)
    i_tbl.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#eee')), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#555')), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('TOPPADDING', (0,0), (-1,-1), 1.2), ('BOTTOMPADDING', (0,0), (-1,-1), 1.2)]))
    story.append(i_tbl)
    story.append(Spacer(1, max(4*mm, (3-len(items))*4.5*mm)))

    l_blk = Paragraph("<b>Terms & Conditions:</b><br/>1. Gold purity certified as per BIS guidelines.<br/>2. Breakage responsibility ceases once delivered.<br/>3. Making charges & GST are non-refundable.<br/>4. Old metal exchange confirms lawful property.<br/>5. Disputes subject to Jamshedpur jurisdiction.", t_terms)
    old_ex = meta_info.get('old_exchange', 0.0)
    lbl_ex = meta_info.get('exchange_label', 'Less Old Exchange')
    lbl_fin = meta_info.get('final_balance_label', 'Net Payable')

    if is_gst:
        s_rows = [
            [l_blk, Paragraph("<b>Subtotal:</b>", c_txt), Paragraph(f"Rs.{meta_info['subtotal']:,.2f}", c_rt)],
            ["", Paragraph("<b>CGST (1.5%):</b>", c_txt), Paragraph(f"Rs.{meta_info['cgst']:,.2f}", c_rt)],
            ["", Paragraph("<b>SGST (1.5%):</b>", c_txt), Paragraph(f"Rs.{meta_info['sgst']:,.2f}", c_rt)],
            ["", Paragraph("<b>Gross Total:</b>", c_txt), Paragraph(f"Rs.{meta_info['gross']:,.2f}", c_rt)],
        ]
        if old_ex > 0:
            s_rows.append(["", Paragraph(f"<b>{lbl_ex}:</b>", c_txt), Paragraph(f"- Rs.{old_ex:,.2f}", c_rt)])
        s_rows.append(["", Paragraph("<b>Round Off:</b>", c_txt), Paragraph(f"Rs.{meta_info['round_off']:+.2f}", c_rt)])
        s_rows.append(["", Paragraph(f"<b>{lbl_fin}:</b>", c_bld), Paragraph(f"<b>Rs.{meta_info['net_payable']:,.2f}</b>", c_rt_b)])
    else:
        s_rows = [[l_blk, Paragraph("<b>Total Value:</b>", c_txt), Paragraph(f"Rs.{meta_info['subtotal']:,.2f}", c_rt)]]
        if old_ex > 0:
            s_rows.append(["", Paragraph(f"<b>{lbl_ex}:</b>", c_txt), Paragraph(f"- Rs.{old_ex:,.2f}", c_rt)])
        s_rows.append(["", Paragraph("<b>Round Off:</b>", c_txt), Paragraph(f"Rs.{meta_info['round_off']:+.2f}", c_rt)])
        s_rows.append(["", Paragraph(f"<b>{lbl_fin}:</b>", c_bld), Paragraph(f"<b>Rs.{meta_info['net_payable']:,.2f}</b>", c_rt_b)])

    sum_tbl = Table(s_rows, colWidths=[68*mm, 36*mm, 32*mm])
    sum_tbl.setStyle(TableStyle([('SPAN', (0,0), (0, len(s_rows)-1)), ('VALIGN', (0,0), (-1,-1), 'TOP'), ('BOX', (1,0), (-1,-1), 0.5, colors.HexColor('#444')), ('INNERGRID', (1,0), (-1,-1), 0.4, colors.HexColor('#777')), ('TOPPADDING', (0,0), (-1,-1), 0.8), ('BOTTOMPADDING', (0,0), (-1,-1), 0.8)]))
    story.append(sum_tbl)
    story.append(Spacer(1, 3.5*mm))

    sig_data = [[
        Paragraph("Customer Signature / अंगूठा<br/><br/><br/>______________________<br/><font size='4.5' color='#555'>(Lawful owner of exchange metal)</font>", s_style),
        Paragraph("For SOVAA JEWELLERS<br/><br/><br/>______________________<br/><font size='4.5' color='#555'>(Authorised Signatory)</font>", s_style)
    ]]
    sig_tbl = Table(sig_data, colWidths=[68*mm, 68*mm])
    sig_tbl.setStyle(TableStyle([('ALIGN', (0,0), (0,0), 'LEFT'), ('ALIGN', (1,0), (1,0), 'RIGHT'), ('VALIGN', (0,0), (-1,-1), 'TOP'), ('TOPPADDING', (0,0), (-1,-1), 0.5), ('BOTTOMPADDING', (0,0), (-1,-1), 0.5)]))
    story.append(sig_tbl)

    doc.build(story)

def render_mobile_print_button(pdf_base64, doc_no):
    html = f"""<div style="margin:4px 0;"><button style="background:#222;color:#fff;padding:12px;font-size:16px;font-weight:bold;border:none;border-radius:8px;width:100%;cursor:pointer;" onclick="pDoc()">🖨️ Quick Print ({doc_no}) to Epson L3350</button><script>function pDoc(){{var b='{pdf_base64}';var c=atob(b);var n=new Array(c.length);for(var i=0;i<c.length;i++){{n[i]=c.charCodeAt(i);}}var blob=new Blob([new Uint8Array(n)],{{type:'application/pdf'}});var u=URL.createObjectURL(blob);var w=window.open(u);if(w){{w.focus();w.print();}}else{{window.location.href=u;}}}}</script></div>"""
    components.html(html, height=60)

st.title("💎 SOVAA JEWELLERS")
st.write(f"**Hallmark Reg No:** `{HALLMARK_REG_NO}` | **GSTIN:** `{FIRM_GSTIN}`")

with st.expander("📈 **Today's Market Rates (Aaj ka Gold & Silver Bhav)**", expanded=False):
    r1, r2, r3 = st.columns(3)
    with r1:
        ng22 = st.number_input("🪙 22K Gold Rate (Per 10g) ₹", min_value=0.0, step=100.0, format="%.2f", value=float(st.session_state.daily_gold_rate))
    with r2:
        ng18 = st.number_input("🪙 18K Gold Rate (Per 10g) ₹", min_value=0.0, step=100.0, format="%.2f", value=float(st.session_state.daily_gold_18k_rate))
    with r3:
        nsil = st.number_input("🥈 Silver Rate (Per 10g) ₹", min_value=0.0, step=10.0, format="%.2f", value=float(st.session_state.daily_silver_rate))
    if st.button("💾 Save Rates (Bhav Save Karein)"):
        st.session_state.daily_gold_rate = ng22
        st.session_state.daily_gold_18k_rate = ng18
        st.session_state.daily_silver_rate = nsil
        save_rates_to_file(ng22, ng18, nsil)
        st.success("✅ Rates save ho gaye!")
        st.rerun()

tab_bill, tab_srch = st.tabs(["📝 New Bill / Estimate", "🔍 Search Customer / History"])

with tab_bill:
    mode = st.radio("Select Document Type", ["Tax Invoice (GST)", "Estimate (Without GST)"], index=1, horizontal=True)

    if "current_mode" not in st.session_state or st.session_state.current_mode != mode:
        st.session_state.current_mode = mode
        st.session_state.active_doc_number = get_next_number(mode)
    if "active_doc_number" not in st.session_state:
        st.session_state.active_doc_number = get_next_number(mode)

    c_h1, c_h2 = st.columns([4, 1])
    with c_h1:
        st.markdown("##### 👤 Customer Information")
    with c_h2:
        if st.button("🧹 Clear Form / New Customer"):
            st.session_state.c_name = ""
            st.session_state.c_mob = ""
            st.session_state.c_addr = "MANGO, JAMSHEDPUR"
            st.session_state.c_id = ""
            st.session_state.c_pan = ""
            st.session_state.c_gstin = ""
            st.session_state.items_list = []
            st.session_state.form_reset_count += 1
            st.session_state.active_doc_number = get_next_number(mode)
            st.rerun()

    rk = st.session_state.form_reset_count
    mt = "gst" if mode == "Tax Invoice (GST)" else "est"

    col1, col2 = st.columns([1.5, 1.5])
    with col1:
        nc1, nc2 = st.columns([3, 1])
        with nc1:
            doc_number = st.text_input("Invoice / Estimate No", value=st.session_state.active_doc_number, key=f"doc_{mt}_{rk}")
        with nc2:
            st.write("")
            st.write("")
            if st.button("🔄 Auto"):
                st.session_state.active_doc_number = get_next_number(mode)
                st.rerun()
        cust_name_in = st.text_input("Customer Name", value=st.session_state.c_name, key=f"cn_{rk}", placeholder="e.g. SOMNATH DAS")
        cust_address_in = st.text_input("Address", value=st.session_state.c_addr, key=f"ca_{rk}")
        cust_id_in = st.text_input("Aadhaar / Voter ID (Exchange ke liye anivarya)", value=st.session_state.c_id, key=f"cid_{rk}", placeholder="e.g. Aadhaar / Voter No")
    with col2:
        doc_date = st.date_input("Date", value=datetime.today(), key=f"cd_{rk}")
        cust_mob = st.text_input("Mobile No", value=st.session_state.c_mob, key=f"cm_{rk}", placeholder="10 Digit Number")
        cust_pan_in = st.text_input("Customer PAN (Mandatory if ≥ ₹2 Lakh)", value=st.session_state.c_pan, key=f"cp_{rk}", placeholder="e.g. ABCDE1234F")
        pmode_sel = st.selectbox("Payment Mode", ["UPI / QR (PhonePe/GPay)", "Cash", "Bank Transfer / Card", "Multiple Modes"], key=f"pm_{rk}")
        cust_gstin_in = st.text_input("Party GSTIN (Optional)", value=st.session_state.c_gstin, key=f"cg_{rk}", placeholder="OPTIONAL") if mode == "Tax Invoice (GST)" else ""

    cust_name = cust_name_in.strip().upper()
    cust_address = cust_address_in.strip().upper()
    cust_id = cust_id_in.strip().upper()
    cust_pan = cust_pan_in.strip().upper()
    cust_gstin = cust_gstin_in.strip().upper()

    st.markdown("---")
    st.subheader("🛒 Item Details")
    sel_preset = st.selectbox("⚡ Quick Select Item (Optional)", COMMON_ITEMS)

    p_idx = 2 if "SILVER" in sel_preset else 0
    p1, p2 = st.columns([1.5, 2.5])
    with p1:
        purity_choice = st.selectbox("Purity", ["22K (916)", "18K (750)", "SILVER", "24K (999)"], index=p_idx, key="pur_sel")
    with p2:
        st.write("")
        st.caption(f"🛡️ Hallmarking: {purity_choice} under {HALLMARK_REG_NO}")

    if purity_choice == "18K (750)":
        auto_rt = st.session_state.daily_gold_18k_rate
    elif purity_choice == "SILVER":
        auto_rt = st.session_state.daily_silver_rate
    else:
        auto_rt = st.session_state.daily_gold_rate

    with st.form("itm_form", clear_on_submit=True):
        i1, i2, i3, i4, i5, i6, i7 = st.columns([2.2, 1.2, 1.1, 1.1, 1.3, 1.1, 0.8])
        d_val = "" if sel_preset in ["-- Select Common Item --", "➕ Custom / Other Item"] else sel_preset
        with i1:
            i_desc = st.text_input("Item Description", value=d_val, placeholder="e.g. GOLD LOCKET")
        with i2:
            i_huid = st.text_input("HUID Code (Optional)", placeholder="e.g. AB1234")
        with i3:
            g_wt = st.number_input("Gross Wt (g)", min_value=0.0, step=0.01, format="%.2f", value=None, placeholder="0.00")
        with i4:
            n_wt = st.number_input("Net Wt (g)", min_value=0.0, step=0.01, format="%.2f", value=None, placeholder="0.00")
        with i5:
            r_10g = st.number_input("Rate (Per 10g)", min_value=0.0, step=10.0, format="%.2f", value=auto_rt)
        with i6:
            m_val = st.number_input("Making Charge", min_value=0.0, step=1.0, value=None, placeholder="Making")
        with i7:
            m_typ = st.selectbox("Unit", ["%", "₹"])

        sub_btn = st.form_submit_button(f"➕ Add Item #{len(st.session_state.items_list) + 1} ({purity_choice})", use_container_width=True)
        if sub_btn:
            if i_desc and n_wt and n_wt > 0 and r_10g and r_10g > 0:
                val_m = m_val if m_val is not None else 0.0
                gw = g_wt if (g_wt and g_wt > 0) else n_wt
                m_cost = (n_wt * r_10g) / 10.0
                if m_typ == "%":
                    m_amt = m_cost * (val_m / 100.0)
                    disp_m = f"{val_m}%"
                else:
                    m_amt = val_m
                    disp_m = f"Rs.{val_m:,.0f}"

                st.session_state.items_list.append({
                    "desc": i_desc.strip().upper(),
                    "huid": i_huid.strip().upper() if i_huid else "NA",
                    "gross_wt": gw, "net_wt": n_wt, "hsn": "7113",
                    "purity": purity_choice, "rate": r_10g,
                    "making_display": disp_m, "total": round(m_cost + m_amt, 2)
                })
                st.rerun()
            else:
                st.warning("Kripya Item Description, Net Weight aur Rate bharein.")

    if st.session_state.items_list:
        df_i = pd.DataFrame(st.session_state.items_list)
        cols_show = ["desc", "huid", "gross_wt", "net_wt", "purity", "rate", "making_display", "total"]
        if mode == "Tax Invoice (GST)":
            cols_show.insert(4, "hsn")
        st.table(df_i[cols_show])

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
            ex_typ = st.radio("Exchange Type", ["None", "Old Gold", "Old Silver", "Both (Gold & Silver)"], horizontal=True)
            old_val = 0.0
            ex_lbl = "Less Old Exchange"
            if ex_typ == "Old Gold":
                old_val = st.number_input("Purane Sone ka Value (₹)", min_value=0.0, step=100.0, format="%.2f", value=0.0)
                ex_lbl = "Less Old Gold Exchange"
            elif ex_typ == "Old Silver":
                old_val = st.number_input("Purani Chandi ka Value (₹)", min_value=0.0, step=100.0, format="%.2f", value=0.0)
                ex_lbl = "Less Old Silver Exchange"
            elif ex_typ == "Both (Gold & Silver)":
                cg_col, cs_col = st.columns(2)
                with cg_col:
                    gv = st.number_input("Old Gold (₹)", min_value=0.0, step=100.0, format="%.2f", value=0.0)
                with cs_col:
                    sv = st.number_input("Old Silver (₹)", min_value=0.0, step=100.0, format="%.2f", value=0.0)
                old_val = gv + sv
                ex_lbl = "Less Old Gold & Silver Exchange"

        net_diff = gross_total - old_val
        if net_diff < 0:
            is_ref = True
            abs_d = abs(net_diff)
            f_amt = int(math.floor(abs_d / 10.0) * 10)
            r_off = round(f_amt - abs_d, 2)
            b_lbl = "Refund / Return to Customer"
            b_icon = "🟢"
        else:
            is_ref = False
            f_amt = int(math.floor(net_diff / 10.0) * 10)
            r_off = round(f_amt - net_diff, 2)
            b_lbl = "Net Payable"
            b_icon = "💰"

        with sc2:
            st.subheader("📊 Bill Summary")
            if mode == "Tax Invoice (GST)":
                st.markdown(f"**Total Taxable:** Rs. {subtotal:,.2f}")
                st.markdown(f"**CGST (1.5%):** Rs. {cgst:,.2f}")
                st.markdown(f"**SGST (1.5%):** Rs. {sgst:,.2f}")
                st.markdown(f"**Gross Bill Total:** Rs. {gross_total:,.2f}")
            else:
                st.markdown(f"**Total Item Value:** Rs. {subtotal:,.2f}")
            if old_val > 0:
                st.markdown(f"**{ex_lbl}:** - Rs. {old_val:,.2f}")
            st.markdown(f"**Round Off:** Rs. {r_off:,.2f}")
            st.subheader(f"{b_icon} {b_lbl}: Rs. {f_amt:,.2f}")

        if f_amt >= 200000 and not cust_pan:
            st.warning("⚠️ **Legal Alert:** Bill ₹2,00,000 se zyada hai. Customer PAN Card likhna anivarya hai.")
        if old_val > 0 and not cust_id:
            st.info("💡 **Legal Tip:** Purana sona/chandi lene par Customer ka Aadhaar ya Voter ID number likhna zaroori hai.")

        if st.button(f"💾 Generate & Save {mode}", type="primary", use_container_width=True):
            if not cust_name:
                st.error("Customer Name zaroori hai!")
            else:
                f_date = doc_date.strftime("%d %b %Y")
                s_pfx = str(doc_number).replace('/', '_')
                p_name = f"{s_pfx}.pdf"
                p_path = os.path.join(PDF_DIR, p_name)

                m_info = {
                    "doc_no": doc_number, "date": f_date, "customer_name": cust_name,
                    "customer_mob": cust_mob, "customer_address": cust_address,
                    "customer_id": cust_id, "customer_pan": cust_pan, "customer_gstin": cust_gstin,
                    "payment_mode": pmode_sel, "subtotal": subtotal, "cgst": cgst, "sgst": sgst,
                    "gross": gross_total, "old_exchange": old_val, "exchange_label": ex_lbl,
                    "round_off": r_off, "net_payable": f_amt, "final_balance_label": b_lbl
                }

                generate_a5_pdf(mode, m_info, st.session_state.items_list, p_path)

