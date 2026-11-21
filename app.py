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
st.set_page_config(
    page_title="SOVAA JEWELLERS - Billing & Sync",
    layout="wide",
    page_icon="💎"
)

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
GOOGLE_SCRIPT_URL = (
    "https://script.google.com/macros/s/"
    "AKfycbwyLXOi4yCglsCueHm_nVQ9wYjPUI0eKsPlxYxYGqFnwIvcML66W5JWfRCwtqj8WuhU/exec"
)

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
        data = {
            "gold_22k": float(g22),
            "gold_18k": float(g18),
            "silver": float(sil)
        }
        with open(RATES_FILE, "w") as f:
            json.dump(data, f)
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
    "-- Select Common Item --",
    "GOLD LOCKET",
    "SILVER LOCKET",
    "GOLD NATHIYA",
    "SILVER BABY BALA",
    "SILVER BABY PAYAL",
    "SILVER RING",
    "SILVER BICHHIYA",
    "GOLD CHAIN",
    "GOLD RING",
    "GOLD TOPS / EARING",
    "GOLD NOSE PIN",
    "GOLD PENDANT",
    "SILVER PAYAL",
    "SILVER CHAIN",
    "SILVER BRACELET",
    "SILVER COIN",
    "SILVER MURTI / IDOL",
    "➕ Custom / Other Item"
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
                            raw_addr = re.sub(
                                r'\s*\(Aadhaar:\s*[0-9X\s-]+\)', '',
                                raw_addr, flags=re.IGNORECASE
                            ).strip()

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
                key_name = "last_gst" if doc_type == "Tax Invoice (GST)" else "last_est"
                c_num = c_data.get(key_name, 0)
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

    next_seq = str(max_val + 1).zfill(3)
    return f"{prefix}-{fy}/{next_seq}"

def save_to_database(row_dict, doc_type, gstr1_row=None):
    db = "Sales_Database.xlsx" if doc_type == "Tax Invoice (GST)" else "Estimate_Database.xlsx"
    try:
        if os.path.exists(db):
            df = pd.read_excel(db)
            doc_str = str(row_dict["Document No"]).strip()
            df = df[df["Document No"].astype(str).str.strip() != doc_str]
            df = pd.concat([df, pd.DataFrame([row_dict])], ignore_index=True)
        else:
            df = pd.DataFrame([row_dict])
        
        df["_sort_dt"] = pd.to_datetime(df["Date"], format="%d %b %Y", errors="coerce")
        df = df.sort_values(
            by=["_sort_dt", "Document No"], na_position="last"
        ).drop(columns=["_sort_dt"])
        df.to_excel(db, index=False)
    except Exception:
        pass

    if doc_type == "Tax Invoice (GST)" and gstr1_row:
        try:
            if os.path.exists(GST_FILING_FILE):
                gdf = pd.read_excel(GST_FILING_FILE)
                inv_str = str(gstr1_row["Invoice No"]).strip()
                gdf = gdf[gdf["Invoice No"].astype(str).str.strip() != inv_str]
                gdf = pd.concat([gdf, pd.DataFrame([gstr1_row])], ignore_index=True)
            else:
                gdf = pd.DataFrame([gstr1_row])
            
            gdf["_sort_dt"] = pd.to_datetime(gdf["Invoice Date"], format="%d-%m-%Y", errors="coerce")
            gdf = gdf.sort_values(
                by=["_sort_dt", "Invoice No"], na_position="last"
            ).drop(columns=["_sort_dt"])
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
        requests.post(
            GOOGLE_SCRIPT_URL,
            json={"action": "delete", "doc_no": doc_clean},
            timeout=5
        )
    except Exception:
        pass

    return deleted_any

def sync_to_google_sheet(payload):
    try:
        res = requests.post(
            GOOGLE_SCRIPT_URL,
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=12
        )
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

    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=12,
        alignment=1
    )
    cell_meta = ParagraphStyle(
        'CellMeta',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=6.2,
        leading=7.5
    )
    cell_style = ParagraphStyle(
        'CellText',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=item_font,
        leading=item_leading
    )
    cell_header = ParagraphStyle(
        'CellHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=item_font,
        leading=item_leading,
        alignment=1,
        textColor=colors.HexColor('#000000')
    )
    cell_right = ParagraphStyle(
        'CellRight',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=5.8,
        leading=7.0,
        alignment=2
    )
    cell_right_bold = ParagraphStyle(
        'CellRightBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=6.2,
        leading=7.4,
        alignment=2
    )
    terms_style = ParagraphStyle(
        'TermsText',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=terms_font,
        leading=terms_leading
    )
    sign_style = ParagraphStyle(
        'SignText',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=5.8,
        leading=7.2
    )

    is_gst = (doc_type == "Tax Invoice (GST)")
    doc_title_html = "<u>TAX INVOICE</u>" if is_gst else "<u>ESTIMATE</u>"
    story.append(Paragraph(doc_title_html, title_style))
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

    p_mode = meta_info.get('payment_mode', 'UPI/Cash')

    if is_gst:
        p_gst = meta_info['customer_gstin'] or 'NA'
        meta_data = [
            [
                Paragraph(f"<b>Date:</b> {meta_info['date']}", cell_meta),
                Paragraph(f"<b>GSTIN:</b> {FIRM_GSTIN}", cell_meta)
            ],
            [
                Paragraph(f"<b>Invoice No:</b> {meta_info['doc_no']}", cell_meta),
                Paragraph(f"<b>Hallmark Reg:</b> {HALLMARK_REG_NO}", cell_meta)
            ],
            [
                Paragraph(cust_name_pan, cell_meta),
                Paragraph(f"<b>Payment Mode:</b> {p_mode}", cell_meta)
            ],
            [
                Paragraph(f"<b>Customer Mob:</b> {c_mob}", cell_meta),
                Paragraph("<b>State:</b> 20 - JHARKHAND", cell_meta)
            ],
            [
                Paragraph(cust_addr_aadhaar, cell_meta),
                Paragraph(f"<b>Party GSTIN:</b> {p_gst}", cell_meta)
            ]
        ]
    else:
        meta_data = [
            [
                Paragraph(f"<b>Date:</b> {meta_info['date']}", cell_meta),
                Paragraph(f"<b>Hallmark Reg:</b> {HALLMARK_REG_NO}", cell_meta)
            ],
            [
                Paragraph(f"<b>Estimate No:</b> {meta_info['doc_no']}", cell_meta),
                Paragraph(f"<b>Payment Mode:</b> {p_mode}", cell_meta)
            ],
            [
                Paragraph(cust_name_pan, cell_meta),
                Paragraph(f"<b>Customer Mob:</b> {c_mob}", cell_meta)
            ],
            [
                Paragraph(cust_addr_aadhaar, cell_meta),
                Paragraph("<b>State:</b> JHARKHAND", cell_meta)
            ]
        ]

    meta_table = Table(meta_data, colWidths=[70 * mm, 66 * mm])
    meta_table.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor('#444')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 0.7),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0.7),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 1.2 * mm))

    # Items Table
    if is_gst:
        item_rows = [[
            Paragraph("Sr", cell_header),
            Paragraph("Description & HUID", cell_header),
            Paragraph("Gross Wt", cell_header),
            Paragraph("Net Wt", cell_header),
            Paragraph("HSN", cell_header),
            Paragraph("Purity", cell_header),
            Paragraph("Rate/10g", cell_header),
            Paragraph("Making", cell_header),
            Paragraph("Total", cell_header)
        ]]
        col_widths = [6*mm, 32*mm, 13*mm, 13*mm, 10*mm, 14*mm, 18*mm, 13*mm, 17*mm]
    else:
        item_rows = [[
            Paragraph("Sr", cell_header),
            Paragraph("Description & HUID", cell_header),
            Paragraph("Gross Wt", cell_header),
            Paragraph("Net Wt", cell_header),
            Paragraph("Purity", cell_header),
            Paragraph("Rate/10g", cell_header),
            Paragraph("Making", cell_header),
            Paragraph("Total Amount", cell_header)
        ]]
        col_widths = [6*mm, 36*mm, 15*mm, 15*mm, 15*mm, 19*mm, 13*mm, 17*mm]

    for idx, itm in enumerate(items, 1):
        desc_huid_text = f"<b>{itm['desc']}</b>"
        if itm.get('huid') and itm['huid'] != "NA":
            h_font = round(item_font - 0.8, 1)
            desc_huid_text += f"<br/><font size='{h_font}' color='#333'>HUID: {itm['huid']}</font>"

        if is_gst:
            item_rows.append([
                Paragraph(str(idx), cell_style),
                Paragraph(desc_huid_text, cell_style),
                Paragraph(f"{itm['gross_wt']:.2f}g", cell_style),
                Paragraph(f"{itm['net_wt']:.2f}g", cell_style),
                Paragraph(str(itm['hsn']), cell_style),
                Paragraph(str(itm['purity']), cell_style),
                Paragraph(f"Rs.{itm['rate']:,.2f}", cell_style),
                Paragraph(str(itm['making_display']), cell_style),
                Paragraph(f"Rs.{itm['total']:,.2f}", cell_style)
            ])
        else:
            item_rows.append([
                Paragraph(str(idx), cell_style),
                Paragraph(desc_huid_text, cell_style),
                Paragraph(f"{itm['gross_wt']:.2f}g", cell_style),
                Paragraph(f"{itm['net_wt']:.2f}g", cell_style),
                Paragraph(str(itm['purity']), cell_style),
                Paragraph(f"Rs.{itm['rate']:,.2f}", cell_style),
                Paragraph(str(itm['making_display']), cell_style),
                Paragraph(f"Rs.{itm['total']:,.2f}", cell_style)
            ])

    item_table = Table(item_rows, colWidths=col_widths)
    item_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#EAEAEA')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#555')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), item_pad),
        ('BOTTOMPADDING', (0, 0), (-1, -1), item_pad),
    ]))
    story.append(item_table)

    story.append(Spacer(1, dynamic_bottom_gap))

    # Sub-Table for Signatures
    sig_subtable = Table([
        [
            Paragraph("Customer Signature<br/><br/><br/>", sign_style),
            "",
            Paragraph("For SOVAA JEWELLERS<br/><br/><br/>(Authorised Signatory)", sign_style)
        ]
    ], colWidths=[27 * mm, 12.7 * mm, 29 * mm])
    sig_subtable.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
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
            [
                left_block,
                Paragraph("<b>Subtotal:</b>", cell_right),
                Paragraph(f"Rs.{meta_info['subtotal']:,.2f}", cell_right)
            ],
            [
                "",
                Paragraph("<b>CGST (1.5%):</b>", cell_right),
                Paragraph(f"Rs.{meta_info['cgst']:,.2f}", cell_right)
            ],
            [
                "",
                Paragraph("<b>SGST (1.5%):</b>", cell_right),
                Paragraph(f"Rs.{meta_info['sgst']:,.2f}", cell_right)
            ],
            [
                "",
                Paragraph("<b>Gross Total:</b>", cell_right),
                Paragraph(f"Rs.{meta_info['gross']:,.2f}", cell_right)
            ],
        ]
        if old_exchange > 0:
            summary_rows.append([
                "",
                Paragraph(f"<b>{exchange_label}:</b>", cell_right),
                Paragraph(f"- Rs.{old_exchange:,.2f}", cell_right)
            ])
        summary_rows.append([
            "",
            Paragraph("<b>Round Off:</b>", cell_right),
            Paragraph(f"Rs.{meta_info['round_off']:+.2f}", cell_right)
        ])
        summary_rows.append([
            "",
            Paragraph(f"<b>{final_balance_label}:</b>", cell_right_bold),
            Paragraph(f"<b>Rs.{meta_info['net_payable']:,.2f}</b>", cell_right_bold)
        ])
    else:
        summary_rows = [
            [
                left_block,
                Paragraph("<b>Total Value:</b>", cell_right),
                Paragraph(f"Rs.{meta_info['subtotal']:,.2f}", cell_right)
            ],
        ]
        if old_exchange > 0:
            summary_rows.append([
                "",
                Paragraph(f"<b>{exchange_label}:</b>", cell_right),
                Paragraph(f"- Rs.{old_exchange:,.2f}", cell_right)
            ])
        summary_rows.append([
            "",
            Paragraph("<b>Round Off:</b>", cell_right),
            Paragraph(f"Rs.{meta_info['round_off']:+.2f}", cell_right)
        ])
        summary_rows.append([
            "",
            Paragraph(f"<b>{final_balance_label}:</b>", cell_right_bold),
            Paragraph(f"<b>Rs.{meta_info['net_payable']:,.2f}</b>", cell_right_bold)
        ])

    summary_table = Table(summary_rows, colWidths=[70 * mm, 34 * mm, 32 * mm])
    summary_table.setStyle(TableStyle([
        ('SPAN', (0, 0), (0, len(summary_rows) - 1)),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOX', (1, 0), (-1, -1), 0.5, colors.HexColor('#444')),
        ('INNERGRID', (1, 0), (-1, -1), 0.4, colors.HexColor('#777')),
        ('TOPPADDING', (0, 0), (-1, -1), 0.8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0.8),
    ]))
    story.append(summary_table)

    doc.build(story)

def render_mobile_print_button(pdf_base64, doc_no):
    btn_style = (
        "background:#222;color:#fff;padding:12px;font-size:16px;"
        "font-weight:bold;border:none;border-radius:8px;width:100%;cursor:pointer;"
    )
    js_code = (
        f"var b='{pdf_base64}';var c=atob(b);var n=new Array(c.length);"
        "for(var i=0;i<c.length;i++){n[i]=c.charCodeAt(i);}"
        "var blob=new Blob([new Uint8Array(n)],{type:'application/pdf'});"
        "var u=URL.createObjectURL(blob);var w=window.open(u);"
        "if(w){w.focus();w.print();}else{window.location.href=u;}"
    )
    html = (
        f'<div style="margin:4px 0;">'
        f'<button style="{btn_style}" onclick="pDoc()">'
        f'🖨️ Quick Print ({doc_no}) to Epson L3350</button>'
        f'<script>function pDoc(){{{js_code}}}</script></div>'
    )
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
        d_col1, d_col2 = st.columns([2, 1.3])
        with d_col1:
            if st.session_state.lock_bill_date:
                default_dt = st.session_state.locked_date_val
            else:
                default_dt = datetime.today().date()
            doc_date = st.date_input("Date", value=default_dt, key=f"cdate_{r_key}")
            if st.session_state.lock_bill_date:
                st.session_state.locked_date_val = doc_date
        with d_col2:
            st.write("")
            st.write("")
            is_locked = st.checkbox(
                "🔒 Lock Date",
                value=st.session_state.lock_bill_date,
                help="Back-date entries karte samay date ko lock rakhein"
            )
            if is_locked != st.session_state.lock_bill_date:
                st.session_state.lock_bill_date = is_locked
                st.session_state.locked_date_val = doc_date
                st.rerun()

        cust_mob = st.text_input(
            "Mobile No",
            value=st.session_state.c_mob,
            key=f"cmob_{r_key}",
            placeholder="10 Digit Number"
        )

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

        payment_mode_selected = st.selectbox(
            "Payment Mode",
            ["UPI / QR (PhonePe/GPay)", "Cash", "Bank Transfer / Card", "Multiple Modes"],
            key=f"pmode_{r_key}"
        )
        if mode == "Tax Invoice (GST)":
            cust_gstin_in = st.text_input(
                "Party GSTIN (Optional)",
                value=st.session_state.c_gstin,
                key=f"cgst_{r_key}",
                placeholder="OPTIONAL"
            )
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
                
        cust_name_in = st.text_input(
            "Customer Name",
            value=st.session_state.c_name,
            key=f"cname_{r_key}",
            placeholder="e.g. SOMNATH DAS"
        )
        cust_pan_in = st.text_input(
            "Customer PAN (Name ke sath print hoga)",
            value=st.session_state.c_pan,
            key=f"cpan_{r_key}",
            placeholder="e.g. ABCDE1234F"
        )
        cust_address_in = st.text_input(
            "Address",
            value=st.session_state.c_addr,
            key=f"caddr_{r_key}"
        )
        cust_aadhaar_in = st.text_input(
            "Aadhaar No (Address ke sath print hoga)",
            value=st.session_state.c_aadhaar,
            key=f"caadhaar_{r_key}",
            placeholder="e.g. 12 Digit Aadhaar"
        )

    cust_name = cust_name_in.strip().upper()
    cust_pan = cust_pan_in.strip().upper()
    cust_address = cust_address_in.strip().upper()
    cust_aadhaar = cust_aadhaar_in.strip().upper()
    cust_gstin = cust_gstin_in.strip().upper()

    st.markdown("---")
    st.subheader("🛒 Item Details")

    selected_preset = st.selectbox("⚡ Quick Select Item (Optional)", COMMON_ITEMS)

    purity_options = ["22K (916)", "18K (750)", "SILVER", "24K (999)"]

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
        submitted = st.form_submit_button(
            f"➕ Add Item #{next_item_no} ({purity_choice})",
            use_container_width=True
        )
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
                    st.caption(
                        f"Net: {itm['net_wt']}g | Rate: ₹{itm['rate']:,.0f} | "
                        f"Making: {itm['making_display']} | Total: ₹{itm['total']:,.2f}"
                    )
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

            if (
                st.session_state.editing_item_idx is not None
                and st.session_state.editing_item_idx < len(st.session_state.items_list)
            ):
                e_idx = st.session_state.editing_item_idx
                e_itm = st.session_state.items_list[e_idx]
                st.markdown("---")
                st.markdown(f"##### 🔧 Editing Item #{e_idx + 1}: {e_itm['desc']}")
                with st.form(f"edit_item_form_{e_idx}"):
                    f1, f2, f3, f4 = st.columns([2, 1.2, 1.2, 1.2])
                    with f1:
                        new_desc = st.text_input("Description", value=e_itm["desc"])
                    with f2:
                        h_val = "" if e_itm.get("huid") == "NA" else e_itm.get("huid", "")
                        new_huid = st.text_input("HUID", value=h_val)
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
                item_labels = []
                for idx_lbl, itm_lbl in enumerate(st.session_state.items_list):
                    lbl_str = "#" + str(idx_lbl + 1) + " " + str(itm_lbl["desc"]) + " (" + str(itm_lbl["net_wt"]) + "g)"
                    item_labels.append(lbl_str)
                target_item_idx = st.selectbox(
                    "Select Item to Adjust",
                    range(len(item_labels)),
                    format_func=lambda i: item_labels[i]
                )
            with rc_col2:
                customer_giving_amt = st.number_input("Customer Final Amount (₹)", min_value=0.0, step=100.0, format="%.2f", value=0.0)
            with rc_col3:
                st.write("")
                st.write("")
                if st.button("⚡ Apply Auto Making", use_container_width=True):
                    if customer_giving_amt > 0:
                        other_items_total = sum(
                            x["total"] for i, x in enumerate(st.session_state.items_list) if i != target_item_idx
                        )
                        if mode == "Tax Invoice (GST)":
                            target_subtotal = customer_giving_amt / 1.03
                        else:
                            target_subtotal = customer_giving_amt
                        
                        target_item_total = target_subtotal - other_items_total
                        t_itm = st.session_state.items_list[target_item_idx]
                        metal_val = (t_itm["net_wt"] * t_itm["rate"]) / 10.0
                        calc_making_rs = round(target_item_total - metal_val, 2)

                        if calc_making_rs >= 0:
                            st.session_state.items_list[target_item_idx]["making_val"] = calc_making_rs
                            st.session_state.items_list[target_item_idx]["making_unit"] = "₹"
                            st.session_state.items_list[target_item_idx]["making_display"] = f"Rs. {calc_making_rs:,.0f}"
                            st.session_state.items_list[target_item_idx]["total"] = round(metal_val + calc_making_rs, 2)
                            st.success(f"✅ Making Charge ₹{calc_making_rs:,.2f} सेट हो गया!")
                            st.rerun()
                        else:
                            st.error(f"⚠️ कस्टमर का अमाउंट धातु की मूल कीमत (₹{metal_val:,.2f}) से भी कम है!")

        if st.button("🗑️ Clear All Items"):
            st.session_state.items_list = []
            st.session_state.editing_item_idx = None
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

        net_difference = gross_total - old_val_input
        
        if net_difference < 0:
            is_refund = True
            abs_diff = abs(net_difference)
            final_amt = int(math.floor(abs_diff / 10.0) * 10)
            round_off = round(final_amt - abs_diff, 2)
            balance_label = "Refund / Return to Customer"
            display_box_color = "🟢"
        else:
            is_refund = False
            final_amt = int(math.floor(net_difference / 10.0) * 10)
            round_off = round(final_amt - net_difference, 2)
            balance_label = "Net Payable"
            display_box_color = "💰"

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
            st.subheader(f"{display_box_color} {balance_label}: Rs. {final_amt:,.2f}")

        if final_amt >= 200000 and not cust_pan:
            st.warning("⚠️ **Legal Alert:** Bill amount ₹2,00,000 se zyada hai. Income Tax kanoon ke anusar Customer PAN Card likhna anivarya hai.")

        btn_text = f"💾 Generate & Save {mode}"
        if st.button(btn_text, type="primary", use_container_width=True):
            if not cust_name:
                st.error("Customer Name zaroori hai!")
            else:
                if st.session_state.lock_bill_date:
                    st.session_state.locked_date_val = doc_date

                formatted_date = doc_date.strftime("%d %b %Y")
                gstr_date = doc_date.strftime("%d-%m-%Y")
                safe_file_prefix = str(doc_number).replace('/', '_')
                pdf_filename = f"{safe_file_prefix}.pdf"
                pdf_filepath = os.path.join(PDF_DIR, pdf_filename)

                meta_info = {
                    "doc_no": doc_number,
                    "date": formatted_date,
                    "customer_name": cust_name,
                    "customer_mob": cust_mob,
                    "customer_address": cust_address,
                    "customer_aadhaar": cust_aadhaar,
                    "customer_pan": cust_pan,
                    "customer_gstin": cust_gstin,
                    "payment_mode": payment_mode_selected,
                    "subtotal": subtotal,
                    "cgst": cgst,
                    "sgst": sgst,
                    "gross": gross_total,
                    "old_exchange": old_val_input,
                    "exchange_label": exchange_display_label,
                    "round_off": round_off,
                    "net_payable": final_amt,
                    "final_balance_label": balance_label
                }

                try:
                    generate_a5_pdf(mode, meta_info, st.session_state.items_list, pdf_filepath)
                except Exception as pdf_err:
                    st.error(f"PDF Banane me dikkat aayi: {str(pdf_err)}")
                    st.stop()

                with open(pdf_filepath, "rb") as f:
                    pdf_base64 = base64.b64encode(f.read()).decode('utf-8')

                total_net_wt = sum(i["net_wt"] for i in st.session_state.items_list)
                items_summary_str = ", ".join([
                    f"{it['desc']} (HUID:{it.get('huid','NA')}, {it['net_wt']}g)"
                    for it in st.session_state.items_list
                ])

                gstr1_data = None
                if mode == "Tax Invoice (GST)":
                    if cust_gstin:
                        invoice_category = "B2B (Registered)"
                    elif gross_total > 250000:
                        invoice_category = "B2CL (>2.5L)"
                    else:
                        invoice_category = "B2CS (Retail)"

                    gstr1_data = {
                        "Month-Year": doc_date.strftime("%b-%Y"),
                        "Invoice Date": gstr_date,
                        "Invoice No": doc_number,
                        "Invoice Category": invoice_category,
                        "Party Name": cust_name,
                        "Party GSTIN": cust_gstin if cust_gstin else "UNREGISTERED",
                        "Party PAN": cust_pan if cust_pan else "NA",
                        "Place of Supply": "20-JHARKHAND",
                        "HSN Code": "7113",
                        "Total Net Wt (g)": round(total_net_wt, 3),
                        "Taxable Value (₹)": round(subtotal, 2),
                        "GST Rate (%)": 3.0,
                        "CGST (1.5%) (₹)": round(cgst, 2),
                        "SGST (1.5%) (₹)": round(sgst, 2),
                        "Total Invoice Value (₹)": round(gross_total, 2),
                        "Payment Mode": payment_mode_selected
                    }

                addr_full = f"{cust_address} (Aadhaar: {cust_aadhaar})" if cust_aadhaar else cust_address
                name_full = f"{cust_name} (PAN: {cust_pan})" if cust_pan else cust_name

                payload = {
                    "doc_no": doc_number,
                    "date": formatted_date,
                    "type": mode,
                    "customer_name": cust_name,
                    "customer_mob": str(cust_mob).strip() if cust_mob else "NA",
                    "customer_address": addr_full,
                    "customer_pan": cust_pan or "NA",
                    "payment_mode": payment_mode_selected,
                    "gross": gross_total,
                    "old_exchange": old_val_input,
                    "net_payable": -final_amt if is_refund else final_amt,
                    "cgst": cgst,
                    "sgst": sgst,
                    "pdf_base64": pdf_base64
                }

                db_row = {
                    "Document No": doc_number,
                    "Date": formatted_date,
                    "Customer Name": name_full,
                    "Mobile No": str(cust_mob).strip() if cust_mob else "NA",
                    "Address": addr_full,
                    "Party GSTIN": cust_gstin if cust_gstin else "NA",
                    "Payment Mode": payment_mode_selected,
                    "Type": mode,
                    "Items": items_summary_str,
                    "Gross Amount": gross_total,
                    "Old Exchange": old_val_input,
                    "Net Payable": f"Refund: Rs. {final_amt:,.2f}" if is_refund else final_amt,
                    "SGST": sgst,
                    "CGST": cgst
                }
                save_to_database(db_row, mode, gstr1_row=gstr1_data)

                with st.spinner("☁️ Google Sheet & Drive par upload ho raha hai..."):
                    sync_ok, drive_pdf_url, error_msg = sync_to_google_sheet(payload)

                st.success(f"✅ {mode} ({doc_number}) ban gaya!")
                if sync_ok and drive_pdf_url:
                    st.markdown(f"🔗 [Google Drive par PDF dekhein]({drive_pdf_url})")

                render_mobile_print_button(pdf_base64, doc_number)

                with open(pdf_filepath, "rb") as f:
                    st.download_button(
                        label=f"📄 Download A5 PDF File ({doc_number})",
                        data=f,
                        file_name=pdf_filename,
                        mime="application/pdf",
                        use_container_width=True
                    )

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

with tab_gst_report:
    st.subheader("📊 GSTR-1 Ready Report (CA & GST Portal Filing)")
    st.write("Yahan sirf **Tax Invoice (GST)** ka clean data HSN `7113`, Taxable Value, CGST (1.5%), SGST (1.5%) aur B2B/B2C breakup ke sath taiyar rehta hai:")

    if os.path.exists(GST_FILING_FILE):
        try:
            gdf = pd.read_excel(GST_FILING_FILE)
            if not gdf.empty:
                months_avail = ["All Months"] + sorted(gdf["Month-Year"].dropna().unique().tolist(), reverse=True)
                sel_month = st.selectbox("📅 Select Filing Month (Mahina chunein)", months_avail)

                if sel_month != "All Months":
                    filtered_gst = gdf[gdf["Month-Year"] == sel_month]
                else:
                    filtered_gst = gdf

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Total GST Bills", len(filtered_gst))
                m2.metric("Total Taxable Value", f"₹{filtered_gst['Taxable Value (₹)'].sum():,.2f}")
                m3.metric("Total CGST (1.5%)", f"₹{filtered_gst['CGST (1.5%) (₹)'].sum():,.2f}")
                m4.metric("Total SGST (1.5%)", f"₹{filtered_gst['SGST (1.5%) (₹)'].sum():,.2f}")

                st.dataframe(filtered_gst, use_container_width=True)

                export_path = f"GSTR1_Return_{sel_month.replace(' ', '_')}.xlsx"
                filtered_gst.to_excel(export_path, index=False)
                with open(export_path, "rb") as ef:
                    st.download_button(
                        label=f"📥 Download GSTR-1 Filing Excel ({sel_month})",
                        data=ef,
                        file_name=export_path,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )
            else:
                st.info("Abhi tak koi GST Tax Invoice generate nahi hua hai.")
        except Exception as e:
            st.error(f"GST file load karne mein error: {e}")
    else:
        st.info("Jaise hi aap pehla **Tax Invoice (GST)** banayenge, aapki GSTR-1 filing sheet yahan apne aap ban jayegi.")

with tab_search:
    st.subheader("🔍 Customer Bill & History Search")
    st.write("Customer ka **Naam** ya **Mobile No** search karein. Naya bill banane ke liye **'⚡ Auto-Fill'** par click karein ya purane bill ko **'🗑️ Del'** karein:")

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
        search_query = st.text_input("🔎 Search by Name or Mobile No", placeholder="e.g. RAJU YA 77177...").strip()

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
                    c_name_val = str(row.get("Customer Name", ""))
                    c_mob_val = str(row.get("Mobile No", ""))
                    c_addr_val = str(row.get("Address", "MANGO, JAMSHEDPUR"))
                    c_gst_val = str(row.get("Party GSTIN", ""))
                    if c_gst_val == "NA":
                        c_gst_val = ""
                    safe_p_name = d_no.replace('/', '_')
                    p_file = os.path.join(PDF_DIR, f"{safe_p_name}.pdf")

                    card = st.container()
                    with card:
                        r_col1, r_col2, r_col3, r_col4, r_col5 = st.columns([1.8, 2.4, 1.7, 1.0, 1.0])
                        with r_col1:
                            st.markdown(f"**{d_no}** ({row.get('Type', '')})")
                            st.caption(f"📅 {row.get('Date', '')}")
                        with r_col2:
                            st.markdown(f"👤 **{c_name_val}** | 📞 {c_mob_val}")
                            st.caption(f"📦 Items: {row.get('Items', 'NA')}")
                        with r_col3:
                            st.markdown(f"💰 **₹{row.get('Net Payable', 0)}**")
                            if st.button("⚡ Auto-Fill", key=f"btn_s_fill_{idx}_{safe_p_name}"):
                                found_via_btn = lookup_customer_by_mobile(c_mob_val)
                                if found_via_btn:
                                    st.session_state.c_name = found_via_btn["name"]
                                    st.session_state.c_addr = found_via_btn["address"]
                                    st.session_state.c_pan = found_via_btn["pan"]
                                    st.session_state.c_aadhaar = found_via_btn["aadhaar"]
                                    st.session_state.c_gstin = found_via_btn["gstin"]
                                else:
                                    st.session_state.c_name = c_name_val.upper()
                                    st.session_state.c_addr = c_addr_val.upper()
                                    st.session_state.c_gstin = c_gst_val.upper()
                                st.session_state.c_mob = "" if c_mob_val == "NA" else c_mob_val
                                st.session_state.form_reset_count += 1
                                st.success(f"✅ {c_name_val} ki details load ho gayi hain! Upar 'New Bill / Estimate' tab par click karein.")
                        with r_col4:
                            if os.path.exists(p_file):
                                with open(p_file, "rb") as pf:
                                    st.download_button(
                                        label="⬇️ PDF",
                                        data=pf,
                                        file_name=f"{safe_p_name}.pdf",
                                        mime="application/pdf",
                                        key=f"btn_s_dl_{idx}_{safe_p_name}"
                                    )
                        with r_col5:
                            if st.button("🗑️ Del", key=f"btn_s_del_{idx}_{safe_p_name}"):
                                if delete_bill_record(d_no):
                                    st.success(f"🗑️ {d_no} delete ho gaya!")
                                    st.rerun()
                        st.divider()
            else:
                st.warning(f"'{search_query}' ke naam ya number se koi record nahi mila.")
        else:
            st.info("💡 Upar search box mein customer ka naam ya mobile number likhein.")
            st.markdown("##### 🕒 Recent 5 Transactions:")
            recent_df = all_data.tail(5).iloc[::-1]
            for idx, row in recent_df.iterrows():
                d_no = str(row.get("Document No", ""))
                c_name_val = str(row.get("Customer Name", ""))
                c_mob_val = str(row.get("Mobile No", ""))
                c_addr_val = str(row.get("Address", "MANGO, JAMSHEDPUR"))
                c_gst_val = str(row.get("Party GSTIN", ""))
                if c_gst_val == "NA":
                    c_gst_val = ""
                safe_r_name = d_no.replace('/', '_')
                
                rc1, rc2, rc3, rc4 = st.columns([2.2, 2.3, 1.5, 1.0])
                with rc1:
                    st.markdown(f"**{d_no}** - {c_name_val}")
                with rc2:
                    st.markdown(f"📞 {c_mob_val} | 💰 ₹{row.get('Net Payable', 0)}")
                with rc3:
                    if st.button("⚡ Use Details", key=f"btn_r_fill_{idx}_{safe_r_name}"):
                        found_via_btn = lookup_customer_by_mobile(c_mob_val)
                        if found_via_btn:
                            st.session_state.c_name = found_via_btn["name"]
                            st.session_state.c_addr = found_via_btn["address"]
                            st.session_state.c_pan = found_via_btn["pan"]
                            st.session_state.c_aadhaar = found_via_btn["aadhaar"]
                            st.session_state.c_gstin = found_via_btn["gstin"]
                        else:
                            st.session_state.c_name = c_name_val.upper()
                            st.session_state.c_addr = c_addr_val.upper()
                            st.session_state.c_gstin = c_gst_val.upper()
                        st.session_state.c_mob = "" if c_mob_val == "NA" else c_mob_val
                        st.session_state.form_reset_count += 1
                        st.success(f"✅ {c_name_val} ki details load ho gayi hain! Upar 'New Bill / Estimate' tab par click karein.")
                with rc4:
                    if st.button("🗑️ Del", key=f"btn_r_del_{idx}_{safe_r_name}"):
                        if delete_bill_record(d_no):
                            st.success(f"🗑️ {d_no} delete ho gaya!")
                            st.rerun()
                st.divider()
    else:
        st.info("Abhi tak koi transaction record nahi hua hai.")
