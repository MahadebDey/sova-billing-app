import streamlit as st
import pandas as pd
from datetime import datetime
import os
from reportlab.lib.pagesizes import A5
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm

# Page Configuration
st.set_page_config(page_title="Billing App", layout="wide", page_icon="💎")

# Database & PDF Folder Setup
DB_FILE = "Sales_Database.xlsx"
PDF_DIR = "Invoices_PDF"
os.makedirs(PDF_DIR, exist_ok=True)

def get_next_invoice_number():
    if os.path.exists(DB_FILE):
        try:
            df = pd.read_excel(DB_FILE)
            if not df.empty and "Invoice No" in df.columns:
                last_inv = str(df["Invoice No"].iloc[-1])
                if "-" in last_inv:
                    parts = last_inv.split("-")
                    num = int(parts[-1]) + 1
                    return f"SV-{datetime.now().year}-{str(num).zfill(3)}"
        except Exception:
            pass
    return f"SV-{datetime.now().year}-001"

def save_to_database(row_dict):
    if os.path.exists(DB_FILE):
        try:
            df = pd.read_excel(DB_FILE)
            df = pd.concat([df, pd.DataFrame([row_dict])], ignore_index=True)
        except Exception:
            df = pd.DataFrame([row_dict])
    else:
        df = pd.DataFrame([row_dict])
    df.to_excel(DB_FILE, index=False)

def generate_a5_pdf(inv_data, items, pdf_path):
    # Exact A5 size: 148mm x 210mm
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A5,
        leftMargin=7*mm,
        rightMargin=7*mm,
        topMargin=8*mm,
        bottomMargin=7*mm
    )
    story = []
    styles = getSampleStyleSheet()
    
    inv_title_style = ParagraphStyle(
        'InvTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=12,
        alignment=1
    )
    cell_style = ParagraphStyle(
        'CellText',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=7,
        leading=8.5
    )
    cell_bold = ParagraphStyle(
        'CellBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=7,
        leading=8.5
    )
    cell_header = ParagraphStyle(
        'CellHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=7,
        leading=8.5,
        textColor=colors.whitesmoke,
        alignment=1
    )

    # 🟢 TAX INVOICE Title Header
    story.append(Paragraph("<u>TAX INVOICE</u>", inv_title_style))
    story.append(Spacer(1, 2.5*mm))

    # Meta Info Table
    meta_data = [
        [
            Paragraph(f"<b>Date:</b> {inv_data['date']}", cell_style),
            Paragraph("<b>GSTIN:</b> 20AUJPD1127G1ZE", cell_style)
        ],
        [
            Paragraph(f"<b>Invoice No:</b> {inv_data['invoice_no']}", cell_style),
            Paragraph("<b>Mobile No:</b> 7717740697", cell_style)
        ],
        [
            Paragraph(f"<b>Customer Name:</b> {inv_data['customer_name']}", cell_style),
            Paragraph("<b>GST Type:</b> CGST + SGST (1.5% Each)", cell_style)
        ],
        [
            Paragraph(f"<b>Customer Mob:</b> {inv_data['customer_mob']}", cell_style),
            Paragraph("<b>State & Code:</b> 20 - JHARKHAND", cell_style)
        ],
        [
            Paragraph(f"<b>Address:</b> {inv_data['customer_address']}", cell_style),
            Paragraph(f"<b>Party GSTIN:</b> {inv_data['customer_gstin'] or 'NA'}", cell_style)
        ]
    ]
    meta_table = Table(meta_data, colWidths=[67*mm, 67*mm])
    meta_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#0b2f64')),
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f7f9fc')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 1.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.5),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 2*mm))

    # Items Header
    item_rows = [[
        Paragraph("Sr", cell_header),
        Paragraph("Particular / Description", cell_header),
        Paragraph("Gross(g)", cell_header),
        Paragraph("Net(g)", cell_header),
        Paragraph("HSN", cell_header),
        Paragraph("Purity", cell_header),
        Paragraph("Rate/10g", cell_header),
        Paragraph("Making", cell_header),
        Paragraph("Total Amount", cell_header)
    ]]

    # Items Data
    for idx, itm in enumerate(items, 1):
        item_rows.append([
            Paragraph(str(idx), cell_style),
            Paragraph(str(itm['desc']), cell_style),
            Paragraph(f"{itm['gross_wt']:.2f}", cell_style),
            Paragraph(f"{itm['net_wt']:.2f}", cell_style),
            Paragraph(str(itm['hsn']), cell_style),
            Paragraph(str(itm['purity']), cell_style),
            Paragraph(f"Rs. {itm['rate']:,.2f}", cell_style),
            Paragraph(f"{itm['making_pct']}%", cell_style),
            Paragraph(f"Rs. {itm['total']:,.2f}", cell_style)
        ])

    for _ in range(max(0, 4 - len(items))):
        item_rows.append(["", "", "", "", "", "", "", "", ""])

    item_table = Table(item_rows, colWidths=[6*mm, 33*mm, 12*mm, 12*mm, 10*mm, 13*mm, 19*mm, 11*mm, 18*mm])
    item_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0b2f64')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#b0bec5')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(item_table)
    story.append(Spacer(1, 2*mm))

    # Summary & Terms Table
    terms_text = Paragraph(
        "<b>Terms & Conditions:</b><br/>"
        "1. We are not responsible for any breakage/damage.<br/>"
        "2. All disputes subject to Jamshedpur jurisdiction.",
        cell_style
    )
    
    summary_rows = [
        [terms_text, Paragraph("<b>Subtotal:</b>", cell_style), Paragraph(f"Rs. {inv_data['subtotal']:,.2f}", cell_style)],
        ["", Paragraph("<b>CGST (1.5%):</b>", cell_style), Paragraph(f"Rs. {inv_data['cgst']:,.2f}", cell_style)],
        ["", Paragraph("<b>SGST (1.5%):</b>", cell_style), Paragraph(f"Rs. {inv_data['sgst']:,.2f}", cell_style)],
        ["", Paragraph("<b>Gross Total:</b>", cell_style), Paragraph(f"Rs. {inv_data['gross']:,.2f}", cell_style)],
        ["", Paragraph("<b>Round Off:</b>", cell_style), Paragraph(f"Rs. {inv_data['round_off']:+.2f}", cell_style)],
        ["", Paragraph("<b>Net Payable:</b>", cell_bold), Paragraph(f"<b>Rs. {inv_data['net_payable']:,.2f}</b>", cell_bold)],
    ]

    summary_table = Table(summary_rows, colWidths=[68*mm, 32*mm, 34*mm])
    summary_table.setStyle(TableStyle([
        ('SPAN', (0,0), (0,5)),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('BOX', (1,0), (-1,-1), 0.5, colors.HexColor('#0b2f64')),
        ('INNERGRID', (1,0), (-1,-1), 0.5, colors.HexColor('#cfd8dc')),
        ('BACKGROUND', (1,5), (2,5), colors.HexColor('#e8f5e9')),
        ('TOPPADDING', (0,0), (-1,-1), 1.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.5),
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 3*mm))

    # Signatory Section
    sig_data = [
        [Paragraph("", cell_style), Paragraph("<b>Authorised Signatory</b><br/><br/><br/>", ParagraphStyle('Sig', parent=styles['Normal'], alignment=1, fontSize=7, leading=9))]
    ]
    sig_table = Table(sig_data, colWidths=[80*mm, 54*mm])
    story.append(sig_table)

    doc.build(story)

# --- STREAMLIT UI ---
st.title("💎 TAX INVOICE - Billing System")

col1, col2 = st.columns([1.5, 1.5])
with col1:
    invoice_no = st.text_input("Invoice No", value=get_next_invoice_number())
    cust_name = st.text_input("Customer Name", placeholder="e.g. Somnath Das")
    cust_address = st.text_input("Address", value="Mango, Jamshedpur")
with col2:
    inv_date = st.date_input("Invoice Date", value=datetime.today())
    cust_mob = st.text_input("Mobile No", placeholder="10 Digit Number")
    cust_gstin = st.text_input("Party GSTIN (Optional)", placeholder="Optional")

st.markdown("---")
st.subheader("🛒 Item Details")

if "items_list" not in st.session_state:
    st.session_state.items_list = []

with st.container():
    ic1, ic2, ic3, ic4, ic5, ic6 = st.columns([2.5, 1.2, 1.2, 1.5, 1.2, 1])
    with ic1:
        item_desc = st.text_input("Item Description", placeholder="e.g. SILVER CHAIN", key="desc_in")
    with ic2:
        gross_wt = st.number_input("Gross Wt (g)", min_value=0.0, step=0.01, format="%.2f", key="gwt_in")
    with ic3:
        net_wt = st.number_input("Net Wt (g)", min_value=0.0, step=0.01, format="%.2f", key="nwt_in")
    with ic4:
        rate_10g = st.number_input("Rate (Per 10g)", min_value=0.0, step=10.0, format="%.2f", key="rate_in")
    with ic5:
        making_pct = st.number_input("Making %", min_value=0.0, step=1.0, value=12.0, key="making_in")
    with ic6:
        purity = st.selectbox("Purity", ["SILVER", "22K (916)", "18K (750)", "24K (999)"], key="purity_in")

    if st.button("➕ Add Item"):
        if item_desc and net_wt > 0 and rate_10g > 0:
            metal_cost = (net_wt * rate_10g) / 10.0
            total_item_amt = metal_cost * (1 + (making_pct / 100.0))
            st.session_state.items_list.append({
                "desc": item_desc.upper(),
                "gross_wt": gross_wt or net_wt,
                "net_wt": net_wt,
                "hsn": "7113",
                "purity": purity,
                "rate": rate_10g,
                "making_pct": making_pct,
                "total": round(total_item_amt, 2)
            })
            st.rerun()
        else:
            st.warning("Kripya Item Description, Net Weight aur Rate sahi se bharein.")

if st.session_state.items_list:
    df_items = pd.DataFrame(st.session_state.items_list)
    st.table(df_items[["desc", "gross_wt", "net_wt", "hsn", "purity", "rate", "making_pct", "total"]])

    if st.button("🗑️ Clear All Items"):
        st.session_state.items_list = []
        st.rerun()

    subtotal = sum(i["total"] for i in st.session_state.items_list)
    cgst = round(subtotal * 0.015, 2)
    sgst = round(subtotal * 0.015, 2)
    gross_total = subtotal + cgst + sgst
    net_payable = round(gross_total)
    round_off = round(net_payable - gross_total, 2)

    st.markdown("---")
    sc1, sc2 = st.columns([2, 1])
    with sc2:
        st.markdown(f"**Total Taxable:** Rs. {subtotal:,.2f}")
        st.markdown(f"**CGST (1.5%):** Rs. {cgst:,.2f}")
        st.markdown(f"**SGST (1.5%):** Rs. {sgst:,.2f}")
        st.markdown(f"**Gross Total:** Rs. {gross_total:,.2f}")
        st.markdown(f"**Round Off:** Rs. {round_off:+.2f}")
        st.subheader(f"💰 Net Payable: Rs. {net_payable:,.2f}")

    if st.button("💾 Generate A5 Invoice & Save to Excel", type="primary", use_container_width=True):
        if not cust_name:
            st.error("Customer Name zaroori hai!")
        else:
            formatted_date = inv_date.strftime("%d %b %Y")
            pdf_filename = f"{invoice_no}.pdf"
            pdf_filepath = os.path.join(PDF_DIR, pdf_filename)

            inv_data = {
                "invoice_no": invoice_no,
                "date": formatted_date,
                "customer_name": cust_name.upper(),
                "customer_mob": cust_mob,
                "customer_address": cust_address,
                "customer_gstin": cust_gstin,
                "subtotal": subtotal,
                "cgst": cgst,
                "sgst": sgst,
                "gross": gross_total,
                "round_off": round_off,
                "net_payable": net_payable
            }

            generate_a5_pdf(inv_data, st.session_state.items_list, pdf_filepath)

            db_row = {
                "Invoice No": invoice_no,
                "Date": formatted_date,
                "Customer Name": cust_name.upper(),
                "Mobile No": cust_mob,
                "State Code": "20-JHARKHAND",
                "Total Amount": net_payable,
                "SGST": sgst,
                "CGST": cgst
            }
            save_to_database(db_row)

            st.success(f"✅ Invoice {invoice_no} ban gaya hai aur Database mein save ho chuka hai!")

            with open(pdf_filepath, "rb") as f:
                st.download_button(
                    label="📄 Download & Print A5 PDF",
                    data=f,
                    file_name=pdf_filename,
                    mime="application/pdf",
                    use_container_width=True
                )
