import os
import re
import tempfile
from datetime import datetime
from io import BytesIO
import base64
import zipfile
from docx import Document
import streamlit as st
import pandas as pd
from pypdf import PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors

# Configuration
TEMPLATE_PATH = "invoice_template.docx"
OUTPUT_FOLDER = "generated_invoices"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def sanitize_filename(name):
    return re.sub(r'[\\/:*?"<>|]', '_', name)

def generate_pdf_from_docx(docx_path, pdf_path):
    """Generate PDF using pypdf and reportlab with professional formatting"""
    try:
        # Read DOCX content
        doc = Document(docx_path)
        
        # Create PDF buffer
        buffer = BytesIO()
        
        # Setup PDF document with margins
        doc_template = SimpleDocTemplate(
            buffer,
            pagesize=letter,
            leftMargin=40,
            rightMargin=40,
            topMargin=40,
            bottomMargin=40
        )
        
        # Set up styles
        styles = getSampleStyleSheet()
        elements = []
        
        # Process document content
        for para in doc.paragraphs:
            if para.text.strip():  # Skip empty paragraphs
                p = Paragraph(para.text, styles["Normal"])
                elements.append(p)
        
        # Process tables with improved formatting
        for table in doc.tables:
            table_data = []
            for row in table.rows:
                row_data = []
                for cell in row.cells:
                    row_data.append(cell.text)
                table_data.append(row_data)
            
            # Create table with styling
            tbl = Table(table_data, colWidths='*')
            tbl.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#4F81BD')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,0), 10),
                ('BOTTOMPADDING', (0,0), (-1,0), 12),
                ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#DCE6F1')),
                ('GRID', (0,0), (-1,-1), 1, colors.black),
                ('VALIGN', (0,0), (-1,-1), 'TOP')
            ]))
            elements.append(tbl)
        
        # Build PDF
        doc_template.build(elements)
        
        # Save to file
        with open(pdf_path, "wb") as f:
            f.write(buffer.getvalue())
        
        return True
        
    except Exception as e:
        st.error(f"PDF generation error: {str(e)}")
        return False

def generate_invoice(row_data, invoice_number):
    """Generate invoice with all data and proper error handling"""
    try:
        # Create temporary files
        temp_docx = os.path.join(tempfile.gettempdir(), f"temp_{invoice_number}.docx")
        customer_name = sanitize_filename(row_data.get('MARK', 'Customer'))
        pdf_name = f"Invoice_{invoice_number}_{customer_name}.pdf"
        pdf_path = os.path.join(OUTPUT_FOLDER, pdf_name)
        
        # Load template
        doc = Document(TEMPLATE_PATH)
        
        # Set current date and invoice number
        current_date = datetime.now().strftime("%Y-%m-%d")
        row_data.update({
            "DATE": current_date,
            "INVOICE NUMBER": str(invoice_number)
        })
        
        # Replace placeholders in paragraphs
        for paragraph in doc.paragraphs:
            for key, value in row_data.items():
                if isinstance(value, (int, float)):
                    value = f"{value:.2f}" if pd.notna(value) else ""
                for placeholder in [f"{{{{{key}}}}}", f"{{{{{key}.}}}}"]:
                    if placeholder in paragraph.text:
                        paragraph.text = paragraph.text.replace(placeholder, str(value))
        
        # Replace placeholders in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for key, value in row_data.items():
                        if isinstance(value, (int, float)):
                            value = f"{value:.2f}" if pd.notna(value) else ""
                        for placeholder in [f"{{{{{key}}}}}", f"{{{{{key}.}}}}"]:
                            if placeholder in cell.text:
                                cell.text = cell.text.replace(placeholder, str(value))
        
        # Save temporary DOCX
        doc.save(temp_docx)
        
        # Convert to PDF
        if generate_pdf_from_docx(temp_docx, pdf_path):
            os.remove(temp_docx)
            return pdf_path
        return None
        
    except Exception as e:
        st.error(f"Invoice generation failed: {str(e)}")
        return None

def update_notification_sheet(output_folder, pdf_name, customer_name, invoice_number, contact_number, invoice_total):
    """Update the notification spreadsheet"""
    sheet_path = os.path.join(output_folder, "Customer_Notification_Sheet.xlsx")
    
    new_entry = pd.DataFrame([{
        "CUSTOMER": customer_name,
        "INVOICE NO": invoice_number,
        "CONTACT NO": contact_number,
        "INVOICE TOTAL": invoice_total,
        "FILE PATH": pdf_name
    }])
    
    if os.path.exists(sheet_path):
        try:
            existing_data = pd.read_excel(sheet_path)
            updated_data = pd.concat([existing_data, new_entry], ignore_index=True)
        except:
            updated_data = new_entry
    else:
        updated_data = new_entry
    
    updated_data.to_excel(sheet_path, index=False)

def consolidate_rows(df):
    """Consolidate multiple rows for the same customer"""
    consolidated_data = []
    for customer_name, group in df.groupby("MARK"):
        total_qty = group["QTY"].sum(skipna=True)
        total_cbm = group["CBM"].sum(skipna=True)
        
        parking_charges = group["PARKING CHARGES"].dropna().iloc[0] if not group["PARKING CHARGES"].dropna().empty else 0
        
        if total_cbm < 0.05:
            calculated_charges = 10.00
        else:
            calculated_charges = (group["CBM"] * group["PER CHARGES"]).sum(skipna=True)
        
        total_charges = calculated_charges + parking_charges
        first_row = group.iloc[0]
        
        # Process multi-line values
        receipt_nos = []
        qtys = []
        descriptions = []
        cbms = []
        weights = []
        
        for _, row in group.iterrows():
            receipt_nos.append(str(row["RECEIPT NO."]) if pd.notna(row["RECEIPT NO."]) else "")
            qtys.append(f"{row['QTY']:.2f}" if pd.notna(row['QTY']) else "")
            descriptions.append(str(row["DESCRIPTION"]) if pd.notna(row["DESCRIPTION"]) else "")
            cbms.append(f"{row['CBM']:.2f}" if pd.notna(row['CBM']) else "")
            weights.append(f"{row['WEIGHT(KG)']:.2f}" if pd.notna(row['WEIGHT(KG)']) else "")
        
        consolidated_data.append({
            "RECEIPT NO.": "\n".join(receipt_nos),
            "QTY": "\n".join(qtys),
            "DESCRIPTION": "\n".join(descriptions),
            "CBM": "\n".join(cbms),
            "WEIGHT(KG)": "\n".join(weights),
            "PARKING CHARGES": f"{parking_charges:.2f}",
            "PER CHARGES": f"{first_row['PER CHARGES']:.2f}" if pd.notna(first_row['PER CHARGES']) else "",
            "TOTAL CHARGES": f"{total_charges:.2f}",
            "MARK": customer_name,
            "CONTACT NUMBER": str(first_row.get("CONTACT NUMBER", "")),
            "CARGO NUMBER": str(first_row.get("CARGO NUMBER", "")),
            "TRACKING NUMBER": str(first_row.get("TRACKING NUMBER", "")),
            "TERMS": str(first_row.get("TERMS", "")),
            "TOTAL QTY": f"{total_qty:.2f}",
            "TOTAL CBM": f"{total_cbm:.2f}",
            "TOTAL CHARGES_SUM": f"{total_charges:.2f}",
            "FLAT_RATE_APPLIED": "Yes" if total_cbm < 0.05 else "No"
        })
    return consolidated_data

def get_binary_file_downloader_html(file_path, file_label):
    """Generate a download link for files"""
    with open(file_path, 'rb') as f:
        data = f.read()
    bin_str = base64.b64encode(data).decode()
    href = f'<a href="data:application/octet-stream;base64,{bin_str}" download="{os.path.basename(file_path)}">{file_label}</a>'
    return href

# Streamlit UI
st.title("Professional Invoice Generator")

# Sidebar for quick actions
st.sidebar.header("Quick Actions")
st.sidebar.markdown("[Sample Excel Template](#)")

# File upload section
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        
        # Initialize missing columns
        for col in ["CARGO NUMBER", "TRACKING NUMBER", "TERMS", "PARKING CHARGES"]:
            if col not in df.columns:
                df[col] = "" if col != "PARKING CHARGES" else 0.0
        
        if "Weight Rate" not in df.columns:
            df["Weight Rate"] = 0.0
        if "PER CHARGES" not in df.columns:
            df["PER CHARGES"] = 0.0

        # Global Settings
        st.header("Global Settings")
        col1, col2, col3 = st.columns(3)
        with col1:
            global_per_charge = st.number_input("Default Rate", value=float(df["PER CHARGES"].iloc[0]) if not df["PER CHARGES"].empty else 0.0)
        with col2:
            global_weight_rate = st.number_input("Weight Rate", value=float(df["Weight Rate"].iloc[0]) if not df["Weight Rate"].empty else 0.0)
        with col3:
            global_parking = st.number_input("Parking Fee", value=0.0)
        
        if st.button("Apply Globally"):
            df["PER CHARGES"] = global_per_charge
            df["Weight Rate"] = global_weight_rate
            df["PARKING CHARGES"] = global_parking
            st.success("Settings applied to all customers!")

        # Customer-level Editing
        st.header("Customer Adjustments")
        customers = df["MARK"].unique()
        
        for customer in customers:
            with st.expander(f"Customer: {customer}"):
                cols = st.columns(3)
                with cols[0]:
                    new_rate = st.number_input(f"Rate for {customer}", 
                        value=float(df[df["MARK"] == customer]["PER CHARGES"].iloc[0]),
                        key=f"rate_{customer}")
                with cols[1]:
                    new_weight = st.number_input(f"Weight Rate for {customer}",
                        value=float(df[df["MARK"] == customer]["Weight Rate"].iloc[0]),
                        key=f"weight_{customer}")
                with cols[2]:
                    new_parking = st.number_input(f"Parking for {customer}",
                        value=float(df[df["MARK"] == customer]["PARKING CHARGES"].iloc[0]),
                        key=f"parking_{customer}")
                
                if st.button(f"Update {customer}"):
                    df.loc[df["MARK"] == customer, "PER CHARGES"] = new_rate
                    df.loc[df["MARK"] == customer, "Weight Rate"] = new_weight
                    df.loc[df["MARK"] == customer, "PARKING CHARGES"] = new_parking
                    st.success(f"Updated {customer}")

        # Calculations
        df["Weight CBM"] = df["WEIGHT(KG)"] / df["Weight Rate"]
        df["CBM"] = df[["MEAS.(CBM)", "Weight CBM"]].max(axis=1)
        df["Calculated Charges"] = df["CBM"] * df["PER CHARGES"]
        
        consolidated_data = consolidate_rows(df)
        consolidated_df = pd.DataFrame(consolidated_data)
        
        # Store in session state
        st.session_state.consolidated_df = consolidated_df
        
        st.header("Processed Data")
        st.dataframe(consolidated_df)

        # Invoice Generation
        st.header("Invoice Generation")
        last_invoice = st.number_input("Starting Invoice Number", min_value=1, value=1)
        
        # Individual Invoice Generation
        st.subheader("Generate Single Invoice")
        selected_customer = st.selectbox(
            "Select Customer", 
            options=consolidated_df["MARK"].unique()
        )
        single_inv_num = st.number_input("Invoice Number", min_value=1, value=last_invoice)
        
        if st.button("Generate Selected Invoice"):
            row = consolidated_df[consolidated_df["MARK"] == selected_customer].iloc[0]
            pdf_path = generate_invoice(row, single_inv_num)
            if pdf_path:
                update_notification_sheet(
                    OUTPUT_FOLDER,
                    os.path.basename(pdf_path),
                    row["MARK"],
                    single_inv_num,
                    row["CONTACT NUMBER"],
                    row["TOTAL CHARGES_SUM"]
                )
                with open(pdf_path, "rb") as f:
                    st.download_button(
                        label="Download Invoice",
                        data=f,
                        file_name=os.path.basename(pdf_path),
                        mime="application/pdf"
                    )
                st.success(f"Invoice #{single_inv_num} generated successfully!")

        # Bulk Generation
        st.subheader("Generate All Invoices")
        if st.button("Generate All Invoices"):
            if os.path.exists(OUTPUT_FOLDER):
                shutil.rmtree(OUTPUT_FOLDER)
            os.makedirs(OUTPUT_FOLDER)
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for i, (_, row) in enumerate(consolidated_df.iterrows()):
                status_text.text(f"Processing {i+1}/{len(consolidated_df)}: {row['MARK']}")
                progress_bar.progress((i + 1) / len(consolidated_df))
                
                pdf_path = generate_invoice(row, last_invoice + i)
                if pdf_path:
                    update_notification_sheet(
                        OUTPUT_FOLDER,
                        os.path.basename(pdf_path),
                        row["MARK"],
                        last_invoice + i,
                        row["CONTACT NUMBER"],
                        row["TOTAL CHARGES_SUM"]
                    )
            
            progress_bar.empty()
            status_text.success("Invoice generation complete!")
            
            # Create zip of all invoices
            zip_path = os.path.join(OUTPUT_FOLDER, "invoices.zip")
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for file in os.listdir(OUTPUT_FOLDER):
                    if file.endswith('.pdf'):
                        zipf.write(os.path.join(OUTPUT_FOLDER, file), file)
            
            st.markdown(get_binary_file_downloader_html(zip_path, "Download All Invoices"), unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
