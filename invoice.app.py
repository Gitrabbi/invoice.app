import os
import re
import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
import tempfile
import shutil
from datetime import datetime
import base64
from io import BytesIO
import zipfile
from fpdf import FPDF
from pdf2docx import Converter

# Cloud-friendly path configuration
TEMPLATE_PATH = "invoice_template.docx"
OUTPUT_FOLDER = "generated_invoices"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def sanitize_filename(name):
    return re.sub(r'[\\/:*?"<>|]', '_', name)

def convert_docx_to_pdf(docx_path, pdf_path):
    """Cloud-compatible DOCX to PDF conversion with fallbacks"""
    try:
        # First try pdf2docx (better formatting)
        cv = Converter(docx_path)
        cv.convert(pdf_path)
        cv.close()
        return True
    except Exception as e:
        st.warning(f"PDF conversion (pdf2docx) warning: {str(e)}")
        try:
            # Fallback to FPDF
            doc = Document(docx_path)
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=8)
            
            for para in doc.paragraphs:
                pdf.cell(200, 5, txt=para.text, ln=True)
                
            pdf.output(pdf_path)
            return True
        except Exception as e:
            st.error(f"PDF conversion (FPDF) failed: {str(e)}")
            return False

def generate_pdf_from_template(template_path, row_data, output_folder, invoice_number):
    try:
        doc = Document(template_path)
        
        # Set default font size to 8pt
        style = doc.styles['Normal']
        font = style.font
        font.size = Pt(8)
        
        # Format numeric values
        for key, value in row_data.items():
            if isinstance(value, (int, float)):
                row_data[key] = f"{value:.2f}" if pd.notna(value) else ""
        
        # Set invoice number and date
        current_date = datetime.now().strftime("%Y-%m-%d")
        row_data.update({
            "DATE": current_date,
            "INVOICE NUMBER": str(invoice_number)
        })
        
        # Add invoice header
        if len(doc.paragraphs) > 0:
            p = doc.paragraphs[0]
            p.text = f"Invoice #: {invoice_number}\nDate: {current_date}\n" + p.text
            for run in p.runs:
                run.font.size = Pt(8)
        else:
            p = doc.add_paragraph(f"Invoice #: {invoice_number}\nDate: {current_date}")
            p.style.font.size = Pt(8)
        
        # Replace placeholders
        for paragraph in doc.paragraphs:
            for key, value in row_data.items():
                for placeholder in [f"{{{{{key}}}}}", f"{{{{{key}.}}}}"]:
                    if placeholder in paragraph.text:
                        paragraph.text = paragraph.text.replace(placeholder, str(value))
            for run in paragraph.runs:
                run.font.size = Pt(8)
        
        # Process tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for key, value in row_data.items():
                        for placeholder in [f"{{{{{key}}}}}", f"{{{{{key}.}}}}"]:
                            if placeholder in cell.text:
                                cell.text = cell.text.replace(placeholder, str(value))
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(8)
        
        # Generate filename
        customer_name = row_data.get("MARK", "Customer")
        contact_number = row_data.get("CONTACT NUMBER", "")
        invoice_total = row_data.get("TOTAL CHARGES_SUM", "0.00")
        
        pdf_name = f"Invoice_{invoice_number}_{sanitize_filename(customer_name)}_{sanitize_filename(contact_number)}_{sanitize_filename(invoice_total)}.pdf"
        pdf_path = os.path.join(output_folder, pdf_name)
        
        # Handle filename conflicts
        counter = 1
        while os.path.exists(pdf_path):
            pdf_name = f"Invoice_{invoice_number}_{sanitize_filename(customer_name)}_{counter}.pdf"
            pdf_path = os.path.join(output_folder, pdf_name)
            counter += 1
        
        # Save and convert
        temp_docx = os.path.join(output_folder, f"temp_{invoice_number}.docx")
        doc.save(temp_docx)
        
        if convert_docx_to_pdf(temp_docx, pdf_path):
            os.remove(temp_docx)
            return pdf_path
        return None
        
    except Exception as e:
        st.error(f"Template processing error: {str(e)}")
        raise

def update_notification_sheet(output_folder, pdf_name, customer_name, invoice_number, contact_number, invoice_total):
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
st.title("Invoice Generation System")

# Sidebar
st.sidebar.header("Quick Actions")
st.sidebar.markdown("[Sample Excel Template](#)")

# File Upload
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

        # Global Updates Section
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
        
        st.header("Processed Data")
        st.dataframe(consolidated_df)

        # Invoice Generation
        st.header("Invoice Generation")
        last_invoice = st.number_input("Starting Invoice Number", min_value=1, value=1)
        
        if st.button("Generate All Invoices"):
            if os.path.exists(OUTPUT_FOLDER):
                shutil.rmtree(OUTPUT_FOLDER)
            os.makedirs(OUTPUT_FOLDER)
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for i, (_, row) in enumerate(consolidated_df.iterrows()):
                status_text.text(f"Processing {i+1}/{len(consolidated_df)}: {row['MARK']}")
                progress_bar.progress((i + 1) / len(consolidated_df))
                
                template_data = {
                    "RECEIPT NO": row["RECEIPT NO."],
                    "QTY": row["QTY"],
                    "DESCRIPTION": row["DESCRIPTION"],
                    "WEIGHT(KG)": row["WEIGHT(KG)"],
                    "CONTACT NUMBER": row["CONTACT NUMBER"],
                    "CBM": row["CBM"],
                    "PER CHARGES": row["PER CHARGES"],
                    "PARKING CHARGES": row["PARKING CHARGES"],
                    "TOTAL CHARGES": row["TOTAL CHARGES_SUM"],
                    "MARK": row["MARK"],
                    "TOTAL QTY": row["TOTAL QTY"],
                    "TOTAL CBM": row["TOTAL CBM"],
                    "CARGO NUMBER": row["CARGO NUMBER"],
                    "TRACKING NUMBER": row["TRACKING NUMBER"],
                    "TERMS": row["TERMS"]
                }
                
                pdf_path = generate_pdf_from_template(
                    TEMPLATE_PATH,
                    template_data,
                    OUTPUT_FOLDER,
                    last_invoice + i
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