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
from PIL import Image
import pythoncom
import win32com.client

# Cloud-friendly path configuration
TEMPLATE_PATH = "invoice_template.docx"
OUTPUT_FOLDER = "generated_invoices"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def sanitize_filename(name):
    return re.sub(r'[\\/:*?"<>|]', '_', name)

def convert_docx_to_pdf(docx_path, pdf_path):
    """More reliable DOCX to PDF conversion with multiple fallbacks"""
    try:
        # Method 1: Try Word COM object (best formatting)
        try:
            pythoncom.CoInitialize()
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(os.path.abspath(docx_path))
            doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
            doc.Close()
            word.Quit()
            pythoncom.CoUninitialize()
            return True
        except Exception as com_error:
            st.warning(f"Word COM conversion failed, trying alternatives: {str(com_error)}")
            
            # Method 2: Try pdf2docx
            try:
                cv = Converter(docx_path)
                cv.convert(pdf_path)
                cv.close()
                return True
            except Exception as pdf_error:
                st.warning(f"pdf2docx failed, trying FPDF: {str(pdf_error)}")
                
                # Method 3: Fallback to FPDF (basic)
                doc = Document(docx_path)
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", size=8)
                
                for para in doc.paragraphs:
                    pdf.cell(200, 5, txt=para.text, ln=True)
                
                pdf.output(pdf_path)
                return True
    except Exception as e:
        st.error(f"All PDF conversion methods failed: {str(e)}")
        return False

def generate_pdf_from_template(template_path, row_data, output_folder, invoice_number):
    try:
        # Load template
        doc = Document(template_path)
        
        # Set default font
        style = doc.styles['Normal']
        font = style.font
        font.size = Pt(8)
        
        # Format values
        for key, value in row_data.items():
            if isinstance(value, (int, float)):
                row_data[key] = f"{value:.2f}" if pd.notna(value) else ""
        
        # Set metadata
        current_date = datetime.now().strftime("%Y-%m-%d")
        row_data.update({
            "DATE": current_date,
            "INVOICE NUMBER": str(invoice_number)
        })
        
        # Process document
        for paragraph in doc.paragraphs:
            for key, value in row_data.items():
                for placeholder in [f"{{{{{key}}}}}", f"{{{{{key}.}}}}"]:
                    if placeholder in paragraph.text:
                        paragraph.text = paragraph.text.replace(placeholder, str(value))
        
        # Process tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for key, value in row_data.items():
                        for placeholder in [f"{{{{{key}}}}}", f"{{{{{key}.}}}}"]:
                            if placeholder in cell.text:
                                cell.text = cell.text.replace(placeholder, str(value))
        
        # Generate filename
        customer_name = row_data.get("MARK", "Customer")
        contact_number = row_data.get("CONTACT NUMBER", "")
        invoice_total = row_data.get("TOTAL CHARGES_SUM", "0.00")
        
        pdf_name = f"Invoice_{invoice_number}_{sanitize_filename(customer_name)}.pdf"
        pdf_path = os.path.join(output_folder, pdf_name)
        
        # Save and convert
        temp_docx = os.path.join(tempfile.gettempdir(), f"temp_{invoice_number}.docx")
        doc.save(temp_docx)
        
        if convert_docx_to_pdf(temp_docx, pdf_path):
            os.remove(temp_docx)
            return pdf_path
        return None
        
    except Exception as e:
        st.error(f"Template processing error: {str(e)}")
        return None

def generate_single_invoice(row, invoice_number):
    """Generate one invoice with preview option"""
    with st.spinner(f"Generating invoice #{invoice_number}..."):
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
            invoice_number
        )
        
        if pdf_path:
            st.success(f"Invoice #{invoice_number} generated successfully!")
            with open(pdf_path, "rb") as f:
                st.download_button(
                    label="Download Invoice",
                    data=f,
                    file_name=os.path.basename(pdf_path),
                    mime="application/pdf"
                )
            return pdf_path
        else:
            st.error("Failed to generate invoice")
            return None

# ... [rest of your existing functions remain the same] ...

# Streamlit UI Updates
st.title("Invoice Generation System")

# New Individual Invoice Section
if 'consolidated_df' in st.session_state:
    st.sidebar.header("Individual Invoice")
    customer = st.sidebar.selectbox(
        "Select Customer",
        options=st.session_state.consolidated_df["MARK"].unique()
    )
    inv_number = st.sidebar.number_input("Invoice Number", min_value=1, value=1)
    
    if st.sidebar.button("Generate Single Invoice"):
        row = st.session_state.consolidated_df[
            st.session_state.consolidated_df["MARK"] == customer
        ].iloc[0]
        generate_single_invoice(row, inv_number)

# In your main processing code, add:
if uploaded_file is not None:
    # ... [your existing processing code] ...
    
    # Store the consolidated data in session state
    st.session_state.consolidated_df = consolidated_df
    
    # Add individual generation in the main interface too
    st.header("Individual Invoice Generation")
    selected_customer = st.selectbox(
        "Customer", 
        options=consolidated_df["MARK"].unique()
    )
    single_inv_num = st.number_input("Invoice Number", min_value=1, value=last_invoice)
    
    if st.button("Generate This Invoice"):
        row = consolidated_df[consolidated_df["MARK"] == selected_customer].iloc[0]
        generate_single_invoice(row, single_inv_num)
