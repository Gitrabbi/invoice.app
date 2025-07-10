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
import zipfile
import subprocess
from typing import Optional

# Configuration
TEMPLATE_PATH = "invoice_template.docx"
OUTPUT_FOLDER = "generated_invoices"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def sanitize_filename(name: str) -> str:
    """Make strings safe for filenames"""
    return re.sub(r'[\\/:*?"<>|]', '_', name)

def convert_docx_to_pdf(docx_path: str, pdf_path: str) -> bool:
    """Convert DOCX to PDF using LibreOffice"""
    try:
        # First try headless LibreOffice conversion
        result = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", 
             os.path.dirname(pdf_path), docx_path],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=30
        )
        
        # Get the automatically generated PDF name
        base_name = os.path.splitext(os.path.basename(docx_path))[0]
        temp_pdf = os.path.join(os.path.dirname(pdf_path), f"{base_name}.pdf")
        
        if os.path.exists(temp_pdf):
            os.rename(temp_pdf, pdf_path)
            return True
        return False
        
    except Exception as e:
        st.error(f"PDF conversion failed. Please ensure LibreOffice is installed. Error: {str(e)}")
        return False

def validate_pdf(pdf_path: str) -> bool:
    """Verify PDF is valid"""
    try:
        with open(pdf_path, "rb") as f:
            return f.read(4) == b"%PDF"
    except:
        return False

def generate_pdf_from_template(
    template_path: str,
    row_data: dict,
    output_folder: str,
    invoice_number: int
) -> Optional[str]:
    """Generate PDF invoice from template"""
    try:
        doc = Document(template_path)
        style = doc.styles['Normal']
        style.font.size = Pt(8)

        # Format values
        formatted_data = {
            k: f"{v:.2f}" if isinstance(v, (int, float)) else str(v)
            for k, v in row_data.items()
        }
        formatted_data.update({
            "DATE": datetime.now().strftime("%Y-%m-%d"),
            "INVOICE NUMBER": str(invoice_number)
        })

        # Replace placeholders in paragraphs
        for paragraph in doc.paragraphs:
            for key, value in formatted_data.items():
                for ph in [f"{{{{{key}}}}}", f"{{{{{key}.}}}}"]:
                    if ph in paragraph.text:
                        paragraph.text = paragraph.text.replace(ph, value)

        # Replace placeholders in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for key, value in formatted_data.items():
                        for ph in [f"{{{{{key}}}}}", f"{{{{{key}.}}}}"]:
                            if ph in cell.text:
                                cell.text = cell.text.replace(ph, value)

        # Generate filename
        customer = sanitize_filename(formatted_data.get("MARK", "Customer"))
        pdf_name = f"Invoice_{invoice_number}_{customer}.pdf"
        pdf_path = os.path.join(output_folder, pdf_name)
        
        # Handle duplicates
        counter = 1
        while os.path.exists(pdf_path):
            pdf_name = f"Invoice_{invoice_number}_{customer}_{counter}.pdf"
            pdf_path = os.path.join(output_folder, pdf_name)
            counter += 1

        # Save and convert
        temp_docx = os.path.join(output_folder, f"temp_{invoice_number}.docx")
        doc.save(temp_docx)
        
        if convert_docx_to_pdf(temp_docx, pdf_path) and validate_pdf(pdf_path):
            os.remove(temp_docx)
            return pdf_path
        
        os.remove(temp_docx)
        return None

    except Exception as e:
        st.error(f"Template processing failed: {str(e)}")
        return None

def update_notification_sheet(output_folder: str, pdf_name: str, customer: str, 
                            invoice_number: int, contact: str, total: str):
    """Update tracking spreadsheet"""
    sheet_path = os.path.join(output_folder, "notification_log.xlsx")
    new_data = pd.DataFrame([{
        "Customer": customer,
        "Invoice": invoice_number,
        "Contact": contact,
        "Amount": total,
        "File": pdf_name
    }])
    
    if os.path.exists(sheet_path):
        try:
            existing = pd.read_excel(sheet_path)
            updated = pd.concat([existing, new_data])
            updated.to_excel(sheet_path, index=False)
            return
        except:
            pass
    new_data.to_excel(sheet_path, index=False)

def consolidate_data(df: pd.DataFrame) -> pd.DataFrame:
    """Process raw data into invoice-ready format"""
    consolidated = []
    for customer, group in df.groupby("MARK"):
        # Calculations
        total_cbm = group["CBM"].sum()
        charges = 10.00 if total_cbm < 0.05 else (group["CBM"] * group["PER CHARGES"]).sum()
        total = charges + group["PARKING CHARGES"].iloc[0]
        
        # Multi-line fields
        fields = ["RECEIPT NO.", "QTY", "DESCRIPTION", "CBM", "WEIGHT(KG)"]
        joined = {f: "\n".join(group[f].astype(str)) for f in fields}
        
        # Build record
        first = group.iloc[0]
        consolidated.append({
            **joined,
            "PARKING CHARGES": f"{group['PARKING CHARGES'].iloc[0]:.2f}",
            "PER CHARGES": f"{first['PER CHARGES']:.2f}",
            "TOTAL CHARGES": f"{total:.2f}",
            "MARK": customer,
            "CONTACT NUMBER": first.get("CONTACT NUMBER", ""),
            "CARGO NUMBER": first.get("CARGO NUMBER", ""),
            "TOTAL CBM": f"{total_cbm:.2f}",
            "TOTAL CHARGES_SUM": f"{total:.2f}"
        })
    return pd.DataFrame(consolidated)

def create_download_link(file_path: str, label: str) -> str:
    """Generate HTML download link"""
    with open(file_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    return f'<a href="data:application/octet-stream;base64,{b64}" download="{os.path.basename(file_path)}">{label}</a>'

def main():
    st.title("📄 Invoice Generation System")
    
    # Initialize session state
    if 'consolidated_df' not in st.session_state:
        st.session_state.consolidated_df = None
    
    # File Upload
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
    
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            
            # Initialize missing columns with defaults
            for col in ["PARKING CHARGES", "PER CHARGES", "Weight Rate"]:
                if col not in df.columns:
                    default_value = 0.0 if col != "Weight Rate" else 1.0
                    df[col] = default_value
            
            # Global Settings UI
            st.header("⚙️ Global Settings")
            col1, col2, col3 = st.columns(3)
            with col1:
                default_per_charge = st.number_input(
                    "Default Per Charge ($/CBM)",
                    value=float(df["PER CHARGES"].iloc[0]),
                    min_value=0.0,
                    step=0.1
                )
            with col2:
                default_weight_rate = st.number_input(
                    "Default Weight Rate (kg/CBM)",
                    value=float(df["Weight Rate"].iloc[0]),
                    min_value=0.1,
                    step=0.1
                )
            with col3:
                default_parking = st.number_input(
                    "Default Parking Charge ($)",
                    value=float(df["PARKING CHARGES"].iloc[0]),
                    min_value=0.0,
                    step=0.1
                )
            
            if st.button("💾 Apply Global Settings", key="apply_global"):
                df["PER CHARGES"] = default_per_charge
                df["Weight Rate"] = default_weight_rate
                df["PARKING CHARGES"] = default_parking
                st.success("Global settings applied to all customers!")
            
            # Process data
            df["CBM"] = df[["MEAS.(CBM)", "WEIGHT(KG)"]].max(axis=1)
            st.session_state.consolidated_df = consolidate_data(df)
            
            st.header("📊 Processed Data")
            st.dataframe(st.session_state.consolidated_df)
            
            # Invoice Generation
            st.header("🖨️ Invoice Generation")
            start_num = st.number_input(
                "Starting Invoice Number", 
                min_value=1, 
                value=1,
                key="invoice_start"
            )
            
            if st.button("🔄 Generate All Invoices", key="generate_all"):
                if os.path.exists(OUTPUT_FOLDER):
                    shutil.rmtree(OUTPUT_FOLDER)
                os.makedirs(OUTPUT_FOLDER)
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                for i, (_, row) in enumerate(st.session_state.consolidated_df.iterrows()):
                    status_text.text(f"Processing {i+1}/{len(st.session_state.consolidated_df)}: {row['MARK']}")
                    progress_bar.progress((i + 1) / len(st.session_state.consolidated_df))
                    
                    pdf_path = generate_pdf_from_template(
                        TEMPLATE_PATH,
                        row.to_dict(),
                        OUTPUT_FOLDER,
                        start_num + i
                    )
                    
                    if pdf_path:
                        update_notification_sheet(
                            OUTPUT_FOLDER,
                            os.path.basename(pdf_path),
                            row["MARK"],
                            start_num + i,
                            row["CONTACT NUMBER"],
                            row["TOTAL CHARGES_SUM"]
                        )
                
                # Create ZIP archive
                zip_path = os.path.join(OUTPUT_FOLDER, "invoices.zip")
                with zipfile.ZipFile(zip_path, 'w') as zipf:
                    for file in os.listdir(OUTPUT_FOLDER):
                        if file.endswith('.pdf'):
                            zipf.write(
                                os.path.join(OUTPUT_FOLDER, file), 
                                file
                            )
                
                st.markdown(
                    create_download_link(zip_path, "📥 Download All Invoices"),
                    unsafe_allow_html=True
                )
                st.success("✅ All invoices generated successfully!")
            
            # Single Invoice Generation
            st.subheader("🖨️ Single Invoice")
            if st.session_state.consolidated_df is not None:
                customer = st.selectbox(
                    "Select Customer",
                    options=st.session_state.consolidated_df["MARK"].unique(),
                    key="customer_select"
                )
                single_num = st.number_input(
                    "Invoice Number", 
                    min_value=1, 
                    value=start_num,
                    key="single_invoice_num"
                )
                
                if st.button("🖨️ Generate Selected Invoice", key="generate_single"):
                    row = st.session_state.consolidated_df[
                        st.session_state.consolidated_df["MARK"] == customer
                    ].iloc[0]
                    
                    pdf_path = generate_pdf_from_template(
                        TEMPLATE_PATH,
                        row.to_dict(),
                        OUTPUT_FOLDER,
                        single_num
                    )
                    
                    if pdf_path:
                        with open(pdf_path, "rb") as f:
                            st.download_button(
                                "📥 Download Invoice",
                                f,
                                file_name=os.path.basename(pdf_path),
                                mime="application/pdf"
                            )
                        st.success(f"✅ Invoice #{single_num} generated for {customer}!")
        
        except Exception as e:
            st.error(f"❌ Processing error: {str(e)}")

if __name__ == "__main__":
    main()
