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
    """Convert DOCX to PDF using LibreOffice's unoconv"""
    try:
        result = subprocess.run(
            ["unoconv", "-f", "pdf", "-o", pdf_path, docx_path],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=30  # Prevent hanging
        )
        return os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0
    except subprocess.TimeoutExpired:
        st.error("PDF conversion timed out (30s)")
        return False
    except subprocess.CalledProcessError as e:
        st.error(f"LibreOffice conversion failed: {e.stderr.decode()}")
        return False
    except Exception as e:
        st.error(f"Unexpected conversion error: {str(e)}")
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
        # Load and modify template
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
                        for run in paragraph.runs:
                            run.font.size = Pt(8)

        # Replace placeholders in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for key, value in formatted_data.items():
                        for ph in [f"{{{{{key}}}}}", f"{{{{{key}.}}}}"]:
                            if ph in cell.text:
                                cell.text = cell.text.replace(ph, value)
                                for paragraph in cell.paragraphs:
                                    for run in paragraph.runs:
                                        run.font.size = Pt(8)

        # Generate unique filename
        customer = sanitize_filename(formatted_data.get("MARK", "Customer"))
        pdf_name = f"Invoice_{invoice_number}_{customer}.pdf"
        pdf_path = os.path.join(output_folder, pdf_name)
        
        # Handle duplicates
        counter = 1
        while os.path.exists(pdf_path):
            pdf_name = f"Invoice_{invoice_number}_{customer}_{counter}.pdf"
            pdf_path = os.path.join(output_folder, pdf_name)
            counter += 1

        # Convert to PDF
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

# Streamlit UI
def main():
    st.title("Invoice Generation System")
    
    # File Upload
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
    
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            
            # Initialize missing columns
            for col in ["PARKING CHARGES", "PER CHARGES", "Weight Rate"]:
                if col not in df.columns:
                    df[col] = 0.0 if col != "Weight Rate" else 1.0
            
            # Global Settings UI
            st.header("Global Settings")
            defaults = {
                "PER CHARGES": st.number_input("Default Rate", value=float(df["PER CHARGES"].iloc[0])),
                "PARKING CHARGES": st.number_input("Parking Fee", value=0.0)
            }
            
            if st.button("Apply Globally"):
                df["PER CHARGES"] = defaults["PER CHARGES"]
                df["PARKING CHARGES"] = defaults["PARKING CHARGES"]
                st.success("Settings applied!")
            
            # Process data
            df["CBM"] = df[["MEAS.(CBM)", "WEIGHT(KG)"]].max(axis=1)
            consolidated_df = consolidate_data(df)
            
            st.header("Processed Data")
            st.dataframe(consolidated_df)
            
            # Invoice Generation
            st.header("Invoice Generation")
            start_num = st.number_input("Starting Invoice #", min_value=1, value=1)
            
            if st.button("Generate All Invoices"):
                if os.path.exists(OUTPUT_FOLDER):
                    shutil.rmtree(OUTPUT_FOLDER)
                os.makedirs(OUTPUT_FOLDER)
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                for i, (_, row) in enumerate(consolidated_df.iterrows()):
                    status_text.text(f"Processing {i+1}/{len(consolidated_df)}")
                    progress_bar.progress((i + 1) / len(consolidated_df))
                    
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
                    create_download_link(zip_path, "Download All Invoices"),
                    unsafe_allow_html=True
                )
                st.success("All invoices generated!")
            
            # Single Invoice UI
            st.subheader("Single Invoice")
            customer = st.selectbox(
                "Select Customer",
                options=consolidated_df["MARK"].unique()
            )
            single_num = st.number_input(
                "Invoice Number", 
                min_value=1, 
                value=start_num
            )
            
            if st.button("Generate Selected Invoice"):
                row = consolidated_df[consolidated_df["MARK"] == customer].iloc[0]
                pdf_path = generate_pdf_from_template(
                    TEMPLATE_PATH,
                    row.to_dict(),
                    OUTPUT_FOLDER,
                    single_num
                )
                
                if pdf_path:
                    with open(pdf_path, "rb") as f:
                        st.download_button(
                            "Download Invoice",
                            f,
                            file_name=os.path.basename(pdf_path),
                            mime="application/pdf"
                        )
                    st.success("Invoice generated!")
        
        except Exception as e:
            st.error(f"Processing error: {str(e)}")

if __name__ == "__main__":
    main()
