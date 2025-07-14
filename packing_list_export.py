import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

def export_packing_list(df):
    # Prepare output Excel file
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Packing List Print"

    # Styles
    bold_center = Font(bold=True)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sea_blue_fill = PatternFill(start_color="B7DEE8", end_color="B7DEE8", fill_type="solid")

    # Header content (adjust as needed)
    ws.merge_cells("A1:N1")
    ws["A1"] = "AUGUST CARGO LOGISTICS"
    ws["A1"].font = Font(size=14, bold=True)
    ws["A1"].alignment = center_align

    ws.merge_cells("A2:N2")
    ws["A2"] = "PACKING LIST"
    ws["A2"].font = Font(size=12, bold=True)
    ws["A2"].alignment = center_align

    ws.append([])  # Row 3 empty
    ws.append([])  # Row 4 empty

    # Table headers (starting at row 5)
    headers = [
        "RECEIPT NO.", "MARK", "DESCRIPTION", "QTY", "CTNS", "MEAS. (CBM)", "WEIGHT(KG)",
        "WEIGHT RATE", "WEIGHT CBM", "CBM", "PER CHARGES", "TOTAL CHARGES",
        "CONTACT NUMBER", "业务员/ Supplier"
    ]
    ws.append(headers)
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=5, column=col)
        cell.font = bold_center
        cell.alignment = center_align

    # Group by customer (MARK)
    grouped = df.groupby("MARK")

    for mark, group in grouped:
        group = group.copy()

        # Add row(s) for each item
        for _, row in group.iterrows():
            ws.append([
                row.get("RECEIPT NO.", ""),
                row.get("MARK", ""),
                row.get("DESCRIPTION", ""),
                row.get("QTY", ""),
                "",  # CTNS not available
                row.get("CBM", ""),
                row.get("WEIGHT(KG)", ""),
                row.get("Weight Rate", ""),
                "",  # WEIGHT CBM not directly available
                row.get("CBM", ""),
                row.get("PER CHARGES", ""),
                row.get("TOTAL CHARGES", ""),
                row.get("CONTACT NUMBER", ""),
                ""  # Supplier field
            ])

        # If more than one row, add subtotal in sea blue row
        if len(group) > 1:
            total_qty = group["QTY"].astype(float).sum()
            total_cbm = group["CBM"].astype(float).sum()
            total_weight = group["WEIGHT(KG)"].astype(float).sum()
            total_charges = group["TOTAL CHARGES"].astype(float).sum()

            total_row = [
                "", f"{mark} TOTAL", "", total_qty, "", "", total_weight, "", "", total_cbm, "", total_charges, "", ""
            ]
            ws.append(total_row)

            # Apply sea blue fill
            last_row = ws.max_row
            for col in range(1, len(headers) + 1):
                cell = ws.cell(row=last_row, column=col)
                cell.fill = sea_blue_fill
                cell.font = bold_center
                cell.alignment = center_align

    # Adjust column widths
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max(10, min(max_length + 2, 25))

    # Save to BytesIO
    wb.save(output)
    output.seek(0)
    return output
