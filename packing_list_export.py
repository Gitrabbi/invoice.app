import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

def export_packing_list(df: pd.DataFrame):
    buffer = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Packing List Print"

    # Define header row
    headers = [
        "MARK", "CONTACT NUMBER", "CARGO NUMBER", "TRACKING NUMBER", "TERMS",
        "RECEIPT NO.", "DESCRIPTION", "QTY", "CBM", "WEIGHT(KG)",
        "Weight Rate", "PER CHARGES", "PARKING CHARGES", "TOTAL CHARGES"
    ]
    ws.append(headers)

    # Style header
    header_font = Font(bold=True)
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Write data row by row per customer
    row_start = 2
    for _, row in df.iterrows():
        item_count = row["RECEIPT NO."].count("\n") + 1 if "\n" in row["RECEIPT NO."] else 1

        ws.cell(row=row_start, column=1).value = row["MARK"]
        ws.cell(row=row_start, column=2).value = row.get("CONTACT NUMBER", "")
        ws.cell(row=row_start, column=3).value = row.get("CARGO NUMBER", "")
        ws.cell(row=row_start, column=4).value = row.get("TRACKING NUMBER", "")
        ws.cell(row=row_start, column=5).value = row.get("TERMS", "")

        # Split multiline fields
        receipt_list = str(row["RECEIPT NO."]).split("\n")
        desc_list = str(row["DESCRIPTION"]).split("\n")
        qty_list = str(row["QTY"]).split("\n")
        cbm_list = str(row["CBM"]).split("\n")
        weight_list = str(row["WEIGHT(KG)"]).split("\n")

        for i in range(item_count):
            r = row_start + i
            ws.cell(row=r, column=6).value = receipt_list[i] if i < len(receipt_list) else ""
            ws.cell(row=r, column=7).value = desc_list[i] if i < len(desc_list) else ""
            ws.cell(row=r, column=8).value = qty_list[i] if i < len(qty_list) else ""
            ws.cell(row=r, column=9).value = cbm_list[i] if i < len(cbm_list) else ""
            ws.cell(row=r, column=10).value = weight_list[i] if i < len(weight_list) else ""

        row_after_items = row_start + item_count

        if item_count > 1:
            # Add totals row only if more than one item
            total_row = row_after_items
            ws.cell(row=total_row, column=1).value = f"{row['MARK']} TOTALS"
            ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=5)

            ws.cell(row=total_row, column=11).value = float(row.get("Weight Rate", 0))
            ws.cell(row=total_row, column=12).value = float(row.get("PER CHARGES", 0))
            ws.cell(row=total_row, column=13).value = float(row.get("PARKING CHARGES", 0))
            ws.cell(row=total_row, column=14).value = float(row.get("TOTAL CHARGES", 0))

            # Apply sea blue fill
            fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
            for col in range(1, 15):
                cell = ws.cell(row=total_row, column=col)
                cell.fill = fill

            row_start = total_row + 2
        else:
            row_start = row_after_items + 1

    # Adjust column widths safely
    for col in ws.columns:
        first_cell = next((cell for cell in col if cell and cell.value is not None), None)
        if first_cell:
            col_letter = first_cell.column_letter
            ws.column_dimensions[col_letter].width = 20

    # Save to buffer
    wb.save(buffer)
    buffer.seek(0)
    return buffer
