import pdfplumber
import pandas as pd
import re
from datetime import datetime
import logging

# === Config ===
pdf_path = "1_C25TDM_00030840.pdf"
output_path = "extracted_invoice_data.xlsx"
order_date = "02-07-2025"  # Manual or parsed externally

# === Logging setup ===
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

# === Read PDF ===
with pdfplumber.open(pdf_path) as pdf:
    page = pdf.pages[0]
    text = page.extract_text()
    tables = page.extract_tables()

logging.info(f"→ Combined invoice_number = {text}")
# === Extract invoice metadata ===
invoice_date_match = re.search(r'Ngày \(Date\) (\d{1,2}) tháng \(month\) (\d{1,2}) năm \(year\) (\d{4})', text)
invoice_date = f"{int(invoice_date_match.group(1)):02d}-{int(invoice_date_match.group(2)):02d}-{invoice_date_match.group(3)}" if invoice_date_match else ""

# 1. Extract the series
series_match = re.search(r'Ký hiệu \(Series\):\s*(\S+)', text)
series = series_match.group(1) if series_match else None

# 2. Extract the number — look for a standalone 8-digit number near 'Số (No.)'
number_match = re.search(r'(\d{8})\s*Số\s*\(No\.\)', text)
number = number_match.group(1) if number_match else None

# Combine if both are found
if series and number:
    invoice_number = f"{number}.{series}"
    logging.info(f"✅ Found invoice number: Series = {series}, Number = {number}")
    logging.info(f"→ Combined invoice_number = {invoice_number}")
else:
    invoice_number = ""
    logging.warning("⚠️ Could not extract invoice number and/or series.")

buyer_match = re.search(r'Họ tên người mua hàng \(Buyer\):\s+(.*?)\s+Tên đơn vị', text, re.DOTALL)
buyer_name = buyer_match.group(1).strip().replace("\n", " ") if buyer_match else ""

vat_match = re.search(r'Thuế suất GTGT.*?:\s+(\d+)%.*?Tiền thuế GTGT.*?:\s+([\d.]+)', text)
vat_percent = vat_match.group(1) + "%" if vat_match else "8%"
vat_amount = vat_match.group(2).replace(".", "") if vat_match else "0"

total_before_tax_match = re.search(r'Cộng tiền hàng.*?:\s+([\d.]+)', text)
total_before_tax = total_before_tax_match.group(1).replace(".", "") if total_before_tax_match else "0"

total_after_tax_match = re.search(r'Tổng cộng tiền thanh toán.*?:\s+([\d.]+)', text)
total_after_tax = total_after_tax_match.group(1).replace(".", "") if total_after_tax_match else "0"

# === Extract product rows ===
product_rows = []
for t in tables:
    for row in t:
        if row:
            if len(row) > 1 and isinstance(row[1], str) and re.match(r"\d{13}", row[1]):
                product_rows.append(row)
            else:
                logging.warning(f"Skipped row (not a product): {row}")

# === Build final Excel rows ===
rows = []

for i, row in enumerate(product_rows):
    try:
        stt = row[0]
        product_code = row[1]
        product_name = row[2].replace("\n", " ").strip()
        unit = row[3]
        quantity = float(row[4])
        unit_price = float(row[5].replace(",", "").replace(".", ""))
        total = float(row[6].replace(",", "").replace(".", ""))
    except Exception as e:
        logging.error(f"Error parsing product row {i}: {row} → {e}")
        continue

    rows.append({
        "Ngày đặt hàng": order_date if i == 0 else "",
        "Ngày hóa đơn": invoice_date if i == 0 else "",
        "Số hóa đơn": invoice_number if i == 0 else "",
        "Tên khách hàng": buyer_name if i == 0 else "",
        "Tên đơn vị": "-" if i == 0 else "",
        "Địa chỉ": "-" if i == 0 else "",
        "Mã sản phẩm": product_code,
        "Tên sản phẩm": product_name,
        "Đơn vị tính": unit,
        "Số lượng": quantity,
        "Đơn giá bán (vnđ)": unit_price,
        "Thành tiền trước thuế (vnđ)": total_before_tax if i == 0 else "",
        "VAT %": vat_percent if i == 0 else "",
        "Tiền thuế (vnđ)": vat_amount if i == 0 else "",
        "Thành tiền sau thuế (vnđ)": total_after_tax if i == 0 else "",
        "Ghi chú": "TRUE",
        "Đơn giá bán (HD)": unit_price
    })

# === Save to Excel ===
df = pd.DataFrame(rows)

try:
    existing = pd.read_excel(output_path)
    result = pd.concat([existing, df], ignore_index=True)
except FileNotFoundError:
    result = df

result.to_excel(output_path, index=False)
print(f"✅ Exported {len(rows)} rows to '{output_path}'")

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# === Step 6: Write Excel file (overwrite or append) ===
df = pd.DataFrame(rows)
df.to_excel(output_path, index=False)

# === Step 7: Merge cells for invoice fields ===
# Columns to merge (by name)
merge_columns = [
    "Ngày đặt hàng",
    "Ngày hóa đơn",
    "Số hóa đơn",
    "Tên khách hàng",
    "Tên đơn vị",
    "Địa chỉ",
    "Thành tiền trước thuế (vnđ)",
    "VAT %",
    "Tiền thuế (vnđ)",
    "Thành tiền sau thuế (vnđ)",
]

# Reload with openpyxl to merge cells
wb = load_workbook(output_path)
ws = wb.active

# Find column indexes for the merge columns
header = [cell.value for cell in ws[1]]
merge_col_indexes = {col: header.index(col) + 1 for col in merge_columns}

# Calculate row range per invoice group (rows with same invoice number)
start_row = 2  # skip header
while start_row <= ws.max_row:
    invoice_no = ws.cell(row=start_row, column=merge_col_indexes["Số hóa đơn"]).value
    end_row = start_row
    # Find how many rows share the same invoice number
    while (
        end_row + 1 <= ws.max_row and
        ws.cell(row=end_row + 1, column=merge_col_indexes["Số hóa đơn"]).value in [None, "", invoice_no]
    ):
        end_row += 1

    # Merge cells in all specified columns
    for col_name, col_idx in merge_col_indexes.items():
        if end_row > start_row:
            ws.merge_cells(
                start_row=start_row,
                start_column=col_idx,
                end_row=end_row,
                end_column=col_idx
            )

    start_row = end_row + 1  # move to next invoice group

# Save the merged result
wb.save(output_path)
print(f"✅ Excel written with merged rows → {output_path}")
