import os
import re
import pdfplumber
import pandas as pd
from datetime import datetime
import logging
from openpyxl import load_workbook

# === Configuration ===
PDF_FOLDER = "./invoices"  # Folder containing your PDFs
OUTPUT_FILE = "combined_invoices.xlsx"
ORDER_DATE_DEFAULT = "01-07-2025"  # Optional default value

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
SKIP_FREE_GIFTS = True  # Set to False if you want to include gifts

all_rows = []

# === Loop through all PDFs ===
for filename in os.listdir(PDF_FOLDER):
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(PDF_FOLDER, filename)
        logging.info(f"🔍 Processing: {filename}")

        try:
            with pdfplumber.open(pdf_path) as pdf:
                page = pdf.pages[0]
                text = page.extract_text()
                tables = page.extract_tables()

            # === Extract Invoice Metadata ===
            series_match = re.search(r'Ký hiệu \(Series\):\s*(\S+)', text)
            series = series_match.group(1) if series_match else ""
            number_match = re.search(r'(\d{6,8})\s*Số\s*\(No\.\)', text)
            number = number_match.group(1) if number_match else ""
            invoice_number = f"{number}.{series}" if number and series else ""

            date_match = re.search(r'Ngày \(Date\) (\d{1,2}) tháng \(month\) (\d{1,2}) năm \(year\) (\d{4})', text)
            invoice_date = f"{int(date_match.group(1)):02d}-{int(date_match.group(2)):02d}-{date_match.group(3)}" if date_match else ""

            buyer_match = re.search(r'Họ tên người mua hàng \(Buyer\):\s+(.*?)\s+Tên đơn vị', text, re.DOTALL)
            buyer_name = buyer_match.group(1).strip().replace("\n", " ") if buyer_match else ""

            vat_match = re.search(r'Thuế suất GTGT.*?:\s+(\d+)%.*?Tiền thuế GTGT.*?:\s+([\d.]+)', text)
            vat_percent = vat_match.group(1) + "%" if vat_match else ""
            vat_amount = vat_match.group(2).replace(".", "") if vat_match else ""

            total_match = re.search(r'Cộng tiền hàng.*?:\s+([\d.]+)', text)
            total_before_tax = total_match.group(1).replace(".", "") if total_match else ""

            total_with_vat = re.search(r'Tổng cộng tiền thanh toán.*?:\s+([\d.]+)', text)
            total_after_tax = total_with_vat.group(1).replace(".", "") if total_with_vat else ""

            # === Extract Product Rows ===
            product_rows = []
            for t in tables:
                for row in t:
                    if (
                        row and len(row) >= 7 and
                        isinstance(row[1], str) and
                        re.match(r"\d{13}", row[1])
                    ):
                        product_rows.append(row)

            if not product_rows:
                logging.warning(f"⚠️ No products found in {filename}")
                continue

            for i, row in enumerate(product_rows):
                try:
                    product_code = row[1]
                    product_name = row[2].replace("\n", " ").strip()
                    
                    # 🔥 Skip "hàng tặng không bán"
                    if SKIP_FREE_GIFTS and "hàng tặng không bán" in product_name.lower():
                        logging.info(f"⏭️ Skipping gift row: {product_name}")
                        continue

                    unit = row[3]
                    quantity = float(row[4])
                    unit_price = float(row[5].replace(",", "").replace(".", ""))
                    total = float(row[6].replace(",", "").replace(".", ""))
                except Exception as e:
                    logging.error(f"❌ Error in product row {i}: {row} → {e}")
                    continue

                all_rows.append({
                    "Ngày đặt hàng": ORDER_DATE_DEFAULT if i == 0 else "",
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

        except Exception as e:
            logging.error(f"💥 Failed to process {filename}: {e}")

# === Write to Excel ===
df = pd.DataFrame(all_rows)
df.to_excel(OUTPUT_FILE, index=False)

# Now insert dynamic formulas into "Ghi chú" column
wb = load_workbook(OUTPUT_FILE)
ws = wb.active

# Find column indexes
headers = [cell.value for cell in ws[1]]
ghi_chu_col = headers.index("Ghi chú") + 1  # Excel columns are 1-based
col_Q = headers.index("Đơn giá bán (HD)") + 1
col_K = headers.index("Đơn giá bán (vnđ)") + 1

for row in range(2, ws.max_row + 1):
    cell_ref_q = f"{chr(64 + col_Q)}{row}"  # e.g., Q15
    cell_ref_k = f"{chr(64 + col_K)}{row}"  # e.g., K15
    formula = f"={cell_ref_q}={cell_ref_k}"
    ws.cell(row=row, column=ghi_chu_col).value = formula

wb.save(OUTPUT_FILE)
logging.info(f"✅ Saved all invoices to {OUTPUT_FILE}")

# === Merge invoice-level fields in Excel ===
from openpyxl import load_workbook

merge_columns = [
    "Ngày đặt hàng", "Ngày hóa đơn", "Số hóa đơn", "Tên khách hàng", "Tên đơn vị", "Địa chỉ",
    "Thành tiền trước thuế (vnđ)", "VAT %", "Tiền thuế (vnđ)", "Thành tiền sau thuế (vnđ)"
]

wb = load_workbook(OUTPUT_FILE)
ws = wb.active
header = [cell.value for cell in ws[1]]
merge_col_indexes = {col: header.index(col) + 1 for col in merge_columns}

start_row = 2
while start_row <= ws.max_row:
    invoice = ws.cell(row=start_row, column=merge_col_indexes["Số hóa đơn"]).value
    end_row = start_row
    while (
        end_row + 1 <= ws.max_row and
        ws.cell(row=end_row + 1, column=merge_col_indexes["Số hóa đơn"]).value in [None, "", invoice]
    ):
        end_row += 1

    for col, idx in merge_col_indexes.items():
        if end_row > start_row:
            ws.merge_cells(start_row=start_row, start_column=idx, end_row=end_row, end_column=idx)

    start_row = end_row + 1

wb.save(OUTPUT_FILE)
logging.info("✅ Merged invoice-level cells in Excel.")
