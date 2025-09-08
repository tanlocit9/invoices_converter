import os
import re
import sys
import json
import logging
from typing import List, Dict, Any

import pdfplumber
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ========================
#  Portable Base Directory
# ========================
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)  # when packed as .exe
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ========================
#  Logging Setup
# ========================
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

# ========================
#  Load Config
# ========================
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
if not os.path.exists(CONFIG_PATH):
    logging.error(f"Config file not found: {CONFIG_PATH}")
    sys.exit(1)

with open(CONFIG_PATH, "r", encoding="utf-8") as f:
    config = json.load(f)

def resolve_path(p: str) -> str:
    # Allow absolute or relative; relative becomes BASE_DIR/path
    return p if os.path.isabs(p) else os.path.join(BASE_DIR, p)

PDF_FOLDER = resolve_path(config.get("PDF_FOLDER", "promotions"))
OUTPUT_FILE = resolve_path(config.get("OUTPUT_FILE", "combined_promotions.xlsx"))
ORDER_DATE_DEFAULT = config.get("ORDER_DATE_DEFAULT", "")
SKIP_FREE_GIFTS = bool(config.get("SKIP_FREE_GIFTS", True))
APPEND_IF_EXISTS = bool(config.get("APPEND_IF_EXISTS", False))  # append to existing excel

# ========================
#  Helpers
# ========================
def to_float_safe(val: Any) -> float:
    """
    Convert numbers with Vietnamese thousands separators or commas safely.
    Examples:
      "1.234.567" -> 1234567
      "1,234.5"   -> 1234.5
      "1.234,5"   -> 1234.5
    """
    if val is None:
        return 0.0
    s = str(val).strip()
    if not s:
        return 0.0

    if "." in s and "," in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        if s.count(".") > 1 and "," not in s:
            s = s.replace(".", "")
        elif s.count(",") > 1 and "." not in s:
            s = s.replace(",", "")
        elif "," in s and "." not in s:
            s = s.replace(",", ".")

    try:
        return float(s)
    except ValueError:
        return 0.0

def extract_invoice_meta(text: str) -> Dict[str, str]:
    series = ""
    number = ""
    invoice_date = ""
    buyer_name = ""
    vat_percent = ""
    vat_amount = ""
    total_before_tax = ""
    total_after_tax = ""

    series_match = re.search(r'Ký\s*hiệu\s*\(Series\)\s*:\s*(\S+)', text, re.IGNORECASE)
    if series_match: series = series_match.group(1)

    number_match = re.search(r'(\d{6,10})\s*Số\s*\(No\.\)', text, re.IGNORECASE)
    if number_match: number = number_match.group(1)

    date_match = re.search(
        r'Ngày\s*\(Date\)\s*(\d{1,2})\s+tháng\s*\(month\)\s*(\d{1,2})\s+năm\s*\(year\)\s*(\d{4})',
        text, re.IGNORECASE
    )
    if date_match:
        d, m, y = int(date_match.group(1)), int(date_match.group(2)), date_match.group(3)
        invoice_date = f"{d:02d}-{m:02d}-{y}"

    buyer_match = re.search(
        r'Họ\s*tên\s*người\s*mua\s*hàng\s*\(Buyer\)\s*:\s+(.*?)\s+Tên\s*đơn\s*vị',
        text, re.IGNORECASE | re.DOTALL
    )
    if buyer_match:
        buyer_name = buyer_match.group(1).strip().replace("\n", " ")

    vat_match = re.search(
        r'Thuế\s*suất\s*GTGT.*?:\s*([0-9]+)%.*?Tiền\s*thuế\s*GTGT.*?:\s*([\d\.,]+)',
        text, re.IGNORECASE | re.DOTALL
    )
    if vat_match:
        vat_percent = vat_match.group(1) + "%"
        vat_amount = vat_match.group(2).replace(".", "").replace(",", "")

    total_match = re.search(r'Cộng\s*tiền\s*hàng.*?:\s*([\d\.,]+)', text, re.IGNORECASE)
    if total_match:
        total_before_tax = total_match.group(1).replace(".", "").replace(",", "")

    total_with_vat = re.search(r'Tổng\s*cộng\s*tiền\s*thanh\s*toán.*?:\s*([\d\.,]+)', text, re.IGNORECASE)
    if total_with_vat:
        total_after_tax = total_with_vat.group(1).replace(".", "").replace(",", "")

    invoice_number = f"{number}.{series}" if number and series else (number or "")

    return {
        "invoice_number": invoice_number,
        "invoice_date": invoice_date,
        "buyer_name": buyer_name,
        "vat_percent": vat_percent,
        "vat_amount": vat_amount,
        "total_before_tax": total_before_tax,
        "total_after_tax": total_after_tax,
    }

def find_product_rows(tables: List[List[List[Any]]]) -> List[List[Any]]:
    out = []

    # --- Step 1: check the special conditions ---
    try:
        if tables[0][8][12] == "Time applying":
            table0 = fill_down(tables[0])  # fill missing values
            out.extend(transform_row(r) for r in table0[9:] if r)
        if tables[2][0][3] == "Time applying":
            table2 = fill_down(tables[2])
            out.extend(table2[2][1:-1])  # rows after index 0
    except (IndexError, TypeError):
        # Ignore if the tables are not deep enough
        pass

    # --- Step 2: regular product row detection ---
    for t in tables or []:
        for row in t or []:
            if (
                row and len(row) >= 7
                and isinstance(row[1], str)
                and re.fullmatch(r"\d{13}", row[1].strip())
            ):
                out.append(row)

    return out

def walk_pdfs(folder: str) -> List[str]:
    files = []
    for root, _, filenames in os.walk(folder):
        for fn in filenames:
            if fn.lower().endswith(".pdf"):
                files.append(os.path.join(root, fn))
    return sorted(files)

def transform_row(row: list) -> list:
    """
    Transform a raw row into the compact format:
    [id, shop_name, product_name, date_range, amount]
    """
    if not row:
        return []

    # Example: pick fixed positions
    return [
        row[0],  # or generate sequentially instead of hardcoding
        row[1] if len(row) > 1 else "",                 # shop name
        row[3] if len(row) > 3 else "",                 # product name
        row[12] if len(row) > 12 else "",               # date range
        row[14] if len(row) > 14 else "",               # amount
    ]

def fill_down(table: List[List[Any]]) -> List[List[Any]]:
    """
    For each column in the table, if a row has None, 
    copy the value from the previous row in the same column.
    """
    for i in range(1, len(table)):  # start from 2nd row
        for j in range(len(table[i])):
            if table[i][j] is None and j < len(table[i-1]):
                table[i][j] = table[i-1][j]
    return table

def split_products(table: List[List[Any]], product_col: int = 2) -> List[List[Any]]:
    """
    Split rows where product_col contains ';' into multiple rows,
    each with one product.
    """
    new_table = []
    for row in table:
        if row and len(row) > product_col and isinstance(row[product_col], str) and ";" in row[product_col]:
            products = [p.strip() for p in row[product_col].split(";") if p.strip()]
            for p in products:
                new_row = row.copy()
                new_row[product_col] = p
                new_table.append(new_row)
        else:
            new_table.append(row)
    return new_table

def normalize_rows(rows: List[List[Any]]) -> List[List[Any]]:
    """
    Convert product rows so that trailing '(xx%)' (with decimals/newlines/trailing dot) 
    is extracted into its own column.
    """
    out = []
    # regex: capture product + percent (allow decimals, optional . or whitespace after %)
    pattern = re.compile(r"^(.*)\((\d+(?:\.\d+)?)%\)\.?\s*$")

    for row in rows:
        if not isinstance(row, list):
            continue  # skip junk

        new_row = list(row)
        if len(new_row) >= 3 and isinstance(new_row[2], str):
            cell = new_row[2].replace("\n", " ").strip()
            m = pattern.match(cell)
            if m:
                product, percent = m.groups()
                new_row[2] = product.strip()
                new_row.insert(3, f"{percent}%")
        out.append(new_row)

    return out

# ========================
#  Data Storage
# ========================
all_rows: List[Dict[str, Any]] = []

# ========================
#  Process PDFs
# ========================
if not os.path.isdir(PDF_FOLDER):
    logging.error(f"PDF folder not found: {PDF_FOLDER}")
    sys.exit(1)

pdf_files = walk_pdfs(PDF_FOLDER)
if not pdf_files:
    logging.warning(f"No PDFs found in: {PDF_FOLDER}")

for pdf_path in pdf_files:
    filename = os.path.basename(pdf_path)
    logging.info(f"🔍 Processing: {filename}")

    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = []
            all_tables = []
            for page in pdf.pages:
                full_text.append(page.extract_text() or "")
                all_tables.extend(page.extract_tables() or [])
            text = "\n".join(full_text)

        meta = extract_invoice_meta(text)
        product_rows = find_product_rows(all_tables)
        split_products_rows = split_products(product_rows)
        all_rows = normalize_rows(split_products_rows)

    except Exception as e:
        logging.error(f"💥 Failed to process {filename}: {e}")

if not all_rows:
    logging.info("No rows to write. Exiting.")
    sys.exit(0)

# ========================
#  Write to Excel
# ========================
df = pd.DataFrame(all_rows, columns=["STT", "Customer", "SKU", "%", "Timeline - Delivery", "Amount"])

if APPEND_IF_EXISTS and os.path.exists(OUTPUT_FILE):
    try:
        existing = pd.read_excel(OUTPUT_FILE)
        df = pd.concat([existing, df], ignore_index=True)
    except Exception as e:
        logging.warning(f"Could not append to existing file ({OUTPUT_FILE}). Overwriting. Reason: {e}")

os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
df.to_excel(OUTPUT_FILE, index=False)
logging.info(f"✅ All done! Saved to {OUTPUT_FILE}")