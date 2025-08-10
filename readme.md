# 📄 PDF Invoice to Excel Converter

A Python-based tool to batch convert PDF invoices into a single Excel file with:
- Multiple product rows per invoice
- Merged invoice-level fields
- Optional skipping of gift items
- Dynamic Excel formulas for data validation

---

## 📂 Folder Structure

invoice-converter/
├── convert_all.py # Main script
├── config.json # Configuration file
├── invoices/ # Folder for PDF invoices
├── requirements.txt # Python dependencies
└── README.md # This file

yaml
Copy
Edit

---

## 🚀 Features

- ✅ Process all PDF invoices in a folder
- ✅ Extract invoice metadata (dates, buyer, VAT, totals)
- ✅ Extract multiple product lines per invoice
- ✅ Merge invoice-level fields across rows in Excel
- ✅ Add dynamic formulas (`=Q15=K15`)
- ✅ Optionally skip free gift items

---

## ⚙️ Installation

### 1️⃣ Install Python Dependencies
```bash
pip install -r requirements.txt
requirements.txt

nginx
Copy
Edit
pdfplumber
pandas
openpyxl
2️⃣ Configure the Script
Edit config.json:

json
Copy
Edit
{
    "PDF_FOLDER": "invoices",
    "OUTPUT_FILE": "combined_invoices.xlsx",
    "ORDER_DATE_DEFAULT": "01-07-2025",
    "SKIP_FREE_GIFTS": true
}
3️⃣ Add Your Invoices
Place all .pdf files in the invoices/ folder.

4️⃣ Run the Script
bash
Copy
Edit
python convert_all.py
Result: combined_invoices.xlsx will be created in the project folder.

📦 Portable Executable
You can package the script into a standalone .exe for Windows or .app for macOS.

Windows
bash
Copy
Edit
pip install pyinstaller
pyinstaller --onefile --add-data "config.json;." --add-data "invoices;invoices" convert_all.py
macOS/Linux
bash
Copy
Edit
pip install pyinstaller
pyinstaller --onefile --add-data "config.json:." --add-data "invoices:invoices" convert_all.py
The executable will be in the dist/ folder.

Run from command line:

bash
Copy
Edit
dist/convert_all.exe
