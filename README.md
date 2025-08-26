# Invoice2Excel/BillBot 🧾➡️📊

Invoice2Excel is a simple MVP that converts PDF invoices into structured Excel sheets using OCR (Tesseract + Poppler + Python).  
It extracts **Invoice No, Date, Buyer, Services, Quantity, Rate, and Total Amount** and exports them neatly into Excel.

## 🚀 Features
- Drag & Drop PDF invoices
- Automatic OCR text extraction
- Structured Excel output with multiple sheets (Summary + Line Items + Raw OCR text)
- Works with scanned PDFs and images

## 📂 Output Example
| Invoice No | Invoice Date | S.No | Description of Services | Quantity | Rate | Total Amount |
|------------|--------------|------|-------------------------|----------|------|--------------|
| H/AMC/2425/0289 | 27-Feb-25 | 1 | AMC Services PCs, Printers, Laptops & Network (01-11-2024–31-01-2025) | 1.00 | 406450.00 | 406450.00 |

## 🛠️ Installation
```bash
git clone https://github.com/yourusername/Invoice2Excel.git
cd Invoice2Excel
pip install -r requirements.txt
