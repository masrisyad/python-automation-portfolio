# PDF Invoice Data Extractor

Extracts header fields and line items from a folder of invoice PDFs into a structured Excel workbook, and flags invoices whose numbers do not add up for manual review.

**Client use cases:** bulk entry of supplier invoices into accounting software, accounts-payable audits, migrating years of PDF invoices into a database, expense reporting.

## What it extracts
| Sheet | Columns |
|---|---|
| Invoices | vendor, invoice number, invoice date, due date, bill-to, subtotal, tax, total, line-item count, status |
| Line Items | invoice number, description, quantity, unit price, amount |
| Issues | anything that needs a human: missing fields, line items that do not sum to the subtotal, subtotal + tax not equal to total |

Validation is built in, so the client gets a short review list instead of checking every row.

## Run
```bash
pip install -r requirements.txt
python generate_sample_invoices.py     # creates 12 sample PDFs in ./invoices
python extract_invoices.py             # writes output/invoices.xlsx
```

## Sample output
`output/invoices.xlsx` is included in this repo.

## Adapting to your invoices
Field patterns are listed in one dictionary (`PATTERNS`). For a new supplier layout, add or adjust a regex; the table parser, validation and Excel export stay the same. Scanned (image-only) PDFs can be supported by adding an OCR step with Tesseract.
