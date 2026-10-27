# Python Automation Portfolio

Three small, production-style automation tools that solve the kind of problems businesses post every day: "get this data out of that website/file and into a clean spreadsheet". Each one runs with a single command, includes sample input and sample output, and is written to be adapted to a client's own data in minutes.

| Project | Problem it solves | Stack |
|---|---|---|
| [01 Web Scraper to Excel](01-web-scraper-to-excel/) | Pull a 1,000-item catalogue from a website into a clean CSV/Excel file, with pagination, retries and rate limiting | requests, BeautifulSoup, openpyxl |
| [02 Excel Report Automation](02-excel-report-automation/) | Merge messy monthly CSV exports into one Excel report with cleaned data, KPIs, pivots, charts and a heat-map | pandas, openpyxl |
| [03 PDF Invoice Extractor](03-pdf-invoice-extractor/) | Extract header fields and line items from invoice PDFs into Excel, with automatic validation of totals | pdfplumber, pandas, openpyxl |

## How I work
- Clear written communication, async-first. You get a short daily written update and a screen recording of the final result.
- Every delivery includes the source code, a README, and a sample run so you can re-run it yourself.
- Scripts are built to be scheduled (Task Scheduler, cron, GitHub Actions) or extended later.

## Running any project
```bash
cd <project folder>
pip install -r requirements.txt
python <script>.py --help
```
Python 3.10 or newer.
