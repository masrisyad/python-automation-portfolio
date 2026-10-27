# Web Scraper to Excel

Scrapes a 1,000-product online catalogue (pagination, retries, rate limiting) and exports clean data to CSV and a formatted Excel workbook.

**Client use cases:** competitor price monitoring, product catalogue migration, lead list building, market research datasets.

## What it does
- Walks every page of the catalogue until the site reports no more pages
- Retries failed requests with exponential backoff, waits politely between pages
- Converts raw text into typed data: price as a number, rating as 1-5, stock as Yes/No
- Writes `output/books.csv` and `output/books.xlsx` (bold header, filters, frozen header row, currency format)

## Run
```bash
pip install -r requirements.txt
python scraper.py --max-pages 3     # quick test
python scraper.py                   # full catalogue
```

## Sample output
See `output/books.xlsx` (included in this repo).

| Title | Price (GBP) | Rating (1-5) | In Stock | URL |
|---|---|---|---|---|
| A Light in the Attic | 51.77 | 3 | Yes | https://books.toscrape.com/catalogue/... |

## Adapting to your site
The site-specific parts live in one function (`parse_page`). Point `BASE_URL` at your target and adjust the CSS selectors; everything else (pagination loop, retries, export) stays the same.

Target site: https://books.toscrape.com, a public sandbox built for scraping practice.
