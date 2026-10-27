# Excel Report Automation

Turns a folder of messy monthly CSV exports into one management-ready Excel report: cleaned data, KPIs, pivot tables, native Excel charts and a heat-map. One command replaces hours of manual copy-paste.

**Client use cases:** monthly sales/finance reporting, consolidating exports from Shopify, QuickBooks, POS or CRM systems, recurring KPI dashboards for managers.

## What it handles
Real exports are never clean. The script fixes:
- Duplicate rows
- Inconsistent casing and whitespace in names, regions, products (`" SOUTH "` becomes `South`)
- Mixed date formats (`2026-01-05`, `01/05/2026`, `Jan 5, 2026`)
- Missing or non-numeric quantities and prices (removed and counted)

Then it builds a workbook with these sheets:

| Sheet | Content |
|---|---|
| Summary | KPIs (revenue, orders, AOV, best month, top region/product) and a data-quality log |
| Monthly | Orders, units, revenue, month-over-month growth, line chart |
| By Region | Revenue share per region, bar chart |
| By Product | Units, revenue, average price, bar chart |
| Top Customers | Top 10 by revenue |
| Region x Month | Pivot with colour-scale heat-map |
| Clean Data | The full cleaned dataset for further analysis |

## Run
```bash
pip install -r requirements.txt
python generate_sample_data.py     # creates 6 messy sample files in ./input
python build_report.py             # writes output/sales_report.xlsx
```

## Sample output
`output/sales_report.xlsx` is included in this repo.

## Adapting to your data
Column names are mapped in one place (`clean()`), so pointing this at your own exports is usually a 15-minute change. The script can also be scheduled (Windows Task Scheduler, cron, or GitHub Actions) to run automatically each month and email the result.
