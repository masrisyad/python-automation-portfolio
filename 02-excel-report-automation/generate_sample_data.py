"""
Generates 6 months of intentionally messy sales CSV exports into ./input.
Simulates what real clients send: duplicate rows, inconsistent casing and
whitespace, mixed date formats, missing quantities.

    python generate_sample_data.py
"""

import csv
import random
from datetime import date, timedelta
from pathlib import Path

random.seed(42)

PRODUCTS = [("Wireless Mouse", 25.0), ("USB-C Hub", 45.0), ("Laptop Stand", 39.0),
            ("Mechanical Keyboard", 89.0), ("Webcam 1080p", 59.0), ("Desk Lamp", 32.0),
            ("Monitor 27in", 229.0), ("Headset", 69.0)]
REGIONS = ["North", "South", "East", "West"]
CUSTOMERS = [f"Customer {chr(65 + i % 26)}{i:02d}" for i in range(40)]
MESSY_REGION = [str.upper, str.lower, lambda s: f" {s} ", lambda s: s]
DATE_FORMATS = ["%Y-%m-%d", "%m/%d/%Y", "%b %d, %Y"]


def month_rows(year: int, month: int, start_id: int) -> list[dict]:
    rows = []
    first = date(year, month, 1)
    days = 31 if month in (1, 3, 5, 7, 8, 10, 12) else 30 if month != 2 else 28
    for i in range(random.randint(260, 340)):
        product, price = random.choice(PRODUCTS)
        d = first + timedelta(days=random.randint(0, days - 1))
        qty = random.randint(1, 12)
        rows.append({
            "order_id": f"ORD-{start_id + i:06d}",
            "order_date": d.strftime(random.choice(DATE_FORMATS)),
            "customer": random.choice(CUSTOMERS),
            "region": random.choice(MESSY_REGION)(random.choice(REGIONS)),
            "product": product if random.random() > 0.1 else product.lower(),
            "quantity": "" if random.random() < 0.02 else qty,
            "unit_price": f"{price * random.uniform(0.9, 1.1):.2f}",
        })
    # inject ~3% exact duplicates
    for _ in range(len(rows) // 33):
        rows.append(dict(random.choice(rows)))
    random.shuffle(rows)
    return rows


def main() -> None:
    out = Path("input")
    out.mkdir(exist_ok=True)
    next_id = 1
    for month in range(1, 7):
        rows = month_rows(2026, month, next_id)
        next_id += len(rows)
        path = out / f"sales_2026_{month:02d}.csv"
        with path.open("w", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)
        print(f"wrote {path} ({len(rows)} rows)")


if __name__ == "__main__":
    main()
