"""
Generates realistic sample invoice PDFs into ./invoices so the extractor can
be demonstrated without any client data.

    python generate_sample_invoices.py --count 12
"""

import argparse
import random
from datetime import date, timedelta
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

random.seed(7)

VENDORS = ["Northwind Supplies Ltd", "Acme Office Solutions", "BlueSky Logistics Inc",
           "Greenfield Hardware Co", "Summit Cloud Services"]
CLIENTS = ["Harbor Retail Group", "Pinecrest Clinic", "Lumen Studio", "Riverside Cafe"]
ITEMS = [("Printer paper A4 (box)", 24.50), ("Toner cartridge black", 89.00),
         ("Ergonomic chair", 215.00), ("Monthly hosting plan", 49.00),
         ("Delivery service", 35.00), ("Consulting hour", 120.00),
         ("LED desk lamp", 32.90), ("Network cable 5m", 7.80)]


def build_invoice(path: Path, n: int) -> None:
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(str(path), pagesize=A4, leftMargin=20 * mm, rightMargin=20 * mm, topMargin=20 * mm)
    vendor = random.choice(VENDORS)
    client = random.choice(CLIENTS)
    inv_no = f"INV-2026-{n:04d}"
    inv_date = date(2026, random.randint(1, 9), random.randint(1, 28))
    due = inv_date + timedelta(days=random.choice([14, 30, 45]))
    tax_rate = random.choice([0.0, 0.08, 0.10])

    rows = [["Description", "Qty", "Unit Price", "Amount"]]
    subtotal = 0.0
    for desc, price in random.sample(ITEMS, random.randint(2, 5)):
        qty = random.randint(1, 10)
        amount = round(qty * price, 2)
        subtotal += amount
        rows.append([desc, str(qty), f"${price:,.2f}", f"${amount:,.2f}"])
    subtotal = round(subtotal, 2)
    tax = round(subtotal * tax_rate, 2)
    total = round(subtotal + tax, 2)

    story = [
        Paragraph(f"<b>{vendor}</b>", styles["Title"]),
        Paragraph("123 Commerce Street, Springfield", styles["Normal"]),
        Spacer(1, 8 * mm),
        Paragraph("<b>INVOICE</b>", styles["Heading2"]),
        Paragraph(f"Invoice Number: {inv_no}", styles["Normal"]),
        Paragraph(f"Invoice Date: {inv_date:%Y-%m-%d}", styles["Normal"]),
        Paragraph(f"Due Date: {due:%Y-%m-%d}", styles["Normal"]),
        Spacer(1, 4 * mm),
        Paragraph(f"Bill To: {client}", styles["Normal"]),
        Spacer(1, 8 * mm),
    ]
    table = Table(rows, colWidths=[90 * mm, 20 * mm, 30 * mm, 30 * mm])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
    ]))
    story += [
        table,
        Spacer(1, 6 * mm),
        Paragraph(f"Subtotal: ${subtotal:,.2f}", styles["Normal"]),
        Paragraph(f"Tax ({tax_rate * 100:.0f}%): ${tax:,.2f}", styles["Normal"]),
        Paragraph(f"<b>Total Due: ${total:,.2f}</b>", styles["Normal"]),
        Spacer(1, 8 * mm),
        Paragraph("Payment terms: bank transfer. Thank you for your business.", styles["Normal"]),
    ]
    doc.build(story)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--count", type=int, default=12)
    args = parser.parse_args()
    out = Path("invoices")
    out.mkdir(exist_ok=True)
    for n in range(1, args.count + 1):
        p = out / f"invoice_{n:04d}.pdf"
        build_invoice(p, n)
        print("wrote", p)


if __name__ == "__main__":
    main()
