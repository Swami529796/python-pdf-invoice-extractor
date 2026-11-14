"""Create sample PDF invoices from three made-up vendors, each with its own layout.

Repeatable (fixed seed). Two problem files are planted on purpose:
  - an invoice whose total does not equal subtotal + tax
  - a delivery note that is not an invoice at all
One invoice number also appears twice (the same invoice sent again).

    python make_sample_invoices.py
"""
from __future__ import annotations

import random
from datetime import date, timedelta
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

OUTPUT = Path(__file__).parent / "sample-invoices"
CUSTOMER = "Northgate Retail Pvt Ltd"
ITEMS = [("A4 Copier Paper (box)", 1850), ("Toner Cartridge", 4200), ("Wireless Mouse", 850),
         ("Freight - Chennai to Pune", 12500), ("Warehouse Handling", 3600), ("Annual Support Plan", 48000),
         ("Laptop Repair", 6500), ("Network Cabling (per point)", 1200), ("Office Chair", 7400)]


def money(value: float) -> str:
    return f"{value:,.2f}"


def draw_items(pdf: canvas.Canvas, top: float, lines: list[tuple[str, int, float]]) -> float:
    pdf.setFont("Helvetica-Bold", 10)
    pdf.drawString(25 * mm, top, "Description")
    pdf.drawRightString(125 * mm, top, "Qty")
    pdf.drawRightString(155 * mm, top, "Rate")
    pdf.drawRightString(185 * mm, top, "Amount")
    pdf.line(25 * mm, top - 2 * mm, 185 * mm, top - 2 * mm)
    pdf.setFont("Helvetica", 10)
    y = top - 8 * mm
    for name, quantity, rate in lines:
        pdf.drawString(25 * mm, y, name)
        pdf.drawRightString(125 * mm, y, str(quantity))
        pdf.drawRightString(155 * mm, y, money(rate))
        pdf.drawRightString(185 * mm, y, money(quantity * rate))
        y -= 7 * mm
    pdf.line(25 * mm, y + 3 * mm, 185 * mm, y + 3 * mm)
    return y - 4 * mm


def totals_block(pdf: canvas.Canvas, y: float, rows: list[tuple[str, str]]) -> None:
    for index, (label, value) in enumerate(rows):
        pdf.setFont("Helvetica-Bold" if index == len(rows) - 1 else "Helvetica", 10)
        pdf.drawString(120 * mm, y, label)
        pdf.drawRightString(185 * mm, y, value)
        y -= 7 * mm


def layout_apex(pdf, number, day, lines, subtotal, tax, total):
    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawString(25 * mm, 275 * mm, "Apex Office Supplies")
    pdf.setFont("Helvetica", 9)
    pdf.drawString(25 * mm, 269 * mm, "14 Mount Road, Chennai 600002  |  GSTIN: 33AAACA1234F1Z5")
    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawRightString(185 * mm, 275 * mm, "TAX INVOICE")
    pdf.setFont("Helvetica", 10)
    pdf.drawString(25 * mm, 255 * mm, f"Invoice No: {number}")
    pdf.drawString(25 * mm, 249 * mm, f"Invoice Date: {day:%d/%m/%Y}")
    pdf.drawString(25 * mm, 243 * mm, f"Bill To: {CUSTOMER}")
    y = draw_items(pdf, 228 * mm, lines)
    totals_block(pdf, y, [("Subtotal", money(subtotal)), ("GST (18%)", money(tax)),
                          ("Grand Total", f"Rs {money(total)}")])


def layout_bluewave(pdf, number, day, lines, subtotal, tax, total):
    pdf.setFont("Helvetica-Bold", 20)
    pdf.drawString(25 * mm, 272 * mm, "BLUEWAVE LOGISTICS")
    pdf.setFont("Helvetica", 9)
    pdf.drawString(25 * mm, 266 * mm, "Plot 7, MIDC Chakan, Pune 410501")
    pdf.setFont("Helvetica-Bold", 14)
    pdf.drawString(25 * mm, 252 * mm, "INVOICE")
    pdf.setFont("Helvetica", 10)
    pdf.drawString(120 * mm, 252 * mm, f"Invoice # {number}")
    pdf.drawString(120 * mm, 246 * mm, f"Date {day:%B} {day.day}, {day.year}")
    pdf.drawString(25 * mm, 240 * mm, f"Customer: {CUSTOMER}")
    y = draw_items(pdf, 226 * mm, lines)
    totals_block(pdf, y, [("Sub Total", money(subtotal)), ("Tax @ 18%", money(tax)),
                          ("Amount Due", f"INR {money(total)}")])
    pdf.setFont("Helvetica-Oblique", 8)
    pdf.drawString(25 * mm, 20 * mm, "Payment due within 30 days. Thank you for your business.")


def layout_crestline(pdf, number, day, lines, subtotal, tax, total):
    pdf.setFillColorRGB(0.12, 0.31, 0.47)
    pdf.rect(0, 277 * mm, 210 * mm, 20 * mm, stroke=0, fill=1)
    pdf.setFillColorRGB(1, 1, 1)
    pdf.setFont("Helvetica-Bold", 15)
    pdf.drawString(25 * mm, 284 * mm, "Crestline IT Services LLP")
    pdf.setFillColorRGB(0, 0, 0)
    pdf.setFont("Helvetica", 10)
    pdf.drawString(25 * mm, 265 * mm, f"Billed To: {CUSTOMER}")
    pdf.drawString(25 * mm, 259 * mm, f"Bill Number: {number}")
    pdf.drawString(25 * mm, 253 * mm, f"Dated: {day:%Y-%m-%d}")
    y = draw_items(pdf, 238 * mm, lines)
    totals_block(pdf, y, [("Net Amount", money(subtotal)), ("IGST @ 18%", money(tax)),
                          ("Total Payable", f"Rs. {money(total)}")])


def delivery_note(path: Path) -> None:
    pdf = canvas.Canvas(str(path), pagesize=A4)
    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawString(25 * mm, 275 * mm, "Apex Office Supplies")
    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawString(25 * mm, 260 * mm, "DELIVERY NOTE")
    pdf.setFont("Helvetica", 10)
    pdf.drawString(25 * mm, 250 * mm, "Delivered 12 boxes of A4 Copier Paper to the reception desk.")
    pdf.drawString(25 * mm, 244 * mm, "Received by: ____________________")
    pdf.save()


def main() -> None:
    rng = random.Random(5)
    OUTPUT.mkdir(exist_ok=True)
    for old in OUTPUT.glob("*.pdf"):
        old.unlink()

    vendors = [("apex", layout_apex, "AOS-2026-{:04d}", 141), ("bluewave", layout_bluewave, "BW/{}", 1057),
               ("crestline", layout_crestline, "CR-{}", 88213)]
    count = 0
    for name, layout, pattern, first_number in vendors:
        for offset in range(4):
            number = pattern.format(first_number + offset)
            day = date(2026, 5, 4) + timedelta(days=rng.randint(0, 50))
            lines = [(item, rng.randint(1, 12), float(rate)) for item, rate in rng.sample(ITEMS, rng.randint(2, 4))]
            subtotal = sum(quantity * rate for _, quantity, rate in lines)
            tax = round(subtotal * 0.18, 2)
            total = subtotal + tax
            if name == "bluewave" and offset == 2:
                total += 1000            # planted: total does not add up
            path = OUTPUT / f"{name}_{number.replace('/', '-')}.pdf"
            pdf = canvas.Canvas(str(path), pagesize=A4)
            layout(pdf, number, day, lines, subtotal, tax, total)
            pdf.save()
            count += 1
            if name == "crestline" and offset == 1:
                copy = OUTPUT / f"{name}_{number}_resent.pdf"   # planted: the same invoice sent twice
                copy.write_bytes(path.read_bytes())
                count += 1
    delivery_note(OUTPUT / "apex_delivery_note.pdf")
    print(f"Wrote {count + 1} PDF files to {OUTPUT}")


if __name__ == "__main__":
    main()
