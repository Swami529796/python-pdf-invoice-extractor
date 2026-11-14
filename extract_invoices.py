"""PDF Invoice Extractor.

Reads every PDF invoice in a folder, pulls out the header fields (vendor, invoice
number, date, subtotal, tax, total), checks them, and writes one Excel workbook:

    Invoices   one row per PDF, with a Status column
    Issues     every problem found, with the file name and the reason

Vendors word their invoices differently ("Invoice No", "Invoice #", "Bill Number";
"Grand Total", "Amount Due", "Total Payable"), so each field has a list of patterns
that are tried in turn. Supporting a new vendor usually means adding one pattern.

Usage:
    python extract_invoices.py sample-invoices
    python extract_invoices.py "C:/Invoices/May" --output May_Invoices.xlsx

Author : Swaminathan Narayanan
Repo   : https://github.com/Swami529796/python-pdf-invoice-extractor
"""
from __future__ import annotations

import argparse
import re
import sys
import time
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import pdfplumber
from openpyxl.styles import Font, PatternFill

AMOUNT = r"(?:Rs\.?|INR|\$)?\s*(-?[\d,]+\.\d{2})"
TOLERANCE = 1.00   # rupees: subtotal + tax may differ from the total by this much (rounding)

# field -> patterns tried in order; the first capture group is the value
FIELD_PATTERNS: dict[str, list[str]] = {
    "invoice_number": [
        r"Invoice\s*(?:No\.?|Number|#)\s*:?\s*([A-Z0-9][A-Z0-9/\-]+)",
        r"Bill\s*(?:No\.?|Number)\s*:?\s*([A-Z0-9][A-Z0-9/\-]+)",
    ],
    "invoice_date": [
        r"(?:Invoice\s*Date|Dated|Date)\s*:?\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})",
        r"(?:Invoice\s*Date|Dated|Date)\s*:?\s*(\d{4}-\d{2}-\d{2})",
        r"(?:Invoice\s*Date|Dated|Date)\s*:?\s*([A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})",
        r"(?:Invoice\s*Date|Dated|Date)\s*:?\s*(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4})",
    ],
    "customer": [r"(?:Bill(?:ed)?\s*To|Customer)\s*:?\s*(.+)"],
    "subtotal": [rf"(?:Sub\s*-?\s*Total|Net\s*Amount)\s*:?\s*{AMOUNT}"],
    "tax": [rf"(?:I?GST|Tax)[^\n\d]*(?:\d+\s*%\s*\)?)?\s*:?\s*{AMOUNT}"],
    "total": [rf"(?:Grand\s*Total|Amount\s*Due|Total\s*Payable|Total\s*Amount)\s*:?\s*{AMOUNT}"],
}
DATE_FORMATS = ["%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%Y-%m-%d", "%B %d, %Y", "%B %d %Y", "%b %d, %Y",
                "%d %B %Y", "%d %b %Y"]


@dataclass
class Invoice:
    file: str
    vendor: str = ""
    invoice_number: str = ""
    invoice_date: date | None = None
    customer: str = ""
    subtotal: float | None = None
    tax: float | None = None
    total: float | None = None
    status: str = "OK"
    issues: list[str] = field(default_factory=list)


def find(text: str, field_name: str) -> str:
    """First match for a field, or an empty string."""
    for pattern in FIELD_PATTERNS[field_name]:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return ""


def to_amount(text: str) -> float | None:
    return float(text.replace(",", "")) if text else None


def to_date(text: str) -> date | None:
    for date_format in DATE_FORMATS:
        try:
            return datetime.strptime(text.strip(), date_format).date()
        except ValueError:
            continue
    return None


def parse_invoice(text: str, file_name: str) -> Invoice:
    """Turn the text of one PDF into an Invoice and record anything that looks wrong."""
    invoice = Invoice(file=file_name)
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        invoice.issues.append("No readable text (a scanned image? OCR is needed)")
    else:
        invoice.vendor = lines[0].title() if lines[0].isupper() else lines[0]
        invoice.invoice_number = find(text, "invoice_number")
        invoice.customer = find(text, "customer")
        invoice.subtotal = to_amount(find(text, "subtotal"))
        invoice.tax = to_amount(find(text, "tax"))
        invoice.total = to_amount(find(text, "total"))
        date_text = find(text, "invoice_date")
        invoice.invoice_date = to_date(date_text)

        if not invoice.invoice_number and invoice.total is None:
            invoice.issues.append("Does not look like an invoice (no invoice number or total found)")
        else:
            if not invoice.invoice_number:
                invoice.issues.append("Invoice number not found")
            if invoice.invoice_date is None:
                invoice.issues.append(f"Invoice date not readable: '{date_text}'" if date_text else "Invoice date not found")
            if invoice.total is None:
                invoice.issues.append("Total not found")
            elif invoice.subtotal is not None and invoice.tax is not None:
                expected = invoice.subtotal + invoice.tax
                if abs(expected - invoice.total) > TOLERANCE:
                    invoice.issues.append(f"Total {invoice.total:,.2f} does not equal subtotal + tax {expected:,.2f}")
    if invoice.issues:
        invoice.status = "Check"
    return invoice


def read_pdf_text(path: Path) -> str:
    with pdfplumber.open(path) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)


def extract_folder(folder: Path) -> list[Invoice]:
    invoices: list[Invoice] = []
    seen: dict[tuple[str, str], str] = {}
    for path in sorted(folder.glob("*.pdf"), key=lambda p: p.name.lower()):
        try:
            invoice = parse_invoice(read_pdf_text(path), path.name)
        except Exception as error:   # a damaged PDF must not stop the run
            invoice = Invoice(file=path.name, status="Check", issues=[f"Could not read PDF: {error}"])
        key = (invoice.vendor.lower(), invoice.invoice_number.lower())
        if invoice.invoice_number:
            if key in seen:
                invoice.status = "Check"
                invoice.issues.append(f"Duplicate of {seen[key]} (same vendor and invoice number)")
            else:
                seen[key] = invoice.file
        invoices.append(invoice)
    return invoices


def write_workbook(invoices: list[Invoice], output: Path) -> None:
    rows = []
    for invoice in invoices:
        row = asdict(invoice)
        row["issues"] = "; ".join(invoice.issues)
        rows.append(row)
    table = pd.DataFrame(rows).rename(columns=lambda name: name.replace("_", " ").title())
    issues = pd.DataFrame([{"File": invoice.file, "Issue": issue} for invoice in invoices for issue in invoice.issues],
                          columns=["File", "Issue"])

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        table.to_excel(writer, sheet_name="Invoices", index=False)
        issues.to_excel(writer, sheet_name="Issues", index=False)
        for sheet in writer.sheets.values():
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", start_color="1F4E78")
            sheet.freeze_panes = "A2"
            for column in sheet.columns:
                longest = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column)
                sheet.column_dimensions[column[0].column_letter].width = min(longest + 2, 70)
        sheet = writer.sheets["Invoices"]
        headers = [cell.value for cell in sheet[1]]
        for row in sheet.iter_rows(min_row=2):
            status = row[headers.index("Status")]
            status.fill = PatternFill("solid", start_color="C6EFCE" if status.value == "OK" else "FFEB9C")
            row[headers.index("Invoice Date")].number_format = "dd-mmm-yyyy"
            for name in ("Subtotal", "Tax", "Total"):
                row[headers.index(name)].number_format = "#,##0.00"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Extract invoice data from every PDF in a folder.")
    parser.add_argument("folder", type=Path)
    parser.add_argument("-o", "--output", type=Path, default=Path("Invoices.xlsx"))
    args = parser.parse_args(argv)
    if not args.folder.is_dir():
        print(f"Folder not found: {args.folder}", file=sys.stderr)
        return 1

    started = time.perf_counter()
    invoices = extract_folder(args.folder)
    if not invoices:
        print(f"No PDF files found in: {args.folder}", file=sys.stderr)
        return 1
    write_workbook(invoices, args.output)

    ok = [invoice for invoice in invoices if invoice.status == "OK"]
    print(f"PDFs read:      {len(invoices)}")
    print(f"Extracted OK:   {len(ok)}   (total value {sum(invoice.total or 0 for invoice in ok):,.2f})")
    print(f"Need checking:  {len(invoices) - len(ok)}")
    for invoice in invoices:
        for issue in invoice.issues:
            print(f"  ! {invoice.file}: {issue}")
    print(f"Time taken:     {time.perf_counter() - started:.1f} sec")
    print(f"Output:         {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
