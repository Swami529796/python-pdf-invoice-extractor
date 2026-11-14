"""Tests for extract_invoices.py. Run with: python -m pytest"""
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import extract_invoices as ex  # noqa: E402

SAMPLES = Path(__file__).resolve().parents[1] / "sample-invoices"

APEX = """Apex Office Supplies
TAX INVOICE
Invoice No: AOS-2026-0141
Invoice Date: 14/05/2026
Bill To: Northgate Retail Pvt Ltd
Subtotal 12,500.00
GST (18%) 2,250.00
Grand Total Rs 14,750.00"""


def test_fields_are_read_from_one_layout():
    invoice = ex.parse_invoice(APEX, "a.pdf")
    assert invoice.vendor == "Apex Office Supplies"
    assert invoice.invoice_number == "AOS-2026-0141"
    assert invoice.invoice_date == date(2026, 5, 14)
    assert invoice.customer == "Northgate Retail Pvt Ltd"
    assert (invoice.subtotal, invoice.tax, invoice.total) == (12500.0, 2250.0, 14750.0)
    assert invoice.status == "OK"


def test_other_wordings_are_understood():
    text = "BLUEWAVE LOGISTICS\nInvoice # BW/1057\nDate May 14, 2026\nSub Total 100.00\nTax @ 18% 18.00\nAmount Due INR 118.00"
    invoice = ex.parse_invoice(text, "b.pdf")
    assert invoice.vendor == "Bluewave Logistics"
    assert invoice.invoice_number == "BW/1057"
    assert invoice.invoice_date == date(2026, 5, 14)
    assert invoice.total == 118.0 and invoice.status == "OK"

    text = "Crestline IT\nBill Number: CR-88213\nDated: 2026-05-20\nNet Amount 200.00\nIGST @ 18% 36.00\nTotal Payable Rs. 236.00"
    invoice = ex.parse_invoice(text, "c.pdf")
    assert invoice.invoice_number == "CR-88213" and invoice.tax == 36.0 and invoice.status == "OK"


def test_a_total_that_does_not_add_up_is_flagged():
    invoice = ex.parse_invoice(APEX.replace("14,750.00", "15,750.00"), "a.pdf")
    assert invoice.status == "Check"
    assert "does not equal subtotal + tax" in invoice.issues[0]


def test_rounding_differences_are_accepted():
    assert ex.parse_invoice(APEX.replace("14,750.00", "14,750.50"), "a.pdf").status == "OK"


def test_a_non_invoice_is_flagged():
    invoice = ex.parse_invoice("Apex Office Supplies\nDELIVERY NOTE\nDelivered 12 boxes.", "note.pdf")
    assert invoice.status == "Check" and "not look like an invoice" in invoice.issues[0]


def test_empty_text_suggests_ocr():
    assert "OCR" in ex.parse_invoice("", "scan.pdf").issues[0]


def test_sample_folder_end_to_end(tmp_path):
    invoices = ex.extract_folder(SAMPLES)
    assert len(invoices) == 14
    assert sum(invoice.status == "OK" for invoice in invoices) == 11
    problems = {invoice.file: invoice.issues[0] for invoice in invoices if invoice.status == "Check"}
    assert "Duplicate of crestline_CR-88214.pdf" in problems["crestline_CR-88214_resent.pdf"]
    assert "does not equal" in problems["bluewave_BW-1059.pdf"]

    output = tmp_path / "out.xlsx"
    ex.write_workbook(invoices, output)
    assert output.exists()
