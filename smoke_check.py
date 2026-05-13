from __future__ import annotations

import unittest
from io import BytesIO

from openpyxl import Workbook

from msg_to_excel import (
    ExtractedTable,
    canonicalize_table,
    extract_rows_from_excel_bytes,
    extract_rows_from_pdf_bytes,
)


class ExtractionSmokeTests(unittest.TestCase):
    def test_canonicalize_html_like_headers(self) -> None:
        table = ExtractedTable(
            headers=["Invoice Number", "Invoice Date", "Open Balance"],
            rows=[["32855", "1/16/2026", "$211.00"]],
        )
        rows = canonicalize_table(table)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].invoice_number, "32855")
        self.assertEqual(rows[0].invoice_date, "2026-01-16")
        self.assertEqual(rows[0].outstanding_amount, "211.00")

    def test_extract_rows_from_excel_bytes(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Vendor Statement"])
        sheet.append([])
        sheet.append(["Invoice Number", "Invoice Date", "Open Balance"])
        sheet.append(["32855", "1/16/2026", "$211.00"])
        sheet.append(["32920", "1/19/2026", "$597.30"])

        buffer = BytesIO()
        workbook.save(buffer)
        workbook.close()

        rows, detail = extract_rows_from_excel_bytes(buffer.getvalue())
        self.assertEqual(len(rows), 2)
        self.assertIn("worksheet", detail)

    def test_extract_rows_from_pdf_bytes(self) -> None:
        import fitz

        doc = fitz.open()
        page = doc.new_page()
        text = (
            "Invoice# Date Open Balance\n"
            "7349700 02/16/2026 BOL 37888 $340.00 13 $340.00\n"
            "7348900 02/16/2026 PRO 380350862 $575.00 13 $575.00\n"
        )
        page.insert_text((72, 72), text, fontsize=11)
        rows, detail = extract_rows_from_pdf_bytes(doc.tobytes())

        self.assertEqual(len(rows), 2)
        self.assertIn("PDF", detail)


if __name__ == "__main__":
    unittest.main()
