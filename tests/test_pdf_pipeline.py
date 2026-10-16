from __future__ import annotations

import unittest

import fitz

from formatter.formatter_engine import FormattingOptions
from formatter.pdf_pipeline import process_uploaded_pdf


def _make_text_pdf() -> bytes:
    pdf = fitz.open()
    first = pdf.new_page()
    first.insert_text(
        (72, 72),
        "Results\nFig. 2. Current response\nAs shown in Fig. 2, the output is stable.",
        fontsize=12,
    )
    second = pdf.new_page()
    second.insert_text((72, 72), "Table I\nMEASURED RESULTS\nSee Table I for details.", fontsize=12)
    payload = pdf.tobytes()
    pdf.close()
    return payload


class PdfPipelineTests(unittest.TestCase):
    def test_pdf_analysis_extracts_text_objects_and_page_locations_without_formatting(self):
        run = process_uploaded_pdf(
            report_bytes=_make_text_pdf(),
            input_file_name="sample.pdf",
            options=FormattingOptions(ui_language="en"),
        )

        self.assertEqual(run.page_count, 2)
        self.assertEqual(run.input_file_name, "sample.pdf")
        self.assertFalse(run.result.actions)
        self.assertTrue(run.result.figures)
        self.assertEqual(run.result.figures[0].pdf_page_number, 1)
        self.assertIn("PDF analysis only", run.result.report_text)
        self.assertIn("Page 2", run.result.report_text)
