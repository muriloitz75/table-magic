import unittest
from pathlib import Path

from backend.extraction.report_parser import OUTPUT_HEADERS, parse_report

MODEL_PDF = Path(__file__).resolve().parent.parent.parent / "modelo" / "Setor21.pdf"


class TestReportParser(unittest.TestCase):
    def test_parse_report_modelo_setor21(self):
        if not MODEL_PDF.exists():
            self.skipTest("modelo/Setor21.pdf not present")
        pdf_bytes = MODEL_PDF.read_bytes()
        result = parse_report(pdf_bytes)

        self.assertIsNotNone(result)
        self.assertEqual(result["headers"], OUTPUT_HEADERS)
        self.assertEqual(len(result["rows"]), 1917)
        self.assertTrue(result["count_only"])
        self.assertEqual(result["rows"][0][0], "001216 - 2025")
        self.assertEqual(result["rows"][0][2], "22/09/2025")
        self.assertEqual(result["rows"][0][3], "set")
        self.assertEqual(result["rows"][0][4], "2025")


if __name__ == "__main__":
    unittest.main()

