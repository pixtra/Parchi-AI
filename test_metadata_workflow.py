import io
import json
import unittest

from openpyxl import load_workbook

import app
from metadata_workflow import exact_template_label, suggestions


RECORDS = [
    {"document_name": "invoice.pdf", "label": "Invoice Number", "value": "INV-1", "confidence_score": 91},
    {"document_name": "invoice.pdf", "label": "Invoice Number", "value": "INV-2", "confidence_score": 73},
    {"document_name": "po.pdf", "label": "Supplier Name", "value": "ABC", "confidence_score": 88},
]


class MetadataWorkflowTests(unittest.TestCase):
    def test_exact_matching_ignores_case_and_outer_whitespace(self):
        self.assertEqual(
            exact_template_label("  invoice number ", ["Invoice Number"]),
            "Invoice Number",
        )

    def test_fuzzy_suggestions_require_rapidfuzz_threshold(self):
        self.assertTrue(suggestions("Invoice No", ["Invoice Number"], min_score=70))
        self.assertEqual(suggestions("Unrelated", ["Invoice Number"], min_score=70), [])

    def test_exports_preserve_duplicates_and_four_fields(self):
        expected_headers = ["Document Name", "Label", "Value", "Confidence Score"]
        builders = (
            lambda: app.build_metadata_export(RECORDS, "Main Metadata"),
            lambda: app.build_metadata_psv(RECORDS, "Main Metadata"),
            lambda: app.build_metadata_xlsx(RECORDS, "Main Metadata"),
            lambda: app.build_metadata_export(RECORDS, "Client Template"),
            lambda: app.build_metadata_psv(RECORDS, "Client Template"),
            lambda: app.build_metadata_xlsx(RECORDS, "Client Template"),
        )
        for build in builders:
            data, count = build()
            self.assertEqual(count, len(RECORDS))
            if data[:2] == b"PK":
                headers = [cell.value for cell in load_workbook(io.BytesIO(data)).active[1]]
                self.assertEqual(headers, expected_headers)
            elif data.lstrip().startswith(b"["):
                self.assertEqual(list(json.loads(data)[0]), expected_headers)
            else:
                self.assertEqual(data.decode("utf-8-sig").splitlines()[0], "|".join(expected_headers))


if __name__ == "__main__":
    unittest.main()
