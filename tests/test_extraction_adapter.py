"""Regression tests for nullable extraction values and confidence sidecars."""

import json

from app import parse_batch_payload
from extraction_adapter import coerce_nullable_value, confidence_map


def test_nullable_gemini_fields_preserve_original_fact_shape_and_confidence():
    # Simulate Gemini returning clear, blurry, and handwritten occurrences.
    payload = {
        "documents": [{
            "document": "invoice.pdf",
            "document_type": "Invoice",
            "facts": [
                {"label": "Invoice Number", "value": "INV-1", "evidence": "INV-1", "confidence_score": 96},
                {"label": "Total Amount", "value": None, "evidence": "blurred", "confidence_score": 31},
                {"label": "Driver Name", "value": "Rajesh", "evidence": "handwritten",
                 "ambiguity_notes": "last character unclear", "confidence_score": 58},
            ],
        }],
        "cross_references": [],
    }

    # Parse through the same adapter used by the extraction pipeline.
    documents, references = parse_batch_payload(json.dumps(payload), ["invoice.pdf"])
    records = documents["invoice.pdf"]["facts"]

    assert references == []
    assert [record["value"] for record in records] == ["INV-1", "", "Rajesh"]
    assert all({"page_number", "label", "concept", "value", "evidence", "confidence_score"}
               <= set(record) for record in records)
    assert "last character unclear" in records[2]["evidence"]
    assert confidence_map(records) == {
        "Invoice Number#0": 96.0,
        "Total Amount#0": 31.0,
        "Driver Name#0": 58.0,
    }


def test_nullable_numeric_values_coerce_and_zero_confidence():
    # Numeric nulls become schema-safe zero values with zero confidence.
    assert coerce_nullable_value(None, str, 40) == ("", 40.0)
    assert coerce_nullable_value(None, int, 40) == (0, 0.0)
    assert coerce_nullable_value(None, float, 40) == (0.0, 0.0)


def test_legacy_fact_name_alias_is_normalized_to_a_real_label():
    payload = {
        "documents": [{
            "document": "scan.jpg",
            "document_type": "Invoice",
            "facts": [{"field_name": "Invoice Number", "value": "INV-2"}],
        }],
        "cross_references": [],
    }
    documents, _ = parse_batch_payload(json.dumps(payload), ["scan.jpg"])
    assert documents["scan.jpg"]["facts"][0]["label"] == "Invoice Number"
