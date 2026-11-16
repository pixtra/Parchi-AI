import json
import sqlite3
import tempfile
import unittest
from contextlib import contextmanager
from pathlib import Path
from unittest.mock import patch

import app
from validated_data import (
    VALIDATED_DATA_SCHEMA,
    apply_validated_records,
    edit_validated_record,
    list_validated_edit_history,
    set_review_approval,
    undo_validated_edit,
)


class ValidatedDataTests(unittest.TestCase):
    def setUp(self):
        self.db = sqlite3.connect(":memory:")
        self.db.row_factory = sqlite3.Row
        self.db.executescript(VALIDATED_DATA_SCHEMA)
        self.db.execute("""CREATE TABLE facts (
            id TEXT PRIMARY KEY, label TEXT, value TEXT, confidence_score INTEGER
        )""")
        self.db.executemany("INSERT INTO facts VALUES (?, ?, ?, ?)", [
            ("f1", "Driver Name", "Asha", 72),
            ("f2", "Driver Name", "Asha", 61),
            ("f3", "Driver", "Ravi", 88),
            ("f4", "Vehicle Number", "AB-123", 93),
        ])
        self.baseline = [
            {"record_id": "f1", "document_name": "one.pdf", "source_label": "Driver Name",
             "label": "Driver", "value": "Asha", "confidence_score": 72},
            {"record_id": "f2", "document_name": "two.pdf", "source_label": "Driver Name",
             "label": "Driver", "value": "Asha", "confidence_score": 61},
            {"record_id": "f3", "document_name": "three.pdf", "source_label": "Driver",
             "label": "Driver", "value": "Ravi", "confidence_score": 88},
            {"record_id": "f4", "document_name": "four.pdf", "source_label": "Vehicle Number",
             "label": "Vehicle Number", "value": "AB-123", "confidence_score": 93},
        ]
        self.main_snapshot = [tuple(row) for row in self.db.execute(
            "SELECT id, label, value, confidence_score FROM facts ORDER BY id")]
        self.docket_id = "dock"
        self.template_id = "template"

    def tearDown(self):
        self.db.close()

    def edit(self, record_id, label=None, value=None):
        current = next(row for row in apply_validated_records(
            self.db, self.docket_id, self.template_id, self.baseline) if row["record_id"] == record_id)
        edit_validated_record(
            self.db, self.docket_id, self.template_id, self.baseline, record_id,
            label if label is not None else current["label"],
            value if value is not None else current["value"],
        )

    def rows(self):
        return apply_validated_records(self.db, self.docket_id, self.template_id, self.baseline)

    def test_label_only_value_only_and_both_field_edits_set_confidence_100(self):
        self.edit("f1", label="Primary Driver")
        self.edit("f2", value="Asha K.")
        self.edit("f3", label="Operator", value="Ravi K.")
        current = {row["record_id"]: row for row in self.rows()}
        self.assertEqual((current["f1"]["label"], current["f1"]["value"],
                          current["f1"]["confidence_score"]), ("Primary Driver", "Asha", 100))
        self.assertEqual((current["f2"]["label"], current["f2"]["value"],
                          current["f2"]["confidence_score"]), ("Driver", "Asha K.", 100))
        self.assertEqual((current["f3"]["label"], current["f3"]["value"],
                          current["f3"]["confidence_score"]), ("Operator", "Ravi K.", 100))
        self.assertEqual(current["f4"]["confidence_score"], 93)
        self.assertEqual([tuple(row) for row in self.db.execute(
            "SELECT id, label, value, confidence_score FROM facts ORDER BY id")], self.main_snapshot)

    def test_undo_selected_field_keeps_other_record_and_field_edits(self):
        self.edit("f1", label="Primary Driver")
        self.edit("f1", value="Asha K.")
        self.edit("f2", label="Second Driver")
        history = list_validated_edit_history(self.db, self.docket_id)
        label_edit = next(row for row in history if row["record_id"] == "f1" and row["field_name"] == "label")
        undo_validated_edit(self.db, self.docket_id, label_edit["id"])
        current = {row["record_id"]: row for row in self.rows()}
        self.assertEqual((current["f1"]["label"], current["f1"]["value"]), ("Driver", "Asha K."))
        self.assertEqual(current["f1"]["confidence_score"], 100)
        self.assertEqual(current["f2"]["label"], "Second Driver")

    def test_undo_several_individually_restores_base_values_and_confidence(self):
        self.edit("f1", label="Primary Driver", value="Asha K.")
        history = list_validated_edit_history(self.db, self.docket_id)
        for item in history:
            undo_validated_edit(self.db, self.docket_id, item["id"])
        current = {row["record_id"]: row for row in self.rows()}
        self.assertEqual((current["f1"]["label"], current["f1"]["value"],
                          current["f1"]["confidence_score"]), ("Driver", "Asha", 72))
        history_after = list_validated_edit_history(self.db, self.docket_id)
        self.assertTrue(all(item["undone_at"] for item in history_after))

    def test_edit_and_history_survive_connection_reload(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "validated.sqlite"
            self.db.close()
            db = sqlite3.connect(path)
            db.row_factory = sqlite3.Row
            db.executescript(VALIDATED_DATA_SCHEMA)
            db.execute("CREATE TABLE facts (id TEXT PRIMARY KEY, label TEXT, value TEXT, confidence_score INTEGER)")
            db.executemany("INSERT INTO facts VALUES (?, ?, ?, ?)", self.main_snapshot)
            edit_validated_record(db, self.docket_id, self.template_id, self.baseline,
                                  "f1", "Primary Driver", "Asha")
            db.commit()
            db.close()
            reopened = sqlite3.connect(path)
            reopened.row_factory = sqlite3.Row
            current = apply_validated_records(reopened, self.docket_id, self.template_id, self.baseline)
            self.assertEqual(current[0]["label"], "Primary Driver")
            self.assertEqual(current[0]["confidence_score"], 100)
            self.assertEqual(len(list_validated_edit_history(reopened, self.docket_id)), 1)
            self.assertEqual([tuple(row) for row in reopened.execute(
                "SELECT id, label, value, confidence_score FROM facts ORDER BY id")], self.main_snapshot)
            reopened.close()
            self.db = sqlite3.connect(":memory:")

    def test_full_metadata_copy_edits_independently_of_original(self):
        scope = app.VALIDATED_METADATA_SCOPE
        edit_validated_record(self.db, self.docket_id, scope, self.baseline,
                              "f2", "Updated Driver", "Asha")
        validated = apply_validated_records(self.db, self.docket_id, scope, self.baseline)
        self.assertEqual(len(validated), len(self.baseline))
        self.assertEqual(validated[1]["label"], "Updated Driver")
        self.assertEqual(validated[1]["confidence_score"], 100)
        self.assertEqual(self.baseline[1]["label"], "Driver")
        self.assertEqual([tuple(row) for row in self.db.execute(
            "SELECT id, label, value, confidence_score FROM facts ORDER BY id")], self.main_snapshot)

    def test_review_checklist_persists_per_occurrence_without_changing_source_fact(self):
        scope = app.VALIDATED_METADATA_SCOPE
        self.assertTrue(set_review_approval(self.db, self.docket_id, scope, self.baseline, "f1", True))
        reviewed = apply_validated_records(self.db, self.docket_id, scope, self.baseline)
        self.assertTrue(next(row for row in reviewed if row["record_id"] == "f1")["review_approved"])
        self.assertFalse(next(row for row in reviewed if row["record_id"] == "f2")["review_approved"])
        review_export, _ = app.build_metadata_export(reviewed, "Review Metadata")
        self.assertTrue(json.loads(review_export)[0]["review_approved"])
        self.assertEqual([tuple(row) for row in self.db.execute(
            "SELECT id, label, value, confidence_score FROM facts ORDER BY id")], self.main_snapshot)

    def test_main_and_validated_json_psv_xlsx_exports_are_separate(self):
        edit_validated_record(self.db, self.docket_id, app.VALIDATED_METADATA_SCOPE,
                              self.baseline, "f1", "Primary Driver", "Asha K.")
        validated = apply_validated_records(
            self.db, self.docket_id, app.VALIDATED_METADATA_SCOPE, self.baseline
        )
        canonical_rows = [{
            "document_name": row["document_name"], "label": row["source_label"],
            "value": row["value"], "confidence_score": row["confidence_score"],
        } for row in self.baseline]
        with patch.object(app, "fetch_metadata", return_value=canonical_rows):
            main = app.metadata_records(self.docket_id)
        self.assertEqual(main, canonical_rows)
        validated_export = validated

        main_json, _ = app.build_metadata_export(main, "Main Metadata")
        validated_json, _ = app.build_metadata_export(validated_export, "Review Metadata")
        self.assertEqual(json.loads(main_json), main)
        self.assertEqual(json.loads(validated_json)[0]["label"], "Primary Driver")
        self.assertEqual(json.loads(validated_json)[0]["confidence_score"], 100)
        self.assertEqual(json.loads(validated_json)[0]["document_name"], "one.pdf")
        self.assertEqual(json.loads(validated_json)[0]["record_id"], "f1")

        main_psv, _ = app.build_metadata_psv(main, "Main Metadata")
        validated_psv, _ = app.build_metadata_psv(validated_export, "Review Metadata")
        self.assertNotEqual(main_psv, validated_psv)
        self.assertEqual(main_psv.splitlines()[0].lstrip(b"\xef\xbb\xbf"),
                         b"document_name|label|value|confidence_score")
        self.assertIn(b"source_label|record_id", validated_psv.splitlines()[0])

        main_xlsx, _ = app.build_metadata_xlsx(main, "Main Metadata")
        validated_xlsx, _ = app.build_metadata_xlsx(validated_export, "Review Metadata")
        self.assertNotEqual(main_xlsx, validated_xlsx)
        self.assertEqual([tuple(row) for row in self.db.execute(
            "SELECT id, label, value, confidence_score FROM facts ORDER BY id")], self.main_snapshot)

    def test_client_template_group_emits_only_its_prioritized_occurrence(self):
        db = sqlite3.connect(":memory:")
        db.row_factory = sqlite3.Row
        db.executescript("""
            CREATE TABLE documents (id TEXT PRIMARY KEY, filename TEXT);
            CREATE TABLE facts (
                id TEXT PRIMARY KEY, docket_id TEXT, document_id TEXT, label TEXT, value TEXT,
                confidence_score INTEGER, concept TEXT, page_number INTEGER, created_at TEXT
            );
            CREATE TABLE value_groups (
                id TEXT PRIMARY KEY, name TEXT, primary_fact_id TEXT, created_at TEXT, updated_at TEXT
            );
            CREATE TABLE value_group_members (
                group_id TEXT, fact_id TEXT, priority INTEGER, PRIMARY KEY(group_id,fact_id)
            );
        """)
        db.executemany("INSERT INTO documents VALUES (?, ?)", [("d1", "first.pdf"), ("d2", "second.pdf")])
        db.executemany("INSERT INTO facts VALUES (?, 'dock', ?, ?, ?, ?, 'driver', 1, ?)", [
            ("f1", "d1", "Driver", "Higher confidence", 98, "1"),
            ("f2", "d2", "Name of Driver", "User priority", 52, "2"),
        ])
        db.execute("INSERT INTO value_groups VALUES ('g1','Driver', 'f2','1','1')")
        db.executemany("INSERT INTO value_group_members VALUES ('g1', ?, ?)", [("f2", 0), ("f1", 1)])

        @contextmanager
        def connected():
            yield db

        field = {"id": "field1", "source_type": "group", "source_id": "g1",
                 "destination_label": "Driver", "position": 0}
        try:
            with patch.object(app, "connect", connected), \
                 patch.object(app, "template_field_rows", return_value=[field]), \
                 patch.object(app, "approved_mappings", return_value={}):
                records, missing = app.client_template_records(
                    "dock", {"id": "template1", "labels_json": '["Driver"]'}
                )
            self.assertEqual(missing, [])
            self.assertEqual(len(records), 1)
            self.assertEqual(records[0]["record_id"], "f2")
            self.assertEqual(records[0]["source_label"], "Name of Driver")
            self.assertEqual(records[0]["value"], "User priority")
        finally:
            db.close()


if __name__ == "__main__":
    unittest.main()
