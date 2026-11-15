import sqlite3
import unittest
from types import SimpleNamespace
from unittest.mock import patch

import app
from value_groups import (
    VALUE_GROUP_SCHEMA,
    create_group,
    load_groups,
    replace_members,
    set_primary_member,
)


class PoolingTests(unittest.TestCase):
    def setUp(self):
        self.keys = [{"id": "a"}, {"id": "b"}, {"id": "c"}]
        app._KEY_ROTATION_INDEX = 0

    def test_round_robin_advances_after_success(self):
        self.assertEqual([row["id"] for row in app.rotated_keys(self.keys)], ["a", "b", "c"])
        app.record_key_success(self.keys, "b")
        self.assertEqual([row["id"] for row in app.rotated_keys(self.keys)], ["c", "a", "b"])

    def test_cooldown_wait_uses_first_available_model(self):
        models = list(app.GEMINI_MODELS)
        delays = {"a": 0, "a:" + models[0]: 12, "a:" + models[1]: 40}
        with patch.object(app, "GEMINI_MODELS", models[:2]), \
             patch.object(app, "cooldown_remaining", side_effect=lambda key: delays.get(key, 0)):
            self.assertEqual(app.earliest_cooldown([self.keys[0]]), 12)

    def test_key_wide_cooldown_blocks_all_models_until_key_recovers(self):
        delays = {"a": 25, "a:" + app.GEMINI_MODELS[0]: 8}
        with patch.object(app, "cooldown_remaining", side_effect=lambda key: delays.get(key, 0)):
            self.assertEqual(app.earliest_cooldown([self.keys[0]]), 25)

    def test_quota_failure_falls_through_to_next_key(self):
        payload = '{"documents":[{"document":"doc.pdf","document_type":"Invoice",' \
                  '"facts":[{"label":"Invoice Number","value":"INV-1",' \
                  '"evidence":"INV-1","confidence_score":90}]}],"cross_references":[]}'

        class Files:
            def delete(self, name):
                pass

        class Client:
            def __init__(self, fail):
                self.fail = fail
                self.files = Files()
                self.calls = 0
                self.models = self

            def generate_content(self, **kwargs):
                self.calls += 1
                if self.fail:
                    raise RuntimeError("429 RESOURCE_EXHAUSTED quota")
                return SimpleNamespace(text=payload, candidates=[SimpleNamespace(finish_reason="STOP")])

        clients = {"a": Client(True), "b": Client(False)}
        keys = [dict(id=key, key=key, masked=f"mask-{key}") for key in ("a", "b")]
        unit = {"label": "doc.pdf", "pages": 1}
        with patch.object(app, "GEMINI_MODELS", ["test-model"]), \
             patch.object(app, "upload_parts", side_effect=lambda entry, batch:
                          (clients[entry["id"]], [SimpleNamespace(name=entry["id"])])), \
             patch.object(app.types, "GenerateContentConfig", side_effect=lambda **kwargs: kwargs), \
             patch.object(app, "cooldown_remaining", return_value=0), \
             patch.object(app, "set_cooldown"):
            parsed, used = app.call_gemini_batch([unit], keys)
        self.assertEqual(used, "mask-b · test-model")
        self.assertEqual(parsed[0]["doc.pdf"]["facts"][0]["value"], "INV-1")
        self.assertGreaterEqual(clients["a"].calls, 1)
        self.assertEqual(clients["b"].calls, 1)


class GroupPriorityTests(unittest.TestCase):
    def setUp(self):
        self.db = sqlite3.connect(":memory:")
        self.db.row_factory = sqlite3.Row
        self.db.executescript("""
            CREATE TABLE facts (
                id TEXT PRIMARY KEY, docket_id TEXT, document_id TEXT, label TEXT,
                value TEXT, confidence_score INTEGER, review_status TEXT, concept TEXT,
                page_number INTEGER, created_at TEXT
            );
            CREATE TABLE documents (id TEXT PRIMARY KEY, filename TEXT);
        """)
        self.db.executescript(VALUE_GROUP_SCHEMA)
        self.db.execute("INSERT INTO documents VALUES ('d1', 'one.pdf')")
        self.db.execute("INSERT INTO documents VALUES ('d2', 'two.pdf')")
        self.db.execute("INSERT INTO facts VALUES ('f1','dock','d1','Driver','A',95,'pending','driver',1,'1')")
        self.db.execute("INSERT INTO facts VALUES ('f2','dock','d2','Driver Name','B',60,'pending','driver_name',1,'2')")
        self.group_id = create_group(self.db, "dock", "Driver", ["f1", "f2"])

    def tearDown(self):
        self.db.close()

    def test_manual_primary_overrides_confidence_and_survives_reload(self):
        set_primary_member(self.db, "dock", self.group_id, "f2")
        group = load_groups(self.db, "dock")[0]
        self.assertEqual(group["primary_fact_id"], "f2")
        self.assertEqual(group["primary_label"], "Driver Name")
        self.assertTrue(group["manual_primary"])

    def test_removed_manual_primary_resets_to_confidence_choice(self):
        set_primary_member(self.db, "dock", self.group_id, "f2")
        replace_members(self.db, "dock", self.group_id, ["f1"])
        group = load_groups(self.db, "dock")[0]
        self.assertEqual(group["primary_fact_id"], "f1")
        self.assertFalse(group["manual_primary"])


if __name__ == "__main__":
    unittest.main()
