"""Regression tests for key pooling and manual group priority behavior."""

import sqlite3
import unittest
from types import SimpleNamespace
from unittest.mock import patch

import app
from value_groups import (
    VALUE_GROUP_SCHEMA,
    add_member,
    create_group,
    delete_group,
    load_groups,
    remove_member,
    reorder_members,
)


class PoolingTests(unittest.TestCase):
    """Verify that key rotation skips cooling and invalid credentials."""
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
    """Verify that the first saved member controls manual group priority."""
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

    def test_saved_occurrence_order_overrides_confidence_and_survives_reload(self):
        reorder_members(self.db, "dock", self.group_id, ["f2", "f1"])
        group = load_groups(self.db, "dock")[0]
        self.assertEqual(group["primary_fact_id"], "f2")
        self.assertEqual(group["primary_label"], "Driver Name")
        self.assertEqual([m.fact_id for m in group["members"]], ["f2", "f1"])

    def test_reorder_requires_exact_current_membership(self):
        with self.assertRaises(ValueError):
            reorder_members(self.db, "dock", self.group_id, ["f1"])
        self.assertEqual([m.fact_id for m in load_groups(self.db, "dock")[0]["members"]],
                         ["f1", "f2"])

    def test_removing_first_member_preserves_remaining_order(self):
        reorder_members(self.db, "dock", self.group_id, ["f2", "f1"])
        remove_member(self.db, "dock", self.group_id, "f2")
        group = load_groups(self.db, "dock")[0]
        self.assertEqual(group["primary_fact_id"], "f1")
        self.assertEqual(group["primary_label"], "Driver")

    def test_add_and_remove_change_one_occurrence_without_deleting_fact(self):
        self.db.execute("INSERT INTO documents VALUES ('d3', 'three.pdf')")
        self.db.execute("INSERT INTO facts VALUES ('f3','dock','d3','Driver','A',70,'pending','driver',1,'3')")
        add_member(self.db, "dock", self.group_id, "f3")
        remove_member(self.db, "dock", self.group_id, "f1")
        members = {member.fact_id for member in load_groups(self.db, "dock")[0]["members"]}
        self.assertEqual(members, {"f2", "f3"})
        self.assertEqual(self.db.execute("SELECT COUNT(*) FROM facts WHERE id = 'f1'").fetchone()[0], 1)
        self.assertEqual(self.db.execute("SELECT COUNT(*) FROM facts WHERE id = 'f3'").fetchone()[0], 1)

    def test_deleting_group_keeps_facts_and_ungroups_all_occurrences(self):
        delete_group(self.db, "dock", self.group_id)
        self.assertEqual(load_groups(self.db, "dock"), [])
        self.assertEqual(self.db.execute("SELECT COUNT(*) FROM facts").fetchone()[0], 2)
        self.assertEqual(self.db.execute("SELECT COUNT(*) FROM value_group_members").fetchone()[0], 0)


if __name__ == "__main__":
    unittest.main()
