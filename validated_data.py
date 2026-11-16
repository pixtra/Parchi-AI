"""Persist user-reviewed template values separately from canonical extraction facts."""
from __future__ import annotations

import sqlite3
import uuid
from datetime import datetime, timezone
from typing import Iterable, List, Mapping, Optional


VALIDATED_DATA_SCHEMA = """
CREATE TABLE IF NOT EXISTS validated_template_records (
    docket_id TEXT NOT NULL,
    template_id TEXT NOT NULL,
    record_id TEXT NOT NULL,
    base_label TEXT NOT NULL,
    base_value TEXT NOT NULL,
    base_confidence INTEGER NOT NULL,
    destination_label TEXT NOT NULL,
    value TEXT NOT NULL,
    confidence_score INTEGER NOT NULL,
    review_approved INTEGER NOT NULL DEFAULT 0 CHECK(review_approved IN (0, 1)),
    updated_at TEXT NOT NULL,
    PRIMARY KEY (docket_id, template_id, record_id)
);
CREATE TABLE IF NOT EXISTS validated_edit_history (
    edit_order INTEGER PRIMARY KEY AUTOINCREMENT,
    id TEXT NOT NULL UNIQUE,
    docket_id TEXT NOT NULL,
    template_id TEXT NOT NULL,
    record_id TEXT NOT NULL,
    document_name TEXT NOT NULL,
    source_label TEXT NOT NULL,
    field_name TEXT NOT NULL CHECK(field_name IN ('label', 'value')),
    old_value TEXT NOT NULL,
    new_value TEXT NOT NULL,
    created_at TEXT NOT NULL,
    undone_at TEXT
);
CREATE INDEX IF NOT EXISTS idx_validated_history_docket
    ON validated_edit_history(docket_id, edit_order DESC);
CREATE INDEX IF NOT EXISTS idx_validated_history_record
    ON validated_edit_history(docket_id, template_id, record_id, field_name, edit_order);
"""

# This sidecar stores user edits while the facts table remains the canonical extraction source.


def _now() -> str:
    """Return the UTC timestamp used by review records and history entries."""
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _dict(row: sqlite3.Row | Mapping) -> dict:
    """Convert either a SQLite row or mapping into a normal dictionary."""
    return dict(row)


def apply_validated_records(conn: sqlite3.Connection, docket_id: str, template_id: str,
                            records: Iterable[Mapping]) -> List[dict]:
    """Overlay saved review values onto live template occurrences without writing facts."""
    # Build a display copy so the caller never receives mutated source mappings.
    output = []
    for source in records:
        record = _dict(source)
        # Look up by stable occurrence ID; duplicate labels remain independent rows.
        saved = conn.execute(
            """SELECT destination_label, value, confidence_score, review_approved
               FROM validated_template_records
               WHERE docket_id = ? AND template_id = ? AND record_id = ?""",
            (docket_id, template_id, str(record["record_id"])),
        ).fetchone()
        if saved is not None:
            # Apply only sidecar values to the copy used by Viewer and Export.
            record["label"] = saved["destination_label"]
            record["value"] = saved["value"]
            record["confidence_score"] = int(saved["confidence_score"])
            record["review_approved"] = bool(saved["review_approved"])
            if saved["review_approved"]:
                record["review_status"] = "approved"
        else:
            record["review_approved"] = False
        output.append(record)
    return output


def edit_validated_record(conn: sqlite3.Connection, docket_id: str, template_id: str,
                          baseline_records: Iterable[Mapping], record_id: str,
                          destination_label: str, value: str) -> bool:
    """Edit a single mapped occurrence, recording each changed field independently.

    The source records are supplied by the read-only template mapping layer and are used
    only to validate identity and seed the separate validated-data row.
    """
    # Normalize the identity received from the table widget.
    record_id = str(record_id or "").strip()
    base = next((row for row in baseline_records if str(row.get("record_id", "")) == record_id), None)
    if not record_id or base is None:
        raise ValueError("That template occurrence is no longer available. Reload the Viewer.")
    new_label = str(destination_label or "").strip()
    new_value = str(value if value is not None else "")
    if not new_label:
        raise ValueError("Destination label cannot be empty.")
    if len(new_label) > 500:
        raise ValueError("Destination labels must be 500 characters or fewer.")

    # Seed a sidecar row from the original fact the first time it is edited.
    existing = conn.execute(
        """SELECT * FROM validated_template_records
           WHERE docket_id = ? AND template_id = ? AND record_id = ?""",
        (docket_id, template_id, record_id),
    ).fetchone()
    if existing is None:
        base_label = str(base["label"])
        base_value = str(base["value"] if base["value"] is not None else "")
        base_confidence = int(base["confidence_score"])
        current_label, current_value = base_label, base_value
        conn.execute(
            """INSERT INTO validated_template_records
               (docket_id, template_id, record_id, base_label, base_value, base_confidence,
                destination_label, value, confidence_score, review_approved, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (docket_id, template_id, record_id, base_label, base_value, base_confidence,
             base_label, base_value, base_confidence, 0, _now()),
        )
    else:
        base_label, base_value = existing["base_label"], existing["base_value"]
        base_confidence = int(existing["base_confidence"])
        current_label, current_value = existing["destination_label"], existing["value"]

    # Record each changed field separately so each edit can be undone independently.
    changes = []
    if new_label != current_label:
        changes.append(("label", current_label, new_label))
    if new_value != current_value:
        changes.append(("value", current_value, new_value))
    if not changes:
        return False

    # Use one timestamp for all fields changed in one table submission.
    now = _now()
    for field_name, old_value, new_value_for_field in changes:
        conn.execute(
            """INSERT INTO validated_edit_history
               (id, docket_id, template_id, record_id, document_name, source_label,
                field_name, old_value, new_value, created_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (str(uuid.uuid4()), docket_id, template_id, record_id,
             str(base.get("document_name", "")), str(base.get("source_label", "")),
             field_name, old_value, new_value_for_field, now),
        )
    # Edited review records always receive full review confidence in the sidecar.
    conn.execute(
        """UPDATE validated_template_records
           SET destination_label = ?, value = ?, confidence_score = 100, updated_at = ?
           WHERE docket_id = ? AND template_id = ? AND record_id = ?""",
        (new_label, new_value, now, docket_id, template_id, record_id),
    )
    return True


def set_review_approval(conn: sqlite3.Connection, docket_id: str, template_id: str,
                        baseline_records: Iterable[Mapping], record_id: str, approved: bool) -> bool:
    """Persist one occurrence's review checklist state without altering source facts."""
    record_id = str(record_id or "").strip()
    base = next((row for row in baseline_records if str(row.get("record_id", "")) == record_id), None)
    if not record_id or base is None:
        raise ValueError("That review occurrence is no longer available. Reload the Viewer.")
    current = conn.execute(
        """SELECT review_approved, base_confidence FROM validated_template_records
           WHERE docket_id = ? AND template_id = ? AND record_id = ?""",
        (docket_id, template_id, record_id),
    ).fetchone()
    # SQLite stores the checklist value as an integer constrained to 0 or 1.
    target = int(bool(approved))
    if current is not None and int(current["review_approved"]) == target:
        return False
    if current is None:
        # Approval creates a sidecar copy without changing the source fact.
        conn.execute(
            """INSERT INTO validated_template_records
               (docket_id, template_id, record_id, base_label, base_value, base_confidence,
                destination_label, value, confidence_score, review_approved, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (docket_id, template_id, record_id, str(base["label"]),
             str(base["value"] if base["value"] is not None else ""), int(base["confidence_score"]),
             str(base["label"]), str(base["value"] if base["value"] is not None else ""),
             int(base["confidence_score"]), target, _now()),
        )
    else:
        # Existing edits remain at confidence 100 even if the approval checkbox is cleared.
        has_active_edits = conn.execute(
            """SELECT 1 FROM validated_edit_history
               WHERE docket_id = ? AND template_id = ? AND record_id = ?
                 AND undone_at IS NULL LIMIT 1""",
            (docket_id, template_id, record_id),
        ).fetchone() is not None
        confidence = 100 if target or has_active_edits else int(current["base_confidence"])
        conn.execute(
            """UPDATE validated_template_records
               SET review_approved = ?, confidence_score = ?, updated_at = ?
               WHERE docket_id = ? AND template_id = ? AND record_id = ?""",
            (target, confidence, _now(), docket_id, template_id, record_id),
        )
    if current is None and target:
        conn.execute(
            """UPDATE validated_template_records SET confidence_score = 100
               WHERE docket_id = ? AND template_id = ? AND record_id = ?""",
            (docket_id, template_id, record_id),
        )
    return True


def list_validated_edit_history(conn: sqlite3.Connection, docket_id: str,
                                template_id: Optional[str] = None) -> List[dict]:
    """Read edit history newest first, optionally limited to one template."""
    # Keep the query parameterized because docket and template IDs come from the UI.
    sql = """SELECT id, template_id, record_id, document_name, source_label, field_name,
                    old_value, new_value, created_at, undone_at
             FROM validated_edit_history WHERE docket_id = ?"""
    params: list = [docket_id]
    if template_id:
        sql += " AND template_id = ?"
        params.append(template_id)
    sql += " ORDER BY edit_order DESC"
    return [dict(row) for row in conn.execute(sql, params).fetchall()]


def undo_validated_edit(conn: sqlite3.Connection, docket_id: str, edit_id: str) -> bool:
    """Undo one field edit while preserving later edits to the same or other fields."""
    # Load the history event and reject repeated undo requests.
    history = conn.execute(
        """SELECT * FROM validated_edit_history WHERE docket_id = ? AND id = ?""",
        (docket_id, edit_id),
    ).fetchone()
    if history is None:
        raise ValueError("That edit is no longer available.")
    if history["undone_at"]:
        return False
    validated = conn.execute(
        """SELECT * FROM validated_template_records
           WHERE docket_id = ? AND template_id = ? AND record_id = ?""",
        (docket_id, history["template_id"], history["record_id"]),
    ).fetchone()
    if validated is None:
        raise ValueError("The validated occurrence no longer exists.")

    # Map the safe, constrained field name to the sidecar column being restored.
    field_name = history["field_name"]
    current_column = "destination_label" if field_name == "label" else "value"
    latest_active = conn.execute(
        """SELECT edit_order, new_value FROM validated_edit_history
           WHERE docket_id = ? AND template_id = ? AND record_id = ?
             AND field_name = ? AND undone_at IS NULL
           ORDER BY edit_order DESC LIMIT 1""",
        (docket_id, history["template_id"], history["record_id"], field_name),
    ).fetchone()
    if latest_active and int(latest_active["edit_order"]) == int(history["edit_order"]):
        # Restore the previous active value only when this event is currently visible.
        previous = conn.execute(
            """SELECT new_value FROM validated_edit_history
               WHERE docket_id = ? AND template_id = ? AND record_id = ?
                 AND field_name = ? AND undone_at IS NULL AND edit_order < ?
               ORDER BY edit_order DESC LIMIT 1""",
            (docket_id, history["template_id"], history["record_id"], field_name,
             history["edit_order"]),
        ).fetchone()
        if previous:
            restore = previous["new_value"]
        else:
            restore = validated["base_label"] if field_name == "label" else validated["base_value"]
        conn.execute(
            f"UPDATE validated_template_records SET {current_column} = ?, updated_at = ? "
            "WHERE docket_id = ? AND template_id = ? AND record_id = ?",
            (restore, _now(), docket_id, history["template_id"], history["record_id"]),
        )

    # Mark the event undone even when a newer edit means the displayed value stays put.
    undone_at = _now()
    conn.execute("UPDATE validated_edit_history SET undone_at = ? WHERE id = ?", (undone_at, edit_id))
    still_edited = conn.execute(
        """SELECT 1 FROM validated_edit_history
           WHERE docket_id = ? AND template_id = ? AND record_id = ? AND undone_at IS NULL LIMIT 1""",
        (docket_id, history["template_id"], history["record_id"]),
    ).fetchone()
    score = 100 if still_edited or bool(validated["review_approved"]) \
        else int(validated["base_confidence"])
    conn.execute(
        """UPDATE validated_template_records SET confidence_score = ?, updated_at = ?
           WHERE docket_id = ? AND template_id = ? AND record_id = ?""",
        (score, undone_at, docket_id, history["template_id"], history["record_id"]),
    )
    return True
