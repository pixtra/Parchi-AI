"""User-controlled value groups: manual grouping of extracted facts across documents.

Grouping never reads label text or values to decide membership. The user chooses every
member. Main Metadata (the `facts` table) is never written here; only the two group
tables change.
"""
from __future__ import annotations

import sqlite3
import uuid
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Iterable, List, Optional

VALUE_GROUP_SCHEMA = """
CREATE TABLE IF NOT EXISTS value_groups (
    id TEXT PRIMARY KEY,
    docket_id TEXT NOT NULL REFERENCES dockets(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    UNIQUE(docket_id, name)
);
CREATE TABLE IF NOT EXISTS value_group_members (
    group_id TEXT NOT NULL REFERENCES value_groups(id) ON DELETE CASCADE,
    fact_id TEXT NOT NULL UNIQUE REFERENCES facts(id) ON DELETE CASCADE,
    PRIMARY KEY (group_id, fact_id)
);
CREATE INDEX IF NOT EXISTS idx_value_groups_docket ON value_groups(docket_id);
"""


class GroupError(ValueError):
    """A user-facing validation failure. The message is safe to show in the UI."""


@dataclass(frozen=True)
class GroupMember:
    fact_id: str
    label: str
    value: str
    document_name: str
    confidence_score: int   # raw extraction score, never adjusted by grouping


def primary_member(members: Iterable[GroupMember]) -> Optional[GroupMember]:
    """Highest raw score wins. Ties go to the lowest fact_id, so the result is deterministic."""
    members = list(members)
    if not members:
        return None
    return min(members, key=lambda m: (-m.confidence_score, m.fact_id))


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _clean_ids(fact_ids: Iterable[str]) -> List[str]:
    """Drop duplicates while keeping order. Repeated browser events cannot add a record twice."""
    seen: dict[str, None] = {}
    for fact_id in fact_ids:
        text = str(fact_id).strip()
        if text:
            seen.setdefault(text, None)
    return list(seen)


def _check_facts(conn: sqlite3.Connection, docket_id: str, fact_ids: List[str]) -> None:
    """Every fact must exist in this docket and must not be rejected."""
    if not fact_ids:
        return
    placeholders = ",".join("?" * len(fact_ids))
    rows = conn.execute(
        f"""SELECT id, review_status FROM facts
            WHERE docket_id = ? AND id IN ({placeholders})""",
        (docket_id, *fact_ids),
    ).fetchall()
    found = {row["id"] if isinstance(row, sqlite3.Row) else row[0]: row for row in rows}
    missing = [f for f in fact_ids if f not in found]
    if missing:
        raise GroupError("Some selected records no longer exist. Refresh the page and try again.")
    rejected = [f for f in fact_ids
                if (found[f]["review_status"] if isinstance(found[f], sqlite3.Row) else found[f][1])
                == "rejected"]
    if rejected:
        raise GroupError("Rejected records cannot be grouped. Restore them in Review first.")


def _check_not_grouped(conn: sqlite3.Connection, fact_ids: List[str], except_group: Optional[str]) -> None:
    placeholders = ",".join("?" * len(fact_ids))
    params: list = [*fact_ids]
    sql = f"SELECT fact_id, group_id FROM value_group_members WHERE fact_id IN ({placeholders})"
    if except_group is not None:
        sql += " AND group_id != ?"
        params.append(except_group)
    if conn.execute(sql, params).fetchone():
        raise GroupError("A selected record already belongs to another group. Remove it there first.")


def _validate_name(name: str) -> str:
    cleaned = " ".join(str(name or "").split())
    if not cleaned:
        raise GroupError("Enter a group name before saving.")
    if len(cleaned) > 120:
        raise GroupError("Group names must be 120 characters or fewer.")
    return cleaned


def create_group(conn: sqlite3.Connection, docket_id: str, name: str, fact_ids: Iterable[str]) -> str:
    name = _validate_name(name)
    ids = _clean_ids(fact_ids)
    if not ids:
        raise GroupError("Add at least one record to the group before saving.")
    _check_facts(conn, docket_id, ids)
    _check_not_grouped(conn, ids, except_group=None)
    group_id = str(uuid.uuid4())
    now = _now()
    try:
        conn.execute(
            "INSERT INTO value_groups (id, docket_id, name, created_at, updated_at) VALUES (?, ?, ?, ?, ?)",
            (group_id, docket_id, name, now, now),
        )
    except sqlite3.IntegrityError as exc:
        raise GroupError(f"A group named “{name}” already exists.") from exc
    conn.executemany(
        "INSERT INTO value_group_members (group_id, fact_id) VALUES (?, ?)",
        [(group_id, fact_id) for fact_id in ids],
    )
    return group_id


def replace_members(conn: sqlite3.Connection, docket_id: str, group_id: str, fact_ids: Iterable[str]) -> None:
    """Set a group's exact membership. Removed records return to their document sections."""
    ids = _clean_ids(fact_ids)
    if not ids:
        raise GroupError("A group needs at least one record. Delete the group instead.")
    _owned_group(conn, docket_id, group_id)
    _check_facts(conn, docket_id, ids)
    _check_not_grouped(conn, ids, except_group=group_id)
    conn.execute("DELETE FROM value_group_members WHERE group_id = ?", (group_id,))
    conn.executemany(
        "INSERT INTO value_group_members (group_id, fact_id) VALUES (?, ?)",
        [(group_id, fact_id) for fact_id in ids],
    )
    conn.execute("UPDATE value_groups SET updated_at = ? WHERE id = ?", (_now(), group_id))


def rename_group(conn: sqlite3.Connection, docket_id: str, group_id: str, name: str) -> None:
    name = _validate_name(name)
    _owned_group(conn, docket_id, group_id)
    try:
        conn.execute(
            "UPDATE value_groups SET name = ?, updated_at = ? WHERE id = ?", (name, _now(), group_id)
        )
    except sqlite3.IntegrityError as exc:
        raise GroupError(f"A group named “{name}” already exists.") from exc


def merge_groups(conn: sqlite3.Connection, docket_id: str, target_id: str, source_id: str) -> None:
    """Move every record from source into target, then delete source. No record is lost."""
    if target_id == source_id:
        raise GroupError("Choose two different groups to merge.")
    _owned_group(conn, docket_id, target_id)
    _owned_group(conn, docket_id, source_id)
    conn.execute("UPDATE value_group_members SET group_id = ? WHERE group_id = ?",
                 (target_id, source_id))
    conn.execute("DELETE FROM value_groups WHERE id = ?", (source_id,))
    conn.execute("UPDATE value_groups SET updated_at = ? WHERE id = ?", (_now(), target_id))


def delete_group(conn: sqlite3.Connection, docket_id: str, group_id: str) -> None:
    """Delete a group. Its records are ungrouped, not deleted, and reappear in their documents."""
    _owned_group(conn, docket_id, group_id)
    conn.execute("DELETE FROM value_groups WHERE id = ?", (group_id,))


def _owned_group(conn: sqlite3.Connection, docket_id: str, group_id: str) -> None:
    row = conn.execute("SELECT 1 FROM value_groups WHERE id = ? AND docket_id = ?",
                       (group_id, docket_id)).fetchone()
    if row is None:
        raise GroupError("That group no longer exists. Refresh the page and try again.")


def load_groups(conn: sqlite3.Connection, docket_id: str) -> List[dict]:
    """Every group with its members and computed primary label. Read-only."""
    groups = conn.execute(
        "SELECT id, name FROM value_groups WHERE docket_id = ? ORDER BY name COLLATE NOCASE",
        (docket_id,),
    ).fetchall()
    result = []
    for group in groups:
        rows = conn.execute(
            """SELECT f.id AS fact_id, f.label, f.value, f.confidence_score, d.filename AS document_name
               FROM value_group_members m
               JOIN facts f ON f.id = m.fact_id
               JOIN documents d ON d.id = f.document_id
               WHERE m.group_id = ? AND f.review_status != 'rejected'
               ORDER BY d.filename, f.page_number, f.created_at""",
            (group["id"],),
        ).fetchall()
        members = [
            GroupMember(r["fact_id"], r["label"], r["value"], r["document_name"],
                        int(r["confidence_score"]))
            for r in rows
        ]
        primary = primary_member(members)
        result.append({
            "id": group["id"],
            "name": group["name"],
            "members": members,
            "primary_label": primary.label if primary else "",
        })
    return result


def ungrouped_fact_ids(conn: sqlite3.Connection, docket_id: str) -> set:
    rows = conn.execute(
        """SELECT f.id FROM facts f
           WHERE f.docket_id = ? AND f.review_status != 'rejected'
             AND f.concept != 'page_processing_status'
             AND f.id NOT IN (SELECT fact_id FROM value_group_members)""",
        (docket_id,),
    ).fetchall()
    return {r[0] for r in rows}