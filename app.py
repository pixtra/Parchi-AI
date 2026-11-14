"""Parchi AI: supplier document extraction, review and verification (Streamlit)."""
from __future__ import annotations
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
import csv
import functools
import hashlib
import html
import io
import json
import logging
import os
import re
import shutil
import sqlite3
import tempfile
import time
import threading
import traceback
import uuid
from concurrent.futures import ThreadPoolExecutor
from contextlib import contextmanager
from datetime import datetime, timezone
from itertools import islice
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, Iterator, List, Optional, Tuple

import streamlit as st
import pandas as pd
from psv_workflow import band_for_score, label_value_mismatch
from metadata_workflow import (
    json_bytes,
    sheet_text,
)
try:
    import rapidfuzz
    import rapidfuzz.fuzz  # noqa: F401  (load the submodules explicitly)
    import rapidfuzz.process  # noqa: F401
except ImportError:  # pip install rapidfuzz; without it fuzzy matching falls back to difflib
    rapidfuzz = None


try:
    from PIL import Image
except ImportError:
    Image = None

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from google import genai
    from google.genai import types
except ImportError:
    genai = None
    types = None

LOG = logging.getLogger("parchai")
# One worker: dockets are extracted sequentially, never concurrently (checklist §6).
PROCESSING_EXECUTOR = ThreadPoolExecutor(max_workers=1, thread_name_prefix="parchai-extraction")
PROCESSING_LOCK = threading.RLock()
PROCESSING_JOBS: Dict[str, Dict[str, Any]] = {}
WORKER_STATE = threading.local()


# ---------------------------------------------------------------- configuration
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = Path(os.getenv("PARCHAI_DB", BASE_DIR / "parchai_v3.db"))
STORAGE_DIR = Path(os.getenv("PARCHAI_STORAGE", BASE_DIR / "parchai_storage")).resolve()
CONFIG_PATH = Path(os.getenv("PARCHAI_CONFIG", BASE_DIR / "parchai.config.json"))


def load_config() -> Dict[str, Any]:
    """Application configuration, e.g. {"gemini_api_keys": [...], "gemini_models": [...]}.

    Read fresh on every call so API keys edited in the file take effect on the next
    request without restarting the app. A malformed file is reported (never crashes
    the app) via the "_error" key, which the Settings page surfaces.
    """
    try:
        raw = CONFIG_PATH.read_text(encoding="utf-8")
    except FileNotFoundError:
        return {}
    except OSError as exc:
        return {"_error": f"Could not read {CONFIG_PATH.name} ({exc.__class__.__name__})."}
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        return {"_error": f"{CONFIG_PATH.name} is not valid JSON (line {exc.lineno}): {exc.msg}."}
    if not isinstance(data, dict):
        return {"_error": f"{CONFIG_PATH.name} must contain one JSON object."}
    return data


_STARTUP_CONFIG = load_config()
_MODEL_CONFIG = _STARTUP_CONFIG.get("gemini_models")
if isinstance(_MODEL_CONFIG, str):
    _MODEL_CONFIG = [part.strip() for part in _MODEL_CONFIG.split(",")]
GEMINI_MODELS = tuple(
    m.strip()
    for m in (
        os.getenv("PARCHAI_GEMINI_MODELS", "").split(",")
        if os.getenv("PARCHAI_GEMINI_MODELS", "").strip()
        else (
            _MODEL_CONFIG
            if isinstance(_MODEL_CONFIG, list) and any(str(m).strip() for m in _MODEL_CONFIG)
            else "gemini-3.8-flash,gemini-3.6-flash,gemini-3.1-flash-lite".split(",")
        )
    )
    if str(m).strip()
)
GEMINI_TIMEOUT_MS = 180_000
MAX_OUTPUT_TOKENS = int(os.getenv("PARCHAI_MAX_OUTPUT_TOKENS", "32768"))
CHUNK_PAGES = int(os.getenv("PARCHAI_CHUNK_PAGES", "10"))
TOKENS_PER_PAGE = int(os.getenv("PARCHAI_TOKENS_PER_PAGE", "1500"))
MAX_BATCH_TOKENS = int(os.getenv("PARCHAI_MAX_BATCH_TOKENS", "400000"))
MAX_PARTS_PER_BATCH = int(os.getenv("PARCHAI_MAX_PARTS_PER_BATCH", "12"))
MAX_UPLOAD_BYTES = int(os.getenv("PARCHAI_MAX_UPLOAD_MB", "50")) * 1024 * 1024
KEY_COOLDOWN_SECONDS = 60
AUTH_COOLDOWN_SECONDS = 3600
SUPPORTED_TYPES = ["pdf", "png", "jpg", "jpeg", "txt", "xlsx", "csv"]
MIME_TYPES = {
    "pdf": "application/pdf",
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "txt": "text/plain",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "csv": "text/csv",
}
BANDS = ("Green", "Yellow", "Red")
REVIEW_STATUSES = ("pending", "approved", "rejected")
# Document lifecycle: saved → queued → processing → completed / partial / failed.
# "pending" documents exist only in pre-split databases and mean "saved, not processed".
DOCUMENT_STATES = ("saved", "pending", "queued", "saving", "processing", "completed", "partial", "failed")
PROCESSABLE_STATES = ("saved", "pending", "failed", "partial")
DOCUMENT_STATE_META: Dict[str, Tuple[str, str]] = {
    "saved": ("Saved", ""),
    "pending": ("Saved", ""),
    "queued": ("Queued", "accent"),
    "saving": ("Saving…", "accent"),
    "processing": ("Processing", "accent"),
    "completed": ("Completed", "ok"),
    "partial": ("Partial", "warn"),
    "failed": ("Failed", "bad"),
}


def document_state_label(status: str) -> str:
    label, _ = DOCUMENT_STATE_META.get(status, (status.capitalize(), ""))
    return label


def document_state_badge(status: str) -> str:
    label, kind = DOCUMENT_STATE_META.get(status, (status.capitalize(), ""))
    return badge(label, kind)
PAGES = ["Dockets", "Workspace", "Review", "Organize", "Viewer", "Validation", "Export", "Settings"]
PAGE_ICONS = {
    "Dockets": "▤", "Workspace": "⬆", "Review": "✎", "Organize": "⌘", "Viewer": "▣",
    "Validation": "✔", "Export": "⇩", "Settings": "⚙",
}
IDENTITY_CONCEPTS = {
    "invoice_number", "po_number", "purchase_order_number", "challan_number",
    "delivery_challan_number", "grn_number", "gstin", "vehicle_number",
    "supplier_name", "buyer_name", "invoice_date", "total_amount", "grand_total",
}
DOCKET_TYPES = (
    "Purchase",
    "Inventory Transfer In",
    "Inventory Transfer Out",
    "Sales",
    "Miscellaneous",
    "Payment",
)

SCHEMA = """
CREATE TABLE IF NOT EXISTS dockets (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    docket_type TEXT NOT NULL DEFAULT 'Miscellaneous',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS documents (
    id TEXT PRIMARY KEY,
    docket_id TEXT NOT NULL REFERENCES dockets(id) ON DELETE CASCADE,
    filename TEXT NOT NULL,
    mime_type TEXT NOT NULL,
    sha256 TEXT NOT NULL,
    size_bytes INTEGER NOT NULL,
    storage_path TEXT NOT NULL,
    page_count INTEGER NOT NULL,
    document_type TEXT NOT NULL,
    status TEXT NOT NULL,
    extraction_mode TEXT NOT NULL,
    processed_units INTEGER NOT NULL DEFAULT 0,
    failed_units INTEGER NOT NULL DEFAULT 0,
    completeness_status TEXT NOT NULL DEFAULT 'pending',
    completeness_message TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    UNIQUE(docket_id, sha256)
);
CREATE TABLE IF NOT EXISTS pages (
    document_id TEXT NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
    page_number INTEGER NOT NULL,
    sheet_name TEXT NOT NULL DEFAULT '',
    text TEXT NOT NULL,
    readable INTEGER NOT NULL,
    PRIMARY KEY (document_id, page_number)
);
-- Deprecated: kept only for backward compatibility with older database files.
-- Main Metadata is now derived live from `facts` (see fetch_metadata) so that document
-- deletion, fact correction and rejection are always reflected consistently.
CREATE TABLE IF NOT EXISTS metadata (
    id TEXT PRIMARY KEY,
    docket_id TEXT NOT NULL REFERENCES dockets(id) ON DELETE CASCADE,
    document_name TEXT NOT NULL,
    label TEXT NOT NULL,
    value TEXT NOT NULL,
    confidence_score INTEGER NOT NULL
);
CREATE TABLE IF NOT EXISTS templates (
    id TEXT PRIMARY KEY,
    docket_id TEXT NOT NULL REFERENCES dockets(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    labels_json TEXT NOT NULL,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    UNIQUE(docket_id, name)
);
CREATE TABLE IF NOT EXISTS label_mappings (
    id TEXT PRIMARY KEY,
    docket_id TEXT NOT NULL REFERENCES dockets(id) ON DELETE CASCADE,
    template_id TEXT NOT NULL REFERENCES templates(id) ON DELETE CASCADE,
    source_label TEXT NOT NULL,
    requested_label TEXT NOT NULL,
    method TEXT NOT NULL,
    approved INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL,
    UNIQUE(docket_id, template_id, source_label)
);
CREATE TABLE IF NOT EXISTS rejected_suggestions (
    docket_id TEXT NOT NULL REFERENCES dockets(id) ON DELETE CASCADE,
    template_id TEXT NOT NULL REFERENCES templates(id) ON DELETE CASCADE,
    source_label TEXT NOT NULL,
    requested_label TEXT NOT NULL,
    created_at TEXT NOT NULL,
    PRIMARY KEY (docket_id, template_id, source_label, requested_label)
);
CREATE TABLE IF NOT EXISTS facts (
    id TEXT PRIMARY KEY,
    docket_id TEXT NOT NULL REFERENCES dockets(id) ON DELETE CASCADE,
    document_id TEXT NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
    page_number INTEGER NOT NULL,
    label TEXT NOT NULL,
    concept TEXT NOT NULL,
    value TEXT NOT NULL,
    original_value TEXT NOT NULL,
    source_evidence TEXT NOT NULL,
    confidence_band TEXT NOT NULL,
    confidence_score INTEGER NOT NULL,
    confidence_reasons TEXT NOT NULL,
    review_status TEXT NOT NULL,
    corrected INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS fact_history (
    id TEXT PRIMARY KEY,
    fact_id TEXT NOT NULL REFERENCES facts(id) ON DELETE CASCADE,
    old_value TEXT NOT NULL,
    new_value TEXT NOT NULL,
    old_status TEXT NOT NULL,
    new_status TEXT NOT NULL,
    changed_at TEXT NOT NULL,
    reason TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_facts_docket ON facts(docket_id);
CREATE INDEX IF NOT EXISTS idx_documents_docket ON documents(docket_id);
CREATE INDEX IF NOT EXISTS idx_metadata_docket ON metadata(docket_id);
"""

BATCH_PROMPT = """You are an exhaustive document transcription and extraction engine for ONE
supplier document part in this request.
Each document part is preceded by a line "=== DOCUMENT: <id> | <label> ===". Use that <id>
(for example P2) EXACTLY as the "document" value. Page numbers restart at 1 for every labelled part. Spreadsheet
parts contain every worksheet, including hidden worksheets, represented as a page with a
sheet name. Account for every supplied page or worksheet in the response.
Return ONLY a JSON object of this exact shape:
{"documents": [
   {"document": "<label exactly as given>", "document_type": "short type name",
    "facts": [{"page_number": 1, "label": "Human readable field name",
               "concept": "snake_case_name", "value": "value as written",
               "evidence": "text copied verbatim from that page or sheet (max 500 chars)",
               "confidence_score": 87}]}
 ],
 "cross_references": [
   {"from_document": "<label>", "to_document": "<label>", "page_number": 1,
    "concept": "references_invoice", "value": "referenced number",
    "evidence": "verbatim text showing the link"}
 ]}
Rules:
- Read every supplied page and worksheet completely before responding.
- Extract every distinct visible occurrence as a separate fact, including repeated or
  conflicting labels and values. Never deduplicate, merge, overwrite, summarize, or discard.
- Capture printed text, legible handwriting, headers, footers, stamps, seals, signatures,
  notes, tables, every line item, small print and supplier-specific labels.
- Distinguish the document title or heading from actual metadata fields. Never emit a fact
  where a title such as TAX INVOICE or PURCHASE ORDER is both the label and the value.
- Use specific contextual labels such as Supplier Name, Invoice Number, Invoice Date,
  Gross Weight or Taxable Amount when the document establishes that meaning. Do not invent
  context when it is genuinely ambiguous.
- Copy values exactly as written. Do not normalize, calculate, correct, translate, or infer.
- Never fill a missing value from another document. Evidence must be verbatim from the same
  page or worksheet.
- Assign every fact an individual integer confidence_score from 1 through 100. Make it
  granular and evidence-based; do not reuse a small set of default scores and never use 0.
- Do not omit a page or worksheet. If it contains no identifiable value, return a
  "Page processing status" fact with an empty value and explain the limitation in evidence.
- Report links between documents (invoice to PO, challan to GRN, weighment slip to challan)
  in cross_references only when the text shows the link."""


# ---------------------------------------------------------------- utilities
_CONTROL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
_NUMERIC_RE = re.compile(r"^[+-]?[\d,]*\.?\d+$")


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def display_time(stamp: str) -> str:
    return stamp[:16].replace("T", " ") if stamp else "—"


def normalise(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip().lower()


def to_concept(label: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", (label or "").lower()).strip("_")[:80] or "other"


def extension_of(filename: str) -> str:
    return Path(filename).suffix.lower().lstrip(".")


def clean_text(value: Any) -> str:
    """Return text that is always valid UTF-8: lone surrogates and control characters are removed."""
    text = "" if value is None else str(value)
    text = text.encode("utf-8", "replace").decode("utf-8", "replace")
    return _CONTROL_RE.sub("", text)


def deep_clean(obj: Any) -> Any:
    if isinstance(obj, str):
        return clean_text(obj)
    if isinstance(obj, dict):
        return {clean_text(k): deep_clean(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [deep_clean(v) for v in obj]
    return obj


def safe_cell(value: Any) -> str:
    """Return the stored text unchanged. Exported values are never altered; the workbook
    writers store every such cell as literal text (data_type "s"), so nothing becomes a formula."""
    return clean_text(value)


def esc(value: Any) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def fmt_size(num_bytes: int) -> str:
    return f"{num_bytes / 1024:,.0f} KB" if num_bytes < 1024 * 1024 else f"{num_bytes / 1048576:,.1f} MB"


def upload_name(upload: Any) -> str:
    return str(upload["name"] if isinstance(upload, dict) else upload.name)


def upload_data(upload: Any) -> bytes:
    return bytes(upload["data"] if isinstance(upload, dict) else upload.getvalue())


def mask(key: str) -> str:
    return f"••••{key[-4:]}" if len(key) >= 8 else "••••"


def key_id(key: str) -> str:
    return hashlib.sha256(key.encode("utf-8", "replace")).hexdigest()[:12]


def go_to(page: str) -> None:
    st.session_state["nav"] = page


def queue_flash(level: str, message: str) -> None:
    st.session_state.setdefault("flash", []).append((level, message))


def render_flash() -> None:
    for level, message in st.session_state.pop("flash", []):
        if level == "success":
            st.toast(message, icon="✅")
        elif level == "error":
            st.error(message, icon="🚫")
        elif level == "warning":
            st.warning(message, icon="⚠️")
        else:
            st.info(message)


def _streamlit_version() -> Tuple[int, int]:
    try:
        major, minor = st.__version__.split(".")[:2]
        return int(major), int(minor)
    except Exception:
        return (0, 0)


_NEW_WIDTH_API = _streamlit_version() >= (1, 50)


def stretch() -> Dict[str, Any]:
    """Full-width kwargs that work on both old and new Streamlit releases."""
    return {"width": "stretch"} if _NEW_WIDTH_API else {"use_container_width": True}


def is_control_flow(exc: BaseException) -> bool:
    """Streamlit implements st.rerun()/st.stop() with exceptions that must never be swallowed."""
    return type(exc).__name__ in {"RerunException", "StopException", "ScriptControlException"}


def guarded(fn: Callable[..., Any]) -> Callable[..., Any]:
    """Wrap a widget callback so a failure becomes a visible error message, not a crash."""
    @functools.wraps(fn)
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        try:
            return fn(*args, **kwargs)
        except Exception as exc:  # noqa: BLE001
            if is_control_flow(exc):
                raise
            LOG.exception("callback %s failed", fn.__name__)
            queue_flash("error", f"Could not complete “{fn.__name__.replace('_', ' ')}”: {clean_text(exc)[:300]}")
    return wrapper


# ---------------------------------------------------------------- safe file writing
def atomic_write(target: Path, data: bytes) -> None:
    """Write bytes so that `target` is either the complete new content or untouched."""
    target.parent.mkdir(parents=True, exist_ok=True)
    partial = target.with_name(f"{target.name}.{uuid.uuid4().hex[:8]}.part")
    try:
        with open(partial, "wb") as handle:
            handle.write(data)
            handle.flush()
            os.fsync(handle.fileno())
        if partial.stat().st_size != len(data):
            raise OSError("Short write detected; the file was not saved.")
        os.replace(partial, target)
    finally:
        partial.unlink(missing_ok=True)


def safe_filename(name: str) -> str:
    base = re.sub(r"[^A-Za-z0-9._ -]", "_", Path(name).name).strip()
    return base[:150] or "document"


def store_original(docket_id: str, doc_id: str, filename: str, data: bytes) -> str:
    """Atomically store an upload and verify it by reading it back. Returns a path relative to STORAGE_DIR."""
    target = STORAGE_DIR / docket_id / doc_id / safe_filename(filename)
    atomic_write(target, data)
    if hashlib.sha256(target.read_bytes()).hexdigest() != hashlib.sha256(data).hexdigest():
        target.unlink(missing_ok=True)
        raise OSError("Stored file failed its checksum after writing; nothing was saved.")
    return str(target.relative_to(STORAGE_DIR))


def resolve_stored(relative: str) -> Path:
    path = (STORAGE_DIR / relative).resolve()
    if STORAGE_DIR not in path.parents:
        raise ValueError("Stored path is outside the storage directory.")
    return path


def resolve_stored_safe(relative: str) -> Optional[Path]:
    try:
        path = resolve_stored(relative)
    except ValueError:
        return None
    return path if path.is_file() else None


def read_verified(relative: str, sha256: str) -> Tuple[Optional[bytes], Optional[str]]:
    """Read a stored original and confirm it still matches its recorded checksum."""
    path = resolve_stored_safe(relative)
    if path is None:
        return None, "The original file is missing from storage."
    try:
        data = path.read_bytes()
    except OSError as exc:
        return None, f"The original file could not be read ({exc.__class__.__name__})."
    if hashlib.sha256(data).hexdigest() != sha256:
        return None, "The stored file no longer matches its checksum and may be corrupt."
    return data, None


def remove_stored_document(relative: str) -> None:
    try:
        path = resolve_stored(relative)
    except ValueError:
        return
    path.unlink(missing_ok=True)
    for parent in (path.parent, path.parent.parent):
        try:
            parent.rmdir()  # only removes empty folders
        except OSError:
            break


def remove_docket_storage(docket_id: str) -> None:
    shutil.rmtree(STORAGE_DIR / docket_id, ignore_errors=True)


def validate_upload(filename: str, data: bytes) -> Optional[str]:
    """Return a human-readable problem, or None when the upload looks sound."""
    extension = extension_of(filename)
    if extension not in SUPPORTED_TYPES:
        return f"Unsupported file type “.{extension or '?'}”. Use PDF, PNG, JPG, TXT, CSV or XLSX."
    if not data:
        return "The file is empty (0 bytes)."
    if len(data) > MAX_UPLOAD_BYTES:
        return f"The file is {fmt_size(len(data))}; the limit is {fmt_size(MAX_UPLOAD_BYTES)}."
    signatures = {"png": b"\x89PNG\r\n\x1a\n", "jpg": b"\xff\xd8\xff", "jpeg": b"\xff\xd8\xff"}
    if extension == "pdf" and b"%PDF-" not in data[:1024]:
        return "The content is not a PDF. The file may be corrupt or renamed."
    if extension in signatures and not data.startswith(signatures[extension]):
        return f"The content is not a valid .{extension} image. The file may be corrupt or renamed."
    if extension in {"png", "jpg", "jpeg"} and Image is not None:
        try:
            with Image.open(io.BytesIO(data)) as image:
                image.verify()
        except Exception as exc:  # noqa: BLE001
            return f"The image could not be opened and may be corrupt ({exc.__class__.__name__})."
    if extension == "txt" and b"\x00" in data:
        return "The file contains binary data and is not plain text."
    if extension == "csv" and b"\x00" in data:
        return "The file contains binary data and is not CSV text."
    if extension == "xlsx":
        try:
            import openpyxl
            workbook = openpyxl.load_workbook(io.BytesIO(data), read_only=False, data_only=False)
            if not workbook.worksheets:
                return "The workbook contains no worksheets."
        except Exception as exc:  # noqa: BLE001
            return f"The workbook could not be opened ({exc.__class__.__name__})."
    if extension == "pdf" and fitz is not None:
        try:
            with fitz.open(stream=data, filetype="pdf") as document:
                if document.needs_pass:
                    return "The PDF is password-protected. Remove the password and upload it again."
                if document.page_count == 0:
                    return "The PDF contains no pages."
        except Exception as exc:  # noqa: BLE001
            return f"The PDF could not be opened and may be corrupt ({exc.__class__.__name__})."
    return None


# ---------------------------------------------------------------- database
@contextmanager
def connect() -> Iterator[sqlite3.Connection]:
    connection = sqlite3.connect(DB_PATH, timeout=15)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    connection.execute("PRAGMA busy_timeout = 15000")
    connection.execute("PRAGMA synchronous = FULL")
    try:
        yield connection
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def _retire_legacy_table(connection: sqlite3.Connection, table: str) -> None:
    """Move a pre-migration table aside instead of dropping it, so no saved approval or
    rejection is ever erased by an upgrade. Rows without a template_id cannot be
    attributed to a template, so they are preserved for manual recovery."""
    legacy = f"{table}_legacy"
    if connection.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?", (legacy,)
    ).fetchone():
        legacy = f"{table}_legacy_{int(time.time())}"
    connection.execute(f"ALTER TABLE {table} RENAME TO {legacy}")


def _migrate_template_tables(connection: sqlite3.Connection) -> None:
    """Upgrade template tables to template_id scoping without losing data."""
    mapping_columns = {row["name"] for row in connection.execute("PRAGMA table_info(label_mappings)")}
    if mapping_columns and "template_id" not in mapping_columns:
        _retire_legacy_table(connection, "label_mappings")
        connection.executescript("""
            CREATE TABLE label_mappings (
                id TEXT PRIMARY KEY,
                docket_id TEXT NOT NULL REFERENCES dockets(id) ON DELETE CASCADE,
                template_id TEXT NOT NULL REFERENCES templates(id) ON DELETE CASCADE,
                source_label TEXT NOT NULL,
                requested_label TEXT NOT NULL,
                method TEXT NOT NULL,
                approved INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                UNIQUE(docket_id, template_id, source_label)
            );
        """)
    rejection_columns = {row["name"] for row in connection.execute("PRAGMA table_info(rejected_suggestions)")}
    if rejection_columns and "template_id" not in rejection_columns:
        _retire_legacy_table(connection, "rejected_suggestions")
        connection.executescript("""
            CREATE TABLE rejected_suggestions (
                docket_id TEXT NOT NULL REFERENCES dockets(id) ON DELETE CASCADE,
                template_id TEXT NOT NULL REFERENCES templates(id) ON DELETE CASCADE,
                source_label TEXT NOT NULL,
                requested_label TEXT NOT NULL,
                created_at TEXT NOT NULL,
                PRIMARY KEY (docket_id, template_id, source_label, requested_label)
            );
        """)


@st.cache_resource
def init_db() -> bool:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    STORAGE_DIR.mkdir(parents=True, exist_ok=True)
    with connect() as connection:
        connection.execute("PRAGMA journal_mode = WAL")
        connection.executescript(SCHEMA)
        columns = {row["name"] for row in connection.execute("PRAGMA table_info(dockets)")}
        if "docket_type" not in columns:
            connection.execute(
                "ALTER TABLE dockets ADD COLUMN docket_type TEXT NOT NULL DEFAULT 'Miscellaneous'"
            )
        document_columns = {row["name"] for row in connection.execute("PRAGMA table_info(documents)")}
        for name, definition in (
            ("processed_units", "INTEGER NOT NULL DEFAULT 0"),
            ("failed_units", "INTEGER NOT NULL DEFAULT 0"),
            ("completeness_status", "TEXT NOT NULL DEFAULT 'pending'"),
            ("completeness_message", "TEXT NOT NULL DEFAULT ''"),
        ):
            if name not in document_columns:
                connection.execute(f"ALTER TABLE documents ADD COLUMN {name} {definition}")
        page_columns = {row["name"] for row in connection.execute("PRAGMA table_info(pages)")}
        if "sheet_name" not in page_columns:
            connection.execute("ALTER TABLE pages ADD COLUMN sheet_name TEXT NOT NULL DEFAULT ''")
        _migrate_template_tables(connection)
        # A previous process may have stopped mid-run; documents stuck in an in-flight
        # status would otherwise never be processable again. Their originals and page
        # rows are intact (facts are only written atomically at finalize), so they
        # safely return to the "saved" state, ready to be queued again.
        connection.execute(
            "UPDATE documents SET status = 'saved' WHERE status IN ('queued', 'processing', 'saving')"
        )
    return True


def create_docket(name: str, docket_type: str) -> str:
    docket_id = str(uuid.uuid4())
    now = utc_now()
    clean = clean_text(name).strip() or f"Docket {now[:10]}"
    clean_type = clean_text(docket_type).strip()
    if clean_type not in DOCKET_TYPES:
        raise ValueError("Select a valid docket type.")
    with connect() as connection:
        connection.execute(
            "INSERT INTO dockets (id, name, docket_type, created_at, updated_at) VALUES (?, ?, ?, ?, ?)",
            (docket_id, clean[:120], clean_type, now, now),
        )
    return docket_id


def rename_docket(docket_id: str, name: str) -> None:
    clean = clean_text(name).strip()
    if not clean:
        raise ValueError("Docket name cannot be empty.")
    with connect() as connection:
        connection.execute(
            "UPDATE dockets SET name = ?, updated_at = ? WHERE id = ?",
            (clean[:120], utc_now(), docket_id),
        )


def delete_docket(docket_id: str) -> None:
    """Removes the docket, all its rows (cascade) and its stored original files."""
    with connect() as connection:
        connection.execute("DELETE FROM dockets WHERE id = ?", (docket_id,))
    remove_docket_storage(docket_id)


def delete_document(docket_id: str, document_id: str) -> Optional[str]:
    with connect() as connection:
        row = connection.execute(
            "SELECT storage_path FROM documents WHERE id = ? AND docket_id = ?",
            (document_id, docket_id),
        ).fetchone()
        if row is None:
            return None
        connection.execute("DELETE FROM documents WHERE id = ?", (document_id,))
        connection.execute("UPDATE dockets SET updated_at = ? WHERE id = ?", (utc_now(), docket_id))
    remove_stored_document(row["storage_path"])
    return row["storage_path"]


def list_dockets() -> List[sqlite3.Row]:
    with connect() as connection:
        return connection.execute(
            """
            SELECT d.id, d.name, d.docket_type, d.created_at, d.updated_at,
                (SELECT COUNT(*) FROM documents x WHERE x.docket_id = d.id) AS document_count,
                (SELECT COUNT(*) FROM facts f WHERE f.docket_id = d.id) AS fact_count,
                (SELECT COUNT(*) FROM facts f WHERE f.docket_id = d.id
                    AND f.review_status = 'pending') AS pending_count
            FROM dockets d ORDER BY d.updated_at DESC
            """
        ).fetchall()


def docket_status(row: sqlite3.Row) -> str:
    if row["fact_count"] == 0:
        return "Empty"
    return "Needs review" if row["pending_count"] else "Reviewed"


def list_documents(docket_id: str) -> List[sqlite3.Row]:
    with connect() as connection:
        return connection.execute(
            """SELECT id, filename, mime_type, sha256, size_bytes, storage_path, document_type, status,
                      extraction_mode, page_count, processed_units, failed_units,
                      completeness_status, completeness_message, created_at
               FROM documents WHERE docket_id = ? ORDER BY created_at""",
            (docket_id,),
        ).fetchall()


def set_document_status(document_id: str, status: str) -> None:
    with connect() as connection:
        connection.execute(
            "UPDATE documents SET status = ? WHERE id = ?",
            (status, document_id),
        )


def fetch_facts(docket_id: str) -> List[sqlite3.Row]:
    with connect() as connection:
        return connection.execute(
            """SELECT f.*, d.filename, d.storage_path, d.mime_type
               FROM facts f JOIN documents d ON d.id = f.document_id
               WHERE f.docket_id = ?
               ORDER BY d.filename, f.page_number, f.created_at""",
            (docket_id,),
        ).fetchall()


def fetch_metadata(docket_id: str) -> List[sqlite3.Row]:
    with connect() as connection:
        return connection.execute(
            """SELECT d.filename AS document_name, f.label AS label, f.value AS value,
                      f.confidence_score AS confidence_score
               FROM facts f JOIN documents d ON d.id = f.document_id
               WHERE f.docket_id = ? AND f.review_status != 'rejected'
                 AND f.concept != 'page_processing_status'
               ORDER BY d.filename, f.page_number, f.created_at""",
            (docket_id,),
        ).fetchall()


def list_templates(docket_id: Optional[str] = None) -> List[sqlite3.Row]:
    """Every saved template in the workspace. Templates are reusable client templates,
    so they are NOT limited to the docket where they were created; `docket_id` is kept
    only for call-site compatibility and records nothing here."""
    with connect() as connection:
        return connection.execute(
            "SELECT * FROM templates ORDER BY name COLLATE NOCASE"
        ).fetchall()


def save_template(docket_id: str, name: str, labels: List[str], template_id: Optional[str] = None) -> str:
    """Create a template, or update one in place when `template_id` is passed.

    - Labels are deduplicated by comparison form while keeping the first spelling.
    - Creating refuses an already-used name (edit that template instead).
    - Updating keeps kept labels verbatim, drops deselected ones, and prunes
      approved mappings pointing at labels the template no longer offers, so a
      template can never export under a label it does not list.
    Returns the template id. Never touches Main Metadata."""
    clean_name = clean_text(name).strip()
    clean_labels = _dedupe_labels(labels)
    if not clean_name or not clean_labels:
        raise ValueError("A template name and at least one label are required.")
    now = utc_now()
    removed: List[str] = []
    resolved_id: str
    with connect() as connection:
        existing = None
        if template_id:
            existing = connection.execute(
                "SELECT id, labels_json FROM templates WHERE id = ?",
                (template_id,),
            ).fetchone()
            if existing is None:
                raise ValueError("That template no longer exists.")
        if existing:
            try:
                connection.execute(
                    "UPDATE templates SET name = ?, labels_json = ?, updated_at = ? WHERE id = ?",
                    (clean_name, json.dumps(clean_labels, ensure_ascii=False), now, existing["id"]),
                )
            except sqlite3.IntegrityError as exc:
                raise ValueError(f"A template named “{clean_name}” already exists.") from exc
            resolved_id = str(existing["id"])
            kept = {normalize_label(label) for label in clean_labels}
            removed = [
                label
                for label in json.loads(existing["labels_json"])
                if normalize_label(label) not in kept
            ]
        else:
            clash = connection.execute(
                "SELECT id FROM templates WHERE name = ?",
                (clean_name,),
            ).fetchone()
            if clash:
                raise ValueError(
                    f"A template named “{clean_name}” already exists. Choose another name, "
                    "or edit that template instead."
                )
            resolved_id = str(uuid.uuid4())
            connection.execute(
                """INSERT INTO templates (id, docket_id, name, labels_json, created_at, updated_at)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (resolved_id, docket_id, clean_name, json.dumps(clean_labels, ensure_ascii=False), now, now),
            )
    if removed:
        delete_mappings_to_labels(docket_id, resolved_id, removed)
    return resolved_id


def delete_template(docket_id: str, template_id: str) -> None:
    """Delete a template; its mappings and rejected suggestions cascade with it."""
    with connect() as connection:
        connection.execute(
            "DELETE FROM templates WHERE id = ?", (template_id,)
        )


def append_template_labels(docket_id: str, template_id: str, labels: List[str]) -> List[str]:
    """Explicitly add labels to a template (the "Add to template" action in the apply
    workflow). Existing labels and their spelling are never altered; labels already
    present in any casing are skipped. Returns the labels actually added."""
    added: List[str] = []
    with connect() as connection:
        row = connection.execute(
            "SELECT labels_json FROM templates WHERE id = ?",
            (template_id,),
        ).fetchone()
        if row is None:
            raise ValueError("That template no longer exists.")
        current_labels = json.loads(row["labels_json"])
        seen = {normalize_label(label) for label in current_labels}
        for label in labels:
            clean = clean_text(label).strip()
            normalized = normalize_label(clean)
            if normalized and normalized not in seen:
                current_labels.append(clean)
                seen.add(normalized)
                added.append(clean)
        connection.execute(
            "UPDATE templates SET labels_json = ?, updated_at = ? WHERE id = ?",
            (json.dumps(current_labels, ensure_ascii=False), utc_now(), template_id),
        )
    return added


def save_mapping(docket_id: str, template_id: str, source_label: str, requested_label: str, method: str) -> None:
    """Approve a mapping from an extracted label to one of THIS template's labels.
    A source maps to exactly one target per template; re-approving replaces it."""
    source_label, requested_label = clean_text(source_label).strip(), clean_text(requested_label).strip()
    if not source_label or not requested_label:
        return
    with connect() as connection:
        connection.execute(
            """INSERT INTO label_mappings (id, docket_id, template_id, source_label, requested_label, method, approved, created_at)
               VALUES (?, ?, ?, ?, ?, ?, 1, ?)
               ON CONFLICT(docket_id, template_id, source_label)
               DO UPDATE SET requested_label = excluded.requested_label, method = excluded.method, approved = 1""",
            (str(uuid.uuid4()), docket_id, template_id, source_label, requested_label, method, utc_now()),
        )


def approved_mappings(docket_id: str, template_id: str) -> Dict[str, str]:
    """Maps each source label to the one target approved WITHIN this template."""
    with connect() as connection:
        rows = connection.execute(
            """SELECT source_label, requested_label FROM label_mappings
               WHERE docket_id = ? AND template_id = ? AND approved = 1""",
            (docket_id, template_id),
        ).fetchall()
    return {row["source_label"]: row["requested_label"] for row in rows}


def mapping_details(docket_id: str, template_id: str) -> Dict[str, Dict[str, str]]:
    """Approved mappings with their method: {source: {"target": ..., "method": ...}}."""
    with connect() as connection:
        rows = connection.execute(
            """SELECT source_label, requested_label, method FROM label_mappings
               WHERE docket_id = ? AND template_id = ? AND approved = 1""",
            (docket_id, template_id),
        ).fetchall()
    return {
        row["source_label"]: {"target": row["requested_label"], "method": row["method"]}
        for row in rows
    }


def delete_mappings_to_labels(docket_id: str, template_id: str, labels: Iterable[str]) -> None:
    """Remove mappings pointing at labels this template no longer offers, in EVERY docket
    that uses the template, since the template itself is workspace-wide."""
    labels = [str(label) for label in labels if str(label).strip()]
    if not labels:
        return
    with connect() as connection:
        connection.executemany(
            "DELETE FROM label_mappings WHERE template_id = ? AND requested_label = ?",
            [(template_id, label) for label in labels],
        )


def delete_mapping(docket_id: str, template_id: str, source_label: str, requested_label: str) -> None:
    """Remove one approved mapping (Reject-after-approval or explicit Unmap)."""
    with connect() as connection:
        connection.execute(
            """DELETE FROM label_mappings
               WHERE docket_id = ? AND template_id = ? AND source_label = ? AND requested_label = ?""",
            (docket_id, template_id, clean_text(source_label).strip(), clean_text(requested_label).strip()),
        )


def reject_suggestion(docket_id: str, template_id: str, source_label: str, requested_label: str) -> None:
    # Normalized the same way save_mapping/delete_mapping store and look up labels,
    # so a rejection always matches on later lookups regardless of incidental whitespace.
    source_label, requested_label = clean_text(source_label).strip(), clean_text(requested_label).strip()
    if not source_label or not requested_label:
        return
    with connect() as connection:
        connection.execute(
            """INSERT OR IGNORE INTO rejected_suggestions
               (docket_id, template_id, source_label, requested_label, created_at)
               VALUES (?, ?, ?, ?, ?)""",
            (docket_id, template_id, source_label, requested_label, utc_now()),
        )


def rejected_suggestions(docket_id: str, template_id: str) -> set:
    with connect() as connection:
        rows = connection.execute(
            """SELECT source_label, requested_label FROM rejected_suggestions
               WHERE docket_id = ? AND template_id = ?""",
            (docket_id, template_id),
        ).fetchall()
    return {(row["source_label"], row["requested_label"]) for row in rows}


def fact_history(fact_id: str) -> List[sqlite3.Row]:
    with connect() as connection:
        return connection.execute(
            "SELECT * FROM fact_history WHERE fact_id = ? ORDER BY changed_at DESC", (fact_id,)
        ).fetchall()


# ---------------------------------------------------------------- reading documents
def decode_text(data: bytes) -> str:
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", "replace")


def read_pages(filename: str, data: bytes) -> List[Dict[str, Any]]:
    extension = extension_of(filename)
    if extension == "pdf":
        if fitz is None:
            raise RuntimeError("PyMuPDF is not installed (pip install pymupdf).")
        pages = []
        with fitz.open(stream=data, filetype="pdf") as document:
            for index, page in enumerate(document, start=1):
                text = clean_text(page.get_text("text")).strip()
                pages.append({"page_number": index, "sheet_name": "", "text": text, "readable": bool(text)})
        if not pages:
            raise RuntimeError("The PDF contains no pages.")
        return pages
    if extension == "txt":
        text = clean_text(decode_text(data)).strip()
        return [{"page_number": 1, "sheet_name": "", "text": text, "readable": bool(text)}]
    if extension == "xlsx":
        import openpyxl
        workbook = openpyxl.load_workbook(io.BytesIO(data), read_only=False, data_only=False)
        return sheet_text(workbook)
    if extension == "csv":
        text = clean_text(decode_text(data)).strip()
        return [{"page_number": 1, "sheet_name": "CSV", "text": text, "readable": bool(text)}]
    return [{"page_number": 1, "sheet_name": "", "text": "", "readable": False}]  # images: vision-only


def split_pdf(data: bytes, pages_per_part: int) -> List[Tuple[int, int, bytes]]:
    with fitz.open(stream=data, filetype="pdf") as source:
        total = source.page_count
        if total <= pages_per_part:
            return [(0, total, data)]
        parts = []
        for start in range(0, total, pages_per_part):
            end = min(start + pages_per_part, total)
            with fitz.open() as part:
                part.insert_pdf(source, from_page=start, to_page=end - 1)
                parts.append((start, end - start, part.tobytes()))
        return parts


# ---------------------------------------------------------------- confidence
def band_for(score: int) -> str:
    return band_for_score(score)


def clamp_confidence(score: Any, default: int = 50) -> int:
    """Scores are always integers within 0–100 (0 allowed; the model guidance avoids it)."""
    try:
        return max(0, min(100, round(float(score))))
    except (TypeError, ValueError):
        return default


DOCUMENT_TITLE_WORDS = {
    "tax invoice", "invoice", "purchase order", "delivery challan",
    "goods receipt note", "grn", "weighment slip", "credit note",
    "debit note", "quotation", "receipt", "transport receipt",
}


def _normalised_phrase(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").lower()).strip()


def is_document_title_fact(
    label: str, value: str, document_type: str = "", page_text: str = ""
) -> bool:
    label_text = _normalised_phrase(label)
    value_text = _normalised_phrase(value)
    if not value_text or label_text != value_text:
        return False
    title = value_text
    known_type = _normalised_phrase(document_type)
    if title in DOCUMENT_TITLE_WORDS or title == known_type:
        return True
    first_lines = {
        _normalised_phrase(line)
        for line in page_text.splitlines()[:5]
        if line.strip()
    }
    return title in first_lines and len(title.split()) <= 6


def contextual_label(label: str, document_type: str = "", page_text: str = "") -> str:
    original = clean_text(label).strip()
    generic = _normalised_phrase(original)
    context = f"{_normalised_phrase(document_type)} {_normalised_phrase(page_text[:2500])}"
    replacements = {
        "name": (
            ("supplier", "Supplier Name"), ("vendor", "Supplier Name"),
            ("buyer", "Buyer Name"), ("consignee", "Consignee Name"),
            ("transporter", "Transporter Name"), ("driver", "Driver Name"),
        ),
        "number": (
            ("invoice", "Invoice Number"), ("purchase order", "Purchase Order Number"),
            ("vehicle", "Vehicle Registration Number"), ("challan", "Delivery Challan Number"),
            ("grn", "Goods Receipt Note Number"),
        ),
        "date": (
            ("invoice", "Invoice Date"), ("delivery", "Delivery Date"),
            ("purchase order", "Purchase Order Date"), ("due", "Due Date"),
        ),
        "amount": (
            ("grand total", "Invoice Total Amount"), ("total", "Invoice Total Amount"),
            ("taxable", "Taxable Amount"), ("balance", "Outstanding Balance"),
            ("tax amount", "Tax Amount"),
        ),
        "weight": (
            ("gross", "Gross Weight"), ("tare", "Tare Weight"), ("net", "Net Weight"),
        ),
        "quantity": (
            ("ordered", "Ordered Quantity"), ("received", "Received Quantity"),
            ("delivered", "Delivered Quantity"),
        ),
        "address": (
            ("supplier", "Supplier Address"), ("delivery", "Delivery Address"),
            ("billing", "Billing Address"),
        ),
    }
    for token, specific in replacements.get(generic, ()):
        if token in context:
            return specific
    return original


def calibrated_confidence(
    requested: Any,
    value: str,
    evidence: str,
    page_text: Optional[str],
    label: str,
) -> int:
    base = clamp_confidence(requested, 50)
    evidence_len = len(evidence.strip())
    value_len = len(value.strip())
    if page_text is None:
        return max(1, min(100, round(base * 0.78 + min(18, evidence_len / 20))))
    source = normalise(page_text)
    evidence_match = bool(evidence.strip()) and normalise(evidence) in source
    value_match = bool(value.strip()) and normalise(value) in source
    adjustment = 0
    if evidence_match:
        adjustment += min(12, 4 + evidence_len // 35)
    elif value_match:
        adjustment -= 8
    else:
        adjustment -= 24
    if value_len <= 1:
        adjustment -= 10
    if re.search(r"[?*_]{2,}|…", value):
        adjustment -= 12
    if _normalised_phrase(label) in {"name", "number", "date", "amount", "weight", "quantity"}:
        adjustment -= 3
    return max(1, min(100, base + adjustment))


def assess(value: str, evidence: str, page_text: Optional[str]) -> Dict[str, Any]:
    """Confidence reflects evidence quality. page_text is None when the page has no text
    layer, so the value cannot be checked against the source."""
    if not value.strip():
        return {"band": "Red", "score": 1, "reasons": ["MISSING"]}
    if not evidence.strip():
        score = min(45, 18 + len(value.strip()) * 2)
        return {"band": band_for(score), "score": score, "reasons": ["NO_EVIDENCE"]}
    if page_text is None:
        score = min(75, 46 + len(evidence.strip()) // 20)
        return {"band": band_for(score), "score": score, "reasons": ["NOT_TEXT_VERIFIABLE"]}
    text_n = normalise(page_text)
    if normalise(evidence) in text_n:
        score, reasons = min(100, 82 + len(evidence.strip()) // 25), ["GROUNDED"]
    elif normalise(value) and normalise(value) in text_n:
        score, reasons = min(79, 58 + len(value.strip()) // 10), ["VALUE_FOUND_EVIDENCE_MISMATCH"]
    else:
        score = min(55, 15 + len(value.strip()) * 2)
        return {"band": band_for(score), "score": score, "reasons": ["UNSUPPORTED"]}
    if len(value.strip()) < 2:
        score = min(score, 74)
        reasons.append("VERY_SHORT_VALUE")
    return {"band": band_for(score), "score": score, "reasons": reasons}


# ---------------------------------------------------------------- local extraction
def _fact(
    page: int,
    label: str,
    concept: str,
    value: str,
    evidence: str,
    confidence_score: Any = None,
) -> Dict[str, Any]:
    return {
        "page_number": page,
        "label": label[:120],
        "concept": concept,
        "value": value,
        "evidence": evidence,
        "confidence_score": clamp_confidence(confidence_score) if confidence_score is not None else None,
    }


# ---------------------------------------------------------------- API keys and errors
# Gemini keys are opaque ASCII tokens. Current keys use the "AQ." prefix and may contain
# periods in the token body; older AI Studio keys commonly use the "AIza" prefix.
KEY_RE = re.compile(r"(?:AQ\.[A-Za-z0-9._~-]{16,}|AIza[A-Za-z0-9._~-]{16,})")


class _SecretRedactionFilter(logging.Filter):
    """Never let an API key reach a log sink, even inside an exception message."""

    def filter(self, record: logging.LogRecord) -> bool:
        try:
            message = record.getMessage()
            redacted = KEY_RE.sub("[redacted-key]", message)
            if redacted != message:
                record.msg, record.args = redacted, ()
        except Exception:  # noqa: BLE001 - logging must never break the app
            pass
        return True


def install_secret_redaction() -> None:
    """Attach the redaction filter where it actually takes effect: on the logger that
    emits our records and on every root handler (ancestor logger filters do not
    apply to propagated records, so handler-level filtering is required)."""
    redaction = _SecretRedactionFilter()
    LOG.addFilter(redaction)
    for handler in logging.getLogger().handlers:
        handler.addFilter(redaction)


install_secret_redaction()


class InvalidKeyError(Exception):
    """The API key is not shaped like a Gemini key."""


def valid_key(key: str) -> bool:
    return bool(key) and len(key) <= 256 and key.isascii() and bool(KEY_RE.fullmatch(key))





def make_client(api_key: str) -> Any:
    if genai is None or types is None:
        raise RuntimeError("google-genai is not installed (pip install google-genai).")
    if not valid_key(api_key):
        raise InvalidKeyError("API key has an invalid format (expected a Gemini key beginning with “AQ.” or “AIza”).")
    return genai.Client(api_key=api_key, http_options=types.HttpOptions(timeout=GEMINI_TIMEOUT_MS))


def scrub(exc: Exception, secret: str) -> str:
    text = str(exc)
    if secret:
        text = text.replace(secret, "[redacted]")
    return clean_text(text)[:300]


def classify_error(exc: Exception) -> Optional[int]:
    if isinstance(exc, InvalidKeyError):
        return AUTH_COOLDOWN_SECONDS
    text = str(exc).upper()
    if "RESOURCE_EXHAUSTED" in text or "429" in text or "QUOTA" in text:
        return KEY_COOLDOWN_SECONDS
    if "API_KEY_INVALID" in text or "PERMISSION_DENIED" in text or "401" in text or "403" in text:
        return AUTH_COOLDOWN_SECONDS
    return None


def is_size_error(exc: Exception) -> bool:
    text = str(exc).lower()
    if "quota" in text or "rate limit" in text or "per minute" in text or "per day" in text:
        return False  # a quota/rate-limit error, not an oversized request
    return "too large" in text or ("token" in text and ("exceed" in text or "limit" in text))


# Shared by the extraction worker and the UI (one store), so a key that hit its quota
# is skipped by the next run, and the cooling chip and "Clear cooldowns" see the same state.
KEY_COOLDOWNS: Dict[str, float] = {}


def cooldown_remaining(cooldown_id: str) -> float:
    with PROCESSING_LOCK:
        until = KEY_COOLDOWNS.get(cooldown_id, 0.0)
    return max(0.0, until - time.time())


def set_cooldown(cooldown_id: str, seconds: int) -> None:
    with PROCESSING_LOCK:
        KEY_COOLDOWNS[cooldown_id] = time.time() + seconds


def active_cooldown_count() -> int:
    now = time.time()
    with PROCESSING_LOCK:
        return sum(1 for until in KEY_COOLDOWNS.values() if until > now)


def clear_cooldowns() -> None:
    with PROCESSING_LOCK:
        KEY_COOLDOWNS.clear()


def _config_gemini_keys() -> List[Tuple[str, str]]:
    """Keys from the application configuration file (parchai.config.json).

    Supports "gemini_api_keys": ["key1", "key2"] and a single "gemini_api_key": "…".
    Loaded fresh on every pool lookup, so editing the file needs no app restart.
    """
    config = load_config()
    found: List[Tuple[str, str]] = []
    raw_list = config.get("gemini_api_keys", [])
    if isinstance(raw_list, str):
        raw_list = re.split(r"[\n,;]+", raw_list)
    if isinstance(raw_list, list):
        found.extend((str(item).strip(), "config file") for item in raw_list)
    single = str(config.get("gemini_api_key", "") or "").strip()
    if single:
        found.append((single, "config file"))
    return found


def _key_candidates() -> Iterator[Tuple[str, str]]:
    for raw in re.split(r"[\n,;]+", os.getenv("GEMINI_API_KEYS", "")):
        yield raw.strip(), "env pool"
    yield os.getenv("GEMINI_API_KEY", "").strip(), "env"
    for raw, source in _config_gemini_keys():
        yield raw, source


def key_pool() -> List[Dict[str, str]]:
    entries: List[Dict[str, str]] = []
    seen: set = set()
    for key, source in _key_candidates():
        if not key or key in seen or not valid_key(key):
            continue
        seen.add(key)
        entries.append({"key": key, "source": source, "id": key_id(key), "masked": mask(key)})
    for number, entry in enumerate(entries, start=1):
        entry["label"] = f"Key {number}"
    return entries


def rejected_keys() -> List[Dict[str, str]]:
    return [{"source": source, "masked": mask(key)}
            for key, source in _key_candidates() if key and not valid_key(key)]


# Round-robin cursor shared by every worker thread. Each new batch starts on the key
# after the one that served the previous successful request, spreading load across the
# pool instead of hammering "Key 1" until it exhausts its quota.
_KEY_ROTATION_LOCK = threading.Lock()
_KEY_ROTATION_INDEX = 0


def rotated_keys(keys: List[Dict[str, str]]) -> List[Dict[str, str]]:
    """Pool order for the next batch: starts after the last successful key."""
    with _KEY_ROTATION_LOCK:
        start = _KEY_ROTATION_INDEX % len(keys) if keys else 0
    return keys[start:] + keys[:start]


def record_key_success(keys: List[Dict[str, str]], used_id: str) -> None:
    """Advance the rotation cursor past the key that just succeeded."""
    global _KEY_ROTATION_INDEX
    ids = [entry["id"] for entry in keys]
    if used_id in ids:
        with _KEY_ROTATION_LOCK:
            _KEY_ROTATION_INDEX = (ids.index(used_id) + 1) % max(1, len(ids))


def earliest_cooldown(keys: List[Dict[str, str]]) -> float:
    """Seconds until the soonest-to-recover pooled key comes out of cooldown."""
    waits: List[float] = []
    for entry in keys:
        wait = cooldown_remaining(entry["id"])
        model_waits = [cooldown_remaining(f"{entry['id']}:{model}") for model in GEMINI_MODELS]
        recover_at = max([wait, *model_waits], default=0.0)
        if recover_at > 0:
            waits.append(recover_at)
    return min(waits) if waits else 0.0


def finish_reason(response: Any) -> str:
    candidates = getattr(response, "candidates", None) or []
    reason = getattr(candidates[0], "finish_reason", None) if candidates else None
    return str(getattr(reason, "name", reason) or "").upper()


def check_key(key: str) -> str:
    try:
        models = list(islice(make_client(key).models.list(), 1))
        return "Key accepted." if models else "Key accepted (no models listed)."
    except Exception as exc:  # noqa: BLE001
        return f"Failed: {scrub(exc, key)}"


def list_models(key: str) -> List[str]:
    names = []
    for model in make_client(key).models.list():
        actions = getattr(model, "supported_actions", None) or []
        if not actions or "generateContent" in actions:
            names.append(model.name.removeprefix("models/"))
    return sorted(names)


# ---------------------------------------------------------------- batching
class OversizedBatch(Exception):
    """Output was truncated or input was too large; the batch must be split."""


def make_units(doc: Dict[str, Any], label: Optional[str] = None) -> List[Dict[str, Any]]:
    if extension_of(doc["filename"]) == "pdf" and fitz is not None:
        parts = split_pdf(doc["data"], CHUNK_PAGES)
        part_meta = [(start, count, blob, doc["mime"], "." + extension_of(doc["filename"]),
                      None) for start, count, blob in parts]
    elif extension_of(doc["filename"]) == "xlsx":
        part_meta = [
            (
                page["page_number"] - 1,
                1,
                page["text"].encode("utf-8"),
                "text/plain",
                ".txt",
                page.get("sheet_name"),
            )
            for page in doc["pages"]
        ]
    else:
        part_meta = [(0, doc["page_count"], doc["data"], doc["mime"],
                      "." + extension_of(doc["filename"]), None)]
    units = []
    for start, count, blob, mime, suffix, sheet_name in part_meta:
        unit_label = label or doc["filename"]
        if extension_of(doc["filename"]) == "xlsx":
            unit_label = f"{unit_label} [sheet: {sheet_name}]"
        elif len(part_meta) > 1:
            unit_label = f"{unit_label} [pages {start + 1}-{start + count}]"
        unit = {
            "label": unit_label, "doc": doc, "offset": start, "pages": count, "data": blob,
            "mime": mime, "suffix": suffix,
            "pages_slice": doc["pages"][start:start + count],
            "items": None, "status": None, "message": "", "doc_type": None,
        }
        doc["units"].append(unit)
        units.append(unit)
    return units


def pack_batches(units: List[Dict[str, Any]]) -> List[List[Dict[str, Any]]]:
    batches: List[List[Dict[str, Any]]] = []
    current: List[Dict[str, Any]] = []
    tokens = 0
    for unit in units:
        cost = unit["pages"] * TOKENS_PER_PAGE
        if current and (tokens + cost > MAX_BATCH_TOKENS or len(current) >= MAX_PARTS_PER_BATCH):
            batches.append(current)
            current, tokens = [], 0
        current.append(unit)
        tokens += cost
    if current:
        batches.append(current)
    return batches


def delete_handles(client: Any, handles: List[Any]) -> None:
    for handle in handles:
        try:
            client.files.delete(name=handle.name)
        except Exception:  # noqa: BLE001
            pass


def wait_until_active(client: Any, uploaded: Any, timeout: float = 120.0) -> Any:
    """Gemini processes uploads asynchronously; a file must reach ACTIVE before it is used."""
    deadline = time.time() + timeout
    current = uploaded
    while True:
        state_obj = getattr(current, "state", None)
        if state_obj is None:
            return current  # SDK did not report a state; nothing to wait for
        state = str(getattr(state_obj, "name", state_obj)).upper()
        if state == "ACTIVE":
            return current
        if state == "FAILED":
            raise RuntimeError(f"Gemini could not process the uploaded file {current.name}.")
        if time.time() > deadline:
            raise RuntimeError("Gemini took too long to prepare the uploaded file.")
        time.sleep(2)
        current = client.files.get(name=current.name)


def upload_parts(entry: Dict[str, str], batch: List[Dict[str, Any]]) -> Tuple[Any, List[Any]]:
    client = make_client(entry["key"])
    handles: List[Any] = []
    try:
        for unit in batch:
            with tempfile.NamedTemporaryFile(suffix=unit["suffix"], delete=False) as handle:
                handle.write(unit["data"])
                temporary = Path(handle.name)
            try:
                uploaded = client.files.upload(file=str(temporary), config={"mime_type": unit["mime"]})
                handles.append(uploaded)  # tracked first so cleanup still deletes it on failure
                handles[-1] = wait_until_active(client, uploaded)
            finally:
                temporary.unlink(missing_ok=True)
    except Exception:
        delete_handles(client, handles)
        raise
    return client, handles


def parse_batch_payload(text: str, labels: List[str]) -> Tuple[Dict[str, Dict[str, Any]], List[Dict[str, Any]]]:
    cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", (text or "").strip(), flags=re.IGNORECASE)
    payload = json.loads(cleaned)
    if not isinstance(payload, dict) or not isinstance(payload.get("documents"), list):
        raise ValueError("response is missing the documents list")
    ids = {f"p{number}": label for number, label in enumerate(labels, start=1)}
    by_label_lower = {label.lower(): label for label in labels}
    by_base: Dict[str, List[str]] = {}
    for label in labels:
        base = re.sub(r"\s*\[(?:pages|sheet)[^\]]*\]\s*$", "", label, flags=re.IGNORECASE)
        by_base.setdefault(base.strip().lower(), []).append(label)

    def match(name: str) -> Optional[str]:
        name = re.sub(r"^\s*P\d+\s*\|\s*", "", name or "", flags=re.IGNORECASE).strip()
        if name in labels:
            return name
        key = name.lower()
        if key in ids:
            return ids[key]
        if key in by_label_lower:
            return by_label_lower[key]
        candidates = by_base.get(key, [])
        return candidates[0] if len(candidates) == 1 else None  # never guess between chunks

    def page_of(item: Dict[str, Any]) -> Tuple[int, bool]:
        try:
            number = int(item.get("page_number") or 1)
            return max(1, number), number >= 1
        except (TypeError, ValueError):
            return 1, False

    by_label: Dict[str, Dict[str, Any]] = {}
    for doc in payload["documents"]:
        if not isinstance(doc, dict):
            continue
        label = match(str(doc.get("document", "")))
        if label is None:
            continue
        bucket = by_label.setdefault(label, {"document_type": str(doc.get("document_type") or ""), "facts": []})
        document_type = bucket["document_type"]
        for item in doc.get("facts") or []:
            if not isinstance(item, dict):
                continue
            name = str(item.get("label") or "Unlabelled field").strip()
            value = str(item.get("value") or "").strip()
            if is_document_title_fact(name, value, document_type):
                continue
            name = contextual_label(name, document_type)
            page_number, page_valid = page_of(item)
            fact = _fact(
                page_number, name, to_concept(name),
                value,
                str(item.get("evidence") or "").strip()[:500],
                item.get("confidence_score"),
            )
            fact["_page_number_valid"] = page_valid
            bucket["facts"].append(fact)

    references = []
    for ref in payload.get("cross_references") or []:
        if not isinstance(ref, dict):
            continue
        source = match(str(ref.get("from_document", "")))
        target = str(ref.get("to_document", "")).strip()
        if source is None or not target:
            continue
        page_number, page_valid = page_of(ref)
        references.append({
            "from_label": source,
            "page_number": page_number,
            "_page_number_valid": page_valid,
            "label": f"Reference to {target}",
            "concept": f"cross_reference_{to_concept(str(ref.get('concept') or 'reference'))}",
            "value": str(ref.get("value") or "").strip(),
            "evidence": str(ref.get("evidence") or "").strip()[:500],
        })
    return by_label, references


MAX_COOLDOWN_WAIT = 75.0  # at most one per-minute quota window when the whole pool is busy
MAX_POOL_PASSES = 3       # bounded so a genuinely dead pool never loops forever


def call_gemini_batch(batch: List[Dict[str, Any]], keys: List[Dict[str, str]]) -> Tuple[Any, str]:
    """Run one batched extraction request against the pooled keys.

    Order of attempt: models x keys, with keys starting from the round-robin cursor.
    A key (or key+model) in cooldown is skipped; if every candidate is cooling down,
    wait for the soonest recovery and retry (bounded), so a burst of 429s rotates
    through the pool instead of failing the batch. Invalid keys are parked for an
    hour; quota errors park a key+model for a minute. Key material is never logged.
    """
    if genai is None or types is None:
        raise RuntimeError("google-genai is not installed (pip install google-genai).")
    if not keys:
        raise RuntimeError("No valid Gemini API keys are configured.")

    labels = [unit["label"] for unit in batch]
    errors: List[str] = []
    uploads: Dict[str, Tuple[Any, List[Any]]] = {}
    try:
        pool_pass = 0
        while True:
            pool_pass += 1
            for model in GEMINI_MODELS:
                for entry in rotated_keys(keys):
                    masked = entry["masked"]
                    if cooldown_remaining(entry["id"]) > 0:
                        errors.append(f"{masked}: cooling down")
                        continue
                    if cooldown_remaining(f"{entry['id']}:{model}") > 0:
                        errors.append(f"{masked} · {model}: cooling down")
                        continue
                    if entry["id"] not in uploads:
                        try:
                            uploads[entry["id"]] = upload_parts(entry, batch)
                        except Exception as exc:  # noqa: BLE001
                            errors.append(f"{masked} upload: {scrub(exc, entry['key'])}")
                            cooldown = classify_error(exc)
                            if cooldown:
                                set_cooldown(entry["id"], cooldown)
                            continue
                    client, handles = uploads[entry["id"]]

                    contents: List[Any] = []
                    for number, (unit, handle) in enumerate(zip(batch, handles), start=1):
                        contents.append(f"=== DOCUMENT: P{number} | {unit['label']} ===")
                        contents.append(handle)
                    contents.append(BATCH_PROMPT)
                    set_gemini_activity(True, f"Gemini request in progress · {model}")
                    try:
                        response = client.models.generate_content(
                            model=model,
                            contents=contents,
                            config=types.GenerateContentConfig(
                                response_mime_type="application/json",
                                temperature=0.1,
                                max_output_tokens=MAX_OUTPUT_TOKENS,
                            ),
                        )
                    except Exception as exc:  # noqa: BLE001
                        if is_size_error(exc):
                            raise OversizedBatch(scrub(exc, entry["key"]))
                        errors.append(f"{masked} · {model}: {scrub(exc, entry['key'])}")
                        cooldown = classify_error(exc)
                        if cooldown == AUTH_COOLDOWN_SECONDS:
                            # A rejected/forbidden key is dead for every model, not just
                            # this one: park the whole key for the hour.
                            set_cooldown(entry["id"], cooldown)
                        elif cooldown:
                            set_cooldown(f"{entry['id']}:{model}", cooldown)
                        continue
                    finally:
                        set_gemini_activity(False)

                    if "MAX_TOKENS" in finish_reason(response):
                        raise OversizedBatch("response was truncated at the output token limit")
                    try:
                        parsed = parse_batch_payload(response.text or "", labels)
                    except ValueError as exc:  # json.JSONDecodeError is a ValueError
                        errors.append(f"{masked} · {model}: invalid JSON ({str(exc)[:120]})")
                        continue
                    record_key_success(keys, entry["id"])
                    return parsed, f"{masked} · {model}"
            # Full pool pass without a usable candidate: if keys are only rate- or
            # quota-limited, wait for the soonest recovery and go around again.
            wait = earliest_cooldown(keys)
            if wait <= 0 or wait > MAX_COOLDOWN_WAIT or pool_pass >= MAX_POOL_PASSES:
                break
            set_gemini_activity(True, f"Every pooled key is cooling down; retrying in {int(wait)}s")
            time.sleep(wait)
            set_gemini_activity(False)
        raise RuntimeError("; ".join(errors[-6:]) or "every key and model failed")
    finally:
        for client, handles in uploads.values():
            delete_handles(client, handles)


def unit_failed(unit: Dict[str, Any], message: str) -> None:
    """Record that a part was not extracted. No substitute values are created."""
    unit["items"] = None
    unit["doc_type"] = None
    unit["status"] = "failed"
    unit["message"] = f"Not extracted: {message}"


def apply_batch_result(batch: List[Dict[str, Any]], parsed: Any, used: str) -> None:
    by_label, references = parsed
    units_by_label = {unit["label"]: unit for unit in batch}
    for unit in batch:
        payload = by_label.get(unit["label"])
        if payload is None:
            unit_failed(unit, "model returned no entry for this part")
            continue
        unit["items"] = [
            dict(
                raw,
                page_number=unit["offset"] + min(max(1, raw["page_number"]), unit["pages"]),
                _page_number_valid=raw.get("_page_number_valid", True)
                and raw["page_number"] <= unit["pages"],
            )
            for raw in payload["facts"]
        ]
        unit["doc_type"] = payload["document_type"] or unit["doc_type"]
        unit["status"] = "gemini"
        unit["message"] = f"via {used}"
    for ref in references:
        source = units_by_label.get(ref["from_label"])
        if source is None or source["items"] is None:
            continue
        local_page = min(max(1, ref["page_number"]), source["pages"])
        source["items"].append(dict(
            ref,
            page_number=source["offset"] + local_page,
            _page_number_valid=ref.get("_page_number_valid", True) and ref["page_number"] <= source["pages"],
        ))


def process_units(batch: List[Dict[str, Any]], keys: List[Dict[str, str]]) -> None:
    try:
        parsed, used = call_gemini_batch(batch, keys)
    except OversizedBatch as exc:
        if len(batch) > 1:
            middle = len(batch) // 2
            process_units(batch[:middle], keys)
            process_units(batch[middle:], keys)
            return
        unit_failed(batch[0], f"part too large to process even alone ({str(exc)[:150]})")
        return
    except Exception as exc:  # noqa: BLE001
        for unit in batch:
            unit_failed(unit, clean_text(exc)[:300])
        return
    apply_batch_result(batch, parsed, used)


# ---------------------------------------------------------------- ingestion
def prepare_document(docket_id: str, filename: str, data: bytes, mode: str) -> Tuple[Optional[Dict], Optional[Dict]]:
    digest = hashlib.sha256(data).hexdigest()
    with connect() as connection:
        duplicate = connection.execute(
            "SELECT id, filename, status FROM documents WHERE docket_id = ? AND sha256 = ?",
            (docket_id, digest),
        ).fetchone()
    if duplicate:
        if duplicate["status"] == "completed":
            return None, {
                "filename": filename,
                "status": "skipped",
                "fact_count": 0,
                "extracted_count": 0,
                "appended_count": 0,
                "message": f"Identical to “{duplicate['filename']}” already completed in this docket.",
            }
        delete_document(docket_id, duplicate["id"])
    try:
        pages = read_pages(filename, data)
    except Exception as exc:  # noqa: BLE001
        return None, {
            "filename": filename,
            "status": "failed",
            "fact_count": 0,
            "extracted_count": 0,
            "appended_count": 0,
            "message": clean_text(exc)[:300],
        }

    doc_id = str(uuid.uuid4())
    now = utc_now()
    mime = MIME_TYPES.get(extension_of(filename), "application/octet-stream")
    storage_path = store_original(docket_id, doc_id, filename, data)
    try:
        with connect() as connection:
            connection.execute(
                """INSERT INTO documents (id, docket_id, filename, mime_type, sha256, size_bytes,
                   storage_path, page_count, document_type, status, extraction_mode,
                   processed_units, failed_units, completeness_status, completeness_message, created_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (doc_id, docket_id, clean_text(filename), mime, digest, len(data), storage_path, len(pages),
                 "Unclassified", "pending", mode, 0, 0, "pending", "", now),
            )
            connection.executemany(
                "INSERT INTO pages (document_id, page_number, sheet_name, text, readable) VALUES (?, ?, ?, ?, ?)",
                [(doc_id, p["page_number"], p.get("sheet_name", ""), p["text"], int(p["readable"]))
                 for p in pages],
            )
    except Exception:
        remove_stored_document(storage_path)
        raise
    doc = {
        "id": doc_id, "filename": filename, "mime": mime, "data": data,
        "pages": pages, "page_count": len(pages),
        "page_text": {p["page_number"]: (p["text"] if p["readable"] else None) for p in pages},
        "units": [],
    }
    return doc, None


def finalize_document(docket_id: str, doc: Dict[str, Any], status: str,
                      document_type: str, items: List[Dict[str, Any]]) -> int:
    """Save all facts and return the number persisted for this document."""
    now = utc_now()
    # Pages/sheets the extractor actually returned at least one fact for. A page with no
    # text layer (e.g. an image, vision-only) can still land here once Gemini reads it -
    # "no text layer" and "nothing came back" are NOT the same thing.
    accounted_pages = {int(item.get("page_number", 1)) for item in items}
    failed_pages = [page for page in doc["pages"] if page["page_number"] not in accounted_pages]
    for page in failed_pages:
        items.append(_fact(
            page["page_number"],
            "Page processing status",
            "page_processing_status",
            "" if not page["readable"] else "No identifiable value returned",
            "",
        ))
    failed_units = len(failed_pages)
    completeness_status = "complete" if not failed_units and status == "completed" else "complete_with_warnings"
    if completeness_status == "complete":
        completeness_message = "Every page/sheet produced at least one extracted value via Gemini."
    elif status == "failed":
        completeness_message = "Gemini did not return usable values for this document; retry it."
    elif failed_units:
        completeness_message = f"{failed_units} page/sheet(s) returned no extracted value at all."
    else:
        completeness_message = "Some parts were not extracted by Gemini; see the processing message."
    with connect() as connection:
        for item in items:
            page_number = min(max(1, item["page_number"]), doc["page_count"])
            value = clean_text(item["value"])
            evidence = clean_text(item["evidence"])
            assessment = assess(value, evidence, doc["page_text"].get(page_number))
            if item.get("confidence_score") is not None:
                assessment["score"] = calibrated_confidence(
                    item["confidence_score"], value, evidence,
                    doc["page_text"].get(page_number), item["label"],
                )
                assessment["band"] = band_for(assessment["score"])
                assessment["reasons"] = list(dict.fromkeys(
                    assessment["reasons"] + ["MODEL_CONFIDENCE"]
                ))
            if not item.get("_page_number_valid", True):
                assessment["band"] = "Red"
                assessment["score"] = 1
                assessment["reasons"].append("INVALID_PAGE_REFERENCE")
            connection.execute(
                """INSERT INTO facts (id, docket_id, document_id, page_number, label, concept,
                   value, original_value, source_evidence, confidence_band, confidence_score,
                   confidence_reasons, review_status, corrected, created_at, updated_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (str(uuid.uuid4()), docket_id, doc["id"], page_number, clean_text(item["label"]),
                 clean_text(item["concept"]), value, value, evidence,
                 assessment["band"], assessment["score"], json.dumps(assessment["reasons"]),
                 "pending", 0, now, now),
            )
        connection.execute(
            """UPDATE documents SET status = ?, document_type = ?, processed_units = ?,
               failed_units = ?, completeness_status = ?, completeness_message = ?
               WHERE id = ?""",
            (status, clean_text(document_type or "Unclassified")[:80],
             doc["page_count"] - failed_units,
             failed_units, completeness_status, completeness_message, doc["id"]),
        )
        connection.execute("UPDATE dockets SET updated_at = ? WHERE id = ?", (now, docket_id))
    try:
        apply_extraction_integrity(docket_id)
    except Exception:  # noqa: BLE001
        # The facts are already saved; a failed consistency pass must not mark them failed.
        LOG.exception("consistency pass failed for docket %s", docket_id)
    with connect() as connection:
        return int(connection.execute(
            "SELECT COUNT(*) FROM facts WHERE document_id = ?", (doc["id"],)
        ).fetchone()[0])


def _fact_group_key(label: str, concept: str) -> str:
    return normalise(label) or normalise(concept) or "other"


def _append_reason(raw: str, reason: str) -> str:
    reasons = list(dict.fromkeys(json.loads(raw or "[]") + [reason]))
    return json.dumps(reasons)


def apply_extraction_integrity(docket_id: str) -> None:
    """Flag contradictory, repeated or ill-typed values without inventing scores.

    Every fact keeps the model-calibrated confidence it was assigned at extraction —
    duplicate detection, repetition and conflicts only APPEND reasons, they never
    overwrite the score or band. Identical duplicate extractions are still
    auto-approved (workflow preserved), with their own per-occurrence scores intact.
    Only pending, uncorrected extraction results are changed. Human decisions remain
    authoritative and are never silently overwritten.
    """
    with connect() as connection:
        rows = connection.execute(
            """SELECT id, label, concept, value, review_status, corrected, confidence_band,
                      confidence_score, confidence_reasons
               FROM facts WHERE docket_id = ? ORDER BY created_at, id""",
            (docket_id,),
        ).fetchall()
        groups: Dict[str, List[sqlite3.Row]] = {}
        for row in rows:
            # Rejected facts are not part of any group: a rejected value must never make a
            # new pending fact look like a duplicate, or create a conflict.
            if row["value"].strip() and row["review_status"] != "rejected":
                groups.setdefault(_fact_group_key(row["label"], row["concept"]), []).append(row)

        conflict_groups = {
            key for key, group in groups.items()
            if len({normalise(row["value"]) for row in group if normalise(row["value"])}) > 1
        }
        now = utc_now()
        changed = False
        for key, group in groups.items():
            if key in conflict_groups:
                for row in group:
                    if row["review_status"] != "pending" or row["corrected"]:
                        continue
                    # Conflicting occurrences keep their own model scores; the reason
                    # makes the conflict visible in Review and Validation instead.
                    reasons = _append_reason(row["confidence_reasons"], "LABEL_VALUE_CONFLICT")
                    connection.execute(
                        "UPDATE facts SET confidence_reasons = ?, updated_at = ? WHERE id = ?",
                        (reasons, now, row["id"]),
                    )
                    changed = True
                continue

            duplicate_values: Dict[str, List[sqlite3.Row]] = {}
            for row in group:
                duplicate_values.setdefault(normalise(row["value"]), []).append(row)
            for value_rows in duplicate_values.values():
                if len(value_rows) < 2:
                    continue
                for row in value_rows:
                    if row["review_status"] != "pending" or row["corrected"]:
                        continue
                    reasons = _append_reason(
                        row["confidence_reasons"], "DUPLICATE_EXTRACTION_AUTO_APPROVED"
                    )
                    connection.execute(
                        """INSERT INTO fact_history
                           (id, fact_id, old_value, new_value, old_status, new_status,
                            changed_at, reason)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                        (str(uuid.uuid4()), row["id"], row["value"], row["value"],
                         row["review_status"], "approved", now,
                         "Duplicate extraction matched another fact."),
                    )
                    connection.execute(
                        """UPDATE facts SET confidence_reasons = ?, review_status = 'approved',
                           updated_at = ?
                           WHERE id = ?""",
                        (reasons, now, row["id"]),
                    )
                    changed = True

        for row in rows:
            if row["review_status"] != "pending" or row["corrected"]:
                continue
            if _fact_group_key(row["label"], row["concept"]) in conflict_groups:
                continue
            if not label_value_mismatch(row["label"], row["value"]):
                continue
            reasons = _append_reason(row["confidence_reasons"], "LABEL_VALUE_MISMATCH")
            connection.execute(
                "UPDATE facts SET confidence_reasons = ?, updated_at = ? WHERE id = ?",
                (reasons, now, row["id"]),
            )
            changed = True
        if changed:
            connection.execute("UPDATE dockets SET updated_at = ? WHERE id = ?", (now, docket_id))


def gemini_document_status(doc: Dict[str, Any]) -> str:
    statuses = {unit["status"] for unit in doc["units"]}
    if statuses == {"gemini"}:
        return "completed"
    if "gemini" not in statuses:
        return "failed"
    return "partial"


def set_gemini_activity(active: bool, message: str = "") -> None:
    job_id = getattr(WORKER_STATE, "job_id", None)
    if not job_id:
        return
    with PROCESSING_LOCK:
        job = PROCESSING_JOBS.get(job_id)
        if job is not None:
            job["gemini_active"] = active
            job["phase"] = "gemini" if active else "processing"
            if message:
                job["message"] = message


def job_result(filename: str, status: str, message: str,
               extracted: int = 0, appended: int = 0) -> Dict[str, Any]:
    """One result record shape for every outcome, matching what the results panel reads."""
    return {"filename": filename, "status": status, "fact_count": appended,
            "extracted_count": extracted, "appended_count": appended, "message": message}


def _run_processing_job(job_id: str, docket_id: str, uploads: List[Dict[str, Any]],
                        keys: List[Dict[str, str]],
                        initial_results: Optional[List[Dict[str, Any]]] = None) -> None:
    WORKER_STATE.job_id = job_id
    results: List[Dict[str, Any]] = list(initial_results or [])
    documents: List[Dict[str, Any]] = []   # documents successfully stored, awaiting extraction
    finalized: set = set()                 # ids of documents whose facts were saved

    def mark_done() -> None:
        with PROCESSING_LOCK:
            job = PROCESSING_JOBS.get(job_id)
            if job is not None:
                job["done_docs"] = int(job.get("done_docs", 0)) + 1

    def save(
        doc: Dict[str, Any],
        status: str,
        doc_type: str,
        items: List[Dict[str, Any]],
        message: str,
    ) -> Dict[str, Any]:
        set_document_status(doc["id"], "saving")
        try:
            fact_count = finalize_document(docket_id, doc, status, doc_type, items)
            finalized.add(doc["id"])
            return {
                "filename": doc["filename"],
                "status": status,
                "fact_count": fact_count,
                "extracted_count": len(items),
                "appended_count": fact_count,
                "message": message,
            }
        except Exception as exc:  # noqa: BLE001
            LOG.exception("could not save %s", doc["filename"])
            try:
                set_document_status(doc["id"], "failed")
            except Exception:  # noqa: BLE001
                LOG.exception("could not mark %s as failed", doc["filename"])
            return {
                "filename": doc["filename"],
                "status": "failed",
                "fact_count": 0,
                "extracted_count": 0,
                "appended_count": 0,
                "message": f"Could not save results ({clean_text(exc)[:150]}). Retry this document.",
            }

    def progress(message: str) -> None:
        with PROCESSING_LOCK:
            PROCESSING_JOBS[job_id].update(message=message, phase="processing", gemini_active=False)

    try:
        with PROCESSING_LOCK:
            PROCESSING_JOBS[job_id].update(status="running", phase="processing", gemini_active=False)
        progress("Preparing the queued documents…")
        batch_digests: Dict[str, str] = {}  # sha256 -> filename, for duplicates inside this upload
        for upload in uploads:
            data = upload["data"]
            if upload.get("document") is not None:
                documents.append(upload["document"])
                batch_digests[hashlib.sha256(data).hexdigest()] = upload["filename"]
                continue
            problem = validate_upload(upload["filename"], data)
            digest = hashlib.sha256(data).hexdigest()
            if not problem and digest in batch_digests:
                results.append(job_result(upload["filename"], "skipped",
                                          f"Identical to “{batch_digests[digest]}” in this upload."))
                continue
            if problem:
                results.append(job_result(upload["filename"], "failed", problem))
                continue
            try:
                doc, skipped = prepare_document(docket_id, upload["filename"], data, "Gemini")
            except Exception as exc:  # noqa: BLE001
                LOG.exception("prepare failed for %s", upload["filename"])
                results.append(job_result(upload["filename"], "failed",
                                          f"Could not store the file ({clean_text(exc)[:200]})."))
                continue
            if doc:
                documents.append(doc)
                batch_digests[digest] = upload["filename"]
            else:
                results.append(skipped)

        with PROCESSING_LOCK:
            PROCESSING_JOBS[job_id].update(total_docs=len(documents), done_docs=0)

        for document_index, doc in enumerate(documents, start=1):
            set_document_status(doc["id"], "processing")
            progress(
                f"Processing document {document_index}/{len(documents)}: "
                f"{doc['filename']}"
            )
            try:
                try:
                    units = make_units(doc)
                except Exception as exc:  # noqa: BLE001
                    doc["units"] = [{
                        "label": doc["filename"], "doc": doc, "offset": 0,
                        "pages": doc["page_count"], "data": doc["data"],
                        "mime": doc["mime"], "suffix": "",
                        "pages_slice": doc["pages"], "items": None, "status": None,
                        "message": "", "doc_type": None,
                    }]
                    units = doc["units"]
                    LOG.warning(
                        "could not split %s; sending this document whole: %s",
                        doc["filename"], exc,
                    )

                # The batch boundary is deliberately inside this document loop.
                # A Gemini request can therefore never contain units from another
                # uploaded document.
                batches = pack_batches([u for u in units if u["status"] is None])
                for request_index, batch in enumerate(batches, start=1):
                    progress(
                        f"{doc['filename']}: request {request_index}/{len(batches)}"
                    )
                    process_units(batch, keys)

                items = [item for unit in units for item in (unit["items"] or [])]
                doc_type = next((u["doc_type"] for u in units if u["doc_type"]), "")
                messages = sorted({u["message"] for u in units if u["message"]})
                results.append(save(
                    doc,
                    gemini_document_status({"units": units}),
                    doc_type or "Unclassified",
                    items,
                    " | ".join(messages),
                ))
                mark_done()
            except Exception as exc:  # noqa: BLE001
                LOG.exception("document processing failed for %s", doc["filename"])
                set_document_status(doc["id"], "failed")
                results.append(job_result(
                    doc["filename"], "failed",
                    f"Processing failed ({clean_text(exc)[:200]}). "
                    "Completed documents were preserved; retry this document.",
                ))
                mark_done()
        with PROCESSING_LOCK:
            PROCESSING_JOBS[job_id].update(
                status="complete", phase="complete", gemini_active=False,
                results=results, message="Processing complete.", finished_at=time.time(),
            )
    except Exception as exc:  # noqa: BLE001
        LOG.exception("processing run failed")
        for doc in documents:
            if doc["id"] not in finalized:
                set_document_status(doc["id"], "failed")
                already_reported = any(
                    result["filename"] == doc["filename"]
                    and result["status"] == "failed"
                    for result in results
                )
                if already_reported:
                    continue
                results.append(job_result(
                    doc["filename"], "failed",
                    f"Processing was interrupted ({clean_text(exc)[:150]}). "
                    "Completed documents were preserved; retry this document."))
        with PROCESSING_LOCK:
            PROCESSING_JOBS[job_id].update(
                status="error", phase="error", gemini_active=False,
                results=results, message=f"Processing stopped: {clean_text(exc)[:200]}",
                finished_at=time.time(),
            )
    finally:
        WORKER_STATE.job_id = None
        WORKER_STATE.state = None


def processing_job(docket_id: str) -> Optional[Dict[str, Any]]:
    with PROCESSING_LOCK:
        jobs = [job for job in PROCESSING_JOBS.values() if job["docket_id"] == docket_id]
        active = [job for job in jobs if job["status"] in {"queued", "running"}]
        return dict((active or jobs)[-1]) if (active or jobs) else None


_LOADER_CSS = """
<style>
.pp-loader{display:flex;gap:1rem;align-items:center;padding:.9rem 1rem;border-radius:12px;
  border:1px solid rgba(128,128,128,.25);background:rgba(128,128,128,.06);margin:.5rem 0 1rem}
.pp-ring{position:relative;width:64px;height:64px;flex:0 0 64px}
.pp-ring svg{width:64px;height:64px;transform:rotate(-90deg)}
.pp-track{fill:none;stroke:rgba(128,128,128,.25);stroke-width:6}
.pp-bar{fill:none;stroke:#2563eb;stroke-width:6;stroke-linecap:round;transition:stroke-dasharray .6s ease}
.pp-orbit{fill:none;stroke:#60a5fa;stroke-width:6;stroke-linecap:round;
  transform-origin:32px 32px;animation:pp-spin 1.1s linear infinite}
.pp-pct{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;
  font-weight:600;font-size:.85rem}
.pp-title{font-weight:600;margin-bottom:.15rem}
.pp-msg{font-size:.9rem;opacity:.9;word-break:break-word}
.pp-sub{font-size:.8rem;opacity:.7;margin-top:.25rem}
@keyframes pp-spin{to{transform:rotate(360deg)}}
@media (prefers-reduced-motion: reduce){.pp-orbit{animation:none;opacity:.5}}
</style>
"""


def _render_processing_status(job: Optional[Dict[str, Any]]) -> None:
    if not job:
        return
    if job["status"] in {"queued", "running"}:
        total = max(0, int(job.get("total_docs") or 0))
        done = max(0, min(int(job.get("done_docs") or 0), total)) if total else 0
        percent = round(100 * done / total) if total else 0
        if job.get("gemini_active"):
            title = "Gemini is extracting facts"
        elif job["status"] == "queued":
            title = "Queued for processing"
        else:
            title = "Processing saved documents"
        message = esc(job.get("message") or "Preparing the saved documents.")
        st.markdown(
            _LOADER_CSS
            + '<div class="pp-loader" role="status" aria-live="polite">'
            + '<div class="pp-ring">'
            + '<svg viewBox="0 0 64 64" aria-hidden="true">'
            + '<circle class="pp-track" cx="32" cy="32" r="28" pathLength="100"/>'
            + f'<circle class="pp-bar" cx="32" cy="32" r="28" pathLength="100" '
            + f'style="stroke-dasharray:{percent} 100"/>'
            + '<circle class="pp-orbit" cx="32" cy="32" r="28" pathLength="100" '
            + 'style="stroke-dasharray:14 86"/>'
            + '</svg>'
            + f'<div class="pp-pct">{percent}%</div></div>'
            + '<div><div class="pp-title">' + esc(title) + '…</div>'
            + f'<div class="pp-msg">{message}</div>'
            + f'<div class="pp-sub">{done} of {total} document(s) finished · '
            + 'you can switch tabs safely; facts appear as each document completes.</div>'
            + '</div></div>',
            unsafe_allow_html=True,
        )
    elif job["status"] == "error":
        st.error(job["message"], icon="🚫")


@st.fragment(run_every=1)
def render_processing_loader(docket_id: str) -> None:
    _render_processing_status(processing_job(docket_id))


def processing_active(docket_id: Optional[str]) -> bool:
    """True while a queued/running job exists for this docket: used to disable
    duplicate submissions everywhere (requirement: no double submissions)."""
    if not docket_id:
        return False
    job = processing_job(docket_id)
    return bool(job and job["status"] in {"queued", "running"})


# ---------------------------------------------------------------- save ≠ process
def save_uploaded_document(docket_id: str, upload: Any) -> Dict[str, Any]:
    """Store ONE uploaded file without processing it. Saving and extraction are
    deliberately independent: the document lands in the Workspace as “saved”."""
    filename = upload_name(upload)
    data = upload_data(upload)
    problem = validate_upload(filename, data)
    if problem:
        return job_result(filename, "failed", problem)
    try:
        doc, skipped = prepare_document(docket_id, filename, data, "Gemini")
    except Exception as exc:  # noqa: BLE001
        LOG.exception("could not store %s", filename)
        return job_result(filename, "failed",
                          f"Could not store the file ({clean_text(exc)[:200]}).")
    if doc is None:
        return skipped  # identical to an already-completed document
    set_document_status(doc["id"], "saved")
    return job_result(filename, "saved",
                      f"Stored {doc['page_count']} page(s)/sheet(s). Not processed yet.")


def render_upload_panel(docket_id: str) -> None:
    """Workspace step 1: save uploads. No Gemini call ever happens here."""
    with st.container(border=True):
        st.markdown("#### 1 · Upload and save documents")
        st.caption("Saving only stores the original file in the workspace. "
                   "Extraction runs only from step 2 below, on documents that were saved.")
        nonce = st.session_state.get("upload_nonce", 0)
        uploads = st.file_uploader(
            f"Drop PDF, PNG, JPG, TXT, CSV or XLSX files here (up to {fmt_size(MAX_UPLOAD_BYTES)} each)",
            type=SUPPORTED_TYPES,
            accept_multiple_files=True,
            key=f"uploads_{nonce}",
        )
        if not uploads:
            st.caption("Choose one or more files to enable saving.")
        if st.button("Save documents", key="save_documents", type="primary",
                     disabled=not uploads, **stretch()):
            if st.session_state.get("ws_saving"):
                queue_flash("warning", "Already saving the current selection; wait for it to finish.")
                return
            st.session_state["ws_saving"] = True
            try:
                with st.spinner(f"Saving {len(uploads)} document(s) to the workspace…"):
                    results = [save_uploaded_document(docket_id, upload) for upload in uploads]
            finally:
                st.session_state["ws_saving"] = False
            st.session_state["upload_nonce"] = nonce + 1
            saved = sum(result["status"] == "saved" for result in results)
            for result in results:
                if result["status"] == "failed":
                    queue_flash("error", f"“{result['filename']}” was not saved: {result['message']}")
                elif result["status"] == "skipped":
                    queue_flash("warning", f"“{result['filename']}” skipped: {result['message']}")
            if saved:
                queue_flash("success", f"Saved {saved} document(s). Pick them in step 2 to run extraction.")
            st.rerun(scope="app")


def load_saved_document(docket_id: str, document_id: str) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    """Rebuild the worker's document payload from the STORED original and the saved
    page rows — never from temporary session memory — after a checksum verification."""
    with connect() as connection:
        row = connection.execute(
            """SELECT id, filename, mime_type, sha256, storage_path, page_count, status
               FROM documents WHERE id = ? AND docket_id = ?""",
            (document_id, docket_id),
        ).fetchone()
        if row is None:
            return None, "That document no longer exists in this docket."
        pages = connection.execute(
            """SELECT page_number, sheet_name, text, readable FROM pages
               WHERE document_id = ? ORDER BY page_number""",
            (document_id,),
        ).fetchall()
    data, problem = read_verified(row["storage_path"], row["sha256"])
    if problem:
        return None, problem
    page_list = [{"page_number": p["page_number"], "sheet_name": p["sheet_name"],
                  "text": p["text"], "readable": bool(p["readable"])} for p in pages]
    document = {
        "id": row["id"], "filename": row["filename"], "mime": row["mime_type"],
        "data": data, "pages": page_list, "page_count": row["page_count"],
        "page_text": {p["page_number"]: (p["text"] if p["readable"] else None) for p in page_list},
        "units": [],
    }
    return document, None


def clear_document_facts(document_id: str) -> None:
    """Re-processing replaces the document's earlier extraction: drop its facts
    (fact_history cascades) and reset its completeness counters first."""
    with connect() as connection:
        connection.execute("DELETE FROM facts WHERE document_id = ?", (document_id,))
        connection.execute(
            """UPDATE documents SET processed_units = 0, failed_units = 0,
               completeness_status = 'pending', completeness_message = '' WHERE id = ?""",
            (document_id,),
        )


JOB_RETENTION_SECONDS = 3600


def _prune_finished_jobs() -> None:
    cutoff = time.time() - JOB_RETENTION_SECONDS
    with PROCESSING_LOCK:
        for stale_id in [
            job_id for job_id, job in PROCESSING_JOBS.items()
            if job["status"] in {"complete", "error"} and job.get("finished_at", 0) < cutoff
        ]:
            del PROCESSING_JOBS[stale_id]


def queue_saved_documents(docket_id: str, document_ids: List[str]) -> Tuple[int, List[Dict[str, Any]]]:
    """Queue saved documents for background extraction, from stored bytes only.
    Returns (count queued, immediate failure results) and flashes suitable messages."""
    keys = key_pool()
    if not keys:
        queue_flash(
            "error",
            f"No Gemini API keys are configured. Add 4–5 keys to “gemini_api_keys” in "
            f"{CONFIG_PATH.name} (see Settings) before processing.",
        )
        return 0, []
    existing = processing_job(docket_id)
    if existing and existing["status"] in {"queued", "running"}:
        queue_flash("warning", "This docket is already being processed in the background.")
        return 0, []
    _prune_finished_jobs()
    wanted = set(document_ids)
    payloads: List[Dict[str, Any]] = []
    results: List[Dict[str, Any]] = []
    for row in list_documents(docket_id):
        if row["id"] not in wanted or row["status"] not in PROCESSABLE_STATES:
            continue
        document, problem = load_saved_document(docket_id, row["id"])
        if problem:
            set_document_status(row["id"], "failed")
            results.append(job_result(row["filename"], "failed", problem))
            continue
        clear_document_facts(row["id"])
        set_document_status(row["id"], "queued")
        payloads.append({"filename": row["filename"], "data": document["data"], "document": document})
    if not payloads:
        queue_flash("info", "No queued documents to start on.")
        return 0, results
    job_id = str(uuid.uuid4())
    with PROCESSING_LOCK:
        PROCESSING_JOBS[job_id] = {
            "id": job_id, "docket_id": docket_id, "status": "queued", "phase": "queued",
            "gemini_active": False, "message": "Queued for Gemini extraction…",
            "results": results, "total_docs": len(payloads), "done_docs": 0, "finished_at": 0.0,
        }
    PROCESSING_EXECUTOR.submit(_run_processing_job, job_id, docket_id, payloads, keys, [])
    queue_flash("success", f"Queued {len(payloads)} document(s) for Gemini extraction. "
                           "Progress appears in step 2; you can switch tabs safely.")
    return len(payloads), results


@st.fragment(run_every=1)
def render_process_panel(docket_id: str) -> None:
    """Workspace step 2: choose saved documents and start extraction on them.
    Reruns every second while idle too, so the state counts and progress stay live."""
    job = processing_job(docket_id)
    active = bool(job and job["status"] in {"queued", "running"})
    # Surface a finished run into the main page exactly once, so the results panel
    # and the stored-documents table refresh without waiting for a manual rerun.
    if st.session_state.pop("clear_process_pick", False):
        # Cleared here, before the multiselect is created in this run (Streamlit forbids
        # changing a widget's state after it has been instantiated).
        st.session_state.pop("process_pick", None)
    if job and job["status"] in {"complete", "error"} \
            and st.session_state.get("last_surfaced_job_id") != job["id"]:
        st.session_state["last_results"] = job.get("results") or []
        st.session_state["last_surfaced_job_id"] = job["id"]
        st.rerun(scope="app")
    # Live updates: each time a document finishes, refresh the whole app so its facts
    # appear in Workspace, Review and Main Metadata straight away.
    progress_mark = (job["id"], job.get("done_docs", 0)) if job else None
    if progress_mark is not None and progress_mark != st.session_state.get("last_seen_progress"):
        st.session_state["last_seen_progress"] = progress_mark
        if st.session_state.get("last_seen_progress_seeded"):
            st.rerun(scope="app")
        st.session_state["last_seen_progress_seeded"] = True
    with st.container(border=True):
        st.markdown("#### 2 · Process saved documents")
        documents = list_documents(docket_id)
        if documents:
            counts: Dict[str, int] = {}
            for d in documents:
                counts[d["status"]] = counts.get(d["status"], 0) + 1
            st.markdown(
                " ".join(
                    badge(
                        f"{document_state_label(status)}: {counts[status]}",
                        DOCUMENT_STATE_META.get(status, ("", ""))[1],
                    )
                    for status in ("saved", "pending", "queued", "processing", "completed", "partial", "failed")
                    if status in counts
                ),
                unsafe_allow_html=True,
            )
        st.caption("Extraction reads the stored original (checksum-verified) and sends it to the "
                   "pooled Gemini keys. Saved, failed and partially processed documents are listed; "
                   "re-running a partial document replaces its earlier results.")
        processable = [d for d in documents if d["status"] in PROCESSABLE_STATES]
        if not processable and not active:
            st.info("No documents are waiting to be processed. Save files in step 1 above first.")
        elif processable:
            labels = {
                d["id"]: (
                    f"{d['filename']} · {d['page_count']} page(s)/sheet(s) · "
                    f"{document_state_label(d['status'])}"
                    + (" — re-run replaces earlier results" if d["status"] == "partial" else "")
                )
                for d in processable
            }
            chosen = st.multiselect(
                "Documents to process",
                list(labels),
                format_func=lambda identifier: labels[identifier],
                key="process_pick",
                disabled=active,
                help="Pick one or more; or use “Process all ready” for everything waiting.",
            )
            selected_col, all_col = st.columns(2)
            start_selected = selected_col.button(
                f"Process selected ({len(chosen)})",
                type="primary", key="process_selected",
                disabled=active or not chosen, **stretch(),
            )
            start_all = all_col.button(
                f"Process all ready ({len(processable)})",
                key="process_all", disabled=active, **stretch(),
            )
            if start_selected or start_all:
                ids = list(chosen) if start_selected else list(labels)
                queue_saved_documents(docket_id, ids)
                st.session_state["clear_process_pick"] = True
                st.rerun(scope="app")
        _render_processing_status(job)
        for failed in [r for r in st.session_state.get("last_results") or [] if r["status"] == "failed"]:
            st.error(f"{failed['filename']}: {failed['message']}", icon="🚫")
        if active:
            st.caption("A run is active, so the process buttons are disabled to avoid duplicate submissions.")

# ---------------------------------------------------------------- review operations
def update_fact(fact_id: str, value: str, status: str, reason: str) -> None:
    value = clean_text(value)
    with connect() as connection:
        row = connection.execute(
            """SELECT f.*, p.text AS page_text, p.readable AS page_readable
               FROM facts f LEFT JOIN pages p
                 ON p.document_id = f.document_id AND p.page_number = f.page_number
               WHERE f.id = ?""",
            (fact_id,),
        ).fetchone()
        if row is None:
            raise ValueError("That fact no longer exists.")
        now = utc_now()
        if row["value"] != value or row["review_status"] != status:
            connection.execute(
                """INSERT INTO fact_history (id, fact_id, old_value, new_value, old_status,
                   new_status, changed_at, reason) VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                (str(uuid.uuid4()), fact_id, row["value"], value,
                 row["review_status"], status, now, reason),
            )
        corrected = int(bool(row["corrected"]) or value != row["original_value"])
        if status in ("approved", "rejected"):
            # A human decision changes the review status only: the model's per-occurrence
            # confidence score is preserved exactly as extracted (band follows the score).
            score = clamp_confidence(row["confidence_score"], 50)
            band = band_for(score)
            reasons = [r for r in json.loads(row["confidence_reasons"] or "[]")
                       if r not in {"HUMAN_APPROVED", "HUMAN_REJECTED"}]
            reasons.append("HUMAN_APPROVED" if status == "approved" else "HUMAN_REJECTED")
        else:
            page_text = row["page_text"] if row["page_readable"] else None
            assessment = assess(value, row["source_evidence"], page_text)
            band, score, reasons = assessment["band"], assessment["score"], assessment["reasons"]
        if corrected:
            reasons = list(dict.fromkeys(reasons + ["HUMAN_EDITED"]))
        connection.execute(
            """UPDATE facts SET value = ?, corrected = ?, confidence_band = ?, confidence_score = ?,
               confidence_reasons = ?, review_status = ?, updated_at = ? WHERE id = ?""",
            (value, corrected, band, score, json.dumps(reasons), status, now, fact_id),
        )
        connection.execute("UPDATE dockets SET updated_at = ? WHERE id = ?", (now, row["docket_id"]))


def approve_green(docket_id: str) -> int:
    count = 0
    for fact in fetch_facts(docket_id):
        if fact["confidence_band"] == "Green" and fact["review_status"] == "pending":
            update_fact(fact["id"], fact["value"], "approved", "bulk approve Green")
            count += 1
    return count


# ---------------------------------------------------------------- validation
def _number(value: str) -> Optional[float]:
    if "%" in value:
        return None  # a rate (e.g. "18%"), not an absolute amount
    match = re.search(r"-?\d[\d,]*(?:\.\d+)?", value)
    if not match:
        return None
    try:
        return float(match.group(0).replace(",", ""))
    except ValueError:
        return None


def validation_report(docket_id: str) -> List[Dict[str, str]]:
    with connect() as connection:
        documents = connection.execute(
            """SELECT id, filename, status, storage_path, sha256, page_count,
                      processed_units, failed_units, completeness_status, completeness_message
               FROM documents WHERE docket_id = ? ORDER BY filename""",
            (docket_id,),
        ).fetchall()
        unreadable = connection.execute(
            """SELECT d.filename, p.page_number FROM pages p
               JOIN documents d ON d.id = p.document_id
               WHERE d.docket_id = ? AND p.readable = 0 ORDER BY d.filename, p.page_number""",
            (docket_id,),
        ).fetchall()
    if not documents:
        return [{"severity": "Red", "check": "Completeness", "message": "No documents in this docket."}]

    rows: List[Dict[str, str]] = []

    def add(severity: str, check: str, message: str) -> None:
        rows.append({"severity": severity, "check": check, "message": message})

    for document in documents:
        _, problem = read_verified(document["storage_path"], document["sha256"])
        if problem:
            add("Red", "Storage", f"{document['filename']}: {problem}")
        if document["status"] == "partial":
            add("Yellow", "Extraction", f"{document['filename']}: some parts were not extracted by Gemini.")
        if document["status"] == "failed":
            add("Red", "Extraction", f"{document['filename']}: Gemini did not extract this document; retry it.")
        if document["processed_units"] != document["page_count"]:
            add("Red", "Completeness",
                f"{document['filename']}: only {document['processed_units']} of "
                f"{document['page_count']} pages/sheets were accounted for.")
        elif document["completeness_status"] != "complete":
            add("Yellow", "Completeness",
                f"{document['filename']}: {document['completeness_message']}")
    for page in unreadable:
        add("Yellow", "Completeness",
            f"{page['filename']} p.{page['page_number']}: no text layer; Gemini mode or OCR required.")

    facts = fetch_facts(docket_id)
    for fact in facts:
        where = f"{fact['filename']} p.{fact['page_number']} · {fact['label']}"
        if fact["confidence_band"] == "Red" and fact["review_status"] != "rejected":
            add("Red", "Evidence", f"{where}: {', '.join(json.loads(fact['confidence_reasons']))}.")
        elif fact["review_status"] == "pending":
            add("Yellow", "Human review", f"{where}: awaiting review.")

    groups: Dict[str, Dict[str, set]] = {}
    for fact in facts:
        if fact["concept"] in IDENTITY_CONCEPTS and fact["value"].strip() and fact["review_status"] != "rejected":
            groups.setdefault(fact["concept"], {}).setdefault(normalise(fact["value"]), set()).add(fact["filename"])
    for concept, values in groups.items():
        if len(values) > 1:
            add("Yellow", "Cross-document consistency",
                f"{concept} differs across documents: {', '.join(sorted(values))}.")

    total_keys = ("grand_total", "total_amount", "invoice_total", "total")
    subtotal_keys = ("subtotal", "sub_total", "taxable_value")
    tax_keys = ("total_tax", "tax_amount", "gst_total", "tax")
    for document in documents:
        doc_facts = [f for f in facts if f["filename"] == document["filename"] and f["review_status"] != "rejected"]

        def first(keys: Tuple[str, ...]) -> Optional[float]:
            for fact in doc_facts:
                if fact["concept"] in keys:
                    number = _number(fact["value"])
                    if number is not None:
                        return number
            return None

        subtotal, total, tax = first(subtotal_keys), first(total_keys), first(tax_keys)
        if subtotal is not None and total is not None:
            expected = subtotal + (tax or 0.0)
            if abs(expected - total) > 1.0:
                add("Red", "Arithmetic",
                    f"{document['filename']}: subtotal + tax = {expected:,.2f} but total is {total:,.2f}.")

    if not rows:
        add("Green", "Validation", "No unresolved issues found.")
    return rows


def storage_audit() -> List[str]:
    """Database integrity plus a checksum pass over every stored original."""
    problems: List[str] = []
    with connect() as connection:
        result = connection.execute("PRAGMA integrity_check").fetchone()[0]
        if result != "ok":
            problems.append(f"Database integrity check reported: {result}")
        for row in connection.execute("SELECT filename, storage_path, sha256 FROM documents").fetchall():
            _, problem = read_verified(row["storage_path"], row["sha256"])
            if problem:
                problems.append(f"{row['filename']}: {problem}")
    return problems


# ---------------------------------------------------------------- export (PSV and JSON)
class ExportError(Exception):
    """An export failed its integrity check and must not be offered for download."""


def metadata_records(docket_id: str) -> List[dict]:
    return [
        {
            "document_name": row["document_name"],
            "label": row["label"],
            "value": row["value"],
            "confidence_score": int(row["confidence_score"]),
        }
        for row in fetch_metadata(docket_id)
    ]


def client_template_records(docket_id: str, template: sqlite3.Row) -> Tuple[List[dict], List[str]]:
    """Every Main Metadata occurrence whose label has a SAVED mapping to this template.

    Read-only: rendering a page never creates or changes a mapping. Mappings are created only
    by the Organize tab (direct labels on create/edit, exact and approved matches on apply).
    Duplicates and conflicting values stay separate rows with their document names and scores.
    Template labels with no saved mapping are reported in `missing`, never filled in.
    """
    try:
        raw_labels = json.loads(template["labels_json"])
        labels = [str(label) for label in raw_labels] if isinstance(raw_labels, list) else []
    except (TypeError, ValueError):
        labels = []
    mappings = approved_mappings(docket_id, template["id"])
    # Mappings may be stored with one spelling of a source label while facts carry another,
    # so the lookup compares case-insensitively and ignores surrounding whitespace.
    mapping_by_normalized = {normalize_label(source): target for source, target in mappings.items()}
    output: List[dict] = []
    matched = {label: 0 for label in labels}
    for record in metadata_records(docket_id):
        target = mapping_by_normalized.get(normalize_label(record["label"]))
        if target in matched:
            output.append({**record, "label": target})
            matched[target] += 1
    return output, [label for label, count in matched.items() if count == 0]


def build_metadata_export(records: List[dict], description: str) -> Tuple[bytes, int]:
    exported = [
        {
            "Document Name": record["document_name"],
            "Label": record["label"],
            "Value": record["value"],
            "Confidence Score": int(record["confidence_score"]),
        }
        for record in records
    ]
    data = json_bytes(exported)
    if json.loads(data.decode("utf-8")) != exported:
        raise ExportError(f"{description} integrity check failed after re-reading.")
    return data, len(records)


def build_metadata_psv(records: List[dict], description: str) -> Tuple[bytes, int]:
    expected = [METADATA_HEADERS] + [
        [
            safe_cell(record["document_name"]),
            safe_cell(record["label"]),
            safe_cell(record["value"]),
            str(int(record["confidence_score"])),
        ]
        for record in records
    ]
    buffer = io.StringIO(newline="")
    csv.writer(buffer, delimiter="|", quotechar='"', quoting=csv.QUOTE_MINIMAL,
               doublequote=True, lineterminator="\r\n").writerows(expected)
    text = buffer.getvalue()
    try:
        parsed = list(csv.reader(io.StringIO(text, newline=""), delimiter="|", quotechar='"',
                                 doublequote=True, strict=True))
    except csv.Error as exc:
        raise ExportError(f"{description} PSV could not be read back: {exc}") from exc
    if parsed != expected:
        raise ExportError(f"{description} PSV integrity check failed after re-reading.")
    return ("\ufeff" + text).encode("utf-8"), len(records)


def build_metadata_xlsx(records: List[dict], description: str) -> Tuple[bytes, int]:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Metadata"
    sheet.append(METADATA_HEADERS)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for row_number, record in enumerate(records, start=2):
        for column_number, field in enumerate(METADATA_FIELDS, start=1):
            value = record[field]
            cell = sheet.cell(row=row_number, column=column_number, value=value)
            if isinstance(value, str) and value[:1] in ("=", "+", "-", "@"):
                cell.data_type = "s"
        score = int(record["confidence_score"])
        color = {"Red": "FFC7CE", "Yellow": "FFEB9C", "Green": "C6EFCE"}[band_for(score)]
        sheet.cell(row=row_number, column=4).fill = PatternFill(
            fill_type="solid", start_color=color, end_color=color
        )
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:D{max(sheet.max_row, 1)}"
    for letter, width in zip("ABCD", (34, 34, 60, 18)):
        sheet.column_dimensions[letter].width = width
    output = io.BytesIO()
    workbook.save(output)
    data = output.getvalue()
    reloaded = load_workbook(io.BytesIO(data), read_only=True)
    raw_rows = list(reloaded.active.iter_rows(min_row=2, max_col=4, values_only=True))
    reloaded.close()
    # openpyxl has no concept of "blank but present" - an empty string cell round-trips as
    # None. That is expected, not corruption, so None is normalised back to "" for the
    # text columns before the comparison (confidence_score has no such ambiguity).
    got = []
    for document_name, label, value, confidence_score in raw_rows:
        got.append({
            "document_name": document_name or "",
            "label": label or "",
            "value": "" if value is None else value,
            "confidence_score": int(confidence_score) if confidence_score is not None else 0,
        })
    if got != records:
        raise ExportError(f"{description} XLSX integrity check failed after re-reading.")
    return data, len(records)


XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ---------------------------------------------------------------- styling
def inject_css() -> None:
    st.markdown(
        """
        <style>
        :root { --bg:#10141b; --panel:#171c25; --panel2:#1e2430; --line:#2c3545;
                --ink:#eef2f8; --muted:#a4afc1; --accent:#6c8cff; --accent-strong:#4f71f0;
                --ok:#4fd6a8; --warn:#f2c46d; --bad:#ff8585; --focus:#ffd166; }
        .stApp { background: var(--bg); color: var(--ink); }
        .block-container { max-width: 1320px; padding-top: 1rem; padding-bottom: 3rem; }
        header[data-testid="stHeader"] { background: transparent; }
        [data-testid="stSidebar"] { display: none; }

        /* accessibility: always-visible keyboard focus, respect reduced motion */
        button:focus-visible, input:focus-visible, textarea:focus-visible, [role="tab"]:focus-visible,
        [data-baseweb="select"] *:focus-visible { outline: 3px solid var(--focus) !important; outline-offset: 2px !important; }
        @media (prefers-reduced-motion: reduce) { * { transition: none !important; animation: none !important; } }

        /* top bar */
        .brand { display:flex; align-items:center; gap:12px; }
        .brand-mark { width:40px; height:40px; border-radius:10px; display:grid; place-items:center;
                      background: var(--accent-strong); color:#fff; font-weight:800; font-size:19px; }
        .brand-name { font-size:17px; font-weight:700; color:var(--ink); line-height:1.15; }
        .brand-tag { font-size:12.5px; color:var(--muted); margin-top:2px; }
        .chips { display:flex; gap:8px; justify-content:flex-end; flex-wrap:wrap; }
        .chip, .badge { display:inline-flex; align-items:center; gap:5px; font-size:12px; font-weight:600;
                padding:4px 10px; border-radius:999px; border:1px solid var(--line);
                color:var(--muted); background:var(--panel2); white-space:nowrap; }
        .chip.ok, .badge.ok { color:var(--ok); border-color:rgba(79,214,168,.45); }
        .chip.warn, .badge.warn { color:var(--warn); border-color:rgba(242,196,109,.45); }
        .chip.bad, .badge.bad { color:var(--bad); border-color:rgba(255,133,133,.45); }
        .chip.accent { color:#c4d1ff; border-color:rgba(108,140,255,.5); }

        /* navigation */
        button[kind="tertiary"], button[data-testid="stBaseButton-tertiary"] {
            color: var(--muted) !important; font-weight: 600 !important; background: transparent !important;
            border: 1px solid transparent !important; border-radius: 8px !important; }
        button[kind="tertiary"]:hover, button[data-testid="stBaseButton-tertiary"]:hover {
            color: var(--ink) !important; background: var(--panel2) !important; border-color: var(--line) !important; }
        button[kind="primary"], button[data-testid="stBaseButton-primary"] {
            background: var(--accent-strong) !important; border-color: var(--accent-strong) !important;
            color: #fff !important; border-radius: 8px !important; font-weight: 600 !important; }
        button[kind="primary"]:hover, button[data-testid="stBaseButton-primary"]:hover {
            background: #3f60d6 !important; border-color: #3f60d6 !important; }
        button[kind="secondary"], button[data-testid="stBaseButton-secondary"] { border-radius: 8px !important; }
        /* headings */
        .page-head { margin: 6px 0 16px; }
        .page-crumb { font-size:13px; color:var(--muted); }
        .page-title { font-size:26px; font-weight:700; color:var(--ink); margin-top:2px; letter-spacing:-.01em; }
        .page-sub { color:var(--muted); font-size:14px; margin-top:3px; max-width:72ch; }

        /* surfaces */
        [data-testid="stMetric"] { background: var(--panel); border:1px solid var(--line);
            border-radius:10px; padding:12px 16px; }
        [data-testid="stMetricLabel"] p { color: var(--muted) !important; font-size:13px !important; }
        [data-testid="stMetricValue"] { color: var(--ink) !important; }
        div[data-testid="stVerticalBlockBorderWrapper"] { background: var(--panel); border-radius: 10px; }
        [data-testid="stFileUploader"] section { background: var(--panel2); border-radius: 10px; }
        [data-testid="stDataFrame"] { border:1px solid var(--line); border-radius:10px; overflow:hidden; }
        div[data-baseweb="select"] > div, div[data-baseweb="input"] { border-radius: 8px !important; }
        .empty { text-align:center; padding: 44px 20px; border:1px dashed var(--line);
                 border-radius:12px; color:var(--muted); background:var(--panel); }
        .empty h3 { color: var(--ink); margin: 6px 0; font-size:18px; }
        .kv { color:var(--muted); font-size:13px; }

        /* colour-coded confidence filter pills (fixed option order: Green, Yellow, Red) */
        .st-key-review_band div[data-testid="stPills"] button {
            border-width: 1.5px !important; font-weight: 700 !important;
        }
        .st-key-review_band div[data-testid="stPills"] button:nth-child(1) {
            border-color: rgba(79,214,168,.75) !important; color: var(--ok) !important; }
        .st-key-review_band div[data-testid="stPills"] button:nth-child(2) {
            border-color: rgba(242,196,109,.75) !important; color: var(--warn) !important; }
        .st-key-review_band div[data-testid="stPills"] button:nth-child(3) {
            border-color: rgba(255,133,133,.75) !important; color: var(--bad) !important; }
        .st-key-review_band div[data-testid="stPills"] button:nth-child(1)[aria-pressed="true"] {
            background: rgba(79,214,168,.18) !important; }
        .st-key-review_band div[data-testid="stPills"] button:nth-child(2)[aria-pressed="true"] {
            background: rgba(242,196,109,.18) !important; }
        .st-key-review_band div[data-testid="stPills"] button:nth-child(3)[aria-pressed="true"] {
            background: rgba(255,133,133,.18) !important; }

        /* responsive: let column rows wrap on narrow screens */
        @media (max-width: 900px) {
            div[data-testid="stHorizontalBlock"] { flex-wrap: wrap !important; row-gap: .5rem; }
            div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"],
            div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
                min-width: min(100%, 220px) !important; flex: 1 1 220px !important; }
            .chips { justify-content: flex-start; }
            .page-title { font-size: 22px; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


BAND_META = {"Green": ("ok", "✔"), "Yellow": ("warn", "▲"), "Red": ("bad", "✖")}
STATUS_META = {"approved": ("ok", "✔"), "pending": ("warn", "…"), "rejected": ("bad", "✖")}


def badge(text: str, kind: str = "", icon: str = "") -> str:
    inner = f"{esc(icon)} {esc(text)}" if icon else esc(text)
    return f'<span class="badge {esc(kind)}">{inner}</span>'


def band_badge(band: str) -> str:
    kind, icon = BAND_META.get(band, ("", ""))
    return badge(band, kind, icon)


def status_badge(status: str) -> str:
    kind, icon = STATUS_META.get(status, ("", ""))
    return badge(status.capitalize(), kind, icon)


def page_header(title: str, subtitle: str, docket_name: Optional[str]) -> None:
    crumb = f"{esc(docket_name)} › " if docket_name else ""
    st.markdown(
        f'<div class="page-head"><div class="page-crumb">{crumb}{esc(st.session_state.get("nav", ""))}</div>'
        f'<h1 class="page-title">{esc(title)}</h1><div class="page-sub">{esc(subtitle)}</div></div>',
        unsafe_allow_html=True,
    )


def empty_state(title: str, body: str) -> None:
    st.markdown(f'<div class="empty" role="status"><h3>{esc(title)}</h3><div>{esc(body)}</div></div>',
                unsafe_allow_html=True)


def render_error_panel(exc: BaseException, where: str) -> None:
    st.error(f"Something went wrong while loading **{where}**. Your saved data has not been changed. "
             "Try again, or go back to Dockets.", icon="🚫")
    with st.expander("Technical details"):
        st.code("".join(traceback.format_exception(type(exc), exc, exc.__traceback__))[-3000:], language=None)
    retry_col, back_col, _ = st.columns([1, 1, 4])
    retry_col.button("Try again", key="err_retry", **stretch())
    back_col.button("Go to Dockets", key="err_back", on_click=go_to, args=("Dockets",), **stretch())


# ---------------------------------------------------------------- top navigation
def active_docket(dockets: List[sqlite3.Row]) -> Optional[sqlite3.Row]:
    active_id = st.session_state.get("active_docket")
    return next((d for d in dockets if d["id"] == active_id), None)


def render_topbar(dockets: List[sqlite3.Row]) -> Optional[sqlite3.Row]:
    if dockets:
        ids = [d["id"] for d in dockets]
        if st.session_state.get("active_docket") not in ids:
            st.session_state["active_docket"] = ids[0]
    keys = len(key_pool())
    cooling = active_cooldown_count()
    rejected = len(rejected_keys())
    current_docket = active_docket(dockets)

    brand, docket_col, chips_col = st.columns([3, 3, 3], vertical_alignment="center")
    with brand:
        st.markdown(
            '<div class="brand"><div class="brand-mark" aria-hidden="true">P</div><div>'
            '<div class="brand-name">Parchi AI</div>'
            '<div class="brand-tag">Supplier document extraction &amp; verification</div></div></div>',
            unsafe_allow_html=True,
        )
    with docket_col:
        # No dropdown here (removed the stray one): the open docket is shown read-only,
        # and switching dockets happens explicitly on the Dockets tab.
        if current_docket is not None:
            st.markdown(
                '<div class="chips" style="justify-content:center">'
                f'<span class="chip accent" role="status" '
                f'title="Open docket — switch on the Dockets tab">▤ {esc(current_docket["name"])}</span>'
                "</div>",
                unsafe_allow_html=True,
            )
        else:
            st.button("＋ Create your first docket", type="primary", on_click=go_to, args=("Dockets",),
                      **stretch())
    with chips_col:
        extra = ""
        active_job = processing_job(st.session_state.get("active_docket")) if st.session_state.get("active_docket") else None
        if active_job and active_job["status"] in {"queued", "running"}:
            extra += '<span class="chip accent">⏳ Processing in background</span>'
        if cooling:
            extra += f'<span class="chip warn">▲ {cooling} key(s) cooling</span>'
        if rejected:
            extra += f'<span class="chip bad">✖ {rejected} invalid key(s)</span>'
        st.markdown(
            f'<div class="chips"><span class="chip accent">Mode: Gemini</span>'
            f'<span class="chip ok">✔ {keys} key(s) ready</span>{extra}</div>',
            unsafe_allow_html=True,
        )

    current = st.session_state.get("nav", PAGES[0])
    for column, page in zip(st.columns(len(PAGES)), PAGES):
        with column:
            st.button(f"{PAGE_ICONS[page]}  {page}", key=f"nav_{page}",
                      type="primary" if page == current else "tertiary",
                      on_click=go_to, args=(page,), **stretch())
    st.divider()
    return active_docket(dockets)


def need_docket(active: Optional[sqlite3.Row]) -> bool:
    if active is None:
        empty_state("No docket selected",
                    "Create a docket to start organising supplier documents for one shipment, supplier or period.")
        st.button("Go to Dockets", type="primary", on_click=go_to, args=("Dockets",))
        return False
    return True


# ---------------------------------------------------------------- callbacks
@guarded
def create_docket_callback() -> None:
    name = st.session_state.get("new_docket_name", "").strip()
    docket_type = st.session_state.get("new_docket_type", "")
    if not name:
        queue_flash("error", "Enter a docket name first.")
        return
    if docket_type not in DOCKET_TYPES:
        queue_flash("error", "Select a docket type first.")
        return
    docket_id = create_docket(name, docket_type)
    st.session_state["active_docket"] = docket_id
    st.session_state["nav"] = "Workspace"
    st.session_state["new_docket_name"] = ""
    st.session_state["new_docket_type"] = "Select a docket type"
    queue_flash("success", f"Docket “{name}” created.")


def open_docket(docket_id: str) -> None:
    st.session_state["active_docket"] = docket_id
    st.session_state["nav"] = "Workspace"


@guarded
def rename_callback(docket_id: str, input_key: str) -> None:
    try:
        rename_docket(docket_id, st.session_state.get(input_key, ""))
        queue_flash("success", "Docket renamed.")
    except ValueError as exc:
        queue_flash("error", str(exc))


@guarded
def delete_docket_callback(docket_id: str, name: str) -> None:
    delete_docket(docket_id)
    if st.session_state.get("active_docket") == docket_id:
        st.session_state["active_docket"] = None
    queue_flash("success", f"Docket “{name}” and its stored files were permanently deleted.")


@guarded
def delete_document_callback(docket_id: str, document_id: str, filename: str) -> None:
    if delete_document(docket_id, document_id) is not None:
        st.session_state.pop("file_pick", None)
        queue_flash("success", f"“{filename}” and its extracted facts were removed.")
    else:
        queue_flash("warning", "That document was already removed.")


@guarded
def clear_cooldowns_callback() -> None:
    clear_cooldowns()
    queue_flash("success", "Cooldowns cleared.")


@guarded
def approve_green_callback(docket_id: str) -> None:
    queue_flash("success", f"Approved {approve_green(docket_id)} Green fact(s).")


@guarded
def save_review_callback(fact_id: str, value_key: str, status_key: str) -> None:
    try:
        update_fact(fact_id, st.session_state[value_key], st.session_state[status_key], "human review")
        queue_flash("success", "Review saved to history.")
    except ValueError as exc:
        queue_flash("error", str(exc))


# ---------------------------------------------------------------- pages
def page_dockets(dockets: List[sqlite3.Row], active: Optional[sqlite3.Row]) -> None:
    page_header("Dockets", "Each docket groups the supplier documents for one shipment, supplier or period.", None)

    with st.container(border=True):
        name_col, type_col, button_col = st.columns([4, 2, 1], vertical_alignment="bottom")
        name_col.text_input("New docket name", key="new_docket_name",
                            placeholder="e.g. Acme Steel, March deliveries", max_chars=120)
        type_col.selectbox(
            "Docket type",
            ["Select a docket type", *DOCKET_TYPES],
            key="new_docket_type",
            index=0,
        )
        button_col.button("Create docket", type="primary", on_click=create_docket_callback, **stretch())

    if not dockets:
        empty_state("No dockets yet", "Name your first docket above, then upload supplier documents to it.")
        return

    search_col, filter_col = st.columns([3, 1])
    query = search_col.text_input("Search dockets", key="docket_search", placeholder="Filter by name",
                                  label_visibility="collapsed")
    status_filter = filter_col.selectbox("Status", ["All", "Empty", "Needs review", "Reviewed"],
                                         key="docket_status", label_visibility="collapsed")
    rows = [
        d for d in dockets
        if query.lower() in d["name"].lower()
        and (status_filter == "All" or docket_status(d) == status_filter)
    ]
    if not rows:
        st.info("No dockets match the current filters. Clear the search or choose “All”.")
        return

    for docket in rows:
        with st.container(border=True):
            info, stats, actions = st.columns([4, 2, 2], vertical_alignment="center")
            with info:
                st.markdown(f"**{docket['name']}**")
                st.caption(f"Type: {docket['docket_type']}")
                st.caption(f"Created {display_time(docket['created_at'])} · Updated {display_time(docket['updated_at'])}")
            with stats:
                st.caption(f"{docket['document_count']} documents · {docket['fact_count']} facts")
                status = docket_status(docket)
                kind = {"Reviewed": "ok", "Needs review": "warn"}.get(status, "")
                st.markdown(badge(status, kind), unsafe_allow_html=True)
            with actions:
                st.button("Open", key=f"open_{docket['id']}", type="primary",
                          on_click=open_docket, args=(docket["id"],), **stretch())
                st.caption("Rename or delete a docket from its Workspace “Docket actions”.")


def page_workspace(dockets: List[sqlite3.Row], active: Optional[sqlite3.Row]) -> None:
    page_header("Workspace",
                "Keep this docket's supplier documents here: step 1 saves files without "
                "processing them, step 2 runs Gemini extraction on saved documents.",
                active["name"] if active else None)
    if not need_docket(active):
        return
    docket_id = active["id"]
    facts = fetch_facts(docket_id)
    documents = list_documents(docket_id)

    metrics = st.columns(3)
    metrics[0].metric("Documents", len(documents))
    metrics[1].metric("Facts", len(facts))
    metrics[2].metric("Pending review", sum(f["review_status"] == "pending" for f in facts))

    render_upload_panel(docket_id)
    render_process_panel(docket_id)

    st.markdown("#### 3 · Stored documents")
    st.caption("Everything saved in this docket, processing state included. “Saved” documents "
               "are stored but not extracted yet.")
    if not documents:
        empty_state("No documents in this docket",
                    "Save files in step 1 above. They stay available here until you process or remove them.")
        return

    st.dataframe(
        [
            {
                "Document": d["filename"],
                "State": document_state_label(d["status"]),
                "Type": d["document_type"],
                "Pages": d["page_count"],
                "Processed": f"{d['processed_units']}/{d['page_count']}",
                "Completeness": d["completeness_status"],
                "Size": fmt_size(d["size_bytes"]),
                "Extraction": d["extraction_mode"],
                "Added": display_time(d["created_at"]),
                "File on disk": "✔" if resolve_stored_safe(d["storage_path"]) else "missing",
            }
            for d in documents
        ],
        hide_index=True,
        **stretch(),
    )

    st.markdown("#### 4 · Document actions")
    by_id = {d["id"]: d for d in documents}
    chosen_id = st.selectbox(
        "Choose a document",
        list(by_id),
        format_func=lambda i: f"{by_id[i]['filename']} — {document_state_label(by_id[i]['status'])}",
        key="file_pick",
        help="Select the document you want to inspect, download, process or remove.",
    )
    chosen = by_id[chosen_id]
    with st.container(border=True):
        badge_col, hint_col = st.columns([3, 1])
        badge_col.markdown(
            f"**{esc(chosen['filename'])}** · {document_state_badge(chosen['status'])}",
            unsafe_allow_html=True,
        )
        hint_col.caption(f"{document_state_label(chosen['status'])}")
        st.caption(f"{chosen['mime_type']} · {chosen['page_count']} page(s)/sheet(s) · "
                   f"{fmt_size(chosen['size_bytes'])} · added {display_time(chosen['created_at'])}")
        data, problem = read_verified(chosen["storage_path"], chosen["sha256"])
        action_cols = st.columns([2, 1, 2], vertical_alignment="center")
        with action_cols[0]:
            if chosen["status"] in PROCESSABLE_STATES:
                if st.button(
                    "Process this document",
                    type="primary",
                    key=f"process_one_{chosen_id}",
                    disabled=processing_active(docket_id),
                    help="Run Gemini extraction on the stored original now.",
                    **stretch(),
                ):
                    queue_saved_documents(docket_id, [chosen_id])
                    st.rerun(scope="app")
            elif chosen["status"] in ("queued", "processing", "saving"):
                st.button("Process this document", disabled=True,
                          help="This document is already queued or processing.", **stretch())
            else:
                st.button("Process this document", disabled=True,
                          help="Already processed. Remove it and save it again to re-extract.",
                          **stretch())
        with action_cols[1]:
            if data is not None:
                st.download_button("Download", data, file_name=chosen["filename"],
                                   mime=chosen["mime_type"], **stretch())
            else:
                st.button("Unavailable", disabled=True, help=problem, **stretch())
        with action_cols[2]:
            with st.popover("Remove document", **stretch()):
                st.markdown(f"Remove **{chosen['filename']}**?")
                st.caption("Deletes the stored original, its pages and every fact extracted from it. "
                           "Other documents are unaffected.")
                confirm_key = f"confirm_doc_{chosen_id}"
                st.checkbox("I understand", key=confirm_key)
                st.button("Remove permanently", type="primary", key=f"remove_doc_{chosen_id}",
                          disabled=not st.session_state.get(confirm_key, False),
                          on_click=delete_document_callback, args=(docket_id, chosen_id, chosen["filename"]))
        if problem:
            st.error(problem, icon="🚫")

    with st.expander("5 · Docket actions"):
        st.caption("Rename or delete this docket. Deleting removes all of its documents and stored files.")
        rename_key = f"ws_rename_{docket_id}"
        rename_col, rename_btn = st.columns([4, 1], vertical_alignment="bottom")
        rename_col.text_input("Docket name", value=active["name"], key=rename_key, max_chars=120)
        rename_btn.button("Save name", key=f"ws_save_{docket_id}", on_click=rename_callback,
                          args=(docket_id, rename_key), **stretch())
        confirm_key = f"ws_confirm_delete_{docket_id}"
        st.checkbox("I understand deleting this docket cannot be undone", key=confirm_key)
        st.button("Delete this docket permanently", type="primary", key=f"ws_delete_{docket_id}",
                  disabled=not st.session_state.get(confirm_key, False),
                  on_click=delete_docket_callback, args=(docket_id, active["name"]))


def fact_label(fact: sqlite3.Row) -> str:
    return (f"{fact['filename']} · p.{fact['page_number']} · {fact['label']} · "
            f"{fact['confidence_band']} · {fact['review_status']}")


@st.cache_data(show_spinner=False, max_entries=48)
def render_pdf_page(path: str, page_number: int, mtime: float) -> bytes:
    with fitz.open(path) as document:
        index = min(max(1, page_number), document.page_count) - 1
        return document[index].get_pixmap(matrix=fitz.Matrix(1.5, 1.5)).tobytes("png")


def source_preview(fact: sqlite3.Row) -> None:
    """Show the original page next to the extracted value so it can be verified."""
    st.markdown("**Original document**")
    path = resolve_stored_safe(fact["storage_path"])
    if path is None:
        st.warning("The original file is missing from storage, so it cannot be previewed.", icon="⚠️")
        return
    extension = extension_of(path.name)
    try:
        if extension in ("png", "jpg", "jpeg"):
            st.image(path.read_bytes(), caption=f"{fact['filename']}", **stretch())
        elif extension == "pdf" and fitz is not None:
            image = render_pdf_page(str(path), fact["page_number"], path.stat().st_mtime)
            st.image(image, caption=f"{fact['filename']} · page {fact['page_number']}", **stretch())
        elif extension == "txt":
            st.code(decode_text(path.read_bytes())[:4000], language=None)
        elif extension == "csv":
            st.dataframe(pd.read_csv(io.StringIO(decode_text(path.read_bytes())), dtype=str,
                                      keep_default_na=False).head(200), hide_index=True, **stretch())
        elif extension == "xlsx":
            with connect() as connection:
                page = connection.execute(
                    "SELECT sheet_name, text FROM pages WHERE document_id = ? AND page_number = ?",
                    (fact["document_id"], fact["page_number"]),
                ).fetchone()
            if page and page["text"]:
                rows = [line.split(" | ") for line in page["text"].splitlines()]
                st.caption(f"Sheet: {page['sheet_name'] or '(unnamed)'}")
                st.dataframe(rows, hide_index=True, **stretch())
            else:
                st.info("This sheet has no readable rows to preview.")
        else:
            st.info("No preview available for this file type.")
    except Exception as exc:  # noqa: BLE001
        LOG.exception("preview failed")
        st.warning(f"The preview could not be rendered ({exc.__class__.__name__}). "
                   "Download the original from the Workspace instead.", icon="⚠️")


def page_review(dockets: List[sqlite3.Row], active: Optional[sqlite3.Row]) -> None:
    page_header("Review", "Verify extracted values against the source evidence. Every change is recorded.",
                active["name"] if active else None)
    if not need_docket(active):
        return
    docket_id = active["id"]
    facts = fetch_facts(docket_id)
    if not facts:
        empty_state("Nothing to review yet", "Process documents in the Workspace to extract facts.")
        return

    documents = sorted({f["filename"] for f in facts})
    doc_col, band_col, status_col = st.columns(3)
    doc_filter = doc_col.selectbox("Document", ["All documents"] + documents, key="review_doc")
    with band_col:
        st.markdown(
            "**Confidence band** "
            + " ".join(
                badge(f"{band} — {meaning}", kind, icon)
                for band, (kind, icon), meaning in zip(
                    BANDS, BAND_META.values(), ("high", "medium", "low")
                )
            ),
            unsafe_allow_html=True,
        )
        band_filter = st.pills(
            "Confidence band",
            list(BANDS),
            selection_mode="multi",
            default=list(BANDS),
            key="review_band",
            label_visibility="collapsed",
            help="Green = high confidence (score > 80), Yellow = medium (66–80), Red = low (≤ 65).",
        ) or []
    status_filter = status_col.selectbox("Review status", ["All"] + list(REVIEW_STATUSES), key="review_status")

    visible = [
        f for f in facts
        if (doc_filter == "All documents" or f["filename"] == doc_filter)
        and f["confidence_band"] in band_filter
        and (status_filter == "All" or f["review_status"] == status_filter)
    ]
    green_pending = sum(f["confidence_band"] == "Green" and f["review_status"] == "pending" for f in facts)
    summary_col, bulk_col = st.columns([3, 1], vertical_alignment="center")
    band_counts = {band: sum(f["confidence_band"] == band for f in visible) for band in BANDS}
    summary_col.markdown(
        f"Showing **{len(visible)}** of {len(facts)} facts: "
        + " ".join(badge(f"{band} {count}", BAND_META[band][0], BAND_META[band][1])
                   for band, count in band_counts.items()),
        unsafe_allow_html=True,
    )
    bulk_col.button(f"Approve Green pending ({green_pending})", on_click=approve_green_callback,
                    args=(docket_id,), disabled=green_pending == 0, **stretch())

    if not visible:
        st.info("No facts match the current filters. Widen the confidence or status filters.")
        return

    by_id = {f["id"]: f for f in visible}
    fact_id = st.selectbox("Fact", list(by_id), format_func=lambda i: fact_label(by_id[i]), key="review_fact")
    fact = by_id[fact_id]

    editor_col, preview_col = st.columns([3, 2])
    with editor_col:
        with st.container(border=True):
            st.markdown(f"**{fact['label']}** · concept `{fact['concept']}`")
            st.markdown(band_badge(fact["confidence_band"]) + " " + status_badge(fact["review_status"]),
                        unsafe_allow_html=True)
            reasons = ", ".join(json.loads(fact["confidence_reasons"])) or "none"
            st.caption(f"Score {fact['confidence_score']} · Reasons: {reasons} · "
                       f"Original value: {fact['original_value'] or '(empty)'}")
            st.markdown("**Source evidence**")
            st.code(fact["source_evidence"] or "None captured", language=None)
            # The version suffix makes widget state reset after a save or bulk approval.
            version = f"{fact_id}_{fact['updated_at']}"
            value_key, status_key = f"edit_{version}", f"status_{version}"
            st.text_area("Value", fact["value"], key=value_key)
            st.selectbox("Review status", list(REVIEW_STATUSES),
                         index=list(REVIEW_STATUSES).index(fact["review_status"]), key=status_key)
            st.button("Save review", type="primary", on_click=save_review_callback,
                      args=(fact_id, value_key, status_key))
    with preview_col:
        with st.container(border=True):
            source_preview(fact)

    history = fact_history(fact_id)
    with st.expander(f"Change history ({len(history)})"):
        if history:
            st.dataframe(
                [{"When": display_time(h["changed_at"]), "Old value": h["old_value"],
                  "New value": h["new_value"], "Status": f"{h['old_status']} → {h['new_status']}",
                  "Reason": h["reason"]} for h in history],
                hide_index=True, **stretch(),
            )
        else:
            st.caption("No changes recorded yet.")

FUZZY_MIN_SCORE = 70


def normalize_label(label: Any) -> str:
    """Comparison form for label matching: trimmed and case-insensitive."""
    return str(label or "").strip().casefold()


def exact_template_label(source_label: str, requested_labels: Iterable[str]) -> Optional[str]:
    """The template label equal to `source_label` ignoring case and leading/trailing
    whitespace, else None."""
    normalized = normalize_label(source_label)
    return next(
        (str(label) for label in requested_labels if normalize_label(label) == normalized),
        None,
    )


def unique_labels(records: Iterable[Any]) -> List[str]:
    """Every distinct label in first-seen order, original spelling preserved.
    Each result is one unique label - not one extracted value."""
    seen: set = set()
    labels: List[str] = []
    for record in records:
        raw = record.get("label") if isinstance(record, dict) else record
        label = clean_text(raw).strip()
        normalized = normalize_label(label)
        if normalized and normalized not in seen:
            seen.add(normalized)
            labels.append(label)
    return labels


def _dedupe_labels(labels: Iterable[str]) -> List[str]:
    """Collapse repeats of the same label (any casing/spacing) to the first spelling."""
    seen: set = set()
    out: List[str] = []
    for label in labels:
        clean = clean_text(label).strip()
        normalized = normalize_label(clean)
        if normalized and normalized not in seen:
            seen.add(normalized)
            out.append(clean)
    return out


def fuzzy_suggestions(
    source: str, candidates: List[str], min_score: int = FUZZY_MIN_SCORE, limit: int = 5
) -> List[Dict[str, Any]]:
    """RapidFuzz suggestions for one extracted label against template labels.
    Only candidates scoring >= min_score are returned, best first. Exact matches are
    resolved separately and never reach this function."""
    query = normalize_label(source)
    if not query or not candidates:
        return []
    scored: List[Tuple[str, float]] = []
    if rapidfuzz is not None:
        matches = rapidfuzz.process.extract(
            query,
            [normalize_label(candidate) for candidate in candidates],
            scorer=rapidfuzz.fuzz.WRatio,
            score_cutoff=float(min_score),
            limit=limit,
        )
        scored = [(candidates[index], float(score)) for _, score, index in matches]
    else:  # graceful degradation; install rapidfuzz for the specified behaviour
        import difflib

        for candidate in candidates:
            ratio = difflib.SequenceMatcher(None, query, normalize_label(candidate)).ratio() * 100
            if ratio >= min_score:
                scored.append((candidate, ratio))
    scored.sort(key=lambda item: (-item[1], item[0]))
    return [{"label": label, "score": score} for label, score in scored[:limit]]


def template_sections_key(docket_id: str, scope: str) -> str:
    """Session key for {document_name: [submitted labels]} in the pill workflow.
    `scope` isolates Create ("create") from Edit ("edit_<template id>") so their
    selections can never leak into each other."""
    return f"template_sections_{docket_id}_{scope}"


def _document_label_picker(docket_id: str, metadata: List[dict], sections: Dict[str, List[str]], scope: str) -> None:
    """Step 1 (shared by create/edit): choose ONE document, see ONLY its unique
    labels as pills, submit them below. Switching documents never disturbs the
    pills of another document or any submitted section."""
    document_names = sorted({record["document_name"] for record in metadata}, key=str.casefold)
    with st.container(border=True):
        st.markdown("**1 · Pick labels from one document**")
        st.caption(f"{len(document_names)} document(s) with extracted labels are available.")
        document_name = st.selectbox(
            "Uploaded document",
            ["Select a document", *document_names],
            key=f"template_doc_{docket_id}_{scope}",
            help="Labels appear only after a document is chosen; each pill is one unique label.",
        )
        if document_name == "Select a document":
            st.info("Choose a document to display its unique extracted labels.")
            return
        document_labels = unique_labels(
            [record for record in metadata if record["document_name"] == document_name]
        )
        submitted = {normalize_label(label) for label in sections.get(document_name, [])}
        available = [label for label in document_labels if normalize_label(label) not in submitted]
        if not available:
            st.success("Every label from this document has already been submitted.", icon="✅")
            return
        st.caption(f"{len(available)} unique label(s) available from “{document_name}”. "
                   "Selected pills are highlighted; click again to unselect.")
        # The key embeds the submission count, so the widget starts fresh whenever
        # this document's section changes. st.pills renders selected pills in the
        # accent colour and unselected ones outlined, so selection state is visual.
        picked = st.pills(
            f"Labels from “{document_name}”",
            available,
            selection_mode="multi",
            key=f"template_pills_{docket_id}_{scope}_{key_id(document_name)}_{len(sections.get(document_name, []))}",
            width="stretch",
        ) or []
        if st.button(
            "Submit labels",
            type="primary",
            key=f"template_submit_{docket_id}_{scope}",
            disabled=not picked,
            **stretch(),
        ):
            # Duplicate submissions of the same label (any casing) are blocked.
            sections[document_name] = _dedupe_labels([*sections.get(document_name, []), *picked])
            st.session_state[template_sections_key(docket_id, scope)] = sections
            st.rerun()


def _submitted_sections_editor(docket_id: str, sections: Dict[str, List[str]], scope: str) -> None:
    """Step 2 (shared): one bordered section per submitted document, filename as
    heading, selected labels as editable pills beneath it. Selections stay visible
    until explicitly removed."""
    st.markdown("#### 2 · Submitted document sections")
    if not sections:
        st.info("Pick labels above and click Submit labels. Every submitted document gets its own section here.")
        return
    total_labels = sum(len(section_labels) for section_labels in sections.values())
    st.caption(f"{len(sections)} document section(s) · {total_labels} label(s) currently selected. "
               "Click a highlighted pill to remove it from the section; remove a whole section "
               "with its Remove section button.")
    for document, section_labels in list(sections.items()):
        with st.container(border=True):
            head_col, clear_col = st.columns([5, 1], vertical_alignment="center")
            head_col.markdown(f"**{document}**")
            head_col.caption(f"{len(section_labels)} label(s) selected")
            if clear_col.button("Remove section", key=f"template_drop_{docket_id}_{scope}_{key_id(document)}",
                                help=f"Remove the whole “{document}” section from this template.",
                                **stretch()):
                del sections[document]
                st.session_state[template_sections_key(docket_id, scope)] = sections
                st.rerun()
            kept = st.pills(
                f"Labels selected for {document}",
                section_labels,
                selection_mode="multi",
                default=section_labels,
                key=f"template_section_{docket_id}_{scope}_{key_id(document)}_{len(section_labels)}",
                width="stretch",
            )
            kept = list(kept) if kept is not None else list(section_labels)
            if kept != list(section_labels):
                if kept:
                    sections[document] = kept
                else:
                    del sections[document]
                st.session_state[template_sections_key(docket_id, scope)] = sections
                st.rerun()


def create_template_mode(active: sqlite3.Row, metadata: List[dict]) -> None:
    """Create a new template purely by selecting labels per document. No similarity
    scoring, no automatic mapping, no mapping writes; Main Metadata untouched.
    Saving stays on this page: it never switches to Apply."""
    docket_id = active["id"]
    scope = "create"
    if st.session_state.pop(f"template_create_reset_{docket_id}", False):
        # Clear the name box before its widget is rendered (safe: nothing rendered yet).
        st.session_state.pop(f"template_name_{docket_id}", None)
    sections: Dict[str, List[str]] = st.session_state.setdefault(template_sections_key(docket_id, scope), {})
    _document_label_picker(docket_id, metadata, sections, scope)
    _submitted_sections_editor(docket_id, sections, scope)
    st.markdown("#### 3 · Save template")
    all_labels = _dedupe_labels(label for labels in sections.values() for label in labels)
    name = st.text_input("Template name", key=f"template_name_{docket_id}", max_chars=120)
    st.caption(f"{len(all_labels)} unique label(s) will be saved.")
    if st.button(
        "Save Template",
        type="primary",
        key=f"template_save_{docket_id}",
        disabled=not all_labels or not name.strip(),
        **stretch(),
    ):
        try:
            template_id = save_template(docket_id, name, all_labels)
        except ValueError as exc:
            st.error(str(exc), icon="🚫")
            return
        for label in all_labels:
            save_mapping(docket_id, template_id, label, label, "direct")
        st.session_state.pop(template_sections_key(docket_id, scope), None)
        st.session_state[f"template_create_reset_{docket_id}"] = True
        queue_flash("success", f"Template “{clean_text(name).strip()}” saved. Main Metadata was not changed.")
        st.rerun()


def edit_template_mode(active: sqlite3.Row, metadata: List[dict]) -> None:
    """Edit an existing template with the same document-based selection workflow.
    Saving updates labels in place (kept labels keep their exact spelling) and
    prunes mappings pointing at removed labels. Mappings are never edited here."""
    docket_id = active["id"]
    templates = list_templates(docket_id)
    if not templates:
        st.info("No saved templates yet. Create one first.")
        return
    names = [row["name"] for row in templates]
    selected = st.selectbox("Template to edit", names, key=f"template_edit_pick_{docket_id}")
    template = next(row for row in templates if row["name"] == selected)
    template_id = template["id"]
    scope = f"edit_{template_id}"

    state_key = f"template_edit_state_{docket_id}"
    state = st.session_state.get(state_key)
    if not isinstance(state, dict) or state.get("template_id") != template_id:
        state = {"template_id": template_id, "labels": json.loads(template["labels_json"])}
        st.session_state[state_key] = state
        st.session_state.pop(template_sections_key(docket_id, scope), None)

    sections: Dict[str, List[str]] = st.session_state.setdefault(template_sections_key(docket_id, scope), {})

    with st.container(border=True):
        st.markdown("**Current template labels**")
        st.caption("Deselect a pill to remove that label when you save. Labels are never renamed here.")
        kept = st.pills(
            "Current labels in this template — deselect a pill to drop it on save",
            state["labels"],
            selection_mode="multi",
            default=state["labels"],
            key=f"template_edit_labels_{docket_id}_{template_id}_{len(state['labels'])}",
            width="stretch",
        )
        kept = list(kept) if kept is not None else list(state["labels"])
        if kept != list(state["labels"]):
            state["labels"] = kept
            st.rerun()

    _document_label_picker(docket_id, metadata, sections, scope)
    _submitted_sections_editor(docket_id, sections, scope)

    st.markdown("#### 3 · Save changes")
    combined = _dedupe_labels([*state["labels"], *(label for labels in sections.values() for label in labels)])
    name = st.text_input("Template name", value=template["name"], key=f"template_edit_name_{docket_id}", max_chars=120)
    st.caption(f"{len(combined)} unique label(s) will be saved.")
    save_col, delete_col = st.columns(2)
    if save_col.button("Save changes", type="primary", key=f"template_edit_save_{docket_id}", disabled=not combined, **stretch()):
        try:
            template_id = save_template(docket_id, name, combined, template_id=template_id)
        except ValueError as exc:
            st.error(str(exc), icon="🚫")
            return
        for label in combined:
            if label not in state["labels"]:
                save_mapping(docket_id, template_id, label, label, "direct")
        st.session_state.pop(state_key, None)
        st.session_state.pop(template_sections_key(docket_id, scope), None)
        queue_flash("success", "Template updated. Main Metadata was not changed.")
        st.rerun()
    with delete_col.popover("Delete template", **stretch()):
        st.markdown(f"Delete **{template['name']}**?")
        st.caption("Also removes this template's mappings and rejected suggestions. Main Metadata is untouched.")
        confirm_key = f"template_delete_confirm_{template_id}"
        st.checkbox("I understand this cannot be undone", key=confirm_key)
        if st.button("Delete permanently", type="primary", disabled=not st.session_state.get(confirm_key, False), **stretch()):
            delete_template(docket_id, template_id)
            st.session_state.pop(state_key, None)
            st.session_state.pop(template_sections_key(docket_id, scope), None)
            queue_flash("success", f"Template “{template['name']}” deleted.")
            st.rerun()


def apply_template_mode(active: sqlite3.Row, metadata: List[dict]) -> None:
    """Apply a saved template to Main Metadata labels.

    - Exact matches (case-insensitive, trimmed) map automatically, no interaction.
    - Fuzzy suggestions (RapidFuzz >= 70%) require explicit Approve/Reject each.
    - Unmatched labels can be mapped manually or explicitly added to the template.
    Applying never silently rewrites saved labels; Main Metadata is never modified;
    every decision persists in the database across restarts."""
    docket_id = active["id"]
    templates = list_templates(docket_id)
    if not templates:
        st.info("No saved templates yet. Create one first.")
        return
    names = [row["name"] for row in templates]
    apply_pick_key = f"template_apply_pick_{docket_id}"
    selected = st.selectbox("Saved template", names, key=apply_pick_key)
    template = next(row for row in templates if row["name"] == selected)
    template_id = template["id"]
    template_labels = json.loads(template["labels_json"])

    document_names = sorted({record["document_name"] for record in metadata}, key=str.casefold)
    chosen_documents = st.multiselect(
        "Newly extracted document(s) to compare against this template",
        document_names,
        default=document_names,
        key=f"template_apply_docs_{docket_id}_{template_id}",
        help="Only labels from the selected documents are compared below. "
             "Approved mappings still apply to every matching occurrence in Main Metadata.",
    )
    scoped_metadata = [record for record in metadata if record["document_name"] in chosen_documents]
    if not chosen_documents:
        st.info("Select at least one document to compare its labels against this template.")
        return

    labels = unique_labels(scoped_metadata)
    current = approved_mappings(docket_id, template_id)
    details = mapping_details(docket_id, template_id)
    rejected = rejected_suggestions(docket_id, template_id)

    exact_pairs: List[Tuple[str, str]] = []
    manual_pairs: List[Tuple[str, str, str]] = []
    fuzzy: Dict[str, List[Dict[str, Any]]] = {}
    unmatched: List[str] = []

    for source in labels:
        exact_target = exact_template_label(source, template_labels)
        if exact_target is not None:
            if source not in current:
                # Automatic, idempotent, no user interaction. A deliberate human
                # mapping to a different target is respected and shown as such.
                save_mapping(docket_id, template_id, source, exact_target, "exact")
                current[source] = exact_target
            if current[source] == exact_target:
                exact_pairs.append((source, exact_target))
            else:
                manual_pairs.append((source, current[source], details.get(source, {}).get("method", "manual")))
            continue
        if source in current:
            manual_pairs.append((source, current[source], details.get(source, {}).get("method", "manual")))
            continue
        options = [
            match for match in fuzzy_suggestions(source, template_labels)
            if (source, match["label"]) not in rejected
        ]
        if options:
            fuzzy[source] = options
        else:
            unmatched.append(source)

    st.caption(
        f"{len(labels)} unique extracted label(s) from {len(chosen_documents)} document(s), "
        f"compared against {len(template_labels)} template label(s)."
    )

    with st.container(border=True):
        st.markdown("**Current mapping decisions**")
        st.caption("Every approved mapping for this template stays in effect — across the "
                   "Viewer, Export and future runs — until you unmap it here.")
        if details:
            st.dataframe(
                [
                    {
                        "Extracted label": source,
                        "Template label": info["target"],
                        "Decided by": {
                            "exact": "Exact match (automatic)",
                            "user-approved": "Approved by you",
                            "manual": "Mapped by you",
                        }.get(info["method"], info["method"]),
                    }
                    for source, info in sorted(details.items(), key=lambda item: item[0].casefold())
                ],
                hide_index=True, **stretch(),
            )
        else:
            st.caption("No mapping decisions yet — approve a fuzzy suggestion or map an "
                       "unmatched label below to create one.")

    with st.container(border=True):
        st.markdown("**Exact matches (mapped automatically)**")
        if exact_pairs:
            for source, target in exact_pairs:
                st.success(f"`{source}` → `{target}`", icon="✅")
        else:
            st.caption("No exact matches yet.")

    with st.container(border=True):
        st.markdown(f"**Fuzzy suggestions (≥{FUZZY_MIN_SCORE}% similarity — approval required)**")
        if fuzzy:
            for source, options in fuzzy.items():
                st.markdown(f"**{source}**")
                for match in options:
                    target = match["label"]
                    pid = key_id(f"{source}\u241f{target}")
                    approve_col, reject_col, info_col = st.columns([1, 1, 3], vertical_alignment="center")
                    if approve_col.button("Approve", key=f"tpl_approve_{docket_id}_{template_id}_{pid}", width="stretch"):
                        save_mapping(docket_id, template_id, source, target, "user-approved")
                        st.rerun()
                    if reject_col.button("Reject", key=f"tpl_reject_{docket_id}_{template_id}_{pid}", width="stretch"):
                        reject_suggestion(docket_id, template_id, source, target)
                        delete_mapping(docket_id, template_id, source, target)  # Reject also undoes an approval of this pair.
                        st.rerun()
                    info_col.caption(f"`{target}` · {match['score']:.0f}%")
        else:
            st.caption("No suggestions at or above the similarity threshold.")

    with st.container(border=True):
        st.markdown("**Unmatched labels**")
        if unmatched:
            st.caption("Map each label to a template label, or explicitly add it to the template. Nothing is discarded.")
            for source in unmatched:
                c1, c2, c3, c4 = st.columns([2, 2, 1, 1.4], vertical_alignment="center")
                c1.markdown(f"`{source}`")
                target = c2.selectbox(
                    f"Map “{source}” to",
                    template_labels,
                    key=f"tpl_manual_t_{docket_id}_{template_id}_{key_id(source)}",
                )
                if c3.button("Map", key=f"tpl_manual_b_{docket_id}_{template_id}_{key_id(source)}", width="stretch"):
                    save_mapping(docket_id, template_id, source, target, "manual")
                    st.rerun()
                if c4.button("Add to template", key=f"tpl_add_{docket_id}_{template_id}_{key_id(source)}",
                             help="Explicitly extend this template with the label; it will then match exactly.",
                             width="stretch"):
                    append_template_labels(docket_id, template_id, [source])
                    queue_flash("success", f"“{source}” added to template “{template['name']}”.")
                    st.rerun()
        else:
            st.caption("Every extracted label is exactly matched, mapped, or has a pending suggestion.")

    with st.expander("All manual / fuzzy-approved mappings for this template"):
        if manual_pairs:
            for source, target, method in manual_pairs:
                row_c, btn_c = st.columns([4, 1], vertical_alignment="center")
                row_c.markdown(f"`{source}` → `{target}` · {method}")
                if btn_c.button("Unmap", key=f"tpl_unmap_{docket_id}_{template_id}_{key_id(source)}", width="stretch"):
                    delete_mapping(docket_id, template_id, source, target)
                    st.rerun()
        else:
            st.caption("No manual or fuzzy-approved mappings yet.")


ORGANIZE_MODES = ("Create new template", "Apply existing template")
ORGANIZE_MODE_HELP = {
    "Create new template": (
        "Build a template purely from labels you pick document by document in steps 1–3 "
        "below. Nothing is mapped yet and Main Metadata is never changed."
    ),
    "Apply existing template": (
        "Compare the extracted labels with a saved template. Exact matches map automatically, "
        "fuzzy suggestions need your approval, and every mapping decision stays visible below. "
        "Main Metadata values are never modified."
    ),
}


def page_organize(dockets: List[sqlite3.Row], active: Optional[sqlite3.Row]) -> None:
    page_header("Organize", "Create, edit or apply reusable client templates. "
                "Main Metadata is never modified by template operations.", active["name"] if active else None)
    if not need_docket(active):
        return
    metadata = metadata_records(active["id"])
    if not metadata:
        st.info("Process documents first. Save files in Workspace step 1, run extraction in "
                "step 2, then the Main Metadata File will be available here.")
        return
    overview = st.columns(3)
    overview[0].metric("Documents with labels", len({record["document_name"] for record in metadata}))
    overview[1].metric("Unique extracted labels", len(unique_labels(metadata)))
    overview[2].metric("Saved templates", len(list_templates(active["id"])))
    st.caption("Pick one workflow; only its controls appear below. "
               "Your submitted selections and mapping decisions remain visible inside each mode.")
    mode = st.radio(
        "Choose an operation",
        list(ORGANIZE_MODES),
        horizontal=True,
        key="org_mode",
    )
    with st.container(border=True):
        st.markdown(f"**{mode}**")
        st.caption(ORGANIZE_MODE_HELP[mode])
    if mode == "Create new template":
        create_template_mode(active, metadata)
    else:
        apply_template_mode(active, metadata)
    # Rename, edit or delete a saved template. Kept apart from both workflows above.
    with st.expander("Manage saved templates (rename, edit labels, delete)", expanded=False):
        edit_template_mode(active, metadata)
def confidence_score_style(value: object) -> str:
    """Cell colouring for confidence scores — identical thresholds to the bands
    used everywhere else (Green > 80, Yellow 66–80, Red <= 65)."""
    band = band_for(clamp_confidence(value, 50))
    if band == "Green":
        return "background-color: #1f7a4d; color: #f0fff7; font-weight: 700"
    if band == "Yellow":
        return "background-color: #8a6d1d; color: #fff8d6; font-weight: 700"
    return "background-color: #8c3540; color: #fff1f2; font-weight: 700"


VIEWER_COLUMNS = ["Document Name", "Label", "Value", "Confidence Score"]


def viewer_frame(rows: List[dict]) -> None:
    """Render metadata records with stable, human-readable headers (never raw
    snake_case keys) and colour-coded confidence cells."""
    frame = pd.DataFrame(
        [
            {
                "Document Name": record["document_name"],
                "Label": record["label"],
                "Value": record["value"],
                "Confidence Score": clamp_confidence(record["confidence_score"], 50),
            }
            for record in rows
        ],
        columns=VIEWER_COLUMNS,
    )
    st.dataframe(
        frame.style.map(confidence_score_style, subset=["Confidence Score"]),
        hide_index=True,
        width="stretch",
        height=min(720, max(180, 44 + 36 * len(frame))),
    )


def page_viewer(dockets: List[sqlite3.Row], active: Optional[sqlite3.Row]) -> None:
    page_header(
        "Viewer",
        "View the complete Main Metadata file or a Client Template as a table — duplicates, "
        "conflicting values, document names and confidence scores are all preserved.",
        active["name"] if active else None,
    )
    if not need_docket(active):
        return
    docket_id = active["id"]
    try:
        with st.spinner("Loading extracted metadata…"):
            metadata = metadata_records(docket_id)
            templates = list_templates(docket_id)
    except Exception:  # noqa: BLE001
        LOG.exception("viewer could not load metadata")
        st.error("The extracted metadata could not be loaded. Your saved data has not been "
                 "changed. Try again; if the problem persists, run the integrity check in "
                 "Settings.", icon="🚫")
        return
    if not metadata:
        empty_state("No metadata available", "Save and process documents in the Workspace first.")
        return

    view_options = ["Main Metadata"]
    if templates:
        view_options.append("Client Template")
    selected_view = st.selectbox(
        "Table to view",
        view_options,
        key="viewer_table",
        help="Main Metadata lists every extracted record. Client Template lists every record "
             "mapped onto one saved template.",
    )

    if selected_view == "Main Metadata":
        rows = metadata
        st.caption(f"All extracted records · {len(rows)} row(s), including duplicates and "
                   "conflicting values.")
    else:
        template_names = [template["name"] for template in templates]
        selected_name = st.selectbox("Client template", template_names, key="viewer_template")
        template = next(row for row in templates if row["name"] == selected_name)
        try:
            raw_labels = json.loads(template["labels_json"])
            if not isinstance(raw_labels, list):
                raise ValueError("labels_json must be a list")
        except ValueError:
            st.error(f"The template “{template['name']}” could not be read (its label list is "
                     "damaged). Other templates and the Main Metadata view still work.", icon="🚫")
            return
        with st.spinner("Applying template mappings…"):
            rows, missing = client_template_records(docket_id, template)
        if missing:
            st.warning(
                "No extracted record mapped to: " + ", ".join(missing) + ". "
                "If these values exist under a different spelling, approve the mapping in "
                "Organize → Apply existing template.",
                icon="⚠️",
            )
        if not rows:
            st.info("No records to show for this template yet. Approve label mappings in "
                    "Organize → Apply existing template; saved mappings are shown here.",
                    icon="ℹ️")
        st.caption(f"Mapped records · {len(rows)} row(s); duplicate and conflicting values each "
                   "remain a separate row with its own document name and confidence score.")

    viewer_frame(rows)


def page_validation(dockets: List[sqlite3.Row], active: Optional[sqlite3.Row]) -> None:
    page_header("Validation", "Deterministic checks across completeness, evidence, consistency and arithmetic.",
                active["name"] if active else None)
    if not need_docket(active):
        return
    with st.spinner("Running validation checks…"):
        report = validation_report(active["id"])
    counts = {band: sum(r["severity"] == band for r in report) for band in BANDS}
    cols = st.columns(3)
    cols[0].metric("Red issues", counts["Red"])
    cols[1].metric("Yellow issues", counts["Yellow"])
    cols[2].metric("Green", counts["Green"])
    if counts["Red"] == 0 and counts["Yellow"] == 0:
        st.success("No unresolved issues. This docket is ready to export.", icon="✅")
    st.dataframe(
        [{"Severity": f"{BAND_META[r['severity']][1]} {r['severity']}", "Check": r["check"],
          "Details": r["message"]} for r in report],
        hide_index=True, **stretch(),
    )

def page_export(dockets: List[sqlite3.Row], active: Optional[sqlite3.Row]) -> None:
    page_header("Export", "Download the Main Metadata File or a Client Template. Files are generated here from "
                          "saved mappings. This page never changes them.", active["name"] if active else None)
    if not need_docket(active):
        return
    if active["fact_count"] == 0:
        empty_state("Nothing to export yet", "Process documents in the Workspace to extract facts first.")
        return
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d")
    slug = to_concept(active["name"]) or "docket"
    formats = ["JSON", "PSV", "XLSX"]
    mime_by_format = {"JSON": "application/json", "PSV": "text/plain; charset=utf-8", "XLSX": XLSX_MIME}

    def export_panel(title: str, caption: str, records: List[dict], missing: List[str], prefix: str) -> None:
        st.markdown(f"**{title}**")
        st.caption(caption)
        if missing:
            st.warning("Not mapped yet: " + ", ".join(missing) + ". Map them in Organize → Apply existing "
                       "template. Until then they are left out of this file.", icon="⚠️")
        chosen = st.selectbox("Format", formats, key=f"{prefix}_format")
        try:
            if chosen == "JSON":
                data, count = build_metadata_export(records, title)
                extension = "json"
            elif chosen == "PSV":
                data, count = build_metadata_psv(records, title)
                extension = "psv"
            else:
                data, count = build_metadata_xlsx(records, title)
                extension = "xlsx"
        except ExportError as exc:
            st.error(f"{title} export blocked: {exc}", icon="🚫")
            return
        except Exception as exc:  # noqa: BLE001
            LOG.exception("export failed")
            st.error(f"{title} export failed: {clean_text(exc)[:200]}", icon="🚫")
            return
        st.caption(f"{count} row(s) · {fmt_size(len(data))} · integrity checked")
        st.download_button(
            f"Download {title} ({chosen})", data,
            file_name=f"parchai-{slug}-{prefix}-{stamp}.{extension}",
            mime=mime_by_format[chosen], key=f"{prefix}_download", **stretch(),
        )

    with st.container(border=True):
        export_panel("Main Metadata", "The complete, unfiltered four-field record set.",
                     metadata_records(active["id"]), [], "main_metadata")
    templates = list_templates(active["id"])
    with st.container(border=True):
        if templates:
            names = [row["name"] for row in templates]
            chosen_name = st.selectbox("Client template", names, key="export_template_choice")
            template = next(row for row in templates if row["name"] == chosen_name)
            records, missing = client_template_records(active["id"], template)
            export_panel(f"Client Template · {template['name']}",
                         "Only labels with a saved mapping, under the template's exact labels.",
                         records, missing, "client_template")
        else:
            st.markdown("**Client Template**")
            st.info("Create a template in Organize first.")


def page_settings(dockets: List[sqlite3.Row], active: Optional[sqlite3.Row]) -> None:
    page_header("Settings", "Gemini key pool, model fallback, storage and integrity checks.", None)
    if active:
        render_processing_loader(active["id"])
    entries = key_pool()
    config = load_config()

    st.markdown("#### Gemini API key pool")
    st.caption(
        "Keys come from the application configuration — the “gemini_api_keys” list in "
        f"`{CONFIG_PATH.name}` — or the `GEMINI_API_KEYS`/`GEMINI_API_KEY` environment variables. "
        "Keys are always masked here, rotated round-robin while processing, and never appear "
        "in logs. A quota or rate-limit error pauses the affected key and model for a minute; "
        "an invalid or rejected key is parked for an hour. Editing the config file takes "
        "effect on the next request; no restart is needed."
    )
    if config.get("_error"):
        st.error(f"Configuration problem: {config['_error']}", icon="🚫")
    for bad in rejected_keys():
        st.error(f"A key from “{bad['source']}” ending {bad['masked']} was ignored because it is not a valid "
                 "API key (it must be a complete ASCII key beginning with “AQ.” or “AIza”, with no spaces). "
                 f"Check the value in `{CONFIG_PATH.name}` or the GEMINI_API_KEY(S) variable.", icon="🚫")
    if not entries:
        st.warning(
            "No valid Gemini keys configured. Add 4–5 keys to “gemini_api_keys” in "
            f"`{CONFIG_PATH.name}` before processing documents.",
            icon="⚠️",
        )
    else:
        rows = []
        for entry in entries:
            cooling = max((cooldown_remaining(f"{entry['id']}:{m}") for m in GEMINI_MODELS), default=0)
            cooling = max(cooling, cooldown_remaining(entry["id"]))
            rows.append({
                "Key": entry["label"],
                "Value": entry["masked"],
                "Source": entry["source"],
                "Status": "✔ Ready" if cooling == 0 else f"▲ Cooling down ({int(cooling)}s)",
            })
        st.dataframe(rows, hide_index=True, **stretch())

    if entries:
        options = [f"{e['label']} ({e['masked']})" for e in entries]
        with st.container(border=True):
            st.markdown("**Diagnostics**")
            choice = st.selectbox("Key", options, key="test_choice")
            chosen = entries[options.index(choice)]
            test_col, list_col, clear_col = st.columns(3)
            if test_col.button("Test selected key", **stretch()):
                with st.spinner("Contacting Gemini…"):
                    result = check_key(chosen["key"])
                (st.success if result.startswith("Key accepted") else st.error)(result)
            if list_col.button("List models for selected key", **stretch()):
                try:
                    with st.spinner("Fetching models…"):
                        names = list_models(chosen["key"])
                    st.code("\n".join(names) or "No models returned.", language=None)
                except Exception as exc:  # noqa: BLE001
                    st.error(f"Failed: {scrub(exc, chosen['key'])}")
            clear_col.button("Clear all cooldowns", on_click=clear_cooldowns_callback, **stretch())

    st.markdown("#### Model fallback order")
    st.code("\n".join(f"{i}. {m}" for i, m in enumerate(GEMINI_MODELS, start=1)), language=None)
    st.caption("Override with “gemini_models” in the config file or the PARCHAI_GEMINI_MODELS "
               "environment variable (comma-separated; restart required). Use “List models” above "
               "to see valid names.")

    st.markdown("#### Storage and integrity")
    st.caption(f"Database: `{DB_PATH}`")
    st.caption(f"Original files: `{STORAGE_DIR}`")
    if st.button("Run integrity check"):
        with st.spinner("Checking the database and every stored file…"):
            problems = storage_audit()
        if problems:
            for problem in problems:
                st.error(problem, icon="🚫")
        else:
            st.success("Database and all stored originals passed their checks.", icon="✅")


RENDERERS = {
    "Dockets": page_dockets,
    "Workspace": page_workspace,
    "Review": page_review,
    "Organize": page_organize,
    "Viewer": page_viewer,
    "Validation": page_validation,
    "Export": page_export,
    "Settings": page_settings,
}


def main() -> None:
    st.set_page_config(page_title="Parchi AI", page_icon="📄", layout="wide",
                       initial_sidebar_state="collapsed")
    inject_css()
    page = st.session_state.get("nav", PAGES[0])
    if page not in RENDERERS:
        page = PAGES[0]
    try:
        init_db()
        dockets = list_dockets()
        active = render_topbar(dockets)
        render_flash()
        RENDERERS[page](dockets, active)
    except Exception as exc:  # noqa: BLE001  (error boundary)
        if is_control_flow(exc):
            raise
        LOG.exception("page %s failed", page)
        render_error_panel(exc, page)


if __name__ == "__main__":
    main()