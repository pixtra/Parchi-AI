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
import math
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
import streamlit.components.v1 as components
import pandas as pd
from psv_workflow import band_for_score, label_value_mismatch
from metadata_workflow import (
    json_bytes,
    sheet_text,
)
from value_groups import (
    VALUE_GROUP_SCHEMA,
    GroupError,
    add_member,
    create_group,
    delete_group,
    load_groups,
    rename_group,
    remove_member,
    reorder_members,
    ungrouped_fact_ids,
)
from extraction_adapter import coerce_nullable_value, confidence_map
from validated_data import (
    VALIDATED_DATA_SCHEMA,
    apply_validated_records,
    edit_validated_record,
    list_validated_edit_history,
    set_review_approval,
    undo_validated_edit,
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

LOG = logging.getLogger("parchi")
# One worker: dockets are extracted sequentially, never concurrently (checklist §6).
PROCESSING_EXECUTOR = ThreadPoolExecutor(max_workers=1, thread_name_prefix="parchi-extraction")
PROCESSING_LOCK = threading.RLock()
PROCESSING_JOBS: Dict[str, Dict[str, Any]] = {}
WORKER_STATE = threading.local()


# ---------------------------------------------------------------- configuration
BASE_DIR = Path(__file__).resolve().parent

# Drag-and-drop board for manual value grouping (Organize tab). The static asset folder
# ships with the app; declare it once so it is served by Streamlit's component route.
_VALUE_GROUP_COMPONENT_DIR = BASE_DIR / "value_group_component"
_value_groups_component = (
    components.declare_component("parchi_value_groups", path=str(_VALUE_GROUP_COMPONENT_DIR))
    if _VALUE_GROUP_COMPONENT_DIR.is_dir() else None
)
def _env(name: str, default: Any = "") -> Any:
    """Read PARCHI_<name> configuration values."""
    return os.getenv(f"PARCHI_{name}", default)


DB_PATH = Path(_env("DB", BASE_DIR / "parchi_v3.db"))
STORAGE_DIR = Path(_env("STORAGE", BASE_DIR / "parchi_storage")).resolve()
CONFIG_PATH = BASE_DIR / ".streamlit" / "secrets.toml"


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
        _env("GEMINI_MODELS", "").split(",")
        if _env("GEMINI_MODELS", "").strip()
        else (
            _MODEL_CONFIG
            if isinstance(_MODEL_CONFIG, list) and any(str(m).strip() for m in _MODEL_CONFIG)
            else "gemini-3.8-flash,gemini-3.6-flash,gemini-3.1-flash-lite".split(",")
        )
    )
    if str(m).strip()
)
GEMINI_TIMEOUT_MS = 180_000
MAX_OUTPUT_TOKENS = int(_env("MAX_OUTPUT_TOKENS", "32768"))
CHUNK_PAGES = int(_env("CHUNK_PAGES", "10"))
TOKENS_PER_PAGE = int(_env("TOKENS_PER_PAGE", "1500"))
MAX_BATCH_TOKENS = int(_env("MAX_BATCH_TOKENS", "400000"))
MAX_PARTS_PER_BATCH = int(_env("MAX_PARTS_PER_BATCH", "12"))
MAX_UPLOAD_BYTES = int(_env("MAX_UPLOAD_MB", "50")) * 1024 * 1024
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
PAGES = ["Dockets", "Workspace", "Organize", "Viewer", "Validation", "Export", "Settings"]
PAGE_ICONS = {
    "Dockets": "▤", "Workspace": "⬆", "Organize": "⌘", "Viewer": "▣",
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
CREATE TABLE IF NOT EXISTS template_fields (
    id TEXT PRIMARY KEY,
    template_id TEXT NOT NULL REFERENCES templates(id) ON DELETE CASCADE,
    source_type TEXT NOT NULL CHECK(source_type IN ('group', 'label')),
    source_id TEXT NOT NULL,
    source_name TEXT NOT NULL,
    destination_label TEXT NOT NULL,
    position INTEGER NOT NULL,
    UNIQUE(template_id, position)
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
- Write labels in clear, concise language that names what the field means. Prefer precise
  names ("Supplier GSTIN", "Net Weight", "Challan Date") over vague or overly generic text
  ("Number", "Detail 1", "Other"). Read the surrounding headings, table column headers and
  nearby text before choosing a label, and keep the label's original meaning — do not
  pad, abbreviate or reword it beyond what is needed to be clear.
- If the source text of a label is garbled or unreadable (for example OCR noise such as
  "asfasfafw"), use the document's context — nearby headings, table headers, the values on
  the same row or column — to infer the intended field name, and use that corrected label.
  Quote the original garbled text verbatim in evidence, keep the extracted value exactly
  as read, and give the fact a LOW confidence score (20-40) reflecting the uncertainty of
  the correction. If the intended field cannot be inferred even with context, use the label
  "Unclear", give a very low score (5-15), keep the extracted value, and describe what is
  unreadable in evidence. Never invent a meaning or a value, and never drop the occurrence.
- Copy values exactly as written. Do not normalize, calculate, correct, translate, or infer.
- Never fill a missing value from another document. Evidence must be verbatim from the same
  page or worksheet.
- Assign every fact an individual integer confidence_score from 1 through 100, judging THAT
  occurrence on its own merits: the legibility of the text it came from, the quality of the
  scan or writing, ambiguity, and how well the evidence supports the value. Never derive a
  score from other occurrences: a value repeated across documents or pages must NOT get a
  higher score for repeating, and a conflict with another occurrence must NOT lower one.
  Never return 0 or 100, and never reuse one default score for many facts — two facts from
  the same page should get different scores when their legibility differs.
  Calibrate with these anchors and interpolate between them:
    * 95-99: crisp digitally-typed text; label and value cleanly legible; evidence quoted
      verbatim from the page.
    * 75-90: typed text with slight noise, a cramped layout, or a value assembled across
      line breaks.
    * 55-80: clear handwriting or a modest-quality scan, at most one ambiguous character.
    * 25-55: faint, skewed or partially readable text, or a label you had to infer from
      context.
    * 5-20: text you could barely read at all, or a garbled label you could not confidently
      correct.
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


def _migrate_global_value_groups(connection: sqlite3.Connection) -> None:
    columns = {row["name"] for row in connection.execute("PRAGMA table_info(value_groups)")}
    if not columns or "docket_id" not in columns:
        return
    connection.execute("DROP INDEX IF EXISTS idx_value_groups_name")
    connection.execute("ALTER TABLE value_group_members RENAME TO value_group_members_docketed")
    connection.execute("ALTER TABLE value_groups RENAME TO value_groups_docketed")
    connection.executescript(VALUE_GROUP_SCHEMA)
    rows = connection.execute(
        "SELECT id, docket_id, name, created_at, updated_at FROM value_groups_docketed "
        "ORDER BY created_at, id"
    ).fetchall()
    used = set()
    for row in rows:
        name = row["name"]
        if name.casefold() in used:
            docket = connection.execute("SELECT name FROM dockets WHERE id = ?", (row["docket_id"],)).fetchone()
            suffix = docket["name"] if docket else row["docket_id"][:8]
            name = f"{name} ({suffix})"
        used.add(name.casefold())
        connection.execute(
            "INSERT INTO value_groups (id, name, created_at, updated_at) VALUES (?, ?, ?, ?)",
            (row["id"], name, row["created_at"], row["updated_at"]),
        )
    connection.execute(
        """INSERT INTO value_group_members (group_id, fact_id)
           SELECT group_id, fact_id FROM value_group_members_docketed"""
    )
    connection.execute("DROP TABLE value_group_members_docketed")
    connection.execute("DROP TABLE value_groups_docketed")


def _remove_template_snapshots(connection: sqlite3.Connection) -> None:
    columns = {row["name"] for row in connection.execute("PRAGMA table_info(template_fields)")}
    if "snapshot_json" not in columns and "group_version" not in columns:
        return
    connection.execute("ALTER TABLE template_fields RENAME TO template_fields_with_snapshots")
    connection.executescript("""
        CREATE TABLE template_fields (
            id TEXT PRIMARY KEY,
            template_id TEXT NOT NULL REFERENCES templates(id) ON DELETE CASCADE,
            source_type TEXT NOT NULL CHECK(source_type IN ('group', 'label')),
            source_id TEXT NOT NULL,
            source_name TEXT NOT NULL,
            destination_label TEXT NOT NULL,
            position INTEGER NOT NULL,
            UNIQUE(template_id, position)
        );
    """)
    connection.execute(
        """INSERT INTO template_fields
           (id, template_id, source_type, source_id, source_name, destination_label, position)
           SELECT id, template_id, source_type, source_id, source_name, destination_label, position
           FROM template_fields_with_snapshots"""
    )
    connection.execute("DROP TABLE template_fields_with_snapshots")


@st.cache_resource
def init_db() -> bool:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    STORAGE_DIR.mkdir(parents=True, exist_ok=True)
    with connect() as connection:
        connection.execute("PRAGMA journal_mode = WAL")
        connection.executescript(SCHEMA)
        connection.executescript(VALUE_GROUP_SCHEMA)
        connection.executescript(VALIDATED_DATA_SCHEMA)
        review_columns = {row["name"] for row in connection.execute(
            "PRAGMA table_info(validated_template_records)"
        )}
        if review_columns and "review_approved" not in review_columns:
            connection.execute(
                "ALTER TABLE validated_template_records "
                "ADD COLUMN review_approved INTEGER NOT NULL DEFAULT 0"
            )
        _migrate_global_value_groups(connection)
        value_group_columns = {row["name"] for row in connection.execute("PRAGMA table_info(value_groups)")}
        if "primary_fact_id" not in value_group_columns:
            connection.execute("ALTER TABLE value_groups ADD COLUMN primary_fact_id TEXT")
        member_columns = {row["name"] for row in connection.execute("PRAGMA table_info(value_group_members)")}
        if "priority" not in member_columns:
            connection.execute(
                "ALTER TABLE value_group_members ADD COLUMN priority INTEGER NOT NULL DEFAULT 0"
            )
        _remove_template_snapshots(connection)
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
    """Every extracted occurrence in the docket, whatever its review status.

    Main Metadata is the canonical dataset: pending, approved and rejected records are all
    included (spec §6.2) and exports must never silently drop records by review status.
    Only internal bookkeeping rows (page processing status) are excluded."""
    with connect() as connection:
        return connection.execute(
            """SELECT d.filename AS document_name, f.label AS label, f.value AS value,
                      f.confidence_score AS confidence_score, f.review_status AS review_status
               FROM facts f JOIN documents d ON d.id = f.document_id
               WHERE f.docket_id = ? AND f.concept != 'page_processing_status'
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


def save_template_fields(template_id: str, fields: List[Dict[str, Any]]) -> None:
    """Replace a template's ordered live source definitions atomically."""
    with connect() as connection:
        if not connection.execute("SELECT 1 FROM templates WHERE id = ?", (template_id,)).fetchone():
            raise ValueError("That template no longer exists.")
        validated = []
        seen_sources = set()
        seen_destinations = set()
        for field in fields:
            destination = clean_text(field.get("destination_label")).strip()
            source_name = clean_text(field.get("source_name")).strip()
            source_id = clean_text(field.get("source_id")).strip()
            source_type = field.get("source_type")
            source_key = (source_type, normalize_label(source_id))
            destination_key = normalize_label(destination)
            if not destination or not source_name or not source_id or source_type not in {"group", "label"}:
                raise ValueError("Every template source needs a destination label.")
            if source_key in seen_sources:
                raise ValueError("A source can appear only once in a template.")
            if destination_key in seen_destinations:
                raise ValueError("Each template source needs a distinct destination label.")
            seen_sources.add(source_key)
            seen_destinations.add(destination_key)
            validated.append((source_type, source_id, source_name, destination))
        connection.execute("DELETE FROM template_fields WHERE template_id = ?", (template_id,))
        for position, (source_type, source_id, source_name, destination) in enumerate(validated):
            connection.execute(
                """INSERT INTO template_fields
                   (id, template_id, source_type, source_id, source_name, destination_label, position)
                   VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (str(uuid.uuid4()), template_id, source_type, source_id, source_name,
                 destination, position),
            )


def template_field_rows(template_id: str) -> List[sqlite3.Row]:
    with connect() as connection:
        return connection.execute(
            "SELECT * FROM template_fields WHERE template_id = ? ORDER BY position, id",
            (template_id,),
        ).fetchall()


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
        if added and connection.execute(
            "SELECT 1 FROM template_fields WHERE template_id = ? LIMIT 1", (template_id,)
        ).fetchone():
            position = int(connection.execute(
                "SELECT COALESCE(MAX(position), -1) + 1 FROM template_fields WHERE template_id = ?",
                (template_id,),
            ).fetchone()[0])
            connection.executemany(
                """INSERT INTO template_fields
                   (id, template_id, source_type, source_id, source_name, destination_label, position)
                   VALUES (?, ?, 'label', ?, ?, ?, ?)""",
                [(str(uuid.uuid4()), template_id, label, label, label, position + index)
                 for index, label in enumerate(added)],
            )
    return added


def save_mapping(docket_id: str, template_id: str, source_label: str, requested_label: str, method: str) -> None:
    """Approve a mapping from an extracted label to one of THIS template's labels.
    A source maps to exactly one target per template; re-approving replaces it."""
    source_label, requested_label = clean_text(source_label).strip(), clean_text(requested_label).strip()
    if not source_label or not requested_label:
        return
    with connect() as connection:
        existing_rows = connection.execute(
            "SELECT id, source_label FROM label_mappings WHERE template_id = ? "
            "ORDER BY approved DESC, created_at DESC", (template_id,),
        ).fetchall()
        existing = next((row for row in existing_rows
                         if normalize_label(row["source_label"]) == normalize_label(source_label)), None)
        if existing:
            connection.execute(
                """UPDATE label_mappings SET requested_label = ?, method = ?, approved = 1
                   WHERE id = ?""", (requested_label, method, existing["id"]),
            )
        else:
            connection.execute(
                """INSERT INTO label_mappings
                   (id, docket_id, template_id, source_label, requested_label, method, approved, created_at)
                   VALUES (?, ?, ?, ?, ?, ?, 1, ?)""",
                (str(uuid.uuid4()), docket_id, template_id, source_label, requested_label, method, utc_now()),
            )


def approved_mappings(docket_id: str, template_id: str) -> Dict[str, str]:
    """Maps each source label to the one target approved WITHIN this template."""
    with connect() as connection:
        rows = connection.execute(
            """SELECT source_label, requested_label FROM label_mappings
               WHERE template_id = ? AND approved = 1 ORDER BY created_at""",
            (template_id,),
        ).fetchall()
    return {row["source_label"]: row["requested_label"] for row in rows}


def mapping_details(docket_id: str, template_id: str) -> Dict[str, Dict[str, str]]:
    """Approved mappings with their method: {source: {"target": ..., "method": ...}}."""
    with connect() as connection:
        rows = connection.execute(
            """SELECT source_label, requested_label, method FROM label_mappings
               WHERE template_id = ? AND approved = 1 ORDER BY created_at""",
            (template_id,),
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
        rows = connection.execute(
            "SELECT id, source_label, requested_label FROM label_mappings WHERE template_id = ?",
            (template_id,),
        ).fetchall()
        source_key, target_key = normalize_label(source_label), normalize_label(requested_label)
        connection.executemany(
            "DELETE FROM label_mappings WHERE id = ?",
            [(row["id"],) for row in rows
             if normalize_label(row["source_label"]) == source_key
             and normalize_label(row["requested_label"]) == target_key],
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
               WHERE template_id = ?""",
            (template_id,),
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
    except (TypeError, ValueError, OverflowError):
        return default


def _parse_model_score(raw: Any) -> Optional[int]:
    """Parse a model-supplied confidence score, or None when it is unusable.

    A missing or malformed score is never replaced with a fixed default: the caller
    falls back to the deterministic assess() pipeline and records exactly why
    (MODEL_SCORE_MISSING / MODEL_SCORE_INVALID) so the gap is visible in the UI.
    Accepts ints/floats and numeric strings (a trailing % is tolerated); booleans,
    non-numeric strings, infinities and out-of-range values are rejected."""
    if raw is None or isinstance(raw, bool):
        return None
    try:
        text = str(raw).strip().rstrip("%")
        number = float(text)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(number) or number < 0 or number > 100:
        return None
    return round(number)


DOCUMENT_TITLE_WORDS = {
    "tax invoice", "invoice", "purchase order", "delivery challan",
    "goods receipt note", "grn", "weighment slip", "credit note",
    "debit note", "quotation", "receipt", "transport receipt",
}


def _normalised_phrase(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").lower()).strip()


def _looks_gibberish(token: Any) -> bool:
    """True when a token looks like keyboard noise rather than language.

    Language-agnostic heuristics: no vowels at all, or too few distinct letters for
    the token's length (e.g. "asfasfafw" has 4 distinct letters in 8). Deliberately
    coarse — it only ever disqualifies a label, and only when the page text cannot
    ground a reading (see _label_unreadable), so real but terse labels like "Wt"
    survive whenever they actually appear on the page."""
    text = re.sub(r"[^a-zA-Z]", "", str(token or ""))
    if len(text) < 3:
        return False
    if not re.search(r"[aeiouAEIOU]", text):
        return True
    return len(set(text.lower())) / len(text) <= 0.5


def _label_unreadable(label: Any, page_text: Optional[str]) -> bool:
    """A label is unreadable when it looks like gibberish AND the page offers no text
    that could ground a corrected reading. Gibberish that the page actually contains
    (codes, serials) is kept: this flag never discards information, it only marks
    extraction unreliability for the score cap in finalize_document."""
    stripped = str(label or "").strip()
    if not _looks_gibberish(stripped):
        return False
    if page_text is None:
        return True
    return not (_tokens(stripped) & _tokens(page_text))


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


def _tokens(text: Any) -> set:
    return set(re.findall(r"[a-z0-9]+", normalise(text)))


def evidence_overlap(evidence: str, page_text: str) -> float:
    """Share of the evidence's words that also appear on the page (0.0 to 1.0).
    Continuous, so facts with different evidence get different scores."""
    wanted = _tokens(evidence)
    if not wanted:
        return 0.0
    return len(wanted & _tokens(page_text)) / len(wanted)


def _soft_cap(score: float) -> float:
    """Compress scores above 88 instead of clipping them at 100, so high-confidence facts
    keep distinct values rather than piling up on the same ceiling."""
    if score <= 88:
        return score
    return 88 + 12 * (1 - math.exp(-(score - 88) / 10))


def calibrated_confidence(
    requested: Any,
    value: str,
    evidence: str,
    page_text: Optional[str],
    label: str,
) -> int:
    """Blend the model's own score with how well the evidence is found on the page."""
    base = clamp_confidence(requested, 50)
    value_len = len(value.strip())
    evidence_len = len(evidence.strip())
    if page_text is None:
        # No text layer to check against: scale with how much evidence was quoted.
        raw = base * 0.78 + 18 * (1 - math.exp(-evidence_len / 40))
        return max(1, min(100, round(_soft_cap(raw))))
    source = normalise(page_text)
    exact = bool(evidence.strip()) and normalise(evidence) in source
    value_found = bool(value.strip()) and normalise(value) in source
    overlap = evidence_overlap(evidence, page_text)
    adjustment = 0.0
    if exact:
        adjustment += 5 + 5 * (1 - math.exp(-evidence_len / 60))   # +5 .. +10
    elif value_found:
        adjustment -= 6
    else:
        adjustment -= 16
    adjustment += 8 * (overlap - 0.5)                               # -4 .. +4
    if value_len <= 1:
        adjustment -= 10
    if re.search(r"[?*_]{2,}|…", value):
        adjustment -= 12
    if _normalised_phrase(label) in {"name", "number", "date", "amount", "weight", "quantity"}:
        adjustment -= 3
    return max(1, min(100, round(_soft_cap(base + adjustment))))


def assess(value: str, evidence: str, page_text: Optional[str]) -> Dict[str, Any]:
    """Confidence for a fact without a model score (local rules). Continuous, like
    calibrated_confidence: nothing is pinned to a fixed value."""
    if not value.strip():
        return {"band": "Red", "score": 1, "reasons": ["MISSING"]}
    if not evidence.strip():
        score = round(8 + 37 * (1 - math.exp(-len(value.strip()) / 10)))
        return {"band": band_for(score), "score": score, "reasons": ["NO_EVIDENCE"]}
    if page_text is None:
        score = round(46 + 29 * (1 - math.exp(-len(evidence.strip()) / 120)))
        return {"band": band_for(score), "score": score, "reasons": ["NOT_TEXT_VERIFIABLE"]}
    source = normalise(page_text)
    if normalise(evidence) in source:
        reasons = ["GROUNDED"]
    elif normalise(value) and normalise(value) in source:
        reasons = ["VALUE_FOUND_EVIDENCE_MISMATCH"]
    else:
        reasons = ["UNSUPPORTED"]
    score = calibrated_confidence(72, value, evidence, page_text, "")
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
    """Keys from the application configuration file (parchi.config.json).

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
        recover_at = wait if wait > 0 else min((delay for delay in model_waits if delay > 0), default=0.0)
        if recover_at > 0:
            waits.append(recover_at)
    return min(waits) if waits else 0.0


def finish_reason(response: Any) -> str:
    candidates = getattr(response, "candidates", None) or []
    reason = getattr(candidates[0], "finish_reason", None) if candidates else None
    return str(getattr(reason, "name", reason) or "").upper()


def check_key(key: str) -> str:
    client = None
    try:
        client = make_client(key)
        models = list(islice(client.models.list(), 1))
        return "Key accepted." if models else "Key accepted (no models listed)."
    except Exception as exc:  # noqa: BLE001
        return f"Failed: {scrub(exc, key)}"
    finally:
        if client is not None:
            client.close()


def list_models(key: str) -> List[str]:
    client = make_client(key)
    names = []
    try:
        for model in client.models.list():
            actions = getattr(model, "supported_actions", None) or []
            if not actions or "generateContent" in actions:
                names.append(model.name.removeprefix("models/"))
    finally:
        client.close()
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
            value, _ = coerce_nullable_value(item.get("value"), str)
            value = value.strip()
            ambiguity_notes = str(item.get("ambiguity_notes") or "").strip()
            legibility = str(item.get("legibility") or "CLEAR").strip().upper()
            box_2d = item.get("box_2d")
            evidence = str(item.get("evidence") or "").strip()[:500]
            if ambiguity_notes:
                evidence = f"{evidence} Ambiguity notes: {ambiguity_notes}".strip()[:500]
            if is_document_title_fact(name, value, document_type):
                continue
            name = contextual_label(name, document_type)
            page_number, page_valid = page_of(item)
            fact = _fact(
                page_number, name, to_concept(name),
                value,
                evidence,
                item.get("confidence_score"),
            )
            fact["_source_label"] = str(item.get("label") or "Unlabelled field").strip()
            fact["_legibility"] = legibility if legibility in {"CLEAR", "BLURRY", "HANDWRITTEN", "MISSING"} else "CLEAR"
            fact["_ambiguity_notes"] = ambiguity_notes
            fact["_box_2d"] = box_2d
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


def _token_confidences(response: Any) -> Dict[Tuple[str, str], List[float]]:
    """Map confidence to fact label/value occurrences, excluding cross-reference values."""
    try:
        candidate = response.candidates[0]
        chosen = candidate.logprobs_result.chosen_candidates
        tokens = [(item.token, float(item.log_probability)) for item in chosen]
        serialized = "".join(token for token, _ in tokens)
        spans = []
        cursor = 0
        for token, logprob in tokens:
            spans.append((cursor, cursor + len(token), logprob))
            cursor += len(token)
        objects = []
        object_starts = []
        depth = 0
        quoted = escaped = False
        for pos, char in enumerate(serialized):
            if quoted:
                if escaped:
                    escaped = False
                elif char == "\\":
                    escaped = True
                elif char == '"':
                    quoted = False
                continue
            if char == '"':
                quoted = True
            elif char == "{":
                object_starts.append(pos)
                depth += 1
            elif char == "}":
                depth -= 1
                if object_starts:
                    object_start = object_starts.pop()
                    objects.append((object_start, serialized[object_start:pos + 1]))
        result: Dict[Tuple[str, str], List[float]] = {}
        for obj_start, fragment in objects:
            if "{" in fragment[1:-1]:
                continue
            label_match = re.search(r'"label"\s*:\s*("(?:\\.|[^"\\])*")', fragment)
            value_match = re.search(r'"value"\s*:\s*("(?:\\.|[^"\\])*"|null)', fragment)
            if not label_match or not value_match:
                continue
            try:
                label = json.loads(label_match.group(1))
                raw_value = value_match.group(1)
                value = json.loads(raw_value) if raw_value != "null" else ""
            except (TypeError, ValueError):
                continue
            if not isinstance(label, str) or not isinstance(value, str):
                continue
            label, value = label.strip(), value.strip()
            if raw_value == "null":
                score = 0.0
            else:
                value_start = obj_start + value_match.start(1) + 1
                value_end = obj_start + value_match.end(1) - 1
                selected = [prob for left, right, prob in spans
                            if right > value_start and left < value_end]
                score = math.exp(sum(selected) / len(selected)) if selected else 0.0
            result.setdefault((label, value), []).append(score)
        return result
    except (AttributeError, IndexError, TypeError, ValueError, OverflowError):
        return {}


def _audit_confidence(
    model_score: Any, token_probability: Optional[float], value: str,
    evidence: str, page_text: Optional[str], legibility: str, ambiguity: str,
) -> int:
    """Combine existing score calibration with token, OCR-agreement, and legibility signals."""
    base = calibrated_confidence(model_score, value, evidence, page_text, "")
    if token_probability is not None:
        base = round(base * 0.65 + max(0.0, min(1.0, token_probability)) * 100 * 0.35)
    if page_text and value.strip():
        if normalise(value) in normalise(page_text):
            base += 4
        else:
            base -= 8
    if legibility == "HANDWRITTEN":
        base = round(base * 0.88)
    elif legibility == "BLURRY":
        base = min(round(base * 0.75), 65)
    elif legibility == "MISSING" or not value.strip():
        return 0
    if ambiguity.strip():
        base -= 12
    return max(0, min(100, base))


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
                        ocr_context = "\n".join(
                            str(page.get("text") or "") for page in unit.get("pages_slice", [])
                        )
                        contents.append(
                            "Existing OCR text for comparison (may contain errors):\n"
                            + (ocr_context[:12000] if ocr_context else "[No text layer available]")
                        )
                    contents.append(BATCH_PROMPT)
                    contents.append(
                        "For every fact, also return legibility as CLEAR, BLURRY, HANDWRITTEN, "
                        "or MISSING; box_2d as [ymin,xmin,ymax,xmax] normalized 0-1000 "
                        "(use [] when unavailable); and ambiguity_notes as a short string. "
                        "Keep the existing documents/facts JSON structure and all existing keys."
                    )
                    set_gemini_activity(True, f"Gemini request in progress · {model}")
                    try:
                        config_args = dict(response_mime_type="application/json", temperature=0.1,
                                           max_output_tokens=MAX_OUTPUT_TOKENS)
                        try:
                            response = client.models.generate_content(
                                model=model, contents=contents,
                                config=types.GenerateContentConfig(**config_args,
                                                                  response_logprobs=True, logprobs=1),
                            )
                        except Exception as logprob_exc:
                            # Older models/endpoints may not support response log probabilities.
                            LOG.info("Gemini log probabilities unavailable; retrying without them: %s",
                                     str(logprob_exc)[:200])
                            response = client.models.generate_content(
                                model=model, contents=contents,
                                config=types.GenerateContentConfig(**config_args),
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
                        token_scores = _token_confidences(response)
                        for unit in batch:
                            bucket = parsed[0].get(unit["label"])
                            if not bucket:
                                continue
                            for fact in bucket["facts"]:
                                key = (fact.get("_source_label", ""), fact["value"])
                                candidates = token_scores.get(key, [])
                                if candidates:
                                    fact["_token_probability"] = candidates.pop(0)
                    except ValueError as exc:  # json.JSONDecodeError is a ValueError
                        LOG.warning("Gemini extraction schema/payload parse failed for %s · %s: %s",
                                    masked, model, str(exc)[:300])
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
        return None, {
            "filename": filename,
            "status": "skipped",
            "fact_count": 0,
            "extracted_count": 0,
            "appended_count": 0,
            "message": f"Identical to “{duplicate['filename']}” already exists in this docket "
                       f"with status {document_state_label(duplicate['status'])}.",
        }
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
    """Save all facts and return the number persisted for this document.

    Facts from an earlier extraction of the same document are deleted first, inside the
    same transaction, so a failed save rolls back and leaves the previous extraction —
    and its review history — untouched."""
    now = utc_now()
    extracted_items = [
        item for item in items if item.get("concept") != "page_processing_status"
    ]
    if status == "failed" and not extracted_items:
        # A failed retry must not erase a previously successful extraction. Keep the old
        # facts available for review/export while recording that the latest attempt failed.
        with connect() as connection:
            connection.execute(
                """UPDATE documents SET status = ?, document_type = ?, processed_units = 0,
                   failed_units = ?, completeness_status = ?, completeness_message = ?
                   WHERE id = ?""",
                (status, clean_text(document_type or "Unclassified")[:80], doc["page_count"],
                 "complete_with_warnings", "Gemini did not return usable values; previous "
                 "extraction was retained. Retry this document.", doc["id"]),
            )
            connection.execute("UPDATE dockets SET updated_at = ? WHERE id = ?", (now, docket_id))
            return int(connection.execute(
                "SELECT COUNT(*) FROM facts WHERE document_id = ? "
                "AND concept != 'page_processing_status'", (doc["id"],)
            ).fetchone()[0])

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
    financial_by_page: Dict[int, Dict[str, List[float]]] = {}
    for item in extracted_items:
        concept = str(item.get("concept") or "")
        key = next((candidate for candidate in ("subtotal", "tax_amount", "total_amount", "grand_total")
                    if candidate in concept), None)
        if key:
            try:
                financial_by_page.setdefault(int(item.get("page_number", 1)), {}).setdefault(key, []).append(
                    float(str(item.get("value") or "").replace(",", "").replace("₹", "").strip())
                )
            except ValueError:
                pass
    math_consistent = {}
    for page_number, amounts in financial_by_page.items():
        total_key = "total_amount" if "total_amount" in amounts else "grand_total"
        # Do not infer accounting equations when multiple values/sections exist.
        unique = all(len(amounts.get(key, [])) == 1 for key in ("subtotal", "tax_amount", total_key))
        if unique and all(key in amounts for key in ("subtotal", "tax_amount", total_key)):
            math_consistent[page_number] = (
                round(amounts["subtotal"][0] + amounts["tax_amount"][0], 2)
                == round(amounts[total_key][0], 2)
            )
    with connect() as connection:
        # Replace-on-success: drop any earlier extraction of this document (its history
        # cascades) before inserting the new facts, all in one transaction.
        connection.execute("DELETE FROM facts WHERE document_id = ?", (doc["id"],))
        for item in items:
            page_number = min(max(1, item["page_number"]), doc["page_count"])
            value = clean_text(item["value"])
            evidence = clean_text(item["evidence"])
            assessment = assess(value, evidence, doc["page_text"].get(page_number))
            # Model scores, when present, are combined with token and audit signals;
            # otherwise the deterministic assessment supplies the base score.
            model_score = _parse_model_score(item.get("confidence_score"))
            if model_score is None:
                assessment["reasons"].append(
                    "MODEL_SCORE_MISSING" if item.get("confidence_score") is None
                    else "MODEL_SCORE_INVALID"
                )
            assessment["score"] = _audit_confidence(
                model_score if model_score is not None else assessment["score"],
                item.get("_token_probability"), value, evidence,
                doc["page_text"].get(page_number), item.get("_legibility", "CLEAR"),
                item.get("_ambiguity_notes", ""),
            )
            concept = str(item.get("concept") or "")
            if any(key in concept for key in ("subtotal", "tax_amount", "total_amount", "grand_total")):
                if math_consistent.get(page_number) is True:
                    assessment["score"] = min(100, assessment["score"] + 6)
                    assessment["reasons"].append("FINANCIAL_MATH_CHECK_PASSED")
            assessment["band"] = band_for(assessment["score"])
            assessment["reasons"] = list(dict.fromkeys(
                assessment["reasons"] + ["AUDIT_CONFIDENCE"]
            ))
            # An unreadable label makes the whole extraction unreliable even when the
            # value looks right: cap the score low. The value itself is never dropped —
            # the reviewer decides what to do with it.
            if _label_unreadable(item.get("label"), doc["page_text"].get(page_number)):
                assessment["score"] = min(assessment["score"], 12)
                assessment["band"] = band_for(assessment["score"])
                assessment["reasons"] = list(dict.fromkeys(
                    assessment["reasons"] + ["LABEL_UNREADABLE"]
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
            "SELECT COUNT(*) FROM facts WHERE document_id = ? "
            "AND concept != 'page_processing_status'", (doc["id"],)
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


def _job_doc_update(job_id: str, document_id: str, **fields: Any) -> None:
    """Update one document's live entry in a job's queue (created on demand).

    The queue is what the processing panel renders per file: name, live status and
    fact count. Called from the worker thread only; every mutation takes the lock."""
    with PROCESSING_LOCK:
        job = PROCESSING_JOBS.get(job_id)
        if job is None:
            return
        docs = job.setdefault("docs", [])
        for entry in docs:
            if entry["id"] == document_id:
                entry.update(fields)
                return
        docs.append({"id": document_id, **fields})


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
            extracted_count = sum(
                item.get("concept") != "page_processing_status" for item in items
            )
            fact_count = finalize_document(docket_id, doc, status, doc_type, items)
            if status in {"failed", "partial"} and message:
                with connect() as connection:
                    connection.execute(
                        "UPDATE documents SET completeness_message = ? WHERE id = ?",
                        (message[:500], doc["id"]),
                    )
            # Decisions captured when the document was queued are restored onto the fresh
            # extraction where label and value still match; unmatched decisions lapse.
            carried = reapply_review_decisions(doc["id"], doc.get("previous_decisions") or [])
            if carried:
                LOG.info("carried %d review decision(s) onto re-extracted facts of %s",
                         carried, doc["filename"])
            finalized.add(doc["id"])
            return {
                "filename": doc["filename"],
                "status": status,
                "fact_count": fact_count,
                "extracted_count": extracted_count,
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
            # Include files rejected before extraction in the denominator so the
            # overall percentage describes every file selected for this run.
            PROCESSING_JOBS[job_id].update(
                total_docs=len(documents) + len(results), done_docs=len(results)
            )

        for document_index, doc in enumerate(documents, start=1):
            set_document_status(doc["id"], "processing")
            _job_doc_update(job_id, doc["id"], filename=doc["filename"], status="processing",
                            facts=0, message="Sending to Gemini…")
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
                outcome = save(
                    doc,
                    gemini_document_status({"units": units}),
                    doc_type or "Unclassified",
                    items,
                    " | ".join(messages),
                )
                results.append(outcome)
                _job_doc_update(job_id, doc["id"], status=outcome["status"],
                                facts=outcome["appended_count"], message=outcome["message"])
                mark_done()
            except Exception as exc:  # noqa: BLE001
                LOG.exception("document processing failed for %s", doc["filename"])
                set_document_status(doc["id"], "failed")
                message = (f"Processing failed ({clean_text(exc)[:200]}). "
                           "Completed documents were preserved; retry this document.")
                results.append(job_result(doc["filename"], "failed", message))
                _job_doc_update(job_id, doc["id"], status="failed", facts=0, message=message)
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
                message = (f"Processing was interrupted ({clean_text(exc)[:150]}). "
                           "Completed documents were preserved; retry this document.")
                results.append(job_result(doc["filename"], "failed", message))
                _job_doc_update(job_id, doc["id"], status="failed", facts=0, message=message)
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
.pp-spinner{width:24px;height:24px;flex:0 0 24px;border:3px solid rgba(128,128,128,.28);
  border-top-color:#60a5fa;border-right-color:#2563eb;border-radius:50%;
  animation:pp-spin .8s linear infinite}
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
.pp-loader.error .pp-bar,.pp-loader.error .pp-orbit{stroke:#ff8585}
.pp-loader.ok .pp-bar{stroke:#4fd6a8}
.pp-docs{margin:.4rem 0 0 .5rem;display:grid;gap:.3rem;max-height:11.5rem;overflow-y:auto}
.pp-doc{display:flex;gap:.55rem;align-items:baseline;font-size:.86rem;flex-wrap:wrap}
.pp-doc-name{font-weight:600;overflow-wrap:anywhere}
.pp-doc-facts{opacity:.75;font-size:.8rem}
.pp-doc-msg{opacity:.7;font-size:.8rem;overflow-wrap:anywhere}
@keyframes pp-spin{to{transform:rotate(360deg)}}
@media (prefers-reduced-motion: reduce){.pp-orbit,.pp-spinner{animation:none;opacity:.5}}
</style>
"""


PROCESSING_SUMMARY_SECONDS = 90


def queue_rows(job: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Normalize a job's per-document queue for display.

    Every row carries the file name, its live status, the fact count saved so far and a
    short human message. Rows keep queue order (queued documents first, then failures
    that never entered the queue)."""
    rows = []
    for entry in job.get("docs") or []:
        rows.append({
            "filename": entry.get("filename") or "",
            "status": entry.get("status") or "queued",
            "facts": int(entry.get("facts") or 0),
            "message": entry.get("message") or "",
        })
    seen = {row["filename"] for row in rows}
    for result in job.get("results") or []:
        if result.get("filename") in seen:
            continue
        status = result.get("status") or "failed"
        message = result.get("message") or ""
        if status == "skipped":
            status = "failed"
            message = f"Skipped: {message}" if message else "Skipped before extraction."
        rows.append({
            "filename": result.get("filename") or "",
            "status": status,
            "facts": int(result.get("appended_count") or 0),
            "message": message,
        })
    return rows


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
        # Streamlit's native running status icon remains animated after this fragment
        # finishes rendering. The former CSS-only ring could be swallowed or made
        # inconspicuous by the app's Markdown/iframe theme styling.
        status = st.status(f"{title}…", state="running", expanded=True)
        status.markdown(job.get("message") or "Preparing the saved documents.")
        status.progress(percent, text=f"Overall document progress · {done}/{total} finished")
        status.caption(
            f"{done} of {total} document(s) finished · you can switch tabs safely; "
            "facts appear as each document completes."
        )
        rows = queue_rows(job)
        if rows:
            queue = pd.DataFrame([{
                "Document": row["filename"],
                "Status": document_state_label(row["status"]),
                "Records": row["facts"] if row["status"] in {"completed", "partial"} else "",
                "Details": row["message"],
            } for row in rows])
            status.dataframe(
                queue, hide_index=True, width="stretch",
                height=min(240, max(110, 35 + 34 * len(queue))),
            )
    elif job["status"] == "error":
        _render_processing_panel(
            job, title="Processing stopped", message=esc(job.get("message") or ""),
            percent=None, kind="error",
        )
    elif job["status"] == "complete" \
            and time.time() - float(job.get("finished_at") or 0) <= PROCESSING_SUMMARY_SECONDS:
        rows = queue_rows(job)
        counts: Dict[str, int] = {}
        for row in rows:
            counts[row["status"]] = counts.get(row["status"], 0) + 1
        facts = sum(row["facts"] for row in rows)
        st.progress(100, text="Overall document progress · complete")
        summary = " · ".join(
            f"{document_state_label(status)}: {counts[status]}"
            for status in ("completed", "partial", "failed") if status in counts
        )
        _render_processing_panel(
            job, title="Processing complete", message=esc(summary or "No documents were processed."),
            percent=100, kind="ok",
            sub=f"{facts} extracted record(s) saved to Main Metadata." if facts else "",
        )


def _render_processing_panel(job: Dict[str, Any], title: str, message: str,
                             percent: Optional[int], kind: str = "",
                             sub: str = "") -> None:
    """Throbber + overall progress + the live per-document queue, in the app's chip style."""
    ring = ""
    if percent is not None:
        ring = (
            '<div class="pp-ring">'
            '<svg viewBox="0 0 64 64" aria-hidden="true">'
            '<circle class="pp-track" cx="32" cy="32" r="28" pathLength="100"/>'
            + (f'<circle class="pp-bar" cx="32" cy="32" r="28" pathLength="100" '
               f'style="stroke-dasharray:{max(0, min(100, percent))} 100"/>'
               if percent else '')
            + ('<circle class="pp-orbit" cx="32" cy="32" r="28" pathLength="100" '
               'style="stroke-dasharray:14 86"/>' if kind == "" else '')
            + '</svg>'
            + f'<div class="pp-pct">{max(0, min(100, percent))}%</div></div>'
        )
    queue_html = ""
    rows = queue_rows(job)
    if rows:
        cells = []
        for row in rows:
            facts = f'<span class="pp-doc-facts">{row["facts"]} record(s)</span>' \
                if row["status"] in {"completed", "partial"} else ""
            note = f'<span class="pp-doc-msg">{esc(row["message"])}</span>' \
                if row["message"] and row["status"] in {"failed", "partial"} else ""
            cells.append(
                '<div class="pp-doc">'
                + f'<span class="pp-doc-name">{esc(row["filename"])}</span>'
                + document_state_badge(row["status"])
                + facts + note
                + '</div>'
            )
        queue_html = ('<div class="pp-docs" aria-label="Document queue">' + "".join(cells) + '</div>')
    st.markdown(
        _LOADER_CSS
        + f'<div class="pp-loader {esc(kind)}" role="status" aria-live="polite">'
        + ('<div class="pp-spinner" aria-label="Processing"></div>' if kind == "" else '')
        + ring
        + f'<div><div class="pp-title">{esc(title)}</div>'
        + f'<div class="pp-msg">{message}</div>'
        + (f'<div class="pp-sub">{esc(sub)}</div>' if sub else '')
        + '</div></div>'
        + queue_html,
        unsafe_allow_html=True,
    )


def _watch_processing_progress(job: Optional[Dict[str, Any]]) -> None:
    """Rerun the whole app whenever a document finishes, so its facts appear in
    Workspace, Review and Main Metadata without any manual refresh or tab change.

    Shared by every auto-refreshing fragment that shows processing state; the
    session-state keys make multiple fragments (banner + step-2 panel) dedupe to a
    single rerun. The first sighting only seeds the baseline so opening a page
    mid-run doesn't trigger a pointless rerun."""
    if not job:
        return
    progress_mark = (job["id"], job.get("done_docs", 0))
    if progress_mark == st.session_state.get("last_seen_progress"):
        return
    st.session_state["last_seen_progress"] = progress_mark
    if st.session_state.get("last_seen_progress_seeded"):
        st.rerun(scope="app")
    st.session_state["last_seen_progress_seeded"] = True


@st.fragment(run_every=1)
def render_processing_loader(docket_id: str) -> None:
    """Live processing banner: throbber, overall progress and the per-document queue.
    Rendered on every page for the active docket, so progress stays visible across
    tab navigation; auto-refreshes every second while a run is active or just finished."""
    job = processing_job(docket_id)
    if job is None:
        return
    _watch_processing_progress(job)
    _render_processing_status(job)


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


def capture_review_decisions(document_id: str) -> List[Dict[str, Any]]:
    """Snapshot every human review decision on a document's facts before re-extraction.

    A decision is a fact the user corrected and/or approved or rejected. The snapshot is
    keyed by label and value so it can be re-applied to re-extracted occurrences; the
    document's facts themselves are NOT deleted here — they stay visible and are only
    replaced when the new extraction actually succeeds (see finalize_document)."""
    with connect() as connection:
        rows = connection.execute(
            """SELECT label, value, original_value, review_status, corrected
               FROM facts
               WHERE document_id = ?
                 AND concept != 'page_processing_status'
                 AND (review_status IN ('approved', 'rejected') OR corrected = 1)""",
            (document_id,),
        ).fetchall()
    return [
        {
            "label": row["label"],
            "value": row["value"],
            "original_value": row["original_value"],
            "review_status": row["review_status"],
            "corrected": int(row["corrected"]),
        }
        for row in rows
    ]


def reapply_review_decisions(document_id: str, decisions: List[Dict[str, Any]]) -> int:
    """Carry queued review decisions onto re-extracted facts where they still match.

    Matching is by normalized label plus value: a decision applies when the re-extracted
    value equals the value the user decided on, or (for a human-corrected fact) equals the
    original extraction the user corrected. Matching occurrences get the decided status
    (and the corrected value, if the user had edited it). Extracted confidence scores are
    never changed. Returns the number of facts updated."""
    applicable = [d for d in decisions if d.get("review_status") in ("approved", "rejected")]
    if not applicable:
        return 0
    with connect() as connection:
        facts = connection.execute(
            """SELECT id, label, value FROM facts
               WHERE document_id = ? AND concept != 'page_processing_status'""",
            (document_id,),
        ).fetchall()
        now = utc_now()
        applied = 0
        for fact in facts:
            fact_label = normalize_label(fact["label"])
            for decision in applicable:
                if normalize_label(decision["label"]) != fact_label:
                    continue
                final_value = fact["value"]
                if decision["value"] != final_value:
                    if not (decision["corrected"] and decision["original_value"] == final_value):
                        continue
                    # The user had edited this extraction; restore their correction.
                    final_value = decision["value"]
                connection.execute(
                    """UPDATE facts SET value = ?, corrected = ?, review_status = ?, updated_at = ?
                       WHERE id = ?""",
                    (final_value, int(decision["corrected"]), decision["review_status"], now,
                     fact["id"]),
                )
                connection.execute(
                    """INSERT INTO fact_history (id, fact_id, old_value, new_value, old_status,
                       new_status, changed_at, reason) VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                    (str(uuid.uuid4()), fact["id"], fact["value"], final_value, "pending",
                     decision["review_status"], now,
                     "review carried over from the previous extraction (retry)"),
                )
                applied += 1
    return applied


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
    docs: List[Dict[str, Any]] = []
    for row in list_documents(docket_id):
        if row["id"] not in wanted or row["status"] not in PROCESSABLE_STATES:
            continue
        document, problem = load_saved_document(docket_id, row["id"])
        if problem:
            set_document_status(row["id"], "failed")
            results.append(job_result(row["filename"], "failed", problem))
            docs.append({"id": row["id"], "filename": row["filename"], "status": "failed",
                         "facts": 0, "message": problem})
            continue
        set_document_status(row["id"], "queued")
        docs.append({"id": row["id"], "filename": row["filename"], "status": "queued",
                     "facts": 0, "message": ""})
        # Snapshot review decisions now; the document's existing facts stay in place until
        # its new extraction succeeds (finalize_document replaces them transactionally).
        document["previous_decisions"] = capture_review_decisions(row["id"])
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
            "docs": docs,
        }
    PROCESSING_EXECUTOR.submit(_run_processing_job, job_id, docket_id, payloads, keys, results)
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
    # Live updates across the whole app (facts appearing as each document completes) are
    # handled by _watch_processing_progress inside the processing banner fragment, which
    # renders on every page above this panel.
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
                queued_count, immediate_results = queue_saved_documents(docket_id, ids)
                if immediate_results:
                    st.session_state["last_results"] = immediate_results
                st.session_state["clear_process_pick"] = True
                if queued_count or immediate_results:
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


def viewer_metadata_records(docket_id: str) -> List[dict]:
    """Main Metadata rows for the Viewer table, plus each record's review status.

    The record set is identical to the exports — nothing is filtered by review status.
    The review status column is display-only traceability and is never exported."""
    return [
        {
            "document_name": row["document_name"],
            "label": row["label"],
            "value": row["value"],
            "confidence_score": int(row["confidence_score"]),
            "review_status": row["review_status"],
        }
        for row in fetch_metadata(docket_id)
    ]


def viewer_validated_base_records(docket_id: str) -> List[dict]:
    """Full occurrence-level editable copy, seeded from but stored apart from Main Metadata."""
    return [
        {
            "document_name": row["filename"],
            "label": row["label"],
            "value": row["value"],
            "confidence_score": int(row["confidence_score"]),
            "review_status": row["review_status"],
            "source_label": row["label"],
            "record_id": row["id"],
        }
        for row in fetch_facts(docket_id)
        if row["concept"] != "page_processing_status"
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
    field_rows = template_field_rows(template["id"])
    mappings = approved_mappings(docket_id, template["id"])
    # Mappings may be stored with one spelling of a source label while facts carry another,
    # so the lookup compares case-insensitively and ignores surrounding whitespace.
    mapping_by_normalized = {normalize_label(source): target for source, target in mappings.items()}
    output: List[dict] = []
    if field_rows:
        with connect() as connection:
            rows = connection.execute(
                """SELECT f.id, f.label, d.filename AS document_name, f.value, f.confidence_score
                   FROM facts f JOIN documents d ON d.id = f.document_id
                   WHERE f.docket_id = ? AND f.concept != 'page_processing_status'""",
                (docket_id,),
            ).fetchall()
            membership: Dict[str, List[str]] = {}
            member_rows = connection.execute(
                """SELECT m.group_id, f.id AS fact_id
                   FROM value_group_members m
                   JOIN facts f ON f.id = m.fact_id
                   JOIN documents d ON d.id = f.document_id
                   JOIN value_groups g ON g.id = m.group_id
                   WHERE f.docket_id = ? AND f.concept != 'page_processing_status'
                   ORDER BY m.group_id, m.priority,
                     CASE WHEN f.id = g.primary_fact_id THEN 0 ELSE 1 END,
                     f.confidence_score DESC, d.filename, f.page_number, f.created_at, f.id""",
                (docket_id,),
            ).fetchall()
            for member in member_rows:
                membership.setdefault(member["group_id"], []).append(member["fact_id"])
        records_by_id = {record["id"]: record for record in rows}
        matched = {row["id"]: 0 for row in field_rows}
        for field in field_rows:
            if field["source_type"] == "group":
                ordered_ids = membership.get(field["source_id"], [])
                allowed_ids = set(ordered_ids)
            else:
                ordered_ids = []
                allowed_ids = {row["id"] for row in rows
                                if normalize_label(row["label"]) == normalize_label(field["source_id"])}
            source_labels = {normalize_label(source) for source, target in mappings.items()
                             if target == field["destination_label"]}
            if field["source_type"] == "group":
                # A saved group contributes only its first prioritized occurrence to
                # template output; other members remain intact in Main/Review Metadata.
                primary_record = next((records_by_id[fact_id] for fact_id in ordered_ids
                                       if fact_id in records_by_id), None)
                if primary_record is not None:
                    output.append({
                        "document_name": primary_record["document_name"],
                        "label": field["destination_label"], "value": primary_record["value"],
                        "confidence_score": int(primary_record["confidence_score"]),
                        "source_label": primary_record["label"], "record_id": primary_record["id"],
                    })
                    matched[field["id"]] += 1
                # Keep ordinary approved mappings that are not already members of this group.
                for record in rows:
                    if record["id"] in allowed_ids or normalize_label(record["label"]) not in source_labels:
                        continue
                    output.append({
                        "document_name": record["document_name"],
                        "label": field["destination_label"], "value": record["value"],
                        "confidence_score": int(record["confidence_score"]),
                        "source_label": record["label"], "record_id": record["id"],
                    })
                    matched[field["id"]] += 1
                continue
            for record in rows:
                if record["id"] not in allowed_ids and normalize_label(record["label"]) not in source_labels:
                    continue
                output.append({
                    "document_name": record["document_name"],
                    "label": field["destination_label"], "value": record["value"],
                    "confidence_score": int(record["confidence_score"]),
                    "source_label": record["label"], "record_id": record["id"],
                })
                matched[field["id"]] += 1
        return output, [row["destination_label"] for row in field_rows if matched[row["id"]] == 0]
    matched = {label: 0 for label in labels}
    for fact in fetch_facts(docket_id):
        if fact["concept"] == "page_processing_status":
            continue
        target = mapping_by_normalized.get(normalize_label(fact["label"]))
        if target in matched:
            output.append({
                "document_name": fact["filename"], "label": target, "value": fact["value"],
                "confidence_score": int(fact["confidence_score"]),
                "source_label": fact["label"], "record_id": fact["id"],
            })
            matched[target] += 1
    return output, [label for label, count in matched.items() if count == 0]


def build_metadata_export(records: List[dict], description: str) -> Tuple[bytes, int]:
    exported = []
    for record in records:
        item = {field: record[field] for field in METADATA_FIELDS}
        item["confidence_score"] = int(item["confidence_score"])
        if "source_label" in record:
            item["source_label"] = record.get("source_label", "")
            item["record_id"] = record.get("record_id", "")
        if "review_approved" in record:
            item["review_approved"] = bool(record["review_approved"])
        exported.append(item)
    data = json_bytes(exported)
    if json.loads(data.decode("utf-8")) != exported:
        raise ExportError(f"{description} integrity check failed after re-reading.")
    return data, len(records)


def build_metadata_psv(records: List[dict], description: str) -> Tuple[bytes, int]:
    headers = list(METADATA_HEADERS)
    has_source = any("source_label" in record for record in records)
    has_approval = any("review_approved" in record for record in records)
    if has_source:
        headers += ["source_label", "record_id"]
    if has_approval:
        headers += ["review_approved"]
    expected = [headers] + [
        [
            safe_cell(record["document_name"]),
            safe_cell(record["label"]),
            safe_cell(record["value"]),
            str(int(record["confidence_score"])),
        ] + ([safe_cell(record.get("source_label", "")), safe_cell(record.get("record_id", ""))]
             if has_source else []) + ([str(bool(record.get("review_approved", False))).lower()]
                                        if has_approval else [])
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
    fields = list(METADATA_FIELDS)
    headers = list(METADATA_HEADERS)
    if any("source_label" in record for record in records):
        fields += ["source_label", "record_id"]
        headers += ["source_label", "record_id"]
    if any("review_approved" in record for record in records):
        fields += ["review_approved"]
        headers += ["review_approved"]
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Metadata"
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for row_number, record in enumerate(records, start=2):
        for column_number, field in enumerate(fields, start=1):
            value = record.get(field, "")
            cell = sheet.cell(row=row_number, column=column_number, value=value)
            if isinstance(value, str) and value[:1] in ("=", "+", "-", "@"):
                cell.data_type = "s"
        score = int(record["confidence_score"])
        color = {"Red": "FFC7CE", "Yellow": "FFEB9C", "Green": "C6EFCE"}[band_for(score)]
        sheet.cell(row=row_number, column=4).fill = PatternFill(
            fill_type="solid", start_color=color, end_color=color
        )
    sheet.freeze_panes = "A2"
    last_column = sheet.max_column
    sheet.auto_filter.ref = f"A1:{sheet.cell(row=1, column=last_column).column_letter}{max(sheet.max_row, 1)}"
    for letter, width in zip((sheet.cell(row=1, column=i).column_letter for i in range(1, last_column + 1)),
                             (34, 34, 60, 18, 34, 40)):
        sheet.column_dimensions[letter].width = width
    output = io.BytesIO()
    workbook.save(output)
    data = output.getvalue()
    reloaded = load_workbook(io.BytesIO(data), read_only=True)
    raw_rows = list(reloaded.active.iter_rows(min_row=2, max_col=len(fields), values_only=True))
    reloaded.close()
    # openpyxl has no concept of "blank but present" - an empty string cell round-trips as
    # None. That is expected, not corruption, so None is normalised back to "" for the
    # text columns before the comparison (confidence_score has no such ambiguity).
    got = []
    for values in raw_rows:
        row = dict(zip(fields, values))
        row["document_name"] = row["document_name"] or ""
        row["label"] = row["label"] or ""
        row["value"] = "" if row["value"] is None else row["value"]
        row["confidence_score"] = int(row["confidence_score"]) if row["confidence_score"] is not None else 0
        for optional in ("source_label", "record_id"):
            if optional in row:
                row[optional] = row[optional] or ""
        if "review_approved" in row:
            row["review_approved"] = bool(row["review_approved"])
        got.append(row)
    if got != records:
        raise ExportError(f"{description} XLSX integrity check failed after re-reading.")
    return data, len(records)


XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
METADATA_FIELDS = ("document_name", "label", "value", "confidence_score")
METADATA_HEADERS = list(METADATA_FIELDS)
VALIDATED_METADATA_SCOPE = "__validated_main_metadata__"

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
        .chip.processing { color:#ffe3a0; border-color:rgba(242,196,109,.65);
                background:rgba(105,77,20,.22); }
        .topbar-throbber { width:13px; height:13px; flex:0 0 13px; display:inline-block;
                border:2px solid rgba(255,227,160,.35); border-top-color:#ffe3a0;
                border-right-color:#f2c46d; border-radius:50%;
                animation:parchi-topbar-spin .72s linear infinite; }
        @keyframes parchi-topbar-spin { to { transform:rotate(360deg); } }
        @media (prefers-reduced-motion: reduce) {
            .topbar-throbber { animation:none; border-color:#f2c46d; border-top-color:transparent; }
        }

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
        /* Value Grouping board: blend the section into the page instead of boxing it in a
           panel — the drag-and-drop frame provides its own card styling. */
        div[data-testid="stVerticalBlockBorderWrapper"]:has(iframe[title*="parchi_value_groups"]) {
            background: transparent !important; background-color: transparent !important;
            border-color: transparent !important; box-shadow: none !important; border-radius: 0;
            width: calc(100% + 2rem); max-width: none; margin-left: -1rem; }
        div[data-testid="stElementContainer"]:has(iframe[title*="parchi_value_groups"]) {
            background: transparent !important; background-color: transparent !important; }
        iframe[title*="parchi_value_groups"] { width: 100% !important; }
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
    ids = {d["id"] for d in dockets}
    active_id = st.session_state.get("active_docket")
    if active_id not in ids:
        try:
            query_id = st.query_params.get("parchi_docket")
        except Exception:  # tolerate older Streamlit runtimes and malformed URLs
            query_id = None
        active_id = query_id if query_id in ids else None
        if active_id:
            st.session_state["active_docket"] = active_id
    return next((d for d in dockets if d["id"] == active_id), None)


@st.fragment(run_every=1)
def render_topbar_status(docket_id: Optional[str]) -> None:
    """Refresh the Mode/status chips independently of the Workspace extraction fragment."""
    keys = len(key_pool())
    cooling = active_cooldown_count()
    rejected = len(rejected_keys())
    active_job = processing_job(docket_id) if docket_id else None
    processing_chip = ""
    if active_job and active_job["status"] in {"queued", "running"}:
        processing_chip = (
            '<span class="chip processing" role="status" aria-live="polite" '
            'title="Document extraction is active">'
            '<span class="topbar-throbber" aria-hidden="true"></span>'
            'Processing documents, please wait</span>'
        )
    extra = ""
    if cooling:
        extra += f'<span class="chip warn">▲ {cooling} key(s) cooling</span>'
    if rejected:
        extra += f'<span class="chip bad">✖ {rejected} invalid key(s)</span>'
    st.markdown(
        f'<div class="chips"><span class="chip accent">MODE: GEMINI</span>'
        f'{processing_chip}<span class="chip ok">✔ {keys} key(s) ready</span>{extra}</div>',
        unsafe_allow_html=True,
    )


def render_topbar(dockets: List[sqlite3.Row]) -> Optional[sqlite3.Row]:
    if dockets:
        ids = [d["id"] for d in dockets]
        if st.session_state.get("active_docket") not in ids:
            try:
                query_id = st.query_params.get("parchi_docket")
            except Exception:
                query_id = None
            st.session_state["active_docket"] = query_id if query_id in ids else ids[0]
        selected_docket = str(st.session_state["active_docket"])
        try:
            if st.query_params.get("parchi_docket") != selected_docket:
                st.query_params["parchi_docket"] = selected_docket
        except Exception:
            pass
    else:
        try:
            if "parchi_docket" in st.query_params:
                del st.query_params["parchi_docket"]
        except Exception:
            pass
    if st.session_state.get("nav") not in PAGES:
        try:
            query_page = st.query_params.get("parchi_page")
        except Exception:
            query_page = None
        st.session_state["nav"] = query_page if query_page in PAGES else PAGES[0]
    try:
        if st.query_params.get("parchi_page") != st.session_state["nav"]:
            st.query_params["parchi_page"] = st.session_state["nav"]
    except Exception:
        pass
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
        render_topbar_status(str(current_docket["id"]) if current_docket is not None else None)

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


def unique_labels(records: Iterable[Any]) -> List[str]:
    """Every distinct label in first-seen order, original spelling preserved."""
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


def exact_template_label(source_label: str, requested_labels: Iterable[str]) -> Optional[str]:
    """The template label equal to `source_label` ignoring case and surrounding whitespace."""
    normalized = normalize_label(source_label)
    return next(
        (str(label) for label in requested_labels if normalize_label(label) == normalized),
        None,
    )


def documents_by_label(records: Iterable[Any]) -> Dict[str, List[str]]:
    """Map each normalized label to the sorted document names where it appears."""
    found: Dict[str, set] = {}
    for record in records:
        normalized = normalize_label(record["label"])
        if normalized:
            found.setdefault(normalized, set()).add(record["document_name"])
    return {key: sorted(names, key=str.casefold) for key, names in found.items()}

def _dedupe_labels(labels: Iterable[str]) -> List[str]:
    """Collapse repeats of the same label (any casing or spacing) to the first spelling."""
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


def _norm_set(labels: Iterable[str]) -> set:
    return {normalize_label(label) for label in labels}


def _grouping_facts(active: sqlite3.Row) -> Tuple[Dict[str, List[Dict[str, Any]]], List[dict]]:
    """Ungrouped, non-rejected facts for the grouping board.

    Returns (records grouped by document name for the drag zones, flat fact rows for the
    add-members picker). Grouping membership itself comes from value_groups.load_groups;
    a fact that already belongs to a group never appears in a drag zone."""
    docket_id = active["id"]
    with connect() as connection:
        allowed = ungrouped_fact_ids(connection, docket_id)
    documents: Dict[str, List[Dict[str, Any]]] = {}
    flat: List[dict] = []
    for row in fetch_facts(docket_id):
        if row["id"] not in allowed:
            continue
        item = {
            "id": row["id"],
            "label": row["label"],
            "value": row["value"],
            "score": int(row["confidence_score"]),
            "document": row["filename"],
            "document_id": row["document_id"],
        }
        document_key = f"{row['filename']} · {row['document_id'][:8]}"
        documents.setdefault(document_key, []).append(item)
        flat.append(item)
    return documents, flat


def _render_grouping(active: sqlite3.Row) -> None:
    """Manual value grouping: drag equivalent records across documents into groups.

    Grouping is entirely user-driven — nothing is suggested, matched or grouped
    automatically. Main Metadata is only read here; saving writes the group tables."""
    docket_id = active["id"]
    if _value_groups_component is None:
        st.error("The drag-and-drop board (value_group_component/) is missing next to "
                 "app.py, so records cannot be grouped visually. Reinstall the app files "
                 "or contact support.", icon="🚫")
        return

    documents, flat = _grouping_facts(active)
    still_free = {item["id"] for item in flat}
    draft_key = f"vg_draft_{docket_id}"
    saved_draft = [
        str(fact_id) for fact_id in dict.fromkeys(st.session_state.get(draft_key, []))
        if str(fact_id) in still_free
    ]
    st.session_state[draft_key] = saved_draft
    with connect() as connection:
        groups = load_groups(connection, docket_id)
    group_views = [
        {
            "id": group["id"],
            "name": group["name"],
            "primary": group["primary_label"],
            "members": [
                {"id": m.fact_id, "label": m.label, "value": m.value, "document": m.document_name,
                 "score": m.confidence_score}
                for m in group["members"]
            ],
        }
        for group in groups
    ]

    if not documents and not groups:
        empty_state("Nothing to group", "Every record is either grouped or rejected. "
                    "Re-extract documents in the Workspace to add more records.")
        return

    result = _value_groups_component(
        documents=documents,
        groups=group_views,
        draft=saved_draft,
        storage_key=f"parchi_value_group_draft_{docket_id}",
        reset=int(st.session_state.get("vg_reset", 0)),
        key=f"vg_board_{docket_id}",
    )
    result = result or {}
    component_draft = result.get("draft")
    if isinstance(component_draft, list):
        st.session_state[draft_key] = [
            str(fact_id) for fact_id in dict.fromkeys(component_draft)
            if str(fact_id) in still_free
        ]
    action = result.get("action")
    event_id = str(result.get("event_id") or "")
    event_key = f"vg_last_event_{docket_id}"
    if action in {"create_group", "add_member", "remove_member", "reorder_members", "rename_group", "delete_group"} and event_id \
            and st.session_state.get(event_key) != event_id:
        st.session_state[event_key] = event_id
        try:
            with connect() as connection:
                if action == "create_group":
                    selected_ids = result.get("fact_ids")
                    if not isinstance(selected_ids, list) or not selected_ids:
                        selected_ids = result.get("draft")
                    if not isinstance(selected_ids, list) or not selected_ids:
                        selected_ids = st.session_state.get(draft_key, [])
                    create_group(connection, docket_id, str(result.get("name") or ""),
                                 [str(i) for i in dict.fromkeys(selected_ids or [])
                                  if str(i) in still_free])
                elif action == "add_member":
                    add_member(connection, docket_id, str(result.get("group_id") or ""),
                               str(result.get("fact_id") or ""))
                elif action == "remove_member":
                    remove_member(connection, docket_id, str(result.get("group_id") or ""),
                                  str(result.get("fact_id") or ""))
                elif action == "reorder_members":
                    reorder_members(connection, docket_id, str(result.get("group_id") or ""),
                                    result.get("fact_ids") or [])
                elif action == "rename_group":
                    rename_group(connection, docket_id, str(result.get("group_id") or ""),
                                 str(result.get("name") or ""))
                else:
                    delete_group(connection, docket_id, str(result.get("group_id") or ""))
        except GroupError as exc:
            queue_flash("error", str(exc))
        except sqlite3.Error as exc:
            LOG.exception("value group action failed for docket %s", docket_id)
            queue_flash("error", f"Could not save the group in the database: {clean_text(exc)[:200]}. "
                        "Your selected occurrences are still in Group Values; correct the issue and retry.")
        except Exception as exc:  # keep the draft visible if a component/backend error occurs
            LOG.exception("unexpected value group action failure for docket %s", docket_id)
            queue_flash("error", f"Could not save the group: {clean_text(exc)[:200]}. "
                        "Your selected occurrences are still in Group Values; retry after correcting the issue.")
        else:
            if action == "create_group":
                st.session_state[draft_key] = []
                st.session_state["vg_reset"] = int(st.session_state.get("vg_reset", 0)) + 1
            messages = {
                "create_group": "Group saved. Its records stay in Main Metadata untouched.",
                "add_member": "Occurrence added to the group.",
                "remove_member": "Occurrence returned to its document. Main Metadata is unchanged.",
                "reorder_members": "Occurrence priority order saved.",
                "rename_group": "Group renamed.",
                "delete_group": "Group deleted. Its occurrences returned to their documents; Main Metadata is unchanged.",
            }
            queue_flash("success", messages[action])
        st.rerun()

def _render_create(active: sqlite3.Row, metadata: List[dict]) -> None:
    """Create an ordered template from live group or label sources."""
    docket_id = active["id"]
    with connect() as connection:
        groups = load_groups(connection, docket_id)
        all_rows = connection.execute(
            """SELECT f.id, f.label, f.review_status FROM facts f WHERE f.docket_id = ?
               AND f.concept != 'page_processing_status' ORDER BY f.label COLLATE NOCASE""",
            (docket_id,),
        ).fetchall()
        group_rows = connection.execute(
            "SELECT group_id, fact_id FROM value_group_members"
        ).fetchall()
    group_members = {}
    for row in group_rows:
        group_members.setdefault(row["group_id"], set()).add(row["fact_id"])
    selectable_groups = [g for g in groups if g["members"]]
    grouped_ids = set().union(*(group_members.get(g["id"], set()) for g in selectable_groups)) \
        if selectable_groups else set()
    label_ids = {}
    for row in all_rows:
        if row["id"] in grouped_ids:
            continue
        label_ids.setdefault(row["label"], []).append(row["id"])

    st.markdown("**Grouped Values**")
    selected_groups = st.multiselect(
        "Group sources", [g["id"] for g in selectable_groups],
        format_func=lambda gid: next(g["name"] for g in selectable_groups if g["id"] == gid),
        key=f"org_create_groups_{docket_id}",
    )
    st.markdown("**Ungrouped Values**")
    selected_labels = st.multiselect(
        "Original label sources", list(label_ids), key=f"org_create_labels_{docket_id}"
    )
    sources = [("group", gid, next(g["name"] for g in selectable_groups if g["id"] == gid))
               for gid in selected_groups]
    sources += [("label", label, label) for label in selected_labels]
    # Each source keeps its own name as its label, in the order it was added.
    fields = [
        {"source_type": source_type, "source_id": source_id, "source_name": source_name,
         "destination_label": source_name, "position": position + 1}
        for position, (source_type, source_id, source_name) in enumerate(sources)
    ]
    name = st.text_input("Template name", key=f"org_create_name_{docket_id}", max_chars=120)
    if st.button("Save template", type="primary", key=f"org_create_save_{docket_id}",
                 disabled=not fields or not name.strip(), **stretch()):
        if any(not field["destination_label"].strip() for field in fields):
            st.error("Every source needs a destination label.", icon="🚫")
            return
        try:
            template_id = save_template(docket_id, name, [f["destination_label"] for f in fields])
            save_template_fields(template_id, fields)
        except ValueError as exc:
            st.error(str(exc), icon="🚫")
            return
        queue_flash("success", f"Template “{clean_text(name).strip()}” saved with live sources.")
        st.rerun()


def _render_apply(active: sqlite3.Row, metadata: List[dict]) -> None:
    """Apply a saved template. Tabs run in this order: Decided, Suggested, Unmatched."""
    docket_id = active["id"]
    templates = list_templates(docket_id)
    if not templates:
        st.info("No templates yet. Create one first.")
        return

    pick_col, docs_col = st.columns([2, 3], vertical_alignment="bottom")
    selected = pick_col.selectbox("Template", [row["name"] for row in templates],
                                  key=f"org_apply_pick_{docket_id}")
    template = next(row for row in templates if row["name"] == selected)
    template_id = template["id"]
    template_labels = json.loads(template["labels_json"])
    fact_rows = fetch_facts(docket_id)
    document_names = {}
    for fact in fact_rows:
        document_names[fact["document_id"]] = fact["filename"]
    document_ids = sorted(document_names, key=lambda did: (document_names[did].casefold(), did))
    chosen = docs_col.multiselect(
        "Documents", document_ids, default=document_ids,
        format_func=lambda did: (f"{document_names[did]} · {did[:8]}"
                                 if sum(name == document_names[did] for name in document_names.values()) > 1
                                 else document_names[did]),
        key=f"org_apply_docs_{docket_id}_{template_id}")
    if not chosen:
        st.info("Select at least one document.")
        return

    chosen_records = [
        {"document_name": fact["filename"], "label": fact["label"], "value": fact["value"],
         "confidence_score": int(fact["confidence_score"]), "record_id": fact["id"]}
        for fact in fact_rows if fact["document_id"] in set(chosen)
        and fact["concept"] != "page_processing_status"
    ]
    labels = unique_labels(chosen_records)
    sources_map = documents_by_label(chosen_records)

    def from_docs(source: str) -> str:
        names = sources_map.get(normalize_label(source), [])
        return ", ".join(names) if names else "—"

    current = approved_mappings(docket_id, template_id)
    current_normalized = {normalize_label(source): target for source, target in current.items()}
    rejected = rejected_suggestions(docket_id, template_id)
    rejected_normalized = {(normalize_label(source), normalize_label(target))
                           for source, target in rejected}

    suggested: Dict[str, List[Dict[str, Any]]] = {}
    unmatched: List[str] = []
    for source in labels:
        exact_target = exact_template_label(source, template_labels)
        if exact_target is not None:
            if normalize_label(source) not in current_normalized:
                save_mapping(docket_id, template_id, source, exact_target, "exact")
            continue
        if normalize_label(source) in current_normalized:
            continue
        options = [m for m in fuzzy_suggestions(source, template_labels)
                   if (normalize_label(source), normalize_label(m["label"])) not in rejected_normalized]
        if options:
            suggested[source] = options
        else:
            unmatched.append(source)

    details = mapping_details(docket_id, template_id)

    st.info(
        "**How mapping works:** the left label (**source**) comes from Main Metadata, "
        "the right label (**target**) is a field in your template. "
        "The output file uses the target name."
    )

    decided_tab, suggested_tab, unmatched_tab = st.tabs([
        f"Decided ({len(details)})",
        f"Suggested ({sum(len(v) for v in suggested.values())})",
        f"Unmatched ({len(unmatched)})",
    ])

    with decided_tab:
        st.subheader("Decided")
        if not details:
            st.caption("No mappings yet.")
        for source, info in sorted(details.items(), key=lambda item: item[0].casefold()):
            text_col, unmap_col = st.columns([5, 1], vertical_alignment="center")
            text_col.markdown(
                f"**{source}**  →  **{info['target']}**  \n"
                f"<small>from: {esc(from_docs(source))}</small>",
                unsafe_allow_html=True,
            )
            if unmap_col.button("Unmap", key=f"tpl_unmap_{docket_id}_{template_id}_{key_id(source)}",
                                **stretch()):
                delete_mapping(docket_id, template_id, source, info["target"])
                st.rerun()

    with suggested_tab:
        st.subheader("Suggested")
        if not suggested:
            st.caption("No suggestions.")
        for source, options in suggested.items():
            for match in options:
                target = match["label"]
                pid = key_id(f"{source}\u241f{target}")
                text_col, approve_col, reject_col = st.columns([4, 1, 1], vertical_alignment="center")
                text_col.markdown(
                    f"**{source}**  →  **{target}**  ({match['score']:.0f}%)  \n"
                    f"<small>from: {esc(from_docs(source))}</small>",
                    unsafe_allow_html=True,
                )
                if approve_col.button("Approve", key=f"tpl_approve_{docket_id}_{template_id}_{pid}",
                                      **stretch()):
                    save_mapping(docket_id, template_id, source, target, "user-approved")
                    st.rerun()
                if reject_col.button("Reject", key=f"tpl_reject_{docket_id}_{template_id}_{pid}",
                                     **stretch()):
                    reject_suggestion(docket_id, template_id, source, target)
                    delete_mapping(docket_id, template_id, source, target)
                    st.rerun()

    with unmatched_tab:
        st.subheader("Unmatched")
        if not unmatched:
            st.caption("Nothing unmatched.")
        for source in unmatched:
            label_col, target_col, map_col, add_col = st.columns([2, 2, 1, 1.5], vertical_alignment="bottom")
            label_col.markdown(f"**{source}**  \n<small>from: {esc(from_docs(source))}</small>",
                               unsafe_allow_html=True)
            occurrences = [record for record in chosen_records
                           if normalize_label(record["label"]) == normalize_label(source)]
            with st.expander(f"Inspect {len(occurrences)} occurrence(s) · {source}"):
                st.dataframe(
                    [{"Source document": row["document_name"], "Original label": row["label"],
                      "Value": row["value"], "Confidence score": row["confidence_score"]}
                     for row in occurrences],
                    hide_index=True, **stretch(),
                )
            target = target_col.selectbox(f"Field for {source}", template_labels,
                                          key=f"tpl_manual_t_{docket_id}_{template_id}_{key_id(source)}",
                                          label_visibility="collapsed")
            if map_col.button("Map", key=f"tpl_manual_b_{docket_id}_{template_id}_{key_id(source)}",
                              **stretch()):
                save_mapping(docket_id, template_id, source, target, "manual")
                st.rerun()
            if add_col.button("Add as field", key=f"tpl_add_{docket_id}_{template_id}_{key_id(source)}",
                              **stretch()):
                append_template_labels(docket_id, template_id, [source])
                queue_flash("success", f"“{source}” added to {template['name']}.")
                st.rerun()

def _render_manage(active: sqlite3.Row, metadata: List[dict]) -> None:
    """Rename a template, change its fields, or delete it."""
    docket_id = active["id"]
    templates = list_templates(docket_id)
    if not templates:
        st.caption("No saved templates yet.")
        return
    selected = st.selectbox("Template", [row["name"] for row in templates],
                            key=f"org_manage_pick_{docket_id}")
    template = next(row for row in templates if row["name"] == selected)
    template_id = template["id"]
    current = json.loads(template["labels_json"])
    source_fields = template_field_rows(template_id)
    addable = [label for label in unique_labels(metadata) if normalize_label(label) not in _norm_set(current)]

    kept = st.pills("Fields", current, selection_mode="multi", default=current,
                    key=f"org_manage_fields_{template_id}_{len(current)}",
                    label_visibility="collapsed", width="stretch")
    kept = list(kept) if kept is not None else list(current)
    to_add = st.multiselect("Add fields from extracted labels", addable,
                            key=f"org_manage_add_{template_id}_{len(current)}")
    name = st.text_input("Template name", value=template["name"], key=f"org_manage_name_{template_id}",
                         max_chars=120)
    new_labels = _dedupe_labels([*kept, *to_add])
    edited_sources = []
    if source_fields:
        st.markdown("**Source fields and order**")
        active_source_fields = [field for field in source_fields
                                if field["destination_label"] in kept]
        for index, field in enumerate(active_source_fields):
            source_col, destination_col, order_col = st.columns([2, 3, 1])
            source_col.caption(f"{field['source_type'].title()} · {field['source_name']}")
            destination = destination_col.text_input(
                "Destination", value=field["destination_label"],
                key=f"org_edit_dest_{field['id']}",
            )
            order = order_col.number_input(
                "Order", min_value=1, value=index + 1, step=1,
                key=f"org_edit_order_{field['id']}",
            )
            edited_sources.append({
                "source_type": field["source_type"], "source_id": field["source_id"],
                "source_name": field["source_name"], "destination_label": destination,
                "position": int(order),
            })
        existing_sources = {(field["source_type"], field["source_id"]) for field in edited_sources}
        for label in to_add:
            if ("label", label) not in existing_sources:
                edited_sources.append({
                    "source_type": "label", "source_id": label, "source_name": label,
                    "destination_label": label, "position": len(edited_sources) + 1,
                })
        edited_sources.sort(key=lambda item: (item["position"], item["source_name"].casefold()))
        new_labels = _dedupe_labels([field["destination_label"] for field in edited_sources])
    changed = new_labels != current or name.strip() != template["name"]
    save_col, delete_col = st.columns(2)
    if save_col.button("Save changes", type="primary", key=f"org_manage_save_{template_id}",
                       disabled=not changed or not new_labels or not name.strip(), **stretch()):
        try:
            save_template(docket_id, name, new_labels, template_id=template_id)
            if source_fields:
                save_template_fields(template_id, edited_sources)
        except ValueError as exc:
            st.error(str(exc), icon="🚫")
            return
        for label in new_labels:
            if label not in current:
                save_mapping(docket_id, template_id, label, label, "direct")
        queue_flash("success", "Template updated. Main Metadata was not changed.")
        st.rerun()
    with delete_col.popover("Delete template", **stretch()):
        st.caption("Also removes this template's mappings and rejected suggestions. Main Metadata is untouched.")
        confirm_key = f"org_delete_confirm_{template_id}"
        st.checkbox("I understand this cannot be undone", key=confirm_key)
        if st.button("Delete permanently", type="primary", key=f"org_delete_{template_id}",
                     disabled=not st.session_state.get(confirm_key, False), **stretch()):
            delete_template(docket_id, template_id)
            queue_flash("success", f"Template “{template['name']}” deleted.")
            st.rerun()


ORGANIZE_MODES = ("Group values", "Create new template", "Apply existing template")


def page_organize(dockets: List[sqlite3.Row], active: Optional[sqlite3.Row]) -> None:
    page_header("Organize", "Group equivalent records manually, or build and apply "
                "reusable client templates. Main Metadata is never modified here.",
                active["name"] if active else None)
    if not need_docket(active):
        return
    metadata = metadata_records(active["id"])
    if not metadata:
        st.info("Process documents first. Once Main Metadata has facts, templates can be built here.")
        return
    st.caption(
        f"{len({record['document_name'] for record in metadata})} document(s) · "
        f"{len(unique_labels(metadata))} unique label(s) · "
        f"{len(list_templates(active['id']))} saved template(s)"
    )
    mode = st.radio("Workflow", list(ORGANIZE_MODES), horizontal=True, key="org_mode",
                    label_visibility="collapsed")
    if mode == ORGANIZE_MODES[0]:
        _render_grouping(active)
    elif mode == ORGANIZE_MODES[1]:
        _render_create(active, metadata)
    else:
        _render_apply(active, metadata)
    if mode != ORGANIZE_MODES[0]:
        with st.expander("Manage saved templates"):
            _render_manage(active, metadata)


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


def viewer_frame(rows: List[dict], show_status: bool = False) -> None:
    """Render metadata records with stable, human-readable headers (never raw
    snake_case keys) and colour-coded confidence cells.

    `show_status` adds a Review Status column (display-only traceability); it is never
    part of the exported four-field schema."""
    template_columns = ["Source Label", "Record ID"] if any("source_label" in row for row in rows) else []
    columns = VIEWER_COLUMNS + template_columns + (["Review Status"] if show_status else [])
    frame = pd.DataFrame(
        [
            {
                "Document Name": record["document_name"],
                "Label": record["label"],
                "Value": record["value"],
                "Confidence Score": clamp_confidence(record["confidence_score"], 50),
                "Source Label": record.get("source_label", ""),
                "Record ID": record.get("record_id", ""),
                "Review Status": record.get("review_status", ""),
            }
            for record in rows
        ],
        columns=columns,
    )
    st.dataframe(
        frame.style.map(confidence_score_style, subset=["Confidence Score"]),
        hide_index=True,
        width="stretch",
        height=min(720, max(180, 44 + 36 * len(frame))),
    )


def viewer_template_editor(docket_id: str, template_id: str, rows: List[dict]) -> None:
    """Editable Review Metadata view backed only by the review-data sidecar."""
    has_status = any("review_status" in row for row in rows)
    data = [{
            "Document Name": row["document_name"],
            "Label": row["label"],
            "Value": row["value"],
            "Confidence Score": int(row["confidence_score"]),
            "Approved": bool(row.get("review_approved", False)),
            "Source Label": row.get("source_label", ""),
            "Record ID": row["record_id"],
            **({"Review Status": row.get("review_status", "")} if has_status else {}),
        } for row in rows]
    columns = ["Approved", "Document Name", "Label", "Value", "Confidence Score", "Source Label", "Record ID"]
    if has_status:
        columns.append("Review Status")
    frame = pd.DataFrame(data, columns=columns)
    styled_frame = frame.style.map(confidence_score_style, subset=["Confidence Score"])
    revision_key = f"viewer_editor_revision_{docket_id}_{template_id}"
    revision = int(st.session_state.get(revision_key, 0))
    edited = st.data_editor(
        styled_frame,
        key=f"viewer_editor_{docket_id}_{template_id}_{revision}",
        hide_index=True,
        num_rows="fixed",
        width="stretch",
        height=min(720, max(180, 44 + 36 * len(frame))),
        column_config={
            "Approved": st.column_config.CheckboxColumn("Approved", help="Approve this label and value."),
            "Document Name": st.column_config.TextColumn("Document Name"),
            "Label": st.column_config.TextColumn("Label", required=True),
            "Value": st.column_config.TextColumn("Value"),
            "Confidence Score": st.column_config.NumberColumn("Confidence Score"),
            "Source Label": st.column_config.TextColumn("Source Label"),
            "Record ID": None,
            **({"Review Status": st.column_config.TextColumn("Review Status")} if has_status else {}),
        },
        disabled=["Document Name", "Confidence Score", "Source Label", "Record ID"]
                 + (["Review Status"] if has_status else []),
    )
    baseline_by_id = {row["record_id"]: row for row in rows}
    changes = []
    approval_changes = []
    for _, edited_row in edited.iterrows():
        record_id = str(edited_row["Record ID"])
        baseline = baseline_by_id.get(record_id)
        if baseline is None:
            continue
        label_cell = edited_row["Label"]
        value_cell = edited_row["Value"]
        label = "" if pd.isna(label_cell) else str(label_cell)
        value = "" if pd.isna(value_cell) else str(value_cell)
        if label != str(baseline["label"]) or value != str(baseline["value"]):
            changes.append((record_id, label, value))
        approved = bool(edited_row["Approved"])
        if approved != bool(baseline.get("review_approved", False)):
            approval_changes.append((record_id, approved))
    if changes or approval_changes:
        try:
            with connect() as connection:
                for record_id, label, value in changes:
                    edit_validated_record(connection, docket_id, template_id, rows,
                                          record_id, label, value)
                for record_id, approved in approval_changes:
                    set_review_approval(connection, docket_id, template_id, rows,
                                        record_id, approved)
        except (ValueError, sqlite3.Error) as exc:
            st.error(f"Could not save Viewer edits: {clean_text(exc)}", icon="🚫")
        else:
            st.session_state[revision_key] = revision + 1
            queue_flash("success", f"Saved {len(changes)} edit(s) and {len(approval_changes)} approval change(s). "
                        "Original Main Metadata is unchanged.")
            st.rerun()


def render_viewer_edit_history(docket_id: str) -> None:
    with connect() as connection:
        history = list_validated_edit_history(connection, docket_id)
    st.markdown("### Edit History")
    if not history:
        st.caption("No Viewer edits have been made.")
        return
    st.caption("Each field edit is recorded separately. Undo restores only that field edit; "
               "Main Metadata remains unchanged.")
    for item in history:
        detail_col, time_col, action_col = st.columns([5, 1.3, 1.1], vertical_alignment="center")
        scope_name = ("Review Metadata" if item["template_id"] == VALIDATED_METADATA_SCOPE
                      else f"template {item['template_id'][:8]}")
        detail_col.write(f"{item['document_name'] or 'Unknown document'} · "
                         f"{item['source_label'] or 'Record'} · {item['field_name'].title()}")
        detail_col.caption(f"{item['old_value']} → {item['new_value']} · "
                           f"record {item['record_id'][:8]} · {scope_name}")
        time_col.caption(display_time(item["created_at"]))
        if item["undone_at"]:
            action_col.caption(f"Undone · {display_time(item['undone_at'])}")
        elif action_col.button("Undo", key=f"viewer_undo_{item['id']}"):
            try:
                with connect() as connection:
                    undo_validated_edit(connection, docket_id, item["id"])
            except (ValueError, sqlite3.Error) as exc:
                st.error(f"Could not undo this edit: {clean_text(exc)}", icon="🚫")
            else:
                revision_key = f"viewer_editor_revision_{docket_id}_{item['template_id']}"
                st.session_state[revision_key] = int(st.session_state.get(revision_key, 0)) + 1
                queue_flash("success", "The selected Viewer edit was undone. Main Metadata is unchanged.")
                st.rerun()


def page_viewer(dockets: List[sqlite3.Row], active: Optional[sqlite3.Row]) -> None:
    page_header(
        "Viewer",
        "Edit Review Metadata separately from the untouched original — duplicates, "
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

    view_options = ["Review Metadata", "Original Main Metadata"]
    if templates:
        view_options.append("Client Template")
    selected_view = st.selectbox(
        "Table to view",
        view_options,
        index=0,
        key="viewer_table_mode_v2",
        help="Review Metadata is editable and saved separately. Original Main Metadata is read-only. "
             "Client Template shows records mapped to a saved template.",
    )

    if selected_view == "Review Metadata":
        baseline_rows = viewer_validated_base_records(docket_id)
        with connect() as connection:
            rows = apply_validated_records(
                connection, docket_id, VALIDATED_METADATA_SCOPE, baseline_rows
            )
        st.caption(f"Editable review copy · {len(rows)} occurrence(s). Approve each reviewed label and value "
                   "are editable; the original Main Metadata remains untouched.")
        editor_template_id = VALIDATED_METADATA_SCOPE
    elif selected_view == "Original Main Metadata":
        rows = viewer_metadata_records(docket_id)
        st.caption(f"Original extraction · {len(rows)} row(s), including duplicates, conflicting "
                   "values, and pending or rejected records. This view is read-only.")
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
            baseline_rows, missing = client_template_records(docket_id, template)
        with connect() as connection:
            rows = apply_validated_records(connection, docket_id, template["id"], baseline_rows)
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
        editor_template_id = template["id"]

    if selected_view == "Original Main Metadata":
        viewer_frame(rows, show_status=True)
    elif rows:
        st.caption("Only Label, Value and Approved are editable. These changes apply to Review Metadata; "
                   "edited records receive confidence 100, while source facts remain untouched.")
        viewer_template_editor(docket_id, editor_template_id, rows)
    render_viewer_edit_history(docket_id)


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
    page_header("Export", "Download untouched Main Metadata or the separately reviewed Review Metadata. "
                          "Files use the saved template mappings.", active["name"] if active else None)
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
            file_name=f"parchi-{slug}-{prefix}-{stamp}.{extension}",
            mime=mime_by_format[chosen], key=f"{prefix}_download", **stretch(),
        )

    with st.container(border=True):
        export_panel("Main Metadata", "The complete, unfiltered four-field record set.",
                     metadata_records(active["id"]), [], "main_metadata")
    validated_base = viewer_validated_base_records(active["id"])
    with connect() as connection:
        validated = apply_validated_records(
            connection, active["id"], VALIDATED_METADATA_SCOPE, validated_base
        )
    review_export = [
        {**{field: row[field] for field in METADATA_FIELDS},
         "review_approved": bool(row.get("review_approved", False))}
        for row in validated
    ]
    with st.container(border=True):
        export_panel("Review Metadata", "The full editable metadata copy, including saved Viewer edits. "
                     "Unedited occurrences retain their original confidence scores.",
                     review_export, [], "review_metadata")
    templates = list_templates(active["id"])
    with st.container(border=True):
        if templates:
            names = [row["name"] for row in templates]
            chosen_name = st.selectbox("Client template", names, key="export_template_choice")
            template = next(row for row in templates if row["name"] == chosen_name)
            baseline_records, missing = client_template_records(active["id"], template)
            with connect() as connection:
                records = apply_validated_records(connection, active["id"], template["id"],
                                                  baseline_records)
            mapped_ids = {record.get("record_id") for record in records if record.get("record_id")}
            unmatched_sources = [
                {"record_id": fact["id"], "document_name": fact["filename"],
                 "label": fact["label"], "value": fact["value"],
                 "confidence_score": int(fact["confidence_score"])}
                for fact in fetch_facts(active["id"])
                if fact["concept"] != "page_processing_status" and fact["id"] not in mapped_ids
            ]
            export_panel(f"Client Template · {template['name']}",
                         "Mapped occurrences keep their source document, original source label and record ID.",
                         records, missing, "client_template")
            if unmatched_sources:
                with st.expander(f"Unmatched occurrences remain in Main Metadata ({len(unmatched_sources)})"):
                    st.dataframe(unmatched_sources, hide_index=True, **stretch())
        else:
            st.markdown("**Client Template**")
            st.info("Create a template in Organize first.")


def page_settings(dockets: List[sqlite3.Row], active: Optional[sqlite3.Row]) -> None:
    page_header("Settings", "Gemini key pool, model fallback, storage and integrity checks.", None)
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
    st.caption("Override with “gemini_models” in the config file or the PARCHI_GEMINI_MODELS "
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
    try:
        init_db()
        dockets = list_dockets()
        active = render_topbar(dockets)
        # render_topbar restores both docket and page from the URL after a browser
        # refresh, so a live run returns to the same workspace and loader.
        page = st.session_state.get("nav", PAGES[0])
        if page not in RENDERERS:
            page = PAGES[0]
        render_flash()
        # Live processing banner on every page: throbber, overall progress and the
        # per-document queue for the active docket's run (see render_processing_loader).
        if active is not None:
            render_processing_loader(active["id"])
        RENDERERS[page](dockets, active)
    except Exception as exc:  # noqa: BLE001  (error boundary)
        if is_control_flow(exc):
            raise
        LOG.exception("page %s failed", page)
        render_error_panel(exc, page)


if __name__ == "__main__":
    main()
