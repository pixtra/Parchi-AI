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
import traceback
import uuid
from contextlib import contextmanager
from datetime import datetime, timezone
from itertools import islice
from pathlib import Path
from typing import Any, Callable, Dict, Iterator, List, Optional, Tuple

import streamlit as st
import pandas as pd
from psv_workflow import (
    build_excel,
    clean_psv,
    label_value_mismatch,
    next_pending_index,
    auto_verified_indices,
)

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


# ---------------------------------------------------------------- configuration
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = Path(os.getenv("PARCHAI_DB", BASE_DIR / "parchai_v3.db"))
STORAGE_DIR = Path(os.getenv("PARCHAI_STORAGE", BASE_DIR / "parchai_storage")).resolve()
GEMINI_MODELS = tuple(
    m.strip()
    for m in os.getenv(
        "PARCHAI_GEMINI_MODELS",
        "gemini-3.8-flash,gemini-3.6-flash,gemini-3.1-flash-lite",
    ).split(",")
    if m.strip()
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
SUPPORTED_TYPES = ["pdf", "png", "jpg", "jpeg", "txt"]
MIME_TYPES = {
    "pdf": "application/pdf",
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "txt": "text/plain",
}
BANDS = ("Green", "Yellow", "Red")
REVIEW_STATUSES = ("pending", "approved", "rejected")
PAGES = ["Dockets", "Workspace", "Review", "PSV Verify", "Validation", "Export", "Settings"]
PAGE_ICONS = {
    "Dockets": "▤", "Workspace": "⬆", "Review": "✎", "PSV Verify": "✓",
    "Validation": "✔", "Export": "⇩", "Settings": "⚙",
}
IDENTITY_CONCEPTS = {
    "invoice_number", "po_number", "purchase_order_number", "challan_number",
    "delivery_challan_number", "grn_number", "gstin", "vehicle_number",
    "supplier_name", "buyer_name", "invoice_date", "total_amount", "grand_total",
}
DOCUMENT_KEYWORDS = [
    ("Weighment slip", ("weighment", "gross weight", "tare weight", "net weight")),
    ("Goods receipt note", ("goods receipt", "grn")),
    ("Delivery challan", ("challan",)),
    ("Purchase order", ("purchase order",)),
    ("Transport receipt", ("lorry receipt", "transport", "consignment")),
    ("Invoice", ("invoice", "tax invoice")),
]
DOCKET_TYPES = (
    "Purchase",
    "Inventory Transfer In",
    "Inventory Transfer Out",
    "Sales",
    "Miscellaneous",
    "Payment",
)

# Fields written to PSV (one row per fact). Order is the column order.
PSV_FIELDS = [
    "fact_id", "filename", "page_number", "label", "concept", "value", "original_value",
    "source_evidence", "confidence_band", "confidence_score", "confidence_reasons",
    "review_status", "corrected", "updated_at",
]

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
    created_at TEXT NOT NULL,
    UNIQUE(docket_id, sha256)
);
CREATE TABLE IF NOT EXISTS pages (
    document_id TEXT NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
    page_number INTEGER NOT NULL,
    text TEXT NOT NULL,
    readable INTEGER NOT NULL,
    PRIMARY KEY (document_id, page_number)
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
"""

KV_RE = re.compile(r"^\s*([A-Za-z][A-Za-z0-9 /().#&'-]{1,48}?)\s*:\s*(\S.*?)\s*$")
PATTERNS = [
    ("gstin", "GSTIN", re.compile(r"\b\d{2}[A-Z]{5}\d{4}[A-Z][A-Z\d]Z[A-Z\d]\b")),
    ("date", "Date", re.compile(r"\b\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}\b")),
    ("amount", "Amount", re.compile(r"(?:₹|Rs\.?|INR)\s?\d[\d,]*(?:\.\d{1,2})?")),
    ("weight", "Weight", re.compile(r"\b\d+(?:\.\d+)?\s?(?:kgs?|MT|tons?|tonnes?)\b", re.IGNORECASE)),
]

BATCH_PROMPT = """You extract data from SEVERAL supplier documents in one request.
Each document part is preceded by a line "=== DOCUMENT: <label> ===". Use that label EXACTLY
as the "document" value. Page numbers restart at 1 for every labelled part.
Return ONLY a JSON object of this exact shape:
{"documents": [
   {"document": "<label exactly as given>", "document_type": "short type name",
    "facts": [{"page_number": 1, "label": "Human readable field name",
               "concept": "snake_case_name", "value": "value as written",
               "evidence": "text copied verbatim from that page (max 200 chars)"}]}
 ],
 "cross_references": [
   {"from_document": "<label>", "to_document": "<label>", "page_number": 1,
    "concept": "references_invoice", "value": "referenced number",
    "evidence": "verbatim text showing the link"}
 ]}
Rules:
- Extract every identifiable field: identifiers, parties, dates, references, quantities,
  weights, rates, taxes, totals and each line item (item_1_description, item_1_quantity, ...).
- Include document-specific fields that do not fit a standard schema.
- Copy values exactly as written, including handwritten entries when legible.
- Never invent or infer values that are not on the page. Use an empty value for illegible fields.
- Evidence must be copied verbatim from the named document and page.
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
    """Neutralise spreadsheet formula injection without touching genuine numbers."""
    text = clean_text(value)
    if text[:1] in ("=", "+", "-", "@", "\t", "\r") and not _NUMERIC_RE.match(text):
        return "'" + text
    return text


def esc(value: Any) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def fmt_size(num_bytes: int) -> str:
    return f"{num_bytes / 1024:,.0f} KB" if num_bytes < 1024 * 1024 else f"{num_bytes / 1048576:,.1f} MB"


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
        return f"Unsupported file type “.{extension or '?'}”. Use PDF, PNG, JPG or TXT."
    if not data:
        return "The file is empty (0 bytes)."
    if len(data) > MAX_UPLOAD_BYTES:
        return f"The file is {fmt_size(len(data))}; the limit is {fmt_size(MAX_UPLOAD_BYTES)}."
    signatures = {"png": b"\x89PNG\r\n\x1a\n", "jpg": b"\xff\xd8\xff", "jpeg": b"\xff\xd8\xff"}
    if extension == "pdf" and b"%PDF-" not in data[:1024]:
        return "The content is not a PDF. The file may be corrupt or renamed."
    if extension in signatures and not data.startswith(signatures[extension]):
        return f"The content is not a valid .{extension} image. The file may be corrupt or renamed."
    if extension == "txt" and b"\x00" in data:
        return "The file contains binary data and is not plain text."
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
                      extraction_mode, page_count, created_at
               FROM documents WHERE docket_id = ? ORDER BY created_at""",
            (docket_id,),
        ).fetchall()


def fetch_facts(docket_id: str) -> List[sqlite3.Row]:
    with connect() as connection:
        return connection.execute(
            """SELECT f.*, d.filename, d.storage_path, d.mime_type
               FROM facts f JOIN documents d ON d.id = f.document_id
               WHERE f.docket_id = ?
               ORDER BY d.filename, f.page_number, f.created_at""",
            (docket_id,),
        ).fetchall()


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
                pages.append({"page_number": index, "text": text, "readable": bool(text)})
        if not pages:
            raise RuntimeError("The PDF contains no pages.")
        return pages
    if extension == "txt":
        text = clean_text(decode_text(data)).strip()
        return [{"page_number": 1, "text": text, "readable": bool(text)}]
    return [{"page_number": 1, "text": "", "readable": False}]  # images: no text layer


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


def classify_local(pages: List[Dict[str, Any]]) -> str:
    text = normalise(" ".join(p["text"] for p in pages))
    for label, keywords in DOCUMENT_KEYWORDS:
        if any(keyword in text for keyword in keywords):
            return label
    return "Unclassified"


# ---------------------------------------------------------------- confidence
def band_for(score: int) -> str:
    return "Green" if score >= 90 else "Yellow" if score >= 60 else "Red"


def assess(value: str, evidence: str, page_text: Optional[str]) -> Dict[str, Any]:
    """Confidence reflects evidence quality. page_text is None when the page has no text
    layer, so the value cannot be checked against the source."""
    if not value.strip():
        return {"band": "Red", "score": 0, "reasons": ["MISSING"]}
    if not evidence.strip():
        return {"band": "Red", "score": 20, "reasons": ["NO_EVIDENCE"]}
    if page_text is None:
        return {"band": "Yellow", "score": 65, "reasons": ["NOT_TEXT_VERIFIABLE"]}
    text_n = normalise(page_text)
    if normalise(evidence) in text_n:
        score, reasons = 95, ["GROUNDED"]
    elif normalise(value) and normalise(value) in text_n:
        score, reasons = 80, ["VALUE_FOUND_EVIDENCE_MISMATCH"]
    else:
        return {"band": "Red", "score": 30, "reasons": ["UNSUPPORTED"]}
    if len(value.strip()) < 2:
        score = min(score, 75)
        reasons.append("VERY_SHORT_VALUE")
    return {"band": band_for(score), "score": score, "reasons": reasons}


# ---------------------------------------------------------------- local extraction
def _fact(page: int, label: str, concept: str, value: str, evidence: str) -> Dict[str, Any]:
    return {"page_number": page, "label": label[:120], "concept": concept,
            "value": value, "evidence": evidence}


def local_extract(pages: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    facts: List[Dict[str, Any]] = []
    for page in pages:
        number = page["page_number"]
        if not page["readable"]:
            facts.append(_fact(number, "Page readability", "page_status", "", ""))
            continue
        seen_pairs: set = set()
        seen_values: set = set()
        structured = 0
        for raw_line in page["text"].splitlines():
            line = raw_line.strip()
            if not line:
                continue
            match = KV_RE.match(line)
            if match:
                label, value = match.group(1).strip(), match.group(2).strip()
                pair = (to_concept(label), value.lower())
                if pair not in seen_pairs:
                    seen_pairs.add(pair)
                    seen_values.add(value.lower())
                    facts.append(_fact(number, label, to_concept(label), value, line))
                    structured += 1
            for concept, label, pattern in PATTERNS:
                for hit in pattern.finditer(line):
                    value = hit.group(0).strip()
                    if value.lower() in seen_values:
                        continue
                    seen_values.add(value.lower())
                    facts.append(_fact(number, label, concept, value, line))
                    structured += 1
        if structured == 0:
            first_line = next((ln.strip() for ln in page["text"].splitlines() if ln.strip()), "")
            facts.append(_fact(number, f"Page {number} text (unstructured)", "raw_text",
                               page["text"][:4000], first_line[:200]))
    return facts


# ---------------------------------------------------------------- API keys and errors
# Gemini keys are opaque ASCII tokens. Current keys use the "AQ." prefix and may contain
# periods in the token body; older AI Studio keys commonly use the "AIza" prefix.
KEY_RE = re.compile(r"(?:AQ\.[A-Za-z0-9._~-]{16,}|AIza[A-Za-z0-9._~-]{16,})")


class InvalidKeyError(Exception):
    """The API key is not shaped like a Gemini key."""


def valid_key(key: str) -> bool:
    return bool(key) and len(key) <= 256 and key.isascii() and bool(KEY_RE.fullmatch(key))


def key_problem(raw: str) -> Optional[str]:
    raw = (raw or "").strip()
    if not raw:
        return "Enter a key first."
    if len(raw) > 256 or re.search(r"\s", raw):
        return (f"That is not a single API key ({len(raw)} characters, with spaces or line breaks). "
                "Paste only the key itself; current Gemini keys usually start with “AQ.”.")
    if not raw.isascii() or not valid_key(raw):
        return ("That does not look like a Gemini API key. Paste the complete key beginning with “AQ.” "
                "(older keys may begin with “AIza”).")
    return None


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
    if isinstance(exc, (InvalidKeyError, UnicodeError)):
        return AUTH_COOLDOWN_SECONDS
    text = str(exc).upper()
    if "RESOURCE_EXHAUSTED" in text or "429" in text or "QUOTA" in text:
        return KEY_COOLDOWN_SECONDS
    if "API_KEY_INVALID" in text or "PERMISSION_DENIED" in text or "401" in text or "403" in text:
        return AUTH_COOLDOWN_SECONDS
    return None


def is_size_error(exc: Exception) -> bool:
    text = str(exc).lower()
    return "too large" in text or ("token" in text and ("exceed" in text or "limit" in text))


def cooldown_remaining(cooldown_id: str) -> float:
    until = st.session_state.get("key_cooldowns", {}).get(cooldown_id, 0.0)
    return max(0.0, until - time.time())


def set_cooldown(cooldown_id: str, seconds: int) -> None:
    st.session_state.setdefault("key_cooldowns", {})[cooldown_id] = time.time() + seconds


def _key_candidates() -> Iterator[Tuple[str, str]]:
    for raw in re.split(r"[\n,;]+", os.getenv("GEMINI_API_KEYS", "")):
        yield raw.strip(), "env pool"
    yield os.getenv("GEMINI_API_KEY", "").strip(), "env"
    for raw in st.session_state.get("session_keys", []):
        yield str(raw).strip(), "session"


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


def make_units(doc: Dict[str, Any]) -> List[Dict[str, Any]]:
    if extension_of(doc["filename"]) == "pdf" and fitz is not None:
        parts = split_pdf(doc["data"], CHUNK_PAGES)
    else:
        parts = [(0, doc["page_count"], doc["data"])]
    units = []
    for start, count, blob in parts:
        label = doc["filename"]
        if len(parts) > 1:
            label = f"{doc['filename']} [pages {start + 1}-{start + count}]"
        unit = {
            "label": label, "doc": doc, "offset": start, "pages": count, "data": blob,
            "mime": doc["mime"], "suffix": "." + extension_of(doc["filename"]),
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


def upload_parts(entry: Dict[str, str], batch: List[Dict[str, Any]]) -> Tuple[Any, List[Any]]:
    client = make_client(entry["key"])
    handles: List[Any] = []
    try:
        for unit in batch:
            with tempfile.NamedTemporaryFile(suffix=unit["suffix"], delete=False) as handle:
                handle.write(unit["data"])
                temporary = Path(handle.name)
            try:
                handles.append(client.files.upload(file=str(temporary), config={"mime_type": unit["mime"]}))
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
    by_label_lower = {label.lower(): label for label in labels}

    def match(name: str) -> Optional[str]:
        return name if name in labels else by_label_lower.get(name.strip().lower())

    def page_of(item: Dict[str, Any]) -> int:
        try:
            return max(1, int(item.get("page_number") or 1))
        except (TypeError, ValueError):
            return 1

    by_label: Dict[str, Dict[str, Any]] = {}
    for doc in payload["documents"]:
        if not isinstance(doc, dict):
            continue
        label = match(str(doc.get("document", "")))
        if label is None:
            continue
        bucket = by_label.setdefault(label, {"document_type": str(doc.get("document_type") or ""), "facts": []})
        for item in doc.get("facts") or []:
            if not isinstance(item, dict):
                continue
            name = str(item.get("label") or "Unlabelled field").strip()
            bucket["facts"].append(_fact(
                page_of(item), name, to_concept(str(item.get("concept") or name)),
                str(item.get("value") or "").strip(), str(item.get("evidence") or "").strip()[:500],
            ))

    references = []
    for ref in payload.get("cross_references") or []:
        if not isinstance(ref, dict):
            continue
        source = match(str(ref.get("from_document", "")))
        target = str(ref.get("to_document", "")).strip()
        if source is None or not target:
            continue
        references.append({
            "from_label": source,
            "page_number": page_of(ref),
            "label": f"Reference to {target}",
            "concept": f"cross_reference_{to_concept(str(ref.get('concept') or 'reference'))}",
            "value": str(ref.get("value") or "").strip(),
            "evidence": str(ref.get("evidence") or "").strip()[:500],
        })
    return by_label, references


def call_gemini_batch(batch: List[Dict[str, Any]], keys: List[Dict[str, str]]) -> Tuple[Any, str]:
    if genai is None or types is None:
        raise RuntimeError("google-genai is not installed (pip install google-genai).")
    if not keys:
        raise RuntimeError("No valid Gemini API keys are configured.")

    labels = [unit["label"] for unit in batch]
    errors: List[str] = []
    uploads: Dict[str, Tuple[Any, List[Any]]] = {}
    try:
        for model in GEMINI_MODELS:
            for entry in keys:
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
                for unit, handle in zip(batch, handles):
                    contents.append(f"=== DOCUMENT: {unit['label']} ===")
                    contents.append(handle)
                contents.append(BATCH_PROMPT)
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
                    if cooldown:
                        set_cooldown(f"{entry['id']}:{model}", cooldown)
                    continue

                if "MAX_TOKENS" in finish_reason(response):
                    raise OversizedBatch("response was truncated at the output token limit")
                try:
                    parsed = parse_batch_payload(response.text or "", labels)
                except ValueError as exc:  # json.JSONDecodeError is a ValueError
                    errors.append(f"{masked} · {model}: invalid JSON ({str(exc)[:120]})")
                    continue
                return parsed, f"{masked} · {model}"
        raise RuntimeError("; ".join(errors[-6:]) or "every key and model failed")
    finally:
        for client, handles in uploads.values():
            delete_handles(client, handles)


def unit_fallback(unit: Dict[str, Any], message: str) -> None:
    unit["items"] = local_extract(unit["pages_slice"])
    unit["doc_type"] = classify_local(unit["pages_slice"])
    unit["status"] = "fallback"
    unit["message"] = f"Local rules used: {message}"


def apply_batch_result(batch: List[Dict[str, Any]], parsed: Any, used: str) -> None:
    by_label, references = parsed
    units_by_label = {unit["label"]: unit for unit in batch}
    for unit in batch:
        payload = by_label.get(unit["label"])
        if payload is None:
            unit_fallback(unit, "model returned no entry for this part")
            continue
        unit["items"] = [
            dict(raw, page_number=unit["offset"] + min(max(1, raw["page_number"]), unit["pages"]))
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
        source["items"].append(dict(ref, page_number=source["offset"] + local_page))


def process_units(batch: List[Dict[str, Any]], keys: List[Dict[str, str]]) -> None:
    try:
        parsed, used = call_gemini_batch(batch, keys)
    except OversizedBatch as exc:
        if len(batch) > 1:
            middle = len(batch) // 2
            process_units(batch[:middle], keys)
            process_units(batch[middle:], keys)
            return
        unit_fallback(batch[0], f"part too large to process even alone ({str(exc)[:150]})")
        return
    except Exception as exc:  # noqa: BLE001
        for unit in batch:
            unit_fallback(unit, clean_text(exc)[:300])
        return
    apply_batch_result(batch, parsed, used)


# ---------------------------------------------------------------- ingestion
def prepare_document(docket_id: str, filename: str, data: bytes, mode: str) -> Tuple[Optional[Dict], Optional[Dict]]:
    digest = hashlib.sha256(data).hexdigest()
    with connect() as connection:
        duplicate = connection.execute(
            "SELECT filename FROM documents WHERE docket_id = ? AND sha256 = ?", (docket_id, digest)
        ).fetchone()
    if duplicate:
        return None, {"filename": filename, "status": "skipped", "facts": 0,
                      "message": f"Identical to “{duplicate['filename']}” already in this docket."}
    try:
        pages = read_pages(filename, data)
    except Exception as exc:  # noqa: BLE001
        return None, {"filename": filename, "status": "failed", "facts": 0,
                      "message": clean_text(exc)[:300]}

    doc_id = str(uuid.uuid4())
    now = utc_now()
    mime = MIME_TYPES.get(extension_of(filename), "application/octet-stream")
    storage_path = store_original(docket_id, doc_id, filename, data)
    try:
        with connect() as connection:
            connection.execute(
                """INSERT INTO documents (id, docket_id, filename, mime_type, sha256, size_bytes,
                   storage_path, page_count, document_type, status, extraction_mode, created_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (doc_id, docket_id, clean_text(filename), mime, digest, len(data), storage_path, len(pages),
                 classify_local(pages), "pending", mode, now),
            )
            connection.executemany(
                "INSERT INTO pages (document_id, page_number, text, readable) VALUES (?, ?, ?, ?)",
                [(doc_id, p["page_number"], p["text"], int(p["readable"])) for p in pages],
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
                      document_type: str, items: List[Dict[str, Any]]) -> None:
    """Save all facts for a document in one transaction: all of them are stored, or none."""
    now = utc_now()
    with connect() as connection:
        for item in items:
            page_number = min(max(1, item["page_number"]), doc["page_count"])
            value = clean_text(item["value"])
            evidence = clean_text(item["evidence"])
            assessment = assess(value, evidence, doc["page_text"].get(page_number))
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
        connection.execute("UPDATE documents SET status = ?, document_type = ? WHERE id = ?",
                           (status, clean_text(document_type or "Unclassified")[:80], doc["id"]))
        connection.execute("UPDATE dockets SET updated_at = ? WHERE id = ?", (now, docket_id))
    apply_extraction_integrity(docket_id)


def _fact_group_key(label: str, concept: str) -> str:
    return normalise(label) or normalise(concept) or "other"


def _append_reason(raw: str, reason: str) -> str:
    reasons = list(dict.fromkeys(json.loads(raw or "[]") + [reason]))
    return json.dumps(reasons)


def apply_extraction_integrity(docket_id: str) -> None:
    """Auto-approve repeated facts and flag contradictory or ill-typed values.

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
            if row["value"].strip():
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
                    reasons = _append_reason(row["confidence_reasons"], "LABEL_VALUE_CONFLICT")
                    connection.execute(
                        """UPDATE facts SET confidence_band = 'Red', confidence_score = 0,
                           confidence_reasons = ?, updated_at = ? WHERE id = ?""",
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
                        """UPDATE facts SET confidence_band = 'Green', confidence_score = 100,
                           confidence_reasons = ?, review_status = 'approved', updated_at = ?
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
                """UPDATE facts SET confidence_band = 'Red', confidence_score = 0,
                   confidence_reasons = ?, updated_at = ? WHERE id = ?""",
                (reasons, now, row["id"]),
            )
            changed = True
        if changed:
            connection.execute("UPDATE dockets SET updated_at = ? WHERE id = ?", (now, docket_id))


def gemini_document_status(doc: Dict[str, Any]) -> str:
    statuses = {unit["status"] for unit in doc["units"]}
    if statuses == {"gemini"}:
        return "gemini"
    if statuses == {"fallback"}:
        return "fallback"
    return "partial"


def process_uploads(docket_id: str, uploads: List[Any]) -> None:
    mode = st.session_state.get("extraction_mode", "Local")
    keys = key_pool() if mode == "Gemini" else []
    if mode == "Gemini" and not keys:
        queue_flash("error", "Gemini mode needs at least one valid API key. Add one in Settings, "
                             "or switch to Local mode.")
        return

    results: List[Dict[str, Any]] = []
    documents: List[Dict[str, Any]] = []
    finalized: set = set()

    def save(doc: Dict[str, Any], status: str, doc_type: str, items: List[Dict[str, Any]], message: str) -> None:
        try:
            finalize_document(docket_id, doc, status, doc_type, items)
            finalized.add(doc["id"])
            results.append({"filename": doc["filename"], "status": status, "facts": len(items), "message": message})
        except Exception as exc:  # noqa: BLE001
            LOG.exception("could not save %s", doc["filename"])
            delete_document(docket_id, doc["id"])  # leave nothing half-saved, so a retry is possible
            results.append({"filename": doc["filename"], "status": "failed", "facts": 0,
                            "message": f"Could not save results ({clean_text(exc)[:150]}). Nothing was kept; try again."})

    with st.status("Processing docket…", expanded=True) as run:
        try:
            for upload in uploads:
                data = upload.getvalue()
                problem = validate_upload(upload.name, data)
                if problem:
                    results.append({"filename": upload.name, "status": "failed", "facts": 0, "message": problem})
                    continue
                try:
                    doc, skipped = prepare_document(docket_id, upload.name, data, mode)
                except Exception as exc:  # noqa: BLE001
                    LOG.exception("prepare failed for %s", upload.name)
                    results.append({"filename": upload.name, "status": "failed", "facts": 0,
                                    "message": f"Could not store the file ({clean_text(exc)[:200]})."})
                    continue
                if doc:
                    documents.append(doc)
                else:
                    results.append(skipped)

            if mode == "Local":
                for doc in documents:
                    run.write(f"Local rules: {doc['filename']}")
                    items = local_extract(doc["pages"])
                    save(doc, "local", classify_local(doc["pages"]), items, "Extracted with local rules.")
            else:
                units: List[Dict[str, Any]] = []
                for doc in documents:
                    try:
                        units.extend(make_units(doc))
                    except Exception as exc:  # noqa: BLE001
                        doc["units"] = [{
                            "label": doc["filename"], "doc": doc, "offset": 0, "pages": doc["page_count"],
                            "data": doc["data"], "mime": doc["mime"], "suffix": "",
                            "pages_slice": doc["pages"], "items": None, "status": None,
                            "message": "", "doc_type": None,
                        }]
                        unit_fallback(doc["units"][0], f"could not split document ({str(exc)[:120]})")
                batches = pack_batches([u for u in units if u["status"] is None])
                run.write(f"{len(units)} part(s) packed into {len(batches)} request(s).")
                for index, batch in enumerate(batches, start=1):
                    run.write(f"Request {index}/{len(batches)}: {len(batch)} part(s)")
                    process_units(batch, keys)
                for doc in documents:
                    items = [item for unit in doc["units"] for item in (unit["items"] or [])]
                    doc_type = next((u["doc_type"] for u in doc["units"] if u["doc_type"]), "")
                    messages = sorted({u["message"] for u in doc["units"] if u["message"]})
                    save(doc, gemini_document_status(doc), doc_type or classify_local(doc["pages"]),
                         items, " | ".join(messages))
            run.update(label="Processing complete", state="complete", expanded=False)
        except Exception as exc:  # noqa: BLE001
            LOG.exception("processing run failed")
            for doc in documents:
                if doc["id"] not in finalized:
                    delete_document(docket_id, doc["id"])
                    results.append({"filename": doc["filename"], "status": "failed", "facts": 0,
                                    "message": f"Processing was interrupted ({clean_text(exc)[:150]}). "
                                               "The file was removed so you can retry."})
            run.update(label="Processing stopped with an error", state="error", expanded=True)

    st.session_state["last_results"] = results
    st.session_state["upload_nonce"] = st.session_state.get("upload_nonce", 0) + 1  # clears the queue
    good = sum(r["status"] not in ("failed", "skipped") for r in results)
    queue_flash("success" if good else "warning",
                f"{good} of {len(results)} file(s) processed. See “Latest processing run” for details.")


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
        if status == "approved":
            band, score, reasons = "Green", 100, ["HUMAN_APPROVED"]
        elif status == "rejected":
            band, score, reasons = "Red", 0, ["HUMAN_REJECTED"]
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
    try:
        return float(re.sub(r"[^\d.\-]", "", value))
    except ValueError:
        return None


def validation_report(docket_id: str) -> List[Dict[str, str]]:
    with connect() as connection:
        documents = connection.execute(
            "SELECT id, filename, status, storage_path, sha256 FROM documents WHERE docket_id = ? ORDER BY filename",
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
        if document["status"] in ("fallback", "partial"):
            add("Yellow", "Extraction", f"{document['filename']}: some parts used local rules after Gemini failed.")
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


def build_psv(docket_id: str) -> Tuple[bytes, int]:
    """Pipe-separated values: UTF-8 with BOM, CRLF rows, fields quoted when they contain
    '|', '"', or a line break. The output is parsed back and must equal the source rows."""
    facts = fetch_facts(docket_id)
    expected: List[List[str]] = [list(PSV_FIELDS)]
    for row in facts:
        record = dict(row)
        record["fact_id"] = row["id"]
        record["confidence_reasons"] = ", ".join(json.loads(row["confidence_reasons"]))
        record["corrected"] = "true" if row["corrected"] else "false"
        expected.append([safe_cell(record[field]) for field in PSV_FIELDS])

    buffer = io.StringIO(newline="")
    csv.writer(buffer, delimiter="|", quotechar='"', quoting=csv.QUOTE_MINIMAL,
               doublequote=True, lineterminator="\r\n").writerows(expected)
    text = buffer.getvalue()

    try:
        parsed = list(csv.reader(io.StringIO(text, newline=""), delimiter="|", quotechar='"',
                                 doublequote=True, strict=True))
    except csv.Error as exc:
        raise ExportError(f"PSV could not be read back: {exc}") from exc
    if any(len(r) != len(PSV_FIELDS) for r in parsed):
        raise ExportError("PSV integrity check failed: a row has the wrong number of columns.")
    if parsed != expected:
        raise ExportError("PSV integrity check failed: re-reading the file did not reproduce the data.")
    try:
        data = ("\ufeff" + text).encode("utf-8")
    except UnicodeEncodeError as exc:
        raise ExportError(f"PSV contains text that cannot be encoded as UTF-8: {exc}") from exc
    if data.decode("utf-8-sig") != text:
        raise ExportError("PSV integrity check failed: UTF-8 round trip mismatch.")
    return data, len(facts)


def build_json(docket: sqlite3.Row) -> Tuple[bytes, int]:
    docket_id = docket["id"]
    facts = []
    for row in fetch_facts(docket_id):
        facts.append({
            "id": row["id"], "document_id": row["document_id"], "filename": row["filename"],
            "page_number": row["page_number"], "label": row["label"], "concept": row["concept"],
            "value": row["value"], "original_value": row["original_value"],
            "source_evidence": row["source_evidence"],
            "confidence": {"band": row["confidence_band"], "score": row["confidence_score"],
                           "reasons": json.loads(row["confidence_reasons"])},
            "review_status": row["review_status"], "corrected": bool(row["corrected"]),
            "created_at": row["created_at"], "updated_at": row["updated_at"],
        })
    with connect() as connection:
        documents = [dict(r) for r in connection.execute(
            """SELECT id, filename, mime_type, sha256, size_bytes, page_count, document_type, status,
                      extraction_mode, created_at FROM documents WHERE docket_id = ? ORDER BY filename""",
            (docket_id,),
        ).fetchall()]
        history = [dict(r) for r in connection.execute(
            """SELECT h.* FROM fact_history h JOIN facts f ON f.id = h.fact_id
               WHERE f.docket_id = ? ORDER BY h.changed_at""",
            (docket_id,),
        ).fetchall()]
    payload = deep_clean({
        "schema_version": 1,
        "docket": {"id": docket["id"], "name": docket["name"],
                   "type": docket["docket_type"],
                   "created_at": docket["created_at"], "updated_at": docket["updated_at"]},
        "exported_at": utc_now(),
        "documents": documents,
        "facts": facts,
        "fact_history": history,
    })
    try:
        text = json.dumps(payload, indent=2, ensure_ascii=False, allow_nan=False)
        data = text.encode("utf-8")
        reloaded = json.loads(data.decode("utf-8"))
    except (TypeError, ValueError, UnicodeError) as exc:
        raise ExportError(f"JSON could not be generated safely: {exc}") from exc
    if reloaded != payload:
        raise ExportError("JSON integrity check failed: re-reading the file did not reproduce the data.")
    if len(reloaded["facts"]) != len(facts) or len(reloaded["documents"]) != len(documents):
        raise ExportError("JSON integrity check failed: record counts do not match.")
    return data, len(facts)
VIEW_HEADERS = ["filename", "label", "value", "confidence_level"]
XLSX_HEADERS = ["Filename", "Label", "Value", "Confidence Level"]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
BAND_FILLS = {"Green": "C6EFCE", "Yellow": "FFEB9C", "Red": "FFC7CE"}


def build_view_rows(docket_id: str) -> List[List[str]]:
    return [
        [clean_text(f["filename"]), clean_text(f["label"]), clean_text(f["value"]), f["confidence_band"]]
        for f in fetch_facts(docket_id)
    ]


def build_view_psv(docket_id: str) -> Tuple[bytes, int]:
    rows = build_view_rows(docket_id)
    expected = [VIEW_HEADERS] + [[safe_cell(cell) for cell in row] for row in rows]
    buffer = io.StringIO(newline="")
    csv.writer(buffer, delimiter="|", quotechar='"', quoting=csv.QUOTE_MINIMAL,
               doublequote=True, lineterminator="\r\n").writerows(expected)
    text = buffer.getvalue()
    try:
        parsed = list(csv.reader(io.StringIO(text, newline=""), delimiter="|", quotechar='"',
                                 doublequote=True, strict=True))
    except csv.Error as exc:
        raise ExportError(f"PSV could not be read back: {exc}") from exc
    if parsed != expected:
        raise ExportError("PSV integrity check failed: re-reading the file did not reproduce the data.")
    return ("\ufeff" + text).encode("utf-8"), len(rows)


def build_view_xlsx(docket_id: str) -> Tuple[bytes, int]:
    rows = build_view_rows(docket_id)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Extraction"
    sheet.append(XLSX_HEADERS)
    for cell in sheet[1]:
        cell.font = Font(bold=True)

    for row_number, row in enumerate(rows, start=2):
        for column_number, value in enumerate(row, start=1):
            cell = sheet.cell(row=row_number, column=column_number)
            cell.value = value
            if isinstance(value, str) and value[:1] in ("=", "+", "-", "@"):
                cell.data_type = "s"  # keep as text, never a formula
        colour = BAND_FILLS.get(row[3], "FFFFFF")
        sheet.cell(row=row_number, column=4).fill = PatternFill(
            fill_type="solid", start_color=colour, end_color=colour
        )

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:D{max(sheet.max_row, 1)}"
    for letter, width in zip("ABCD", (34, 30, 42, 18)):
        sheet.column_dimensions[letter].width = width

    output = io.BytesIO()
    workbook.save(output)
    data = output.getvalue()

    # Re-read the file and confirm it matches the database before offering it.
    reloaded = load_workbook(io.BytesIO(data), read_only=True)
    got = list(reloaded.active.iter_rows(min_row=2, max_col=4, values_only=True))
    reloaded.close()
    if len(got) != len(rows) or [r[3] for r in got] != [r[3] for r in rows]:
        raise ExportError("Excel integrity check failed: re-reading the file did not match the data.")
    return data, len(rows)

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
    return f'<span class="badge {esc(kind)}">{esc(icon)} {esc(text)}</span>'.replace("> ", ">", 1) if not icon else \
        f'<span class="badge {esc(kind)}">{esc(icon)} {esc(text)}</span>'


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
    mode = st.session_state.get("extraction_mode", "Local")
    keys = len(key_pool())
    cooling = sum(1 for until in st.session_state.get("key_cooldowns", {}).values() if until > time.time())
    rejected = len(rejected_keys())

    brand, docket_col, chips_col = st.columns([3, 3, 3], vertical_alignment="center")
    with brand:
        st.markdown(
            '<div class="brand"><div class="brand-mark" aria-hidden="true">P</div><div>'
            '<div class="brand-name">Parchi AI</div>'
            '<div class="brand-tag">Supplier document extraction &amp; verification</div></div></div>',
            unsafe_allow_html=True,
        )
    with docket_col:
        if dockets:
            names = {d["id"]: d["name"] for d in dockets}
            st.selectbox("Active docket", ids, key="active_docket",
                         format_func=lambda i: names[i], label_visibility="collapsed")
        else:
            st.button("＋ Create your first docket", type="primary", on_click=go_to, args=("Dockets",),
                      **stretch())
    with chips_col:
        extra = ""
        if cooling:
            extra += f'<span class="chip warn">▲ {cooling} key(s) cooling</span>'
        if rejected:
            extra += f'<span class="chip bad">✖ {rejected} invalid key(s)</span>'
        st.markdown(
            f'<div class="chips"><span class="chip accent">Mode: {esc(mode)}</span>'
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
def add_key_callback() -> None:
    raw = st.session_state.get("add_key_input", "").strip()
    problem = key_problem(raw)
    if problem:
        queue_flash("error", problem)
    elif any(e["key"] == raw for e in key_pool()):
        queue_flash("warning", "That key is already in the pool.")
    else:
        st.session_state.setdefault("session_keys", []).append(raw)
        queue_flash("success", "Key added for this session. It is not saved to disk.")
    st.session_state["add_key_input"] = ""


@guarded
def remove_key_callback(identifier: str) -> None:
    st.session_state["session_keys"] = [
        k for k in st.session_state.get("session_keys", []) if key_id(str(k).strip()) != identifier
    ]
    queue_flash("success", "Session key removed.")


@guarded
def clear_cooldowns_callback() -> None:
    st.session_state["key_cooldowns"] = {}
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
            "Type Of Docket",
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
                open_col, manage_col = st.columns(2)
                with open_col:
                    st.button("Open", key=f"open_{docket['id']}", type="primary",
                              on_click=open_docket, args=(docket["id"],), **stretch())
                with manage_col:
                    with st.popover("Manage", **stretch()):
                        rename_key = f"rename_{docket['id']}"
                        st.text_input("Rename", value=docket["name"], key=rename_key, max_chars=120)
                        st.button("Save name", key=f"save_name_{docket['id']}",
                                  on_click=rename_callback, args=(docket["id"], rename_key))
                        st.divider()
                        st.markdown(f"**Delete “{docket['name']}”**")
                        st.caption("Removes every document, extracted fact, history entry and stored file.")
                        confirm_key = f"confirm_delete_{docket['id']}"
                        st.checkbox("I understand this cannot be undone", key=confirm_key)
                        st.button("Delete docket permanently", key=f"delete_{docket['id']}", type="primary",
                                  disabled=not st.session_state.get(confirm_key, False),
                                  on_click=delete_docket_callback, args=(docket["id"], docket["name"]))


def page_workspace(dockets: List[sqlite3.Row], active: Optional[sqlite3.Row]) -> None:
    page_header("Workspace", "Upload supplier documents, run extraction and manage the stored originals.",
                active["name"] if active else None)
    if not need_docket(active):
        return
    docket_id = active["id"]
    facts = fetch_facts(docket_id)
    documents = list_documents(docket_id)

    metrics = st.columns(4)
    metrics[0].metric("Documents", len(documents))
    metrics[1].metric("Facts", len(facts))
    metrics[2].metric("Pending review", sum(f["review_status"] == "pending" for f in facts))
    metrics[3].metric("Red confidence", sum(f["confidence_band"] == "Red" for f in facts))

    with st.container(border=True):
        st.markdown("**Add documents**")
        st.radio("Extraction mode", ["Local", "Gemini"], horizontal=True, key="extraction_mode",
                 help="Local uses text-layer rules and works offline. Gemini sends each batch of "
                      "documents in one request and also reads scans and images.")
        uploads = st.file_uploader(
            f"Drop PDF, PNG, JPG or TXT files (up to {fmt_size(MAX_UPLOAD_BYTES)} each)",
            type=SUPPORTED_TYPES,
            accept_multiple_files=True,
            key=f"uploads_{st.session_state.get('upload_nonce', 0)}",
        )
        run = st.button(f"Process {len(uploads)} document(s)" if uploads else "Process documents",
                        type="primary", disabled=not uploads, **stretch())
        if not uploads:
            st.caption("Choose one or more files to enable processing.")
        if run and uploads:
            process_uploads(docket_id, uploads)
            st.rerun()

    results = st.session_state.get("last_results")
    if results:
        with st.expander("Latest processing run", expanded=True):
            for result in results:
                text = (f"**{result['filename']}** · {result['status']} · {result['facts']} record(s)"
                        + (f" · {result['message']}" if result["message"] else ""))
                if result["status"] == "failed":
                    st.error(text, icon="🚫")
                elif result["status"] in ("skipped", "fallback", "partial"):
                    st.warning(text, icon="⚠️")
                else:
                    st.success(text, icon="✅")
            st.button("Dismiss", key="dismiss_results", on_click=lambda: st.session_state.pop("last_results", None))

    st.markdown("#### Stored documents")
    if not documents:
        empty_state("No documents in this docket", "Upload files above. Originals are kept on disk until you remove them.")
        return

    st.dataframe(
        [
            {
                "Document": d["filename"],
                "Type": d["document_type"],
                "Pages": d["page_count"],
                "Size": fmt_size(d["size_bytes"]),
                "Extraction": d["extraction_mode"],
                "Status": d["status"],
                "Added": display_time(d["created_at"]),
                "File on disk": "✔" if resolve_stored_safe(d["storage_path"]) else "missing",
            }
            for d in documents
        ],
        hide_index=True,
        **stretch(),
    )

    st.markdown("#### File actions")
    by_id = {d["id"]: d for d in documents}
    chosen_id = st.selectbox("Document", list(by_id), format_func=lambda i: by_id[i]["filename"], key="file_pick")
    chosen = by_id[chosen_id]
    with st.container(border=True):
        info_col, download_col, remove_col = st.columns([3, 1, 2], vertical_alignment="center")
        with info_col:
            st.markdown(f"**{chosen['filename']}**")
            st.caption(f"{chosen['mime_type']} · {chosen['page_count']} page(s) · "
                       f"{fmt_size(chosen['size_bytes'])} · added {display_time(chosen['created_at'])}")
        data, problem = read_verified(chosen["storage_path"], chosen["sha256"])
        with download_col:
            if data is not None:
                st.download_button("Download", data, file_name=chosen["filename"],
                                   mime=chosen["mime_type"], **stretch())
            else:
                st.button("Unavailable", disabled=True, help=problem, **stretch())
        with remove_col:
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

    with st.expander("Docket actions"):
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
    band_filter = band_col.multiselect("Confidence", list(BANDS), default=list(BANDS), key="review_band")
    status_filter = status_col.selectbox("Review status", ["All"] + list(REVIEW_STATUSES), key="review_status")

    visible = [
        f for f in facts
        if (doc_filter == "All documents" or f["filename"] == doc_filter)
        and f["confidence_band"] in band_filter
        and (status_filter == "All" or f["review_status"] == status_filter)
    ]
    green_pending = sum(f["confidence_band"] == "Green" and f["review_status"] == "pending" for f in facts)
    summary_col, bulk_col = st.columns([3, 1], vertical_alignment="center")
    summary_col.caption(f"{len(visible)} of {len(facts)} facts shown.")
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


def page_psv_verify(dockets: List[sqlite3.Row], active: Optional[sqlite3.Row]) -> None:
    page_header("PSV verification", "Clean a pipe-separated export and verify low-confidence extractions first.",
                active["name"] if active else None)
    upload = st.file_uploader("Upload PSV file", type=["psv", "txt", "csv"], key="psv_upload")
    if upload is not None:
        try:
            payload = upload.getvalue()
            signature = hashlib.sha256(payload).hexdigest()
            frame = clean_psv(payload)
        except (UnicodeDecodeError, ValueError, pd.errors.ParserError) as exc:
            st.error(f"Could not process this PSV: {exc}", icon="🚫")
            return
        if st.session_state.get("psv_signature") != signature:
            st.session_state["psv_frame"] = frame
            st.session_state["psv_name"] = upload.name
            st.session_state["psv_signature"] = signature
            st.session_state["psv_verified"] = auto_verified_indices(frame)
            st.session_state["psv_current"] = -1

    frame = st.session_state.get("psv_frame")
    if frame is None:
        st.info("Upload a PSV containing field, value, source document, and confidence data.")
        return

    verified = st.session_state.setdefault("psv_verified", set())
    current = int(st.session_state.get("psv_current", -1))
    pending_count = sum(
        index not in verified and frame.at[index, "Confidence Level"] in {"Low", "Medium"}
        for index in frame.index
    )
    metrics = st.columns(4)
    metrics[0].metric("Rows", len(frame))
    metrics[1].metric("Low", int((frame["Confidence Level"] == "Low").sum()))
    metrics[2].metric("Medium", int((frame["Confidence Level"] == "Medium").sum()))
    metrics[3].metric("Pending review", pending_count)

    action_col, export_col = st.columns([1, 1])
    if action_col.button("Next pending review", type="primary", disabled=pending_count == 0, **stretch()):
        next_index = next_pending_index(frame, current, verified)
        if next_index is not None:
            st.session_state["psv_current"] = next_index
            st.rerun()
    excel_data = build_excel(frame)
    export_col.download_button(
        "Download formatted Excel",
        excel_data,
        file_name=f"{Path(st.session_state.get('psv_name', 'verification')).stem}-verification.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        **stretch(),
    )

    if not len(frame):
        st.success("The PSV contains no data rows.")
        return
    current = int(st.session_state.get("psv_current", 0))
    current = max(0, min(current, len(frame) - 1))
    st.session_state["psv_current"] = current
    row = frame.iloc[current]
    with st.container(border=True):
        st.markdown(f"**Review item {current + 1} of {len(frame)}** · {row['Confidence Level']}")
        st.markdown(f"**{row['Field Name']}**")
        st.write(row["Extracted Value"] or "(empty)")
        st.caption(f"Source document: {row['Source Document']}")
        is_high = row["Confidence Level"] == "High"
        is_verified = current in verified
        if not is_high and not is_verified:
            if st.button("Mark verified", type="primary", key=f"psv_verify_{current}", **stretch()):
                verified.add(current)
                st.session_state["psv_verified"] = verified
                st.session_state["psv_current"] = next_pending_index(frame, current, verified)
                if st.session_state["psv_current"] is None:
                    st.session_state["psv_current"] = current
                st.rerun()
        else:
            st.success("Pre-verified" if is_high else "Verified", icon="✅")

    color_map = {
        "High": "background-color: #1f7a4d; color: #f0fff7; font-weight: 700",
        "Medium": "background-color: #8a6d1d; color: #fff8d6; font-weight: 700",
        "Low": "background-color: #8c3540; color: #fff1f2; font-weight: 700",
    }
    styled_frame = frame.style.map(
        lambda value: color_map.get(str(value), ""),
        subset=["Confidence Level"],
    )
    st.dataframe(styled_frame, hide_index=True, width="stretch")


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
    page_header("Export", "Download the review view as PSV or Excel, or the full structured JSON. "
                          "Each file is re-read and checked before it can be downloaded.",
                active["name"] if active else None)
    if not need_docket(active):
        return
    if active["fact_count"] == 0:
        empty_state("Nothing to export yet", "Process documents in the Workspace to extract facts first.")
        return
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d")
    slug = to_concept(active["name"]) or "docket"

    # (title, caption, builder, filename, mime, show_text_preview)
    specs = [
        ("PSV (review view)",
         "Four columns: filename, label, value, confidence_level. Pipe-delimited, UTF-8 with BOM. "
         "Can be uploaded straight into PSV Verify.",
         lambda: build_view_psv(active["id"]), f"parchai-{slug}-{stamp}.psv",
         "text/plain; charset=utf-8", True),
        ("Excel (review view)",
         "The same four columns in a spreadsheet. Confidence Level cells are green, yellow or red.",
         lambda: build_view_xlsx(active["id"]), f"parchai-{slug}-{stamp}.xlsx", XLSX_MIME, False),
        ("JSON (full)",
         "Docket, documents, facts with confidence, review state and full change history.",
         lambda: build_json(active), f"parchai-{slug}-{stamp}.json", "application/json", True),
    ]
    for column, (title, caption, builder, filename, mime, show_preview) in zip(st.columns(3), specs):
        with column, st.container(border=True):
            st.markdown(f"**{title}**")
            st.caption(caption)
            try:
                with st.spinner(f"Building and verifying {title}…"):
                    data, count = builder()
            except ExportError as exc:
                st.error(f"{title} export blocked: {exc}", icon="🚫")
                continue
            except Exception as exc:  # noqa: BLE001
                LOG.exception("export failed")
                st.error(f"{title} export failed: {clean_text(exc)[:200]}", icon="🚫")
                continue
            st.success(f"Integrity verified · {count} row(s) · {fmt_size(len(data))}", icon="✅")
            st.caption(f"SHA-256 `{hashlib.sha256(data).hexdigest()[:16]}…`")
            st.download_button(f"Download {title.split(' (')[0]}", data, file_name=filename,
                               mime=mime, key=f"dl_{title}", **stretch())
            if show_preview:
                with st.expander("Preview"):
                    st.code(data.decode("utf-8-sig")[:1500], language=None)


def page_settings(dockets: List[sqlite3.Row], active: Optional[sqlite3.Row]) -> None:
    page_header("Settings", "Gemini key pool, model fallback, storage and integrity checks.", None)
    entries = key_pool()

    st.markdown("#### Gemini API key pool")
    st.caption("Keys are masked. Env keys come from GEMINI_API_KEYS or GEMINI_API_KEY. Session keys live "
               "in memory only. A quota error pauses a key for one model for a minute; an invalid key "
               "is paused for an hour.")
    for bad in rejected_keys():
        st.error(f"A key from “{bad['source']}” ending {bad['masked']} was ignored because it is not a valid "
                 "API key (it must be a complete ASCII key beginning with “AQ.” or “AIza”, with no spaces). "
                 "Check the GEMINI_API_KEY(S) value.", icon="🚫")
    if not entries:
        st.warning("No valid keys configured. Local extraction still works without keys.", icon="⚠️")
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

    session_entries = [e for e in entries if e["source"] == "session"]
    if session_entries:
        st.markdown("**Remove session keys**")
        for entry in session_entries:
            col_key, col_btn = st.columns([4, 1], vertical_alignment="center")
            col_key.code(entry["masked"], language=None)
            col_btn.button("Remove", key=f"remove_{entry['id']}", on_click=remove_key_callback,
                           args=(entry["id"],))

    with st.container(border=True):
        st.markdown("**Add a key for this session**")
        key_col, add_col = st.columns([4, 1], vertical_alignment="bottom")
        key_col.text_input("Gemini API key", type="password", key="add_key_input",
                           help="Paste only the complete key. Current Gemini keys start with “AQ.”.")
        add_col.button("Add key", on_click=add_key_callback, **stretch())

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
    st.caption("Override with PARCHAI_GEMINI_MODELS (comma-separated). Use “List models” above to see valid names.")

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
    "PSV Verify": page_psv_verify,
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