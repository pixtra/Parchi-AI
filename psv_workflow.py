"""PSV cleaning, Excel formatting, and manual verification workflow helpers."""
from __future__ import annotations

import io
import math
import re
from typing import Iterable

import pandas as pd
from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


OUTPUT_COLUMNS = ["Field Name", "Extracted Value", "Source Document", "Confidence Level"]
CONFIDENCE_ORDER = ["Low", "Medium", "High"]
ALIASES = {
    "field": ("field name", "field_name", "label", "attribute", "concept"),
    "value": ("extracted value", "extracted_value", "value", "text"),
    "source": ("source document", "source_document", "filename", "file name", "document"),
    "score": ("confidence score", "confidence_score", "score", "confidence"),
    "level": ("confidence level", "confidence_level", "confidence band", "confidence_band"),
}


def _normalise_column(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower()).strip()


def _find_column(columns: Iterable[object], aliases: Iterable[str]) -> str | None:
    normalised = {_normalise_column(column): str(column) for column in columns}
    for alias in aliases:
        if _normalise_column(alias) in normalised:
            return normalised[_normalise_column(alias)]
    return None


def _score(value: object) -> float | None:
    if value is None or str(value).strip() == "":
        return None
    text = str(value).strip().replace("%", "")
    try:
        number = float(text)
    except ValueError:
        return None
    if not math.isfinite(number):  # "inf" and "nan" are not usable scores
        return None
    # Only a decimal such as 0.95 is treated as a fraction. A bare "1" means 1 out of 100.
    if "." in text and 0 <= number <= 1:
        return number * 100
    return number


def confidence_level(score: object = None, level: object = None) -> str:
    # Thresholds mirror app.band_for (Green >= 90, Yellow >= 60, else Red) exactly, so the
    # same confidence score is never shown as a different band on different screens.
    numeric = _score(score)
    if numeric is not None:
        if numeric >= 90:
            return "High"
        if numeric >= 60:
            return "Medium"
        return "Low"
    text = str(level or "").strip().lower()
    if text in {"high", "green"}:
        return "High"
    if text in {"medium", "yellow", "moderate"}:
        return "Medium"
    return "Low"


def _normalise_fact(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").lower()).strip()


_MONTH = r"(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*"
_DATE_PATTERNS = (
    re.compile(r"\b\d{1,4}[/.-]\d{1,2}[/.-]\d{1,4}\b"),                       # 12/03/2026, 2026-03-12
    re.compile(rf"\b\d{{1,2}}[\s./-]*{_MONTH}[\s./,-]*\d{{2,4}}\b", re.IGNORECASE),  # 12 Jan 2026
    re.compile(rf"\b{_MONTH}[\s./-]*\d{{1,2}}(?:st|nd|rd|th)?[\s.,/-]*\d{{2,4}}\b", re.IGNORECASE),  # Jan 12, 2026
)
_DATE_LABEL_TOKENS = {"date", "dob"}
_AMOUNT_LABEL_TOKENS = {"amount", "price", "total", "tax", "rate", "quantity", "qty", "weight"}


def label_value_mismatch(label: object, value: object) -> bool:
    """Detect clear label/value type mismatches without guessing arbitrary field semantics.

    Labels are matched on whole words, so "Corporate" does not count as "rate" and
    "Update" does not count as "date".
    """
    label_text = _normalise_fact(label)
    tokens = set(label_text.split())
    value_text = str(value or "").strip()
    if not value_text:
        return True
    if tokens & _DATE_LABEL_TOKENS:
        return not any(pattern.search(value_text) for pattern in _DATE_PATTERNS)
    if tokens & _AMOUNT_LABEL_TOKENS:
        return not bool(re.search(r"\d", value_text))
    if "gstin" in tokens or "gst number" in label_text:
        return not bool(re.fullmatch(r"\d{2}[A-Z]{5}\d{4}[A-Z][A-Z0-9]Z[A-Z0-9]", value_text.upper()))
    if "email" in tokens:
        return not bool(re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", value_text))
    if tokens & {"phone", "mobile", "telephone"}:
        return len(re.sub(r"\D", "", value_text)) < 7
    return False


def auto_verified_indices(frame: pd.DataFrame) -> set[int]:
    """Rows duplicated by field/value are trusted for the review queue."""
    keys = frame.apply(
        lambda row: (_normalise_fact(row["Field Name"]), _normalise_fact(row["Extracted Value"])),
        axis=1,
    )
    label_values = frame.groupby(frame["Field Name"].map(_normalise_fact))["Extracted Value"].agg(
        lambda values: {_normalise_fact(value) for value in values if _normalise_fact(value)}
    )
    conflicting_labels = {label for label, values in label_values.items() if len(values) > 1}
    counts = keys.value_counts()
    return {
        index for index, key in keys.items()
        if key[1] and counts[key] > 1 and key[0] not in conflicting_labels
    }


def clean_psv(data: bytes) -> pd.DataFrame:
    """Read a PSV payload and return the required four-column review dataset."""
    text = data.decode("utf-8-sig")
    frame = pd.read_csv(io.StringIO(text), sep="|", dtype=str, keep_default_na=False)
    field_column = _find_column(frame.columns, ALIASES["field"])
    value_column = _find_column(frame.columns, ALIASES["value"])
    source_column = _find_column(frame.columns, ALIASES["source"])
    score_column = _find_column(frame.columns, ALIASES["score"])
    level_column = _find_column(frame.columns, ALIASES["level"])
    missing = [
        name for name, column in (
            ("Field Name", field_column),
            ("Extracted Value", value_column),
            ("Source Document", source_column),
        ) if column is None
    ]
    if missing:
        raise ValueError(f"Missing required PSV column(s): {', '.join(missing)}.")
    if score_column is None and level_column is None:
        raise ValueError("The PSV must include a confidence score or confidence level column.")

    result = pd.DataFrame({
        "Field Name": frame[field_column].astype(str),
        "Extracted Value": frame[value_column].astype(str),
        "Source Document": frame[source_column].astype(str),
        "Confidence Level": [
            confidence_level(
                frame.iloc[index][score_column] if score_column else None,
                frame.iloc[index][level_column] if level_column else None,
            )
            for index in range(len(frame))
        ],
    })
    mismatches = [
        label_value_mismatch(result.iloc[index]["Field Name"], result.iloc[index]["Extracted Value"])
        for index in range(len(result))
    ]
    result.loc[mismatches, "Confidence Level"] = "Low"
    label_values = result.groupby(result["Field Name"].map(_normalise_fact))["Extracted Value"].agg(
        lambda values: {_normalise_fact(value) for value in values if _normalise_fact(value)}
    )
    conflicting_labels = {label for label, values in label_values.items() if len(values) > 1}
    conflict_rows = result["Field Name"].map(_normalise_fact).isin(conflicting_labels)
    result.loc[conflict_rows, "Confidence Level"] = "Low"
    result["_sort"] = result["Confidence Level"].map({name: index for index, name in enumerate(CONFIDENCE_ORDER)})
    result["_review_index"] = range(len(result))
    result = result.sort_values(["_sort", "_review_index"], kind="stable")
    return result[OUTPUT_COLUMNS].reset_index(drop=True)


def build_excel(frame: pd.DataFrame) -> bytes:
    """Create a formatted Excel Table with confidence conditional formatting."""
    output = io.BytesIO()
    export = frame[OUTPUT_COLUMNS].copy()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export.to_excel(writer, index=False, sheet_name="PSV Verification")
    workbook = load_workbook(output)
    sheet = workbook["PSV Verification"]
    # Neutralise formula/CSV injection: a value such as =HYPERLINK(...) or +CMD(...) coming
    # from an uploaded PSV must never become a live formula in the generated workbook.
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str) and cell.value[:1] in ("=", "+", "-", "@"):
                cell.data_type = "s"
    sheet.freeze_panes = "A2"
    # The Table provides its own filter buttons. A separate sheet-level autofilter over the
    # same range makes Excel report the file as damaged, so it is not added.
    table = Table(displayName="PSVVerificationTable", ref=sheet.dimensions)
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False,
                                          showLastColumn=False, showRowStripes=True,
                                          showColumnStripes=False)
    sheet.add_table(table)
    fills = {
        "High": PatternFill("solid", fgColor="C6EFCE"),
        "Medium": PatternFill("solid", fgColor="FFEB9C"),
        "Low": PatternFill("solid", fgColor="FFC7CE"),
    }
    confidence_column = "D"
    for level, fill in fills.items():
        sheet.conditional_formatting.add(
            f"{confidence_column}2:{confidence_column}{max(sheet.max_row, 2)}",
            FormulaRule(formula=[f'${confidence_column}2="{level}"'], fill=fill),
        )
    widths = {"A": 28, "B": 42, "C": 34, "D": 20}
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    workbook.save(output)
    return output.getvalue()


def next_pending_index(
    frame: pd.DataFrame, current: int = -1, verified: set[int] | None = None
) -> int | None:
    """Return the next pending Low/Medium row, wrapping once at the end."""
    verified = verified or set()
    pending = [
        index for index in frame.index
        if index not in verified and frame.at[index, "Confidence Level"] in {"Low", "Medium"}
    ]
    if not pending:
        return None
    after = [index for index in pending if index > current]
    return (after or pending)[0]