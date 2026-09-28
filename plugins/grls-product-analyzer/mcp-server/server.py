#!/usr/bin/env python3
"""Local MCP server for product-list and GRLS analysis.

The server deliberately does not infer INNs or perform fuzzy pharmaceutical
matching. Current-market reports read only the GRLS section whose raw status is
``Действующий``, deduplicate exact registration numbers, and render offline HTML.
"""

from __future__ import annotations

import csv
import hashlib
import html
import io
import json
import os
import posixpath
import re
import shutil
import subprocess
import sys
import zipfile
from collections import Counter
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Iterable, Iterator
from xml.etree import ElementTree as ET


SERVER_NAME = "grls-product-analyzer"
SERVER_VERSION = "0.6.3"
PROTOCOL_VERSION = "2025-06-18"
PLUGIN_ROOT = Path(__file__).resolve().parents[1]
BUNDLED_GRLS_PATH = PLUGIN_ROOT / "data" / "grls-default.zip"
BUNDLED_GRLS_METADATA_PATH = PLUGIN_ROOT / "data" / "grls-default.json"
GRLS_STALE_AFTER_DAYS = 30
ARCHIVE_DATE_RE = re.compile(r"(?<!\d)(20\d{2})[-_.](\d{2})[-_.](\d{2})(?!\d)")

ACTIVE_STATUSES = ["Действующий"]

GRLS_FIELDS = [
    "registration_number",
    "registration_date",
    "expiration_date",
    "cancellation_date",
    "holder",
    "holder_country",
    "trade_name",
    "inn",
    "release_forms",
    "manufacturing_stages",
    "normative_documentation",
    "pharmacotherapeutic_group",
    "ved",
    "controlled_substances",
    "orphan",
    "source_status",
    "source_file",
    "source_row",
]

GRLS_COLUMN_INDEX = {
    "registration_number": 2,
    "registration_date": 3,
    "expiration_date": 4,
    "cancellation_date": 5,
    "holder": 6,
    "holder_country": 7,
    "trade_name": 8,
    "inn": 9,
    "release_forms": 10,
    "manufacturing_stages": 11,
    "normative_documentation": 12,
    "pharmacotherapeutic_group": 13,
    "ved": 14,
    "controlled_substances": 15,
    "orphan": 16,
}

XML_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XML_DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CELL_REF_RE = re.compile(r"([A-Z]+)(\d+)")
DIMENSION_RE = re.compile(br"<dimension\s+ref=\"([^\"]+)\"")


TOOLS = [
    {
        "name": "get_grls_status",
        "description": (
            "Return the active GRLS database, its source (bundled or user update), archive date, "
            "age, checksum, and whether a monthly update should be requested. Call this before "
            "each analysis."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {},
            "additionalProperties": False,
        },
        "annotations": {"readOnlyHint": True},
    },
    {
        "name": "update_grls_archive",
        "description": (
            "Validate a user-provided GRLS ZIP and install it as the active local database. "
            "The bundled archive remains available as a fallback. This is a write action and "
            "must only be called after the user supplies the archive and asks to update it."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "archive_path": {"type": "string", "description": "Path to the uploaded GRLS ZIP."},
                "archive_as_of": {
                    "type": "string",
                    "description": "Optional YYYY-MM-DD archive date if it cannot be inferred from the filename.",
                },
            },
            "required": ["archive_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": True,
        },
    },
    {
        "name": "inspect_product_list",
        "description": (
            "Read raw rows or page lines from a local .xlsb, .xlsx, .csv, or .pdf product list. "
            "Returns source cells without inferring INNs or interpreting pharmaceutical meaning."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Absolute path to the product list."},
                "sheet_name": {"type": "string", "description": "Optional worksheet name."},
                "start_row": {"type": "integer", "minimum": 1, "default": 1},
                "max_rows": {"type": "integer", "minimum": 1, "maximum": 500, "default": 50},
                "max_columns": {"type": "integer", "minimum": 1, "maximum": 100, "default": 20},
                "start_page": {"type": "integer", "minimum": 1, "default": 1},
                "max_pages": {"type": "integer", "minimum": 1, "maximum": 500, "default": 50},
            },
            "required": ["file_path"],
            "additionalProperties": False,
        },
    },
    {
        "name": "inspect_grls_archive",
        "description": (
            "List XLSX sections inside a local GRLS ZIP and return their raw status, headers, "
            "and approximate row counts. No records are matched or deduplicated."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "archive_path": {
                    "type": "string",
                    "description": "Optional GRLS ZIP. Omit to use the active bundled or updated database.",
                }
            },
            "additionalProperties": False,
        },
    },
    {
        "name": "search_grls",
        "description": (
            "Perform literal exact or substring retrieval over selected raw GRLS fields. "
            "Use it to gather candidates for model review; it does not decide equivalence, "
            "normalize related molecules, or remove duplicates."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "archive_path": {"type": "string"},
                "queries": {"type": "array", "items": {"type": "string"}, "minItems": 1, "maxItems": 30},
                "statuses": {"type": "array", "items": {"type": "string"}},
                "fields": {
                    "type": "array",
                    "items": {"type": "string", "enum": GRLS_FIELDS},
                    "default": ["inn"],
                },
                "match_mode": {"type": "string", "enum": ["contains", "exact"], "default": "contains"},
                "require_all_queries": {"type": "boolean", "default": False},
                "case_sensitive": {"type": "boolean", "default": False},
                "cursor": {"type": "integer", "minimum": 0, "default": 0},
                "max_results": {"type": "integer", "minimum": 1, "maximum": 1000, "default": 200},
            },
            "required": ["queries"],
            "additionalProperties": False,
        },
    },
    {
        "name": "distinct_grls_values",
        "description": (
            "Return distinct raw values and source-row counts for one GRLS field after literal filtering. "
            "Useful for reviewing candidate INN spellings before the model makes a decision."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "archive_path": {"type": "string"},
                "field": {"type": "string", "enum": GRLS_FIELDS},
                "contains_terms": {"type": "array", "items": {"type": "string"}, "maxItems": 30},
                "statuses": {"type": "array", "items": {"type": "string"}},
                "case_sensitive": {"type": "boolean", "default": False},
                "max_values": {"type": "integer", "minimum": 1, "maximum": 1000, "default": 200},
            },
            "required": ["field"],
            "additionalProperties": False,
        },
    },
    {
        "name": "build_html_report",
        "description": (
            "Build a standalone sortable HTML report from explicit, model-reviewed mapping decisions. "
            "The report always uses only the GRLS section named Действующий and deduplicates exact "
            "registration numbers. It separates finished medicines from pharmaceutical substances "
            "and counts requested-form competitors and their unique RU holders. It never invents INN mappings."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "archive_path": {"type": "string"},
                "output_path": {"type": "string", "description": "Absolute .html output path."},
                "decisions": {
                    "type": "array",
                    "minItems": 1,
                    "items": {
                        "type": "object",
                        "properties": {
                            "source_row": {"type": ["integer", "string"]},
                            "serial_number": {"type": ["integer", "string"]},
                            "product": {"type": "string"},
                            "dosage": {"type": "string"},
                            "packing": {"type": "string"},
                            "segment": {"type": "string"},
                            "normalized_inn": {"type": "string"},
                            "accepted_registry_inns": {"type": "array", "items": {"type": "string"}},
                            "desired_form": {"type": "string"},
                            "form_terms": {"type": "array", "items": {"type": "string"}},
                            "match_note": {"type": "string"},
                            "confidence": {"type": "string"},
                            "decision_status": {"type": "string", "enum": ["matched", "not_found", "manual_review"]},
                        },
                        "required": ["product", "normalized_inn", "accepted_registry_inns", "decision_status"],
                        "additionalProperties": False,
                    },
                },
                "title": {"type": "string", "default": "Product list × ГРЛС"},
                "product_list_name": {"type": "string"},
                "grls_as_of": {"type": "string"},
            },
            "required": ["output_path", "decisions"],
            "additionalProperties": False,
        },
    },
    {
        "name": "validate_html_report",
        "description": "Validate a generated report's embedded JSON, expected controls, and summary counts.",
        "inputSchema": {
            "type": "object",
            "properties": {"report_path": {"type": "string"}},
            "required": ["report_path"],
            "additionalProperties": False,
        },
    },
]


def _require_file(path_value: str, suffixes: tuple[str, ...] | None = None) -> Path:
    path = Path(path_value).expanduser().resolve()
    if not path.is_file():
        raise ValueError(f"File does not exist: {path}")
    if suffixes and path.suffix.casefold() not in suffixes:
        raise ValueError(f"Unsupported file type {path.suffix}; expected {', '.join(suffixes)}")
    return path


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _read_json(path: Path) -> dict[str, Any]:
    if not path.is_file():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return payload if isinstance(payload, dict) else {}


def _user_data_dir() -> Path:
    explicit = os.environ.get("GRLS_ANALYZER_DATA_DIR") or os.environ.get("PLUGIN_DATA")
    if explicit:
        return Path(explicit).expanduser().resolve()
    if os.name == "nt":
        base = os.environ.get("LOCALAPPDATA")
        if base:
            return Path(base).resolve() / SERVER_NAME
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / SERVER_NAME
    base = os.environ.get("XDG_DATA_HOME")
    if base:
        return Path(base).expanduser().resolve() / SERVER_NAME
    return Path.home() / ".local" / "share" / SERVER_NAME


def _infer_archive_date(filename: str, explicit: str | None = None) -> str:
    if explicit:
        try:
            return date.fromisoformat(explicit).isoformat()
        except ValueError as error:
            raise ValueError("archive_as_of must use YYYY-MM-DD") from error
    match = ARCHIVE_DATE_RE.search(filename)
    if match:
        try:
            return date(int(match.group(1)), int(match.group(2)), int(match.group(3))).isoformat()
        except ValueError:
            pass
    return ""


def _active_grls() -> tuple[Path, str, dict[str, Any]]:
    data_dir = _user_data_dir()
    updated_path = data_dir / "grls-current.zip"
    updated_metadata_path = data_dir / "grls-current.json"
    if updated_path.is_file():
        return updated_path, "user_update", _read_json(updated_metadata_path)
    if not BUNDLED_GRLS_PATH.is_file():
        raise RuntimeError(
            "Bundled GRLS archive is missing. Reinstall the plugin or provide archive_path explicitly."
        )
    return BUNDLED_GRLS_PATH, "bundled", _read_json(BUNDLED_GRLS_METADATA_PATH)


def get_grls_status(arguments: dict[str, Any]) -> dict[str, Any]:
    del arguments
    archive_path, source, metadata = _active_grls()
    archive_as_of = _clean(metadata.get("archive_as_of")) or _infer_archive_date(
        _clean(metadata.get("original_filename")) or archive_path.name
    )
    if not archive_as_of:
        archive_as_of = datetime.fromtimestamp(archive_path.stat().st_mtime).date().isoformat()
    archive_date = date.fromisoformat(archive_as_of)
    age_days = (date.today() - archive_date).days
    next_update_due = archive_date + timedelta(days=GRLS_STALE_AFTER_DAYS)
    checksum = _clean(metadata.get("sha256")) or _sha256(archive_path)
    return {
        "active_archive_path": str(archive_path),
        "source": source,
        "archive_as_of": archive_as_of,
        "age_days": age_days,
        "stale_after_days": GRLS_STALE_AFTER_DAYS,
        "next_update_due": next_update_due.isoformat(),
        "requires_update": age_days >= GRLS_STALE_AFTER_DAYS,
        "sha256": checksum,
        "original_filename": _clean(metadata.get("original_filename")) or archive_path.name,
        "installed_at_utc": _clean(metadata.get("installed_at_utc")),
        "bundled_fallback_path": str(BUNDLED_GRLS_PATH),
        "reminder": (
            "Попросите пользователя прикрепить свежий ZIP-архив ГРЛС и подтвердить обновление."
            if age_days >= GRLS_STALE_AFTER_DAYS
            else "Обновление пока не требуется."
        ),
    }


def _resolve_grls_archive(arguments: dict[str, Any]) -> Path:
    explicit = _clean(arguments.get("archive_path"))
    if explicit:
        return _require_file(explicit, (".zip",))
    path, _, _ = _active_grls()
    return _require_file(str(path), (".zip",))


def _clean(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace("_x000D_", "").strip()


def _column_number(reference: str) -> int:
    match = CELL_REF_RE.fullmatch(reference)
    if not match:
        return 0
    number = 0
    for character in match.group(1):
        number = number * 26 + ord(character) - 64
    return number


def _parse_dimension(reference: str | None) -> tuple[int | None, int | None]:
    if not reference:
        return None, None
    last = reference.split(":")[-1]
    match = CELL_REF_RE.fullmatch(last)
    if not match:
        return None, None
    return int(match.group(2)), _column_number(last)


class XlsxReader:
    def __init__(self, source: Path | bytes):
        stream: Any = source if isinstance(source, Path) else io.BytesIO(source)
        self.archive = zipfile.ZipFile(stream)
        self.sheets = self._load_sheets()
        self.shared_strings = self._load_shared_strings()

    def close(self) -> None:
        self.archive.close()

    def __enter__(self) -> "XlsxReader":
        return self

    def __exit__(self, exc_type: Any, exc: Any, traceback: Any) -> None:
        self.close()

    def _load_sheets(self) -> list[dict[str, str]]:
        workbook = ET.fromstring(self.archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(self.archive.read("xl/_rels/workbook.xml.rels"))
        rel_map = {
            relation.attrib["Id"]: relation.attrib["Target"]
            for relation in relationships.findall(f"{{{XML_PKG_REL}}}Relationship")
        }
        sheets = []
        for sheet in workbook.findall(f".//{{{XML_MAIN}}}sheet"):
            relation_id = sheet.attrib.get(f"{{{XML_DOC_REL}}}id", "")
            target = rel_map.get(relation_id, "")
            if target.startswith("/"):
                path = target.lstrip("/")
            else:
                path = posixpath.normpath(posixpath.join("xl", target))
            sheets.append({"name": sheet.attrib.get("name", ""), "path": path})
        return sheets

    def _load_shared_strings(self) -> list[str]:
        if "xl/sharedStrings.xml" not in self.archive.namelist():
            return []
        root = ET.fromstring(self.archive.read("xl/sharedStrings.xml"))
        values = []
        for item in root.findall(f"{{{XML_MAIN}}}si"):
            values.append("".join(text.text or "" for text in item.iter(f"{{{XML_MAIN}}}t")))
        return values

    def sheet(self, sheet_name: str | None = None) -> dict[str, str]:
        if not self.sheets:
            raise ValueError("Workbook contains no worksheets")
        if not sheet_name:
            return self.sheets[0]
        for sheet in self.sheets:
            if sheet["name"] == sheet_name:
                return sheet
        raise ValueError(f"Worksheet not found: {sheet_name}")

    def dimension(self, sheet_path: str) -> tuple[int | None, int | None]:
        with self.archive.open(sheet_path) as source:
            prefix = source.read(16384)
        match = DIMENSION_RE.search(prefix)
        return _parse_dimension(match.group(1).decode("utf-8") if match else None)

    def rows(
        self,
        sheet_path: str,
        start_row: int = 1,
        max_rows: int | None = None,
        max_columns: int = 50,
    ) -> Iterator[tuple[int, list[str]]]:
        emitted = 0
        with self.archive.open(sheet_path) as source:
            for event, element in ET.iterparse(source, events=("end",)):
                if element.tag != f"{{{XML_MAIN}}}row":
                    continue
                row_number = int(element.attrib.get("r", "0") or 0)
                if row_number < start_row:
                    element.clear()
                    continue
                values = [""] * max_columns
                for cell in element.findall(f"{{{XML_MAIN}}}c"):
                    reference = cell.attrib.get("r", "")
                    column = _column_number(reference)
                    if column < 1 or column > max_columns:
                        continue
                    cell_type = cell.attrib.get("t", "")
                    value_element = cell.find(f"{{{XML_MAIN}}}v")
                    value = value_element.text if value_element is not None else ""
                    if cell_type == "s" and value:
                        index = int(value)
                        value = self.shared_strings[index] if 0 <= index < len(self.shared_strings) else value
                    elif cell_type == "inlineStr":
                        inline = cell.find(f"{{{XML_MAIN}}}is")
                        value = "" if inline is None else "".join(
                            text.text or "" for text in inline.iter(f"{{{XML_MAIN}}}t")
                        )
                    elif cell_type == "b":
                        value = "TRUE" if value == "1" else "FALSE"
                    values[column - 1] = _clean(value)
                yield row_number, values
                emitted += 1
                element.clear()
                if max_rows is not None and emitted >= max_rows:
                    break


def _read_xlsb_via_powershell(
    path: Path, sheet_name: str | None, start_row: int, max_rows: int, max_columns: int
) -> dict[str, Any]:
    script = Path(__file__).with_name("read_excel.ps1")
    powershell = "powershell.exe" if os.name == "nt" else "pwsh"
    command = [
        powershell,
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script),
        "-Path",
        str(path),
        "-StartRow",
        str(start_row),
        "-MaxRows",
        str(max_rows),
        "-MaxColumns",
        str(max_columns),
    ]
    if sheet_name:
        command.extend(["-SheetName", sheet_name])
    completed = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", timeout=120)
    if completed.returncode != 0:
        error = completed.stderr.strip() or completed.stdout.strip()
        raise RuntimeError(
            "Could not read XLSB. Install pyxlsb or use Windows with Microsoft Excel. " + error
        )
    return json.loads(completed.stdout)


def _read_xlsb(
    path: Path, sheet_name: str | None, start_row: int, max_rows: int, max_columns: int
) -> dict[str, Any]:
    try:
        from pyxlsb import open_workbook  # type: ignore
    except ImportError:
        return _read_xlsb_via_powershell(path, sheet_name, start_row, max_rows, max_columns)

    with open_workbook(str(path)) as workbook:
        sheet_names = list(workbook.sheets)
        selected = sheet_name or sheet_names[0]
        if selected not in sheet_names:
            raise ValueError(f"Worksheet not found: {selected}")
        rows: list[dict[str, Any]] = []
        last_row = 0
        last_column = 0
        with workbook.get_sheet(selected) as sheet:
            for row_index, row in enumerate(sheet.rows(), start=1):
                values = [_clean(cell.v) for cell in row[:max_columns]]
                if any(values):
                    last_row = row_index
                    last_column = max(last_column, max((i + 1 for i, value in enumerate(values) if value), default=0))
                if row_index >= start_row and len(rows) < max_rows:
                    rows.append({"row": row_index, "values": values})
        return {
            "file_path": str(path),
            "format": "xlsb",
            "reader": "pyxlsb",
            "sheets": sheet_names,
            "selected_sheet": selected,
            "last_row": last_row,
            "last_column": last_column,
            "rows": rows,
        }


def _read_pdf(
    path: Path,
    start_page: int,
    max_pages: int,
    start_row: int,
    max_rows: int,
) -> dict[str, Any]:
    try:
        from pypdf import PdfReader  # type: ignore
    except ImportError as error:
        raise RuntimeError(
            "PDF support requires pypdf. Run the repository install.ps1 again or install "
            "the plugin requirements with: python -m pip install -r requirements.txt"
        ) from error

    reader = PdfReader(str(path))
    if reader.is_encrypted:
        try:
            unlocked = reader.decrypt("")
        except Exception as error:
            raise ValueError("The PDF is encrypted and cannot be read without a password") from error
        if not unlocked:
            raise ValueError("The PDF is encrypted and cannot be read without a password")

    total_pages = len(reader.pages)
    if start_page > total_pages and total_pages:
        raise ValueError(f"start_page {start_page} exceeds PDF page count {total_pages}")
    end_page = min(total_pages, start_page - 1 + max_pages)
    rows: list[dict[str, Any]] = []
    extracted_line_count = 0
    pages_with_text = 0

    for page_index in range(start_page - 1, end_page):
        page = reader.pages[page_index]
        try:
            text = page.extract_text(extraction_mode="layout") or ""
        except TypeError:
            text = page.extract_text() or ""
        page_has_text = False
        for line_on_page, raw_line in enumerate(text.splitlines(), start=1):
            line = raw_line.rstrip()
            if not line.strip():
                continue
            page_has_text = True
            extracted_line_count += 1
            if extracted_line_count < start_row or len(rows) >= max_rows:
                continue
            rows.append(
                {
                    "row": extracted_line_count,
                    "page": page_index + 1,
                    "line_on_page": line_on_page,
                    "values": [line],
                }
            )
        if page_has_text:
            pages_with_text += 1

    warnings = [
        "PDF text is returned as page lines. The model must reconstruct columns and verify each product row."
    ]
    if pages_with_text == 0:
        warnings.append(
            "No extractable text was found. This PDF may be scanned and requires OCR before analysis."
        )
    elif pages_with_text < max(0, end_page - start_page + 1):
        warnings.append("Some selected PDF pages contained no extractable text and may require OCR.")

    return {
        "file_path": str(path),
        "format": "pdf",
        "reader": "pypdf",
        "sheets": ["PDF"],
        "selected_sheet": "PDF",
        "total_pages": total_pages,
        "selected_page_start": start_page,
        "selected_page_end": end_page,
        "last_row": extracted_line_count,
        "last_column": 1,
        "rows": rows,
        "warnings": warnings,
    }


def inspect_product_list(arguments: dict[str, Any]) -> dict[str, Any]:
    path = _require_file(arguments["file_path"], (".xlsb", ".xlsx", ".csv", ".pdf"))
    sheet_name = arguments.get("sheet_name") or None
    start_row = int(arguments.get("start_row", 1))
    max_rows = int(arguments.get("max_rows", 50))
    max_columns = int(arguments.get("max_columns", 20))
    if path.suffix.casefold() == ".pdf":
        return _read_pdf(
            path,
            start_page=int(arguments.get("start_page", 1)),
            max_pages=int(arguments.get("max_pages", 50)),
            start_row=start_row,
            max_rows=max_rows,
        )
    if path.suffix.casefold() == ".xlsb":
        return _read_xlsb(path, sheet_name, start_row, max_rows, max_columns)
    if path.suffix.casefold() == ".csv":
        rows = []
        last_row = 0
        last_column = 0
        with path.open("r", encoding="utf-8-sig", newline="") as source:
            reader = csv.reader(source)
            for row_number, row in enumerate(reader, start=1):
                last_row = row_number
                last_column = max(last_column, len(row))
                if row_number >= start_row and len(rows) < max_rows:
                    rows.append({"row": row_number, "values": row[:max_columns]})
        return {
            "file_path": str(path),
            "format": "csv",
            "sheets": ["CSV"],
            "selected_sheet": "CSV",
            "last_row": last_row,
            "last_column": last_column,
            "rows": rows,
        }
    with XlsxReader(path) as workbook:
        sheet = workbook.sheet(sheet_name)
        last_row, last_column = workbook.dimension(sheet["path"])
        rows = [
            {"row": row_number, "values": values}
            for row_number, values in workbook.rows(
                sheet["path"], start_row=start_row, max_rows=max_rows, max_columns=max_columns
            )
        ]
        return {
            "file_path": str(path),
            "format": "xlsx",
            "sheets": [item["name"] for item in workbook.sheets],
            "selected_sheet": sheet["name"],
            "last_row": last_row,
            "last_column": last_column,
            "rows": rows,
        }


def _archive_sections(archive_path: Path) -> list[dict[str, Any]]:
    sections = []
    with zipfile.ZipFile(archive_path) as archive:
        for name in archive.namelist():
            if not name.casefold().endswith(".xlsx"):
                continue
            with XlsxReader(archive.read(name)) as workbook:
                sheet = workbook.sheet()
                last_row, last_column = workbook.dimension(sheet["path"])
                rows = dict(
                    workbook.rows(sheet["path"], start_row=5, max_rows=2, max_columns=17)
                )
                headers = rows.get(5, [""] * 17)
                status_row = rows.get(6, [""] * 17)
                status = status_row[2] if len(status_row) > 2 else ""
                sections.append(
                    {
                        "source_file": name,
                        "sheet_name": sheet["name"],
                        "status": status,
                        "last_row": last_row,
                        "last_column": last_column,
                        "approximate_data_rows": max(0, (last_row or 0) - 6),
                        "headers": headers,
                    }
                )
    return sections


def _validate_grls_archive(path: Path) -> list[dict[str, Any]]:
    if not zipfile.is_zipfile(path):
        raise ValueError("The supplied file is not a valid ZIP archive")
    try:
        sections = _archive_sections(path)
    except (KeyError, ET.ParseError, zipfile.BadZipFile) as error:
        raise ValueError(f"The ZIP does not contain readable GRLS XLSX sections: {error}") from error
    if not sections:
        raise ValueError("The ZIP contains no readable .xlsx GRLS sections")
    statuses = {_clean(section.get("status")) for section in sections}
    if not any(status in statuses for status in ACTIVE_STATUSES):
        raise ValueError(
            "The ZIP does not contain any expected current-market GRLS status section"
        )
    return sections


def update_grls_archive(arguments: dict[str, Any]) -> dict[str, Any]:
    source_path = _require_file(arguments["archive_path"], (".zip",))
    sections = _validate_grls_archive(source_path)
    archive_as_of = _infer_archive_date(source_path.name, _clean(arguments.get("archive_as_of")) or None)
    if not archive_as_of:
        raise ValueError(
            "Could not infer the archive date from its filename. Pass archive_as_of as YYYY-MM-DD."
        )

    data_dir = _user_data_dir()
    data_dir.mkdir(parents=True, exist_ok=True)
    target_path = data_dir / "grls-current.zip"
    metadata_path = data_dir / "grls-current.json"
    temporary_archive = data_dir / f"grls-current.{os.getpid()}.tmp"
    temporary_metadata = data_dir / f"grls-current.{os.getpid()}.json.tmp"
    try:
        shutil.copy2(source_path, temporary_archive)
        checksum = _sha256(temporary_archive)
        metadata = {
            "archive_as_of": archive_as_of,
            "installed_at_utc": datetime.now(timezone.utc).isoformat(),
            "original_filename": source_path.name,
            "sha256": checksum,
            "bytes": temporary_archive.stat().st_size,
            "section_count": len(sections),
            "statuses": sorted({_clean(section.get("status")) for section in sections if _clean(section.get("status"))}),
        }
        temporary_metadata.write_text(
            json.dumps(metadata, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        os.replace(temporary_archive, target_path)
        os.replace(temporary_metadata, metadata_path)
    finally:
        temporary_archive.unlink(missing_ok=True)
        temporary_metadata.unlink(missing_ok=True)

    status = get_grls_status({})
    status.update(
        {
            "updated": True,
            "validated_sections": len(sections),
            "message": "Архив ГРЛС проверен и установлен как активная база.",
        }
    )
    return status


def inspect_grls_archive(arguments: dict[str, Any]) -> dict[str, Any]:
    path = _resolve_grls_archive(arguments)
    sections = _archive_sections(path)
    active_path, source, metadata = _active_grls()
    uses_active = path == active_path
    return {
        "archive_path": str(path),
        "archive_source": source if uses_active else "explicit",
        "archive_as_of": _clean(metadata.get("archive_as_of")) if uses_active else _infer_archive_date(path.name),
        "sections": sections,
        "active_status_default": ACTIVE_STATUSES,
        "note": "Counts come from worksheet dimensions; records start at row 7.",
    }


def _iter_grls_records(archive_path: Path, statuses: list[str] | None) -> Iterator[dict[str, Any]]:
    wanted = set(statuses or ACTIVE_STATUSES)
    with zipfile.ZipFile(archive_path) as archive:
        for name in archive.namelist():
            if not name.casefold().endswith(".xlsx"):
                continue
            workbook_bytes = archive.read(name)
            with XlsxReader(workbook_bytes) as workbook:
                sheet = workbook.sheet()
                status_rows = dict(
                    workbook.rows(sheet["path"], start_row=6, max_rows=1, max_columns=17)
                )
                status = status_rows.get(6, ["", "", ""])[2]
                if wanted and status not in wanted:
                    continue
                for row_number, row in workbook.rows(
                    sheet["path"], start_row=7, max_rows=None, max_columns=17
                ):
                    registration_number = row[2] if len(row) > 2 else ""
                    if not registration_number:
                        continue
                    record = {
                        field: _clean(row[index]) if index < len(row) else ""
                        for field, index in GRLS_COLUMN_INDEX.items()
                    }
                    record.update(
                        {"source_status": status, "source_file": name, "source_row": row_number}
                    )
                    yield record


def _normalized(value: Any, case_sensitive: bool) -> str:
    text = _clean(value)
    return text if case_sensitive else text.casefold()


def search_grls(arguments: dict[str, Any]) -> dict[str, Any]:
    archive_path = _resolve_grls_archive(arguments)
    queries = [_clean(query) for query in arguments["queries"] if _clean(query)]
    if not queries:
        raise ValueError("At least one non-empty query is required")
    fields = arguments.get("fields") or ["inn"]
    invalid_fields = [field for field in fields if field not in GRLS_FIELDS]
    if invalid_fields:
        raise ValueError(f"Unsupported fields: {', '.join(invalid_fields)}")
    case_sensitive = bool(arguments.get("case_sensitive", False))
    normalized_queries = [_normalized(query, case_sensitive) for query in queries]
    match_mode = arguments.get("match_mode", "contains")
    require_all = bool(arguments.get("require_all_queries", False))
    cursor = int(arguments.get("cursor", 0))
    max_results = int(arguments.get("max_results", 200))
    results = []
    matched_count = 0
    has_more = False

    for record in _iter_grls_records(archive_path, arguments.get("statuses")):
        values = [_normalized(record.get(field, ""), case_sensitive) for field in fields]
        if match_mode == "exact":
            matches = [any(query == value for value in values) for query in normalized_queries]
        else:
            matches = [any(query in value for value in values) for query in normalized_queries]
        is_match = all(matches) if require_all else any(matches)
        if not is_match:
            continue
        if matched_count < cursor:
            matched_count += 1
            continue
        if len(results) >= max_results:
            has_more = True
            break
        results.append(record)
        matched_count += 1

    return {
        "archive_path": str(archive_path),
        "queries": queries,
        "fields": fields,
        "match_mode": match_mode,
        "cursor": cursor,
        "returned": len(results),
        "has_more": has_more,
        "next_cursor": cursor + len(results) if has_more else None,
        "records": results,
        "warning": "Literal retrieval only. Review candidate values before accepting an INN mapping.",
    }


def distinct_grls_values(arguments: dict[str, Any]) -> dict[str, Any]:
    archive_path = _resolve_grls_archive(arguments)
    field = arguments["field"]
    if field not in GRLS_FIELDS:
        raise ValueError(f"Unsupported field: {field}")
    case_sensitive = bool(arguments.get("case_sensitive", False))
    terms = [_normalized(term, case_sensitive) for term in arguments.get("contains_terms", []) if _clean(term)]
    counts: Counter[str] = Counter()
    statuses_by_value: dict[str, Counter[str]] = {}
    for record in _iter_grls_records(archive_path, arguments.get("statuses")):
        value = _clean(record.get(field, ""))
        normalized_value = _normalized(value, case_sensitive)
        if terms and not any(term in normalized_value for term in terms):
            continue
        counts[value] += 1
        statuses_by_value.setdefault(value, Counter())[record["source_status"]] += 1
    max_values = int(arguments.get("max_values", 200))
    values = [
        {"value": value, "source_rows": count, "statuses": dict(statuses_by_value[value])}
        for value, count in counts.most_common(max_values)
    ]
    return {
        "archive_path": str(archive_path),
        "field": field,
        "returned": len(values),
        "total_distinct": len(counts),
        "values": values,
        "warning": "Distinct raw values are evidence, not an automatic equivalence decision.",
    }


def _decision_records(decision: dict[str, Any], records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    accepted = set(decision.get("accepted_registry_inns") or [])
    candidates = [record for record in records if record["inn"] in accepted]
    unique: list[dict[str, Any]] = []
    seen_registration_numbers: set[str] = set()
    for record in candidates:
        registration_number = _clean(record.get("registration_number"))
        if registration_number and registration_number in seen_registration_numbers:
            continue
        if registration_number:
            seen_registration_numbers.add(registration_number)
        unique.append(record)
    return unique


def _deduplicate_registration_records(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Keep one detail row per exact non-empty registration number."""
    unique: list[dict[str, Any]] = []
    seen_registration_numbers: set[str] = set()
    for record in records:
        registration_number = _clean(record.get("registration_number"))
        if registration_number and registration_number in seen_registration_numbers:
            continue
        if registration_number:
            seen_registration_numbers.add(registration_number)
        unique.append(record)
    return unique


def _is_pharmaceutical_substance(record: dict[str, Any]) -> bool:
    return "субстанц" in _clean(record.get("release_forms")).casefold()


def _decision_summary(decision: dict[str, Any], records: list[dict[str, Any]]) -> dict[str, Any]:
    matching = _decision_records(decision, records)
    form_terms = [term.casefold() for term in decision.get("form_terms", []) if _clean(term)]
    unique_ru = {record["registration_number"] for record in matching if record["registration_number"]}
    unique_trade_names = {record["trade_name"] for record in matching if record["trade_name"]}
    ved_ru = {
        record["registration_number"]
        for record in matching
        if record["registration_number"] and record["ved"].casefold() == "да"
    }
    substance_ru = {
        record["registration_number"]
        for record in matching
        if record["registration_number"] and _is_pharmaceutical_substance(record)
    }
    medicine_ru = unique_ru - substance_ru
    competitive_records = [
        record
        for record in matching
        if record["registration_number"]
        and not _is_pharmaceutical_substance(record)
        and form_terms
        and any(term in record["release_forms"].casefold() for term in form_terms)
    ]
    form_ru = {
        record["registration_number"]
        for record in competitive_records
    }
    market_manufacturers = {
        _clean(record.get("holder"))
        for record in competitive_records
        if _clean(record.get("holder"))
    }
    summary = dict(decision)
    summary.update(
        {
            "registration_count": len(unique_ru),
            "source_row_count": len(matching),
            "trade_name_count": len(unique_trade_names),
            "ved_count": len(ved_ru),
            "medicine_count": len(medicine_ru),
            "substance_count": len(substance_ru),
            "form_registration_count": len(form_ru),
            "market_manufacturer_count": len(market_manufacturers),
        }
    )
    return summary


def _json_for_html(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c")


def _render_template(template: str, replacements: dict[str, str]) -> str:
    rendered = template
    for key, value in replacements.items():
        rendered = rendered.replace("{{" + key + "}}", value)
    unresolved = sorted(set(re.findall(r"\{\{[A-Z0-9_]+\}\}", rendered)))
    if unresolved:
        raise RuntimeError(f"Unresolved report template tokens: {', '.join(unresolved)}")
    return rendered


def build_html_report(arguments: dict[str, Any]) -> dict[str, Any]:
    archive_path = _resolve_grls_archive(arguments)
    output_path = Path(arguments["output_path"]).expanduser().resolve()
    if output_path.suffix.casefold() != ".html":
        raise ValueError("output_path must end with .html")
    decisions = arguments["decisions"]
    accepted = {
        inn
        for decision in decisions
        for inn in decision.get("accepted_registry_inns", [])
        if _clean(inn)
    }
    records = [
        record
        for record in _iter_grls_records(archive_path, ACTIVE_STATUSES)
        if record["inn"] in accepted
    ]
    for record in records:
        record["record_type"] = (
            "Фармацевтическая субстанция"
            if _is_pharmaceutical_substance(record)
            else "Лекарственный препарат"
        )
        record["reregistration"] = "Нет данных в выгрузке"
    summaries = [_decision_summary(decision, records) for decision in decisions]
    selected_record_ids = {
        id(record)
        for decision in decisions
        for record in _decision_records(decision, records)
    }
    selected_records = _deduplicate_registration_records(
        [record for record in records if id(record) in selected_record_ids]
    )
    unique_ru = {record["registration_number"] for record in selected_records if record["registration_number"]}
    ved_ru = {
        record["registration_number"]
        for record in selected_records
        if record["registration_number"] and record["ved"].casefold() == "да"
    }
    unique_groups = {
        decision.get("normalized_inn", "") for decision in decisions if decision.get("normalized_inn")
    }
    found_groups = {
        decision.get("normalized_inn", "")
        for decision, summary in zip(decisions, summaries)
        if summary["registration_count"] > 0
    }
    template_path = Path(__file__).resolve().parents[1] / "assets" / "report-template.html"
    if not template_path.is_file():
        raise RuntimeError(f"Report template is missing: {template_path}")
    template = template_path.read_text(encoding="utf-8")
    title = _clean(arguments.get("title")) or "Product list × ГРЛС"
    source_line_parts = []
    if arguments.get("product_list_name"):
        source_line_parts.append(f"Product list: {_clean(arguments['product_list_name'])}")
    grls_as_of = _clean(arguments.get("grls_as_of"))
    if not grls_as_of and not _clean(arguments.get("archive_path")):
        grls_as_of = get_grls_status({})["archive_as_of"]
    if grls_as_of:
        source_line_parts.append(f"ГРЛС по состоянию на {grls_as_of}")
    source_line_parts.append("Локальный автономный HTML")
    rendered = _render_template(
        template,
        {
            "PAGE_TITLE": html.escape(title),
            "REPORT_TITLE": html.escape(title),
            "SOURCE_LINE": html.escape(" · ".join(source_line_parts)),
            "PRODUCT_COUNT": str(len(decisions)),
            "GROUP_COUNT": str(len(unique_groups)),
            "FOUND_GROUP_COUNT": str(len(found_groups)),
            "UNIQUE_RU_COUNT": str(len(unique_ru)),
            "VED_RU_COUNT": str(len(ved_ru)),
            "PRODUCTS_JSON": _json_for_html(summaries),
            "RECORDS_JSON": _json_for_html(selected_records),
        },
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(rendered, encoding="utf-8")
    return {
        "report_path": str(output_path),
        "archive_path": str(archive_path),
        "grls_as_of": grls_as_of,
        "product_rows": len(decisions),
        "unique_inn_groups": len(unique_groups),
        "found_inn_groups": len(found_groups),
        "source_rows": len(selected_records),
        "unique_registration_numbers": len(unique_ru),
        "ved_registration_numbers": len(ved_ru),
        "bytes": output_path.stat().st_size,
    }


def validate_html_report(arguments: dict[str, Any]) -> dict[str, Any]:
    path = _require_file(arguments["report_path"], (".html",))
    text = path.read_text(encoding="utf-8")
    errors = []
    required_ids = [
        "summaryBody",
        "summaryTopScroll",
        "summaryTopScrollSpacer",
        "summaryTableScroll",
        "registryBody",
        "searchInput",
        "innFilter",
        "formFilter",
        "typeFilter",
        "vedFilter",
        "countryFilter",
        "productsData",
        "recordsData",
        "detailDialog",
        "exportCsv",
    ]
    for element_id in required_ids:
        if f'id="{element_id}"' not in text:
            errors.append(f"Missing element id: {element_id}")

    def embedded_json(element_id: str) -> Any:
        pattern = re.compile(
            rf'<script id="{re.escape(element_id)}" type="application/json">(.*?)</script>',
            re.DOTALL,
        )
        match = pattern.search(text)
        if not match:
            return None
        return json.loads(match.group(1))

    products = None
    records = None
    try:
        products = embedded_json("productsData")
    except Exception as error:
        errors.append(f"productsData JSON is invalid: {error}")
    try:
        records = embedded_json("recordsData")
    except Exception as error:
        errors.append(f"recordsData JSON is invalid: {error}")
    if isinstance(records, list):
        invalid_statuses = sorted(
            {
                _clean(record.get("source_status"))
                for record in records
                if _clean(record.get("source_status")) != "Действующий"
            }
        )
        if invalid_statuses:
            errors.append(
                "Report contains records outside the Действующий section: "
                + ", ".join(invalid_statuses)
            )
        registration_numbers = [
            _clean(record.get("registration_number"))
            for record in records
            if _clean(record.get("registration_number"))
        ]
        if len(registration_numbers) != len(set(registration_numbers)):
            errors.append("Report contains duplicate registration numbers")
    if 'id="statusFilter"' in text or 'data-key="source_status"' in text:
        errors.append("Report still contains a status filter or status column")
    if 'id="manufacturerFilter"' in text:
        errors.append("Report still contains the removed manufacturer filter")
    if 'data-key="trade_name_count"' in text or 'data-key="confidence"' in text:
        errors.append("Report still contains a removed summary column")
    if 'data-key="reregistration"' in text:
        errors.append("Report still contains a re-registration table column")
    if "row.addEventListener('pointerdown'" not in text or "if (dragged) return;" not in text:
        errors.append("Report does not protect summary text selection from row navigation")
    if 'data-primary-metric="medicine"' not in text or "const hasDesiredForm" not in text:
        errors.append("Report does not highlight medicine count when product-list form is absent")
    return {
        "report_path": str(path),
        "valid": not errors,
        "errors": errors,
        "product_rows": len(products) if isinstance(products, list) else None,
        "source_rows": len(records) if isinstance(records, list) else None,
        "bytes": path.stat().st_size,
    }


TOOL_HANDLERS = {
    "get_grls_status": get_grls_status,
    "update_grls_archive": update_grls_archive,
    "inspect_product_list": inspect_product_list,
    "inspect_grls_archive": inspect_grls_archive,
    "search_grls": search_grls,
    "distinct_grls_values": distinct_grls_values,
    "build_html_report": build_html_report,
    "validate_html_report": validate_html_report,
}


def _tool_result(data: dict[str, Any], is_error: bool = False) -> dict[str, Any]:
    text = json.dumps(data, ensure_ascii=False, indent=2)
    result: dict[str, Any] = {"content": [{"type": "text", "text": text}], "isError": is_error}
    if not is_error:
        result["structuredContent"] = data
    return result


def _handle_request(message: dict[str, Any]) -> dict[str, Any] | None:
    method = message.get("method")
    request_id = message.get("id")
    if method == "initialize":
        requested = message.get("params", {}).get("protocolVersion")
        result = {
            "protocolVersion": requested or PROTOCOL_VERSION,
            "capabilities": {"tools": {"listChanged": False}},
            "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION},
        }
    elif method == "ping":
        result = {}
    elif method == "tools/list":
        result = {"tools": TOOLS}
    elif method == "tools/call":
        params = message.get("params", {})
        name = params.get("name")
        arguments = params.get("arguments") or {}
        handler = TOOL_HANDLERS.get(name)
        if not handler:
            result = _tool_result({"error": f"Unknown tool: {name}"}, is_error=True)
        else:
            try:
                result = _tool_result(handler(arguments))
            except Exception as error:
                result = _tool_result(
                    {"error": str(error), "tool": name, "error_type": type(error).__name__},
                    is_error=True,
                )
    elif method in {"notifications/initialized", "notifications/cancelled"}:
        return None
    else:
        if request_id is None:
            return None
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "error": {"code": -32601, "message": f"Method not found: {method}"},
        }
    if request_id is None:
        return None
    return {"jsonrpc": "2.0", "id": request_id, "result": result}


def run_stdio() -> None:
    # MCP JSON-RPC is UTF-8. Windows may otherwise inherit a legacy console
    # code page that cannot encode symbols used in Russian UI descriptions.
    if hasattr(sys.stdin, "reconfigure"):
        sys.stdin.reconfigure(encoding="utf-8")
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    for raw_line in sys.stdin:
        line = raw_line.strip()
        if not line:
            continue
        try:
            message = json.loads(line)
            response = _handle_request(message)
            if response is not None:
                sys.stdout.write(json.dumps(response, ensure_ascii=False, separators=(",", ":")) + "\n")
                sys.stdout.flush()
        except Exception as error:
            sys.stdout.write(
                json.dumps(
                    {
                        "jsonrpc": "2.0",
                        "id": None,
                        "error": {"code": -32700, "message": str(error)},
                    },
                    ensure_ascii=False,
                    separators=(",", ":"),
                )
                + "\n"
            )
            sys.stdout.flush()


if __name__ == "__main__":
    run_stdio()
