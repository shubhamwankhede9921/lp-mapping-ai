#!/usr/bin/env python3
 #input _parser
import json
import logging
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Optional, Any, Tuple
 
import pandas as pd
from openpyxl import load_workbook
 
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
 
 
# =========================
# FIELD STRUCTURE
# =========================
class FieldDefinition:
    def __init__(
        self,
        field_name: str,
        column_category: Optional[str] = None,
        data_type: Optional[str] = None,
        is_required: Optional[bool] = None,
        sample_value: Optional[str] = None,
        description: Optional[str] = None,
        source_sheet: Optional[str] = None
    ):
        self.field_name = field_name
        self.column_category = column_category
        self.data_type = data_type
        self.is_required = is_required
        self.sample_value = sample_value
        self.description = description
        self.source_sheet = source_sheet

    def to_dict(self):
        return self.__dict__


# =========================
# IDENTIFIER DETECTION
# =========================
def _looks_like_identifier(val: str) -> bool:
    val = val.rstrip('.,:;')
    words = val.split()
 
    if not words:
        return False
 
    if re.match(r'^[A-Z][A-Za-z0-9_\s]+$', val):
        return True
 
    if re.match(r'^[a-z][a-zA-Z0-9_]+$', val):
        return True
 
    if re.match(r'^[A-Za-z][A-Za-z0-9_]+$', val):
        return True
 
    if 1 <= len(words) <= 6:
        return True
 
    return False


# =========================
# DOTTED-PATH DETECTION
# =========================
def _is_dotted_path(val: str) -> bool:
    """
    Return True if the value looks like a dotted-path field name,
    e.g. 'loanAccount.cbDateTime', 'a.b.c', 'foo.barBaz'.
    A dotted path must have at least one dot with non-empty segments
    on both sides, and no whitespace.
    """
    if ' ' in val:
        return False
    parts = val.split('.')
    if len(parts) < 2:
        return False
    return all(len(p) > 0 for p in parts)


# =========================
# HEADER DETECTION
# =========================

_EXACT_HEADERS = {
    "field names", "field name", "column name", "column names",
    "header", "headers", "s.no", "s.no.", "sr no", "sr. no", "sr.no",
    "input fields", "input field", "data fields", "data field",
    "description", "remarks", "notes", "mandatory", "required",
}

_SINGLE_WORD_HEADERS = {
    "field", "fields", "column", "columns", "header",
    "input", "inputs", "description", "remarks",
}

_LABEL_SUFFIXES = {
    'input', 'inputs', 'fields', 'field', 'field names', 'field name',
    'data', 'details', 'info', 'information', 'section',
}

_GENERIC_WORDS = {
    'client', 'customer', 'applicant', 'loan',
    'coapplicant', 'details', 'info',
}


def _is_header_label(val: str) -> bool:
    if not val:
        return False
 
    stripped = val.strip()
    lower    = stripped.lower()
 
    if lower in _EXACT_HEADERS:
        return True
 
    words = stripped.split()
 
    if len(words) == 1 and lower in _SINGLE_WORD_HEADERS:
        return True
 
    if len(words) >= 2 and words[-1].lower() in _LABEL_SUFFIXES:
        return True
 
    if len(words) >= 2:
        generic_count = sum(1 for w in words if w.lower() in _GENERIC_WORDS)
        if generic_count == len(words):
            return True
 
    return False


def _is_purely_numeric(val: str) -> bool:
    """Return True if the value is a plain integer/float string (serial numbers, etc.)."""
    try:
        float(val)
        return True
    except ValueError:
        return False


# =========================
# TABULAR HEADER DETECTION
# =========================

FIELD_COL_HEADERS = {
    "loan api fields",
    "fields",
    "field name",
    "field names",
    "parameter",
    "parameter name",
    "input field",
    "input fields",
    "column name",
    "column names",
    "object name",
}

SAMPLE_COL_HEADERS = {
    "sample values",
    "sample value",
    "example",
    "examples",
}

CATEGORY_COL_HEADERS = {
    "field category",
    "category",
    "api name",
    "group",
    "section",
    "module",
}


def find_header_row_and_cols(
    ws, max_scan_rows: int = 10
) -> Tuple[Optional[int], Optional[int], Optional[int], Optional[int]]:
    for row_idx, row in enumerate(
        ws.iter_rows(max_row=max_scan_rows, values_only=True), start=1
    ):
        field_col    = None
        sample_col   = None
        category_col = None
        for col_idx, cell in enumerate(row):
            if cell is None:
                continue
            norm = str(cell).strip().lower()
            if norm in FIELD_COL_HEADERS and field_col is None:
                field_col = col_idx
            elif norm in SAMPLE_COL_HEADERS and sample_col is None:
                sample_col = col_idx
            elif norm in CATEGORY_COL_HEADERS and category_col is None:
                category_col = col_idx
 
        if field_col is not None:
            return row_idx + 1, field_col, sample_col, category_col
 
    return None, None, None, None


# =========================
# TABULAR SHEET PARSER
# =========================
def parse_tabular_sheet(
    ws, sheet_name: str, global_seen: Optional[set] = None
) -> Optional[List[dict]]:
    """
    Try to parse *ws* as a tabular sheet.
    Returns a list of field dicts, or None if no recognised field header found.

    Dedup rules:
    - If a category column exists: dedup key is (category, field_name)
      → same field name under different categories is kept (e.g. CITY under
        Customer vs Co-applicant).
    - Without a category column: dedup key is field_name alone
      → same field name seen anywhere (across sheets) is kept only once.
    - global_seen is shared across all sheets to enforce cross-sheet dedup.
    """
    if global_seen is None:
        global_seen = set()
 
    data_start, field_col, sample_col, category_col = find_header_row_and_cols(ws)
    if field_col is None:
        return None
 
    fields = []
    local_seen = set()   # within-sheet dedup (for category-aware key)
 
    for row in ws.iter_rows(min_row=data_start, values_only=True):
        val = row[field_col] if field_col < len(row) else None
        if val is None:
            continue
 
        field_name = str(val).strip()
        if not field_name or field_name.lower() == 'nan':
            continue
        if _is_header_label(field_name):
            logger.debug(f"Tabular skip (header label): {field_name!r}")
            continue
        if _is_purely_numeric(field_name):
            logger.debug(f"Tabular skip (numeric): {field_name!r}")
            continue
        if _is_dotted_path(field_name):
            logger.debug(f"Tabular skip (dotted-path): {field_name!r}")
            continue
 
        # Resolve category value
        cat_val = None
        if category_col is not None and category_col < len(row):
            cv = row[category_col]
            if cv is not None:
                cat_val = str(cv).strip()
 
        # Dedup key:
        # - With category column  → (category, field_name): same field in
        #   different categories is a distinct entry.
        # - Without category column → field_name only: one entry per unique
        #   name across the entire workbook.
        if category_col is not None:
            dedup_key = (cat_val, field_name)
            if dedup_key in local_seen:
                continue
            local_seen.add(dedup_key)
        else:
            dedup_key = field_name.lower()
            if dedup_key in global_seen:
                logger.debug(f"Tabular skip (global duplicate): {field_name!r}")
                continue
            global_seen.add(dedup_key)
 
        sample = None
        if sample_col is not None and sample_col < len(row):
            sv = row[sample_col]
            if sv is not None:
                sample = str(sv).strip()
 
        fields.append(
            FieldDefinition(
                field_name=field_name,
                column_category=cat_val if cat_val else sheet_name,
                sample_value=sample,
                source_sheet=sheet_name,
            ).to_dict()
        )
 
    return fields


# =========================
# KV SHEET DETECTION
# =========================
def is_key_value_sheet(ws) -> bool:
    values = []
 
    for i in range(1, min(ws.max_row + 1, 20)):
        v = ws.cell(row=i, column=1).value
        if v:
            values.append(str(v).strip())
 
    if len(values) < 2:
        return False
 
    identifier_ratio = sum(_looks_like_identifier(v) for v in values) / len(values)
    return identifier_ratio > 0.6


# =========================
# PARSE KV SHEET
# =========================
def parse_key_value_sheet(
    ws, sheet_name: str, global_seen: Optional[set] = None
) -> List[dict]:
    """
    Parse a key-value sheet (column A = field names).

    global_seen is shared across all sheets so that the same field name
    appearing in multiple KV sheets (or after a tabular sheet) is kept
    only once.
    """
    if global_seen is None:
        global_seen = set()
 
    fields = []
 
    for i in range(1, ws.max_row + 1):
        val = ws.cell(row=i, column=1).value
 
        if not val:
            continue
 
        field_name = str(val).strip()
 
        if not field_name or field_name.lower() == 'nan':
            continue
        if _is_header_label(field_name):
            logger.debug(f"KV skip (header label): {field_name!r}")
            continue
        if _is_purely_numeric(field_name):
            logger.debug(f"KV skip (numeric): {field_name!r}")
            continue
        if _is_dotted_path(field_name):
            logger.debug(f"KV skip (dotted-path): {field_name!r}")
            continue
 
        dedup_key = field_name.lower()
        if dedup_key in global_seen:
            logger.debug(f"KV skip (global duplicate): {field_name!r}")
            continue
        global_seen.add(dedup_key)
 
        sample = None
        if ws.max_column >= 2:
            v = ws.cell(row=i, column=2).value
            if v:
                sample = str(v).strip()
 
        fields.append(
            FieldDefinition(
                field_name=field_name,
                column_category=sheet_name,
                sample_value=sample,
                source_sheet=sheet_name,
            ).to_dict()
        )
 
    return fields


# =========================
# EXCEL PARSER
# Route: tabular first → KV fallback
# =========================
def parse_excel_input(file_path: str) -> List[dict]:
    wb          = load_workbook(file_path, data_only=True)
    all_fields: List[dict] = []
    global_seen: set       = set()   # shared across all sheets for dedup
 
    for sheet in wb.sheetnames:
        ws = wb[sheet]
 
        if ws.sheet_state != 'visible':
            logger.info(f"{sheet} → Skipped (hidden sheet)")
            continue
 
        # --- Attempt tabular parsing first ---
        result = parse_tabular_sheet(ws, sheet, global_seen)
        if result is not None:
            logger.info(f"{sheet} → Tabular format ({len(result)} fields found)")
            all_fields.extend(result)
            continue
 
        # --- Fallback: KV list ---
        if is_key_value_sheet(ws):
            logger.info(f"{sheet} → KV format")
            all_fields.extend(parse_key_value_sheet(ws, sheet, global_seen))
        else:
            logger.info(f"{sheet} → Skipped (unrecognised format)")
 
    wb.close()
    return all_fields


# =========================
# JSON PARSER
# =========================
def parse_json_input(file_path: str) -> List[dict]:
    """
    Nested JSON → field rows.

    Output behavior (generic, not hard-coded):
    - **column_category**: immediate parent object key
      Example: "bankingDetails" becomes column_category for its children
    - **field_name**: leaf key name (e.g. "bankAccountNumber")
    - **sample_value**: leaf scalar value (often null)
    """
    with open(file_path, "r", encoding="utf-8") as f:
        raw = f.read()

    # 1) Strict JSON first
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        data = None

    def _flatten_nested(
        obj: Any,
        path: Optional[List[str]] = None,
    ) -> List[Tuple[str, str, Any]]:
        if path is None:
            path = []
        out: List[Tuple[str, str, Any]] = []

        if isinstance(obj, dict):
            for k, v in obj.items():
                out.extend(_flatten_nested(v, path + [str(k)]))
            return out
        if isinstance(obj, list):
            for i, v in enumerate(obj):
                out.extend(_flatten_nested(v, path + [str(i)]))
            return out

        # scalar leaf (including None)
        field_name = path[-1] if path else ""
        category = path[-2] if len(path) >= 2 else "API"
        if field_name:
            out.append((field_name, category, obj))
        return out

    if data is not None:
        flattened = _flatten_nested(data)
        seen = set()
        fields: List[dict] = []
        for field_name, category, sample in flattened:
            dedup_key = (str(category).strip().lower(), str(field_name).strip().lower())
            if dedup_key in seen:
                continue
            seen.add(dedup_key)
            fields.append(
                FieldDefinition(
                    field_name=str(field_name).strip(),
                    column_category=str(category).strip() or "API",
                    sample_value=None if sample is None else str(sample),
                    source_sheet="JSON",
                ).to_dict()
            )
        return fields

    # 2) Fallback: tolerate JSON-like templates with stray text/comments.
    # Extract `"section": { ... "child": null ... }` and use section as column_category.
    text = raw
    text = re.sub(r"//.*?$", "", text, flags=re.M)
    text = re.sub(r"#.*?$", "", text, flags=re.M)

    section_pat = re.compile(r"\"([A-Za-z0-9_]+)\"\\s*:\\s*\\{", re.M)
    key_null_pat = re.compile(r"\"([A-Za-z0-9_]+)\"\\s*:\\s*null", re.I)

    fields: List[dict] = []
    seen = set()
    stack: List[str] = []
    i = 0
    current_section = "API"
    while i < len(text):
        m = section_pat.match(text, i)
        if m:
            current_section = m.group(1)
            stack.append(current_section)
            i = m.end()
            continue
        if text[i] == "{":
            i += 1
            continue
        if text[i] == "}":
            if stack:
                stack.pop()
                current_section = stack[-1] if stack else "API"
            i += 1
            continue
        km = key_null_pat.match(text, i)
        if km:
            fname = km.group(1)
            category = current_section or "API"
            dedup_key = (category.strip().lower(), fname.strip().lower())
            if dedup_key not in seen:
                seen.add(dedup_key)
                fields.append(
                    FieldDefinition(
                        field_name=fname.strip(),
                        column_category=category.strip(),
                        sample_value=None,
                        source_sheet="JSON",
                    ).to_dict()
                )
            i = km.end()
            continue
        i += 1

    return fields


# =========================
# WORD (.docx) PARSER
# =========================
def _docx_extract_text(file_path: str) -> str:
    """
    Extract text from a .docx file without external dependencies.
    """
    with zipfile.ZipFile(file_path) as zf:
        xml_bytes = zf.read("word/document.xml")
    root = ET.fromstring(xml_bytes)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    parts: List[str] = []
    for para in root.findall(".//w:p", ns):
        texts = [t.text for t in para.findall(".//w:t", ns) if t.text]
        line = "".join(texts).strip()
        if line:
            parts.append(line)
    return "\n".join(parts)


def parse_docx_input(file_path: str) -> List[dict]:
    """
    Parse a Word .docx where field names are listed as lines or table cells.
    Strategy: extract paragraph text, split into lines, then keep identifier-like values.
    """
    text = _docx_extract_text(file_path)
    lines = [ln.strip() for ln in re.split(r"[\r\n]+", text) if ln and ln.strip()]

    # If the document contains JSON-like key blocks, preserve parent → child relationship:
    #   "bankingDetails": { ... "ifscCode": null ... }  => column_category=bankingDetails, field_name=ifscCode
    section_pat = re.compile(r"\"?([A-Za-z0-9_]+)\"?\s*:\s*\{")
    key_null_pat = re.compile(r"\"?([A-Za-z0-9_]+)\"?\s*:\s*null\b", re.I)
    key_scalar_pat = re.compile(r"\"?([A-Za-z0-9_]+)\"?\s*:\s*([A-Za-z0-9_@./-]+)\b")

    seen = set()
    fields: List[dict] = []
    current_section = "WORD"
    section_stack: List[str] = []

    for ln in lines:
        # Track section headers like: bankingDetails: {
        sm = section_pat.search(ln)
        if sm:
            current_section = sm.group(1)
            section_stack.append(current_section)
            continue
        if "}" in ln and section_stack:
            # pop one level when braces appear (best-effort)
            section_stack.pop()
            current_section = section_stack[-1] if section_stack else "WORD"
            continue

        # Extract "key: null" as field name (clean: no quotes/braces/null)
        km = key_null_pat.search(ln)
        if km:
            fname = km.group(1).strip()
            if fname and not _is_header_label(fname) and not _is_purely_numeric(fname) and not _is_dotted_path(fname):
                dedup_key = (current_section.lower(), fname.lower())
                if dedup_key not in seen:
                    seen.add(dedup_key)
                    fields.append(
                        FieldDefinition(
                            field_name=fname,
                            column_category=current_section,
                            sample_value=None,
                            source_sheet="WORD",
                        ).to_dict()
                    )
            continue

        # If line is a plain identifier, accept it as a standalone field
        cleaned = ln.strip().strip(",")
        cleaned = cleaned.replace("{", "").replace("}", "").replace('"', "").replace("'", "").strip()
        if not cleaned:
            continue
        if _is_header_label(cleaned) or _is_purely_numeric(cleaned) or _is_dotted_path(cleaned):
            continue
        if not _looks_like_identifier(cleaned):
            continue
        dedup_key = (current_section.lower(), cleaned.lower())
        if dedup_key in seen:
            continue
        seen.add(dedup_key)
        fields.append(
            FieldDefinition(
                field_name=cleaned,
                column_category=current_section,
                sample_value=None,
                source_sheet="WORD",
            ).to_dict()
        )

    return fields


# =========================
# MAIN ENTRY
# =========================
def parse_input(file_path: str) -> List[dict]:
    file_path = Path(file_path)
 
    if not file_path.exists():
        raise FileNotFoundError(file_path)
 
    if file_path.suffix in ['.xlsx', '.xls']:
        return parse_excel_input(str(file_path))
 
    if file_path.suffix == '.json':
        return parse_json_input(str(file_path))

    if file_path.suffix.lower() == '.docx':
        return parse_docx_input(str(file_path))
 
    raise ValueError("Unsupported file type")


# =========================
# CLI RUN
# =========================
if __name__ == "__main__":
    import sys
 
    file   = sys.argv[1]
    result = parse_input(file)
    print(json.dumps(result, indent=2))