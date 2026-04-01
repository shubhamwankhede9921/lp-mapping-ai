# app/utils/los_json_builder.py
#
# Converts flat mappings (client_column + lms_column dot-path)
# into a nested JSON tree — no LLM required.
#
# Supports two modes:
#   generate_nested_mapping — leaf = client_column value
#   generate_schema         — leaf = null

from typing import List, Dict, Any, Optional
import re


def _set_nested(tree: Dict, path_parts: List[str], value: Any) -> None:
    if not path_parts:
        return
    key  = path_parts[0]
    rest = path_parts[1:]
    if not rest:
        tree[key] = value
        return
    if key not in tree or not isinstance(tree[key], dict):
        tree[key] = {}
    _set_nested(tree[key], rest, value)


def _split_path(lms_column: str) -> List[str]:
    return [p for p in lms_column.split(".") if p.strip()]


def _is_skippable(lms_column: Optional[str], matched_excel_key: Optional[str] = None) -> bool:
    if not lms_column or not lms_column.strip():
        return True
    lms_lower = lms_column.strip().lower()
    if "loanparameter" in lms_lower:
        return True
    if matched_excel_key and "loanparameter" in matched_excel_key.lower():
        return True
    return False


def _normalise_mappings(raw_mappings: List[Any]) -> List[Dict]:
    normalised = []
    for m in raw_mappings:
        if not isinstance(m, dict):
            continue
        if "client_column" in m:
            normalised.append(m)
        else:
            for col_name, col_data in m.items():
                if isinstance(col_data, dict):
                    normalised.append({
                        "client_column":     col_name,
                        "matched_excel_key": col_data.get("matched_excel_key", ""),
                        "lms_column":        col_data.get("lms_column", ""),
                        "confidence":        col_data.get("confidence", 0.0),
                        "tier":              col_data.get("tier", ""),
                        "method":            col_data.get("method", ""),
                    })
    return normalised


def generate_nested_mapping(mappings_result: Dict) -> Dict:
    raw   = mappings_result.get("mappings", [])
    flat  = _normalise_mappings(raw)
    tree: Dict = {}

    for m in flat:
        lms_col    = (m.get("lms_column") or "").strip()
        client_col = (m.get("client_column") or "").strip()
        excel_key  = m.get("matched_excel_key", "")

        if _is_skippable(lms_col, excel_key) or not client_col:
            continue

        parts = _split_path(lms_col)
        if parts:
            _set_nested(tree, parts, client_col)

    return {"mappings": tree}


def generate_schema(mappings_result: Dict) -> Dict:
    raw   = mappings_result.get("mappings", [])
    flat  = _normalise_mappings(raw)
    tree: Dict = {}

    for m in flat:
        lms_col   = (m.get("lms_column") or "").strip()
        excel_key = m.get("matched_excel_key", "")

        if _is_skippable(lms_col, excel_key):
            continue

        parts = _split_path(lms_col)
        if parts:
            _set_nested(tree, parts, None)

    return {"schema": tree}


def transform(mappings_result: Dict, task: str) -> Dict:
    task = (task or "").strip().lower()
    if task == "generate_nested_mapping":
        return generate_nested_mapping(mappings_result)
    elif task == "generate_schema":
        return generate_schema(mappings_result)
    else:
        raise ValueError(
            f"Unknown task '{task}'. "
            f"Supported: 'generate_nested_mapping', 'generate_schema'"
        )