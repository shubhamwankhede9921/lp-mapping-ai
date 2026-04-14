"""
Shared signals for fuzzy, embedding, and prompt filtering:
field name, column_category, entity, process_name (and PUTM process_names).
"""

from __future__ import annotations

import re
from typing import Any, Dict, Optional, Set

# Aligns partner entity labels to internal json_key / excel roles
ENTITY_ROLE: Dict[str, str] = {
    "APPLICANT": "CUSTOMER",
    "COAPPLICANT": "COAPPLICANT",
    "COAPPLICANT1": "COAPPLICANT",
    "COAPPLICANT2": "COAPPLICANT",
    "COAPPLICANT3": "COAPPLICANT",
    "COAPPLICANT4": "COAPPLICANT",
    "GUARANTOR": "GUARANTOR",
    "LOAN": "LOAN",
    "DOCUMENT": "LOAN",
    "FEE": "LOAN",
    "OTHER": "CUSTOMER",
}

# Short hints appended to semantic queries so excel_key prefixes align (e.g. APPLICANT*)
ENTITY_QUERY_HINTS: Dict[str, str] = {
    "APPLICANT": "applicant primary customer borrower",
    "COAPPLICANT": "coapplicant co borrower",
    "COAPPLICANT1": "coapplicant co borrower",
    "COAPPLICANT2": "coapplicant co borrower",
    "COAPPLICANT3": "coapplicant co borrower",
    "COAPPLICANT4": "coapplicant co borrower",
    "GUARANTOR": "guarantor",
    "LOAN": "loan facility account sanction",
    "DOCUMENT": "document upload attachment",
    "FEE": "fee charge pricing",
    "OTHER": "",
}

CATEGORY_SEMANTICS = {
    "date": {"keywords": {"date", "born", "birth", "dob", "since", "from", "to"}, "boost": 1.15},
    "amount": {"keywords": {"amount", "value", "price", "cost", "fee", "charge", "salary", "income"}, "boost": 1.15},
    "name": {"keywords": {"name", "first", "last", "middle", "full", "title", "label"}, "boost": 1.15},
    "id": {"keywords": {"id", "code", "number", "ref", "reference", "identifier", "account", "pan", "aadhar"}, "boost": 1.15},
    "status": {"keywords": {"status", "state", "flag", "type", "category", "active", "approved", "verified"}, "boost": 1.10},
    "address": {"keywords": {"address", "street", "city", "state", "pin", "zip", "postal", "location"}, "boost": 1.12},
    "contact": {"keywords": {"phone", "email", "contact", "mobile", "telephone", "number"}, "boost": 1.12},
    "demographic": {"keywords": {"age", "gender", "caste", "religion", "marital", "occupation", "education"}, "boost": 1.12},
    "financial": {"keywords": {"credit", "score", "debt", "loan", "balance", "account", "bank"}, "boost": 1.10},
}


def _norm(text: str) -> str:
    text = text.lower()
    text = re.sub(r"[_\s.\-'()]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _extract_keywords(text: str) -> Set[str]:
    words = _norm(text).split()
    return {w for w in words if len(w) > 2}


def compute_category_boost(column_category: Optional[str], excel_key: str) -> float:
    if not column_category:
        return 1.0

    col_keywords = _extract_keywords(column_category)
    ek_keywords = _extract_keywords(excel_key)

    if col_keywords & ek_keywords:
        return 1.15

    for _family, config in CATEGORY_SEMANTICS.items():
        family_keywords = config["keywords"]
        col_has = any(kw in col_keywords for kw in family_keywords)
        ek_has = any(kw in ek_keywords for kw in family_keywords)
        if col_has and ek_has:
            return float(config["boost"])

    return 1.0


def effective_process(
    process_override: Optional[str],
    default_process: str = "",
) -> str:
    """Prefer non-empty override; otherwise use default (e.g. engine init process_name)."""
    if process_override is not None and str(process_override).strip():
        return str(process_override).strip()
    return (default_process or "").strip()


def is_process_compatible(
    excel_key: str,
    field_process: str,
    field_dictionary: Dict[str, Any],
) -> bool:
    """
    Whether excel_key may be used under the current pipeline process.

    Honors by_excel_key.process_name, process_names[], and treats COMBINED / empty as universal.
    """
    fp = (field_process or "").strip().upper()
    if not fp:
        return True

    info = (field_dictionary.get("by_excel_key") or {}).get(excel_key) or {}
    ek_primary = (info.get("process_name") or "").strip().upper()
    extras_raw = info.get("process_names") or []
    allowed: Set[str] = set()
    if ek_primary:
        allowed.add(ek_primary)
    for x in extras_raw:
        if x:
            allowed.add(str(x).strip().upper())

    if not allowed:
        return True
    if "COMBINED" in allowed:
        return True
    return fp in allowed


# Partner field names shorter than this (after stripping non-alphanumeric) cannot be
# meaningfully aligned to catalogue keys — category/entity context would dominate
# fuzzy/embedding/LLM and produce false positives (e.g. "as" → APPLICANTBORROWERTYPE).
SEMANTIC_FIELD_MIN_ALNUM_LEN = 3

# Short names that are still legitimate API/catalogue tokens (allowed past the min-length rule).
SEMANTIC_FIELD_ALLOWLIST = frozenset({"id", "all"})

# Standalone field tokens (alphanumeric-normalized) that are never valid as the sole
# business field name for semantic layers, even when length >= SEMANTIC_FIELD_MIN_ALNUM_LEN.
AMBIGUOUS_STANDALONE_FIELD_NAMES = frozenset({
    "the", "and", "for", "not", "but", "nor", "any", "new", "old",
    "yes", "yet", "nil", "ask", "use", "get", "set", "run", "try", "are",
})


def semantic_field_guard_reason(field_name: Optional[str]) -> Optional[str]:
    """
    If non-None, fuzzy/embedding/LLM must not map this partner field name.
    Category/entity hints alone are not enough to justify a catalogue mapping.
    """
    raw = (field_name or "").strip()
    if not raw:
        return "Field name is empty."
    alnum = re.sub(r"[^a-z0-9]+", "", raw.lower())
    if alnum in SEMANTIC_FIELD_ALLOWLIST:
        return None
    if len(alnum) < SEMANTIC_FIELD_MIN_ALNUM_LEN:
        return (
            f"Field name is too short ({len(alnum)} alphanumeric character(s); "
            f"need at least {SEMANTIC_FIELD_MIN_ALNUM_LEN}) for semantic matching."
        )
    if alnum in AMBIGUOUS_STANDALONE_FIELD_NAMES:
        return "Field name is a generic token; semantic matching would be unreliable."
    return None


def coapplicant_slot_from_entity(entity: Optional[str]) -> Optional[int]:
    """
    Co-applicant index n for LOANCOAPP{n}CUSTPARAM* (1-based).
    Returns None if entity is not a co-applicant scope.
    """
    e = (entity or "").strip().upper()
    if not e.startswith("COAPPLICANT"):
        return None
    if e == "COAPPLICANT":
        return 1
    m = re.match(r"COAPPLICANT(\d+)$", e)
    if m:
        return int(m.group(1))
    return 1


def coapplicant_custparam_base(entity: Optional[str]) -> Optional[str]:
    """Unnumbered PUTM base key LOANCOAPP{n}CUSTPARAM for co-applicant custom parameters."""
    slot = coapplicant_slot_from_entity(entity)
    if slot is None:
        return None
    return f"LOANCOAPP{slot}CUSTPARAM"


def build_semantic_query(
    field_name: str,
    column_category: Optional[str] = None,
    entity: Optional[str] = None,
    process_name: Optional[str] = None,
) -> str:
    """
    Single normalized query string combining partner field, category, entity hints, and process.
    Used by fuzzy (token) and embedding (sentence) layers.
    """
    parts: list[str] = [field_name or ""]
    if column_category:
        parts.append(column_category)
    ent = (entity or "").strip().upper()
    if ent and ent in ENTITY_QUERY_HINTS:
        hint = ENTITY_QUERY_HINTS.get(ent, "").strip()
        if hint:
            parts.append(hint)
    if process_name:
        parts.append(process_name)
    return _norm(" ".join(parts))

