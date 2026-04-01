"""
utils/text_utils.py - Shared text normalization utilities used across services.
"""

import re
from typing import Optional


# Entity prefix patterns to strip during normalization
_ENTITY_PREFIXES = re.compile(
    r'^(applicant|coapplicant\d*|guarantor|loan|document|fee|customer)[_\s\.]?',
    re.IGNORECASE
)

# Separators / noise chars
_SEPARATORS = re.compile(r"[_\s\.\-\'\(\)/\\|:,;]")


def normalize_basic(text: str) -> str:
    """
    Lowercase + strip separators, preserving entity prefixes.
    Used for field_dictionary indexing.
    """
    if not text:
        return ""
    text = text.lower()
    text = _SEPARATORS.sub("", text)
    return text


def normalize_field(text: str) -> str:
    """
    Lowercase + strip separators + strip entity prefixes.
    Used for alias registry lookup.
    """
    if not text:
        return ""
    text = text.lower()
    text = _SEPARATORS.sub("", text)
    text = _ENTITY_PREFIXES.sub("", text)
    return text


def extract_entity_from_path(json_path: str) -> str:
    """
    Infer entity from a JSON dot-notation path.
    e.g. 'loanAccount.customer.firstName' → 'APPLICANT'
    """
    path_lower = json_path.lower()
    if "coapplicant" in path_lower or "co_applicant" in path_lower:
        return "COAPPLICANT"
    if "guarantor" in path_lower:
        return "GUARANTOR"
    if "customer" in path_lower or "applicant" in path_lower:
        return "APPLICANT"
    if "loan" in path_lower:
        return "LOAN"
    if "document" in path_lower or "file" in path_lower:
        return "DOCUMENT"
    if "fee" in path_lower or "charge" in path_lower:
        return "FEE"
    return "OTHER"


def truncate(text: str, max_len: int = 200) -> str:
    """Truncate a string for safe logging."""
    if len(text) <= max_len:
        return text
    return text[:max_len] + "..."


def clean_json_fence(text: str) -> str:
    """Remove markdown code fences from an LLM JSON response."""
    text = re.sub(r'^```(?:json)?\s*', '', text, flags=re.MULTILINE)
    text = re.sub(r'\s*```$', '', text, flags=re.MULTILINE)
    return text.strip()


def extract_json_block(text: str) -> Optional[str]:
    """Extract the outermost JSON object/array from a text blob."""
    # Try object first
    start = text.find('{')
    end = text.rfind('}')
    if start != -1 and end != -1 and end > start:
        return text[start:end + 1]
    # Try array
    start = text.find('[')
    end = text.rfind(']')
    if start != -1 and end != -1 and end > start:
        return text[start:end + 1]
    return None