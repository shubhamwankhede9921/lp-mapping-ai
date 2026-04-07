#!/usr/bin/env python3
"""
Deterministic Field Matching Engine

Core matching layer that maps partner fields to internal excel_keys without using an LLM.
Loads reference files (field_dictionary, alias_registry, entity_routing) and evaluates
fields in strict priority order: exact match → alias tiers → document detection →
document ID detection → fee detection → unmatched.

Usage:
    # As a module
    from matching_engine import MatchResult, load_references, match_field, match_batch

    refs = load_references("./references")
    result = match_field("dateOfBirth", "Customer Details", "APPLICANT", refs)
    print(result)

    # As a CLI
    python matching_engine.py --field "dateOfBirth" --category "Customer Details"
    python matching_engine.py --field "processingFee" --category "Loan Details"
    python matching_engine.py --field "aadharFrontLink" --category "Documents"
"""

import json
import argparse
import re
import sys
from dataclasses import dataclass, asdict, field
from pathlib import Path
from typing import Optional, Dict, List, Any, Tuple
from enum import Enum


class MatchType(str, Enum):
    """Enumeration of match types in priority order."""
    EXACT = "exact"
    ALIAS_TIER1 = "alias_tier1"
    ALIAS_TIER2 = "alias_tier2"
    ALIAS_TIER3 = "alias_tier3"
    ALIAS_TIER4 = "alias_tier4"
    DOCUMENT_NAME = "document_name"
    DOCUMENT_ID = "document_id"
    FEE = "fee"
    PARAMETER_FALLBACK = "parameter_fallback"
    UNMATCHED = "unmatched"


class Entity(str, Enum):
    """Enumeration of entity types."""
    APPLICANT = "APPLICANT"
    COAPPLICANT = "COAPPLICANT"
    COAPPLICANT1 = "COAPPLICANT1"
    COAPPLICANT2 = "COAPPLICANT2"
    COAPPLICANT3 = "COAPPLICANT3"
    COAPPLICANT4 = "COAPPLICANT4"
    LOAN = "LOAN"
    DOCUMENT = "DOCUMENT"
    FEE = "FEE"
    OTHER = "OTHER"


@dataclass
class MatchResult:
    """Result of a single field matching operation."""
    partner_field: str
    column_category: Optional[str]
    matched_excel_key: Optional[str]
    matched_json_key: Optional[str]
    confidence: float
    match_type: str  # MatchType enum as string
    reasoning: str
    entity: str  # Entity enum as string

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization."""
        return asdict(self)


def normalize_process_name(process_name: Optional[str]) -> str:
    """Normalize process names for filtering reference candidates."""
    if not process_name:
        return ""
    normalized = process_name.strip().upper()
    aliases = {
        "ORIGINATION": "ORIGINATION",
        "ENROLLMENT": "ENROLLMENT",
        "ENROLMENT": "ENROLLMENT",
        "COMBINED": "",
        "SEPARATE": "",
    }
    return aliases.get(normalized, normalized)


def _is_excel_key_allowed_for_process(
    field_dictionary: Dict[str, Any],
    excel_key: str,
    process_name: Optional[str],
) -> bool:
    """Check whether an excel key is valid for the requested process."""
    normalized_process = normalize_process_name(process_name)
    if not normalized_process:
        return True

    entry = field_dictionary.get("by_excel_key", {}).get(excel_key, {})
    process_names = entry.get("process_names") or []
    if not process_names:
        return True
    return normalized_process in {str(name).upper() for name in process_names}


# Document detection keywords
DOCUMENT_KEYWORDS = {
    "aadhar", "aadhaar", "kyc", "photo", "photograph", "image",
    "bank statement", "bankstatement", "cibil", "credit bureau", "bureau",
    "sanction letter", "sanctionletter", "loan agreement", "loanagreement",
    "delivery receipt", "deliveryreceipt", "dpn", "demand promissory",
    "demandpromissory", "perfios", "vkyc", "video kyc", "tvr",
    "ownership proof", "ownershipproof", "income proof", "incomeproof",
    "pan card", "pancard", "voter id", "voterid", "passport",
    "driving license", "drivinglicense", "insurance",
    "invoice", "form16", "form 16", "itr", "income tax return",
    "gst certificate", "gstcertificate", "udyam", "udyamregistration",
    "shop act", "shopact", "rental agreement", "rentalagreement",
    "property doc", "property document", "noc", "no objection certificate",
    "cheque", "check", "nach", "mandate", "auto-debit",
    "vehicle", "automobile", "car", "bike", "two wheeler",
    "document"
}

FEE_KEYWORDS = {
    "fee", "fees", "charge", "charges", "commission", "premium",
    "processing", "interest", "emi"
}

DOCUMENT_LOCATOR_KEYWORDS = {"link", "url", "path", "id", "identifier"}

CATEGORY_STOPWORDS = {
    "lead", "details", "detail", "data", "info", "information", "section",
    "left", "right", "top", "bottom", "tab", "page", "field", "fields",
    "screen", "view", "module", "group", "grouping", "form",
    "api", "mapping",
}

ENTITY_CATEGORY_HINTS = {
    "COAPPLICANT": ("coapplicant", "coapp", "co-applicant"),
    "GUARANTOR": ("guarantor",),
    "FEE": ("fee", "fees", "charge", "charges", "commission"),
    "DOCUMENT": ("document", "documents", "upload", "attachment", "kyc", "proof"),
    "LOAN": ("loan", "disbursement", "repayment", "emi", "interest", "vehicle"),
    "APPLICANT": ("applicant", "customer", "borrower"),
}


def normalize_category(value: Optional[str]) -> str:
    """Normalize category/grouping strings for robust entity routing."""
    if not value:
        return ""
    normalized = value.strip().lower()
    normalized = normalized.replace("&", "and")
    normalized = re.sub(r"[\s_\-./()]+", "", normalized)
    return normalized


def _tokenize_text(value: Optional[str]) -> List[str]:
    """Tokenize free-form labels, including camelCase and separator-based text."""
    if not value:
        return []
    expanded = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", value)
    tokens = re.findall(r"[A-Za-z0-9]+", expanded.lower())
    return [t for t in tokens if t]


def _category_signal_tokens(column_category: Optional[str]) -> List[str]:
    """Extract meaningful tokens from a category label for candidate scoring."""
    tokens = []
    for token in _tokenize_text(column_category):
        if len(token) <= 2:
            continue
        if token in CATEGORY_STOPWORDS:
            continue
        tokens.append(token)
    return list(dict.fromkeys(tokens))


def _find_category_exact_override(
    field_name: str,
    column_category: Optional[str],
    field_dictionary: Dict[str, Any],
    entity: Optional[str] = None,
    process_name: Optional[str] = None,
) -> Optional[Tuple[str, str]]:
    """
    Find a better category-aligned match for generic exact fields.

    Example:
    - field_name: RegistrationNumber
    - category: LeadVehicle
    - generic exact hit: REGISTRATIONNUMBER
    - category override: VEHICLEREGISTRATIONNUMBER
    """
    category_tokens = _category_signal_tokens(column_category)
    if not category_tokens:
        return None

    basic_normalized = normalize_basic(field_name)
    if not basic_normalized:
        return None

    by_excel_key = field_dictionary.get("by_excel_key", {})
    scored_candidates: List[Tuple[int, str, str]] = []

    for excel_key, info in by_excel_key.items():
        if not _is_excel_key_allowed_for_process(field_dictionary, excel_key, process_name):
            continue
        normalized_excel = normalize_basic(excel_key)
        if normalized_excel == basic_normalized:
            continue

        if not (
            normalized_excel.endswith(basic_normalized)
            or basic_normalized in normalized_excel
        ):
            continue

        json_key = info.get("json_key") or ""
        description = info.get("description") or ""
        candidate_tokens = set(_tokenize_text(" ".join([excel_key, json_key, description])))
        entity_upper = (entity or "").upper()
        entity_score = 0
        candidate_text = " ".join([excel_key, json_key, description]).lower().replace(" ", "")

        if entity_upper in {"APPLICANT", "CUSTOMER"}:
            if "applicant" in candidate_text or "customer" in candidate_text:
                entity_score = 2
        elif entity_upper.startswith("COAPPLICANT"):
            aliases = ["coapplicant"]
            match = re.search(r"COAPPLICANT(\d+)", entity_upper)
            if match:
                idx = int(match.group(1))
                aliases.append(f"coapplicant{idx}")
                if idx > 0:
                    aliases.append(f"coapplicant{idx - 1}")
            if any(alias in candidate_text for alias in aliases):
                entity_score = 2
        elif entity_upper.startswith("GUARANTOR"):
            if "guarantor" in candidate_text:
                entity_score = 2

        score = 0
        if normalized_excel.endswith(basic_normalized):
            score += 3
        elif basic_normalized in normalized_excel:
            score += 1

        overlap = [token for token in category_tokens if token in candidate_tokens]
        # If category is generic/non-informative, only allow a fallback override
        # when the candidate clearly aligns with the detected entity and the
        # field name is a suffix match (e.g. APPLICANTDATEOFBIRTH).
        if not overlap:
            if entity_score <= 0 or not normalized_excel.endswith(basic_normalized):
                continue

        score += len(overlap) * 2
        score += entity_score

        # Prefer candidates where longer/more specific category tokens align.
        score += sum(1 for token in overlap if len(token) >= 6)

        if score > 0:
            scored_candidates.append((score, excel_key, json_key))

    if not scored_candidates:
        return None

    scored_candidates.sort(key=lambda item: item[0], reverse=True)
    _, excel_key, json_key = scored_candidates[0]
    return excel_key, json_key


def _score_category_alignment(
    column_category: Optional[str],
    excel_key: str,
    json_key: Optional[str] = "",
    description: Optional[str] = "",
) -> int:
    """Score how well a candidate aligns with the input category."""
    category_tokens = _category_signal_tokens(column_category)
    if not category_tokens:
        return 0

    candidate_tokens = set(_tokenize_text(" ".join([excel_key, json_key or "", description or ""])))
    overlap = [token for token in category_tokens if token in candidate_tokens]

    score = len(overlap) * 2
    score += sum(1 for token in overlap if len(token) >= 6)
    return score


def _find_category_alias_override(
    normalized_field: str,
    column_category: Optional[str],
    refs: Dict[str, Any],
    default_excel_key: str,
    process_name: Optional[str] = None,
) -> Optional[Tuple[str, str]]:
    """
    For alias matches, prefer another historically seen alias target when the
    category strongly favors it over the default target.
    """
    category_tokens = _category_signal_tokens(column_category)
    if not category_tokens:
        return None

    alias_registry = refs.get("alias_registry", {})
    by_excel_key = refs.get("field_dictionary", {}).get("by_excel_key", {})
    reverse_aliases = alias_registry.get("reverse", {})

    default_info = by_excel_key.get(default_excel_key, {})
    default_score = _score_category_alignment(
        column_category,
        default_excel_key,
        default_info.get("json_key", ""),
        default_info.get("description", ""),
    )

    candidate_keys = set()
    reverse_entry = reverse_aliases.get(default_excel_key, {})
    for alias_value in reverse_entry.get("aliases", []):
        alias_entry = alias_registry.get("forward", {}).get(alias_value, {})
        target_excel_key = alias_entry.get("target_excel_key", "")
        if target_excel_key:
            candidate_keys.add(target_excel_key)

    # Also consider all targets from aliases that share the same normalized input.
    shared_alias_entry = alias_registry.get("forward", {}).get(normalized_field, {})
    if shared_alias_entry.get("target_excel_key"):
        candidate_keys.add(shared_alias_entry["target_excel_key"])

    best_choice = None
    best_score = default_score
    for candidate_excel_key in candidate_keys:
        if not _is_excel_key_allowed_for_process(refs.get("field_dictionary", {}), candidate_excel_key, process_name):
            continue
        info = by_excel_key.get(candidate_excel_key, {})
        score = _score_category_alignment(
            column_category,
            candidate_excel_key,
            info.get("json_key", ""),
            info.get("description", ""),
        )
        if score > best_score:
            best_score = score
            best_choice = (
                candidate_excel_key,
                info.get("json_key", "") or "",
            )

    return best_choice


def normalize_basic(field_name: str) -> str:
    """
    Basic normalization: lowercase + strip separators. Preserves entity prefixes.
    Used for indexing field_dictionary keys where APPLICANT/COAPPLICANT distinction matters.

    Args:
        field_name: Raw field name

    Returns:
        Normalized field name with entity prefixes preserved

    Examples:
        >>> normalize_basic("APPLICANT_DATE_OF_BIRTH")
        "applicantdateofbirth"
        >>> normalize_basic("COAPPLICANT1_GENDER")
        "coapplicant1gender"
    """
    if not field_name:
        return ""
    normalized = field_name.lower()
    normalized = re.sub(r'[_\s\.\-\'\(\)]', '', normalized)
    return normalized


def normalize_field(field_name: str) -> str:
    """
    Normalize a partner field name for alias matching.

    Steps:
    1. Convert to lowercase
    2. Strip whitespace, underscores, dots, hyphens, apostrophes
    3. Remove common entity prefixes (applicant_, coapplicant_, loan_, etc.)
       This allows 'applicantGender' to match the same alias as 'gender'.

    Args:
        field_name: Raw field name from partner

    Returns:
        Normalized field name with entity prefixes stripped

    Examples:
        >>> normalize_field("Date_of_Birth")
        "dateofbirth"
        >>> normalize_field("Applicant.Gender")
        "gender"
        >>> normalize_field("CoApplicant - Annual Income")
        "annualincome"
    """
    if not field_name:
        return ""

    # Lowercase
    normalized = field_name.lower()

    # Remove common characters
    normalized = re.sub(r'[_\s\.\-\'\(\)]', '', normalized)

    # Remove common entity prefixes
    prefixes = [
        r"^applicant",
        r"^coapplicant\d*",
        r"^coapplicant",
        r"^loan",
        r"^document",
        r"^fee",
        r"^partner",
    ]

    for prefix in prefixes:
        normalized = re.sub(prefix, '', normalized, flags=re.IGNORECASE)

    return normalized


def _canonicalize_alias_key(value: Optional[str]) -> str:
    """
    Collapse small naming variations so process-aware alias fallback can find
    equivalent keys like `aadharNumber` vs `aadharNo`.
    """
    normalized = normalize_field(value or "")
    if not normalized:
        return ""
    normalized = normalized.replace("aadhaar", "aadhar")
    normalized = normalized.replace("number", "no")
    normalized = normalized.replace("num", "no")
    return normalized


def _canonicalize_json_key(value: Optional[str]) -> str:
    """
    Normalize JSON paths so equivalent fields across process-specific object
    roots can still be matched generically.
    """
    if not value:
        return ""
    normalized = str(value).strip().lower()
    normalized = re.sub(r"\.\d+\.", ".", normalized)
    normalized = re.sub(r"\.\d+$", "", normalized)
    for prefix in (
        "loanaccount.",
        "customer.",
        "primary.",
        "secondary.",
        "applicant.",
        "coapplicant.",
    ):
        if normalized.startswith(prefix):
            normalized = normalized[len(prefix):]
    normalized = normalized.replace("aadhaar", "aadhar")
    normalized = re.sub(r"[^a-z0-9]+", "", normalized)
    return normalized


def _get_process_entries(
    field_dictionary: Dict[str, Any],
    process_name: Optional[str],
) -> List[Dict[str, Any]]:
    normalized_process = normalize_process_name(process_name)
    if not normalized_process:
        return []
    return field_dictionary.get("by_process", {}).get(normalized_process, [])


def _find_process_equivalent_by_json(
    refs: Dict[str, Any],
    target_json_key: Optional[str],
    entity: Optional[str],
    column_category: Optional[str],
    process_name: Optional[str],
) -> Optional[Tuple[str, str]]:
    """
    Find a process-valid equivalent by comparing canonicalized JSON meaning,
    not just the excel key name.
    """
    field_dictionary = refs.get("field_dictionary", {})
    canonical_json = _canonicalize_json_key(target_json_key)
    if not canonical_json:
        return None

    candidates: List[Tuple[str, str, str]] = []
    for entry in _get_process_entries(field_dictionary, process_name):
        excel_key = entry.get("excel_key", "")
        json_key = entry.get("json_key", "")
        role = entry.get("role", "")
        if not excel_key or not json_key:
            continue
        if _canonicalize_json_key(json_key) != canonical_json:
            continue
        candidates.append((excel_key, json_key, role))

    if not candidates:
        return None

    return _select_best_candidate(
        candidates,
        entity or "OTHER",
        column_category,
        process_name,
        field_dictionary,
    )


def _find_process_aware_alias_fallback(
    normalized_field: str,
    refs: Dict[str, Any],
    process_name: Optional[str],
) -> Optional[Tuple[str, str]]:
    """
    If the primary alias target belongs to a different process, look for another
    alias with the same canonical meaning whose target is allowed for the
    requested process.
    """
    normalized_process = normalize_process_name(process_name)
    if not normalized_process:
        return None

    field_dictionary = refs.get("field_dictionary", {})
    forward_aliases = refs.get("alias_registry", {}).get("forward", {})
    wanted = _canonicalize_alias_key(normalized_field)
    if not wanted:
        return None

    candidates: List[Tuple[int, str, str]] = []
    for alias_key, alias_entry in forward_aliases.items():
        if _canonicalize_alias_key(alias_key) != wanted:
            continue
        target_excel_key = alias_entry.get("target_excel_key", "")
        if not target_excel_key:
            continue
        if not _is_excel_key_allowed_for_process(field_dictionary, target_excel_key, normalized_process):
            continue
        target_json_key = alias_entry.get("target_json_key", "") or ""
        score = alias_entry.get("frequency", 0)
        if normalize_field(alias_key) == normalize_field(normalized_field):
            score += 5
        candidates.append((score, target_excel_key, target_json_key))

    if not candidates:
        return None

    candidates.sort(key=lambda item: item[0], reverse=True)
    _, excel_key, json_key = candidates[0]
    return excel_key, json_key


def detect_entity(
    column_category: Optional[str],
    field_name: str,
    entity_routing: Dict[str, str],
    process_name: Optional[str] = None
) -> str:
    """
    Detect entity from column_category, field name patterns, or process context.

    Priority order:
    1. Lookup column_category in entity_routing
    2. Check field_name for entity patterns (applicant, coapplicant, loan, etc.)
    3. Use process_name context if available
    4. Default to OTHER

    Args:
        column_category: UI grouping/column category
        field_name: Field name (raw)
        entity_routing: Mapping from grouping_to_entity
        process_name: Optional process or context name

    Returns:
        Entity as string (e.g., "APPLICANT", "COAPPLICANT", "LOAN")
    """
    # Route by column_category first
    if column_category:
        if column_category in entity_routing:
            return entity_routing[column_category]

        normalized_category = normalize_category(column_category)
        normalized_routing = {
            normalize_category(key): value for key, value in entity_routing.items()
        }
        if normalized_category in normalized_routing:
            return normalized_routing[normalized_category]

        # Heuristic fallback for partner-provided human-readable categories.
        for entity_name, hints in ENTITY_CATEGORY_HINTS.items():
            if any(hint in normalized_category for hint in hints):
                return entity_name

    # Check field name patterns
    field_lower = field_name.lower()

    if any(x in field_lower for x in ["guarantor", "guaranter", "guaraontor"]):
        # Try to detect guarantor number
        match = re.search(r'guarantor(\d)', field_lower)
        if match:
            num = int(match.group(1))
            return f"GUARANTOR{num}"
        return "GUARANTOR"

    if any(x in field_lower for x in ["coapplicant", "co_applicant", "co-applicant", "coapp"]):
        # Try to detect coapplicant number
        match = re.search(r'(?:coapplicant|coapp)(\d)', field_lower)
        if match:
            num = int(match.group(1))
            return f"COAPPLICANT{num}"
        return "COAPPLICANT"

    if any(x in field_lower for x in ["applicant", "customer", "borrower"]):
        return "APPLICANT"

    if any(x in field_lower for x in ["loan"]):
        return "LOAN"

    if any(x in field_lower for x in ["document", "upload", "file", "attachment"]):
        return "DOCUMENT"

    if any(x in field_lower for x in ["fee", "charge", "commission"]):
        return "FEE"

    # Fallback to process_name context
    if process_name:
        return entity_routing.get(process_name, "APPLICANT")

    # Default to APPLICANT — most partner fields without explicit entity
    # markers are applicant-level (customer data). This prevents entity-aware
    # matching from deprioritizing applicant fields.
    return "APPLICANT"


def load_references(references_dir: str) -> Dict[str, Any]:
    """
    Load all reference files from the references directory into memory.

    Args:
        references_dir: Path to references directory

    Returns:
        Dictionary with keys: field_dictionary, alias_registry, entity_routing

    Raises:
        FileNotFoundError: If any reference file is missing
        json.JSONDecodeError: If reference file is invalid JSON
    """
    ref_path = Path(references_dir)

    required_files = {
        "field_dictionary.json": "field_dictionary",
        "alias_registry.json": "alias_registry",
        "entity_routing.json": "entity_routing"
    }

    references = {}

    for filename, key in required_files.items():
        file_path = ref_path / filename

        if not file_path.exists():
            raise FileNotFoundError(
                f"Reference file not found: {file_path}\n"
                f"Expected in: {references_dir}"
            )

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                references[key] = json.load(f)
        except json.JSONDecodeError as e:
            raise json.JSONDecodeError(
                f"Invalid JSON in {filename}: {e.msg}",
                e.doc,
                e.pos
            )

    return references


def _build_field_dict_index(field_dictionary: Dict[str, Any]) -> Dict[str, List[Tuple[str, str, str]]]:
    """
    Build a normalized index of field_dictionary for fast lookup.

    Returns mapping: normalized_excel_key -> list of (excel_key, json_key, role)
    Multiple entries per normalized key allows entity-aware selection.

    Args:
        field_dictionary: Loaded field_dictionary.json

    Returns:
        Index dictionary with lists of candidates per normalized key
    """
    index: Dict[str, List[Tuple[str, str, str]]] = {}

    if "by_role" in field_dictionary:
        for role, entries in field_dictionary.get("by_role", {}).items():
            if isinstance(entries, list):
                for entry in entries:
                    excel_key = entry.get("excel_key", "")
                    json_key = entry.get("json_key", "")
                    if excel_key:
                        normalized = normalize_basic(excel_key)
                        if normalized not in index:
                            index[normalized] = []
                        index[normalized].append((excel_key, json_key, role))

    return index


def _build_process_field_dict_index(
    field_dictionary: Dict[str, Any],
    process_name: Optional[str],
) -> Dict[str, List[Tuple[str, str, str]]]:
    """
    Build a process-scoped normalized index from field_dictionary.by_process.

    This preserves the old matching flow but narrows exact-match candidates to
    the requested process when process context is available.
    """
    normalized_process = normalize_process_name(process_name)
    if not normalized_process:
        return {}

    index: Dict[str, List[Tuple[str, str, str]]] = {}
    process_entries = field_dictionary.get("by_process", {}).get(normalized_process, [])

    for entry in process_entries:
        excel_key = entry.get("excel_key", "")
        json_key = entry.get("json_key", "")
        role = entry.get("role", "")
        if not excel_key:
            continue
        normalized = normalize_basic(excel_key)
        index.setdefault(normalized, []).append((excel_key, json_key, role))

    return index


def _select_best_candidate(
    candidates: List[Tuple[str, str, str]],
    entity: str,
    column_category: Optional[str] = None,
    process_name: Optional[str] = None,
    field_dictionary: Optional[Dict[str, Any]] = None,
) -> Tuple[str, str]:
    """
    Select the best candidate from multiple matches based on entity context.

    Priority:
    - If entity is APPLICANT/OTHER: prefer CUSTOMER role, then LOAN
    - If entity is COAPPLICANT*: prefer COAPPLICANT role
    - If entity is LOAN: prefer LOAN role
    - If entity is GUARANTOR: prefer GUARANTOR role
    - Default: first candidate

    Args:
        candidates: List of (excel_key, json_key, role) tuples
        entity: Current entity context

    Returns:
        (excel_key, json_key) of best candidate
    """
    if field_dictionary:
        filtered_candidates = [
            (ek, jk, role)
            for ek, jk, role in candidates
            if _is_excel_key_allowed_for_process(field_dictionary, ek, process_name)
        ]
        if filtered_candidates:
            candidates = filtered_candidates

    if len(candidates) == 1:
        return candidates[0][0], candidates[0][1]

    normalized_category = normalize_category(column_category)

    # Category-aware tiebreaker for domains where the same normalized field
    # can exist in multiple roles, such as registration numbers.
    if normalized_category:
        category_tokens = set(_category_signal_tokens(column_category))
        if category_tokens:
            for ek, jk, role in candidates:
                candidate_tokens = set(_tokenize_text(" ".join([ek, jk or ""])))
                if category_tokens & candidate_tokens:
                    return ek, jk

    # Define role preference by entity
    entity_upper = entity.upper() if entity else "OTHER"

    if entity_upper in ("APPLICANT", "OTHER"):
        role_priority = ["CUSTOMER", "LOAN", "COAPPLICANT", "GUARANTOR"]
    elif entity_upper.startswith("COAPPLICANT"):
        role_priority = ["COAPPLICANT", "CUSTOMER", "LOAN", "GUARANTOR"]
    elif entity_upper == "LOAN":
        role_priority = ["LOAN", "CUSTOMER", "COAPPLICANT", "GUARANTOR"]
    elif entity_upper == "GUARANTOR":
        role_priority = ["GUARANTOR", "CUSTOMER", "LOAN", "COAPPLICANT"]
    else:
        role_priority = ["CUSTOMER", "LOAN", "COAPPLICANT", "GUARANTOR"]

    # Sort candidates by role priority
    for preferred_role in role_priority:
        for ek, jk, role in candidates:
            if role == preferred_role:
                return ek, jk

    # Fallback to first candidate
    return candidates[0][0], candidates[0][1]


def _get_alias_tier(frequency: int) -> str:
    """
    Determine confidence tier based on alias frequency.

    Tiers:
    - tier1: 30+ partners → confidence 0.98
    - tier2: 10-29 partners → confidence 0.92
    - tier3: 3-9 partners → confidence 0.85
    - tier4: 1-2 partners → confidence 0.75

    Args:
        frequency: Number of partners using this alias

    Returns:
        Tier string (tier1, tier2, tier3, tier4)
    """
    if frequency >= 30:
        return "tier1"
    elif frequency >= 10:
        return "tier2"
    elif frequency >= 3:
        return "tier3"
    else:
        return "tier4"


def _confidence_for_tier(tier: str) -> float:
    """Get confidence score for a tier."""
    tier_scores = {
        "tier1": 0.98,
        "tier2": 0.92,
        "tier3": 0.85,
        "tier4": 0.75,
    }
    return tier_scores.get(tier, 0.50)


def _try_entity_substitution(
    excel_key: str,
    current_entity: str,
    field_dictionary_index: Dict[str, List[Tuple[str, str, str]]],
    field_dictionary: Dict[str, Any],
    process_name: Optional[str] = None,
) -> Tuple[Optional[str], float]:
    """
    Try to substitute entity prefixes when alias suggests different entity.

    If alias maps to APPLICANTCUSTOMERGENDER but we're in COAPPLICANT context,
    try COAPPLICANT1CUSTOMERGENDER, etc.

    Args:
        excel_key: Original excel_key from alias
        current_entity: Current entity context
        field_dictionary_index: Field dictionary index

    Returns:
        Tuple of (substituted_excel_key, confidence_adjustment) or (None, 0.0)
    """
    # Attempt substitution for coapplicant and guarantor scenarios
    if not (current_entity.startswith("COAPPLICANT") or current_entity.startswith("GUARANTOR")):
        return None, 0.0

    # Check if original key has APPLICANT prefix
    if "APPLICANT" not in excel_key.upper():
        return None, 0.0

    # Try to substitute APPLICANT with current entity
    # For COAPPLICANT: APPLICANTGENDER → COAPPLICANT1GENDER
    # For GUARANTOR: APPLICANTGENDER → GUARANTOR1GENDER
    target_prefix = current_entity
    if current_entity == "COAPPLICANT":
        target_prefix = "COAPPLICANT1"
    elif current_entity == "GUARANTOR":
        target_prefix = "GUARANTOR1"

    substituted = re.sub(
        r"(?i)applicant(?!\d)",
        target_prefix,
        excel_key
    )

    # Check if substituted key exists in field dictionary
    normalized_sub = normalize_basic(substituted)
    if normalized_sub in field_dictionary_index:
        # Use the best candidate from substituted results
        best_ek, _ = _select_best_candidate(
            field_dictionary_index[normalized_sub],
            current_entity,
            None,
            process_name,
            field_dictionary,
        )
        # Reduce confidence for entity substitution
        return best_ek, -0.10

    return None, 0.0


def match_field(
    field_name: str,
    column_category: Optional[str],
    entity: Optional[str],
    refs: Dict[str, Any],
    process_name: Optional[str] = None,
) -> MatchResult:
    """
    Match a single partner field to an internal excel_key.

    Evaluation order:
    1. EXACT MATCH: normalized field matches normalized excel_key
    2. ALIAS MATCH (by tier): lookup in alias_registry.forward
    3. DOCUMENT NAME DETECTION: field contains document keywords
    4. DOCUMENT ID DETECTION: field contains document + locator keywords
    5. FEE DETECTION: field contains fee/charge keywords
    6. UNMATCHED: return for LLM layer to handle

    Args:
        field_name: Partner field name
        column_category: UI grouping/column category
        entity: Entity type (APPLICANT, COAPPLICANT, LOAN, etc.)
        refs: References dictionary from load_references()
        process_name: Optional process context used to filter candidate excel keys

    Returns:
        MatchResult with confidence, match_type, and reasoning
    """
    # Default entity detection
    if not entity:
        entity = detect_entity(
            column_category,
            field_name,
            refs["entity_routing"].get("grouping_to_entity", {})
        )

    normalized_field = normalize_field(field_name)
    field_dictionary = refs.get("field_dictionary", {})
    alias_registry = refs.get("alias_registry", {})

    # Build field dictionary index for fast lookup
    field_dict_index = _build_field_dict_index(field_dictionary)
    process_field_dict_index = _build_process_field_dict_index(field_dictionary, process_name)

    # Category-aware override before generic exact matching. This helps when
    # a partner field is generic by itself, but the category clearly narrows
    # the domain (vehicle, coapplicant, fee, document, loan, etc.).
    category_override = _find_category_exact_override(
        field_name,
        column_category,
        field_dictionary,
        entity,
        process_name,
    )
    if category_override:
        override_excel_key, override_json_key = category_override
        if override_excel_key:
            return MatchResult(
                partner_field=field_name,
                column_category=column_category,
                matched_excel_key=override_excel_key,
                matched_json_key=override_json_key,
                confidence=0.98,
                match_type=MatchType.EXACT.value,
                reasoning=(
                    f"Category-aware exact override: '{field_name}' in category '{column_category}' "
                    f"maps to '{override_excel_key}'"
                ),
                entity=entity,
            )

    # ====== 1. EXACT MATCH ======
    # Try basic normalization first (preserves entity prefixes for precise matching)
    basic_normalized = normalize_basic(field_name)
    if basic_normalized in process_field_dict_index:
        candidates = process_field_dict_index[basic_normalized]
        excel_key, json_key = _select_best_candidate(
            candidates,
            entity,
            column_category,
            process_name,
            field_dictionary,
        )
        return MatchResult(
            partner_field=field_name,
            column_category=column_category,
            matched_excel_key=excel_key,
            matched_json_key=json_key,
            confidence=1.0,
            match_type=MatchType.EXACT.value,
            reasoning=(
                f"Exact match within process '{normalize_process_name(process_name)}': "
                f"normalized '{field_name}' matches '{excel_key}'"
            ),
            entity=entity
        )

    if basic_normalized in field_dict_index:
        candidates = field_dict_index[basic_normalized]
        excel_key, json_key = _select_best_candidate(
            candidates,
            entity,
            column_category,
            process_name,
            field_dictionary,
        )
        if process_name and not _is_excel_key_allowed_for_process(field_dictionary, excel_key, process_name):
            equivalent = _find_process_equivalent_by_json(
                refs,
                json_key,
                entity,
                column_category,
                process_name,
            )
            if equivalent:
                excel_key, json_key = equivalent
        return MatchResult(
            partner_field=field_name,
            column_category=column_category,
            matched_excel_key=excel_key,
            matched_json_key=json_key,
            confidence=1.0,
            match_type=MatchType.EXACT.value,
            reasoning=f"Exact match: normalized '{field_name}' matches '{excel_key}'",
            entity=entity
        )

    # ====== 2. ALIAS MATCH (by tier) ======
    forward_aliases = alias_registry.get("forward", {})

    if normalized_field in forward_aliases:
        alias_entry = forward_aliases[normalized_field]
        target_excel_key = alias_entry.get("target_excel_key", "")
        target_json_key = alias_entry.get("target_json_key", "")
        frequency = alias_entry.get("frequency", 0)

        # Determine tier and confidence
        tier = _get_alias_tier(frequency)
        match_type = f"alias_{tier}"
        base_confidence = _confidence_for_tier(tier)

        # Try entity substitution if applicable
        substituted_key, conf_adjustment = _try_entity_substitution(
            target_excel_key,
            entity,
            field_dict_index,
            field_dictionary,
            process_name,
        )

        final_excel_key = substituted_key or target_excel_key
        final_json_key = target_json_key
        final_confidence = base_confidence + conf_adjustment

        # Category-aware alias override for generic fields in informative categories.
        alias_override = _find_category_alias_override(
            normalized_field,
            column_category,
            refs,
            final_excel_key,
            process_name,
        )
        if alias_override:
            override_excel_key, override_json_key = alias_override
            if override_excel_key and override_excel_key != final_excel_key:
                final_excel_key = override_excel_key
                final_json_key = override_json_key
                final_confidence = min(0.99, final_confidence + 0.02)

        elif not _is_excel_key_allowed_for_process(field_dictionary, final_excel_key, process_name):
            equivalent = _find_process_equivalent_by_json(
                refs,
                final_json_key,
                entity,
                column_category,
                process_name,
            )
            if equivalent:
                final_excel_key, final_json_key = equivalent
                final_confidence = min(final_confidence, 0.86)
                return MatchResult(
                    partner_field=field_name,
                    column_category=column_category,
                    matched_excel_key=final_excel_key,
                    matched_json_key=final_json_key,
                    confidence=final_confidence,
                    match_type=match_type,
                    reasoning=(
                        f"Alias match ({tier}) via process-equivalent JSON mapping: "
                        f"'{field_name}' resolved to '{final_excel_key}' for process "
                        f"'{process_name}'"
                    ),
                    entity=entity,
                )

            fallback = _find_process_aware_alias_fallback(
                normalized_field,
                refs,
                process_name,
            )
            if fallback:
                final_excel_key, final_json_key = fallback
                final_confidence = min(final_confidence, 0.84)
                return MatchResult(
                    partner_field=field_name,
                    column_category=column_category,
                    matched_excel_key=final_excel_key,
                    matched_json_key=final_json_key,
                    confidence=final_confidence,
                    match_type=match_type,
                    reasoning=(
                        f"Alias match ({tier}) via process-aware fallback: "
                        f"'{field_name}' matched alternate target '{final_excel_key}' "
                        f"for process '{process_name}'"
                    ),
                    entity=entity,
                )

            # Skip alias targets that belong to a different process; let later
            # pattern rules or the LLM handle the field instead of cross-mapping.
            return MatchResult(
                partner_field=field_name,
                column_category=column_category,
                matched_excel_key=None,
                matched_json_key=None,
                confidence=0.0,
                match_type=MatchType.UNMATCHED.value,
                reasoning=(
                    f"Alias target '{final_excel_key}' is not allowed for process "
                    f"'{process_name}'"
                ),
                entity=entity,
            )

        return MatchResult(
            partner_field=field_name,
            column_category=column_category,
            matched_excel_key=final_excel_key,
            matched_json_key=final_json_key,
            confidence=final_confidence,
            match_type=match_type,
            reasoning=(
                f"Alias match ({tier}): '{field_name}' → '{target_excel_key}' "
                f"(frequency={frequency})" +
                (f"; entity substitution applied: {substituted_key}" if substituted_key else "") +
                (
                    f"; category override applied: {final_excel_key}"
                    if alias_override and final_excel_key != (substituted_key or target_excel_key)
                    else ""
                )
            ),
            entity=entity
        )

    # ====== 3. DOCUMENT NAME DETECTION ======
    field_lower = field_name.lower()
    # Remove ambiguous single-word matches: require the doc keyword to be a significant
    # part of the field name, not just incidental (e.g., "KYC" in "APPLICANT'S NAME (SAME AS KYC)")
    doc_keywords_found = [kw for kw in DOCUMENT_KEYWORDS if kw in field_lower]
    # A field is a document reference if doc keywords make up significant part of the name
    # or the field clearly references a file/link/upload
    has_strong_doc_signal = any(
        kw in field_lower for kw in [
            "document", "upload", "attachment", "file", "certificate",
            "proof", "letter", "agreement", "receipt", "statement",
            "xml", "pdf", "image", "photo", "photograph"
        ]
    )
    has_doc_keyword = len(doc_keywords_found) > 0 and has_strong_doc_signal
    looks_like_filename = any(
        x in field_lower for x in [".pdf", ".jpg", ".png", "file", "upload", "imagename"]
    ) or (
        # camelCase ending in Name/File/Image suggests doc naming
        any(field_name.endswith(s) for s in ["Name", "File", "Image", "FileName", "ImageName"])
        and has_strong_doc_signal
    )

    if has_doc_keyword and looks_like_filename:
        return MatchResult(
            partner_field=field_name,
            column_category=column_category,
            matched_excel_key="DOCUMENTNAME",
            matched_json_key=None,
            confidence=0.90,
            match_type=MatchType.DOCUMENT_NAME.value,
            reasoning=(
                f"Document name detection: field contains document keywords "
                f"({', '.join(kw for kw in DOCUMENT_KEYWORDS if kw in field_lower)}) "
                f"and looks like a file/name"
            ),
            entity="DOCUMENT"
        )

    # ====== 4. DOCUMENT ID DETECTION ======
    has_locator_keyword = any(kw in field_lower for kw in DOCUMENT_LOCATOR_KEYWORDS)

    if has_doc_keyword and has_locator_keyword:
        return MatchResult(
            partner_field=field_name,
            column_category=column_category,
            matched_excel_key="DOCUMENTID",
            matched_json_key=None,
            confidence=0.90,
            match_type=MatchType.DOCUMENT_ID.value,
            reasoning=(
                f"Document ID detection: field contains document keyword "
                f"and locator keyword ({', '.join(kw for kw in DOCUMENT_LOCATOR_KEYWORDS if kw in field_lower)})"
            ),
            entity="DOCUMENT"
        )

    # ====== 5. FEE DETECTION ======
    has_fee_keyword = any(kw in field_lower for kw in FEE_KEYWORDS)

    if has_fee_keyword:
        return MatchResult(
            partner_field=field_name,
            column_category=column_category,
            matched_excel_key="FEE",
            matched_json_key=None,
            confidence=0.90,
            match_type=MatchType.FEE.value,
            reasoning=(
                f"Fee detection: field contains fee/charge keywords "
                f"({', '.join(kw for kw in FEE_KEYWORDS if kw in field_lower)})"
            ),
            entity="FEE"
        )

    # ====== 6. UNMATCHED ======
    return MatchResult(
        partner_field=field_name,
        column_category=column_category,
        matched_excel_key=None,
        matched_json_key=None,
        confidence=0.0,
        match_type=MatchType.UNMATCHED.value,
        reasoning="No deterministic match found; requires LLM layer for semantic matching",
        entity=entity
    )


def match_batch(
    fields: List[Dict[str, Any]],
    refs: Dict[str, Any],
    process_name: Optional[str] = None
) -> Dict[str, List[MatchResult]]:
    """
    Match a batch of fields and separate into deterministic and unmatched.

    Args:
        fields: List of field dicts with keys: field_name, column_category (optional), entity (optional)
        refs: References dictionary from load_references()
        process_name: Optional process or context name for entity detection

    Returns:
        Dictionary with keys: "matched" (deterministic) and "unmatched" (needs LLM)
    """
    matched = []
    unmatched = []

    for field_dict in fields:
        field_name = field_dict.get("field_name", "")
        column_category = field_dict.get("column_category")
        entity = field_dict.get("entity")

        if not field_name:
            continue

        result = match_field(
            field_name,
            column_category,
            entity,
            refs,
            process_name=process_name,
        )

        if result.match_type == MatchType.UNMATCHED.value:
            unmatched.append(result)
        else:
            matched.append(result)

    return {
        "matched": matched,
        "unmatched": unmatched
    }


def main():
    """CLI interface for testing the matching engine."""
    parser = argparse.ArgumentParser(
        description="Deterministic Field Matching Engine",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python matching_engine.py --field "dateOfBirth" --category "Customer Details"
  python matching_engine.py --field "processingFee" --category "Loan Details"
  python matching_engine.py --field "aadharFrontLink" --category "Documents"
  python matching_engine.py --field "annualIncome" --entity "COAPPLICANT"
        """
    )

    parser.add_argument(
        "--field",
        "-f",
        type=str,
        required=True,
        help="Partner field name to match"
    )
    parser.add_argument(
        "--category",
        "-c",
        type=str,
        default=None,
        help="Column category/UI grouping (optional)"
    )
    parser.add_argument(
        "--entity",
        "-e",
        type=str,
        default=None,
        help="Entity type: APPLICANT, COAPPLICANT, LOAN, DOCUMENT, FEE, OTHER (optional)"
    )
    parser.add_argument(
        "--refs",
        "-r",
        type=str,
        default="./references",
        help="Path to references directory (default: ./references)"
    )
    parser.add_argument(
        "--json",
        "-j",
        action="store_true",
        help="Output as JSON"
    )

    args = parser.parse_args()

    try:
        # Load references
        refs = load_references(args.refs)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"ERROR loading references: {e}", file=sys.stderr)
        sys.exit(1)

    # Perform match
    result = match_field(
        field_name=args.field,
        column_category=args.category,
        entity=args.entity,
        refs=refs
    )

    # Output
    if args.json:
        print(json.dumps(result.to_dict(), indent=2))
    else:
        print(f"Partner Field: {result.partner_field}")
        print(f"Entity: {result.entity}")
        if result.column_category:
            print(f"Category: {result.column_category}")
        print(f"Match Type: {result.match_type}")
        print(f"Matched Excel Key: {result.matched_excel_key or '(none)'}")
        if result.matched_json_key:
            print(f"Matched JSON Key: {result.matched_json_key}")
        print(f"Confidence: {result.confidence:.2f}")
        print(f"Reasoning: {result.reasoning}")


if __name__ == "__main__":
    main()
