#!/usr/bin/env python3  
 # matching
"""
Deterministic Field Matching Engine

Core matching layer that maps partner fields to internal excel_keys without using an LLM.
Loads reference files (field_dictionary, alias_registry, entity_routing) and evaluates
fields in strict priority order: fee/charge semantics (entity FEE, excel_key FEE) →
exact match (raw) → exact match (prefix-stripped) →
alias tiers (entity-aware for COAPPLICANT vs applicant-scoped PUTM keys) → document
detection → document ID detection → unmatched.

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
    partner_field: str
    column_category: Optional[str]
    matched_excel_key: Optional[str]
    matched_json_key: Optional[str]
    confidence: float
    match_type: str
    reasoning: str
    entity: str

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


def normalize_process_name(process_name: Optional[str]) -> str:
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
    normalized_process = normalize_process_name(process_name)
    if not normalized_process:
        return True
    entry = field_dictionary.get("by_excel_key", {}).get(excel_key, {})
    process_names = entry.get("process_names") or []
    if not process_names:
        return True
    return normalized_process in {str(name).upper() for name in process_names}


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

# Do not include generic loan terms like "emi" or "interest" alone — they match
# schedule/pricing columns (e.g. "1ST EMI DATE", "INTEREST RATE") and wrongly
# force entity=FEE before alias / semantic matching can run.
# Substrings for fee/charge semantics. Do not include bare "processing" or "premium":
# they match loan fields like processingFeePercentage and insurance premium amounts.
FEE_KEYWORDS = {
    "fee",
    "fees",
    "charge",
    "charges",
    "commission",
}

DOCUMENT_LOCATOR_KEYWORDS = {"link", "url", "path", "id", "identifier"}

# Substrings that indicate a credential field (password/PIN for a link, etc.). These must
# not be classified as DOCUMENTID: they satisfy document + locator ("link") but are not IDs.
CREDENTIAL_KEYWORD_HINTS = frozenset({"password", "passphrase", "passcode"})

# Alias forward keys must be at least this long to qualify for subsequence bridge (avoid noise).
_CREDENTIAL_ALIAS_BRIDGE_MIN_KEY_LEN = 10


def _is_subsequence(needle: str, haystack: str) -> bool:
    """True if needle can be formed by deleting characters from haystack (order preserved)."""
    if not needle:
        return True
    i = 0
    for c in haystack:
        if i < len(needle) and needle[i] == c:
            i += 1
    return i == len(needle)


def _credential_alias_bridge_variant(
    field_name: str,
    forward_aliases: Dict[str, Any],
) -> Optional[str]:
    """
    Find a credential-related forward alias whose key is a subsequence of the normalized
    field (digits stripped). Prefer the longest key — e.g. bankstatementlinkpassword still
    matches bankstatementsavingsaccount…linkpassword without hardcoding partner names.

    Returns a registry variant string to re-enter matching, or None.
    """
    if not forward_aliases or not field_name:
        return None
    nb_nd = re.sub(r"\d+", "", normalize_basic(field_name))
    if len(nb_nd) < _CREDENTIAL_ALIAS_BRIDGE_MIN_KEY_LEN:
        return None

    cred_in_key = ("password", "passphrase", "passcode", "pin")
    best_key: Optional[str] = None
    best_len = 0
    for key in forward_aliases:
        if len(key) < _CREDENTIAL_ALIAS_BRIDGE_MIN_KEY_LEN:
            continue
        if not any(t in key for t in cred_in_key):
            continue
        if not _is_subsequence(key, nb_nd):
            continue
        if len(key) > best_len:
            best_len = len(key)
            best_key = key

    if not best_key:
        return None

    entry = forward_aliases.get(best_key) or {}
    variants = entry.get("variants") or []
    if variants and isinstance(variants[0], str) and variants[0].strip():
        return variants[0].strip()

    return None


CATEGORY_STOPWORDS = {
    "lead", "details", "detail", "data", "info", "information", "section",
    "left", "right", "top", "bottom", "tab", "page", "field", "fields",
    "screen", "view", "module", "group", "grouping", "form",
    "api", "mapping", "profile",
}

ENTITY_CATEGORY_HINTS = {
    "COAPPLICANT": ("coapplicant", "coapp", "co-applicant"),
    "GUARANTOR": ("guarantor",),
    "FEE": ("fee", "fees", "charge", "charges", "commission"),
    "DOCUMENT": ("document", "documents", "upload", "attachment", "kyc", "proof"),
    "LOAN": ("loan", "disbursement", "repayment", "emi", "interest", "vehicle"),
    "APPLICANT": ("applicant", "customer", "borrower"),
}

# Token hits (camelCase / PascalCase) for fee/charge semantics — avoids missing
# `\bcharge\b` when "Charge" is glued to another word (e.g. siteVisitCharge).
_FEE_CHARGE_TOKENS = frozenset({
    "fee", "fees", "charge", "charges", "charging", "commission", "premium",
})


def partner_field_excludes_generic_fee_bucket(field_name: Optional[str]) -> bool:
    """
    True when the field must not be routed as the generic FEE entity / FEE excel bucket:
    insurance premium or amount columns; loan processing fee *percentage* (loan pricing).
    """
    if not field_name:
        return False
    field_lower = field_name.lower()
    norm = re.sub(r"[^a-z0-9]+", "", field_lower)
    if "insurance" in field_lower:
        return True
    if "processingfee" in norm and re.search(r"percent|pct|per\b", field_lower):
        return True
    return False


def field_implies_fee_charge_semantics(field_name: Optional[str]) -> bool:
    """
    True when the partner field name denotes a fee/charge (maps to entity FEE).
    Uses substring checks aligned with FEE_KEYWORDS plus tokenized camelCase splits
    so names like siteVisitCharge / clearingCharges match reliably.
    """
    if not field_name:
        return False
    if partner_field_excludes_generic_fee_bucket(field_name):
        return False
    field_lower = field_name.lower()
    if any(kw in field_lower for kw in FEE_KEYWORDS):
        return True
    # "processing fee" phrase (substring "processing" alone is intentionally omitted from FEE_KEYWORDS)
    if "processing" in field_lower and "fee" in field_lower:
        return True
    tokens = set(_tokenize_text(field_name)) & _FEE_CHARGE_TOKENS
    if "premium" in tokens and "insurance" in field_lower:
        tokens.discard("premium")
    if "fee" in tokens and re.search(r"processing\s*fee|processingfee", field_lower):
        # processingFeePercentage etc. — loan pricing, not generic FEE bucket
        if re.search(r"percent|pct|per\b", field_lower):
            tokens.discard("fee")
    if tokens:
        return True
    blob = f"{field_lower} "
    if re.search(r"\bcharges?\b", blob, re.I):
        return True
    if re.search(r"\bfee(s)?\b", blob, re.I) and "feedback" not in field_lower:
        return True
    return False


def normalize_category(value: Optional[str]) -> str:
    if not value:
        return ""
    normalized = value.strip().lower()
    normalized = normalized.replace("&", "and")
    normalized = re.sub(r"[\s_\-./()]+", "", normalized)
    return normalized


def _tokenize_text(value: Optional[str]) -> List[str]:
    if not value:
        return []
    expanded = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", value)
    tokens = re.findall(r"[A-Za-z0-9]+", expanded.lower())
    return [t for t in tokens if t]


def _category_signal_tokens(column_category: Optional[str]) -> List[str]:
    tokens = []
    for token in _tokenize_text(column_category):
        if len(token) <= 2:
            continue
        if token in CATEGORY_STOPWORDS:
            continue
        tokens.append(token)
    return list(dict.fromkeys(tokens))


def _is_valid_category_label(
    column_category: Optional[str],
    refs: Dict[str, Any],
) -> bool:
    if not column_category:
        return False
    grouping_to_entity = refs.get("entity_routing", {}).get("grouping_to_entity", {})
    if column_category in grouping_to_entity:
        return True
    normalized_category = normalize_category(column_category)
    if not normalized_category:
        return False
    normalized_routing = {
        normalize_category(key)
        for key in grouping_to_entity.keys()
        if key
    }
    return normalized_category in normalized_routing


def _score_category_alignment(
    column_category: Optional[str],
    excel_key: str,
    json_key: Optional[str] = "",
    description: Optional[str] = "",
) -> int:
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
    if not _is_valid_category_label(column_category, refs):
        return None
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
    Used for exact matching where APPLICANT/COAPPLICANT distinction matters.

    Examples:
        normalize_basic("APPLICANT_DATE_OF_BIRTH") → "applicantdateofbirth"
        normalize_basic("bureauScore")             → "bureauscore"
    """
    if not field_name:
        return ""
    normalized = field_name.lower()
    normalized = re.sub(r'[_\s\.\-\'\(\)]', '', normalized)
    return normalized


def normalize_field(field_name: str) -> str:
    """
    Normalize a partner field name for alias matching.
    Strips separators AND common entity prefixes so that
    'applicantGender' matches the same alias as 'gender'.

    Examples:
        normalize_field("Date_of_Birth")           → "dateofbirth"
        normalize_field("Applicant.Gender")        → "gender"
        normalize_field("CoApplicant - AnnualInc") → "annualinc"
    """
    if not field_name:
        return ""
    normalized = field_name.lower()
    normalized = re.sub(r'[_\s\.\-\'\(\)]', '', normalized)
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
    normalized = normalize_field(value or "")
    if not normalized:
        return ""
    normalized = normalized.replace("aadhaar", "aadhar")
    normalized = normalized.replace("number", "no")
    normalized = normalized.replace("num", "no")
    return normalized


def _canonicalize_json_key(value: Optional[str]) -> str:
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
    field_lower = (field_name or "").lower()
    cat_lower = (column_category or "").lower()
    blob = f"{field_lower} {cat_lower}"
    # Fee / charge semantics override broad sheet labels like "Loan Details"
    # so partner charge columns route to FEE entity (matches fee bucket rules).
    if field_implies_fee_charge_semantics(field_name):
        return "FEE"
    if any(
        t in blob
        for t in (
            "processing fee",
            "processing fees",
            "service fee",
            "gst fee",
            "insurance fee",
            "stamp duty",
            "commission",
            "penalty fee",
            "other charges",
            "other's charges",
        )
    ):
        return "FEE"

    if column_category:
        if column_category in entity_routing:
            return entity_routing[column_category]
        normalized_category = normalize_category(column_category)
        normalized_routing = {
            normalize_category(key): value for key, value in entity_routing.items()
        }
        if normalized_category in normalized_routing:
            return normalized_routing[normalized_category]
        for entity_name, hints in ENTITY_CATEGORY_HINTS.items():
            if any(hint in normalized_category for hint in hints):
                return entity_name

    if any(x in field_lower for x in ["guarantor", "guaranter", "guaraontor"]):
        match = re.search(r'guarantor(\d)', field_lower)
        if match:
            return f"GUARANTOR{int(match.group(1))}"
        return "GUARANTOR"

    if any(x in field_lower for x in ["coapplicant", "co_applicant", "co-applicant", "coapp"]):
        match = re.search(r'(?:coapplicant|coapp)(\d)', field_lower)
        if match:
            return f"COAPPLICANT{int(match.group(1))}"
        return "COAPPLICANT"

    if any(x in field_lower for x in ["applicant", "customer", "borrower"]):
        return "APPLICANT"
    if "loan" in field_lower:
        return "LOAN"
    if any(x in field_lower for x in ["document", "upload", "file", "attachment"]):
        return "DOCUMENT"

    if process_name:
        return entity_routing.get(process_name, "APPLICANT")
    return "APPLICANT"


def load_references(references_dir: str) -> Dict[str, Any]:
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
                f"Reference file not found: {file_path}\nExpected in: {references_dir}"
            )
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                references[key] = json.load(f)
        except json.JSONDecodeError as e:
            raise json.JSONDecodeError(
                f"Invalid JSON in {filename}: {e.msg}", e.doc, e.pos
            )
    return references


# ── Index builders ─────────────────────────────────────────────────────────────

def _build_field_dict_index(
    field_dictionary: Dict[str, Any]
) -> Dict[str, List[Tuple[str, str, str]]]:
    """
    Build lookup index keyed by normalize_basic(excel_key).
    Sourced from by_role entries.
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


def _build_by_excel_key_index(
    field_dictionary: Dict[str, Any]
) -> Dict[str, Tuple[str, str]]:
    """
    Build a lookup index directly from the authoritative by_excel_key dictionary.
    Keyed by normalize_basic(excel_key) → (excel_key, json_key).

    This index is the ground truth — only keys that are actual internal excel_keys
    appear here. Used to VALIDATE that a candidate from the role-based index is a
    genuine excel_key and not a data artifact.
    """
    index: Dict[str, Tuple[str, str]] = {}
    for excel_key, entry in field_dictionary.get("by_excel_key", {}).items():
        if not excel_key:
            continue
        normalized = normalize_basic(excel_key)
        json_key = entry.get("json_key", "") or ""
        if normalized not in index:
            index[normalized] = (excel_key, json_key)
    return index


def _build_process_field_dict_index(
    field_dictionary: Dict[str, Any],
    process_name: Optional[str],
) -> Dict[str, List[Tuple[str, str, str]]]:
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


def _build_process_by_excel_key_index(
    field_dictionary: Dict[str, Any],
    process_name: Optional[str],
) -> Dict[str, Tuple[str, str]]:
    """
    Authoritative by_excel_key index scoped to a specific process.
    Keyed by normalize_basic(excel_key) → (excel_key, json_key).
    """
    normalized_process = normalize_process_name(process_name)
    index: Dict[str, Tuple[str, str]] = {}
    for excel_key, entry in field_dictionary.get("by_excel_key", {}).items():
        if not excel_key:
            continue
        # Filter by process if specified
        if normalized_process:
            process_names = entry.get("process_names") or []
            if process_names and normalized_process not in {str(p).upper() for p in process_names}:
                continue
        normalized = normalize_basic(excel_key)
        json_key = entry.get("json_key", "") or ""
        if normalized not in index:
            index[normalized] = (excel_key, json_key)
    return index


def _validate_exact_match_in_by_excel_key(
    normalized_lookup: str,
    by_excel_key_index: Dict[str, Tuple[str, str]],
) -> Optional[Tuple[str, str]]:
    """
    Verify a normalized partner field name exists as an actual excel_key in by_excel_key.
    Returns (excel_key, json_key) if found, None otherwise.

    This prevents false exact matches where by_role entries contain partner-style
    field names (e.g. 'annualIncome') that collide with unrelated internal keys.
    """
    return by_excel_key_index.get(normalized_lookup)


# ── Candidate selection ─────────────────────────────────────────────────────────

def _select_best_candidate(
    candidates: List[Tuple[str, str, str]],
    entity: str,
    column_category: Optional[str] = None,
    process_name: Optional[str] = None,
    field_dictionary: Optional[Dict[str, Any]] = None,
) -> Tuple[str, str]:
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
    if normalized_category:
        category_tokens = set(_category_signal_tokens(column_category))
        if category_tokens:
            for ek, jk, role in candidates:
                candidate_tokens = set(_tokenize_text(" ".join([ek, jk or ""])))
                if category_tokens & candidate_tokens:
                    return ek, jk

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

    for preferred_role in role_priority:
        for ek, jk, role in candidates:
            if role == preferred_role:
                return ek, jk

    return candidates[0][0], candidates[0][1]


def _get_alias_tier(frequency: int) -> str:
    if frequency >= 30:
        return "tier1"
    elif frequency >= 10:
        return "tier2"
    elif frequency >= 3:
        return "tier3"
    else:
        return "tier4"


def _confidence_for_tier(tier: str) -> float:
    return {
        "tier1": 0.98,
        "tier2": 0.92,
        "tier3": 0.85,
        "tier4": 0.75,
    }.get(tier, 0.50)


def _coapplicant_slot_index(current_entity: str) -> int:
    e = (current_entity or "").strip().upper()
    if e == "COAPPLICANT":
        return 1
    m = re.match(r"^COAPPLICANT(\d+)$", e)
    if m:
        return int(m.group(1))
    return 1


def _is_applicant_scoped_putm_excel_key(excel_key: str) -> bool:
    """
    True when excel_key is clearly primary-applicant or applicant-parameter
    scoped (prefix / naming convention), excluding keys that are already
    co-applicant-scoped and excluding neutral catalogue keys that carry no
    applicant prefix in their excel_key.
    """
    ek = (excel_key or "").strip().upper()
    if not ek or ek.startswith("COAPPLICANT") or _is_coapplicant_scoped_putm_excel_key(ek):
        return False
    if re.match(r"^LOANAPPLICANTPARAM\d+$", ek):
        return True
    if ek.startswith("APPLICANT"):
        return True
    return False


def _remap_applicant_putm_to_coapplicant_catalogue(
    current_entity: str,
    applicant_style_excel_key: str,
    by_excel_key_index: Dict[str, Tuple[str, str]],
    field_dictionary_index: Dict[str, List[Tuple[str, str, str]]],
    field_dictionary: Dict[str, Any],
    process_name: Optional[str],
) -> Optional[Tuple[str, str]]:
    """
    Map applicant-scoped PUTM excel_keys to the matching co-applicant slot
    for index N using prefix transforms, then validate against the field
    dictionary (no per-field or per-partner hardcoding).
    """
    if not (current_entity or "").strip().upper().startswith("COAPPLICANT"):
        return None
    n = _coapplicant_slot_index(current_entity)
    pref = f"COAPPLICANT{n}"
    src = (applicant_style_excel_key or "").strip()
    eu = src.upper()

    candidates: List[str] = []
    m = re.match(r"^APPLICANTCUSTOMER(.+)$", eu)
    if m:
        candidates.append(f"{pref}{m.group(1)}")
    cand_sub = re.sub(r"^(?i)applicant", pref, src, count=1)
    if cand_sub.upper() != eu:
        candidates.append(cand_sub.upper())
    m2 = re.match(r"^LOANAPPLICANTPARAM(\d+)$", eu, re.I)
    if m2:
        candidates.append(f"LOANCOAPP{n}CUSTPARAM{m2.group(1)}".upper())

    seen: set[str] = set()
    for c in candidates:
        cu = c.upper()
        if cu in seen:
            continue
        seen.add(cu)
        nb = normalize_basic(cu)
        if nb in by_excel_key_index:
            ek, jk = by_excel_key_index[nb]
            if process_name and not _is_excel_key_allowed_for_process(
                field_dictionary, ek, process_name
            ):
                continue
            return ek, jk
        if nb in field_dictionary_index:
            best_ek, best_jk = _select_best_candidate(
                field_dictionary_index[nb],
                current_entity,
                None,
                process_name,
                field_dictionary,
            )
            if normalize_basic(best_ek) in by_excel_key_index:
                ek2, jk2 = by_excel_key_index[normalize_basic(best_ek)]
                if process_name and not _is_excel_key_allowed_for_process(
                    field_dictionary, ek2, process_name
                ):
                    continue
                return ek2, jk2 or best_jk or jk2
    return None


def _first_loancoapp_custparam_for_entity(
    current_entity: str,
    field_dictionary: Dict[str, Any],
    process_name: Optional[str],
) -> Optional[Tuple[str, str]]:
    """First numbered LOANCOAPP{n}CUSTPARAM* in catalogue for this co-applicant slot."""
    if not (current_entity or "").strip().upper().startswith("COAPPLICANT"):
        return None
    n = _coapplicant_slot_index(current_entity)
    prefix = f"LOANCOAPP{n}CUSTPARAM"
    by_ek = field_dictionary.get("by_excel_key") or {}

    def _tail_num(ek: str) -> int:
        m = re.search(r"(\d+)$", ek)
        return int(m.group(1)) if m else 0

    keys = sorted(
        (k for k in by_ek if k.upper().startswith(prefix)),
        key=_tail_num,
    )
    for ek in keys:
        if process_name and not _is_excel_key_allowed_for_process(
            field_dictionary, ek, process_name
        ):
            continue
        jk = (by_ek.get(ek) or {}).get("json_key") or ""
        return ek, jk
    return None


def _try_entity_substitution(
    excel_key: str,
    current_entity: str,
    field_dictionary_index: Dict[str, List[Tuple[str, str, str]]],
    field_dictionary: Dict[str, Any],
    process_name: Optional[str] = None,
    by_excel_key_index: Optional[Dict[str, Tuple[str, str]]] = None,
) -> Tuple[Optional[str], Optional[str], float]:
    """
    GUARANTOR: replace leading APPLICANT prefix with GUARANTOR1 where catalogue allows.

    COAPPLICANT substitution is handled in the alias pipeline via
    _remap_applicant_putm_to_coapplicant_catalogue (not here) so APPLICANTCUSTOMER*
    patterns resolve to COAPPLICANTn* keys correctly.
    """
    by_excel_key_index = by_excel_key_index or {}
    if current_entity.startswith("COAPPLICANT"):
        return None, None, 0.0
    if not current_entity.startswith("GUARANTOR"):
        return None, None, 0.0
    if "APPLICANT" not in excel_key.upper():
        return None, None, 0.0

    target_prefix = "GUARANTOR1"
    substituted = re.sub(r"(?i)applicant(?!\d)", target_prefix, excel_key, count=1)
    normalized_sub = normalize_basic(substituted)
    if normalized_sub in by_excel_key_index:
        ek, jk = by_excel_key_index[normalized_sub]
        if process_name and not _is_excel_key_allowed_for_process(
            field_dictionary, ek, process_name
        ):
            return None, None, 0.0
        return ek, jk, -0.10
    if normalized_sub in field_dictionary_index:
        best_ek, best_jk = _select_best_candidate(
            field_dictionary_index[normalized_sub],
            current_entity,
            None,
            process_name,
            field_dictionary,
        )
        nb = normalize_basic(best_ek)
        if nb in by_excel_key_index:
            ek2, jk2 = by_excel_key_index[nb]
            if process_name and not _is_excel_key_allowed_for_process(
                field_dictionary, ek2, process_name
            ):
                return None, None, 0.0
            return ek2, jk2 or best_jk, -0.10
    return None, None, 0.0


def _is_coapplicant_scoped_putm_excel_key(excel_key: str) -> bool:
    """True when excel_key clearly targets co-applicant JSON (not applicant)."""
    ek = (excel_key or "").strip().upper()
    if not ek:
        return False
    if re.match(r"^COAPPLICANT\d+", ek):
        return True
    if re.match(r"^LOANCOAPP\d+", ek):
        return True
    return False


def _remap_putm_key_for_applicant_customer_entity(
    entity: Optional[str],
    excel_key: str,
    by_excel_key_index: Dict[str, Tuple[str, str]],
) -> Optional[Tuple[str, str]]:
    """
    For APPLICANT/CUSTOMER rows, map co-applicant-only PUTM keys to catalogue
    applicant equivalents when they exist (COAPPLICANT1X -> APPLICANTX,
    LOANCOAPPnCUSTPARAMk -> LOANAPPLICANTPARAMk).
    """
    ent = (entity or "").strip().upper()
    if ent not in ("APPLICANT", "CUSTOMER"):
        return None
    ek = (excel_key or "").strip()
    if not _is_coapplicant_scoped_putm_excel_key(ek):
        return None

    candidates: List[str] = []
    if re.match(r"^COAPPLICANT\d+", ek, re.I):
        candidates.append(re.sub(r"^COAPPLICANT\d+", "APPLICANT", ek, count=1, flags=re.I).upper())
    m = re.match(r"^LOANCOAPP\d+CUSTPARAM(\d+)$", ek, re.I)
    if m:
        candidates.append(f"LOANAPPLICANTPARAM{m.group(1)}".upper())

    for cand in candidates:
        nb = normalize_basic(cand)
        if nb in by_excel_key_index:
            return by_excel_key_index[nb]
    return None


# ── Core match function ─────────────────────────────────────────────────────────

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
    0.  FEE / CHARGE SEMANTICS      — fee/charge field names → excel_key FEE, entity FEE
        (before exact/alias so charge columns are not mis-matched to unrelated keys)
    1a. EXACT MATCH (raw, field_dictionary-first)
        — normalize_basic only, no prefix stripping
        — validates candidate exists in by_excel_key before accepting
    1b. EXACT MATCH (prefix-stripped, field_dictionary-first)
        — normalize_field strips entity prefixes
        — validates candidate exists in by_excel_key before accepting
    2.  ALIAS MATCH (tier 1–4)       — lookup in alias_registry.forward
    3.  DOCUMENT NAME DETECTION
    4.  DOCUMENT ID DETECTION
    5.  UNMATCHED                    — hand off to LLM layer
    """
    # ── Entity detection ──────────────────────────────────────────────────────
    if not entity:
        entity = detect_entity(
            column_category,
            field_name,
            refs["entity_routing"].get("grouping_to_entity", {})
        )

    # Fee/charge semantics first: partner names like siteVisitCharge / clearingCharges
    # must map to the FEE bucket, not lose to a coincidental exact/alias hit.
    if field_implies_fee_charge_semantics(field_name):
        fl = (field_name or "").lower()
        matched_kw = [kw for kw in FEE_KEYWORDS if kw in fl]
        if not matched_kw:
            matched_kw = sorted(set(_tokenize_text(field_name)) & _FEE_CHARGE_TOKENS)
        return MatchResult(
            partner_field=field_name,
            column_category=column_category,
            matched_excel_key="FEE",
            matched_json_key=None,
            confidence=0.90,
            match_type=MatchType.FEE.value,
            reasoning=(
                "Fee detection (priority): field denotes fee/charge semantics "
                f"({', '.join(matched_kw) if matched_kw else 'token match'}) — "
                "mapped to FEE before generic exact/alias matching."
            ),
            entity="FEE",
        )

    field_dictionary = refs.get("field_dictionary", {})
    alias_registry = refs.get("alias_registry", {})

    # Build role-based index (all processes)
    field_dict_index = _build_field_dict_index(field_dictionary)

    # Build authoritative by_excel_key indexes for validation
    # These are the ground truth — only genuine internal excel_keys appear here
    by_excel_key_index = _build_by_excel_key_index(field_dictionary)
    by_excel_key_process_index = _build_process_by_excel_key_index(field_dictionary, process_name)

    # Build process-scoped role index
    process_field_dict_index = _build_process_field_dict_index(field_dictionary, process_name)

    # ====== 1a. EXACT MATCH — raw (normalize_basic, prefixes preserved) ======
    #
    # Priority: check by_excel_key (authoritative) FIRST, then fall back to the
    # role-based index. This prevents false matches where by_role contains
    # partner-style field names that collide with unrelated internal keys.
    #
    basic_normalized = normalize_basic(field_name)

    # 1a-i: Check authoritative by_excel_key (process-scoped first)
    process_exact = _validate_exact_match_in_by_excel_key(basic_normalized, by_excel_key_process_index)
    if process_exact:
        excel_key, json_key = process_exact
        return MatchResult(
            partner_field=field_name,
            column_category=column_category,
            matched_excel_key=excel_key,
            matched_json_key=json_key,
            confidence=1.0,
            match_type=MatchType.EXACT.value,
            reasoning=(
                f"Exact match (raw, field_dictionary-first) within process "
                f"'{normalize_process_name(process_name)}': "
                f"normalized '{field_name}' is a known excel_key '{excel_key}'"
            ),
            entity=entity,
        )

    # 1a-ii: Check authoritative by_excel_key (all processes)
    # Only accept if the matched excel_key is actually allowed for this process.
    # Do NOT call _find_process_equivalent_by_json here — swapping to another key
    # that shares the same json_key produces wrong results (e.g. annualIncome →
    # COMMERCIALCIBIL). If the key isn't valid for this process, fall through to
    # alias matching which has proper frequency-ranked candidates.
    global_exact = _validate_exact_match_in_by_excel_key(basic_normalized, by_excel_key_index)
    if global_exact:
        excel_key, json_key = global_exact
        if process_name and not _is_excel_key_allowed_for_process(field_dictionary, excel_key, process_name):
            # Key exists globally but not for this process — do NOT swap, fall through
            pass
        else:
            return MatchResult(
                partner_field=field_name,
                column_category=column_category,
                matched_excel_key=excel_key,
                matched_json_key=json_key,
                confidence=1.0,
                match_type=MatchType.EXACT.value,
                reasoning=(
                    f"Exact match (raw, field_dictionary-first): "
                    f"normalized '{field_name}' is a known excel_key '{excel_key}'"
                ),
                entity=entity,
            )

    # 1a-iii: Fall back to role-based index (legacy path) — only accept if
    # the candidate is also present in by_excel_key to avoid false positives
    if basic_normalized in process_field_dict_index:
        candidates = process_field_dict_index[basic_normalized]
        excel_key, json_key = _select_best_candidate(
            candidates, entity, column_category, process_name, field_dictionary,
        )
        # Validate: only accept if this excel_key is a known key in by_excel_key
        if normalize_basic(excel_key) in by_excel_key_index:
            return MatchResult(
                partner_field=field_name,
                column_category=column_category,
                matched_excel_key=excel_key,
                matched_json_key=json_key,
                confidence=1.0,
                match_type=MatchType.EXACT.value,
                reasoning=(
                    f"Exact match (raw) within process '{normalize_process_name(process_name)}': "
                    f"normalized '{field_name}' matches '{excel_key}'"
                ),
                entity=entity,
            )

    if basic_normalized in field_dict_index:
        candidates = field_dict_index[basic_normalized]
        excel_key, json_key = _select_best_candidate(
            candidates, entity, column_category, process_name, field_dictionary,
        )
        # Validate: only accept if this excel_key is genuinely in by_excel_key
        # AND is allowed for this process. Never swap to a different key here —
        # fall through to alias matching if the process doesn't allow this key.
        if normalize_basic(excel_key) in by_excel_key_index:
            if process_name and not _is_excel_key_allowed_for_process(field_dictionary, excel_key, process_name):
                pass  # fall through to alias matching
            else:
                return MatchResult(
                    partner_field=field_name,
                    column_category=column_category,
                    matched_excel_key=excel_key,
                    matched_json_key=json_key,
                    confidence=1.0,
                    match_type=MatchType.EXACT.value,
                    reasoning=f"Exact match (raw): normalized '{field_name}' matches '{excel_key}'",
                    entity=entity,
                )

    # ====== 1b. EXACT MATCH — prefix-stripped (normalize_field) ======
    #
    # Same field_dictionary-first priority: check by_excel_key before role index.
    #
    normalized_field = normalize_field(field_name)

    if normalized_field and normalized_field != basic_normalized:

        # 1b-i: Check authoritative by_excel_key (process-scoped first)
        process_exact_stripped = _validate_exact_match_in_by_excel_key(
            normalized_field, by_excel_key_process_index
        )
        if process_exact_stripped:
            excel_key, json_key = process_exact_stripped
            return MatchResult(
                partner_field=field_name,
                column_category=column_category,
                matched_excel_key=excel_key,
                matched_json_key=json_key,
                confidence=1.0,
                match_type=MatchType.EXACT.value,
                reasoning=(
                    f"Exact match (prefix-stripped, field_dictionary-first) within process "
                    f"'{normalize_process_name(process_name)}': "
                    f"'{field_name}' → '{normalized_field}' is a known excel_key '{excel_key}'"
                ),
                entity=entity,
            )

        # 1b-ii: Check authoritative by_excel_key (all processes)
        # Same rule: if the key isn't valid for this process, fall through to
        # alias matching — do NOT swap via _find_process_equivalent_by_json.
        global_exact_stripped = _validate_exact_match_in_by_excel_key(
            normalized_field, by_excel_key_index
        )
        if global_exact_stripped:
            excel_key, json_key = global_exact_stripped
            if process_name and not _is_excel_key_allowed_for_process(field_dictionary, excel_key, process_name):
                pass  # fall through to alias matching
            else:
                return MatchResult(
                    partner_field=field_name,
                    column_category=column_category,
                    matched_excel_key=excel_key,
                    matched_json_key=json_key,
                    confidence=1.0,
                    match_type=MatchType.EXACT.value,
                    reasoning=(
                        f"Exact match (prefix-stripped, field_dictionary-first): "
                        f"'{field_name}' → '{normalized_field}' is a known excel_key '{excel_key}'"
                    ),
                    entity=entity,
                )

        # 1b-iii: Fall back to role-based index with validation
        if normalized_field in process_field_dict_index:
            candidates = process_field_dict_index[normalized_field]
            excel_key, json_key = _select_best_candidate(
                candidates, entity, column_category, process_name, field_dictionary,
            )
            if normalize_basic(excel_key) in by_excel_key_index:
                return MatchResult(
                    partner_field=field_name,
                    column_category=column_category,
                    matched_excel_key=excel_key,
                    matched_json_key=json_key,
                    confidence=1.0,
                    match_type=MatchType.EXACT.value,
                    reasoning=(
                        f"Exact match (prefix-stripped) within process "
                        f"'{normalize_process_name(process_name)}': "
                        f"'{field_name}' → '{normalized_field}' matches '{excel_key}'"
                    ),
                    entity=entity,
                )

        if normalized_field in field_dict_index:
            candidates = field_dict_index[normalized_field]
            excel_key, json_key = _select_best_candidate(
                candidates, entity, column_category, process_name, field_dictionary,
            )
            # Never swap to a process-equivalent here — fall through to alias if not allowed
            if normalize_basic(excel_key) in by_excel_key_index:
                if process_name and not _is_excel_key_allowed_for_process(field_dictionary, excel_key, process_name):
                    pass  # fall through to alias matching
                else:
                    return MatchResult(
                        partner_field=field_name,
                        column_category=column_category,
                        matched_excel_key=excel_key,
                        matched_json_key=json_key,
                        confidence=1.0,
                        match_type=MatchType.EXACT.value,
                        reasoning=(
                            f"Exact match (prefix-stripped): "
                            f"'{field_name}' → '{normalized_field}' matches '{excel_key}'"
                        ),
                        entity=entity,
                    )
    else:
        normalized_field = normalize_field(field_name)

    # ====== 2. ALIAS MATCH (tiers 1–4) ======
    forward_aliases = alias_registry.get("forward", {})

    if normalized_field in forward_aliases:
        alias_entry = forward_aliases[normalized_field]
        target_excel_key = alias_entry.get("target_excel_key", "")
        target_json_key = alias_entry.get("target_json_key", "")
        frequency = alias_entry.get("frequency", 0)

        tier = _get_alias_tier(frequency)
        match_type = f"alias_{tier}"
        base_confidence = _confidence_for_tier(tier)

        substituted_key, substituted_json_key, conf_adjustment = _try_entity_substitution(
            target_excel_key,
            entity,
            field_dict_index,
            field_dictionary,
            process_name,
            by_excel_key_index,
        )

        final_excel_key = substituted_key or target_excel_key
        final_json_key = substituted_json_key if substituted_key else target_json_key
        final_confidence = base_confidence + conf_adjustment

        alias_override = _find_category_alias_override(
            normalized_field, column_category, refs, final_excel_key, process_name,
        )
        if alias_override:
            override_excel_key, override_json_key = alias_override
            if override_excel_key and override_excel_key != final_excel_key:
                final_excel_key = override_excel_key
                final_json_key = override_json_key
                final_confidence = min(0.99, final_confidence + 0.02)

        remap_pair = _remap_putm_key_for_applicant_customer_entity(
            entity, final_excel_key, by_excel_key_index
        )
        if remap_pair:
            final_excel_key, final_json_key = remap_pair

        ent_u = (entity or "").strip().upper()
        if ent_u.startswith("COAPPLICANT") and _is_applicant_scoped_putm_excel_key(
            final_excel_key
        ):
            co_hit = _remap_applicant_putm_to_coapplicant_catalogue(
                ent_u,
                final_excel_key,
                by_excel_key_index,
                field_dict_index,
                field_dictionary,
                process_name,
            )
            if co_hit:
                final_excel_key, final_json_key = co_hit
                final_confidence = min(final_confidence, 0.92)
            else:
                gen_hit = _first_loancoapp_custparam_for_entity(
                    ent_u, field_dictionary, process_name
                )
                if gen_hit:
                    g_ek, g_jk = gen_hit
                    return MatchResult(
                        partner_field=field_name,
                        column_category=column_category,
                        matched_excel_key=g_ek,
                        matched_json_key=g_jk,
                        confidence=min(base_confidence, 0.72),
                        match_type=MatchType.PARAMETER_FALLBACK.value,
                        reasoning=(
                            f"Alias ({tier}) target '{target_excel_key}' is applicant-scoped while "
                            f"entity={ent_u}; no exact co-applicant catalogue key — assigned generic "
                            f"parameter slot '{g_ek}'."
                        ),
                        entity=entity,
                    )
                return MatchResult(
                    partner_field=field_name,
                    column_category=column_category,
                    matched_excel_key=None,
                    matched_json_key=None,
                    confidence=0.0,
                    match_type=MatchType.UNMATCHED.value,
                    reasoning=(
                        f"Alias ({tier}) target '{target_excel_key}' is applicant-scoped while "
                        f"entity={ent_u}; no co-applicant PUTM equivalent or LOANCOAPP slot in "
                        f"catalogue."
                    ),
                    entity=entity,
                )

        if (
            ent_u in ("APPLICANT", "CUSTOMER")
            and not remap_pair
            and _is_coapplicant_scoped_putm_excel_key(final_excel_key)
        ):
            return MatchResult(
                partner_field=field_name,
                column_category=column_category,
                matched_excel_key=None,
                matched_json_key=None,
                confidence=0.0,
                match_type=MatchType.UNMATCHED.value,
                reasoning=(
                    f"Alias ({tier}) target is co-applicant-only PUTM key '{final_excel_key}' "
                    f"for entity={ent_u}; no applicant-safe equivalent in catalogue."
                ),
                entity=entity,
            )

        if not alias_override and not _is_excel_key_allowed_for_process(
            field_dictionary, final_excel_key, process_name
        ):
            equivalent = _find_process_equivalent_by_json(
                refs, final_json_key, entity, column_category, process_name,
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
                        f"'{field_name}' resolved to '{final_excel_key}' "
                        f"for process '{process_name}'"
                    ),
                    entity=entity,
                )

            fallback = _find_process_aware_alias_fallback(
                normalized_field, refs, process_name,
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

            return MatchResult(
                partner_field=field_name,
                column_category=column_category,
                matched_excel_key=None,
                matched_json_key=None,
                confidence=0.0,
                match_type=MatchType.UNMATCHED.value,
                reasoning=(
                    f"Alias target '{final_excel_key}' is not allowed "
                    f"for process '{process_name}'"
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
            entity=entity,
        )

    # ====== 3. DOCUMENT NAME DETECTION ======
    field_lower = field_name.lower()
    suggests_credential = any(h in field_lower for h in CREDENTIAL_KEYWORD_HINTS)
    doc_keywords_found = [kw for kw in DOCUMENT_KEYWORDS if kw in field_lower]
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
        any(field_name.endswith(s) for s in ["Name", "File", "Image", "FileName", "ImageName"])
        and has_strong_doc_signal
    )

    if has_doc_keyword and looks_like_filename and not suggests_credential:
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
            entity="DOCUMENT",
        )

    # ====== 4. DOCUMENT ID DETECTION ======
    has_locator_keyword = any(kw in field_lower for kw in DOCUMENT_LOCATOR_KEYWORDS)

    if has_doc_keyword and has_locator_keyword and not suggests_credential:
        return MatchResult(
            partner_field=field_name,
            column_category=column_category,
            matched_excel_key="DOCUMENTID",
            matched_json_key=None,
            confidence=0.90,
            match_type=MatchType.DOCUMENT_ID.value,
            reasoning=(
                f"Document ID detection: field contains document keyword "
                f"and locator keyword "
                f"({', '.join(kw for kw in DOCUMENT_LOCATOR_KEYWORDS if kw in field_lower)})"
            ),
            entity="DOCUMENT",
        )

    # ====== 4b. CREDENTIAL + DOCUMENT/LOCATOR: bridge long names to forward alias ======
    # Long camelCase names often fail exact alias lookup but embed a shorter alias key as a
    # subsequence (after digit strip). Re-resolve via registry variant — no hardcoded targets.
    if suggests_credential and has_doc_keyword and has_locator_keyword:
        bridge_variant = _credential_alias_bridge_variant(field_name, forward_aliases)
        if bridge_variant and normalize_basic(field_name) != normalize_basic(bridge_variant):
            inner = match_field(
                bridge_variant,
                column_category,
                entity,
                refs,
                process_name=process_name,
            )
            if inner.match_type != MatchType.UNMATCHED.value and inner.matched_excel_key:
                return MatchResult(
                    partner_field=field_name,
                    column_category=column_category,
                    matched_excel_key=inner.matched_excel_key,
                    matched_json_key=inner.matched_json_key,
                    confidence=min(float(inner.confidence or 0.0), 0.88),
                    match_type=inner.match_type,
                    reasoning=(
                        "Credential + document/locator field; bridged to catalogue alias via "
                        f"registry variant (long partner name). {inner.reasoning}"
                    ),
                    entity=inner.entity,
                )

    # ====== 5. UNMATCHED ======
    return MatchResult(
        partner_field=field_name,
        column_category=column_category,
        matched_excel_key=None,
        matched_json_key=None,
        confidence=0.0,
        match_type=MatchType.UNMATCHED.value,
        reasoning="No deterministic match found; requires LLM layer for semantic matching",
        entity=entity,
    )


def match_batch(
    fields: List[Dict[str, Any]],
    refs: Dict[str, Any],
    process_name: Optional[str] = None,
) -> Dict[str, List[MatchResult]]:
    matched = []
    unmatched = []
    for field_dict in fields:
        field_name = field_dict.get("field_name", "")
        column_category = field_dict.get("column_category")
        entity = field_dict.get("entity")
        if not field_name:
            continue
        result = match_field(field_name, column_category, entity, refs, process_name=process_name)
        if result.match_type == MatchType.UNMATCHED.value:
            unmatched.append(result)
        else:
            matched.append(result)
    return {"matched": matched, "unmatched": unmatched}


def main():
    parser = argparse.ArgumentParser(
        description="Deterministic Field Matching Engine",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python matching_engine.py --field "dateOfBirth" --category "Customer Details"
  python matching_engine.py --field "processingFee" --category "Loan Details"
  python matching_engine.py --field "aadharFrontLink" --category "Documents"
  python matching_engine.py --field "annualIncome" --entity "APPLICANT"
  python matching_engine.py --field "bureauScore" --category "Credit Details"
        """
    )
    parser.add_argument("--field",    "-f", type=str, required=True, help="Partner field name to match")
    parser.add_argument("--category", "-c", type=str, default=None,  help="Column category/UI grouping (optional)")
    parser.add_argument("--entity",   "-e", type=str, default=None,  help="Entity type: APPLICANT, COAPPLICANT, LOAN, DOCUMENT, FEE, OTHER")
    parser.add_argument("--refs",     "-r", type=str, default="./references", help="Path to references directory")
    parser.add_argument("--json",     "-j", action="store_true",     help="Output as JSON")
    args = parser.parse_args()

    try:
        refs = load_references(args.refs)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"ERROR loading references: {e}", file=sys.stderr)
        sys.exit(1)

    result = match_field(
        field_name=args.field,
        column_category=args.category,
        entity=args.entity,
        refs=refs,
    )

    if args.json:
        print(json.dumps(result.to_dict(), indent=2))
    else:
        print(f"Partner Field:     {result.partner_field}")
        print(f"Entity:            {result.entity}")
        if result.column_category:
            print(f"Category:          {result.column_category}")
        print(f"Match Type:        {result.match_type}")
        print(f"Matched Excel Key: {result.matched_excel_key or '(none)'}")
        if result.matched_json_key:
            print(f"Matched JSON Key:  {result.matched_json_key}")
        print(f"Confidence:        {result.confidence:.2f}")
        print(f"Reasoning:         {result.reasoning}")


if __name__ == "__main__":
    main()