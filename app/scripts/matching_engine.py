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
    if column_category and column_category in entity_routing:
        return entity_routing[column_category]

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


def _select_best_candidate(
    candidates: List[Tuple[str, str, str]],
    entity: str
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
    if len(candidates) == 1:
        return candidates[0][0], candidates[0][1]

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
    field_dictionary_index: Dict[str, List[Tuple[str, str, str]]]
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
        best_ek, _ = _select_best_candidate(field_dictionary_index[normalized_sub], current_entity)
        # Reduce confidence for entity substitution
        return best_ek, -0.10

    return None, 0.0


def match_field(
    field_name: str,
    column_category: Optional[str],
    entity: Optional[str],
    refs: Dict[str, Any]
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

    # ====== 1. EXACT MATCH ======
    # Try basic normalization first (preserves entity prefixes for precise matching)
    basic_normalized = normalize_basic(field_name)
    if basic_normalized in field_dict_index:
        candidates = field_dict_index[basic_normalized]
        excel_key, json_key = _select_best_candidate(candidates, entity)
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
            field_dict_index
        )

        final_excel_key = substituted_key or target_excel_key
        final_confidence = base_confidence + conf_adjustment

        return MatchResult(
            partner_field=field_name,
            column_category=column_category,
            matched_excel_key=final_excel_key,
            matched_json_key=target_json_key,
            confidence=final_confidence,
            match_type=match_type,
            reasoning=(
                f"Alias match ({tier}): '{field_name}' → '{target_excel_key}' "
                f"(frequency={frequency})" +
                (f"; entity substitution applied: {final_excel_key}" if substituted_key else "")
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

        result = match_field(field_name, column_category, entity, refs)

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
