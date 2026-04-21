"""
services/prompt_builder.py
Fills Langfuse template variables from live field_dictionary + alias_registry data.
ENHANCED: Category-aware context building.

Langfuse prompt variables:
  {{entity}}               e.g. "APPLICANT"
  {{client_name}}          e.g. "Indifi"
  {{process_name}}         e.g. "COMBINED"
  {{available_excel_keys}} scoped field list — excel_key | description | json_path
  {{semantic_shortcuts}}   top-30 aliases — alias → excel_key (N partners)
  {{fields_to_map}}        unmatched fields — field_name | category
  {{structured_fee_putm_policy}}  fee PUTM leaves + notes from mapping_policy.json

Prompt ownership:
  The base prompt template lives on the Dvara/Langfuse platform.
  When prompt_template="" (the normal case), fill_prompt() returns a
  standalone context block that is posted directly as the `task` field.
  The gateway merges it with the stored base template on its side.
  Loanparameter / partner-field refinement: `build_loanparameter_refinement_prompts`
  posts JSON template variables only (client, process, PUTM list, FIELDS_LINES, …),
  not the full instruction text — instructions live in the refinement prompt on platform.
"""

import logging
import json
import re
from typing import Any, Dict, List, Optional, Tuple

from app.services.match_context import is_process_compatible

logger = logging.getLogger(__name__)

ENTITY_TO_ROLE = {
    "APPLICANT":    "CUSTOMER",
    "COAPPLICANT":  "COAPPLICANT",
    "COAPPLICANT1": "COAPPLICANT",
    "COAPPLICANT2": "COAPPLICANT",
    "COAPPLICANT3": "COAPPLICANT",
    "COAPPLICANT4": "COAPPLICANT",
    "GUARANTOR":    "GUARANTOR",
    "LOAN":         "LOAN",
    "DOCUMENT":     "LOAN",
    "FEE":          "LOAN",
    "OTHER":        "LOAN",
}

# Category semantic families for filtering available keys
CATEGORY_SEMANTIC_FAMILIES = {
    "date": {"keywords": {"date", "born", "birth", "dob", "since", "from", "to"}},
    "amount": {"keywords": {"amount", "value", "price", "cost", "fee", "charge", "salary", "income"}},
    "name": {"keywords": {"name", "first", "last", "middle", "full", "title", "label"}},
    "id": {"keywords": {"id", "code", "number", "ref", "reference", "identifier", "account", "pan", "aadhar"}},
    "status": {"keywords": {"status", "state", "flag", "type", "category", "active", "approved", "verified"}},
    "address": {"keywords": {"address", "street", "city", "state", "pin", "zip", "postal", "location"}},
    "contact": {"keywords": {"phone", "email", "contact", "mobile", "telephone", "number"}},
    "demographic": {"keywords": {"age", "gender", "caste", "religion", "marital", "occupation", "education"}},
    "financial": {"keywords": {"credit", "score", "debt", "loan", "balance", "account", "bank"}},
}

# Suppress very high-numbered noisy LOAN fields from the available list
_SUPPRESS = re.compile(
    r"^(LOANPARAMETER[5-9]\d|DOCUMENTNAME[6-9]|DOCUMENTID[6-9]|FEE[6-9])"
)

# column_category text signals (not excel_key hardcoding)
_COAPPLICANT_CATEGORY_RE = re.compile(
    r"co[\s\-]?applicant|co[\s\-]?borrower|co\s*borrower|\bcoapp\b|secondary\s+applicant",
    re.I,
)
_APPLICANT_IN_CATEGORY_RE = re.compile(r"\bapplicant\b", re.I)

# DOCUMENT row + identity / KYC — route to applicant-family PUTM list, not generic LOAN docs
_ID_KYC_FIELD_CATEGORY_RE = re.compile(
    r"kyc|identity|proof|document\s*id|id\s*(no|number|type)?|voter|passport|aadhaar|aadhar|pan\b|dl\b|driving",
    re.I,
)


def _norm(text: str) -> str:
    """Normalize text to lowercase with spaces."""
    text = text.lower()
    text = re.sub(r"[_\s.\-'()]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _extract_keywords(text: str) -> set:
    """Extract normalized keywords from text."""
    words = _norm(text).split()
    return set(w for w in words if len(w) > 2)


def _matches_category_family(excel_key: str, category: Optional[str]) -> bool:
    """Check if excel_key matches the semantic family of the category."""
    if not category:
        return True
    
    col_keywords = _extract_keywords(category)
    ek_keywords = _extract_keywords(excel_key)
    
    # Direct keyword overlap
    if col_keywords & ek_keywords:
        return True
    
    # Check semantic families
    for family, config in CATEGORY_SEMANTIC_FAMILIES.items():
        family_keywords = config["keywords"]
        col_has_family = any(kw in col_keywords for kw in family_keywords)
        ek_has_family = any(kw in ek_keywords for kw in family_keywords)
        if col_has_family and ek_has_family:
            return True
    
    return False


def _excel_key_is_coapplicant_prefix_family(ek: str) -> bool:
    """
    Structural check: co-applicant–scoped PUTM keys (prefix / LOANCOAPP / infix co-applicant slots).
    Used to filter PUTM lists and alias shortcuts — not individual business key names.
    """
    u = (ek or "").strip().upper()
    if not u:
        return False
    if u.startswith("COAPPLICANT") or u.startswith("LOANCOAPP"):
        return True
    if "COAPPLICANT" in u:
        return True
    return False


def _column_category_putm_signal(column_category: str) -> str:
    """
    Whether column_category text scopes applicant-only, co-applicant, or is neutral.
    Returns 'applicant' | 'coapplicant' | 'neutral'.
    """
    cat = (column_category or "").strip()
    if not cat:
        return "neutral"
    if _COAPPLICANT_CATEGORY_RE.search(cat):
        return "coapplicant"
    if _APPLICANT_IN_CATEGORY_RE.search(cat):
        return "applicant"
    return "neutral"


def _document_row_use_applicant_putm_family(
    partner_field: str,
    column_category: str,
) -> bool:
    """ENTITY=DOCUMENT but column/field clearly refer to applicant identity/KYC, not generic loan docs."""
    blob = f"{partner_field} {column_category}"
    return bool(_ID_KYC_FIELD_CATEGORY_RE.search(blob))


def _loanparameter_refinement_putm_prefix_scope(row: Dict[str, Any]) -> str:
    """
    Which excel_key prefix families may appear in PUTM for this row.
    Returns 'applicant_family' | 'coapplicant_family' | 'neutral'.
    """
    raw_ent = (row.get("entity") or "OTHER").strip().upper()
    cat = row.get("column_category") or ""
    pf = (row.get("partner_field") or row.get("field_name") or "").strip()
    sig = _column_category_putm_signal(cat)

    if sig == "applicant":
        return "applicant_family"
    if sig == "coapplicant":
        return "coapplicant_family"

    if raw_ent == "DOCUMENT" and _document_row_use_applicant_putm_family(pf, cat):
        return "applicant_family"
    if raw_ent.startswith("COAPPLICANT"):
        return "coapplicant_family"
    if raw_ent in ("APPLICANT", "CUSTOMER"):
        return "applicant_family"
    return "neutral"


def _entity_for_putm_catalog_block(prefix_scope: str, row_entity: str) -> str:
    """Map refinement batch to field_dictionary by_role list (APPLICANT→CUSTOMER, etc.)."""
    pe = (row_entity or "OTHER").strip().upper()
    if prefix_scope == "applicant_family":
        return "APPLICANT"
    if prefix_scope == "coapplicant_family":
        return pe if pe.startswith("COAPPLICANT") else "COAPPLICANT"
    return pe


def _putm_line_allowed_for_prefix_scope(line: str, prefix_scope: Optional[str]) -> bool:
    if not prefix_scope or prefix_scope == "neutral":
        return True
    ek = (line.split("|")[0] or "").strip()
    co = _excel_key_is_coapplicant_prefix_family(ek)
    if prefix_scope == "applicant_family":
        return not co
    if prefix_scope == "coapplicant_family":
        return co
    return True


_UDF_EXCEL_KEY_RE = re.compile(r"UDF\d+")


def _excel_key_is_udf_slot(ek: str) -> bool:
    """True for user-defined field slots: UDF1, COAPPLICANTUDF8, CUSTOMERUDF1, VEHICLEDOC1UDF1, …"""
    return bool(ek and _UDF_EXCEL_KEY_RE.search(ek))


def _available_keys_block(
    entity: str,
    field_dictionary: Dict[str, Any],
    category: Optional[str] = None,
    process_name: Optional[str] = None,
    putm_prefix_scope: Optional[str] = None,
    exclude_udf_keys: bool = False,
) -> str:
    """Generate available keys block, filtered by category and pipeline process_name."""
    role = ENTITY_TO_ROLE.get(entity, "LOAN")
    fields = field_dictionary.get("by_role", {}).get(role, [])
    lines = []
    for f in fields:
        ek   = f.get("excel_key", "")
        desc = (f.get("description", "") or "").strip()
        jk   = (f.get("json_key",   "") or "").strip()

        if role == "LOAN" and _SUPPRESS.match(ek):
            continue

        if process_name and not is_process_compatible(ek, process_name, field_dictionary):
            continue

        # Category filtering: prefer matches but include all if no category
        if category and not _matches_category_family(ek, category):
            continue

        if exclude_udf_keys and _excel_key_is_udf_slot(ek):
            continue

        line = ek
        if desc: line += f" | {desc}"
        if jk:   line += f" | {jk}"
        if putm_prefix_scope and not _putm_line_allowed_for_prefix_scope(line, putm_prefix_scope):
            continue
        lines.append(line)

    return "\n".join(lines) or "(none)"


def _shortcuts_block(
    alias_registry: Dict[str, Any],
    limit: int = 30,
    putm_prefix_scope: Optional[str] = None,
) -> str:
    forward = alias_registry.get("forward", {})
    rows = [
        (v.get("frequency", 0), k, v.get("target_excel_key", ""))
        for k, v in forward.items()
    ]
    rows.sort(key=lambda x: x[0], reverse=True)
    pool = rows[: max(limit * 8, 120)] if putm_prefix_scope and putm_prefix_scope != "neutral" else rows[:limit]
    lines = []
    for freq, alias, target in pool:
        tgt = (target or "").strip()
        if putm_prefix_scope and putm_prefix_scope != "neutral":
            if putm_prefix_scope == "applicant_family" and _excel_key_is_coapplicant_prefix_family(tgt):
                continue
            if putm_prefix_scope == "coapplicant_family" and not _excel_key_is_coapplicant_prefix_family(tgt):
                continue
        lines.append(f"  {alias} → {target}  (seen in {freq} partners)")
        if len(lines) >= limit:
            break
    return "\n".join(lines) or "(none)"


def _mapping_fee_policy_block(mapping_policy: Optional[Dict[str, Any]]) -> str:
    """
    Human-readable block from mapping_policy.json for LLM context (fee PUTM leaves).
    Keys list drives deterministic matching too (via load_references).
    """
    if not isinstance(mapping_policy, dict) or not mapping_policy:
        return ""
    lines: List[str] = []
    keys = mapping_policy.get("structured_fee_putm_base_keys")
    if isinstance(keys, list) and keys:
        cleaned = [str(k).strip().upper() for k in keys if str(k).strip()]
        if cleaned:
            lines.append(
                "STRUCTURED FEE PUTM BASE KEYS (mapping_policy.json → "
                "structured_fee_putm_base_keys): "
                + ", ".join(cleaned)
            )
    note = mapping_policy.get("llm_fee_mapping_note") or mapping_policy.get(
        "fee_field_instructions"
    )
    if isinstance(note, str) and note.strip():
        lines.append("FEE / CHARGE INSTRUCTIONS (mapping_policy.json):")
        lines.append(note.strip())
    leaves = mapping_policy.get("fee_putm_leaves")
    if isinstance(leaves, list) and leaves:
        lines.append("STRUCTURED FEE PUTM LEAVES (excel_key | json_key):")
        for row in leaves[:50]:
            if not isinstance(row, dict):
                continue
            ek = (row.get("excel_key") or row.get("base_key") or "").strip()
            jk = (row.get("json_key") or "").strip()
            if ek:
                lines.append(f"  {ek} | {jk}")
    return "\n".join(lines) if lines else ""


def _fields_block(fields: List[Dict[str, Any]]) -> str:
    lines = []
    for f in fields:
        name = f.get("field_name") or f.get("partner_field", "")
        cat  = f.get("column_category", "") or ""
        proc = f.get("process_name", "") or ""
        ent  = f.get("entity", "") or ""
        lines.append(f"{name} | {cat} | {proc} | {ent}")
    return "\n".join(lines) or "(none)"


def _candidate_matches_block(fields: List[Dict[str, Any]]) -> str:
    lines = []
    for f in fields:
        name = f.get("field_name") or f.get("partner_field", "")
        candidates = f.get("candidate_matches") or []
        if not name or not candidates:
            continue
        lines.append(f"{name}:")
        for candidate in candidates:
            engine = candidate.get("engine", "unknown")
            excel_key = candidate.get("matched_excel_key") or "(none)"
            json_key = candidate.get("json_key") or ""
            confidence = candidate.get("confidence", 0.0)
            reasoning = candidate.get("reasoning") or ""
            line = f"  {engine} | {excel_key} | confidence={confidence:.4f}"
            if json_key:
                line += f" | {json_key}"
            if reasoning:
                line += f" | {reasoning}"
            lines.append(line)
    return "\n".join(lines) or "(none)"


def _build_context_block(
    entity: str,
    fields: List[Dict[str, Any]],
    field_dictionary: Dict[str, Any],
    alias_registry: Dict[str, Any],
    client_name: str = "",
    process_name: str = "",
    pipeline_context_payload: Optional[Dict[str, Any]] = None,
    mapping_policy: Optional[Dict[str, Any]] = None,
) -> str:
    """
    Build the variable block that the gateway merges with the stored base template.
    ENHANCED: Now includes category context and per-field category suggestions.
    
    {
    Client: HDFC Bank
    Process: Home Loan Origination
    Entity scope: APPLICANT

    AVAILABLE INTERNAL excel_key VALUES:
    APPLICANTFIRSTNAME | First name of applicant | loanAccount.customer.firstName
    ...

    SEMANTIC SHORTCUTS:
    dateofbirth → APPLICANTDATEOFBIRTH (seen in 87 partners)
    ...

    FIELDS TO MAP (with category context):
    field_name | column_category
    [Suggestions for Applicant_Age (category=demographic): APPLICANTAGE, APPLICANTDATEOFBIRTH]
    ...

    LOW CONFIDENCE ENGINE SUGGESTIONS:
    ...
    }
    """
    available = _available_keys_block(
        entity, field_dictionary, category=None, process_name=process_name
    )
    shortcuts = _shortcuts_block(alias_registry)
    fields_to = _fields_block(fields)
    candidates = _candidate_matches_block(fields)

    filtered_payload: Optional[Dict[str, Any]] = None
    if isinstance(pipeline_context_payload, dict) and pipeline_context_payload:
        def _filter_list(key: str) -> List[Dict[str, Any]]:
            raw = pipeline_context_payload.get(key) or []
            if not isinstance(raw, list):
                return []
            return [
                item for item in raw
                if isinstance(item, dict)
                and (item.get("entity") or "").upper() == entity.upper()
            ]

        filtered_payload = {
            "unmatched_fields": _filter_list("unmatched_fields"),
            "deterministic_matches": _filter_list("deterministic_matches"),
            "fuzzy_matches": _filter_list("fuzzy_matches"),
            "embedding_matches": _filter_list("embedding_matches"),
            # Keep global LOANPARAMETER assignments visible to all entity prompts
            "loanparameter_assigned_fields": [
                item for item in (pipeline_context_payload.get("loanparameter_assigned_fields") or [])
                if isinstance(item, dict)
            ],
        }

    loanparameter_lines = []
    if filtered_payload and filtered_payload.get("loanparameter_assigned_fields"):
        for item in filtered_payload["loanparameter_assigned_fields"]:
            fname = item.get("field_name") or ""
            target = item.get("matched_target") or ""
            cat = item.get("column_category") or ""
            ent = (item.get("entity") or "").upper()
            proc = item.get("process_name") or ""
            if not target:
                continue
            loanparameter_lines.append(f"{fname} | {cat} | {proc} | {ent} | {target}")
    loanparameter_block = "\n".join(loanparameter_lines) or "(none)"

    payload_json_block = "(none)"
    if filtered_payload:
        try:
            payload_json_block = json.dumps(filtered_payload, indent=2, ensure_ascii=False)
        except (TypeError, ValueError):
            payload_json_block = "(unserializable payload)"

    # Build category-aware suggestions
    category_suggestions_lines = []
    for f in fields:
        field_name = f.get("field_name") or f.get("partner_field", "")
        category = f.get("column_category")
        if field_name and category:
            # Generate category-scoped available keys for this field
            category_available = _available_keys_block(
                entity, field_dictionary, category=category, process_name=process_name
            )
            if category_available and category_available != "(none)":
                # Take top 3 matches
                top_matches = category_available.split("\n")[:3]
                matches_str = ", ".join([m.split(" | ")[0] for m in top_matches])
                category_suggestions_lines.append(
                    f"  [{field_name} (category={category})] → {matches_str}"
                )

    category_suggestions = "\n".join(category_suggestions_lines) if category_suggestions_lines else "(none)"

    fee_policy_block = _mapping_fee_policy_block(mapping_policy)
    fee_policy_section = ""
    if fee_policy_block:
        fee_policy_section = (
            "FEE / CHARGE PUTM POLICY (from references/mapping_policy.json; "
            "same list as deterministic engine):\n"
            f"{fee_policy_block}\n"
            "\n"
        )

    return (
        "{\n"
        f"Client: {client_name or 'Unknown'}\n"
        f"Process: {process_name or 'COMBINED'}\n"
        f"Entity scope: {entity}\n"
        "\n"
        "AVAILABLE INTERNAL excel_key VALUES:\n"
        f"{available}\n"
        "\n"
        "SEMANTIC SHORTCUTS:\n"
        f"{shortcuts}\n"
        "\n"
        "STRUCTURED PIPELINE CONTEXT (JSON):\n"
        f"{payload_json_block}\n"
        "\n"
        "DETERMINISTIC LOANPARAMETER ASSIGNMENTS (already taken):\n"
        "field_name | column_category | process_name | entity | matched_target\n"
        f"{loanparameter_block}\n"
        "\n"
        f"{fee_policy_section}"
        "GUARDRAILS:\n"
        "- You MUST emit exactly one mapping for every row under FIELDS TO MAP "
        "(same field_name spelling as given). Do not skip or merge rows.\n"
        "- Honor Entity scope: choose excel_key values that belong to that role "
        "(e.g. COAPPLICANT1* / LOANCOAPP* only when entity=COAPPLICANT; "
        "APPLICANT* when entity=APPLICANT; LOAN-level keys when entity=LOAN or FEE).\n"
        "- When mapping fee/charge columns (entity FEE or fee-like semantics), prefer a "
        "concrete catalogue excel_key listed under FEE / CHARGE PUTM POLICY above over "
        "the generic excel_key FEE when one clearly fits.\n"
        "- Use column_category as disambiguation when several PUTM keys look similar.\n"
        "- Use deterministic/fuzzy/embedding context above as prior signals.\n"
        "- Do NOT duplicate LOANPARAMETER* targets already listed as taken.\n"
        "- Only reuse/remap a taken LOANPARAMETER* if you are more confident; explain why.\n"
        "\n"
        "FIELDS TO MAP (with category context):\n"
        "field_name | column_category | process_name | entity\n"
        f"{fields_to}\n"
        "\n"
        "CATEGORY-AWARE SUGGESTIONS:\n"
        f"{category_suggestions}\n"
        "\n"
        "LOW CONFIDENCE ENGINE SUGGESTIONS:\n"
        f"{candidates}\n"
        "}"
    )


def fill_prompt(
    template: str,
    entity: str,
    fields: List[Dict[str, Any]],
    field_dictionary: Dict[str, Any],
    alias_registry: Dict[str, Any],
    client_name: str = "",
    process_name: str = "",
    pipeline_context_payload: Optional[Dict[str, Any]] = None,
    mapping_policy: Optional[Dict[str, Any]] = None,
) -> str:
    """
    If template is provided: replace all {{variable}} placeholders and return
    the fully rendered prompt (legacy / local-template path).

    If template is empty: return the context block directly — this is posted
    as the `task` value and the gateway merges it with the stored base template.
    """
    if not template or not template.strip():
        return _build_context_block(
            entity=entity,
            fields=fields,
            field_dictionary=field_dictionary,
            alias_registry=alias_registry,
            client_name=client_name,
            process_name=process_name,
            pipeline_context_payload=pipeline_context_payload,
            mapping_policy=mapping_policy,
        )

    # Legacy path: fill {{variables}} into a local template
    variables = {
        "entity":               entity,
        "client_name":          client_name or "Unknown",
        "process_name":         process_name or "COMBINED",
        "available_excel_keys": _available_keys_block(
            entity, field_dictionary, category=None, process_name=process_name
        ),
        "semantic_shortcuts":   _shortcuts_block(alias_registry),
        "fields_to_map":        _fields_block(fields),
        "structured_fee_putm_policy": _mapping_fee_policy_block(mapping_policy)
        or "(none — add mapping_policy.json under references/)",
    }
    result = template
    for key, value in variables.items():
        result = result.replace("{{" + key + "}}", value)
    return result


def build_entity_prompts(
    unmatched_fields: List[Dict[str, Any]],
    field_dictionary: Dict[str, Any],
    alias_registry: Dict[str, Any],
    prompt_template: str,
    client_name: str = "",
    process_name: str = "",
    pipeline_context_payload: Optional[Dict[str, Any]] = None,
    mapping_policy: Optional[Dict[str, Any]] = None,
) -> List[Dict[str, Any]]:
    """
    Group unmatched fields by entity and build one filled prompt per group.

    Returns list of:
    {
      "entity": "APPLICANT",
      "fields": [...],
      "rendered_prompt": "...",          # context block posted as `task`
      "entity_context": {field_name: {entity, column_category}}
    }
    """
    # Group by entity
    groups: Dict[str, List[Dict]] = {}
    for f in unmatched_fields:
        entity = (f.get("entity") or "OTHER").upper()
        groups.setdefault(entity, []).append(f)

    prompts = []
    for entity, fields in groups.items():
        rendered = fill_prompt(
            template=prompt_template,
            entity=entity,
            fields=fields,
            field_dictionary=field_dictionary,
            alias_registry=alias_registry,
            client_name=client_name,
            process_name=process_name,
            pipeline_context_payload=pipeline_context_payload,
            mapping_policy=mapping_policy,
        )

        # Build lookup for response parsing
        entity_context = {
            (f.get("field_name") or f.get("partner_field", "")): {
                "entity":           entity,
                "column_category":  f.get("column_category"),
                "process_name":     f.get("process_name") or process_name or "",
            }
            for f in fields
        }

        logger.info(
            f"[build_entity_prompts] entity={entity}, "
            f"fields_sent_to_llm={list(entity_context.keys())}"
        )
        logger.debug(
            f"[build_entity_prompts] entity={entity}, "
            f"rendered_prompt_length={len(rendered)} chars"
        )

        prompts.append({
            "entity":          entity,
            "fields":          fields,
            "rendered_prompt": rendered,
            "entity_context":  entity_context,
            "prompt_template": prompt_template,
            "client_name":     client_name,
            "process_name":    process_name,
            "field_dictionary": field_dictionary,
            "alias_registry":   alias_registry,
            "pipeline_context_payload": pipeline_context_payload,
            "mapping_policy": mapping_policy or {},
        })

    return prompts


def build_loanparameter_refinement_prompts(
    loanparameter_rows: List[Dict[str, Any]],
    field_dictionary: Dict[str, Any],
    alias_registry: Dict[str, Any],
    client_name: str = "",
    process_name: str = "",
    mapping_policy: Optional[Dict[str, Any]] = None,
) -> List[Dict[str, Any]]:
    """
    After deterministic matching, rows stuck in LOANPARAMETER* are sent to the refinement
    gateway. Instructions live on the Dvara/Langfuse platform (REFINE_PARTNER_FIELDS…).

    This function only supplies template variables as JSON in ``rendered_prompt`` for the
    gateway to merge with the stored prompt. Keys: client, process, catalogue_entity,
    putm_prefix_scope, PUTM_CANONICAL_EXCEL_KEYS, SEMANTIC_SHORTCUTS, FIELDS_LINES,
    STRUCTURED_FEE_PUTM_BASE_KEYS, STRUCTURED_FEE_PUTM_LEAVES, FEE_PUTM_LLM_NOTE
    (from mapping_policy.json when provided; leaves are auto-derived from field_dictionary).

    UDF slot keys are excluded from PUTM_CANONICAL_EXCEL_KEYS. ``alias_registry`` is kept
    for API compatibility but not embedded here.

    All LOANPARAMETER* rows are included: batches are split by (putm_prefix_scope,
    catalogue_entity) so each row is sent in exactly one gateway call with an appropriate
    PUTM catalogue slice.

    Returns the same shape as build_entity_prompts entries: entity, fields, rendered_prompt,
    entity_context (partner_field -> metadata for response parsing).
    """
    if not loanparameter_rows:
        return []

    groups: Dict[Tuple[str, str], List[Dict[str, Any]]] = {}
    for row in loanparameter_rows:
        pfx = _loanparameter_refinement_putm_prefix_scope(row)
        raw_ent = (row.get("entity") or "OTHER").strip().upper()
        ent_block = _entity_for_putm_catalog_block(pfx, raw_ent)
        groups.setdefault((pfx, ent_block), []).append(row)

    _ = alias_registry  # API compatibility; shortcuts are not sent (platform owns wording).

    mp = mapping_policy if isinstance(mapping_policy, dict) else {}
    fee_keys = mp.get("structured_fee_putm_base_keys")
    fee_key_line = ""
    if isinstance(fee_keys, list) and fee_keys:
        fee_key_line = ", ".join(str(k).strip().upper() for k in fee_keys if str(k).strip())
    fee_note = (
        (mp.get("llm_fee_mapping_note") or mp.get("fee_field_instructions") or "").strip()
    )
    fee_leaf_lines: List[str] = []
    raw_leaves = mp.get("fee_putm_leaves")
    if isinstance(raw_leaves, list):
        for row in raw_leaves[:60]:
            if not isinstance(row, dict):
                continue
            ek = (row.get("excel_key") or row.get("base_key") or "").strip()
            jk = (row.get("json_key") or "").strip()
            if ek:
                fee_leaf_lines.append(f"{ek} | {jk}")
    fee_leaves_block = "\n".join(fee_leaf_lines) if fee_leaf_lines else "(none)"

    prompts: List[Dict[str, Any]] = []

    for (pfx, ent_block), rows in groups.items():
        putm_scope_arg = pfx if pfx != "neutral" else None
        putm_block = _available_keys_block(
            ent_block,
            field_dictionary,
            category=None,
            process_name=process_name,
            putm_prefix_scope=putm_scope_arg,
            exclude_udf_keys=True,
        )

        refine_lines: List[str] = []
        for r in rows:
            pf = (r.get("partner_field") or "").strip()
            if not pf:
                continue
            cat = (r.get("column_category") or "").strip()
            proc = (r.get("process_name") or "").strip() or (process_name or "")
            cur = (r.get("matched_excel_key") or "").strip()
            det = (r.get("entity") or "OTHER").strip().upper()
            row_scope = _loanparameter_refinement_putm_prefix_scope(r)
            prev = ((r.get("reasoning") or "") + " " + (r.get("previous_mapping_reason") or "")).strip()
            if len(prev) > 600:
                prev = prev[:600] + "…"
            refine_lines.append(
                f"{pf} | detected_entity={det} | column_category={cat} | process_name={proc} "
                f"| putm_prefix_scope={row_scope} | current_match={cur} | prior_notes={prev}"
            )

        # Variable bundle for Langfuse / Dvara — full instructions live on the platform only.
        variables: Dict[str, Any] = {
            "client": client_name or "Unknown",
            "process": process_name or "COMBINED",
            "catalogue_entity": ent_block,
            "putm_prefix_scope": pfx,
            "PUTM_CANONICAL_EXCEL_KEYS": putm_block,
            "SEMANTIC_SHORTCUTS": (
                "Not used — map only from PUTM_CANONICAL_EXCEL_KEYS; do not use alias shortcuts."
            ),
            "REFINEMENT_OUTPUT_RULE": (
                "For each field in FIELDS_LINES: output a concrete PUTM excel_key from "
                "PUTM_CANONICAL_EXCEL_KEYS only when there is a clear match. "
                "If no concrete key fits, omit that field from the mapping output or keep the "
                "same current_match LOANPARAMETER*n value — do not invent another generic "
                "LOANPARAMETER* bucket as the target."
            ),
            "FIELDS_LINES": "\n".join(refine_lines) if refine_lines else "(none)",
            "STRUCTURED_FEE_PUTM_BASE_KEYS": fee_key_line or "(none — configure mapping_policy.json)",
            "STRUCTURED_FEE_PUTM_LEAVES": fee_leaves_block,
            "FEE_PUTM_LLM_NOTE": fee_note or "(none)",
        }
        rendered = json.dumps(variables, ensure_ascii=False, indent=2)

        entity_context: Dict[str, Dict[str, Any]] = {}
        for r in rows:
            pf = (r.get("partner_field") or "").strip()
            if not pf:
                continue
            entity_context[pf] = {
                "entity": (r.get("entity") or "OTHER").strip().upper(),
                "column_category": r.get("column_category"),
                "process_name": (r.get("process_name") or "").strip() or (process_name or ""),
                "putm_prefix_scope": _loanparameter_refinement_putm_prefix_scope(r),
            }

        logger.info(
            "[build_loanparameter_refinement_prompts] ent_block=%s pfx=%s fields=%d",
            ent_block,
            pfx,
            len(entity_context),
        )

        prompts.append({
            "entity": ent_block,
            "fields": rows,
            "rendered_prompt": rendered,
            "entity_context": entity_context,
        })

    return prompts


def build_parameter_classifier_prompt(
    fields: List[Dict[str, Any]],
    client_name: str = "",
    process_name: str = "",
) -> Dict[str, Any]:
    """
    Build a single context block for the parameter bucket classifier gateway.

    Each row format:
      partner_field | column_category | detected_entity | current_match | current_reasoning
    """
    lines = []
    entity_context: Dict[str, Dict[str, Any]] = {}

    for field in fields:
        partner_field = field.get("partner_field") or field.get("field_name") or ""
        if not partner_field:
            continue
        column_category = field.get("column_category") or ""
        entity = field.get("entity") or "OTHER"
        current_match = field.get("matched_excel_key") or "unmatched"
        current_reasoning = field.get("reasoning") or ""
        proc = field.get("process_name") or ""
        lines.append(
            f"{partner_field} | {column_category} | {entity} | {proc} | {current_match} | {current_reasoning}"
        )
        entity_context[partner_field] = {
            "entity": entity,
            "column_category": field.get("column_category"),
            "process_name": proc,
            "matched_excel_key": field.get("matched_excel_key"),
            "reasoning": field.get("reasoning", ""),
        }

    rendered = (
        "{\n"
        f"Client: {client_name or 'Unknown'}\n"
        f"Process: {process_name or 'COMBINED'}\n"
        "FIELDS TO CLASSIFY:\n"
        f"{chr(10).join(lines) if lines else '(none)'}\n"
        "}"
    )

    logger.info(
        "[build_parameter_classifier_prompt] fields_sent_to_classifier=%s",
        list(entity_context.keys()),
    )
    return {
        "entity": "PARAMETER_CLASSIFIER",
        "fields": fields,
        "rendered_prompt": rendered,
        "entity_context": entity_context,
    }


def build_entity_classifier_prompt(
    fields: List[Dict[str, Any]],
    client_name: str = "",
    process_name: str = "",
) -> Dict[str, Any]:
    """
    Context block for ENTITY_CLASSIFIER_GATEWAY_URL (optional pre-step before matching).

    The Langfuse/Dvara *base* prompt lives outside this repo (see
    entity_classifier_langfuse_system_prompt.txt at package root for paste-in text).
    This function only builds the `task` payload: Client, Process, INPUT_ROWS.

    Expected model output: JSON keyed by row index strings; each value
    ENTITY|confidence|reasoning (0.0–1.0), validated in llm_service.
    """
    lines: List[str] = []
    # row_index (string) -> metadata for merge + validation
    entity_context: Dict[str, Dict[str, Any]] = {}

    for idx, field in enumerate(fields):
        fn = (field.get("field_name") or "").strip()
        if not fn:
            continue
        key = str(idx)
        cat = (field.get("column_category") or "").strip()
        sheet = (field.get("source_sheet") or "").strip()
        lines.append(f"{key}: {fn} | {cat} | {sheet}")
        entity_context[key] = {
            "field_name": fn,
            "column_category": field.get("column_category"),
            "source_sheet": field.get("source_sheet"),
            "list_index": idx,
        }

    rendered = (
        "{\n"
        f"Client: {client_name or 'Unknown'}\n"
        f"Process: {process_name or 'COMBINED'}\n"
        "TASK: ENTITY_CLASSIFIER (indices match INPUT_ROWS below)\n"
        "INPUT_ROWS (one partner column per line; index is the authoritative key for your output):\n"
        f"{chr(10).join(lines) if lines else '(none)'}\n"
        "}"
    )

    logger.info(
        "[build_entity_classifier_prompt] rows=%d indices=%s",
        len(lines),
        list(entity_context.keys()),
    )
    return {
        "entity": "ENTITY_CLASSIFIER",
        "fields": fields,
        "rendered_prompt": rendered,
        "entity_context": entity_context,
    }
