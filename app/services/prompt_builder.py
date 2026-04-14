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

Prompt ownership:
  The base prompt template lives on the Dvara/Langfuse platform.
  When prompt_template="" (the normal case), fill_prompt() returns a
  standalone context block that is posted directly as the `task` field.
  The gateway merges it with the stored base template on its side.
"""

import logging
import json
import re
from typing import Any, Dict, List, Optional

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


def _available_keys_block(
    entity: str,
    field_dictionary: Dict[str, Any],
    category: Optional[str] = None,
    process_name: Optional[str] = None,
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

        line = ek
        if desc: line += f" | {desc}"
        if jk:   line += f" | {jk}"
        lines.append(line)

    return "\n".join(lines) or "(none)"


def _shortcuts_block(alias_registry: Dict[str, Any], limit: int = 30) -> str:
    forward = alias_registry.get("forward", {})
    rows = [
        (v.get("frequency", 0), k, v.get("target_excel_key", ""))
        for k, v in forward.items()
    ]
    rows.sort(key=lambda x: x[0], reverse=True)
    lines = [
        f"  {alias} → {target}  (seen in {freq} partners)"
        for freq, alias, target in rows[:limit]
    ]
    return "\n".join(lines) or "(none)"


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
        "GUARDRAILS:\n"
        "- You MUST emit exactly one mapping for every row under FIELDS TO MAP "
        "(same field_name spelling as given). Do not skip or merge rows.\n"
        "- Honor Entity scope: choose excel_key values that belong to that role "
        "(e.g. COAPPLICANT1* / LOANCOAPP* only when entity=COAPPLICANT; "
        "APPLICANT* when entity=APPLICANT; LOAN-level keys when entity=LOAN or FEE).\n"
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
        })

    return prompts


def build_loanparameter_refinement_prompts(
    loanparameter_rows: List[Dict[str, Any]],
    field_dictionary: Dict[str, Any],
    alias_registry: Dict[str, Any],
    client_name: str = "",
    process_name: str = "",
) -> List[Dict[str, Any]]:
    """
    After deterministic matching, rows stuck in LOANPARAMETER* (often from bad aliases)
    are sent to a dedicated gateway. Each prompt includes PUTM-scoped excel keys for
    that entity plus field context (category, entity, process_name).

    Returns the same shape as build_entity_prompts entries: entity, fields, rendered_prompt,
    entity_context (partner_field -> metadata for response parsing).
    """
    if not loanparameter_rows:
        return []

    groups: Dict[str, List[Dict[str, Any]]] = {}
    for row in loanparameter_rows:
        ent = (row.get("entity") or "OTHER").upper()
        groups.setdefault(ent, []).append(row)

    prompts: List[Dict[str, Any]] = []

    for entity, rows in groups.items():
        putm_block = _available_keys_block(
            entity,
            field_dictionary,
            category=None,
            process_name=process_name,
        )
        shortcuts = _shortcuts_block(alias_registry, limit=40)

        refine_lines: List[str] = []
        for r in rows:
            pf = (r.get("partner_field") or "").strip()
            if not pf:
                continue
            cat = (r.get("column_category") or "").strip()
            proc = (r.get("process_name") or "").strip() or (process_name or "")
            cur = (r.get("matched_excel_key") or "").strip()
            prev = ((r.get("reasoning") or "") + " " + (r.get("previous_mapping_reason") or "")).strip()
            if len(prev) > 600:
                prev = prev[:600] + "…"
            refine_lines.append(
                f"{pf} | column_category={cat} | process_name={proc} | entity={entity} "
                f"| current_bucket={cur} | prior_notes={prev}"
            )

        alias_warning = (
            "IMPORTANT: Semantic shortcuts (alias_registry) below are aggregated from many "
            "partners and may be WRONG for this file. Do not trust an alias if it disagrees "
            "with the PUTM catalog or the field semantics. Final mapped_excel_key must appear "
            "in PUTM_CANONICAL_EXCEL_KEYS for this process."
        )

        rendered = (
            "{\n"
            "TASK: REFINE_LOANPARAMETER_BUCKET_TO_PUTM_EXCEL_KEY\n"
            f"Client: {client_name or 'Unknown'}\n"
            f"Process (pipeline): {process_name or 'COMBINED'}\n"
            f"Entity scope: {entity}\n"
            "\n"
            f"{alias_warning}\n"
            "\n"
            "SEMANTIC_SHORTCUTS (alias_registry — verify against PUTM; may be incorrect):\n"
            f"{shortcuts}\n"
            "\n"
            "PUTM_CANONICAL_EXCEL_KEYS (you MUST choose mapped_excel_key from this list only):\n"
            f"{putm_block}\n"
            "\n"
            "FIELDS_CURRENTLY_MAPPED_TO_LOANPARAMETER_NEEDING_REFINEMENT:\n"
            f"{chr(10).join(refine_lines) if refine_lines else '(none)'}\n"
            "\n"
            "For each field row above, infer the best real PUTM excel_key using:\n"
            "- partner column name and meaning\n"
            "- column_category (sheet / semantic group)\n"
            "- entity (APPLICANT vs LOAN vs …) — the entity= token on EACH line is authoritative\n"
            "- process_name (origination vs enrollment vs combined)\n"
            "\n"
            "ENTITY / PUTM PREFIX RULES (mandatory):\n"
            "- If entity=APPLICANT or CUSTOMER: you MUST NOT choose COAPPLICANT* or LOANCOAPP* "
            "excel_keys; use APPLICANT* or LOANAPPLICANTPARAM* from the PUTM list only.\n"
            "- If entity starts with COAPPLICANT: prefer COAPPLICANT* / LOANCOAPP* keys over "
            "bare APPLICANT* when both exist.\n"
            "- If there is NO suitable PUTM excel_key for that row's entity in the list above "
            "(e.g. no applicant mobile when entity=APPLICANT): you MUST keep the same "
            "current_bucket LOANPARAMETER* value as matched_excel_key — do NOT substitute "
            "COAPPLICANT* or LOANCOAPP* keys as a workaround.\n"
            "\n"
            "OUTPUT: JSON object keyed by exact partner_field string; each value MUST be:\n"
            "  matched_excel_key|json_key|confidence|SEMANTIC|short_reasoning\n"
            "json_key may be empty (downstream fills from PUTM). confidence 0.0–1.0.\n"
            "You MUST return one entry for every partner_field listed above; "
            "if unsure or no entity-safe key exists, repeat current_bucket exactly with "
            "low confidence and explain.\n"
            "}"
        )

        entity_context: Dict[str, Dict[str, Any]] = {}
        for r in rows:
            pf = (r.get("partner_field") or "").strip()
            if not pf:
                continue
            entity_context[pf] = {
                "entity": entity,
                "column_category": r.get("column_category"),
                "process_name": (r.get("process_name") or "").strip() or (process_name or ""),
            }

        logger.info(
            "[build_loanparameter_refinement_prompts] entity=%s fields=%d",
            entity,
            len(entity_context),
        )

        prompts.append({
            "entity": entity,
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
