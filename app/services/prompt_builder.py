"""
services/prompt_builder.py
Fills Langfuse template variables from live field_dictionary + alias_registry data.

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
import re
from typing import Any, Dict, List

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

# Suppress very high-numbered noisy LOAN fields from the available list
_SUPPRESS = re.compile(
    r"^(LOANPARAMETER[5-9]\d|DOCUMENTNAME[6-9]|DOCUMENTID[6-9]|FEE[6-9])"
)


def _available_keys_block(entity: str, field_dictionary: Dict[str, Any]) -> str:
    role = ENTITY_TO_ROLE.get(entity, "LOAN")
    fields = field_dictionary.get("by_role", {}).get(role, [])
    lines = []
    for f in fields:
        ek   = f.get("excel_key", "")
        desc = (f.get("description", "") or "").strip()
        jk   = (f.get("json_key",   "") or "").strip()
        if role == "LOAN" and _SUPPRESS.match(ek):
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
        lines.append(f"{name} | {cat}" if cat else name)
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
) -> str:
    """
    Build the variable block that the gateway merges with the stored base template.
    This is exactly what Postman showed as the `task` value:

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

      FIELDS TO MAP:
      field_name | column_category
      ...
      }
    """
    available = _available_keys_block(entity, field_dictionary)
    shortcuts = _shortcuts_block(alias_registry)
    fields_to = _fields_block(fields)
    candidates = _candidate_matches_block(fields)

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
        "FIELDS TO MAP:\n"
        f"{fields_to}\n"
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
        )

    # Legacy path: fill {{variables}} into a local template
    variables = {
        "entity":               entity,
        "client_name":          client_name or "Unknown",
        "process_name":         process_name or "COMBINED",
        "available_excel_keys": _available_keys_block(entity, field_dictionary),
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
        )

        # Build lookup for response parsing
        entity_context = {
            (f.get("field_name") or f.get("partner_field", "")): {
                "entity":          entity,
                "column_category": f.get("column_category"),
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
        lines.append(
            f"{partner_field} | {column_category} | {entity} | {current_match} | {current_reasoning}"
        )
        entity_context[partner_field] = {
            "entity": entity,
            "column_category": field.get("column_category"),
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
