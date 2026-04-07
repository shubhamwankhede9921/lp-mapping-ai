"""
services/llm_service.py
Calls the Dvara gateway and parses the response.

Gateway contract (confirmed from Postman):
  POST form-data
    task = <rendered context block>
           ── one field, one value ──
           The gateway fetches the base prompt template from the platform
           and merges it with this context block server-side.
           No base template text is sent from this service.

Value block format (what goes into `task`):
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

Response envelope:
  {
    "status": "completed",
    "workflow_name": "...",
    "workflow_id": "...",
    "result": {
      "result": "```json\\n{ ... }\\n```"
    }
  }

LLM output format (flat pipe-delimited):
  {
    "Date_of_Birth": "APPLICANTDATEOFBIRTH|loanAccount.customer.dateOfBirth|0.97|SEMANTIC|reason",
    "Aadhar_Front":  "DOCUMENTNAME||0.90|PATTERN|aadhaar keyword"
  }
  Each value = matched_excel_key | json_key | confidence | matched_pattern | reasoning
"""

import json
import logging
import re
from typing import Any, Dict, List, Optional

import requests

logger = logging.getLogger(__name__)

# Keys that are NEVER partner field names — gateway envelope keys
ENVELOPE_KEYS = {"result", "error", "status", "message", "data", "output", "response"}


# ── Helpers ────────────────────────────────────────────────────────────────────

def _normalize_key(key: str) -> str:
    return re.sub(r"[\s_\-\.]", "", key).lower()


def _strip_markdown(text: str) -> str:
    if not isinstance(text, str):
        return text
    stripped = text.strip()
    if stripped.startswith("```"):
        stripped = re.sub(r"^```(?:json)?\s*", "", stripped, flags=re.MULTILINE)
        stripped = re.sub(r"\s*```\s*$",       "", stripped, flags=re.MULTILINE)
    return stripped.strip()


def _is_mapping_dict(d: dict, min_pipe_count: int = 4) -> bool:
    if not d:
        return False
    non_env = [k for k in d if k not in ENVELOPE_KEYS]
    if not non_env:
        return False
    samples = [d[k] for k in non_env[:5] if isinstance(d[k], str)]
    if not samples:
        return False
    return all(v.count("|") >= min_pipe_count for v in samples)


def _unwrap_envelope(
    data: Any,
    depth: int = 0,
    min_pipe_count: int = 4,
) -> Optional[dict]:
    if depth > 6:
        return None
    if isinstance(data, str):
        cleaned = _strip_markdown(data)
        try:
            return _unwrap_envelope(
                json.loads(cleaned),
                depth + 1,
                min_pipe_count=min_pipe_count,
            )
        except (json.JSONDecodeError, ValueError):
            return None
    if isinstance(data, list):
        return (
            _unwrap_envelope(data[0], depth + 1, min_pipe_count=min_pipe_count)
            if data else None
        )
    if not isinstance(data, dict):
        return None
    if _is_mapping_dict(data, min_pipe_count=min_pipe_count):
        logger.info(f"✓ Found mapping dict at depth={depth} with {len(data)} keys")
        return data
    if "mappings" in data and isinstance(data["mappings"], list):
        return data
    for key in ("result", "data", "output", "response", "content", "text"):
        if key in data:
            found = _unwrap_envelope(
                data[key],
                depth + 1,
                min_pipe_count=min_pipe_count,
            )
            if found is not None:
                return found
    for key, value in data.items():
        if key not in ENVELOPE_KEYS:
            found = _unwrap_envelope(
                value,
                depth + 1,
                min_pipe_count=min_pipe_count,
            )
            if found is not None:
                return found
    return None


# ── Response parsers ───────────────────────────────────────────────────────────

def _parse_pipe_value(
    partner_field: str,
    value_str: str,
    entity_context: Dict[str, Dict],
) -> Optional[Dict[str, Any]]:
    parts = value_str.split("|", 4)
    if len(parts) != 5:
        logger.warning(
            f"Skipping malformed value for '{partner_field}': "
            f"{value_str!r} (expected 5 parts, got {len(parts)})"
        )
        return None
    matched_excel_key, json_key, confidence_str, matched_pattern, reasoning = parts
    try:
        confidence = float(confidence_str.strip())
    except ValueError:
        logger.warning(f"Invalid confidence '{confidence_str}' for '{partner_field}', defaulting to 0.0")
        confidence = 0.0
    ctx = entity_context.get(partner_field, {})
    return {
        "partner_field":     partner_field,
        "column_category":   ctx.get("column_category"),
        "entity":            ctx.get("entity", "OTHER"),
        "matched_excel_key": matched_excel_key.strip() or None,
        "json_key":          json_key.strip(),
        "confidence":        round(confidence, 4),
        "match_type":        f"llm_{matched_pattern.strip().lower()}",
        "reasoning":         reasoning.strip(),
        "needs_review":      confidence < 0.80,
        "winning_engine":    "llm",
    }


def _parse_structured_item(
    item: Dict[str, Any],
    entity_context: Dict[str, Dict],
) -> Optional[Dict[str, Any]]:
    partner_field = item.get("client_column", "")
    if not partner_field:
        return None
    ctx        = entity_context.get(partner_field, {})
    confidence = float(item.get("confidence", 0.0))
    return {
        "partner_field":     partner_field,
        "column_category":   ctx.get("column_category") or item.get("column_category"),
        "entity":            ctx.get("entity", "OTHER"),
        "matched_excel_key": (item.get("matched_excel_key") or "").strip() or None,
        "json_key":          (item.get("json_key") or "").strip(),
        "confidence":        round(confidence, 4),
        "match_type":        f"llm_{(item.get('matched_pattern') or 'semantic').strip().lower()}",
        "reasoning":         (item.get("reasoning") or "").strip(),
        "needs_review":      confidence < 0.80,
        "winning_engine":    "llm",
    }


# ── Main service ───────────────────────────────────────────────────────────────

class LLMService:

    def __init__(self, settings):
        self.url        = settings.llm_gateway_url
        self.parameter_classifier_url = getattr(settings, "parameter_classifier_gateway_url", "")
        self.token      = settings.llm_gateway_token
        # Form-data field name — confirmed "task" from Postman, configurable via .env
        self.task_field = getattr(settings, "llm_task_field", "task")

    # ── Send ──────────────────────────────────────────────────────────────────

    def _send(
        self,
        rendered_context: str,
        url: Optional[str] = None,
        min_pipe_count: int = 4,
    ) -> dict:
        """
        POST rendered_context to the gateway as form-data:
          task = <rendered_context>

        The gateway reads this field, fetches the stored base prompt from
        the platform, merges them, and returns LLM output.
        We never send the base template — the platform owns it.
        """
        headers = {"Authorization": f"Bearer {self.token}"}
        payload = {self.task_field: rendered_context}
        target_url = url or self.url

        logger.info(
            f"→ Gateway POST  url={target_url!r}  field={self.task_field!r}  "
            f"value_length={len(rendered_context)} chars"
        )

        try:
            resp = requests.post(target_url, headers=headers, data=payload, timeout=600)
            resp.raise_for_status()
            body = resp.json()
        except requests.RequestException as e:
            logger.error(f"Gateway request failed: {e}")
            raise
        except json.JSONDecodeError as e:
            logger.error(f"Gateway returned non-JSON: {e}")
            raise

        logger.debug(
            f"Gateway top-level keys: "
            f"{list(body.keys()) if isinstance(body, dict) else type(body)}"
        )

        mapping_dict = _unwrap_envelope(body, min_pipe_count=min_pipe_count)
        if mapping_dict is None:
            raise ValueError(
                f"Could not find mapping dict in gateway response: {str(body)[:300]}"
            )
        return mapping_dict

    # ── Parse ─────────────────────────────────────────────────────────────────

    def _parse(
        self,
        mapping_dict: dict,
        entity_context: Dict[str, Dict],
    ) -> List[Dict[str, Any]]:
        allowed_normalized: Dict[str, str] = {
            _normalize_key(k): k for k in entity_context
        }
        logger.info(f"Whitelist: {list(allowed_normalized.keys())}")
        results: List[Dict[str, Any]] = []

        # Structured array format: {"mappings": [...]}
        if "mappings" in mapping_dict and isinstance(mapping_dict["mappings"], list):
            logger.info(f"Structured format — {len(mapping_dict['mappings'])} items")
            for item in mapping_dict["mappings"]:
                partner_field = item.get("client_column", "")
                norm = _normalize_key(partner_field)
                if entity_context and norm not in allowed_normalized:
                    logger.warning(f"Hallucinated field '{partner_field}' — skipping")
                    continue
                item = {**item, "client_column": allowed_normalized.get(norm, partner_field)}
                m = _parse_structured_item(item, entity_context)
                if m:
                    results.append(m)
            return results

        # Flat pipe-delimited format
        logger.info(f"Flat format — {len(mapping_dict)} keys")
        for partner_field, value_str in mapping_dict.items():
            if partner_field in ENVELOPE_KEYS:
                continue
            norm = _normalize_key(partner_field)
            if entity_context and norm not in allowed_normalized:
                logger.warning(f"Hallucinated field '{partner_field}' — skipping")
                continue
            original_key = allowed_normalized.get(norm, partner_field)
            if not isinstance(value_str, str):
                logger.warning(f"Skipping '{partner_field}' — value is {type(value_str).__name__}")
                continue
            if value_str.count("|") < 4:
                logger.warning(f"Skipping '{partner_field}' — only {value_str.count('|')} pipes (need 4)")
                continue
            m = _parse_pipe_value(original_key, value_str, entity_context)
            if m:
                results.append(m)

        logger.info(f"Parsed {len(results)} mappings")
        return results

    def _parse_parameter_bucket_response(
        self,
        mapping_dict: dict,
        entity_context: Dict[str, Dict],
    ) -> List[Dict[str, Any]]:
        allowed_normalized: Dict[str, str] = {
            _normalize_key(k): k for k in entity_context
        }
        results: List[Dict[str, Any]] = []

        logger.info(
            "Parameter classifier whitelist: %s",
            list(allowed_normalized.keys()),
        )

        for partner_field, value_str in mapping_dict.items():
            if partner_field in ENVELOPE_KEYS:
                continue
            norm = _normalize_key(partner_field)
            if entity_context and norm not in allowed_normalized:
                logger.warning(
                    "Parameter classifier hallucinated field '%s' — skipping",
                    partner_field,
                )
                continue
            original_key = allowed_normalized.get(norm, partner_field)
            if not isinstance(value_str, str):
                logger.warning(
                    "Skipping classifier value for '%s' — value is %s",
                    partner_field,
                    type(value_str).__name__,
                )
                continue

            parts = value_str.split("|", 2)
            if len(parts) != 3:
                logger.warning(
                    "Skipping malformed classifier value for '%s': %r",
                    partner_field,
                    value_str,
                )
                continue

            bucket, confidence_str, reasoning = parts
            bucket = bucket.strip().upper()
            if bucket not in {"LOANPARAMETER", "LOANAPPLICANTPARAM"}:
                logger.warning(
                    "Skipping classifier bucket '%s' for '%s'",
                    bucket,
                    partner_field,
                )
                continue
            try:
                confidence = float(confidence_str.strip())
            except ValueError:
                confidence = 0.0

            ctx = entity_context.get(original_key, {})
            results.append({
                "partner_field": original_key,
                "column_category": ctx.get("column_category"),
                "entity": ctx.get("entity", "OTHER"),
                "matched_excel_key": bucket,
                "json_key": "",
                "confidence": round(confidence, 4),
                "match_type": "llm_parameter_bucket",
                "reasoning": reasoning.strip(),
                "needs_review": confidence < 0.80,
                "winning_engine": "llm_parameter_bucket",
            })

        logger.info("Parsed %d parameter bucket classifications", len(results))
        return results

    # ── Public entry point ────────────────────────────────────────────────────

    def map_fields(self, entity_prompts: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        One gateway call per entity group.
        entity_prompt["rendered_prompt"] is the filled context block —
        it goes straight into the `task` form-data field.
        """
        all_mappings: List[Dict[str, Any]] = []

        for i, ep in enumerate(entity_prompts, 1):
            entity         = ep["entity"]
            rendered       = ep["rendered_prompt"]
            entity_context = ep["entity_context"]

            logger.info(
                f"[{i}/{len(entity_prompts)}] entity={entity}  "
                f"fields={len(ep['fields'])}  "
                f"context_keys={list(entity_context.keys())}"
            )

            try:
                mapping_dict = self._send(rendered)
                mappings     = self._parse(mapping_dict, entity_context)
                logger.info(f"  → {len(mappings)} mappings for entity={entity}")
                all_mappings.extend(mappings)

            except Exception as e:
                logger.error(f"Failed for entity={entity}: {e}", exc_info=True)
                for field_name, ctx in entity_context.items():
                    all_mappings.append({
                        "partner_field":     field_name,
                        "column_category":   ctx.get("column_category"),
                        "entity":            ctx.get("entity", entity),
                        "matched_excel_key": None,
                        "json_key":          "",
                        "confidence":        0.0,
                        "match_type":        "llm_error",
                        "reasoning":         f"LLM call failed: {e}",
                        "needs_review":      True,
                        "winning_engine":    "none",
                    })

        return all_mappings

    def classify_parameter_buckets(self, prompt_payload: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Single gateway call to classify rows into LOANPARAMETER vs LOANAPPLICANTPARAM.
        """
        if not self.parameter_classifier_url:
            logger.info("Parameter classifier URL not configured — skipping classifier step")
            return []

        rendered = prompt_payload["rendered_prompt"]
        entity_context = prompt_payload["entity_context"]
        logger.info(
            "[parameter-classifier] fields=%d context_keys=%s",
            len(prompt_payload.get("fields", [])),
            list(entity_context.keys()),
        )

        mapping_dict = self._send(
            rendered,
            url=self.parameter_classifier_url,
            min_pipe_count=2,
        )
        return self._parse_parameter_bucket_response(mapping_dict, entity_context)
