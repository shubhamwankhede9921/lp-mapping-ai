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
from collections import deque
from typing import Any, Dict, List, Optional

import requests

from app.services.prompt_builder import fill_prompt

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


def _is_benign_gateway_message(text: str) -> bool:
    """
    Dvara gateway sometimes puts status text in `message` / `error`-shaped fields,
    e.g. \"Workflow 'field_mapping' executed successfully\" — not a failure.
    """
    t = (text or "").strip().lower()
    if not t:
        return True
    if "fail" in t or "error" in t or "invalid" in t or "unauthorized" in t:
        return False
    if any(m in t for m in ("executed successfully", "completed successfully")):
        return True
    if "successfully" in t and "workflow" in t:
        return True
    if t in {"success", "ok", "true"}:
        return True
    return False


def _deep_find_pipe_mapping_dict(
    data: Any,
    min_pipe_count: int = 4,
    max_depth: int = 10,
    max_nodes: int = 500,
) -> Optional[dict]:
    """
    Breadth-first search for a dict of partner_field → pipe-delimited strings.
    Used when the gateway nests the LLM JSON under extra workflow metadata so
    the shallow _unwrap_envelope walk misses it.
    """
    queue: deque = deque([(data, 0)])
    seen_ids: set = set()
    nodes = 0
    while queue and nodes < max_nodes:
        node, depth = queue.popleft()
        nodes += 1
        if depth > max_depth:
            continue
        if isinstance(node, dict):
            i = id(node)
            if i in seen_ids:
                continue
            seen_ids.add(i)
            extracted = _extract_pipe_mapping_dict(node, min_pipe_count=min_pipe_count)
            if extracted:
                logger.info(
                    "Deep search found mapping dict at depth=%d with %d keys",
                    depth,
                    len(extracted),
                )
                return extracted
            if "mappings" in node and isinstance(node["mappings"], list):
                return node
            for v in node.values():
                queue.append((v, depth + 1))
        elif isinstance(node, list):
            i = id(node)
            if i in seen_ids:
                continue
            seen_ids.add(i)
            for v in node:
                queue.append((v, depth + 1))
        elif isinstance(node, str):
            cleaned = _strip_markdown(node)
            if "|" not in cleaned or len(cleaned) < 15:
                continue
            try:
                parsed = json.loads(cleaned)
            except (json.JSONDecodeError, ValueError):
                continue
            queue.append((parsed, depth + 1))
    return None


def _is_mapping_dict(d: dict, min_pipe_count: int = 4) -> bool:
    return _extract_pipe_mapping_dict(d, min_pipe_count=min_pipe_count) is not None


def _extract_pipe_mapping_dict(
    d: dict,
    min_pipe_count: int = 4,
) -> Optional[dict]:
    if not d:
        return None

    extracted = {
        key: value
        for key, value in d.items()
        if key not in ENVELOPE_KEYS
        and isinstance(value, str)
        and value.count("|") >= min_pipe_count
    }
    return extracted or None


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
    extracted = _extract_pipe_mapping_dict(data, min_pipe_count=min_pipe_count)
    if extracted is not None:
        logger.info(
            f"✓ Found mapping dict at depth={depth} with {len(extracted)} keys"
        )
        return extracted
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


def _gateway_error_message(body: Any) -> Optional[str]:
    if not isinstance(body, dict):
        return None

    def _take(msg: Any) -> Optional[str]:
        if not isinstance(msg, str):
            return None
        s = msg.strip()
        if not s or _is_benign_gateway_message(s):
            return None
        return s

    for key in ("error", "message"):
        picked = _take(body.get(key))
        if picked:
            return picked

    result = body.get("result")
    if isinstance(result, dict):
        for key in ("error", "message"):
            picked = _take(result.get(key))
            if picked:
                return picked

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
        "process_name":      ctx.get("process_name") or "",
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
        "process_name":      ctx.get("process_name") or item.get("process_name") or "",
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
        self.loanparameter_refinement_url = getattr(
            settings, "loanparameter_refinement_gateway_url", ""
        ) or ""
        self.entity_classifier_url = (
            getattr(settings, "entity_classifier_gateway_url", "") or ""
        ).strip()
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
            mapping_dict = _deep_find_pipe_mapping_dict(
                body, min_pipe_count=min_pipe_count
            )
        if mapping_dict is None:
            gateway_error = _gateway_error_message(body)
            if gateway_error:
                raise ValueError(f"Gateway returned error: {gateway_error}")
            keys_hint = list(body.keys()) if isinstance(body, dict) else type(body)
            logger.warning(
                "Could not extract mapping dict; top-level keys=%r snippet=%r",
                keys_hint,
                str(body)[:400],
            )
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

    def _parse_entity_classifier_response(
        self,
        mapping_dict: dict,
        entity_context: Dict[str, Dict[str, Any]],
    ) -> Dict[str, Dict[str, Any]]:
        """
        Values: ENTITY|confidence|reasoning (2+ pipes). Keys: row index strings \"0\", \"1\", …
        """

        def _valid_entity_label(ent: str) -> bool:
            e = (ent or "").strip().upper()
            if not e or len(e) > 48:
                return False
            if not re.match(r"^[A-Z][A-Z0-9]*$", e):
                return False
            if re.match(r"^COAPPLICANT\d+$", e):
                return True
            if re.match(r"^GUARANTOR\d+$", e):
                return True
            return e in {
                "APPLICANT",
                "CUSTOMER",
                "COAPPLICANT",
                "GUARANTOR",
                "LOAN",
                "DOCUMENT",
                "FEE",
                "OTHER",
            }

        out: Dict[str, Dict[str, Any]] = {}
        for raw_key, value_str in mapping_dict.items():
            if raw_key in ENVELOPE_KEYS:
                continue
            if raw_key not in entity_context:
                logger.debug("Entity classifier key %r not in context — skipping", raw_key)
                continue
            if not isinstance(value_str, str):
                continue
            if value_str.count("|") < 2:
                logger.warning(
                    "Entity classifier value for %r needs 2+ pipes: %r",
                    raw_key,
                    value_str,
                )
                continue
            parts = value_str.split("|", 2)
            if len(parts) != 3:
                continue
            ent, conf_s, reason = parts[0].strip(), parts[1].strip(), parts[2].strip()
            if not _valid_entity_label(ent):
                logger.warning("Entity classifier rejected label %r for key %r", ent, raw_key)
                continue
            try:
                confidence = float(conf_s)
            except ValueError:
                confidence = 0.0
            confidence = max(0.0, min(1.0, confidence))
            out[raw_key] = {
                "entity": ent.upper(),
                "confidence": round(confidence, 4),
                "reasoning": reason,
            }
        logger.info("Parsed %d entity classifier row(s)", len(out))
        return out

    def classify_entities(
        self,
        prompt_payload: Dict[str, Any],
    ) -> Dict[str, Dict[str, Any]]:
        """
        POST to entity_classifier_gateway_url. Returns row-index -> {entity, confidence, reasoning}.
        """
        if not self.entity_classifier_url:
            logger.info("Entity classifier URL not configured — skipping")
            return {}

        rendered = prompt_payload["rendered_prompt"]
        entity_context = prompt_payload.get("entity_context") or {}
        if not entity_context:
            return {}

        logger.info(
            "[entity-classifier] rows=%d indices=%s",
            len(entity_context),
            list(entity_context.keys()),
        )
        mapping_dict = self._send(
            rendered,
            url=self.entity_classifier_url,
            min_pipe_count=2,
        )
        return self._parse_entity_classifier_response(mapping_dict, entity_context)

    def _parse_parameter_bucket_response(
        self,
        mapping_dict: dict,
        entity_context: Dict[str, Dict],
    ) -> List[Dict[str, Any]]:
        # Gateway often keys rows by plain field_name while we send ENTITY||field
        # in the classifier prompt — index both forms for stable resolution.
        composite_norm_to_key: Dict[str, str] = {
            _normalize_key(k): k for k in entity_context
        }
        field_norm_to_keys: Dict[str, List[str]] = {}
        for k in entity_context:
            if "||" in k:
                tail = k.split("||", 1)[1]
                nk = _normalize_key(tail)
                field_norm_to_keys.setdefault(nk, []).append(k)
        results: List[Dict[str, Any]] = []

        logger.info(
            "Parameter classifier whitelist (composite): %s",
            list(composite_norm_to_key.keys()),
        )

        def _resolve_classifier_field_key(raw_key: str) -> Optional[str]:
            norm = _normalize_key(raw_key)
            if norm in composite_norm_to_key:
                return composite_norm_to_key[norm]
            cands = field_norm_to_keys.get(norm)
            if cands and len(cands) == 1:
                return cands[0]
            return None

        for partner_field, value_str in mapping_dict.items():
            if partner_field in ENVELOPE_KEYS:
                continue
            original_key = _resolve_classifier_field_key(partner_field)
            if entity_context:
                if original_key is None:
                    logger.warning(
                        "Parameter classifier hallucinated field '%s' — skipping",
                        partner_field,
                    )
                    continue
            else:
                original_key = partner_field
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
            if bucket == "LOANAPPLICANTPARAM":
                bucket = "LOANPARAMETER"
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
                retry_results = self._retry_fields_individually(ep, str(e))
                if retry_results is not None:
                    all_mappings.extend(retry_results)
                    continue
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

    def _retry_fields_individually(
        self,
        entity_prompt: Dict[str, Any],
        original_error: str,
    ) -> Optional[List[Dict[str, Any]]]:
        fields = entity_prompt.get("fields") or []
        if len(fields) <= 1:
            return None

        field_dictionary = entity_prompt.get("field_dictionary")
        alias_registry = entity_prompt.get("alias_registry")
        if not isinstance(field_dictionary, dict) or not isinstance(alias_registry, dict):
            return None

        entity = entity_prompt["entity"]
        logger.info(
            "Retrying LLM batch field-by-field after entity=%s failure; fields=%d",
            entity,
            len(fields),
        )

        results: List[Dict[str, Any]] = []
        for field in fields:
            partner_field = field.get("field_name") or field.get("partner_field", "")
            field_context = {
                partner_field: {
                    "entity": entity,
                    "column_category": field.get("column_category"),
                }
            }
            rendered = fill_prompt(
                template=entity_prompt.get("prompt_template", ""),
                entity=entity,
                fields=[field],
                field_dictionary=field_dictionary,
                alias_registry=alias_registry,
                client_name=entity_prompt.get("client_name", ""),
                process_name=entity_prompt.get("process_name", ""),
                pipeline_context_payload=entity_prompt.get("pipeline_context_payload"),
                mapping_policy=entity_prompt.get("mapping_policy"),
            )
            try:
                mapping_dict = self._send(rendered)
                mappings = self._parse(mapping_dict, field_context)
                if mappings:
                    results.extend(mappings)
                    continue
                raise ValueError("No mappings returned for single-field retry")
            except Exception as retry_error:
                logger.error(
                    "Single-field LLM retry failed for field=%s: %s",
                    partner_field,
                    retry_error,
                )
                results.append({
                    "partner_field":     partner_field,
                    "column_category":   field.get("column_category"),
                    "entity":            field.get("entity", entity),
                    "matched_excel_key": None,
                    "json_key":          "",
                    "confidence":        0.0,
                    "match_type":        "llm_error",
                    "reasoning":         (
                        f"LLM batch failed: {original_error}; "
                        f"single-field retry failed: {retry_error}"
                    ),
                    "needs_review":      True,
                    "winning_engine":    "none",
                })

        return results

    def classify_parameter_buckets(self, prompt_payload: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Single gateway call to classify rows into LOANPARAMETER vs LOANAPPLICANTPARAM.

        LOANAPPLICANTPARAM responses are coerced to LOANPARAMETER so the LLM never
        assigns applicant-parameter buckets via this path.
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

    def refine_loanparameter_mappings(
        self,
        entity_prompts: List[Dict[str, Any]],
    ) -> List[Dict[str, Any]]:
        """
        One gateway call per entity group to the LOANPARAMETER refinement URL
        (same auth token as main LLM). Parses pipe-delimited responses like map_fields.

        ``rendered_prompt`` is JSON (template variables only); the gateway merges with the
        Dvara/Langfuse prompt — instructions are not duplicated in this service.
        """
        if not self.loanparameter_refinement_url:
            logger.info("loanparameter_refinement_url not configured — skipping")
            return []

        all_mappings: List[Dict[str, Any]] = []

        for i, ep in enumerate(entity_prompts, 1):
            entity = ep["entity"]
            rendered = ep["rendered_prompt"]
            entity_context = ep["entity_context"]

            logger.info(
                "[LOANPARAMETER refinement %d/%d] entity=%s fields=%d keys=%s",
                i,
                len(entity_prompts),
                entity,
                len(ep.get("fields") or []),
                list(entity_context.keys()),
            )

            try:
                mapping_dict = self._send(
                    rendered,
                    url=self.loanparameter_refinement_url,
                )
                mappings = self._parse(mapping_dict, entity_context)
                all_mappings.extend(mappings)
            except Exception as e:
                logger.error(
                    "LOANPARAMETER refinement failed for entity=%s: %s",
                    entity,
                    e,
                    exc_info=True,
                )

        return all_mappings
