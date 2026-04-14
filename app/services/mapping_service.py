"""
services/mapping_service.py
Orchestrates all three layers end-to-end.

Layer 1 — deterministic (rule_engine wrapping matching_engine.py)
Layer 2 — hybrid: fuzzy → embeddings → LLM
Layer 3 — post-processing + Excel output (post_processor + generate_output)

Prompt ownership:
  The base prompt template lives on the Dvara/Langfuse platform.
  prompt_builder builds the rendered context block (client, entity,
  available_excel_keys, semantic_shortcuts, fields_to_map) and that
  block is posted as-is to the gateway via the `task` form-data field.
  No local prompt file is read or needed anywhere in this service.
"""

import json
import logging
import re
import sys
import importlib.util
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from app.services.match_context import coapplicant_custparam_base

logger = logging.getLogger(__name__)


CUSTOMER_PARAM_CATEGORY_HINTS = (
    "customer",
    "applicant",
    "coapplicant",
    "co applicant",
    "borrower",
    "personal",
    "employment",
    "income",
    "bank",
    "bureau",
    "kyc",
    "address",
    "reference",
    "demographic",
)

LOAN_LEVEL_CATEGORY_HINTS = (
    "loan",
    "disbursement",
    "repayment",
    "emi",
    "pricing",
    "sanction",
    "scheme",
    "product",
    "facility",
)

LOAN_LEVEL_FIELD_HINTS = {
    "irr", "iir", "foir", "ltv", "loantovalue", "marginmoney",
    "downpayment", "schemename", "schemeid", "schemecode",
    "subproduct", "reducedemi", "tenure", "interest", "apr",
    "loanamount", "sanctionamount", "approvedamount", "emiamount",
    "disbursementamount", "repaymentfrequency",
}


def _normalize_hint_text(value: Optional[str]) -> str:
    if not value:
        return ""
    return re.sub(r"[^a-z0-9]+", " ", value.strip().lower()).strip()


def _fallback_parameter_bucket(item: Dict[str, Any]) -> str:
    entity = (item.get("entity") or "").upper()
    category = _normalize_hint_text(item.get("column_category") or item.get("category") or "")
    field = _normalize_hint_text(item.get("partner_field") or item.get("field_name") or "")
    compact_field = field.replace(" ", "")

    coapp_base = coapplicant_custparam_base(entity)
    if coapp_base and compact_field not in LOAN_LEVEL_FIELD_HINTS:
        return coapp_base

    if entity in {"APPLICANT", "CUSTOMER"}:
        if compact_field not in LOAN_LEVEL_FIELD_HINTS:
            return "LOANAPPLICANTPARAM"

    if any(hint in category for hint in CUSTOMER_PARAM_CATEGORY_HINTS):
        if compact_field not in LOAN_LEVEL_FIELD_HINTS:
            cb = coapplicant_custparam_base(entity)
            if cb:
                return cb
            return "LOANAPPLICANTPARAM"

    if (
        entity == "LOAN"
        and any(hint in category for hint in LOAN_LEVEL_CATEGORY_HINTS)
        and any(hint in field for hint in CUSTOMER_PARAM_CATEGORY_HINTS)
    ):
        return "LOANAPPLICANTPARAM"

    return "LOANPARAMETER"


_RE_REFINE_DETERMINISTIC_BUCKET = re.compile(
    r"Deterministic\s+bucket\s+was\s+(LOANPARAMETER\d+)",
    re.I,
)
_RE_ANY_LOANPARAMETER_SLOT = re.compile(r"\b(LOANPARAMETER\d+)\b", re.I)


def _recover_prior_loanparameter_excel_key(
    row: Dict[str, Any],
    by_excel_key: Dict[str, Any],
) -> Optional[str]:
    """
    Best-effort recovery of the numbered LOANPARAMETER* slot from refinement /
    deterministic history so we can revert illegal co-applicant PUTM keys for
    APPLICANT/CUSTOMER rows.
    """
    if not by_excel_key or not isinstance(row, dict):
        return None

    def _ok(k: str) -> Optional[str]:
        kk = (k or "").strip()
        if not kk:
            return None
        if kk in by_excel_key:
            return kk
        ku = kk.upper()
        for cand in (kk, ku):
            if cand in by_excel_key:
                return cand
        return None

    stored = (row.get("deterministic_loanparameter_bucket") or "").strip()
    hit = _ok(stored)
    if hit:
        return hit

    for blob in (row.get("llm_change_reason") or "", row.get("previous_mapping_reason") or ""):
        m = _RE_REFINE_DETERMINISTIC_BUCKET.search(blob)
        if m:
            hit = _ok(m.group(1))
            if hit:
                return hit

    for m in _RE_ANY_LOANPARAMETER_SLOT.finditer(
        f"{row.get('llm_change_reason') or ''} {row.get('previous_mapping_reason') or ''}"
    ):
        hit = _ok(m.group(1))
        if hit:
            return hit

    fb = (_fallback_parameter_bucket(row) or "").strip()
    return _ok(fb)


def _is_parameter_bucket_candidate(item: Dict[str, Any]) -> bool:
    excel_key = (item.get("matched_excel_key") or "").upper()
    return (
        not excel_key
        or excel_key == "LOANPARAMETER"
        or excel_key.startswith("LOANPARAMETER")
        or excel_key == "LOANAPPLICANTPARAM"
        or excel_key.startswith("LOANAPPLICANTPARAM")
        or item.get("match_type") == "unmatched_parameter_fallback"
    )


# ── sys.path helper ────────────────────────────────────────────────────────────
def _add_scripts_to_path(scripts_dir: str):
    p = str(Path(scripts_dir).resolve())
    if p not in sys.path:
        sys.path.insert(0, p)
        logger.debug(f"Added {p} to Python path")
    parent = str(Path(p).parent)
    if parent not in sys.path:
        sys.path.insert(0, parent)
        logger.debug(f"Added {parent} to Python path")


def _import_from_scripts(scripts_dir: str, module_name: str):
    scripts_path = Path(scripts_dir).resolve()
    module_path  = scripts_path / f"{module_name}.py"
    if not module_path.exists():
        raise FileNotFoundError(f"Module file not found: {module_path}")
    _add_scripts_to_path(scripts_dir)
    spec = importlib.util.spec_from_file_location(module_name, module_path)
    if spec and spec.loader:
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    raise ImportError(f"Cannot import {module_name} from {module_path}")


# ── Stats ──────────────────────────────────────────────────────────────────────
def compute_stats(mappings: List[Dict]) -> Dict[str, Any]:
    total = len(mappings)
    if total == 0:
        return dict(total_fields=0, matched=0, unmatched=0,
                    match_rate_pct=0.0, needs_review=0, avg_confidence=0.0,
                    by_match_type={}, by_entity={}, by_confidence_band={})

    matched   = sum(1 for m in mappings if m.get("matched_excel_key"))
    needs_rev = sum(1 for m in mappings if m.get("needs_review"))
    confs     = [m.get("confidence", 0.0) for m in mappings]
    avg_conf  = round(sum(confs) / total, 4) if total > 0 else 0.0
    by_type:   Dict[str, int] = {}
    by_entity: Dict[str, int] = {}
    bands = {"0.90-1.00": 0, "0.80-0.89": 0, "0.70-0.79": 0, "0.00-0.69": 0}

    for m in mappings:
        mt = m.get("match_type", "unknown")
        by_type[mt] = by_type.get(mt, 0) + 1
        ent = m.get("entity", "OTHER")
        by_entity[ent] = by_entity.get(ent, 0) + 1
        c = m.get("confidence", 0.0)
        if c >= 0.90:   bands["0.90-1.00"] += 1
        elif c >= 0.80: bands["0.80-0.89"] += 1
        elif c >= 0.70: bands["0.70-0.79"] += 1
        else:           bands["0.00-0.69"] += 1

    return dict(
        total_fields=total, matched=matched, unmatched=total - matched,
        match_rate_pct=round(matched / total * 100, 1) if total > 0 else 0.0,
        needs_review=needs_rev, avg_confidence=avg_conf,
        by_match_type=by_type, by_entity=by_entity, by_confidence_band=bands,
    )


# ── Reference building ─────────────────────────────────────────────────────────
def build_references_from_db_direct(
    putm_xlsx: str,
    mapping_csv: str,
    references_dir: str,
    scripts_dir: str,
) -> Dict[str, int]:
    import pandas as pd

    scripts_path = Path(scripts_dir)
    search_paths = [
        scripts_path / "build_references.py",
        scripts_path.parent / "scripts" / "build_references.py",
        Path.cwd() / "scripts" / "build_references.py",
        Path.cwd() / "build_references.py",
    ]
    build_script_path = next((p for p in search_paths if p.exists()), None)
    if not build_script_path:
        raise FileNotFoundError(
            "build_references.py not found. Searched:\n" +
            "\n".join(f"  - {p}" for p in search_paths)
        )

    logger.info(f"Found build_references.py at {build_script_path}")
    build_module = _import_from_scripts(str(build_script_path.parent), "build_references")

    FieldDictionaryBuilder = getattr(build_module, "FieldDictionaryBuilder")
    AliasRegistryBuilder   = getattr(build_module, "AliasRegistryBuilder")
    EntityRoutingBuilder   = getattr(build_module, "EntityRoutingBuilder")

    logger.info(f"Reading PUTM dump from {putm_xlsx}")
    putm_df = pd.read_excel(putm_xlsx)
    logger.info(f"  → {len(putm_df)} PUTM rows")

    logger.info(f"Reading generic-mapping dump from {mapping_csv}")
    generic_df = pd.read_csv(mapping_csv, encoding="utf-8")
    logger.info(f"  → {len(generic_df)} generic-mapping rows")

    ref_dir = Path(references_dir)
    ref_dir.mkdir(parents=True, exist_ok=True)

    logger.info("Building field_dictionary.json …")
    fd = FieldDictionaryBuilder(putm_df).build()
    out = ref_dir / "field_dictionary.json"
    out.write_text(json.dumps(fd, indent=2, ensure_ascii=False), encoding="utf-8")
    logger.info(f"✓ Wrote {out}  ({fd['metadata']['total']} entries)")

    logger.info("Building alias_registry.json …")
    ar = AliasRegistryBuilder(generic_df).build()
    out = ref_dir / "alias_registry.json"
    out.write_text(json.dumps(ar, indent=2, ensure_ascii=False), encoding="utf-8")
    logger.info(f"✓ Wrote {out}  ({len(ar['forward'])} forward mappings)")

    logger.info("Building entity_routing.json …")
    er = EntityRoutingBuilder(generic_df).build()
    out = ref_dir / "entity_routing.json"
    out.write_text(json.dumps(er, indent=2, ensure_ascii=False), encoding="utf-8")
    logger.info(f"✓ Wrote {out}  ({er['metadata']['total_groupings']} groupings)")

    return {
        "field_count":           len(fd.get("all_excel_keys", [])),
        "alias_count":           len(ar.get("forward", {})),
        "entity_routing_exists": True,
        "references_dir":        references_dir,
        "field_dict_total":      fd["metadata"]["total"],
        "total_partners":        ar["metadata"]["total_partners"],
    }


# ── Optional LLM entity labels (pre–match_batch) ───────────────────────────────
def _apply_llm_entity_classifier(
    fields: List[Dict[str, Any]],
    settings,
    client_name: str = "",
    process_name: str = "",
) -> List[Dict[str, Any]]:
    """
    When `use_llm_entity_classifier` is enabled and ENTITY_CLASSIFIER_GATEWAY_URL is set,
    the entity-assignment LLM (`entity_assignment` gateway) runs **before** `match_batch`.
    Its `entity` per row (APPLICANT, LOAN, FEE, DOCUMENT, etc.) is what deterministic
    matching and downstream prompts use unless the call fails—in that case rows keep
    prior entity and `match_batch` falls back to `detect_entity` heuristics.

    Later pipeline steps may still adjust entity from resolved json paths or business
    rules (e.g. FEE entity override after mapping).
    """
    url = (getattr(settings, "entity_classifier_gateway_url", None) or "").strip()
    if not url:
        logger.info("ENTITY_CLASSIFIER_GATEWAY_URL not set — skipping LLM entity classification")
        return fields

    from app.services.llm_service import LLMService
    from app.services.prompt_builder import build_entity_classifier_prompt

    payload = build_entity_classifier_prompt(
        fields,
        client_name=client_name or "",
        process_name=process_name or "",
    )
    if not payload.get("entity_context"):
        return fields

    try:
        llm = LLMService(settings)
        classified = llm.classify_entities(payload)
    except Exception as e:
        logger.warning(
            "LLM entity classifier failed (%s) — falling back to heuristic entity detection",
            e,
            exc_info=True,
        )
        return fields

    applied = 0
    for key, meta in classified.items():
        ctx = payload["entity_context"].get(key)
        if not ctx:
            continue
        idx = ctx.get("list_index")
        if idx is None or not isinstance(idx, int) or idx < 0 or idx >= len(fields):
            continue
        ent = (meta.get("entity") or "").strip().upper()
        if not ent:
            continue
        row = fields[idx]
        row["entity"] = ent
        row["entity_classifier_confidence"] = meta.get("confidence")
        row["entity_classifier_reasoning"] = (meta.get("reasoning") or "").strip()
        applied += 1

    logger.info("LLM entity classifier applied entity to %d/%d field row(s)", applied, len(fields))
    return fields


# ── Layer 1 — deterministic ────────────────────────────────────────────────────
def run_deterministic(
    input_file: str,
    settings,
    process_name: str = "COMBINED",
    sheet_filter: Optional[str] = None,
    client_name: str = "",
    use_llm_entity_classifier: bool = False,
) -> Dict[str, Any]:
    """
    Layer 1: deterministic matching.

    If `use_llm_entity_classifier` is True, entity labels come from the LLM entity
    classifier first; otherwise entities come from sheet/category heuristics inside
    `match_batch` / `detect_entity`.

    Also prepares entity_prompts for Layer 2 — each prompt's
    rendered_prompt is the context block (no base template) that
    will be posted to the gateway as the `task` field value.
    """
    _add_scripts_to_path(settings.scripts_dir)

    try:
        from input_parser    import parse_input
        from matching_engine import load_references, match_batch
        from app.services.prompt_builder import build_entity_prompts
    except ImportError as e:
        logger.error(f"Failed to import required modules: {e}")
        logger.error(f"Python path: {sys.path}")
        raise

    refs = load_references(settings.references_dir)
    fd   = refs["field_dictionary"]
    ar   = refs["alias_registry"]

    fields = parse_input(input_file)
    if sheet_filter:
        fields = [f for f in fields if f.get("source_sheet") == sheet_filter]

    if use_llm_entity_classifier:
        fields = _apply_llm_entity_classifier(
            fields,
            settings,
            client_name=client_name,
            process_name=process_name,
        )

    batch     = match_batch(fields, refs, process_name=process_name)
    matched   = batch["matched"]
    unmatched = batch["unmatched"]

    det_dicts = []

    for m in matched:
        det_dicts.append({
            "partner_field":     m.partner_field,
            "column_category":   m.column_category,
            "entity":            m.entity,
            "process_name":      process_name or "",
            "matched_excel_key": m.matched_excel_key,
            "json_key":          m.matched_json_key or "",
            "confidence":        m.confidence,
            "match_type":        m.match_type,
            "reasoning":         m.reasoning,
            "previous_mapping_reason": (
                f"Deterministic first-match: match_type={m.match_type} "
                f"confidence={round(float(m.confidence or 0.0), 4)}. "
                f"Reason: {(m.reasoning or '').strip()}"
            ).strip(),
            "llm_change_reason": "",
            "llm_param_bucket_reason": "",
            "needs_review":      m.confidence < settings.review_threshold,
            "winning_engine":    "deterministic",
        })

    unmatched_dicts = [
        {
            "field_name":      u.partner_field,
            "partner_field":   u.partner_field,
            "column_category": u.column_category,
            "entity":          u.entity,
            "process_name":    process_name or "",
        }
        for u in unmatched
    ]

    # NOTE: Deterministic LOANPARAMETER assignments are sent to the LLM
    # as context (loanparameter_assigned_fields), not mixed into unmatched_fields.

    # Build entity_prompts — rendered_prompt = context block only.
    # The gateway merges it with the stored base template on its side.
    entity_prompts = (
        build_entity_prompts(
            unmatched_fields=unmatched_dicts,
            field_dictionary=fd,
            alias_registry=ar,
            prompt_template="",   # gateway owns the base template
            client_name="",
            process_name=process_name,
        )
        if unmatched_dicts
        else []
    )

    return {
        "deterministic_results": det_dicts,
        "unmatched_fields":      unmatched_dicts,
        "entity_prompts":        entity_prompts,
        "field_dictionary":      fd,
        "alias_registry":        ar,
    }


def merge_deterministic_with_hybrid_phase(
    det_results: List[Dict[str, Any]],
    phase2: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """
    Append hybrid-phase rows to deterministic rows without duplicating fields.

    If the same (entity, partner_field) appears in phase2 — for example after
    LLM verification of a deterministic LOANPARAMETER* or alias_tier4 mapping —
    the phase2 row is kept and the corresponding deterministic row is omitted.
    """

    def _row_key(row: Dict[str, Any]) -> str:
        ent = (row.get("entity") or "OTHER").strip().upper()
        pf = (row.get("partner_field") or row.get("field_name") or "").strip()
        if not pf:
            return ""
        return f"{ent}||{pf}"

    phase_keys: Set[str] = set()
    for row in phase2 or []:
        if isinstance(row, dict):
            k = _row_key(row)
            if k:
                phase_keys.add(k)

    merged: List[Dict[str, Any]] = []
    for row in det_results or []:
        if not isinstance(row, dict):
            continue
        k = _row_key(row)
        if k and k in phase_keys:
            continue
        merged.append(row)
    merged.extend(phase2 or [])
    return merged


def refine_loanparameter_after_deterministic(
    deterministic_results: List[Dict[str, Any]],
    field_dictionary: Dict[str, Any],
    alias_registry: Dict[str, Any],
    settings,
    client_name: str = "",
    process_name: str = "",
) -> List[Dict[str, Any]]:
    """
    After Layer 1: remap rows stuck on LOANPARAMETER* using a dedicated LLM + PUTM catalog.
    Requires settings.loanparameter_refinement_gateway_url (same token as main LLM).
    """
    refine_url = getattr(settings, "loanparameter_refinement_gateway_url", None) or ""
    if deterministic_results is None:
        logger.warning("refine_loanparameter_after_deterministic received None — returning []")
        return []
    # ── DEBUG LOGS ─────────────────────────────────────────────────────────────
    logger.info("REFINE_DEBUG: loanparameter_refinement_gateway_url = %r", refine_url)
    logger.info("REFINE_DEBUG: type = %s", type(refine_url).__name__)
    logger.info("REFINE_DEBUG: stripped = %r", str(refine_url).strip())
    logger.info("REFINE_DEBUG: bool check = %s", not str(refine_url).strip())
    logger.info("REFINE_DEBUG: total det_results = %d", len(deterministic_results))
    lp_candidate_count = sum(
        1 for row in deterministic_results
        if (row.get("matched_excel_key") or "").upper().startswith("LOANPARAMETER")
    )
    logger.info("REFINE_DEBUG: LOANPARAMETER* rows in det_results = %d", lp_candidate_count)
    # ──────────────────────────────────────────────────────────────────────────

    if not str(refine_url).strip():
        logger.info("LOANPARAMETER refinement URL not set — skipping PUTM refinement step")
        return deterministic_results

    lp_rows = [
        row
        for row in deterministic_results
        if (row.get("matched_excel_key") or "").upper().startswith("LOANPARAMETER")
    ]

    logger.info("REFINE_DEBUG: lp_rows selected for refinement = %d", len(lp_rows))
    if not lp_rows:
        logger.info("REFINE_DEBUG: early exit — no LOANPARAMETER rows to refine")
        return deterministic_results

    from app.services.prompt_builder import build_loanparameter_refinement_prompts
    from app.services.llm_service import LLMService

    try:
        prompts = build_loanparameter_refinement_prompts(
            loanparameter_rows=lp_rows,
            field_dictionary=field_dictionary,
            alias_registry=alias_registry,
            client_name=client_name,
            process_name=process_name,
        )
        logger.info("REFINE_DEBUG: prompts built = %d", len(prompts))
    except Exception as e:
        logger.error("build_loanparameter_refinement_prompts failed: %s", e, exc_info=True)
        return deterministic_results

    if not prompts:
        logger.info("REFINE_DEBUG: no prompts returned — skipping")
        return deterministic_results

    llm = LLMService(settings)
    logger.info("REFINE_DEBUG: LLMService.loanparameter_refinement_url = %r", llm.loanparameter_refinement_url)

    try:
        refined_list = llm.refine_loanparameter_mappings(prompts)
        if refined_list is None:
            logger.warning("refine_loanparameter_mappings returned None — treating as empty")
            refined_list = []
        logger.info("REFINE_DEBUG: refined_list returned = %d rows", len(refined_list))
    except Exception as e:
        logger.error("LOANPARAMETER refinement gateway failed: %s", e, exc_info=True)
        return deterministic_results or []

    # ── Apply refined mappings back onto deterministic_results ─────────────────
    _add_scripts_to_path(settings.scripts_dir)
    from matching_engine import (
        _build_by_excel_key_index,
        _is_coapplicant_scoped_putm_excel_key,
        _remap_putm_key_for_applicant_customer_entity,
    )

    by_field = {m.get("partner_field"): m for m in refined_list if m.get("partner_field")}
    by_excel = field_dictionary.get("by_excel_key") or {}
    by_excel_fd_idx = _build_by_excel_key_index(field_dictionary)
    review_threshold = getattr(settings, "review_threshold", 0.80)

    out: List[Dict[str, Any]] = []
    for row in deterministic_results:
        pf = row.get("partner_field")
        is_lp = (row.get("matched_excel_key") or "").upper().startswith("LOANPARAMETER")
        if not is_lp or not pf or pf not in by_field:
            out.append(row)
            continue

        new_m = by_field[pf]
        new_key = (new_m.get("matched_excel_key") or "").strip()
        if not new_key:
            out.append(row)
            continue

        ent_row = (row.get("entity") or "OTHER").strip().upper()
        if ent_row in ("APPLICANT", "CUSTOMER") and _is_coapplicant_scoped_putm_excel_key(new_key):
            alt = _remap_putm_key_for_applicant_customer_entity(ent_row, new_key, by_excel_fd_idx)
            if alt:
                new_key, _alt_jk = alt
                new_m = {**new_m, "matched_excel_key": new_key}
            else:
                old_lp = row.get("matched_excel_key") or ""
                logger.warning(
                    "PUTM refinement rejected co-applicant-only key %r for entity=%s field=%r — keeping %s",
                    (new_m.get("matched_excel_key") or "").strip(),
                    ent_row,
                    pf,
                    old_lp,
                )
                out.append({
                    **row,
                    "deterministic_loanparameter_bucket": old_lp,
                    "needs_review": True,
                    "llm_change_reason": (
                        f"PUTM refinement proposed {(new_m.get('matched_excel_key') or '').strip()} "
                        f"which is co-applicant-scoped while entity={ent_row}; kept bucket {old_lp}."
                    ).strip(),
                })
                continue

        if new_key not in by_excel:
            logger.warning(
                "LOANPARAMETER refinement: unknown excel_key %r for field %r — keeping bucket",
                new_key,
                pf,
            )
            out.append({
                **row,
                "needs_review": True,
                "llm_change_reason": (
                    (row.get("llm_change_reason") or "").strip()
                    + f" PUTM refinement proposed invalid key {new_key}."
                ).strip(),
            })
            continue

        info = by_excel[new_key]
        old_key = row.get("matched_excel_key")
        conf = float(new_m.get("confidence") or 0.0)
        new_reason = (new_m.get("reasoning") or "").strip()

        out.append({
            **row,
            "deterministic_loanparameter_bucket": old_key,
            "matched_excel_key": new_key,
            "json_key":          info.get("json_key") or new_m.get("json_key") or "",
            "confidence":        conf,
            "match_type":        "llm_putm_refinement",
            "winning_engine":    "llm_putm_refinement",
            "reasoning":         new_reason or row.get("reasoning", ""),
            "llm_change_reason": (
                f"Deterministic bucket was {old_key}; PUTM refinement → {new_key}."
                + (f" {new_reason}" if new_reason else "")
            ).strip(),
            "needs_review":      conf < review_threshold,
        })

    logger.info(
        "LOANPARAMETER PUTM refinement applied where possible: refined_rows=%d",
        len(by_field),
    )
    return out


# ── Layer 2 — hybrid + LLM ────────────────────────────────────────────────────
def run_hybrid_llm(
    unmatched_fields: List[Dict],
    field_dictionary: Dict,
    alias_registry: Dict,
    entity_prompts: List[Dict],
    settings,
    deterministic_matches: Optional[List[Dict[str, Any]]] = None,
    use_fuzzy: bool      = True,
    use_embeddings: bool = False,
    use_llm: bool        = True,
    client_name: str     = "",
    process_name: str    = "",
) -> Tuple[List[Dict], Dict[str, int]]:
    """
    Layer 2: fuzzy → embeddings → LLM.

    For LLM: re-builds entity_prompts scoped to whatever fields are still
    remaining after fuzzy/embedding, then posts each rendered_prompt
    directly to the gateway as the `task` field value.
    """
    _add_scripts_to_path(settings.scripts_dir)

    from app.services.fuzzy_engine     import FuzzyEngine,     HAS_RAPIDFUZZ
    from app.services.embedding_engine import EmbeddingEngine, HAS_ST
    from app.services.llm_service      import LLMService
    from app.services.prompt_builder   import build_entity_prompts
    from app.services.match_context    import semantic_field_guard_reason

    remaining = list(unmatched_fields)
    results:   List[Dict] = []
    breakdown = {"fuzzy": 0, "embedding": 0, "llm": 0, "unmatched": 0, "invalid_field_name": 0}
    review_threshold = getattr(settings, "review_threshold", 0.80)
    review_candidates_by_field: Dict[str, Dict] = {}
    fuzzy_seen: List[Dict[str, Any]] = []
    embedding_seen: List[Dict[str, Any]] = []

    deterministic_matches = deterministic_matches or []

    def _field_id(item: Dict[str, Any]) -> str:
        return item.get("partner_field") or item.get("field_name") or ""

    def _make_llm_context_item(
        item: Dict[str, Any],
        *,
        process_name_value: str,
        matched_target_override: Optional[str] = None,
    ) -> Dict[str, Any]:
        field_name = item.get("field_name") or item.get("partner_field") or ""
        return {
            "field_name":      field_name,
            "column_category": item.get("column_category") or "",
            "process_name":    (item.get("process_name") or process_name_value or ""),
            "entity":          (item.get("entity") or "OTHER").upper(),
            "matched_target": (
                matched_target_override
                if matched_target_override is not None
                else (item.get("matched_excel_key") or "")
            ),
        }

    def _candidate_payload(item: Dict[str, Any]) -> Dict[str, Any]:
        return {
            "engine": item.get("winning_engine") or item.get("match_type"),
            "matched_excel_key": item.get("matched_excel_key"),
            "json_key": item.get("json_key", ""),
            "confidence": item.get("confidence", 0.0),
            "reasoning": item.get("reasoning", ""),
        }

    def _ensure_trace_columns(item: Dict[str, Any]) -> None:
        item.setdefault("previous_mapping_reason", "")
        item.setdefault("llm_change_reason", "")
        item.setdefault("llm_param_bucket_reason", "")

    def _format_previous_mapping_reason(item: Dict[str, Any]) -> str:
        candidates = item.get("candidate_matches") or []
        if isinstance(candidates, list) and candidates:
            usable = [
                c for c in candidates
                if isinstance(c, dict) and (c.get("matched_excel_key") or "")
            ]
            usable.sort(key=lambda c: float(c.get("confidence") or 0.0), reverse=True)
            best = usable[0] if usable else None
            parts: List[str] = []
            if best:
                parts.append(
                    "Pre-LLM best candidate: "
                    f"engine={best.get('engine')} target={best.get('matched_excel_key')} "
                    f"confidence={round(float(best.get('confidence') or 0.0), 4)}. "
                    f"Reason: {(best.get('reasoning') or '').strip()}"
                )
            others = usable[1:4]
            if others:
                compact = [
                    f"{c.get('engine')}->{c.get('matched_excel_key')}({round(float(c.get('confidence') or 0.0), 4)})"
                    for c in others
                ]
                parts.append("Other candidates: " + ", ".join(compact))
            return " ".join(p for p in parts if p).strip()

        engine = item.get("winning_engine") or item.get("match_type") or "unknown"
        target = item.get("matched_excel_key") or ""
        conf = round(float(item.get("confidence") or 0.0), 4)
        reason = (item.get("reasoning") or "").strip()
        if not target and not reason:
            return ""
        return (
            f"Pre-LLM mapping: engine={engine} target={target or '(unmatched)'} "
            f"confidence={conf}. Reason: {reason}"
        ).strip()

    def _queue_review_candidate(item: Dict[str, Any]) -> None:
        field_id = _field_id(item)
        if not field_id:
            return

        entry = review_candidates_by_field.get(field_id)
        if not entry:
            entry = {
                **item,
                "field_name": item.get("field_name") or item.get("partner_field"),
                "partner_field": item.get("partner_field") or item.get("field_name"),
                "candidate_matches": [],
            }
            review_candidates_by_field[field_id] = entry

        candidate = _candidate_payload(item)
        existing = entry["candidate_matches"]
        for idx, current in enumerate(existing):
            if current.get("engine") == candidate["engine"]:
                if candidate["confidence"] > current.get("confidence", 0.0):
                    existing[idx] = candidate
                break
        else:
            existing.append(candidate)

        if item.get("confidence", 0.0) >= entry.get("confidence", 0.0):
            old_reasoning = entry.get("reasoning", "")
            new_reasoning = item.get("reasoning", "")

            # Preserve reasoning chain if confidence improves
            combined_reasoning = new_reasoning
            if old_reasoning and old_reasoning != new_reasoning:
                combined_reasoning = f"{old_reasoning} → {new_reasoning}"

            entry.update({
                "matched_excel_key": item.get("matched_excel_key"),
                "json_key": item.get("json_key", ""),
                "confidence": item.get("confidence", 0.0),
                "match_type": item.get("match_type", entry.get("match_type")),
                "reasoning": combined_reasoning,
                "needs_review": True,
                "winning_engine": item.get("winning_engine", entry.get("winning_engine")),
            })
            if item.get("fuzzy_score") is not None:
                entry["fuzzy_score"] = item.get("fuzzy_score")
            if item.get("embedding_score") is not None:
                entry["embedding_score"] = item.get("embedding_score")

    def _clear_review_candidate(item: Dict[str, Any]) -> None:
        field_id = _field_id(item)
        if field_id:
            review_candidates_by_field.pop(field_id, None)

    def _dedupe_fields(items):
        deduped = {}
        for item in items:
            entity = (item.get("entity") or "OTHER").upper()
            field_id = f"{entity}||{_field_id(item)}"
            deduped[field_id] = item
        return list(deduped.values())

    def _split_by_confidence(items: List[Dict]) -> Tuple[List[Dict], List[Dict]]:
        accepted: List[Dict] = []
        review: List[Dict] = []
        for item in items:
            if item.get("confidence", 0.0) < review_threshold:
                _queue_review_candidate(item)
                review.append(item)
            else:
                _clear_review_candidate(item)
                accepted.append(item)
        return accepted, review

    # 2a0 — Reject unusable partner field names before fuzzy/embedding/LLM.
    # Short or generic tokens (e.g. "as") still build rich semantic queries from
    # category + entity and produce spurious high-confidence matches.
    gated: List[Dict[str, Any]] = []
    kept_remaining: List[Dict[str, Any]] = []
    for item in remaining:
        fid = _field_id(item)
        guard = semantic_field_guard_reason(fid)
        if guard:
            _ensure_trace_columns(item)
            gated.append({
                **item,
                "matched_excel_key": None,
                "json_key": "",
                "confidence": 0.0,
                "match_type": "invalid_field_name",
                "reasoning": guard,
                "needs_review": True,
                "winning_engine": None,
                "previous_mapping_reason": "",
                "llm_change_reason": "",
                "llm_param_bucket_reason": "",
            })
        else:
            kept_remaining.append(item)
    remaining = kept_remaining
    if gated:
        logger.info(
            "Semantic field-name guard rejected %d field(s); skipping fuzzy, embedding, LLM.",
            len(gated),
        )
    results.extend(gated)
    breakdown["invalid_field_name"] = len(gated)

    # 2a — Fuzzy
    if use_fuzzy and remaining and HAS_RAPIDFUZZ:
        try:
            fe = FuzzyEngine(
                field_dictionary,
                threshold=settings.fuzzy_threshold,
                process_name=process_name,
            )
            fuzzy_matched, remaining = fe.run_batch(remaining, process_name=process_name)
            fuzzy_seen.extend(fuzzy_matched)
            accepted, review = _split_by_confidence(fuzzy_matched)
            for item in accepted:
                _ensure_trace_columns(item)
                item["previous_mapping_reason"] = (
                    item.get("previous_mapping_reason") or _format_previous_mapping_reason(item)
                )
            if review:
                logger.info(
                    "FuzzyEngine routed %d low-confidence fields to LLM review "
                    "(threshold=%.2f, process=%s)",
                    len(review),
                    review_threshold,
                    process_name or "COMBINED",
                )
            results.extend(accepted)
            breakdown["fuzzy"] = len(accepted)
            remaining = _dedupe_fields(remaining + review)
        except Exception as e:
            logger.warning(f"FuzzyEngine skipped: {e}")

    # 2b — Embeddings
    if use_embeddings and remaining and HAS_ST:
        try:
            ee = EmbeddingEngine(
                field_dictionary,
                threshold=settings.embedding_threshold,
                process_name=process_name,
            )
            emb_matched, remaining = ee.run_batch(remaining, process_name=process_name)
            embedding_seen.extend(emb_matched)
            accepted, review = _split_by_confidence(emb_matched)
            for item in accepted:
                _ensure_trace_columns(item)
                item["previous_mapping_reason"] = (
                    item.get("previous_mapping_reason") or _format_previous_mapping_reason(item)
                )
            if review:
                logger.info(
                    "EmbeddingEngine routed %d low-confidence fields to LLM review "
                    "(threshold=%.2f, process=%s)",
                    len(review),
                    review_threshold,
                    process_name or "COMBINED",
                )
            results.extend(accepted)
            breakdown["embedding"] = len(accepted)
            remaining = _dedupe_fields(remaining + review)
        except Exception as e:
            logger.warning(f"EmbeddingEngine skipped: {e}")

    # 2c — LLM — only LOANPARAMETER* buckets; do not re-send alias_tier4 rows that
    # already map to concrete catalogue keys (avoids gateway/classifier overwriting
    # keys like APPLICATIONFILEID, LOANID, etc.).
    deterministic_llm_recheck_rows = [
        item
        for item in (deterministic_matches or [])
        if isinstance(item, dict)
        and _field_id(item)
        and (item.get("matched_excel_key") or "").upper().startswith("LOANPARAMETER")
    ]
    if deterministic_llm_recheck_rows:
        lp_ids = [
            (_field_id(i) or "").strip()
            for i in deterministic_llm_recheck_rows
            if (i.get("matched_excel_key") or "").upper().startswith("LOANPARAMETER")
        ]
        t4_ids = [
            (_field_id(i) or "").strip()
            for i in deterministic_llm_recheck_rows
            if (i.get("match_type") or "").strip() == "alias_tier4"
        ]
        logger.info(
            "Including deterministic LOANPARAMETER* rows for LLM verification: count=%d "
            "fields=%s (alias_tier4 among them=%d: %s)",
            len(lp_ids),
            lp_ids,
            len(t4_ids),
            t4_ids,
        )

    llm_input_fields = _dedupe_fields(
        list(review_candidates_by_field.values())
        + [
            item
            for item in remaining
            if _field_id(item) not in review_candidates_by_field
        ]
        + deterministic_llm_recheck_rows
    )
    if use_llm and llm_input_fields:
        try:
            low_conf_fields_for_llm = list(review_candidates_by_field.values())
            low_conf_candidate_count = sum(
                len(item.get("candidate_matches") or [])
                for item in low_conf_fields_for_llm
            )
            logger.info(
                "Preparing LLM input: total_fields=%d, low_confidence_fields=%d, "
                "deterministic_llm_recheck_rows=%d, low_confidence_candidates=%d, threshold=%.2f",
                len(llm_input_fields),
                len(low_conf_fields_for_llm),
                len(deterministic_llm_recheck_rows),
                low_conf_candidate_count,
                review_threshold,
            )

            loanparameter_assigned_fields = [
                _make_llm_context_item(
                    item,
                    process_name_value=process_name,
                    matched_target_override=(item.get("matched_excel_key") or ""),
                )
                for item in deterministic_matches
                if (item.get("matched_excel_key") or "").upper().startswith("LOANPARAMETER")
            ]
            llm_context_payload = {
                "unmatched_fields": [
                    _make_llm_context_item(item, process_name_value=process_name)
                    for item in llm_input_fields
                ],
                "deterministic_matches": [
                    _make_llm_context_item(
                        item,
                        process_name_value=process_name,
                        matched_target_override=(item.get("matched_excel_key") or ""),
                    )
                    for item in deterministic_matches
                ],
                "fuzzy_matches": [
                    _make_llm_context_item(item, process_name_value=process_name)
                    for item in fuzzy_seen
                ],
                "embedding_matches": [
                    _make_llm_context_item(item, process_name_value=process_name)
                    for item in embedding_seen
                ],
                "loanparameter_assigned_fields": loanparameter_assigned_fields,
            }

            pre_llm_by_field: Dict[str, Dict[str, Any]] = {
                _field_id(item): item
                for item in llm_input_fields
                if isinstance(item, dict) and _field_id(item)
            }
            for item in llm_input_fields:
                if isinstance(item, dict):
                    item.setdefault("process_name", process_name or "")
                    item["entity"] = (item.get("entity") or "OTHER").upper()
                    _ensure_trace_columns(item)
            fresh_prompts = build_entity_prompts(
                unmatched_fields=llm_input_fields,
                field_dictionary=field_dictionary,
                alias_registry=alias_registry,
                prompt_template="",   # gateway owns the base template
                client_name=client_name,
                process_name=process_name,
                pipeline_context_payload=llm_context_payload,
            )
            llm_svc      = LLMService(settings)
            llm_mappings = llm_svc.map_fields(fresh_prompts)

            def _expected_field_ids_from_prompts(
                prompts: List[Dict[str, Any]],
            ) -> Tuple[Set[str], Dict[str, Dict[str, Any]]]:
                ids: Set[str] = set()
                meta: Dict[str, Dict[str, Any]] = {}
                for ep in prompts:
                    for fname, ctx in (ep.get("entity_context") or {}).items():
                        if not fname:
                            continue
                        ids.add(fname)
                        meta[fname] = ctx if isinstance(ctx, dict) else {}
                return ids, meta

            expected_ids, expected_meta = _expected_field_ids_from_prompts(fresh_prompts)
            returned_ids = {
                (m.get("partner_field") or m.get("field_name") or "").strip()
                for m in llm_mappings
                if isinstance(m, dict) and (m.get("partner_field") or m.get("field_name"))
            }
            omitted_ids = expected_ids - returned_ids
            if omitted_ids:
                logger.warning(
                    "LLM batch omitted %d/%d field keys; re-attaching pre-LLM state: %s",
                    len(omitted_ids),
                    len(expected_ids),
                    sorted(omitted_ids)[:50],
                )
                for fid in omitted_ids:
                    prev = pre_llm_by_field.get(fid)
                    if not isinstance(prev, dict):
                        prev = {}
                    src = next(
                        (x for x in llm_input_fields if _field_id(x) == fid),
                        {},
                    )
                    if not isinstance(src, dict):
                        src = {}
                    exp_ctx = expected_meta.get(fid) or {}
                    mt = (prev.get("match_type") or prev.get("winning_engine") or "none")
                    llm_mappings.append({
                        "partner_field":     fid,
                        "column_category":   (
                            exp_ctx.get("column_category")
                            or prev.get("column_category")
                            or src.get("column_category")
                        ),
                        "entity":            (
                            (exp_ctx.get("entity") or prev.get("entity") or src.get("entity") or "OTHER")
                            .upper()
                        ),
                        "process_name":      (
                            exp_ctx.get("process_name")
                            or prev.get("process_name")
                            or src.get("process_name")
                            or process_name
                            or ""
                        ),
                        "matched_excel_key": prev.get("matched_excel_key"),
                        "json_key":          prev.get("json_key", "") or "",
                        "confidence":        float(prev.get("confidence") or 0.0),
                        "match_type":        "llm_batch_omitted",
                        "reasoning":         (
                            "Gateway/LLM output did not include this field key; "
                            "retained the best pre-LLM mapping when present, "
                            "otherwise left unmatched for review."
                        ),
                        "needs_review":      True,
                        "winning_engine":    prev.get("winning_engine") or mt,
                        "previous_mapping_reason": (
                            _format_previous_mapping_reason(prev) if prev else ""
                        ),
                        "llm_change_reason":   "",
                        "llm_param_bucket_reason": "",
                    })

            for mapping in llm_mappings:
                if not isinstance(mapping, dict):
                    continue
                field_id = (mapping.get("partner_field") or mapping.get("field_name") or "").strip()
                prev = pre_llm_by_field.get(field_id, {})

                _ensure_trace_columns(mapping)
                mapping.setdefault("process_name", process_name or "")
                mapping["entity"] = (
                    mapping.get("entity")
                    or prev.get("entity")
                    or "OTHER"
                )
                mapping["entity"] = (mapping["entity"] or "OTHER").upper()

                prev_reason = _format_previous_mapping_reason(prev) if isinstance(prev, dict) else ""
                if prev_reason:
                    mapping["previous_mapping_reason"] = prev_reason

                prev_target = (prev.get("matched_excel_key") or "").strip() if isinstance(prev, dict) else ""
                prev_engine = (prev.get("winning_engine") or prev.get("match_type") or "").strip() if isinstance(prev, dict) else ""
                prev_conf = round(float(prev.get("confidence") or 0.0), 4) if isinstance(prev, dict) else 0.0
                new_target = (mapping.get("matched_excel_key") or "").strip()
                new_conf = round(float(mapping.get("confidence") or 0.0), 4)
                if prev_target and new_target and prev_target != new_target:
                    mapping["llm_change_reason"] = (
                        f"Previous mapping was {prev_target} via {prev_engine or 'prior_engine'} "
                        f"(confidence={prev_conf}). LLM changed to {new_target} "
                        f"(confidence={new_conf}). LLM reasoning: {(mapping.get('reasoning') or '').strip()}"
                    ).strip()
            resolved_fields = {
                m.get("partner_field") or m.get("field_name")
                for m in llm_mappings
                if m.get("partner_field") or m.get("field_name")
            }
            results.extend(llm_mappings)
            breakdown["llm"] = len(llm_mappings)
            remaining = [
                f for f in remaining
                if (f.get("partner_field") or f.get("field_name")) not in resolved_fields
            ]
            review_candidates_by_field = {
                field_id: item
                for field_id, item in review_candidates_by_field.items()
                if field_id not in resolved_fields
            }
        except Exception as e:
            logger.error(f"LLMService error: {e}", exc_info=True)

    if review_candidates_by_field:
        for item in review_candidates_by_field.values():
            if isinstance(item, dict):
                _ensure_trace_columns(item)
                item["previous_mapping_reason"] = (
                    item.get("previous_mapping_reason") or _format_previous_mapping_reason(item)
                )
        results.extend(review_candidates_by_field.values())

    # Still unmatched after all engines
    unmatched_count = 0
    fallback_loan_parameter = 0
    fallback_loan_applicant_parameter = 0
    for f in remaining:
        field_id = _field_id(f)
        if field_id and field_id in review_candidates_by_field:
            continue
        fallback_excel_key = _fallback_parameter_bucket(f)
        if fallback_excel_key == "LOANAPPLICANTPARAM":
            fallback_loan_applicant_parameter += 1
        else:
            fallback_loan_parameter += 1
        results.append({
            **f,
            "matched_excel_key": fallback_excel_key,
            "json_key":          "",
            "confidence":        0.0,
            "match_type":        "unmatched_parameter_fallback",
            "reasoning":         (
                "No match found in deterministic, fuzzy, embedding, or LLM engines; "
                f"defaulted to {fallback_excel_key} using entity/category/field context"
            ),
            "previous_mapping_reason": (
                "No earlier engine produced an accepted match; "
                "field remained unresolved after fuzzy/embedding/LLM review."
            ),
            "llm_change_reason": "",
            "llm_param_bucket_reason": "",
            "needs_review":      True,
            "winning_engine":    "none",
        })
        unmatched_count += 1
    breakdown["unmatched"] = unmatched_count
    if unmatched_count:
        logger.info(
            "Fallback-labeled unmatched fields: total=%d, LOANPARAMETER=%d, "
            "LOANAPPLICANTPARAM=%d",
            unmatched_count,
            fallback_loan_parameter,
            fallback_loan_applicant_parameter,
        )

    return results, breakdown


def refine_parameter_buckets(
    all_mappings: List[Dict[str, Any]],
    settings,
    client_name: str = "",
    process_name: str = "",
) -> List[Dict[str, Any]]:
    """
    Extra classifier pass:
    send parameter-like rows to a second LLM gateway that decides only
    LOANPARAMETER vs LOANAPPLICANTPARAM.
    """
    if not getattr(settings, "parameter_classifier_gateway_url", ""):
        logger.info("Parameter classifier gateway not configured — skipping refinement")
        return all_mappings

    from app.services.llm_service import LLMService
    from app.services.prompt_builder import build_parameter_classifier_prompt

    def _classifier_id(item: Dict[str, Any]) -> str:
        partner_field = (item.get("partner_field") or item.get("field_name") or "").strip()
        entity = (item.get("entity") or "OTHER").strip().upper()
        return f"{entity}||{partner_field}"

    candidates = [
        {
            **item,
            "partner_field": _classifier_id(item),
            "_original_partner_field": item.get("partner_field") or item.get("field_name"),
        }
        for item in all_mappings
        if _is_parameter_bucket_candidate(item)
        and (item.get("partner_field") or item.get("field_name"))
    ]

    if not candidates:
        logger.info("No parameter-bucket candidates found for classifier step")
        return all_mappings

    logger.info(
        "Preparing parameter bucket classifier input: total_candidates=%d",
        len(candidates),
    )

    classifier = LLMService(settings)
    prompt_payload = build_parameter_classifier_prompt(
        candidates,
        client_name=client_name,
        process_name=process_name,
    )

    try:
        classified = classifier.classify_parameter_buckets(prompt_payload)
    except Exception as e:
        logger.error("Parameter bucket classifier failed: %s", e, exc_info=True)
        return all_mappings

    if not classified:
        logger.info("Parameter bucket classifier returned no usable results")
        return all_mappings

    by_field = {item["partner_field"]: item for item in classified if item.get("partner_field")}
    updated = 0
    to_loan_parameter = 0
    to_loan_applicant = 0

    for item in all_mappings:
        original_field_name = item.get("partner_field") or item.get("field_name")
        if not original_field_name:
            continue
        classifier_key = _classifier_id(item)
        if classifier_key not in by_field:
            continue
        classification = by_field[classifier_key]
        new_bucket = classification["matched_excel_key"]
        old_bucket = item.get("matched_excel_key")

        old_reasoning = (item.get("reasoning") or "").strip()
        classifier_reasoning = classification.get("reasoning", "").strip()

        item["matched_excel_key"] = new_bucket
        item["json_key"] = ""
        item["confidence"] = classification.get("confidence", item.get("confidence", 0.0))
        item["needs_review"] = classification.get("needs_review", item.get("needs_review", True))
        item["winning_engine"] = "llm_parameter_bucket"
        item["match_type"] = "llm_parameter_bucket"

        item.setdefault("previous_mapping_reason", "")
        item.setdefault("llm_change_reason", "")
        item.setdefault("llm_param_bucket_reason", "")
        item["llm_param_bucket_reason"] = (
            "LLM parameter bucket classifier decision: "
            f"old_bucket={old_bucket or '(none)'} new_bucket={new_bucket}. "
            f"entity={(item.get('entity') or 'OTHER')} "
            f"process_name={process_name or ''} "
            f"column_category={(item.get('column_category') or '')}. "
            f"Reason: {classifier_reasoning}"
        ).strip()

        if old_reasoning:
            item["reasoning"] = f"{old_reasoning} → Parameter bucket classifier chose {new_bucket}. {classifier_reasoning}"
        else:
            item["reasoning"] = f"Parameter bucket classifier chose {new_bucket}. {classifier_reasoning}"

        updated += 1
        if new_bucket == "LOANAPPLICANTPARAM":
            to_loan_applicant += 1
        else:
            to_loan_parameter += 1
        logger.info(
            "Parameter classifier: field=%s old_bucket=%s new_bucket=%s confidence=%.4f",
            original_field_name,
            old_bucket or "(none)",
            new_bucket,
            item["confidence"],
        )

    logger.info(
        "Parameter bucket classifier summary: updated=%d LOANPARAMETER=%d LOANAPPLICANTPARAM=%d",
        updated,
        to_loan_parameter,
        to_loan_applicant,
    )
    return all_mappings


# ── Layer 3 — post-processing + Excel output ───────────────────────────────────
def _base_excel_key(excel_key: Optional[str]) -> str:
    value = (excel_key or "").strip()
    i = len(value) - 1
    while i >= 0 and value[i].isdigit():
        i -= 1
    return value[: i + 1]


def _apply_fee_entity_override(all_mappings: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Final business rule: if entity is FEE, force the mapping into the FEE
    bucket regardless of any earlier alias/fuzzy/LLM match.
    """
    from matching_engine import partner_field_excludes_generic_fee_bucket

    def _fee_entity_override_should_skip(row: Dict[str, Any]) -> bool:
        """Skip collapsing to generic FEE when the field is insurance or proc.-fee % pricing."""
        pf = (row.get("partner_field") or row.get("field_name") or "").strip()
        if partner_field_excludes_generic_fee_bucket(pf):
            return True
        ek = _base_excel_key(row.get("matched_excel_key") or "").upper()
        if ek.startswith("APPLICANTINSURANCE"):
            return True
        if "FEEPERCENTAGE" in ek:
            return True
        return False

    updated: List[Dict[str, Any]] = []
    overridden = 0
    skipped = 0

    for item in all_mappings:
        row = dict(item)
        entity = (row.get("entity") or "").upper()
        if entity == "FEE" and _base_excel_key(row.get("matched_excel_key")) != "FEE":
            if _fee_entity_override_should_skip(row):
                skipped += 1
                updated.append(row)
                continue
            old_key = row.get("matched_excel_key") or "(none)"
            old_reasoning = (row.get("reasoning") or "").strip()
            row["matched_excel_key"] = "FEE"
            row["json_key"] = ""
            row["match_type"] = "fee_entity_override"
            row["winning_engine"] = "fee_entity_override"
            if old_reasoning:
                row["reasoning"] = (
                    f"{old_reasoning} → Fee entity override applied: forced mapping from {old_key} to FEE."
                )
            else:
                row["reasoning"] = (
                    f"Fee entity override applied: forced mapping from {old_key} to FEE."
                )
            overridden += 1
        updated.append(row)

    if skipped:
        logger.info(
            "Skipped FEE entity override for %d mapping(s) (insurance / processing-fee %% / structured key)",
            skipped,
        )
    if overridden:
        logger.info("Applied final FEE entity override to %d mapping(s)", overridden)
    return updated


def _align_putm_key_to_row_entity(
    row: Dict[str, Any],
    by_excel_key: Dict[str, Any],
) -> Dict[str, Any]:
    """
    Correct obvious entity/key mismatches (e.g. APPLICANT row with COAPPLICANT1*
    or LOANCOAPP* PUTM keys) when the equivalent APPLICANT* / LOANAPPLICANTPARAM*
    key exists in the field dictionary.
    """
    ek = (row.get("matched_excel_key") or "").strip()
    if not ek or not isinstance(row, dict):
        return row

    ent = (row.get("entity") or "OTHER").strip().upper()
    if not by_excel_key:
        return row

    def _note_append(base: str, extra: str) -> str:
        b = (base or "").strip()
        return f"{b}; {extra}".strip("; ") if b else extra

    out = dict(row)

    if ent in ("APPLICANT", "CUSTOMER"):
        from matching_engine import (
            _build_by_excel_key_index,
            _is_coapplicant_scoped_putm_excel_key,
            _remap_putm_key_for_applicant_customer_entity,
        )
        idx = _build_by_excel_key_index({"by_excel_key": by_excel_key})
        rem = _remap_putm_key_for_applicant_customer_entity(ent, ek, idx)
        if rem:
            cand, jk_c = rem
            info = by_excel_key.get(cand) or {}
            out["matched_excel_key"] = cand
            out["json_key"] = info.get("json_key") or jk_c or out.get("json_key") or ""
            out["needs_review"] = True
            msg = f"Entity alignment ({ent}): retargeted {ek} → {cand}."
            out["reasoning"] = _note_append(out.get("reasoning", ""), msg)
            lr = (out.get("llm_change_reason") or "").strip()
            out["llm_change_reason"] = f"{lr}; {msg}".strip("; ") if lr else msg
            mt = (out.get("match_type") or "").strip()
            out["match_type"] = f"{mt}_entity_aligned".strip("_") if mt else "entity_aligned"
            return out
        if _is_coapplicant_scoped_putm_excel_key(ek):
            revert_key = _recover_prior_loanparameter_excel_key(out, by_excel_key)
            if revert_key:
                info = by_excel_key.get(revert_key) or {}
                msg = (
                    f"Entity alignment ({ent}): reverted co-applicant-only key {ek} → "
                    f"{revert_key} (no applicant-scoped PUTM equivalent; stay on parameter bucket)."
                )
                out["matched_excel_key"] = revert_key
                out["json_key"] = info.get("json_key") or out.get("json_key") or ""
                out["reasoning"] = _note_append(out.get("reasoning", ""), msg)
                lr = (out.get("llm_change_reason") or "").strip()
                out["llm_change_reason"] = f"{lr}; {msg}".strip("; ") if lr else msg
                mt = (out.get("match_type") or "").strip()
                out["match_type"] = f"{mt}_entity_reverted_to_loanparameter".strip("_") if mt else (
                    "entity_reverted_to_loanparameter"
                )
                out["needs_review"] = True
                out["winning_engine"] = out.get("winning_engine") or "deterministic"
                return out
            warn = (
                f"[needs_review] entity={ent} but matched_excel_key={ek} is "
                "co-applicant-scoped and no applicant-safe equivalent exists in PUTM."
            )
            out["reasoning"] = _note_append(out.get("reasoning", ""), warn)
            out["needs_review"] = True
            return out

    # COAPPLICANT* rows should not keep applicant-scoped PUTM keys when a
    # catalogue co-applicant equivalent exists (same rules as alias matching).
    if ent.startswith("COAPPLICANT"):
        from matching_engine import (
            _build_by_excel_key_index,
            _is_applicant_scoped_putm_excel_key,
            _remap_applicant_putm_to_coapplicant_catalogue,
        )
        if _is_applicant_scoped_putm_excel_key(ek):
            idx = _build_by_excel_key_index({"by_excel_key": by_excel_key})
            rem = _remap_applicant_putm_to_coapplicant_catalogue(
                ent, ek, idx, {}, {}, None,
            )
            if rem:
                cand, jk_c = rem
                info = by_excel_key.get(cand) or {}
                out["matched_excel_key"] = cand
                out["json_key"] = info.get("json_key") or jk_c or out.get("json_key") or ""
                out["needs_review"] = True
                msg = f"Entity alignment ({ent}): retargeted applicant-scoped key {ek} → {cand}."
                out["reasoning"] = _note_append(out.get("reasoning", ""), msg)
                lr = (out.get("llm_change_reason") or "").strip()
                out["llm_change_reason"] = f"{lr}; {msg}".strip("; ") if lr else msg
                mt = (out.get("match_type") or "").strip()
                out["match_type"] = f"{mt}_entity_aligned".strip("_") if mt else "entity_aligned"
                return out
            idx_suffix = ent.replace("COAPPLICANT", "").strip() or "1"
            cand = re.sub(r"^APPLICANT", f"COAPPLICANT{idx_suffix}", ek, count=1, flags=re.I).upper()
            if cand in by_excel_key:
                info = by_excel_key.get(cand) or {}
                out["matched_excel_key"] = cand
                out["json_key"] = info.get("json_key") or out.get("json_key") or ""
                out["needs_review"] = True
                msg = (
                    f"Entity alignment ({ent}): retargeted applicant key {ek} → {cand}."
                )
                out["reasoning"] = _note_append(out.get("reasoning", ""), msg)
                lr = (out.get("llm_change_reason") or "").strip()
                out["llm_change_reason"] = f"{lr}; {msg}".strip("; ") if lr else msg
                mt = (out.get("match_type") or "").strip()
                out["match_type"] = f"{mt}_entity_aligned".strip("_") if mt else "entity_aligned"
                return out

    return out


def _apply_entity_excel_key_alignment(
    rows: List[Dict[str, Any]],
    by_excel_key: Dict[str, Any],
) -> List[Dict[str, Any]]:
    if not rows or not by_excel_key:
        return rows
    fixed = 0
    out: List[Dict[str, Any]] = []
    for r in rows:
        if not isinstance(r, dict):
            out.append(r)
            continue
        before = (r.get("matched_excel_key") or "").strip()
        aligned = _align_putm_key_to_row_entity(r, by_excel_key)
        if (aligned.get("matched_excel_key") or "").strip() != before:
            fixed += 1
        out.append(aligned)
    if fixed:
        logger.info(
            "Entity/excel_key alignment: adjusted %d row(s) to match row entity vs PUTM prefix",
            fixed,
        )
    return out


def dedupe_one_row_per_excel_key(mappings: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    PUTM excel keys are unique in the target: if multiple partner rows map to the
    same matched_excel_key (case-insensitive), keep one row — highest confidence
    wins; on a tie, the earliest row in the input list wins. Entity is ignored
    for grouping so e.g. two sources cannot both claim APPLICATIONFILEID.
    Rows without matched_excel_key are left unchanged.
    """
    if not mappings or len(mappings) < 2:
        return mappings

    def _identity_key(m: Dict[str, Any]) -> Optional[str]:
        ek = (m.get("matched_excel_key") or "").strip()
        if not ek:
            return None
        return ek.upper()

    best_idx: Dict[str, int] = {}
    best_conf: Dict[str, float] = {}
    for i, m in enumerate(mappings):
        key = _identity_key(m)
        if key is None:
            continue
        c = float(m.get("confidence") or 0.0)
        if key not in best_conf or c > best_conf[key]:
            best_conf[key] = c
            best_idx[key] = i

    out: List[Dict[str, Any]] = []
    for i, m in enumerate(mappings):
        key = _identity_key(m)
        if key is None:
            out.append(m)
            continue
        if best_idx.get(key) == i:
            out.append(m)

    dropped = len(mappings) - len(out)
    if dropped:
        logger.info(
            "dedupe_one_row_per_excel_key: removed %d duplicate row(s) "
            "sharing the same matched_excel_key",
            dropped,
        )
    return out


def finalize_mappings(
    all_mappings: List[Dict[str, Any]],
    settings,
) -> List[Dict[str, Any]]:
    """
    Apply final overrides and resolve numbered keys/json paths once so every
    downstream consumer sees the same final mapping output.
    """
    _add_scripts_to_path(settings.scripts_dir)

    from matching_engine import load_references
    from post_processor import PostProcessor

    adjusted = _apply_fee_entity_override(all_mappings)
    refs = load_references(settings.references_dir)
    by_ek = refs.get("field_dictionary", {}).get("by_excel_key") or {}
    adjusted = _apply_entity_excel_key_alignment(adjusted, by_ek)
    valid_mappings = [m for m in adjusted if m.get("matched_excel_key")]
    unmatched_mappings = [m for m in adjusted if not m.get("matched_excel_key")]

    processor = PostProcessor(refs["field_dictionary"])
    processed_valid = processor.process_results(valid_mappings)
    processed_valid = dedupe_one_row_per_excel_key(processed_valid)
    return processed_valid + unmatched_mappings


def post_process_and_output(
    all_mappings: List[Dict],
    settings,
    output_path: str,
    client_name: str,
    process_name: str,
) -> str:
    _add_scripts_to_path(settings.scripts_dir)

    try:
        from generate_output import generate_mapping_excel
        use_original_excel = True
    except ImportError as e:
        logger.warning(f"Failed to import generate_mapping_excel: {e}. Falling back to pandas.")
        use_original_excel = False

    valid_mappings     = [m for m in all_mappings if m.get("matched_excel_key")]
    unmatched_mappings = [m for m in all_mappings if not m.get("matched_excel_key")]
    valid_mappings     = dedupe_one_row_per_excel_key(valid_mappings)

    if unmatched_mappings:
        logger.warning(
            f"{len(unmatched_mappings)} mappings have no matched_excel_key "
            f"— they will appear as unmatched in output: "
            f"{[m.get('partner_field') for m in unmatched_mappings]}"
        )

    processed_mappings = valid_mappings + unmatched_mappings

    logger.info(
        f"Output generation: {len(valid_mappings)} valid, "
        f"{len(unmatched_mappings)} unmatched, "
        f"{len(processed_mappings)} total"
    )

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    if use_original_excel:
        generate_mapping_excel(processed_mappings, output_path, client_name, process_name)
        logger.info(f"Excel written (original format) → {output_path}")
    else:
        try:
            import pandas as pd
            df = pd.DataFrame(processed_mappings)
            df.to_excel(output_path, index=False, engine='openpyxl')
            logger.info(f"Excel written (fallback) → {output_path}")
        except ImportError:
            from openpyxl import Workbook
            wb = Workbook()
            ws = wb.active
            if processed_mappings:
                headers = list(processed_mappings[0].keys())
                ws.append(headers)
                for row in processed_mappings:
                    ws.append([row.get(h, "") for h in headers])
            wb.save(output_path)
            logger.info(f"Excel written (openpyxl fallback) → {output_path}")

    return output_path