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
from typing import Any, Dict, List, Optional, Tuple

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

    if entity in {"APPLICANT", "CUSTOMER", "COAPPLICANT", "COAPPLICANT1", "COAPPLICANT2", "COAPPLICANT3", "COAPPLICANT4"}:
        if compact_field not in LOAN_LEVEL_FIELD_HINTS:
            return "LOANAPPLICANTPARAM"

    if any(hint in category for hint in CUSTOMER_PARAM_CATEGORY_HINTS):
        if compact_field not in LOAN_LEVEL_FIELD_HINTS:
            return "LOANAPPLICANTPARAM"

    if (
        entity == "LOAN"
        and any(hint in category for hint in LOAN_LEVEL_CATEGORY_HINTS)
        and any(hint in field for hint in CUSTOMER_PARAM_CATEGORY_HINTS)
    ):
        return "LOANAPPLICANTPARAM"

    return "LOANPARAMETER"


def _is_parameter_bucket_candidate(item: Dict[str, Any]) -> bool:
    excel_key = (item.get("matched_excel_key") or "").upper()
    return (
        not excel_key
        or excel_key.startswith("LOANPARAMETER")
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


# ── Layer 1 — deterministic ────────────────────────────────────────────────────
def run_deterministic(
    input_file: str,
    settings,
    process_name: str = "COMBINED",
    sheet_filter: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Layer 1: deterministic matching.

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

    batch     = match_batch(fields, refs, process_name=process_name)
    matched   = batch["matched"]
    unmatched = batch["unmatched"]

    det_dicts = []
    for m in matched:
        det_dicts.append({
            "partner_field":     m.partner_field,
            "column_category":   m.column_category,
            "entity":            m.entity,
            "matched_excel_key": m.matched_excel_key,
            "json_key":          m.matched_json_key or "",
            "confidence":        m.confidence,
            "match_type":        m.match_type,
            "reasoning":         m.reasoning,
            "needs_review":      m.confidence < settings.review_threshold,
            "winning_engine":    "deterministic",
        })

    unmatched_dicts = [
        {
            "field_name":      u.partner_field,
            "partner_field":   u.partner_field,
            "column_category": u.column_category,
            "entity":          u.entity,
        }
        for u in unmatched
    ]

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


# ── Layer 2 — hybrid + LLM ────────────────────────────────────────────────────
def run_hybrid_llm(
    unmatched_fields: List[Dict],
    field_dictionary: Dict,
    alias_registry: Dict,
    entity_prompts: List[Dict],
    settings,
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

    remaining = list(unmatched_fields)
    results:   List[Dict] = []
    breakdown = {"fuzzy": 0, "embedding": 0, "llm": 0, "unmatched": 0}
    review_threshold = getattr(settings, "review_threshold", 0.80)
    review_candidates_by_field: Dict[str, Dict] = {}

    def _field_id(item: Dict[str, Any]) -> str:
        return item.get("partner_field") or item.get("field_name") or ""

    def _candidate_payload(item: Dict[str, Any]) -> Dict[str, Any]:
        return {
            "engine": item.get("winning_engine") or item.get("match_type"),
            "matched_excel_key": item.get("matched_excel_key"),
            "json_key": item.get("json_key", ""),
            "confidence": item.get("confidence", 0.0),
            "reasoning": item.get("reasoning", ""),
        }

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
            entry.update({
                "matched_excel_key": item.get("matched_excel_key"),
                "json_key": item.get("json_key", ""),
                "confidence": item.get("confidence", 0.0),
                "match_type": item.get("match_type", entry.get("match_type")),
                "reasoning": item.get("reasoning", ""),
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

    def _dedupe_fields(items: List[Dict]) -> List[Dict]:
        deduped: Dict[str, Dict] = {}
        for item in items:
            field_id = _field_id(item)
            if field_id:
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

    # 2a — Fuzzy
    if use_fuzzy and remaining and HAS_RAPIDFUZZ:
        try:
            fe = FuzzyEngine(field_dictionary, threshold=settings.fuzzy_threshold)
            fuzzy_matched, remaining = fe.run_batch(remaining)
            accepted, review = _split_by_confidence(fuzzy_matched)
            if review:
                logger.info(
                    "FuzzyEngine routed %d low-confidence fields to LLM review "
                    "(threshold=%.2f)",
                    len(review),
                    review_threshold,
                )
            results.extend(accepted)
            breakdown["fuzzy"] = len(accepted)
            remaining = _dedupe_fields(remaining + review)
        except Exception as e:
            logger.warning(f"FuzzyEngine skipped: {e}")

    # 2b — Embeddings
    if use_embeddings and remaining and HAS_ST:
        try:
            ee = EmbeddingEngine(field_dictionary, threshold=settings.embedding_threshold)
            emb_matched, remaining = ee.run_batch(remaining)
            accepted, review = _split_by_confidence(emb_matched)
            if review:
                logger.info(
                    "EmbeddingEngine routed %d low-confidence fields to LLM review "
                    "(threshold=%.2f)",
                    len(review),
                    review_threshold,
                )
            results.extend(accepted)
            breakdown["embedding"] = len(accepted)
            remaining = _dedupe_fields(remaining + review)
        except Exception as e:
            logger.warning(f"EmbeddingEngine skipped: {e}")

    # 2c — LLM
    # Re-build prompts scoped to remaining fields only.
    # rendered_prompt = context block → posted as `task` to gateway.
    llm_input_fields = list(review_candidates_by_field.values()) + [
        item for item in remaining
        if _field_id(item) not in review_candidates_by_field
    ]
    if use_llm and llm_input_fields:
        try:
            low_conf_fields_for_llm = list(review_candidates_by_field.values())
            low_conf_candidate_count = sum(
                len(item.get("candidate_matches") or [])
                for item in low_conf_fields_for_llm
            )
            logger.info(
                "Preparing LLM input: total_fields=%d, low_confidence_fields=%d, "
                "low_confidence_candidates=%d, threshold=%.2f",
                len(llm_input_fields),
                len(low_conf_fields_for_llm),
                low_conf_candidate_count,
                review_threshold,
            )
            fresh_prompts = build_entity_prompts(
                unmatched_fields=llm_input_fields,
                field_dictionary=field_dictionary,
                alias_registry=alias_registry,
                prompt_template="",   # gateway owns the base template
                client_name=client_name,
                process_name=process_name,
            )
            llm_svc      = LLMService(settings)
            llm_mappings = llm_svc.map_fields(fresh_prompts)
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

    candidates = [
        {
            **item,
            "partner_field": item.get("partner_field") or item.get("field_name"),
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

    by_field = {item["partner_field"]: item for item in classified}
    updated = 0
    to_loan_parameter = 0
    to_loan_applicant = 0

    for item in all_mappings:
        field_name = item.get("partner_field") or item.get("field_name")
        if not field_name or field_name not in by_field:
            continue
        classification = by_field[field_name]
        new_bucket = classification["matched_excel_key"]
        old_bucket = item.get("matched_excel_key")

        item["matched_excel_key"] = new_bucket
        item["json_key"] = ""
        item["confidence"] = classification.get("confidence", item.get("confidence", 0.0))
        item["needs_review"] = classification.get("needs_review", item.get("needs_review", True))
        item["winning_engine"] = "llm_parameter_bucket"
        item["match_type"] = "llm_parameter_bucket"
        item["reasoning"] = (
            f"Parameter bucket classifier chose {new_bucket}. "
            f"{classification.get('reasoning', '').strip()}"
        ).strip()
        updated += 1
        if new_bucket == "LOANAPPLICANTPARAM":
            to_loan_applicant += 1
        else:
            to_loan_parameter += 1
        logger.info(
            "Parameter classifier: field=%s old_bucket=%s new_bucket=%s confidence=%.4f",
            field_name,
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
def post_process_and_output(
    all_mappings: List[Dict],
    settings,
    output_path: str,
    client_name: str,
    process_name: str,
) -> str:
    _add_scripts_to_path(settings.scripts_dir)

    try:
        from matching_engine import load_references
        from post_processor import PostProcessor
        from generate_output import generate_mapping_excel
        use_original_excel = True
    except ImportError as e:
        logger.warning(f"Failed to import generate_mapping_excel: {e}. Falling back to pandas.")
        use_original_excel = False
        from matching_engine import load_references
        from post_processor import PostProcessor

    refs = load_references(settings.references_dir)
    valid_mappings     = [m for m in all_mappings if m.get("matched_excel_key")]
    unmatched_mappings = [m for m in all_mappings if not m.get("matched_excel_key")]

    if unmatched_mappings:
        logger.warning(
            f"{len(unmatched_mappings)} mappings have no matched_excel_key "
            f"— they will appear as unmatched in output: "
            f"{[m.get('partner_field') for m in unmatched_mappings]}"
        )

    processor = PostProcessor(refs["field_dictionary"])
    processed_valid    = processor.process_results(valid_mappings)
    processed_mappings = processed_valid + unmatched_mappings

    logger.info(
        f"Post-processing: {len(processed_valid)} valid, "
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
