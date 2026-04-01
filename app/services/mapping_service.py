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
import sys
import importlib.util
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)


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
    try:
        return __import__(module_name)
    except ImportError as e:
        logger.debug(f"Regular import failed: {e}, trying spec import")
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

    # 2a — Fuzzy
    if use_fuzzy and remaining and HAS_RAPIDFUZZ:
        try:
            fe = FuzzyEngine(field_dictionary, threshold=settings.fuzzy_threshold)
            fuzzy_matched, remaining = fe.run_batch(remaining)
            results.extend(fuzzy_matched)
            breakdown["fuzzy"] = len(fuzzy_matched)
        except Exception as e:
            logger.warning(f"FuzzyEngine skipped: {e}")

    # 2b — Embeddings
    if use_embeddings and remaining and HAS_ST:
        try:
            ee = EmbeddingEngine(field_dictionary, threshold=settings.embedding_threshold)
            emb_matched, remaining = ee.run_batch(remaining)
            results.extend(emb_matched)
            breakdown["embedding"] = len(emb_matched)
        except Exception as e:
            logger.warning(f"EmbeddingEngine skipped: {e}")

    # 2c — LLM
    # Re-build prompts scoped to remaining fields only.
    # rendered_prompt = context block → posted as `task` to gateway.
    if use_llm and remaining:
        try:
            fresh_prompts = build_entity_prompts(
                unmatched_fields=remaining,
                field_dictionary=field_dictionary,
                alias_registry=alias_registry,
                prompt_template="",   # gateway owns the base template
                client_name=client_name,
                process_name=process_name,
            )
            llm_svc      = LLMService(settings)
            llm_mappings = llm_svc.map_fields(fresh_prompts)
            results.extend(llm_mappings)
            breakdown["llm"] = len(llm_mappings)
            remaining = []
        except Exception as e:
            logger.error(f"LLMService error: {e}", exc_info=True)

    # Still unmatched after all engines
    for f in remaining:
        results.append({
            **f,
            "matched_excel_key": None,
            "json_key":          "",
            "confidence":        0.0,
            "match_type":        "unmatched",
            "reasoning":         "No match found in any engine",
            "needs_review":      True,
            "winning_engine":    "none",
        })
    breakdown["unmatched"] = len(remaining)

    return results, breakdown


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
    flat_dict = {
        ek: info.get("json_key", "")
        for ek, info in refs["field_dictionary"].get("by_excel_key", {}).items()
    }

    valid_mappings     = [m for m in all_mappings if m.get("matched_excel_key")]
    unmatched_mappings = [m for m in all_mappings if not m.get("matched_excel_key")]

    if unmatched_mappings:
        logger.warning(
            f"{len(unmatched_mappings)} mappings have no matched_excel_key "
            f"— they will appear as unmatched in output: "
            f"{[m.get('partner_field') for m in unmatched_mappings]}"
        )

    processor = PostProcessor(flat_dict)
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