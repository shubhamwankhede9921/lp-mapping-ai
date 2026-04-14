# app/api/mapping_controller.py

import logging
import os
import re
import shutil
import tempfile
import time
import json
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, Optional

from fastapi import (
    APIRouter, BackgroundTasks, Depends, File,
    Form, HTTPException, UploadFile,
)
from fastapi.responses import Response
from sqlalchemy.orm import Session

from app.config import Settings, get_settings
from app.models.schemas import (
    BuildRefsRequest, BuildRefsResponse,
    DeterministicResponse, HybridLLMResponse,
    FieldMapping, Stats,
    LOSJsonRequest,
    NestedMappingResponse,
    SchemaResponse,
)
from app.repository import database as db_repo
from app.repository.db_writer import upsert_mappings   # opens its own target-DB connection
from app.services import mapping_service as svc

from app.utils.los_json_builder import (
    _is_skippable,
    _normalise_mappings,
    generate_nested_mapping,
    generate_schema,
)

try:
    from app.repository.database import get_db
except ImportError:
    def get_db():
        yield None

logger = logging.getLogger(__name__)
router = APIRouter(prefix="/api/llm_mapping", tags=["LP Field Mapping"])


def _run_build_references_from_db(
    settings: Settings,
    putm_table_override: Optional[str] = None,
    mapping_table_override: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Extract PUTM + generic mapping from the source DB, run build_references.py outputs
    into references_dir, and return the same stats payload as POST /references/build.
    """
    from app.repository.database import extract_putm_dump, extract_generic_mapping

    dumps_dir = Path(settings.references_dir) / "dumps"
    dumps_dir.mkdir(parents=True, exist_ok=True)

    putm_path = dumps_dir / "putm_dump.xlsx"
    mapping_path = dumps_dir / "generic_mapping.csv"

    putm_rows = extract_putm_dump(settings, str(putm_path), putm_table_override)
    mapping_rows = extract_generic_mapping(settings, str(mapping_path), mapping_table_override)

    result = svc.build_references_from_db_direct(
        putm_xlsx=str(putm_path),
        mapping_csv=str(mapping_path),
        references_dir=settings.references_dir,
        scripts_dir=settings.scripts_dir,
    )

    return {**result, "putm_rows": putm_rows, "mapping_rows": mapping_rows}


# ── shared helpers ─────────────────────────────────────────────────────────────

def _to_field_mapping(m: dict) -> FieldMapping:
    return FieldMapping(
        partner_field=m.get("partner_field", ""),
        column_category=m.get("column_category"),
        entity=m.get("entity", "OTHER"),
        matched_excel_key=m.get("matched_excel_key"),
        json_key=m.get("json_key"),
        confidence=m.get("confidence", 0.0),
        match_type=m.get("match_type", "unmatched"),
        reasoning=m.get("reasoning", ""),
        needs_review=m.get("needs_review", False),
        fuzzy_score=m.get("fuzzy_score"),
        embedding_score=m.get("embedding_score"),
        llm_confidence=m.get("llm_confidence"),
        winning_engine=m.get("winning_engine"),
    )


def _to_stats(raw: dict) -> Stats:
    return Stats(**raw)


def _save_upload(file: UploadFile) -> str:
    suffix = Path(file.filename).suffix
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    shutil.copyfileobj(file.file, tmp)
    tmp.close()
    return tmp.name


def _safe_unlink(path: Optional[str]) -> None:
    if path and os.path.exists(path):
        try:
            os.unlink(path)
        except PermissionError:
            logger.debug(f"Could not delete temp file (Windows lock): {path}")


def _sanitize_path_component(value: str, fallback: str = "output") -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", (value or "").strip())
    cleaned = re.sub(r"\s+", "_", cleaned).strip(" ._")
    return cleaned or fallback


def _request_to_builder_input(request: LOSJsonRequest) -> Dict:
    mappings = []
    for m in request.mappings:
        row = m.dict()
        if not row.get("lms_column") and row.get("json_key"):
            row["lms_column"] = row["json_key"]
        mappings.append(row)
    return {"mappings": mappings}


def _count_stats(request: LOSJsonRequest) -> Dict[str, int]:
    flat = _normalise_mappings([
        {**m.dict(), "lms_column": m.get_lms_column()}
        for m in request.mappings
    ])
    total   = len(flat)
    skipped = sum(
        1 for m in flat
        if _is_skippable(m.get("lms_column"), m.get("matched_excel_key"))
        or not (m.get("client_column") or "").strip()
    )
    return {
        "total_input":   total,
        "skipped_count": skipped,
        "mapped_count":  total - skipped,
    }


# ── 1. Build References ────────────────────────────────────────────────────────

@router.post("/references/build", summary="Extract DB tables and build reference JSON files")
async def build_references(
    settings: Settings = Depends(get_settings),
    putm_table_override: Optional[str] = None,
    mapping_table_override: Optional[str] = None,
) -> Dict[str, Any]:
    try:
        data = _run_build_references_from_db(
            settings,
            putm_table_override=putm_table_override,
            mapping_table_override=mapping_table_override,
        )

        return {
            "status":  "success",
            "message": "References built successfully",
            "data": data,
        }

    except FileNotFoundError as e:
        raise HTTPException(status_code=500, detail=f"Required file not found: {e}")
    except ImportError as e:
        raise HTTPException(status_code=500, detail=f"Failed to import required modules: {e}")
    except Exception as e:
        logger.exception("Reference build failed")
        raise HTTPException(status_code=500, detail=f"Failed to build references: {e}")


# ── 2. Deterministic ──────────────────────────────────────────────────────────

@router.post(
    "/mapping/deterministic",
    response_model=DeterministicResponse,
    summary="Phase 1 — deterministic alias + exact matching",
)
async def deterministic(
    file: UploadFile  = File(...),
    client_name: str  = Form(...),
    process_name: str = Form("COMBINED"),
    sheet_filter: str = Form(None),
    use_loanparameter_refinement: bool = Form(
        True,
        description="If true and LOANPARAMETER_REFINEMENT_GATEWAY_URL is set, remap LOANPARAMETER* via PUTM LLM",
    ),
    use_llm_entity_classifier: bool = Form(
        False,
        description="If true and ENTITY_CLASSIFIER_GATEWAY_URL is set, assign entity via LLM before deterministic matching",
    ),
    settings: Settings = Depends(get_settings),
):
    # ── DEBUG: confirm settings loaded correctly ───────────────────
    logger.info("SETTINGS_DEBUG [deterministic]: settings object id = %d", id(settings))
    logger.info("SETTINGS_DEBUG [deterministic]: llm_gateway_url = %r", settings.llm_gateway_url)
    logger.info("SETTINGS_DEBUG [deterministic]: loanparameter_refinement_gateway_url = %r", settings.loanparameter_refinement_gateway_url)
    logger.info("SETTINGS_DEBUG [deterministic]: parameter_classifier_gateway_url = %r", settings.parameter_classifier_gateway_url)
    logger.info("SETTINGS_DEBUG [deterministic]: entity_classifier_gateway_url = %r", settings.entity_classifier_gateway_url)
    logger.info("SETTINGS_DEBUG [deterministic]: use_loanparameter_refinement (form) = %r", use_loanparameter_refinement)
    logger.info("SETTINGS_DEBUG [deterministic]: use_llm_entity_classifier (form) = %r", use_llm_entity_classifier)
    # ──────────────────────────────────────────────────────────────

    tmp = None
    try:
        tmp    = _save_upload(file)
        result = svc.run_deterministic(
            input_file=tmp,
            settings=settings,
            process_name=process_name,
            sheet_filter=sheet_filter or None,
            client_name=client_name,
            use_llm_entity_classifier=use_llm_entity_classifier,
        )
        det       = result.get("deterministic_results") or []
        fd0       = result["field_dictionary"]
        ar0       = result["alias_registry"]

        logger.info("SETTINGS_DEBUG [deterministic]: det_results count = %d", len(det))
        lp_count = sum(1 for r in det if (r.get("matched_excel_key") or "").upper().startswith("LOANPARAMETER"))
        logger.info("SETTINGS_DEBUG [deterministic]: LOANPARAMETER* rows = %d", lp_count)

        if use_loanparameter_refinement:
            logger.info("SETTINGS_DEBUG [deterministic]: calling refine_loanparameter_after_deterministic ...")
            det = svc.refine_loanparameter_after_deterministic(
                deterministic_results=det,
                field_dictionary=fd0,
                alias_registry=ar0,
                settings=settings,
                client_name=client_name,
                process_name=process_name,
            ) or det
            logger.info("SETTINGS_DEBUG [deterministic]: refine returned %d rows", len(det))
        else:
            logger.info("SETTINGS_DEBUG [deterministic]: skipping refinement — use_loanparameter_refinement=False")

        unm       = result["unmatched_fields"]
        eps       = result["entity_prompts"]
        stats_raw = svc.compute_stats(det)

        return DeterministicResponse(
            client_name=client_name,
            process_name=process_name,
            mappings=[_to_field_mapping(m) for m in det],
            unmatched_fields=unm,
            llm_prompts_count=len(eps),
            stats=_to_stats(stats_raw),
        )
    except FileNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        logger.exception("Deterministic mapping failed")
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        _safe_unlink(tmp)


# ── 3. Hybrid + LLM ──────────────────────────────────────────────────────────

@router.post(
    "/mapping/hybrid-llm",
    response_model=HybridLLMResponse,
    summary="Phase 1 + 2 — deterministic then fuzzy/embedding/LLM",
)
async def hybrid_llm(
    file: UploadFile         = File(...),
    client_name: str         = Form(...),
    process_name: str        = Form("COMBINED"),
    use_fuzzy: bool          = Form(True),
    use_embeddings: bool     = Form(False),
    use_llm: bool            = Form(True),
    use_loanparameter_refinement: bool = Form(
        True,
        description="Remap deterministic LOANPARAMETER* via dedicated PUTM LLM when URL configured",
    ),
    use_llm_entity_classifier: bool = Form(
        False,
        description="If true and ENTITY_CLASSIFIER_GATEWAY_URL is set, assign entity via LLM before deterministic matching",
    ),
    master_id: Optional[int] = Form(None),
    save_to_db: bool         = Form(False),
    skip_unmatched: bool     = Form(False),
    settings: Settings       = Depends(get_settings),
):
    # ── DEBUG: confirm settings loaded correctly ───────────────────
    logger.info("SETTINGS_DEBUG [hybrid_llm]: settings object id = %d", id(settings))
    logger.info("SETTINGS_DEBUG [hybrid_llm]: llm_gateway_url = %r", settings.llm_gateway_url)
    logger.info("SETTINGS_DEBUG [hybrid_llm]: loanparameter_refinement_gateway_url = %r", settings.loanparameter_refinement_gateway_url)
    logger.info("SETTINGS_DEBUG [hybrid_llm]: parameter_classifier_gateway_url = %r", settings.parameter_classifier_gateway_url)
    logger.info("SETTINGS_DEBUG [hybrid_llm]: entity_classifier_gateway_url = %r", settings.entity_classifier_gateway_url)
    logger.info("SETTINGS_DEBUG [hybrid_llm]: use_loanparameter_refinement (form) = %r", use_loanparameter_refinement)
    logger.info("SETTINGS_DEBUG [hybrid_llm]: use_llm_entity_classifier (form) = %r", use_llm_entity_classifier)
    # ──────────────────────────────────────────────────────────────

    tmp = None
    try:
        tmp = _save_upload(file)

        p1          = svc.run_deterministic(
            input_file=tmp,
            settings=settings,
            process_name=process_name,
            client_name=client_name,
            use_llm_entity_classifier=use_llm_entity_classifier,
        )
        det_results = p1.get("deterministic_results") or []
        unmatched   = p1["unmatched_fields"]
        ep          = p1["entity_prompts"]
        fd          = p1["field_dictionary"]
        ar          = p1["alias_registry"]

        logger.info("SETTINGS_DEBUG [hybrid_llm]: det_results count = %d", len(det_results))
        lp_count = sum(1 for r in det_results if (r.get("matched_excel_key") or "").upper().startswith("LOANPARAMETER"))
        logger.info("SETTINGS_DEBUG [hybrid_llm]: LOANPARAMETER* rows = %d", lp_count)

        if use_loanparameter_refinement:
            logger.info("SETTINGS_DEBUG [hybrid_llm]: calling refine_loanparameter_after_deterministic ...")
            det_results = svc.refine_loanparameter_after_deterministic(
                deterministic_results=det_results,
                field_dictionary=fd,
                alias_registry=ar,
                settings=settings,
                client_name=client_name,
                process_name=process_name,
            ) or det_results
            logger.info("SETTINGS_DEBUG [hybrid_llm]: refine returned %d rows", len(det_results))
        else:
            logger.info("SETTINGS_DEBUG [hybrid_llm]: skipping refinement — use_loanparameter_refinement=False")

        phase2, breakdown = svc.run_hybrid_llm(
            unmatched_fields=unmatched,
            field_dictionary=fd,
            alias_registry=ar,
            entity_prompts=ep,
            deterministic_matches=det_results,
            settings=settings,
            use_fuzzy=use_fuzzy,
            use_embeddings=use_embeddings,
            use_llm=use_llm,
            client_name=client_name,
            process_name=process_name,
        )

        all_mappings = svc.merge_deterministic_with_hybrid_phase(det_results, phase2)
        all_mappings = svc.refine_parameter_buckets(
            all_mappings=all_mappings,
            settings=settings,
            client_name=client_name,
            process_name=process_name,
        )
        all_mappings = svc.finalize_mappings(
            all_mappings=all_mappings,
            settings=settings,
        )
        stats_raw    = svc.compute_stats(all_mappings)

        if save_to_db and master_id is not None:
            try:
                db_result = upsert_mappings(
                    mappings=all_mappings,
                    master_id=master_id,
                    client_name=client_name,
                    process_name=process_name,
                    settings=settings,
                    skip_unmatched=skip_unmatched,
                )
                logger.info(f"DB write (hybrid-llm): {db_result}")
            except Exception as db_exc:
                logger.error(f"DB write failed (non-fatal): {db_exc}")

        return HybridLLMResponse(
            client_name=client_name,
            process_name=process_name,
            mappings=[_to_field_mapping(m) for m in all_mappings],
            stats=_to_stats(stats_raw),
            engine_breakdown=breakdown,
        )

    except Exception as e:
        logger.exception("Hybrid+LLM mapping failed")
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        _safe_unlink(tmp)


# ── 4. Full Pipeline ──────────────────────────────────────────────────────────

@router.post(
    "/mapping/full-pipeline",
    summary="All phases → ZIP with Excel, nested mapping JSON, schema JSON, optional reference JSONs, and optional DB write",
)
async def full_pipeline(
    background_tasks: BackgroundTasks,
    file: UploadFile     = File(...),
    client_name: str     = Form(...),
    process_name: str    = Form("COMBINED"),
    use_fuzzy: bool      = Form(True),
    use_embeddings: bool = Form(False),
    use_llm: bool        = Form(True),
    use_loanparameter_refinement: bool = Form(
        True,
        description="Remap deterministic LOANPARAMETER* via dedicated PUTM LLM when URL configured",
    ),
    use_llm_entity_classifier: bool = Form(
        False,
        description="If true and ENTITY_CLASSIFIER_GATEWAY_URL is set, assign entity via LLM before deterministic matching",
    ),
    sheet_filter: str    = Form(None),
    master_id: int       = Form(...,   description="FK stored in master_id column of the mapping table"),
    save_to_db: bool     = Form(True,  description="Write results to GENERIC_MAPPING_TABLE in TARGET DB"),
    skip_unmatched: bool = Form(False, description="Skip rows with no matched_excel_key when writing to DB"),
    include_build_references: bool = Form(
        True,
        description=(
            "Extract from source DB, rebuild field_dictionary / alias_registry / entity_routing "
            "before mapping (same as POST /references/build), and add those JSON files under references/ in the ZIP"
        ),
    ),
    settings: Settings   = Depends(get_settings),
):
    # ── DEBUG: confirm settings loaded correctly ───────────────────
    logger.info("SETTINGS_DEBUG [full_pipeline]: settings object id = %d", id(settings))
    logger.info("SETTINGS_DEBUG [full_pipeline]: llm_gateway_url = %r", settings.llm_gateway_url)
    logger.info("SETTINGS_DEBUG [full_pipeline]: loanparameter_refinement_gateway_url = %r", settings.loanparameter_refinement_gateway_url)
    logger.info("SETTINGS_DEBUG [full_pipeline]: parameter_classifier_gateway_url = %r", settings.parameter_classifier_gateway_url)
    logger.info("SETTINGS_DEBUG [full_pipeline]: entity_classifier_gateway_url = %r", settings.entity_classifier_gateway_url)
    logger.info("SETTINGS_DEBUG [full_pipeline]: use_loanparameter_refinement (form) = %r", use_loanparameter_refinement)
    logger.info("SETTINGS_DEBUG [full_pipeline]: use_llm_entity_classifier (form) = %r", use_llm_entity_classifier)
    # ──────────────────────────────────────────────────────────────

    tmp = None
    ref_meta: Optional[Dict[str, Any]] = None
    try:
        tmp = _save_upload(file)

        if include_build_references:
            ref_meta = _run_build_references_from_db(settings)

        # ── Phase 1: deterministic ─────────────────────────────────────────────
        p1 = svc.run_deterministic(
            input_file=tmp,
            settings=settings,
            process_name=process_name,
            sheet_filter=sheet_filter or None,
            client_name=client_name,
            use_llm_entity_classifier=use_llm_entity_classifier,
        )
        det_results = p1.get("deterministic_results") or []
        unmatched   = p1["unmatched_fields"]
        ep          = p1["entity_prompts"]
        fd          = p1["field_dictionary"]
        ar          = p1["alias_registry"]

        logger.info("SETTINGS_DEBUG [full_pipeline]: det_results count = %d", len(det_results))
        lp_count = sum(1 for r in det_results if (r.get("matched_excel_key") or "").upper().startswith("LOANPARAMETER"))
        logger.info("SETTINGS_DEBUG [full_pipeline]: LOANPARAMETER* rows = %d", lp_count)

        if use_loanparameter_refinement:
            logger.info("SETTINGS_DEBUG [full_pipeline]: calling refine_loanparameter_after_deterministic ...")
            det_results = svc.refine_loanparameter_after_deterministic(
                deterministic_results=det_results,
                field_dictionary=fd,
                alias_registry=ar,
                settings=settings,
                client_name=client_name,
                process_name=process_name,
            ) or det_results
            logger.info("SETTINGS_DEBUG [full_pipeline]: refine returned %d rows", len(det_results))
        else:
            logger.info("SETTINGS_DEBUG [full_pipeline]: skipping refinement — use_loanparameter_refinement=False")

        # ── Phase 2: hybrid + LLM ──────────────────────────────────────────────
        phase2, _ = svc.run_hybrid_llm(
            unmatched_fields=unmatched,
            field_dictionary=fd,
            alias_registry=ar,
            entity_prompts=ep,
            deterministic_matches=det_results,
            settings=settings,
            use_fuzzy=use_fuzzy,
            use_embeddings=use_embeddings,
            use_llm=use_llm,
            client_name=client_name,
            process_name=process_name,
        )

        all_mappings = svc.merge_deterministic_with_hybrid_phase(det_results, phase2)
        all_mappings = svc.refine_parameter_buckets(
            all_mappings=all_mappings,
            settings=settings,
            client_name=client_name,
            process_name=process_name,
        )

        all_mappings = svc.finalize_mappings(
            all_mappings=all_mappings,
            settings=settings,
        )

        db_result = {"inserted": 0, "skipped": 0, "errors": 0}
        if save_to_db:
            try:
                db_result = upsert_mappings(
                    mappings=all_mappings,
                    master_id=master_id,
                    client_name=client_name,
                    process_name=process_name,
                    settings=settings,
                    skip_unmatched=skip_unmatched,
                )
                logger.info(
                    f"DB write complete — master_id={master_id} "
                    f"inserted={db_result['inserted']} "
                    f"skipped={db_result['skipped']} "
                    f"errors={db_result['errors']}"
                )
            except Exception as db_exc:
                logger.error(f"DB write failed (non-fatal, ZIP still returned): {db_exc}")

        # ── Build nested mapping + schema ──────────────────────────────────────
        mapping_list = [
            {
                "client_column": m.get("partner_field", ""),
                "lms_column":    m.get("json_key", ""),
                "entity":        m.get("entity", "OTHER"),
            }
            for m in all_mappings
        ]
        builder_input = {"mappings": mapping_list}
        nested_result = generate_nested_mapping(builder_input)
        schema_result = generate_schema(builder_input)

        # ── Excel output ───────────────────────────────────────────────────────
        safe_client_name  = _sanitize_path_component(client_name, "client")
        safe_process_name = _sanitize_path_component(process_name, "process")

        out_dir    = settings.output_path / safe_client_name
        out_dir.mkdir(parents=True, exist_ok=True)
        excel_path = out_dir / f"mapping_{safe_client_name}_{safe_process_name}.xlsx"

        svc.post_process_and_output(
            all_mappings=all_mappings,
            settings=settings,
            output_path=str(excel_path),
            client_name=client_name,
            process_name=process_name,
        )

        # ── Assemble ZIP ───────────────────────────────────────────────────────
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(excel_path, arcname=excel_path.name)
            zf.writestr(
                f"nested_mapping_{safe_client_name}_{safe_process_name}.json",
                json.dumps(nested_result["mappings"], indent=2, ensure_ascii=False),
            )
            zf.writestr(
                f"schema_{safe_client_name}_{safe_process_name}.json",
                json.dumps(schema_result["schema"], indent=2, ensure_ascii=False),
            )
            if include_build_references:
                ref_dir = Path(settings.references_dir)
                for fname in (
                    "field_dictionary.json",
                    "alias_registry.json",
                    "entity_routing.json",
                ):
                    fp = ref_dir / fname
                    if fp.is_file():
                        zf.write(fp, arcname=f"references/{fname}")

        zip_buffer.seek(0)
        background_tasks.add_task(_safe_unlink, tmp)
        tmp = None

        response_headers: Dict[str, str] = {
            "Content-Disposition": (
                f"attachment; filename={safe_client_name}_{safe_process_name}_outputs.zip"
            ),
            "X-DB-Inserted": str(db_result["inserted"]),
            "X-DB-Skipped":  str(db_result["skipped"]),
            "X-DB-Errors":   str(db_result["errors"]),
            "X-Master-Id":   str(master_id),
        }
        if ref_meta is not None:
            response_headers["X-References-PutM-Rows"] = str(ref_meta.get("putm_rows", ""))
            response_headers["X-References-Mapping-Rows"] = str(ref_meta.get("mapping_rows", ""))

        return Response(
            content=zip_buffer.getvalue(),
            media_type="application/zip",
            headers=response_headers,
        )

    except Exception as e:
        logger.exception("Full pipeline failed")
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        _safe_unlink(tmp)


# ── 5. Generate Nested Mapping ────────────────────────────────────────────────

@router.post(
    "/generate-nested-mapping",
    response_model=NestedMappingResponse,
    summary="Convert flat mappings → nested JSON tree (leaf = client_column)",
)
async def generate_nested_mapping_endpoint(
    request: LOSJsonRequest,
    db: Session = Depends(get_db),
):
    try:
        start         = time.time()
        builder_input = _request_to_builder_input(request)
        result        = generate_nested_mapping(builder_input)
        stats         = _count_stats(request)

        return NestedMappingResponse(
            client_name        = request.client_name,
            los_json           = result["mappings"],
            processing_time_ms = round((time.time() - start) * 1000, 2),
            **stats,
        )
    except Exception as e:
        logger.error(f"Error in /generate-nested-mapping: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


# ── 6. Generate Schema ────────────────────────────────────────────────────────

@router.post(
    "/generate-schema",
    response_model=SchemaResponse,
    summary="Generate blank LOS schema from mapped paths (leaf = null)",
)
async def generate_schema_endpoint(
    request: LOSJsonRequest,
    db: Session = Depends(get_db),
):
    try:
        start         = time.time()
        builder_input = _request_to_builder_input(request)
        result        = generate_schema(builder_input)
        stats         = _count_stats(request)

        return SchemaResponse(
            client_name        = request.client_name,
            los_schema         = result["schema"],
            processing_time_ms = round((time.time() - start) * 1000, 2),
            **stats,
        )
    except Exception as e:
        logger.error(f"Error in /generate-schema: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


# ── Health + status ───────────────────────────────────────────────────────────

@router.get("/health", summary="Health check")
async def health():
    return {"status": "ok"}


@router.get("/references/status", summary="Check reference files")
async def ref_status(settings: Settings = Depends(get_settings)):
    files  = ["field_dictionary.json", "alias_registry.json", "entity_routing.json"]
    status = {}
    for f in files:
        p = settings.refs_path / f
        status[f] = {
            "exists":  p.exists(),
            "size_kb": round(p.stat().st_size / 1024, 1) if p.exists() else None,
        }
    return {"ready": all(v["exists"] for v in status.values()), "files": status}