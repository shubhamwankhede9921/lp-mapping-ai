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
    Body, Query,
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
    EditSessionApproveRequest,
    EditSessionCreateRequest,
    EditSessionResponse,
    EditSessionSummary,
    EditSessionUpdateRequest,
)
from app.repository import database as db_repo
from app.repository.db_writer import upsert_mappings   # opens its own target-DB connection
from app.services import mapping_service as svc
from app.services import edit_session_store as session_store

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

def _db_insert_script_template() -> str:
    # NOTE: This is intentionally self-contained so it can run from the extracted ZIP
    # without needing the whole FastAPI project on PYTHONPATH.
    return r'''#!/usr/bin/env python
"""
Insert LP Mapping ZIP output into the target DB.

This script is shipped inside the ZIP returned by the FastAPI `/mapping/full-pipeline` endpoint.

What it does
------------
- Reads `flat_mappings_*.json` produced by the pipeline.
- Upserts rows into the TARGET DB table (default: `generic_excel_upload_definition_fields`).

Pipeline field -> DB column
---------------------------
- partner_field -> excel_column_name
- matched_excel_key -> table_column_name
- json_key -> ui_key
- master_id -> master_id
- entity -> ui_grouping

Required environment variables
------------------------------
- TARGET_DB_HOST
- TARGET_DB_NAME
- TARGET_DB_USER
- TARGET_DB_PASSWORD

Optional environment variables
------------------------------
- TARGET_DB_PORT (default: 3306)
- GENERIC_MAPPING_TABLE (default: generic_excel_upload_definition_fields)

Example
-------
python insert_output_to_db.py --input flat_mappings_client_process.json --master-id 123 --client "ACME" --process "COMBINED"
"""

from __future__ import annotations

import argparse
import json
import os
from datetime import datetime
from typing import Any, Dict, List

from sqlalchemy import create_engine, text
from sqlalchemy.orm import sessionmaker


def _env(name: str, default: str | None = None) -> str | None:
    value = os.getenv(name)
    if value is None or value == "":
        return default
    return value


def _get_target_db_url() -> str:
    host = _env("TARGET_DB_HOST")
    name = _env("TARGET_DB_NAME") or _env("TARGET_DB_DATABASE")
    user = _env("TARGET_DB_USER") or _env("TARGET_DB_USERNAME")
    password = _env("TARGET_DB_PASSWORD") or _env("TARGET_DB_PASS")
    port = int(_env("TARGET_DB_PORT", "3306") or "3306")

    missing = [k for k, v in {
        "TARGET_DB_HOST": host,
        "TARGET_DB_NAME": name,
        "TARGET_DB_USER": user,
        "TARGET_DB_PASSWORD": password,
    }.items() if not v]
    if missing:
        raise SystemExit(f"Missing required env vars: {', '.join(missing)}")

    return f"mysql+pymysql://{user}:{password}@{host}:{port}/{name}"


def _get_table_name() -> str:
    return _env("GENERIC_MAPPING_TABLE", "generic_excel_upload_definition_fields") or "generic_excel_upload_definition_fields"


def _build_row(mapping: Dict[str, Any], master_id: int, client_name: str, process_name: str) -> Dict[str, Any]:
    excel_col = (mapping.get("partner_field") or "").strip()
    inferred_type = "date" if "date" in excel_col.lower() else ""
    return {
        "excel_column_name": excel_col,
        "table_column_name": (mapping.get("matched_excel_key") or "").strip(),
        "partner_api_key":   "",
        "ui_key":            (mapping.get("json_key") or "").strip(),
        "master_id":         master_id,
        "ui_grouping":       (mapping.get("entity") or "OTHER").strip(),
        "type":              inferred_type,
        "description":       (
            f"[auto] client={client_name} process={process_name} "
            f"confidence={round(float(mapping.get('confidence') or 0.0), 4)} "
            f"needs_review={bool(mapping.get('needs_review', False))}"
        ),
        "created_at":        datetime.utcnow(),
        "updated_at":        datetime.utcnow(),
        "created_by":        f"llm_mapping_pipeline:{client_name}",
    }


_UPSERT_SQL = """
    INSERT INTO `{table}`
        (excel_column_name, table_column_name, partner_api_key,
         master_id, ui_grouping,
         type, description,
         created_at, updated_at, created_by,
         ui_key)
    VALUES
        (:excel_column_name, :table_column_name, :partner_api_key,
         :master_id, :ui_grouping,
         :type, :description,
         :created_at, :updated_at, :created_by,
         :ui_key)
    ON DUPLICATE KEY UPDATE
        table_column_name = VALUES(table_column_name),
        ui_key            = VALUES(ui_key),
        partner_api_key   = VALUES(partner_api_key),
        ui_grouping       = VALUES(ui_grouping),
        type              = VALUES(type),
        description       = VALUES(description),
        updated_at        = VALUES(updated_at),
        created_by        = VALUES(created_by)
"""


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, help="Path to flat_mappings_*.json from the ZIP")
    ap.add_argument("--master-id", type=int, required=True, help="Value to store in master_id column")
    ap.add_argument("--client", required=True, help="Client name (for audit columns)")
    ap.add_argument("--process", default="COMBINED", help="Process name (for audit columns)")
    ap.add_argument("--skip-unmatched", action="store_true", help="Skip rows that have no matched_excel_key")
    args = ap.parse_args()

    with open(args.input, "r", encoding="utf-8") as f:
        mappings: List[Dict[str, Any]] = json.load(f)

    table = _get_table_name()
    engine = create_engine(_get_target_db_url(), pool_pre_ping=True, pool_recycle=3600)
    SessionLocal = sessionmaker(bind=engine, autocommit=False, autoflush=False)
    db = SessionLocal()
    try:
        stmt = text(_UPSERT_SQL.format(table=table))
        inserted = skipped = errors = 0

        for m in mappings:
            if args.skip_unmatched and not m.get("matched_excel_key"):
                skipped += 1
                continue
            if not (m.get("partner_field") or "").strip():
                skipped += 1
                continue

            row = _build_row(m, args.master_id, args.client, args.process)
            try:
                db.execute(stmt, row)
                inserted += 1
            except Exception as exc:
                errors += 1
                print(f"DB write failed for partner_field={m.get('partner_field')!r}: {exc}")

        db.commit()
        print(f"Done. inserted/updated={inserted} skipped={skipped} errors={errors} table={table}")
        return 0 if errors == 0 else 2
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()


if __name__ == "__main__":
    raise SystemExit(main())
'''


def _run_build_references_from_db(
    settings: Settings,
    putm_table_override: Optional[str] = None,
    mapping_table_override: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Fetch PUTM + generic mapping from the source DB (no dumps),
    build reference dictionaries, and return them + row counts.
    """
    from app.repository.database import fetch_generic_mapping_dataframe, fetch_putm_dataframe
    from app.scripts.build_references import (
        AliasRegistryBuilder,
        EntityRoutingBuilder,
        FieldDictionaryBuilder,
    )

    putm_df = fetch_putm_dataframe(settings, putm_table_override)
    generic_df = fetch_generic_mapping_dataframe(settings, mapping_table_override)

    fd = FieldDictionaryBuilder(putm_df).build()
    ar = AliasRegistryBuilder(generic_df).build()
    er = EntityRoutingBuilder(generic_df).build()

    return {
        "putm_rows": len(putm_df),
        "mapping_rows": len(generic_df),
        "field_dictionary": fd,
        "alias_registry": ar,
        "entity_routing": er,
    }


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


def _save_json_payload(payload: Any) -> str:
    """
    Persist a JSON payload to a temporary file so the existing pipeline
    (`input_parser.parse_input`) can consume it as a `.json` input.
    """
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".json")
    try:
        tmp.write(json.dumps(payload, ensure_ascii=False).encode("utf-8"))
        tmp.flush()
        return tmp.name
    finally:
        tmp.close()


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

        svc.assign_sequential_loanparameter_slots(det)

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
    save_to_db: bool     = Form(False, description="Write results to GENERIC_MAPPING_TABLE in TARGET DB"),
    skip_unmatched: bool = Form(False, description="Skip rows with no matched_excel_key when writing to DB"),
    include_build_references: bool = Form(
        False,
        description=(
            "Fetch from source DB and rebuild field_dictionary / alias_registry / entity_routing "
            "before mapping, and add those JSON files under references/ in the ZIP (no dumps written)"
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
            ref_meta = _run_build_references_from_db(
                settings,
                putm_table_override=None,
                mapping_table_override=None,
            )

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
            # flat mappings for DB/script usage (same dicts used by db_writer.upsert_mappings)
            zf.writestr(
                f"flat_mappings_{safe_client_name}_{safe_process_name}.json",
                json.dumps(all_mappings, indent=2, ensure_ascii=False),
            )
            zf.writestr(
                f"nested_mapping_{safe_client_name}_{safe_process_name}.json",
                json.dumps(nested_result["mappings"], indent=2, ensure_ascii=False),
            )
            zf.writestr(
                f"schema_{safe_client_name}_{safe_process_name}.json",
                json.dumps(schema_result["schema"], indent=2, ensure_ascii=False),
            )
            # script to insert output into DB after download
            zf.writestr("insert_output_to_db.py", _db_insert_script_template())
            if include_build_references and ref_meta:
                zf.writestr(
                    "references/field_dictionary.json",
                    json.dumps(ref_meta["field_dictionary"], indent=2, ensure_ascii=False),
                )
                zf.writestr(
                    "references/alias_registry.json",
                    json.dumps(ref_meta["alias_registry"], indent=2, ensure_ascii=False),
                )
                zf.writestr(
                    "references/entity_routing.json",
                    json.dumps(ref_meta["entity_routing"], indent=2, ensure_ascii=False),
                )

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

    except HTTPException:
        raise
    except Exception as e:
        logger.exception("Full pipeline failed")
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        _safe_unlink(tmp)


@router.post(
    "/mapping/full-pipeline-json",
    summary="Full pipeline from JSON body (no Excel upload) → ZIP outputs",
)
async def full_pipeline_json(
    background_tasks: BackgroundTasks,
    payload: Dict[str, Any] = Body(..., description="Raw JSON input to be flattened into fields"),
    client_name: str = Query(...),
    process_name: str = Query("COMBINED"),
    use_fuzzy: bool = Query(True),
    use_embeddings: bool = Query(False),
    use_llm: bool = Query(True),
    use_loanparameter_refinement: bool = Query(True),
    use_llm_entity_classifier: bool = Query(False),
    master_id: int = Query(..., description="FK stored in master_id column of the mapping table"),
    save_to_db: bool = Query(False, description="Write results to GENERIC_MAPPING_TABLE in TARGET DB"),
    skip_unmatched: bool = Query(False, description="Skip rows with no matched_excel_key when writing to DB"),
    include_build_references: bool = Query(False),
    settings: Settings = Depends(get_settings),
):
    tmp = None
    ref_meta: Optional[Dict[str, Any]] = None
    try:
        tmp = _save_json_payload(payload)

        if include_build_references:
            ref_meta = _run_build_references_from_db(
                settings,
                putm_table_override=None,
                mapping_table_override=None,
            )

        # Phase 1: deterministic (JSON input is flattened by input_parser.parse_json_input)
        p1 = svc.run_deterministic(
            input_file=tmp,
            settings=settings,
            process_name=process_name,
            sheet_filter=None,
            client_name=client_name,
            use_llm_entity_classifier=use_llm_entity_classifier,
        )

        det_results = p1.get("deterministic_results") or []
        unmatched = p1["unmatched_fields"]
        ep = p1["entity_prompts"]
        fd = p1["field_dictionary"]
        ar = p1["alias_registry"]

        if use_loanparameter_refinement:
            det_results = svc.refine_loanparameter_after_deterministic(
                deterministic_results=det_results,
                field_dictionary=fd,
                alias_registry=ar,
                settings=settings,
                client_name=client_name,
                process_name=process_name,
            ) or det_results

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
        all_mappings = svc.finalize_mappings(all_mappings=all_mappings, settings=settings)

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
            except Exception as db_exc:
                logger.error(f"DB write failed (non-fatal, ZIP still returned): {db_exc}")

        mapping_list = [
            {"client_column": m.get("partner_field", ""), "lms_column": m.get("json_key", ""), "entity": m.get("entity", "OTHER")}
            for m in all_mappings
        ]
        builder_input = {"mappings": mapping_list}
        nested_result = generate_nested_mapping(builder_input)
        schema_result = generate_schema(builder_input)

        safe_client_name = _sanitize_path_component(client_name, "client")
        safe_process_name = _sanitize_path_component(process_name, "process")

        out_dir = settings.output_path / safe_client_name
        out_dir.mkdir(parents=True, exist_ok=True)
        excel_path = out_dir / f"mapping_{safe_client_name}_{safe_process_name}.xlsx"

        svc.post_process_and_output(
            all_mappings=all_mappings,
            settings=settings,
            output_path=str(excel_path),
            client_name=client_name,
            process_name=process_name,
        )

        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(excel_path, arcname=excel_path.name)
            zf.writestr(
                f"flat_mappings_{safe_client_name}_{safe_process_name}.json",
                json.dumps(all_mappings, indent=2, ensure_ascii=False),
            )
            zf.writestr(
                f"nested_mapping_{safe_client_name}_{safe_process_name}.json",
                json.dumps(nested_result["mappings"], indent=2, ensure_ascii=False),
            )
            zf.writestr(
                f"schema_{safe_client_name}_{safe_process_name}.json",
                json.dumps(schema_result["schema"], indent=2, ensure_ascii=False),
            )
            zf.writestr("insert_output_to_db.py", _db_insert_script_template())
            if include_build_references and ref_meta:
                zf.writestr("references/field_dictionary.json", json.dumps(ref_meta["field_dictionary"], indent=2, ensure_ascii=False))
                zf.writestr("references/alias_registry.json", json.dumps(ref_meta["alias_registry"], indent=2, ensure_ascii=False))
                zf.writestr("references/entity_routing.json", json.dumps(ref_meta["entity_routing"], indent=2, ensure_ascii=False))

        zip_buffer.seek(0)
        background_tasks.add_task(_safe_unlink, tmp)
        tmp = None

        return Response(
            content=zip_buffer.getvalue(),
            media_type="application/zip",
            headers={
                "Content-Disposition": f"attachment; filename={safe_client_name}_{safe_process_name}_outputs.zip",
                "X-DB-Inserted": str(db_result["inserted"]),
                "X-DB-Skipped": str(db_result["skipped"]),
                "X-DB-Errors": str(db_result["errors"]),
                "X-Master-Id": str(master_id),
            },
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.exception("Full pipeline (json) failed")
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


@router.get("/references/putm-keys", summary="List available excel_key ↔ json_key from PUTM (via field_dictionary.json)")
async def list_putm_keys(
    process_name: Optional[str] = None,
    q: Optional[str] = None,
    limit: int = 2000,
    settings: Settings = Depends(get_settings),
) -> Dict[str, Any]:
    """
    Returns key pairs derived from PUTM (as persisted in references/field_dictionary.json).

    Query params:
      - process_name: filter by row primary process_name (case-insensitive); use 'ALL' or omit for no filter
      - q: substring filter across excel_key/json_key/role/description/example
      - limit: max rows returned (hard-capped)
    """
    fp = settings.refs_path / "field_dictionary.json"
    if not fp.exists():
        raise HTTPException(status_code=404, detail="field_dictionary.json not found. Run /references/build first.")

    try:
        raw = json.loads(fp.read_text(encoding="utf-8"))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to read field_dictionary.json: {e}")

    by_excel_key = raw.get("by_excel_key") or {}
    if not isinstance(by_excel_key, dict):
        raise HTTPException(status_code=500, detail="field_dictionary.json is invalid (by_excel_key missing).")

    pn = (process_name or "").strip().upper()
    if pn in ("", "ALL", "*"):
        pn = ""

    qq = (q or "").strip().lower()
    hard_cap = 50000
    limit = max(1, min(int(limit or 2000), hard_cap))

    def _row_matches_process_filter(
        filter_pn: str,
        primary: Optional[str],
        names_norm: list[str],
    ) -> bool:
        """Match the catalog row's primary process only (PUTM rows are per-process in practice)."""
        if not filter_pn:
            return True
        pr = (primary or "").strip().upper()
        if pr == filter_pn:
            return True
        clean_names = [n for n in names_norm if n]
        if not pr and len(clean_names) == 1 and clean_names[0] == filter_pn:
            return True
        return False

    rows = []
    processes = set()
    matched_total = 0
    for excel_key, meta in by_excel_key.items():
        if not isinstance(meta, dict):
            continue
        ek = str(excel_key or "").strip()
        jk = str(meta.get("json_key") or "").strip()
        role = str(meta.get("role") or "").strip().upper() or None
        primary_process = str(meta.get("process_name") or "").strip().upper() or None
        desc = meta.get("description")
        example = meta.get("example")
        process_names = meta.get("process_names") if isinstance(meta.get("process_names"), list) else []
        process_names_norm = [str(x).strip().upper() for x in process_names if str(x).strip()]

        if not ek or not jk:
            continue

        if pn and not _row_matches_process_filter(pn, primary_process, process_names_norm):
            continue

        # text search filter
        if qq:
            hay = " ".join(
                [
                    ek,
                    jk,
                    role or "",
                    primary_process or "",
                    " ".join(process_names_norm),
                    str(desc or ""),
                    str(example or ""),
                ]
            ).lower()
            if qq not in hay:
                continue

        matched_total += 1

        if primary_process:
            processes.add(primary_process)
        for p in process_names_norm:
            processes.add(p)

        if len(rows) < limit:
            rows.append(
                {
                    "excel_key": ek,
                    "json_key": jk,
                    "role": role,
                    "process_name": primary_process,
                    "process_names": process_names_norm,
                    "description": desc,
                    "example": example,
                }
            )

    rows.sort(key=lambda r: (r.get("process_name") or "", r.get("excel_key") or ""))
    return {
        "total": len(rows),
        "matched_total": matched_total,
        "truncated": matched_total > len(rows),
        "limit": limit,
        "processes": sorted([p for p in processes if p]),
        "rows": rows,
    }


# ── 7. Edit sessions (draft → approve → DB write) ─────────────────────────────

@router.get(
    "/edit-sessions",
    response_model=list[EditSessionSummary],
    summary="List editable mapping sessions",
)
async def list_edit_sessions(settings: Settings = Depends(get_settings)):
    return session_store.list_sessions(output_path=settings.output_path)


@router.post(
    "/edit-sessions",
    response_model=EditSessionResponse,
    summary="Create an editable mapping session (draft)",
)
async def create_edit_session(
    request: EditSessionCreateRequest,
    settings: Settings = Depends(get_settings),
):
    data = session_store.create_session(
        output_path=settings.output_path,
        client_name=request.client_name,
        process_name=request.process_name,
        master_id=request.master_id,
        mappings=request.mappings,
        created_by=request.created_by,
    )
    return EditSessionResponse(**data)


@router.get(
    "/edit-sessions/{session_id}",
    response_model=EditSessionResponse,
    summary="Get an editable mapping session",
)
async def get_edit_session(session_id: str, settings: Settings = Depends(get_settings)):
    try:
        data = session_store.get_session(output_path=settings.output_path, session_id=session_id)
        return EditSessionResponse(**data)
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="Session not found")


@router.put(
    "/edit-sessions/{session_id}",
    response_model=EditSessionResponse,
    summary="Update an editable mapping session (draft only)",
)
async def update_edit_session(
    session_id: str,
    request: EditSessionUpdateRequest,
    settings: Settings = Depends(get_settings),
):
    try:
        data = session_store.update_session(
            output_path=settings.output_path,
            session_id=session_id,
            mappings=request.mappings,
            client_name=request.client_name,
            process_name=request.process_name,
            master_id=request.master_id,
            updated_by=request.updated_by,
            note=request.note,
        )
        return EditSessionResponse(**data)
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="Session not found")
    except PermissionError as e:
        raise HTTPException(status_code=409, detail=str(e))


@router.delete(
    "/edit-sessions/{session_id}",
    summary="Delete an editable mapping session",
)
async def delete_edit_session(session_id: str, settings: Settings = Depends(get_settings)):
    try:
        # allow deleting any session (draft/approved) - simple cleanup
        session_store.delete_session(output_path=settings.output_path, session_id=session_id)
        return {"status": "deleted", "session_id": session_id}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post(
    "/edit-sessions/{session_id}/approve",
    response_model=EditSessionResponse,
    summary="Approve edited results and write to TARGET DB",
)
async def approve_edit_session(
    session_id: str,
    request: EditSessionApproveRequest,
    settings: Settings = Depends(get_settings),
):
    try:
        session = session_store.get_session(output_path=settings.output_path, session_id=session_id)
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="Session not found")

    if session.get("status") != "draft":
        raise HTTPException(status_code=409, detail="Only draft sessions can be approved")

    master_id = request.master_id if request.master_id is not None else session.get("master_id")
    if master_id is None:
        raise HTTPException(status_code=400, detail="master_id is required to approve and write to DB")

    mappings = session.get("mappings") or []
    if not isinstance(mappings, list):
        raise HTTPException(status_code=400, detail="Session mappings are invalid (expected list)")

    # Write to target DB using the edited mappings as-is.
    try:
        db_result = upsert_mappings(
            mappings=mappings,
            master_id=int(master_id),
            client_name=session.get("client_name") or "",
            process_name=session.get("process_name") or "",
            settings=settings,
            skip_unmatched=bool(request.skip_unmatched),
        )
    except Exception as e:
        logger.exception("Approval DB write failed")
        raise HTTPException(status_code=500, detail=f"DB write failed: {e}")

    approved = session_store.approve_session(
        output_path=settings.output_path,
        session_id=session_id,
        approved_by=request.approved_by,
        approval_result={"db_result": db_result, "master_id": int(master_id)},
    )
    return EditSessionResponse(**approved)