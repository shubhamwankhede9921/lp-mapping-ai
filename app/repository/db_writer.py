"""
repository/db_writer.py

Writes pipeline mapping results to the TARGET database.
Column names match the actual schema from DESCRIBE generic_excel_upload_definition_fields.

Pipeline field          →  DB column
─────────────────────────────────────────────────────
partner_field           →  excel_column_name   ← was wrongly 'excel_column'
matched_excel_key       →  table_column_name
json_key                →  partner_api_key
master_id (Form param)  →  master_id
entity                  →  ui_grouping         ← was wrongly 'ui_group'
"""

import logging
from datetime import datetime
from typing import Any, Dict, List

from sqlalchemy import create_engine, text
from sqlalchemy.orm import sessionmaker, Session

logger = logging.getLogger(__name__)

# Module-level engine cache — one engine per target_db_url for the process lifetime
_engine_cache: Dict[str, Any] = {}


def _get_target_session(settings) -> Session:
    url = settings.target_db_url
    if url not in _engine_cache:
        logger.info(
            f"[db_writer] Creating engine for target DB: "
            f"{settings.target_db_host}/{settings.target_db_name}"
        )
        engine = create_engine(url, pool_pre_ping=True, pool_recycle=3600)
        _engine_cache[url] = sessionmaker(bind=engine, autocommit=False, autoflush=False)

    return _engine_cache[url]()


def _get_table_name(settings) -> str:
    table = getattr(settings, "generic_mapping_table", None)
    if not table:
        raise ValueError(
            "GENERIC_MAPPING_TABLE is not set in .env. "
            "Add:  GENERIC_MAPPING_TABLE=generic_excel_upload_definition_fields"
        )
    return table


def _build_row(
    mapping: Dict[str, Any],
    master_id: int,
    client_name: str,
    process_name: str,
) -> Dict[str, Any]:
    """
    Map pipeline output keys → exact DB column names.
    Only columns we actually populate are included;
    all other columns (ui_type, ui_title, length, etc.) keep their DB defaults.
    """
    excel_col = (mapping.get("partner_field") or "").strip()
    inferred_type = "date" if "date" in excel_col.lower() else ""

    # partner_api_key rule:
    # - if column_category present (and not a generic label like API/APPLICANT), use:
    #     loanAccounts.<column_category>.<partner_field>
    # - else:
    #     loanAccounts.<partner_field>
    raw_cat = (mapping.get("column_category") or mapping.get("columnCategory") or "").strip()
    cat_norm = raw_cat.upper()
    invalid_cats = {
        "API",
        "APPLICANT",
        "API FIELD MAPPING",
        "LOAN API WITH BRE CHECK",
        "MIGRATION API MAPPING",
        "ONLY COAPPLICANT",
    }
    if raw_cat and cat_norm not in invalid_cats:
        partner_api_key = f"loanAccounts.{raw_cat}.{excel_col}"
    else:
        partner_api_key = f"loanAccounts.{excel_col}" if excel_col else ""

    # Per requirement:
    # - excel_column_name/table_column_name are written
    # - partner_api_key uses the partner dotted key (see rule above)
    # - ui_key stores the JSON key used in partner API
    # - type is "date" when excel_column_name contains "date", else empty
    return {
        # ── core mapping columns ───────────────────────────────────────
        "excel_column_name": excel_col,
        "table_column_name": (mapping.get("matched_excel_key") or "").strip(),
        "partner_api_key":   partner_api_key,
        "ui_key":            (mapping.get("json_key") or "").strip(),
        "master_id":         master_id,
        "ui_grouping":       (mapping.get("entity") or "OTHER").strip(),

        # ── audit / traceability columns ──────────────────────────────
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


# ── MySQL upsert ───────────────────────────────────────────────────────────────
#
# ON DUPLICATE KEY UPDATE requires a UNIQUE index on (master_id, excel_column_name).
# If that index doesn't exist yet, run:
#   ALTER TABLE `generic_excel_upload_definition_fields`
#     ADD UNIQUE KEY `uq_master_excel` (`master_id`, `excel_column_name`);

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


# ── public API ─────────────────────────────────────────────────────────────────

def upsert_mappings(
    mappings: List[Dict[str, Any]],
    master_id: int,
    client_name: str,
    process_name: str,
    settings,
    skip_unmatched: bool = False,
) -> Dict[str, int]:
    """
    Insert or update mapping rows in the TARGET MySQL table.
    Opens its own connection to settings.target_db_url.

    Parameters
    ----------
    mappings        Flat mapping dicts from the pipeline
    master_id       Integer stored in master_id column
    client_name     Written to created_by / description
    process_name    Written to description
    settings        App settings (reads TARGET_DB_* + GENERIC_MAPPING_TABLE)
    skip_unmatched  If True, rows with no matched_excel_key are skipped

    Returns
    -------
    {"inserted": N, "skipped": N, "errors": N}
    """
    table    = _get_table_name(settings)
    stmt     = text(_UPSERT_SQL.format(table=table))
    inserted = skipped = errors = 0

    db = _get_target_session(settings)
    try:
        for m in mappings:
            # skip rows that never got a PUTM match (if caller asked)
            if skip_unmatched and not m.get("matched_excel_key"):
                skipped += 1
                continue

            # skip rows with no partner field at all
            if not (m.get("partner_field") or "").strip():
                skipped += 1
                continue

            row = _build_row(m, master_id, client_name, process_name)
            try:
                db.execute(stmt, row)
                inserted += 1
            except Exception as exc:
                logger.error(
                    f"DB write failed for partner_field={m.get('partner_field')!r}: {exc}"
                )
                errors += 1

        db.commit()

    except Exception as exc:
        logger.error(f"DB commit failed: {exc}")
        db.rollback()
        raise
    finally:
        db.close()

    logger.info(
        f"[db_writer] target={settings.target_db_host}/{settings.target_db_name} "
        f"table=`{table}` master_id={master_id} "
        f"inserted/updated={inserted} skipped={skipped} errors={errors}"
    )
    return {"inserted": inserted, "skipped": skipped, "errors": errors}


def delete_by_master_id(master_id: int, settings) -> int:
    """
    Delete all rows for a master_id from the TARGET DB.
    Use before re-running if you prefer delete+insert over upsert.
    """
    table = _get_table_name(settings)
    db    = _get_target_session(settings)
    try:
        result  = db.execute(
            text(f"DELETE FROM `{table}` WHERE master_id = :mid"),
            {"mid": master_id},
        )
        db.commit()
        deleted = result.rowcount
    finally:
        db.close()

    logger.info(f"[db_writer] Deleted {deleted} rows for master_id={master_id} from `{table}`")
    return deleted