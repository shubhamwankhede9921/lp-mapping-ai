"""
app/repository/database.py
 
SQLAlchemy-based DB access layer.
Provides:
  - get_engine()                          → SQLAlchemy engine (singleton)
  - get_db()                              → FastAPI session dependency
  - fetch_putm_dataframe / fetch_generic_mapping_dataframe → read from MySQL (no files)
"""
 
from __future__ import annotations
 
import logging
from functools import lru_cache
from pathlib import Path
from typing import Generator, Optional
 
import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.orm import Session, sessionmaker
 
logger = logging.getLogger(__name__)
 
# ── Engine singleton ──────────────────────────────────────────────────────────
 
@lru_cache(maxsize=4)
def _engine_for_url(url: str):
    """Cache engines by DB URL (supports multiple Settings instances)."""
    return create_engine(
        url,
        pool_pre_ping=True,
        pool_recycle=3600,
        connect_args={"connect_timeout": 30},
    )


def get_engine(settings=None):
    """
    Return a SQLAlchemy engine for the configured SOURCE DB.

    Prefer passing the request-scoped `settings` object so the engine always reflects
    the env/.env values loaded by the FastAPI dependency.
    """
    if settings is None:
        from app.config import get_settings

        settings = get_settings()

    url = settings.source_db_url
    logger.info(
        "Creating SQLAlchemy engine → %s:%s/%s (user=%s)",
        settings.db_host,
        settings.db_port or 3306,
        settings.db_name,
        settings.db_user,
    )
    return _engine_for_url(url)
 
 
# ── Session factory + FastAPI dependency ──────────────────────────────────────
 
@lru_cache(maxsize=4)
def _get_session_factory_for_url(url: str):
    return sessionmaker(bind=_engine_for_url(url), autocommit=False, autoflush=False)
 
 
def get_db() -> Generator[Session, None, None]:
    """
    FastAPI dependency.  Usage:
        def my_endpoint(db: Session = Depends(get_db)): ...
    """
    from app.config import get_settings

    settings = get_settings()
    SessionLocal = _get_session_factory_for_url(settings.source_db_url)
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
 
 
# ── SQL definitions ───────────────────────────────────────────────────────────
 
_PUTM_QUERY = text("""
    SELECT
        process_name,
        excel_key,
        json_key,
        description,
        example,
        json_key_role
    FROM financialForms.putm_upload_api_excel_json_mapping
    WHERE process_name IN ('Origination', 'Enrollment')
""")
 
_GENERIC_QUERY = text("""
    SELECT
        geudf.*
    FROM
        financialForms.generic_excel_upload_definition_fields geudf
    INNER JOIN
        financialForms.generic_excel_upload_definition geud
            ON geudf.master_id = geud.id
    WHERE
        geud.upload_name = :upload_name
        AND geudf.is_exclude = 0
""")
 
 
# ── Public fetch helpers (no intermediate dump files) ─────────────────────────
 
def fetch_putm_dataframe(
    settings,
    table_override: Optional[str] = None,
):
    """
    Load PUTM mapping rows from the source MySQL database.
 
    Args:
        settings:        Reserved for future use (e.g. alternate DB URLs).
        table_override:  Full table reference (e.g. schema.table) to query instead of default.
    """
    query = _PUTM_QUERY
    if table_override:
        logger.info(f"PUTM table override: {table_override}")
        query = text(
            f"""
            SELECT
                process_name,
                excel_key,
                json_key,
                description,
                example,
                json_key_role
            FROM {table_override}
            WHERE process_name IN ('Origination', 'Enrollment')
            """
        )
 
    logger.info("Fetching PUTM data from MySQL via SQLAlchemy …")
    with get_engine(settings).connect() as conn:
        df = pd.read_sql(query, conn)
 
    logger.info(f"  → {len(df)} PUTM rows  (columns: {list(df.columns)})")
    return df
 
 
def fetch_generic_mapping_dataframe(
    settings,
    table_override: Optional[str] = None,
):
    """
    Load generic excel definition field rows from the source MySQL database.
    Only active fields (is_exclude = 0) for 'Individual Loan Upload v3' are returned.
    """
    upload_name = getattr(settings, "generic_upload_name", None) or "Individual Loan Upload v3"
    params = {"upload_name": upload_name}
    query = _GENERIC_QUERY
    if table_override:
        logger.info(f"Generic-mapping table override: {table_override}")
        query = text(f"""
            SELECT
                geudf.*
            FROM
                {table_override} geudf
            INNER JOIN
                financialForms.generic_excel_upload_definition geud
                    ON geudf.master_id = geud.id
            WHERE
                geud.upload_name = :upload_name
                AND geudf.is_exclude = 0
        """)
 
    logger.info("Fetching generic-excel-upload data from MySQL via SQLAlchemy …")
    with get_engine(settings).connect() as conn:
        df = pd.read_sql(query, conn, params=params)
 
    logger.info(f"  → {len(df)} generic-mapping rows  (columns: {list(df.columns)})")
    return df
 
 
# ── Dumps disabled ────────────────────────────────────────────────────────────
 
def extract_putm_dump(
    settings,
    output_path: str,
    table_override: Optional[str] = None,
) -> int:
    raise RuntimeError(
        "Dumps are disabled. Use fetch_putm_dataframe(...) to read live data from DB."
    )
 
 
def extract_generic_mapping(
    settings,
    output_path: str,
    table_override: Optional[str] = None,
) -> int:
    raise RuntimeError(
        "Dumps are disabled. Use fetch_generic_mapping_dataframe(...) to read live data from DB."
    )
 
 
# ── Backward-compatible aliases ───────────────────────────────────────────────
 
def fetch_putm_df() -> pd.DataFrame:
    """Return raw PUTM DataFrame (default table, no override)."""
    return fetch_putm_dataframe(None)
 
 
def fetch_generic_df() -> pd.DataFrame:
    """Return raw generic-excel-upload DataFrame (default join, no override)."""
    return fetch_generic_mapping_dataframe(None)
 