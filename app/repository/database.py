"""
app/repository/database.py

SQLAlchemy-based DB access layer.
Provides:
  - get_engine()                          → SQLAlchemy engine (singleton)
  - get_db()                              → FastAPI session dependency
  - extract_putm_dump(settings, path)     → writes putm rows to xlsx, returns row count
  - extract_generic_mapping(settings, path) → writes mapping rows to csv, returns row count
"""

from __future__ import annotations

import logging
import os
from functools import lru_cache
from pathlib import Path
from typing import Generator, Optional

import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.engine import URL
from sqlalchemy.orm import Session, sessionmaker
from dotenv import load_dotenv
load_dotenv()  # Must be called before get_engine()
logger = logging.getLogger(__name__)

# ── Engine singleton ──────────────────────────────────────────────────────────

@lru_cache(maxsize=1)
def get_engine():
    """
    Build and cache one SQLAlchemy engine for the lifetime of the process.
    Reads credentials from environment variables (already loaded from .env by
    app/config.py or the top of main.py via python-dotenv).
    """
    host     = os.getenv("DB_HOST")
    port     = os.getenv("DB_PORT", "3306")
    database = os.getenv("DB_NAME") or os.getenv("DB_DATABASE")
    user     = os.getenv("DB_USER")
    password = os.getenv("DB_PASSWORD")

    missing = [
        k for k, v in {
            "DB_HOST": host,
            "DB_NAME / DB_DATABASE": database,
            "DB_USER": user,
            "DB_PASSWORD": password,
        }.items()
        if not v
    ]
    if missing:
        raise EnvironmentError(
            f"Missing required environment variables: {missing}. "
            "Add them to your .env file."
        )

    # Keep the source DB driver aligned with requirements.txt and app.config.Settings.
    url = URL.create(
        "mysql+pymysql",
        username=user,
        password=password,
        host=host,
        port=int(port),
        database=database,
    )
    logger.info(f"Creating SQLAlchemy engine → {host}:{port}/{database} (user={user})")

    engine = create_engine(
        url,
        pool_pre_ping=True,       # auto-reconnect on stale connections
        pool_recycle=3600,        # recycle connections every hour
        connect_args={"connect_timeout": 30},
    )
    return engine


# ── Session factory + FastAPI dependency ──────────────────────────────────────

@lru_cache(maxsize=1)
def _get_session_factory():
    return sessionmaker(bind=get_engine(), autocommit=False, autoflush=False)


def get_db() -> Generator[Session, None, None]:
    """
    FastAPI dependency.  Usage:
        def my_endpoint(db: Session = Depends(get_db)): ...
    """
    SessionLocal = _get_session_factory()
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


# ── SQL definitions ───────────────────────────────────────────────────────────

_PUTM_QUERY = text("""
    SELECT *
    FROM financialForms.putm_upload_api_excel_json_mapping
    WHERE process_name IN ('Origination', 'Enrollment')
""")

_GENERIC_QUERY = text("""
    SELECT
        geudf.*,
        geud.upload_name,
        geud.id AS upload_definition_id
    FROM
        financialForms.generic_excel_upload_definition_fields geudf
    INNER JOIN
        financialForms.generic_excel_upload_definition geud
            ON geud.id = geudf.master_id
            AND geud.upload_name = 'Individual Loan Upload v3'
""")


# ── Public extraction helpers ─────────────────────────────────────────────────

def extract_putm_dump(
    settings,
    output_path: str,
    table_override: Optional[str] = None,
) -> int:
    """
    Fetch PUTM rows from MySQL and write them to *output_path* as an Excel file.

    Args:
        settings:        app Settings instance (used only for logging context).
        output_path:     Destination .xlsx path (parent dirs created if needed).
        table_override:  If supplied, replaces the default table name in the query.

    Returns:
        Number of rows written.
    """
    query = _PUTM_QUERY
    if table_override:
        logger.info(f"PUTM table override: {table_override}")
        query = text(
            f"SELECT * FROM {table_override} WHERE process_name IN ('Origination', 'Enrollment')"
        )

    logger.info("Fetching PUTM data from MySQL via SQLAlchemy …")
    with get_engine().connect() as conn:
        df = pd.read_sql(query, conn)

    logger.info(f"  → {len(df)} PUTM rows fetched  (columns: {list(df.columns)})")

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False)
    logger.info(f"  → PUTM dump written to {output_path}")
    return len(df)


def extract_generic_mapping(
    settings,
    output_path: str,
    table_override: Optional[str] = None,
) -> int:
    """
    Fetch generic-excel-upload rows from MySQL and write them to *output_path*
    as a CSV file.

    Args:
        settings:        app Settings instance (used only for logging context).
        output_path:     Destination .csv path (parent dirs created if needed).
        table_override:  If supplied, substitutes the fields table name in the query.

    Returns:
        Number of rows written.
    """
    query = _GENERIC_QUERY
    if table_override:
        logger.info(f"Generic-mapping table override: {table_override}")
        query = text(
            f"""
            SELECT
                geudf.*,
                geud.upload_name,
                geud.id AS upload_definition_id
            FROM
                {table_override} geudf
            INNER JOIN
                financialForms.generic_excel_upload_definition geud
                    ON geud.id = geudf.master_id
                    AND geud.upload_name = 'Individual Loan Upload v3'
            """
        )

    logger.info("Fetching generic-excel-upload data from MySQL via SQLAlchemy …")
    with get_engine().connect() as conn:
        df = pd.read_sql(query, conn)

    logger.info(f"  → {len(df)} generic-mapping rows fetched  (columns: {list(df.columns)})")

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(output_path, index=False, encoding="utf-8")
    logger.info(f"  → Generic mapping written to {output_path}")
    return len(df)


# ── Convenience: fetch both as DataFrames (used by build_references pipeline) ──

def fetch_putm_df() -> pd.DataFrame:
    """Return raw PUTM DataFrame without writing any file."""
    with get_engine().connect() as conn:
        return pd.read_sql(_PUTM_QUERY, conn)


def fetch_generic_df() -> pd.DataFrame:
    """Return raw generic-excel-upload DataFrame without writing any file."""
    with get_engine().connect() as conn:
        return pd.read_sql(_GENERIC_QUERY, conn)
