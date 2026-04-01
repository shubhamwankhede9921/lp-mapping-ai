"""Persistence for existing column mappings (historical reference)."""
from datetime import datetime
from sqlalchemy import Column, Integer, String, Float, DateTime
from sqlalchemy.orm import Session
from collections import defaultdict

from app.repository.database import Base


class ExistingColumnMapping(Base):
    """Table: existing_column_mapping."""

    __tablename__ = "existing_column_mapping"

    id = Column(Integer, primary_key=True, autoincrement=True)
    client_name = Column(String(255), nullable=False)
    client_column = Column(String(255), nullable=False)
    lms_column = Column(String(255), nullable=False)
    confidence_score = Column(Float, default=0.0)
    created_at = Column(DateTime, default=datetime.utcnow)


def _ensure_db():
    try:
        init_db()
    except Exception:
        pass


def save_mapping(
    client_name: str,
    client_column: str,
    lms_column: str,
    confidence_score: float = 0.0,
    db: Session | None = None,
) -> None:
    """Store one mapping for future reference."""
    _ensure_db()
    session = db or get_db()
    try:
        rec = ExistingColumnMapping(
            client_name=client_name,
            client_column=client_column,
            lms_column=lms_column,
            confidence_score=confidence_score,
        )
        session.add(rec)
        if not db:
            session.commit()
    finally:
        if not db:
            session.close()


def save_mappings_bulk(
    client_name: str,
    mappings: list[tuple[str, str, float]],
    db: Session | None = None,
) -> None:
    """Save multiple (client_column, lms_column, confidence) for a client."""
    _ensure_db()
    session = db or get_db()
    try:
        for client_col, lms_col, conf in mappings:
            rec = ExistingColumnMapping(
                client_name=client_name,
                client_column=client_col,
                lms_column=lms_col,
                confidence_score=conf,
            )
            session.add(rec)
        if not db:
            session.commit()
    finally:
        if not db:
            session.close()


def get_historical_patterns(db: Session | None = None) -> dict[str, list[str]]:
    """
    Group previous client columns by lms_column.
    Returns: { "lms_column": ["client_col_1", "client_col_2", ...] }
    Used to score new client columns by historical pattern.
    """
    _ensure_db()
    session = db or get_db()
    try:
        rows = session.query(ExistingColumnMapping).all()
        by_lms: dict[str, list[str]] = defaultdict(list)
        for r in rows:
            by_lms[r.lms_column].append(r.client_column)
        return dict(by_lms)
    finally:
        if not db:
            session.close()


def historical_match_score(
    client_column: str, lms_column: str, patterns: dict[str, list[str]]
) -> float:
    """
    Score 0-1: how well client_column fits historical pattern for lms_column.
    Simple heuristic: if any historical client column for this lms_column
    is a substring of or equal to client_column (normalized), return high score.
    """
    from app.utils.text_utils import normalize_column_name
    client_norm = normalize_column_name(client_column)
    if lms_column not in patterns:
        return 0.0
    for past in patterns[lms_column]:
        past_norm = normalize_column_name(past)
        if client_norm == past_norm:
            return 1.0
        if past_norm in client_norm or client_norm in past_norm:
            return 0.85
    return 0.0


def find_best_historical_match(
    client_column: str,
    lms_columns: list[str],
    patterns: dict[str, list[str]] | None = None,
    db: Session | None = None,
) -> tuple[str | None, float]:
    """Best LMS column by historical pattern. Returns (lms_column, 0-1 score)."""
    pat = patterns if patterns is not None else get_historical_patterns(db)
    best_col: str | None = None
    best_score = 0.0
    for lms in lms_columns:
        s = historical_match_score(client_column, lms, pat)
        if s > best_score:
            best_score = s
            best_col = lms
    return (best_col, best_score)
