"""Rule-based column matching (exact / normalized)."""
from app.utils.text_utils import (
    normalize_column_name,
    normalized_equals,
    normalized_equals_no_underscore,
)

# Score constants
EXACT_MATCH = 100.0
NORMALIZED_MATCH = 95.0
NO_MATCH = 0.0


def match_score(client_column: str, lms_column: str) -> float:
    """
    Returns 0-100 score.
    - 100 if exact match (after normalize)
    - 95 if match when ignoring underscores
    - 0 otherwise
    """
    if not client_column or not lms_column:
        return NO_MATCH
    n_client = normalize_column_name(client_column)
    n_lms = normalize_column_name(lms_column)
    if n_client == n_lms:
        return EXACT_MATCH
    if n_client.replace("_", "") == n_lms.replace("_", ""):
        return NORMALIZED_MATCH
    return NO_MATCH


def find_best_rule_match(
    client_column: str, lms_columns: list[str]
) -> tuple[str | None, float]:
    """
    Find best LMS column by rule matching.
    Returns (lms_column, score) or (None, 0).
    """
    best_col: str | None = None
    best_score = NO_MATCH
    for lms in lms_columns:
        s = match_score(client_column, lms)
        if s > best_score:
            best_score = s
            best_col = lms
    return (best_col, best_score)
