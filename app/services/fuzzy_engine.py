"""Fuzzy string matching for column names using rapidfuzz."""
from rapidfuzz import fuzz

# Threshold above which we consider a strong match
STRONG_MATCH_THRESHOLD = 85


def fuzzy_score(client_column: str, lms_column: str) -> float:
    """
    Returns 0-100 similarity (ratio).
    Uses token_sort_ratio for order-independent matching (e.g. name_full vs full_name).
    """
    if not client_column or not lms_column:
        return 0.0
    # ratio: simple string similarity
    r = fuzz.ratio(client_column.lower(), lms_column.lower())
    # token_sort_ratio: ignores word order
    ts = fuzz.token_sort_ratio(client_column.lower(), lms_column.lower())
    return max(r, ts) / 100.0  # return 0-1 for consistency with other engines


def find_best_fuzzy_match(
    client_column: str, lms_columns: list[str]
) -> tuple[str | None, float]:
    """
    Find best LMS column by fuzzy match.
    Returns (lms_column, score in 0-1) or (None, 0).
    """
    best_col: str | None = None
    best_score = 0.0
    for lms in lms_columns:
        s = fuzzy_score(client_column, lms)
        if s > best_score:
            best_score = s
            best_col = lms
    return (best_col, best_score)
