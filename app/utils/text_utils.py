"""Text normalization and preprocessing for column names."""
import re
from typing import Optional

# Common abbreviations -> full form for better matching
ABBREVIATIONS = {
    "cust": "customer",
    "id": "id",
    "no": "number",
    "num": "number",
    "dob": "date_of_birth",
    "pan": "pan",
    "aadhaar": "aadhar",
    "addr": "address",
    "amt": "amount",
    "mob": "mobile",
    "ph": "phone",
    "fn": "firstname",
    "ln": "lastname",
    "name": "name",
}

# Stopwords to remove (optional - can make matching worse for short columns, use carefully)
STOPWORDS = {"the", "a", "an", "of", "and", "or"}


def normalize_column_name(raw: str) -> str:
    """
    Normalize a column name for comparison.
    - Lowercase
    - Replace separators (space, hyphen, dot) with underscore
    - Remove special chars (keep alphanumeric and underscore)
    - Optionally expand common abbreviations
    """
    if not raw or not isinstance(raw, str):
        return ""
    s = raw.strip().lower()
    # Replace separators with underscore
    s = re.sub(r"[\s\-\.]+", "_", s)
    # Remove any character that is not alphanumeric or underscore
    s = re.sub(r"[^a-z0-9_]", "", s)
    # Collapse multiple underscores
    s = re.sub(r"_+", "_", s).strip("_")
    return s


def normalize_with_abbreviations(raw: str) -> str:
    """Normalize and expand known abbreviations (e.g. cust -> customer)."""
    normalized = normalize_column_name(raw)
    words = normalized.split("_")
    expanded = []
    for w in words:
        expanded.append(ABBREVIATIONS.get(w, w))
    return "_".join(expanded)


def normalized_equals(a: str, b: str) -> bool:
    """True if normalized forms are equal."""
    return normalize_column_name(a) == normalize_column_name(b)


def normalized_equals_no_underscore(a: str, b: str) -> bool:
    """Compare ignoring underscores (e.g. full_name vs fullname)."""
    na = normalize_column_name(a).replace("_", "")
    nb = normalize_column_name(b).replace("_", "")
    return na == nb
