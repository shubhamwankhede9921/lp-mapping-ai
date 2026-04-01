"""
services/fuzzy_engine.py
Layer 2a — rapidfuzz token_sort_ratio matching against all known excel_keys.
"""

import logging
import re
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

try:
    from rapidfuzz import fuzz, process as rfprocess
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False
    logger.warning("rapidfuzz not installed — fuzzy engine disabled")

ENTITY_ROLE = {
    "APPLICANT":   "CUSTOMER",
    "COAPPLICANT": "COAPPLICANT",
    "COAPPLICANT1":"COAPPLICANT",
    "COAPPLICANT2":"COAPPLICANT",
    "GUARANTOR":   "GUARANTOR",
    "LOAN":        "LOAN",
    "DOCUMENT":    "LOAN",
    "FEE":         "LOAN",
}


def _norm(text: str) -> str:
    text = text.lower()
    text = re.sub(r"[_\s.\-'()]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


class FuzzyEngine:
    def __init__(self, field_dictionary: Dict[str, Any], threshold: float = 0.72):
        if not HAS_RAPIDFUZZ:
            raise ImportError("pip install rapidfuzz")
        self.threshold = threshold

        # Build corpus: (normalized_label, excel_key, role)
        self._corpus: List[Tuple[str, str, str]] = []
        # Map excel_key → json_key
        self._json_key_map: Dict[str, str] = {}

        for ek, info in field_dictionary.get("by_excel_key", {}).items():
            role = info.get("role") or info.get("json_key_role", "")
            desc = (info.get("description", "") or "").strip()
            self._corpus.append((_norm(ek),   ek, role))
            if desc:
                self._corpus.append((_norm(desc), ek, role))

            # Store json_key for this excel_key (may be None)
            self._json_key_map[ek] = info.get("json_key", "")

        self._labels = [c[0] for c in self._corpus]
        logger.info(f"FuzzyEngine: {len(self._corpus)} candidate labels indexed")

    def match(self, field_name: str, entity: Optional[str] = None) -> Optional[Dict[str, Any]]:
        query = _norm(field_name)
        if not query:
            return None

        hits = rfprocess.extract(query, self._labels, scorer=fuzz.token_sort_ratio, limit=15)
        expected_role = ENTITY_ROLE.get(entity or "", "")
        best = None

        for label, score_100, idx in hits:
            score = score_100 / 100.0
            if score < self.threshold:
                continue
            _, ek, role = self._corpus[idx]
            # Penalise cross-role matches
            if expected_role and role and role != expected_role:
                score *= 0.85
            if score < self.threshold:
                continue
            if best is None or score > best["confidence"]:
                best = {
                    "matched_excel_key": ek,
                    "confidence":        round(score, 4),
                    "match_type":        "fuzzy",
                    "reasoning":         f"Fuzzy: '{field_name}' ~ '{ek}' score={score:.2f}",
                    "fuzzy_score":       round(score, 4),
                    "winning_engine":    "fuzzy",
                    "json_key":          self._json_key_map.get(ek, ""),  # included here
                }
        return best

    def run_batch(self, unmatched: List[Dict[str, Any]]) -> Tuple[List[Dict], List[Dict]]:
        matched, still_unmatched = [], []
        for f in unmatched:
            result = self.match(
                field_name=f.get("partner_field", f.get("field_name", "")),
                entity=f.get("entity"),
            )
            if result:
                merged = {**f, **result, "needs_review": result["confidence"] < 0.80}
                matched.append(merged)
            else:
                still_unmatched.append(f)
        logger.info(f"FuzzyEngine: matched={len(matched)}, still_unmatched={len(still_unmatched)}")
        return matched, still_unmatched