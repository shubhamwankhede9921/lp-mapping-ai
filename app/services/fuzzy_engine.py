"""
services/fuzzy_engine.py
Layer 2a — rapidfuzz token_sort_ratio matching against all known excel_keys.
Uses shared match_context: field, column_category, entity, process_name.
"""

import logging
import re
from typing import Any, Dict, List, Optional, Tuple

from app.services.match_context import (
    ENTITY_ROLE,
    build_semantic_query,
    compute_category_boost,
    effective_process,
    is_process_compatible,
    semantic_field_guard_reason,
)

logger = logging.getLogger(__name__)

try:
    from rapidfuzz import fuzz, process as rfprocess
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False
    logger.warning("rapidfuzz not installed — fuzzy engine disabled")


def _norm(text: str) -> str:
    text = text.lower()
    text = re.sub(r"[_\s.\-'()]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


class FuzzyEngine:
    def __init__(
        self,
        field_dictionary: Dict[str, Any],
        threshold: float = 0.72,
        process_name: str = "",
    ):
        if not HAS_RAPIDFUZZ:
            raise ImportError("pip install rapidfuzz")
        self.threshold = threshold
        self.process_name = process_name
        self._field_dictionary = field_dictionary

        # Build corpus: (normalized_label, excel_key, role, description)
        self._corpus: List[Tuple[str, str, str, str]] = []
        self._json_key_map: Dict[str, str] = {}

        for ek, info in field_dictionary.get("by_excel_key", {}).items():
            role = info.get("role") or info.get("json_key_role", "")
            desc = (info.get("description", "") or "").strip()

            self._corpus.append((_norm(ek), ek, role, desc))
            if desc:
                self._corpus.append((_norm(desc), ek, role, desc))

            self._json_key_map[ek] = info.get("json_key", "")

        self._labels = [c[0] for c in self._corpus]
        logger.info(
            f"FuzzyEngine: {len(self._corpus)} labels indexed, "
            f"process_name={process_name!r}"
        )

    def match(
        self,
        field_name: str,
        entity: Optional[str] = None,
        column_category: Optional[str] = None,
        process_name: Optional[str] = None,
    ) -> Optional[Dict[str, Any]]:
        proc = effective_process(process_name, self.process_name)
        if semantic_field_guard_reason(field_name):
            return None
        query = build_semantic_query(field_name, column_category, entity, proc)
        if not query:
            return None

        expected_role = ENTITY_ROLE.get((entity or "").strip().upper(), "")
        hits = rfprocess.extract(query, self._labels, scorer=fuzz.token_sort_ratio, limit=25)
        best = None

        for label, score_100, idx in hits:
            score = score_100 / 100.0
            if score < self.threshold:
                continue

            _, ek, role, desc = self._corpus[idx]

            if proc and not is_process_compatible(ek, proc, self._field_dictionary):
                logger.debug(
                    "Fuzzy: Filtered %r — incompatible with process=%r",
                    ek,
                    proc,
                )
                continue

            if expected_role and role and role != expected_role:
                score *= 0.85

            category_boost = compute_category_boost(column_category, ek)
            score *= category_boost

            if score < self.threshold:
                continue

            if best is None or score > best["confidence"]:
                best = {
                    "matched_excel_key": ek,
                    "confidence": round(score, 4),
                    "match_type": "fuzzy",
                    "reasoning": (
                        f"Fuzzy: field={field_name!r} entity={entity!r} category={column_category!r} "
                        f"process={proc!r} ~ {ek!r} "
                        f"(base_score={score_100 / 100:.2f}, "
                        f"entity_adj={1 if not expected_role or role == expected_role else '0.85x'}, "
                        f"category_boost={category_boost:.2f})"
                    ),
                    "fuzzy_score": round(score, 4),
                    "winning_engine": "fuzzy",
                    "json_key": self._json_key_map.get(ek, ""),
                }

        return best

    def run_batch(
        self,
        unmatched: List[Dict[str, Any]],
        process_name: Optional[str] = None,
    ) -> Tuple[List[Dict], List[Dict]]:
        matched, still_unmatched = [], []
        proc = effective_process(process_name, self.process_name)
        for f in unmatched:
            result = self.match(
                field_name=f.get("partner_field", f.get("field_name", "")),
                entity=f.get("entity"),
                column_category=f.get("column_category"),
                process_name=proc or None,
            )
            if result:
                merged = {**f, **result, "needs_review": result["confidence"] < 0.80}
                if proc:
                    merged["process_name"] = proc
                matched.append(merged)
            else:
                still_unmatched.append(f)
        logger.info(
            "FuzzyEngine: matched=%d, still_unmatched=%d",
            len(matched),
            len(still_unmatched),
        )
        return matched, still_unmatched
