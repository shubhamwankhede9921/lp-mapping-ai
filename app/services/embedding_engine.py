"""
services/embedding_engine.py
Layer 2b — sentence-transformers cosine similarity (optional).
Uses shared match_context: field, column_category, entity, process_name.
Install: pip install sentence-transformers
"""

import logging
from typing import Any, Dict, List, Optional, Tuple

import numpy as np

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
    from sentence_transformers import SentenceTransformer
    HAS_ST = True
except ImportError:
    HAS_ST = False
    logger.warning("sentence-transformers not installed — embedding engine disabled")


class EmbeddingEngine:
    MODEL = "all-MiniLM-L6-v2"

    def __init__(
        self,
        field_dictionary: Dict[str, Any],
        threshold: float = 0.60,
        process_name: str = "",
    ):
        if not HAS_ST:
            raise ImportError("pip install sentence-transformers")
        self.threshold = threshold
        self.process_name = process_name
        self._field_dictionary = field_dictionary
        self.model = SentenceTransformer(self.MODEL)

        self._corpus: List[Tuple[str, str, str]] = []  # (text, excel_key, role)
        self._json_key_map: Dict[str, str] = {}

        for ek, info in field_dictionary.get("by_excel_key", {}).items():
            desc = (info.get("description", "") or "").strip()
            role = info.get("role") or info.get("json_key_role", "")
            proc = (info.get("process_name") or "").strip()
            pnames = info.get("process_names") or []
            proc_extra = " ".join(str(p) for p in pnames if p)

            # Corpus: excel key + description + role + process tags for alignment with queries
            text = " ".join(
                x for x in (ek, desc, f"role {role}" if role else "", proc, proc_extra) if x
            ).strip()
            self._corpus.append((text, ek, role))
            self._json_key_map[ek] = info.get("json_key", "")

        logger.info("EmbeddingEngine: encoding %d entries …", len(self._corpus))
        texts = [c[0] for c in self._corpus]
        self._embs = self.model.encode(texts, normalize_embeddings=True)
        logger.info("EmbeddingEngine ready (process_name=%r)", self.process_name)

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
        query_text = build_semantic_query(field_name, column_category, entity, proc)
        if not query_text:
            return None

        q = self.model.encode([query_text], normalize_embeddings=True)[0]
        sims = np.dot(self._embs, q)

        expected = ENTITY_ROLE.get((entity or "").strip().upper(), "")
        best_idx = -1
        best_sim = -1.0
        best_conf = 0.0

        for idx in np.argsort(sims)[::-1]:
            sim = float(sims[idx])
            if sim < self.threshold:
                break

            _, ek, role = self._corpus[idx]

            if proc and not is_process_compatible(ek, proc, self._field_dictionary):
                logger.debug(
                    "Embedding: Filtered %r — incompatible with process=%r",
                    ek,
                    proc,
                )
                continue

            conf = sim

            if expected and role and role != expected:
                conf *= 0.90

            category_boost = compute_category_boost(column_category, ek)
            conf *= category_boost

            if conf < self.threshold:
                continue

            if conf > best_conf:
                best_conf = conf
                best_sim = sim
                best_idx = idx

        if best_idx < 0:
            return None

        _, ek, role = self._corpus[best_idx]
        conf = round(best_conf, 4)
        cat_b = compute_category_boost(column_category, ek)

        return {
            "matched_excel_key": ek,
            "confidence": conf,
            "match_type": "embedding",
            "reasoning": (
                f"Embedding: field={field_name!r} entity={entity!r} category={column_category!r} "
                f"process={proc!r} → {ek!r} (sim={best_sim:.3f}, "
                f"entity_adj={1 if not expected or role == expected else '0.90x'}, "
                f"category_boost={cat_b:.2f})"
            ),
            "embedding_score": conf,
            "winning_engine": "embedding",
            "json_key": self._json_key_map.get(ek, ""),
        }

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
            "EmbeddingEngine: matched=%d, still_unmatched=%d",
            len(matched),
            len(still_unmatched),
        )
        return matched, still_unmatched
