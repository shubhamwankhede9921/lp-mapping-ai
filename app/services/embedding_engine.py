"""
services/embedding_engine.py
Layer 2b — sentence-transformers cosine similarity (optional).
Install: pip install sentence-transformers
"""

import logging
from typing import Any, Dict, List, Optional, Tuple

import numpy as np

logger = logging.getLogger(__name__)

try:
    from sentence_transformers import SentenceTransformer
    HAS_ST = True
except ImportError:
    HAS_ST = False
    logger.warning("sentence-transformers not installed — embedding engine disabled")

ENTITY_ROLE = {
    "APPLICANT": "CUSTOMER", "COAPPLICANT": "COAPPLICANT",
    "GUARANTOR": "GUARANTOR", "LOAN": "LOAN",
}


class EmbeddingEngine:
    MODEL = "all-MiniLM-L6-v2"

    def __init__(self, field_dictionary: Dict[str, Any], threshold: float = 0.60):
        if not HAS_ST:
            raise ImportError("pip install sentence-transformers")
        self.threshold = threshold
        self.model = SentenceTransformer(self.MODEL)

        self._corpus: List[Tuple[str, str, str]] = []  # (text, excel_key, role)
        for ek, info in field_dictionary.get("by_excel_key", {}).items():
            desc = (info.get("description", "") or "").strip()
            role = info.get("role") or info.get("json_key_role", "")
            text = f"{ek} {desc}".strip()
            self._corpus.append((text, ek, role))

        logger.info(f"EmbeddingEngine: encoding {len(self._corpus)} entries …")
        texts = [c[0] for c in self._corpus]
        self._embs = self.model.encode(texts, normalize_embeddings=True)
        logger.info("EmbeddingEngine ready")

    def match(self, field_name: str, entity: Optional[str] = None) -> Optional[Dict[str, Any]]:
        q = self.model.encode([field_name], normalize_embeddings=True)[0]
        sims = np.dot(self._embs, q)
        idx = int(np.argmax(sims))
        sim = float(sims[idx])
        if sim < self.threshold:
            return None

        _, ek, role = self._corpus[idx]
        expected = ENTITY_ROLE.get(entity or "", "")
        conf = sim * (0.90 if (expected and role and role != expected) else 1.0)

        if conf < self.threshold:
            return None

        return {
            "matched_excel_key": ek,
            "confidence":        round(conf, 4),
            "match_type":        "embedding",
            "reasoning":         f"Embedding: '{field_name}' → '{ek}' sim={sim:.3f}",
            "embedding_score":   round(conf, 4),
            "winning_engine":    "embedding",
        }

    def run_batch(
        self, unmatched: List[Dict[str, Any]]
    ) -> Tuple[List[Dict], List[Dict]]:
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
        logger.info(f"EmbeddingEngine: matched={len(matched)}, still_unmatched={len(still_unmatched)}")
        return matched, still_unmatched