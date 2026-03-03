"""Semantic similarity using sentence embeddings."""
from __future__ import annotations

import numpy as np
from typing import Optional

# Lazy load model to avoid slow startup when not used
_model = None
_embedding_dim: Optional[int] = None


def _get_model():
    global _model
    if _model is None:
        from sentence_transformers import SentenceTransformer
        # Lightweight model; swap for larger one if needed
        _model = SentenceTransformer("all-MiniLM-L6-v2")
    return _model


def embed(texts: list[str]) -> np.ndarray:
    """Get embeddings for a list of strings. Shape (n, dim)."""
    if not texts:
        return np.array([]).reshape(0, 384)
    model = _get_model()
    return model.encode(texts, convert_to_numpy=True)


def cosine_similarity(a: np.ndarray, b: np.ndarray) -> float:
    """Single pair cosine similarity. a, b are 1d arrays."""
    na = np.asarray(a, dtype=float).flatten()
    nb = np.asarray(b, dtype=float).flatten()
    if na.size != nb.size or na.size == 0:
        return 0.0
    dot = float(np.dot(na, nb))
    norm_a = float(np.linalg.norm(na))
    norm_b = float(np.linalg.norm(nb))
    if norm_a == 0 or norm_b == 0:
        return 0.0
    return dot / (norm_a * norm_b)


def find_best_embedding_match(
    client_column: str,
    lms_columns: list[str],
    lms_embeddings: Optional[np.ndarray] = None,
) -> tuple[str | None, float]:
    """
    Find best LMS column by embedding similarity.
    If lms_embeddings is provided, shape must be (len(lms_columns), dim).
    Returns (lms_column, score in 0-1) or (None, 0).
    """
    if not lms_columns:
        return (None, 0.0)
    model = _get_model()
    client_emb = model.encode([client_column], convert_to_numpy=True)
    if lms_embeddings is not None and lms_embeddings.shape[0] == len(lms_columns):
        lms_em = lms_embeddings
    else:
        lms_em = model.encode(lms_columns, convert_to_numpy=True)
    best_idx = -1
    best_score = 0.0
    for i in range(len(lms_columns)):
        sim = cosine_similarity(client_emb[0], lms_em[i])
        if sim > best_score:
            best_score = sim
            best_idx = i
    if best_idx < 0:
        return (None, 0.0)
    return (lms_columns[best_idx], float(best_score))


def compute_lms_embeddings(lms_columns: list[str]) -> np.ndarray:
    """Precompute embeddings for all LMS columns (for reuse)."""
    if not lms_columns:
        return np.array([]).reshape(0, 384)
    return embed(lms_columns)
