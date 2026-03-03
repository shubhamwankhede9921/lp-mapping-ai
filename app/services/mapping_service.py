"""Main mapping orchestration: rule + fuzzy + embedding + historical + optional LLM."""
from app.config import get_settings
from app.models.response_model import ColumnMapping, MappingResponse, MappingSuggestion
from app.services.rule_engine import find_best_rule_match
from app.services.fuzzy_engine import find_best_fuzzy_match
from app.services.embedding_engine import find_best_embedding_match, compute_lms_embeddings
from app.services.llm_service import validate_mapping_with_llm
from app.repository.mapping_repository import (
    get_historical_patterns,
    find_best_historical_match,
)

# Weights for final score (must sum to 1.0)
W_RULE = 0.4
W_FUZZY = 0.2
W_EMBEDDING = 0.2
W_HISTORICAL = 0.2


def _tier(confidence: float) -> str:
    settings = get_settings()
    if confidence >= settings.auto_map_threshold:
        return "auto_map"
    if confidence >= settings.suggest_threshold:
        return "suggest"
    return "manual_required"


def generate_mapping(
    client_name: str,
    client_columns: list[str],
    lms_columns: list[str],
    *,
    use_llm: bool = False,
    llm_model: str | None = None,
) -> MappingResponse:
    """
    For each client column, compute best LMS match and confidence using
    rule + fuzzy + embedding + historical, then optionally validate with LLM.
    """
    if not lms_columns:
        return MappingResponse(mappings=[], client_name=client_name)

    settings = get_settings()
    patterns = get_historical_patterns()
    lms_embeddings = compute_lms_embeddings(lms_columns)

    suggestions: list[dict] = []
    for client_col in client_columns:
        # 1. Rule (0-100 -> 0-1)
        rule_lms, rule_raw = find_best_rule_match(client_col, lms_columns)
        rule_score = rule_raw / 100.0 if rule_raw else 0.0

        # 2. Fuzzy (0-1)
        fuzzy_lms, fuzzy_score = find_best_fuzzy_match(client_col, lms_columns)

        # 3. Embedding (0-1)
        emb_lms, emb_score = find_best_embedding_match(
            client_col, lms_columns, lms_embeddings
        )

        # 4. Historical (0-1)
        hist_lms, hist_score = find_best_historical_match(
            client_col, lms_columns, patterns=patterns
        )

        # Weighted combination: pick best LMS per engine then combine scores
        # Simplified: take best LMS by weighted vote (by score) and average scores for that
        candidates: dict[str, list[float]] = {}
        for lms, s in [(rule_lms, rule_score), (fuzzy_lms, fuzzy_score), (emb_lms, emb_score), (hist_lms, hist_score)]:
            if lms is None:
                continue
            if lms not in candidates:
                candidates[lms] = []
            candidates[lms].append(s)

        # Final score for each candidate: weighted sum of the 4 components
        # We have one "best" per engine; combine by taking the candidate that appears
        # and weighting by the engine that suggested it.
        best_lms: str | None = None
        best_final = 0.0
        for lms, scores in candidates.items():
            # Recompute weighted score: we need rule/fuzzy/emb/hist for THIS lms
            r = rule_score if rule_lms == lms else 0.0
            f = fuzzy_score if fuzzy_lms == lms else 0.0
            e = emb_score if emb_lms == lms else 0.0
            h = hist_score if hist_lms == lms else 0.0
            final = W_RULE * r + W_FUZZY * f + W_EMBEDDING * e + W_HISTORICAL * h
            if final > best_final:
                best_final = final
                best_lms = lms

        if best_lms is None:
            best_lms = lms_columns[0]
            best_final = 0.0

        suggestions.append({
            "client_column": client_col,
            "lms_column": best_lms,
            "confidence": round(best_final, 4),
        })

    if use_llm and suggestions:
        suggestions = validate_mapping_with_llm(
            llm_model or settings.llm_model_name,
            lms_columns,
            client_columns,
            suggestions,
        )
        # Normalize to our shape
        suggestions = [
            {
                "client_column": s.get("client_column", ""),
                "lms_column": s.get("lms_column", ""),
                "confidence": float(s.get("confidence", 0)),
            }
            for s in suggestions
        ]

    mappings = [
        ColumnMapping(
            client_column=s["client_column"],
            lms_column=s["lms_column"],
            confidence=s["confidence"],
        )
        for s in suggestions
    ]
    return MappingResponse(mappings=mappings, client_name=client_name)


def generate_mapping_with_tiers(
    client_name: str,
    client_columns: list[str],
    lms_columns: list[str],
    *,
    use_llm: bool = False,
    llm_model: str | None = None,
) -> list[MappingSuggestion]:
    """Same as generate_mapping but returns list of MappingSuggestion with tier for UI."""
    resp = generate_mapping(client_name, client_columns, lms_columns, use_llm=use_llm, llm_model=llm_model)
    return [
        MappingSuggestion(
            client_column=m.client_column,
            lms_column=m.lms_column,
            confidence=m.confidence,
            tier=_tier(m.confidence),
        )
        for m in resp.mappings
    ]
