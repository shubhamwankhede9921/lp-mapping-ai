"""Mapping API endpoints."""
from fastapi import APIRouter, Query
from pydantic import BaseModel, Field
from app.models.request_model import MappingRequest
from app.models.response_model import MappingResponse, MappingSuggestion
from app.services.mapping_service import generate_mapping, generate_mapping_with_tiers
from app.repository.mapping_repository import save_mappings_bulk

router = APIRouter(prefix="/mapping", tags=["mapping"])


class SaveMappingsRequest(BaseModel):
    """Save confirmed mappings for future learning."""
    client_name: str = Field(..., description="Client name")
    mappings: list[dict] = Field(..., description="List of {client_column, lms_column, confidence}")


@router.post("/auto-map", response_model=MappingResponse)
def auto_map(request: MappingRequest):
    """
    Suggest column mappings from client columns to LMS columns using
    rule-based + fuzzy + embedding + historical patterns. Optionally validate with LLM.
    """
    result = generate_mapping(
        request.client_name,
        request.client_columns,
        request.lms_columns,
        use_llm=False,
    )
    return result


@router.post("/auto-map-with-llm", response_model=MappingResponse)
def auto_map_with_llm(
    request: MappingRequest,
    model: str | None = Query(None, description="LLM model name"),
):
    """Same as auto-map but runs final validation through Dvara LLM."""
    result = generate_mapping(
        request.client_name,
        request.client_columns,
        request.lms_columns,
        use_llm=True,
        llm_model=model,
    )
    return result


@router.post("/suggestions", response_model=list[MappingSuggestion])
def mapping_suggestions(request: MappingRequest):
    """Returns mappings with confidence tier: auto_map | suggest | manual_required (for UI)."""
    return generate_mapping_with_tiers(
        request.client_name,
        request.client_columns,
        request.lms_columns,
    )


@router.post("/save")
def save_mappings(request: SaveMappingsRequest):
    """Persist confirmed mappings so the system learns for future clients."""
    tuples = [
        (m["client_column"], m["lms_column"], float(m.get("confidence", 0)))
        for m in request.mappings
        if isinstance(m, dict) and "client_column" in m and "lms_column" in m
    ]
    save_mappings_bulk(request.client_name, tuples)
    return {"status": "saved", "count": len(tuples)}
