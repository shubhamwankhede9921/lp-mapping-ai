"""Response models for mapping API."""
from pydantic import BaseModel, Field


class ColumnMapping(BaseModel):
    """A single client column to LMS column mapping."""

    client_column: str
    lms_column: str
    confidence: float = Field(..., ge=0, le=1)


class MappingSuggestion(BaseModel):
    """Mapping with confidence tier for UI (auto-map / suggest / manual)."""

    client_column: str
    lms_column: str
    confidence: float = Field(..., ge=0, le=1)
    tier: str = Field(..., description="auto_map | suggest | manual_required")


class MappingResponse(BaseModel):
    """Response from auto-map endpoint."""

    mappings: list[ColumnMapping] = Field(..., description="Suggested column mappings")
    client_name: str = ""

    model_config = {"json_schema_extra": {"example": {
        "mappings": [
            {"client_column": "full_name", "lms_column": "name", "confidence": 0.94},
            {"client_column": "dob", "lms_column": "date_of_birth", "confidence": 0.91}
        ],
        "client_name": "ABC Finance"
    }}}
