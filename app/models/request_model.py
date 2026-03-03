"""Request models for mapping API."""
from pydantic import BaseModel, Field


class MappingRequest(BaseModel):
    """Request body for auto-map endpoint."""

    client_name: str = Field(..., description="Name of the client (e.g. ABC Finance)")
    client_columns: list[str] = Field(..., description="Column names from client upload")
    lms_columns: list[str] = Field(..., description="LMS master column names")

    model_config = {"json_schema_extra": {"example": {
        "client_name": "ABC Finance",
        "client_columns": ["full_name", "dob", "pan_no"],
        "lms_columns": ["name", "date_of_birth", "pan_number"]
    }}}
