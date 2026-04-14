from typing import Any, Dict, List, Optional
from pydantic import BaseModel, Field, model_validator


# ── Request models ─────────────────────────────────────────────────────────────

class BuildRefsRequest(BaseModel):
    force_rebuild: bool = False
    putm_table_override: Optional[str] = None
    mapping_table_override: Optional[str] = None


class DeterministicRequest(BaseModel):
    client_name: str
    process_name: str = "COMBINED"
    sheet_filter: Optional[str] = None


class HybridLLMRequest(BaseModel):
    client_name: str
    process_name: str = "COMBINED"
    use_fuzzy: bool = True
    use_embeddings: bool = False
    use_llm: bool = True


class FullPipelineRequest(BaseModel):
    client_name: str
    process_name: str = "COMBINED"
    use_fuzzy: bool = True
    use_embeddings: bool = False
    use_llm: bool = True
    sheet_filter: Optional[str] = None


# ── Response models ────────────────────────────────────────────────────────────

class FieldMapping(BaseModel):
    partner_field: str
    column_category: Optional[str]
    entity: str
    matched_excel_key: Optional[str]
    json_key: Optional[str]
    confidence: float
    match_type: str
    reasoning: str
    previous_mapping_reason: Optional[str] = None
    llm_change_reason: Optional[str] = None
    llm_param_bucket_reason: Optional[str] = None
    needs_review: bool
    fuzzy_score: Optional[float] = None
    embedding_score: Optional[float] = None
    llm_confidence: Optional[float] = None
    winning_engine: Optional[str] = None


class Stats(BaseModel):
    total_fields: int
    matched: int
    unmatched: int
    match_rate_pct: float
    needs_review: int
    avg_confidence: float
    by_match_type: Dict[str, int]
    by_entity: Dict[str, int]
    by_confidence_band: Dict[str, int]


class BuildRefsResponse(BaseModel):
    success: bool
    message: str
    putm_rows: Optional[int] = None
    mapping_rows: Optional[int] = None
    field_count: Optional[int] = None
    alias_count: Optional[int] = None


class DeterministicResponse(BaseModel):
    client_name: str
    process_name: str
    mappings: List[FieldMapping]
    unmatched_fields: List[Dict[str, Any]]
    llm_prompts_count: int
    stats: Stats


class HybridLLMResponse(BaseModel):
    client_name: str
    process_name: str
    mappings: List[FieldMapping]
    stats: Stats
    engine_breakdown: Dict[str, int]


# ── LOS JSON / nested-mapping models ──────────────────────────────────────────

class MappingItem(BaseModel):
    """
    A single field mapping row — accepted by /generate-nested-mapping and /generate-schema.

    Supports BOTH:
    1. direct output of /mapping/hybrid-llm or /mapping/full-pipeline
       (uses 'partner_field' + 'json_key')
    2. transformed/manual payloads
       (uses 'client_column' + 'lms_column')
    """

    # Accept both formats
    client_column: Optional[str] = None
    partner_field: Optional[str] = None

    matched_excel_key: Optional[str] = None

    # lms_column = preferred target path
    # json_key = alias/fallback target path
    lms_column: Optional[str] = None
    json_key: Optional[str] = None

    confidence: Optional[float] = None
    entity: Optional[str] = None
    match_type: Optional[str] = None
    reasoning: Optional[str] = None
    needs_review: Optional[bool] = False

    @model_validator(mode="before")
    @classmethod
    def normalize_input(cls, values):
        """
        Normalize incoming payloads so chained calls work without transformation.
        - If client_column is missing, use partner_field
        - If lms_column is missing, use json_key
        """
        if isinstance(values, dict):
            if not values.get("client_column") and values.get("partner_field"):
                values["client_column"] = values["partner_field"]

            if not values.get("lms_column") and values.get("json_key"):
                values["lms_column"] = values["json_key"]

        return values

    def get_client_column(self) -> Optional[str]:
        """Return client_column, falling back to partner_field."""
        return self.client_column or self.partner_field

    def get_lms_column(self) -> Optional[str]:
        """Return lms_column, falling back to json_key."""
        return self.lms_column or self.json_key


class LOSJsonRequest(BaseModel):
    """
    Request body for /generate-nested-mapping and /generate-schema.

    Accepts the direct output of /mapping/hybrid-llm or /mapping/full-pipeline
    so the two calls can be chained with no transformation.
    """
    client_name: Optional[str] = Field(None, description="Client identifier")
    process_name: Optional[str] = Field(None, description="Process identifier")
    mappings: List[MappingItem] = Field(..., description="List of field mappings")


class NestedMappingResponse(BaseModel):
    """Response for POST /generate-nested-mapping"""
    client_name: Optional[str]
    los_json: Dict[str, Any]
    total_input: int
    mapped_count: int
    skipped_count: int
    processing_time_ms: float


class SchemaResponse(BaseModel):
    """Response for POST /generate-schema (leaf = null instead of client_column)"""
    client_name: Optional[str]

    # NOTE: 'schema' is a reserved name on Pydantic BaseModel — use 'los_schema'
    los_schema: Dict[str, Any]

    total_input: int
    mapped_count: int
    skipped_count: int
    processing_time_ms: float
