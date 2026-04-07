from functools import lru_cache
from pathlib import Path
from pydantic_settings import BaseSettings
from pydantic import Field


class Settings(BaseSettings):
    # LLM Gateway
    llm_gateway_url: str    = Field(...,                env="LLM_GATEWAY_URL")
    llm_gateway_token: str  = Field(...,                env="LLM_GATEWAY_TOKEN")
    parameter_classifier_gateway_url: str = Field(
        "",
        env="PARAMETER_CLASSIFIER_GATEWAY_URL",
    )
    llm_use_form_data: bool = Field(True,               env="LLM_USE_FORM_DATA")
    llm_task_field: str     = Field("task",             env="LLM_TASK_FIELD")
    llm_task_value: str     = Field("lp_field_mapping", env="LLM_TASK_VALUE")
    llm_prompt_field: str   = Field("prompt",           env="LLM_PROMPT_FIELD")

    # ── Source DB (PUTM + generic_excel_upload reads) ─────────────────────────
    db_host: str       = Field(...,      env="DB_HOST")
    db_port: int       = Field(3306,     env="DB_PORT")
    db_name: str       = Field(...,      env="DB_NAME")
    db_user: str       = Field(...,      env="DB_USER")
    db_password: str   = Field(...,      env="DB_PASSWORD")
    db_schema: str     = Field("public", env="DB_SCHEMA")
    putm_table: str    = Field("putm_field_definitions",                 env="PUTM_TABLE")
    mapping_table: str = Field("generic_excel_upload_definition_fields", env="MAPPING_TABLE")

    # ── Target DB (pipeline results are written here) ─────────────────────────
    target_db_host: str        = Field(...,  env="TARGET_DB_HOST")
    target_db_port: int        = Field(3306, env="TARGET_DB_PORT")
    target_db_name: str        = Field(...,  env="TARGET_DB_NAME")
    target_db_user: str        = Field(...,  env="TARGET_DB_USER")
    target_db_password: str    = Field(...,  env="TARGET_DB_PASSWORD")
    generic_mapping_table: str = Field(
        "generic_excel_upload_definition_fields",
        env="GENERIC_MAPPING_TABLE",
    )
    # ─────────────────────────────────────────────────────────────────────────

    # App
    app_host: str  = Field("0.0.0.0", env="APP_HOST")
    app_port: int  = Field(8000,      env="APP_PORT")
    log_level: str = Field("INFO",    env="LOG_LEVEL")

    # Paths
    references_dir: str = Field(
        r"D:\LP_AUTOMATION\column-mapping-service\app\references",
        env="REFERENCES_DIR",
    )
    scripts_dir: str = Field(
        r"D:\LP_AUTOMATION\column-mapping-service\scripts",
        env="SCRIPTS_DIR",
    )
    output_dir: str = Field("./pipeline_output", env="OUTPUT_DIR")

    # Thresholds
    fuzzy_threshold: float     = Field(0.72, env="FUZZY_THRESHOLD")
    embedding_threshold: float = Field(0.60, env="EMBEDDING_THRESHOLD")
    review_threshold: float    = Field(0.80, env="REVIEW_THRESHOLD")

    class Config:
        env_file = ".env"
        case_sensitive = False

    @property
    def refs_path(self) -> Path:
        return Path(self.references_dir)

    @property
    def scripts_path(self) -> Path:
        return Path(self.scripts_dir)

    @property
    def output_path(self) -> Path:
        return Path(self.output_dir)

    @property
    def source_db_url(self) -> str:
        """SQLAlchemy URL for the source DB (PUTM reads)."""
        return (
            f"mysql+pymysql://{self.db_user}:{self.db_password}"
            f"@{self.db_host}:{self.db_port}/{self.db_name}"
        )

    @property
    def target_db_url(self) -> str:
        """SQLAlchemy URL for the target DB (mapping results write)."""
        return (
            f"mysql+pymysql://{self.target_db_user}:{self.target_db_password}"
            f"@{self.target_db_host}:{self.target_db_port}/{self.target_db_name}"
        )


@lru_cache()
def get_settings() -> Settings:
    return Settings()
