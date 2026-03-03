"""Application configuration."""
from pydantic_settings import BaseSettings
from functools import lru_cache


class Settings(BaseSettings):
    """App settings from environment."""

    app_name: str = "LP Mapping Service"
    debug: bool = False

    # Dvara LLM API (set in env or .env)
    llm_base_url: str = "http://localhost:8001"
    llm_model_name: str = "default"
    llm_api_key: str = ""

    # Database (either use DATABASE_URL or DB_* pieces for MySQL)
    database_url: str = "sqlite:///./mapping_data.db"
    db_host: str | None = None
    db_port: int | None = None
    db_user: str | None = None
    db_password: str | None = None
    db_database: str | None = None
    db_driver: str = "mysql+pymysql"

    # Confidence thresholds
    auto_map_threshold: float = 0.85
    suggest_threshold: float = 0.70

    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"

    @property
    def sqlalchemy_database_url(self) -> str:
        """
        Final SQLAlchemy URL.
        - If DB_HOST/DB_USER/DB_DATABASE are set, build MySQL URL from DB_*.
        - Otherwise fall back to DATABASE_URL (env) or default sqlite.
        """
        # If full URL is explicitly provided and no DB_* pieces, use it as-is
        if (
            self.database_url
            and not (self.db_host and self.db_user and self.db_database)
        ):
            return self.database_url

        # If DB_* pieces are available, use them to build a MySQL URL
        if self.db_host and self.db_user and self.db_database:
            port = f":{self.db_port}" if self.db_port else ""
            password = f":{self.db_password}" if self.db_password else ""
            return (
                f"{self.db_driver}://{self.db_user}{password}"
                f"@{self.db_host}{port}/{self.db_database}"
            )

        # Fallback to whatever is in database_url
        return self.database_url


@lru_cache
def get_settings() -> Settings:
    return Settings()
