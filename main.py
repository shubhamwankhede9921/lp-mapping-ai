"""
main.py — FastAPI entry point for LP Field Mapping API
"""

import logging
import sys
from contextlib import asynccontextmanager
from pathlib import Path

import uvicorn
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.config import get_settings
from app.api.mapping_controller import router

settings = get_settings()

logging.basicConfig(
    level=getattr(logging, settings.log_level.upper(), logging.INFO),
    format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
)
logger = logging.getLogger("main")


@asynccontextmanager
async def lifespan(app: FastAPI):
    # Ensure scripts dir on sys.path so services can import matching_engine etc.
    scripts = str(settings.scripts_path.resolve())
    if scripts not in sys.path:
        sys.path.insert(0, scripts)
        logger.info(f"scripts dir on sys.path: {scripts}")

    settings.refs_path.mkdir(parents=True, exist_ok=True)
    settings.output_path.mkdir(parents=True, exist_ok=True)

    if (settings.refs_path / "field_dictionary.json").exists():
        logger.info(f"References ready at {settings.refs_path}")
    else:
        logger.warning(
            "References not found. Call POST /api/v1/references/build first."
        )
    yield
    logger.info("Shutting down.")


app = FastAPI(
    title="LP Field Mapping API",
    description=(
        "3-layer hybrid field mapping for Dvara lending partner onboarding.\n\n"
        "**Layer 1** — Deterministic alias + exact matching (~75–97% of fields)\n\n"
        "**Layer 2** — Fuzzy + optional embeddings + LLM semantic matching\n\n"
        "**Layer 3** — Auto-numbering, json_key resolution, formatted Excel output"
    ),
    version="2.0.0",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(router)

if __name__ == "__main__":
    uvicorn.run(
        "main:app",
        host=settings.app_host,
        port=settings.app_port,
        reload=True,
        log_level=settings.log_level.lower(),
    )