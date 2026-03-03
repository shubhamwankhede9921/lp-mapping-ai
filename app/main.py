"""LP Mapping Service - Intelligent Schema Matching Engine."""
from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.config import get_settings
from app.api import mapping_router
from app.repository.database import init_db


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Create DB tables on startup."""
    init_db()
    yield
    # shutdown if needed
    pass


app = FastAPI(
    title=get_settings().app_name,
    description="AI-assisted column mapping from client uploads to LMS master schema.",
    version="1.0.0",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(mapping_router, prefix="/api")


@app.get("/health")
def health():
    return {"status": "ok", "service": get_settings().app_name}


@app.get("/")
def root():
    return {
        "service": get_settings().app_name,
        "docs": "/docs",
        "api": "/api/mapping/auto-map",
    }
