# app/main.py
# Changes from desktop: DB init, APScheduler lifespan, updated CORS.

import logging
import os
import threading
from contextlib import asynccontextmanager
from datetime import datetime

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.api.validation_routes import router, _job_store
from app.core.config import settings
from app.db.session import init_db
from app.db.repository import RunRepository
from app.scheduler.pipeline import run_pipeline, set_job_store
from app.scheduler.setup import init_scheduler, shutdown_scheduler

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(name)s: %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


def _nearest_slot() -> str:
    """Return the slot label ('06:00'|'12:00'|'18:00') nearest to now.
    Used to label the startup run so it appears in BayCentral's run list."""
    hour = datetime.now().hour
    if hour < 6:
        return "06:00"
    if hour < 12:
        return "12:00"
    if hour < 18:
        return "18:00"
    return "06:00"   # after 18:00 — label as morning so it shows until 06:00 fires


def _maybe_run_initial_pipeline() -> None:
    """On first install (no completed runs in DB), fire one pipeline run immediately
    in a background daemon thread so the service has data as soon as it's ready."""
    if RunRepository.get_last_completed() is not None:
        return
    slot = _nearest_slot()
    logger.info(
        f"First install detected — no completed runs found. "
        f"Starting initial pipeline run (slot={slot}) in background..."
    )
    t = threading.Thread(
        target=run_pipeline,
        args=(slot,),
        daemon=True,
        name="initial-pipeline",
    )
    t.start()


@asynccontextmanager
async def lifespan(app: FastAPI):
    # STARTUP
    os.makedirs(settings.report_dir, exist_ok=True)
    init_db()
    RunRepository.purge_old_runs()
    set_job_store(_job_store)
    init_scheduler()
    _maybe_run_initial_pipeline()
    logger.info(
        f"EMR Validation Service v{settings.app_version} started — port {settings.port}"
    )
    logger.info(f"Reports directory : {settings.report_dir}")
    logger.info(f"Database          : {settings.database_path}")
    logger.info(f"Timezone          : {settings.scheduler_timezone}")
    logger.info(f"API docs          : http://127.0.0.1:{settings.port}/docs")
    yield
    # SHUTDOWN
    shutdown_scheduler()
    logger.info("EMR Validation Service stopped.")


app = FastAPI(
    title="EMR Validation Service",
    description=(
        "HIV Programme RADET Data Quality Validation — Central Service.\n\n"
        "Runs automatically at 06:00, 12:00, and 18:00 daily.\n"
        "BayCentral connects to read completed results only — users cannot trigger validation."
    ),
    version=settings.app_version,
    lifespan=lifespan,
    redoc_url=None,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.allowed_origins,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)

app.include_router(router)
