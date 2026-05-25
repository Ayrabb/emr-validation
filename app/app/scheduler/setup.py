# app/scheduler/setup.py
# Creates and configures the BackgroundScheduler with three CronTrigger entries.
# The scheduler is started in the FastAPI lifespan (app/main.py) and shut down
# cleanly on application exit.

import logging
from datetime import datetime

import pytz
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

from app.core.config import settings
from app.scheduler.pipeline import run_pipeline

logger = logging.getLogger(__name__)

# Module-level scheduler instance — shared across the app lifetime
scheduler = BackgroundScheduler(timezone=settings.scheduler_timezone)


def init_scheduler() -> None:
    """Register all three daily jobs and start the scheduler.
    Called once from the FastAPI lifespan on startup.
    """
    tz = pytz.timezone(settings.scheduler_timezone)

    jobs = [
        ("06:00",  6,  0),
        ("12:00", 12,  0),
        ("18:00", 18,  0),
    ]

    for slot, hour, minute in jobs:
        scheduler.add_job(
            func=run_pipeline,
            trigger=CronTrigger(hour=hour, minute=minute, timezone=tz),
            args=[slot],
            id=f"pipeline_{slot.replace(':', '')}",
            name=f"RADET Validation — {slot}",
            replace_existing=True,
            misfire_grace_time=300,   # tolerate up to 5 min late start
        )
        logger.info(f"Scheduled job: {slot} ({settings.scheduler_timezone})")

    scheduler.start()
    logger.info("APScheduler started — 3 daily jobs active")


def shutdown_scheduler() -> None:
    """Stop the scheduler gracefully.  Called from FastAPI lifespan on shutdown."""
    if scheduler.running:
        scheduler.shutdown(wait=False)
        logger.info("APScheduler stopped")


def get_next_run_time() -> str | None:
    """Return the ISO 8601 UTC timestamp of the next scheduled job execution."""
    next_times = []
    for job in scheduler.get_jobs():
        if job.next_run_time:
            next_times.append(job.next_run_time)
    if not next_times:
        return None
    earliest = min(next_times)
    return earliest.astimezone(pytz.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
