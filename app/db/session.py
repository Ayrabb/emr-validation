# app/db/session.py
# Engine, session factory, and DB initialisation.
# All other modules import get_db() and use it as a context manager.

import logging
import os
from contextlib import contextmanager
from pathlib import Path

from sqlalchemy import create_engine, event, text
from sqlalchemy.orm import sessionmaker, Session

from app.core.config import settings
from app.db.models import Base

logger = logging.getLogger(__name__)

# ── Engine ────────────────────────────────────────────────────────────────────
# check_same_thread=False is required for SQLite when the engine is shared
# across threads (background scheduler + API workers).
# WAL journal mode gives concurrent reads without blocking writes.

def _create_engine():
    db_path = Path(settings.database_path)
    db_path.parent.mkdir(parents=True, exist_ok=True)

    engine = create_engine(
        settings.database_url,
        connect_args={"check_same_thread": False},
        echo=False,
    )

    # Enable WAL mode on every new connection
    @event.listens_for(engine, "connect")
    def _set_wal(dbapi_connection, _):
        dbapi_connection.execute("PRAGMA journal_mode=WAL")
        dbapi_connection.execute("PRAGMA synchronous=NORMAL")
        dbapi_connection.execute("PRAGMA foreign_keys=ON")

    return engine


engine = _create_engine()
SessionLocal = sessionmaker(bind=engine, autocommit=False, autoflush=False)


def _migrate_db() -> None:
    """Apply column and index additions that create_all cannot handle on existing tables."""
    with engine.connect() as conn:

        # ── facility_errors column migrations ─────────────────────────────────
        result = conn.execute(text("PRAGMA table_info(facility_errors)"))
        existing_cols = {row[1] for row in result}

        if "state" not in existing_cols:
            conn.execute(text("ALTER TABLE facility_errors ADD COLUMN state VARCHAR(100)"))
            conn.commit()
            logger.info("DB migration: added 'state' column to facility_errors")

        if "patient_uid" not in existing_cols:
            try:
                conn.execute(text("ALTER TABLE facility_errors ADD COLUMN patient_uid VARCHAR(100)"))
                conn.commit()
                logger.info("DB migration: added 'patient_uid' column to facility_errors")
            except Exception:
                conn.rollback()
                logger.debug("DB migration: 'patient_uid' column already exists (race condition ignored)")

        # ── Indexes (safe to run — IF NOT EXISTS prevents duplicates) ─────────
        conn.execute(text(
            "CREATE INDEX IF NOT EXISTS ix_fe_run_state "
            "ON facility_errors (run_id_fk, state)"
        ))
        conn.execute(text(
            "CREATE INDEX IF NOT EXISTS ix_vr_status_runtime "
            "ON validation_runs (status, run_time)"
        ))
        conn.commit()


def _cleanup_stale_runs() -> None:
    """Mark any runs left in 'running'/'pending' state as 'error'.
    These are runs that were interrupted by a service restart and will
    never complete — leaving them in a non-terminal state causes them
    to appear as in-progress indefinitely.
    """
    from app.db.models import ValidationRun  # local import to avoid circular
    with get_db() as db:
        stale = (
            db.query(ValidationRun)
            .filter(ValidationRun.status.in_(["running", "pending"]))
            .all()
        )
        if stale:
            for run in stale:
                run.status        = "error"
                run.error_message = "Service restarted while run was in progress"
            logger.warning(f"Marked {len(stale)} stale run(s) as error on startup")


def init_db() -> None:
    """Create all tables if they don't exist, then apply migrations."""
    Base.metadata.create_all(bind=engine)
    _migrate_db()
    _cleanup_stale_runs()
    logger.info(f"Database ready: {settings.database_path}")


@contextmanager
def get_db() -> Session:
    """Context manager that yields a session and commits/rolls back cleanly."""
    session = SessionLocal()
    try:
        yield session
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()
