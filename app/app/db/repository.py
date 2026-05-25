# app/db/repository.py
# All database queries live here.  No raw SQL anywhere else.

import math
import logging
from dataclasses import asdict
from datetime import datetime, date, timedelta
from typing import Optional

import pandas as pd
from sqlalchemy.orm import Session

from app.db.models import ValidationRun, FacilityError
from app.db.session import get_db
from app.core.config import settings

logger = logging.getLogger(__name__)

# ── Schedule slots ────────────────────────────────────────────────────────────
SLOTS = ["06:00", "12:00", "18:00"]


# ═══════════════════════════════════════════════════════════════════════════════
# RUN REPOSITORY
# ═══════════════════════════════════════════════════════════════════════════════

class RunRepository:

    @staticmethod
    def create(run_id: str, schedule_slot: str, blob_date: date) -> ValidationRun:
        with get_db() as db:
            run = ValidationRun(
                run_id=run_id,
                schedule_slot=schedule_slot,
                run_time=datetime.utcnow(),
                blob_date=blob_date,
                status="pending",
            )
            db.add(run)
            db.flush()
            db.refresh(run)
            # Detach so the caller can use it outside the session
            db.expunge(run)
        return run

    @staticmethod
    def set_running(run_id: str) -> None:
        with get_db() as db:
            run = db.query(ValidationRun).filter_by(run_id=run_id).first()
            if run:
                run.status = "running"

    @staticmethod
    def set_done(
        run_id: str,
        total_rows: int,
        total_violations: int,
        rules_run: int,
        validation_time_seconds: float,
        report_path: str,
        rule_summaries: list,
        facility_summaries: list,
    ) -> None:
        with get_db() as db:
            run = db.query(ValidationRun).filter_by(run_id=run_id).first()
            if run:
                run.status                  = "done"
                run.total_rows              = total_rows
                run.total_violations        = total_violations
                run.rules_run               = rules_run
                run.validation_time_seconds = validation_time_seconds
                run.report_path             = report_path
                run.set_result_json(rule_summaries, facility_summaries)

    @staticmethod
    def set_error(run_id: str, error_message: str) -> None:
        with get_db() as db:
            run = db.query(ValidationRun).filter_by(run_id=run_id).first()
            if run:
                run.status        = "error"
                run.error_message = str(error_message)[:2000]

    @staticmethod
    def get_latest_per_slot() -> list[ValidationRun]:
        """Return at most 3 runs — the latest completed run for each slot.
        This is what GET /runs returns to BayCentral.
        """
        result = []
        with get_db() as db:
            for slot in SLOTS:
                run = (
                    db.query(ValidationRun)
                    .filter(ValidationRun.schedule_slot == slot,
                            ValidationRun.status == "done")
                    .order_by(ValidationRun.run_time.desc())
                    .first()
                )
                if run:
                    db.expunge(run)
                    result.append(run)
        return result

    @staticmethod
    def get_by_run_id(run_id: str) -> Optional[ValidationRun]:
        with get_db() as db:
            run = db.query(ValidationRun).filter_by(run_id=run_id).first()
            if run:
                db.expunge(run)
            return run

    @staticmethod
    def get_all_completed(days: int = 90) -> list[ValidationRun]:
        """Return all completed runs within the last `days` days, newest first.
        Used by GET /runs/history for the trend chart.
        """
        cutoff = datetime.utcnow() - timedelta(days=days)
        with get_db() as db:
            runs = (
                db.query(ValidationRun)
                .filter(
                    ValidationRun.status == "done",
                    ValidationRun.run_time >= cutoff,
                )
                .order_by(ValidationRun.run_time.desc())
                .all()
            )
            for r in runs:
                db.expunge(r)
        return runs

    @staticmethod
    def get_last_completed() -> Optional[ValidationRun]:
        with get_db() as db:
            run = (
                db.query(ValidationRun)
                .filter(ValidationRun.status == "done")
                .order_by(ValidationRun.run_time.desc())
                .first()
            )
            if run:
                db.expunge(run)
            return run

    @staticmethod
    def purge_old_runs() -> int:
        """Delete runs older than settings.retention_days.  Returns count deleted."""
        cutoff = datetime.utcnow() - timedelta(days=settings.retention_days)
        with get_db() as db:
            old = (
                db.query(ValidationRun)
                .filter(ValidationRun.run_time < cutoff)
                .all()
            )
            count = len(old)
            for run in old:
                db.delete(run)
        if count:
            logger.info(f"Purged {count} run(s) older than {settings.retention_days} days")
        return count


# ═══════════════════════════════════════════════════════════════════════════════
# ERROR REPOSITORY
# ═══════════════════════════════════════════════════════════════════════════════

class ErrorRepository:

    @staticmethod
    def bulk_insert(run_id: str, violations_df: pd.DataFrame) -> int:
        """Write violations_df rows to facility_errors.
        Returns the count inserted.
        """
        if violations_df.empty:
            return 0

        # Resolve the integer PK for the run
        with get_db() as db:
            run = db.query(ValidationRun).filter_by(run_id=run_id).first()
            if not run:
                logger.error(f"bulk_insert: run_id={run_id} not found in DB")
                return 0
            run_pk = run.id

        # Map DataFrame column names to model columns
        col_map = {
            "_row_number":       "row_number",
            "patient_id":        "patient_id",
            "patient_uid":       "patient_uid",
            "facility_name":     "facility_name",
            "state":             "state",
            "lga_of_residence":  "lga_of_residence",
            "rule_id":           "rule_id",
            "rule_title":        "rule_title",
            "severity":          "severity",
            "failing_field":     "failing_field",
            "current_value":     "current_value",
            "expected_condition":"expected_condition",
        }

        records = []
        for _, row in violations_df.iterrows():
            kwargs = {"run_id_fk": run_pk}
            for df_col, model_col in col_map.items():
                val = row.get(df_col)
                if pd.isna(val) if isinstance(val, float) else val is None:
                    val = None
                else:
                    val = str(val) if model_col not in ("row_number",) else (
                        int(val) if val is not None else None
                    )
                kwargs[model_col] = val
            records.append(FacilityError(**kwargs))

        with get_db() as db:
            db.bulk_save_objects(records)

        logger.info(f"Inserted {len(records)} violation rows for run_id={run_id}")
        return len(records)

    @staticmethod
    def paginate(
        run_id: str,
        page: int = 1,
        page_size: int = 200,
        rule_id: Optional[str] = None,
        facility: Optional[str] = None,
        lga: Optional[str] = None,
        state: Optional[str] = None,
        facilities: Optional[list] = None,
    ) -> dict:
        """Return a page of violations for a run, with optional filters.
        Returns: {total_records, total_pages, violations: list[dict]}
        """
        with get_db() as db:
            run = db.query(ValidationRun).filter_by(run_id=run_id).first()
            if not run:
                return {"total_records": 0, "total_pages": 0, "violations": []}

            q = db.query(FacilityError).filter(FacilityError.run_id_fk == run.id)

            if rule_id:
                q = q.filter(FacilityError.rule_id == rule_id.upper())
            if facility:
                q = q.filter(FacilityError.facility_name.ilike(f"%{facility}%"))
            if lga:
                q = q.filter(FacilityError.lga_of_residence.ilike(f"%{lga}%"))
            if state:
                q = q.filter(FacilityError.state.ilike(state))
            if facilities:
                q = q.filter(FacilityError.facility_name.in_(facilities))

            total = q.count()
            total_pages = max(1, math.ceil(total / page_size))
            offset = (page - 1) * page_size
            rows = q.order_by(FacilityError.id).offset(offset).limit(page_size).all()

            violations = [
                {
                    "row_number":         r.row_number,
                    "patient_id":         r.patient_id or "",
                    "patient_uid":        r.patient_uid or "",
                    "facility_name":      r.facility_name or "",
                    "state":              r.state or "",
                    "lga_of_residence":   r.lga_of_residence or "",
                    "rule_id":            r.rule_id or "",
                    "rule_title":         r.rule_title or "",
                    "severity":           r.severity or "",
                    "failing_field":      r.failing_field or "",
                    "current_value":      r.current_value or "",
                    "expected_condition": r.expected_condition or "",
                }
                for r in rows
            ]

        return {
            "total_records": total,
            "total_pages":   total_pages,
            "violations":    violations,
        }
