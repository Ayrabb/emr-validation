# app/scheduler/pipeline.py
# ─────────────────────────────────────────────────────────────────────────────
# run_pipeline(schedule_slot) is called by APScheduler at 06:00, 12:00, 18:00.
#
# Flow:
#   1. Create ValidationRun record (status=pending)
#   2. Download today's RADET file from Azure Blob
#   3. Call ValidationService.run() — same engine the desktop app uses
#   4. Generate Excel report to disk (same reporter.py)
#   5. Persist violations to SQLite (facility_errors table)
#   6. Persist summaries + metadata to ValidationRun record (status=done)
#   7. Populate _job_store so in-memory endpoints work immediately
#
# If any step raises, the run is marked status=error and the exception is
# logged — the scheduler continues and will fire again at the next slot.
# ─────────────────────────────────────────────────────────────────────────────

import logging
import uuid
from datetime import date, datetime
from pathlib import Path

from app.blob.connector import blob_connector
from app.core.config import settings
from app.db.repository import ErrorRepository, RunRepository
from app.engine.reporter import generate_report
from app.services.validation_service import ValidationService

logger = logging.getLogger(__name__)

# Shared with validation_routes.py — see that module for the definition.
# Imported lazily to avoid circular imports.
_job_store_ref: dict | None = None


def set_job_store(store: dict) -> None:
    """Called from main.py so the pipeline can populate the in-memory store."""
    global _job_store_ref
    _job_store_ref = store


def run_pipeline(schedule_slot: str) -> None:
    """Full validation pipeline for one schedule slot.
    schedule_slot: "06:00" | "12:00" | "18:00"
    """
    today    = date.today()
    run_id   = str(uuid.uuid4())[:8]

    logger.info(
        f"[Pipeline] STARTING  slot={schedule_slot}  "
        f"date={today}  run_id={run_id}"
    )

    # ── 1. Create DB record ───────────────────────────────────────────────────
    RunRepository.create(run_id=run_id, schedule_slot=schedule_slot, blob_date=today)
    RunRepository.set_running(run_id)

    if _job_store_ref is not None:
        _job_store_ref[run_id] = {
            "status":   "processing",
            "progress": 5,
            "message":  "Downloading RADET file from Azure Blob…",
            "result":   None,
        }

    try:
        # ── 2. Download blob ──────────────────────────────────────────────────
        _update(run_id, 10, "Downloading RADET file from Azure Blob…")
        file_bytes = blob_connector.download(for_date=today)
        filename   = f"radet_{today}.xlsx"
        size_mb    = len(file_bytes) / 1_048_576
        all_rules  = ValidationService.get_rules()
        rule_count = len(all_rules)
        _update(run_id, 20, f"Downloaded {size_mb:.1f} MB — running {rule_count} rules…")

        # ── 3. Validate ───────────────────────────────────────────────────────
        def _progress(pct: int, msg: str) -> None:
            _update(run_id, 20 + int(pct * 0.6), msg)   # map 0-100 → 20-80%

        result = ValidationService.run(
            file_bytes=file_bytes,
            filename=filename,
            enabled_rule_ids=None,      # always run all enabled rules
            reporting_start=None,       # auto from file date
            reporting_end=None,
            on_progress=_progress,
        )

        # ── 4. Generate report to disk ────────────────────────────────────────
        _update(run_id, 85, "Generating Excel report…")
        report_dir  = Path(settings.report_dir)
        report_dir.mkdir(parents=True, exist_ok=True)
        report_path = report_dir / f"EMR_Validation_Report_{run_id}.xlsx"
        generate_report(result, output_path=report_path)
        logger.info(f"[Pipeline] Report saved: {report_path}")

        # ── 5. Persist violations to SQLite ───────────────────────────────────
        _update(run_id, 92, "Saving results to database…")
        ErrorRepository.bulk_insert(run_id=run_id, violations_df=result.violations_df)

        # ── 6. Persist run metadata ───────────────────────────────────────────
        rule_summaries_json = [
            {
                "rule_id":           rs.rule_id,
                "rule_title":        rs.rule_title,
                "category":          rs.category,
                "severity":          rs.severity,
                "total_violations":  rs.total_violations,
                "pct_rows_affected": rs.pct_rows_affected,
                "skipped":           rs.skipped,
                "skip_reason":       rs.skip_reason,
            }
            for rs in result.rule_summaries
        ]
        facility_summaries_json = [
            {
                "facility_name":     fs.facility_name,
                "state":             fs.state,
                "rule_id":           fs.rule_id,
                "rule_title":        fs.rule_title,
                "severity":          fs.severity,
                "violation_count":   fs.violation_count,
                "pct_facility_rows": fs.pct_facility_rows,
            }
            for fs in result.facility_summaries
        ]

        RunRepository.set_done(
            run_id=run_id,
            total_rows=result.total_rows,
            total_violations=result.total_violations,
            rules_run=result.rules_run,
            validation_time_seconds=result.validation_time_seconds,
            report_path=str(report_path),
            rule_summaries=rule_summaries_json,
            facility_summaries=facility_summaries_json,
        )

        # ── 7. Populate in-memory job_store for immediate API access ──────────
        if _job_store_ref is not None:
            _job_store_ref[run_id] = {
                "status":   "done",
                "progress": 100,
                "message":  (
                    f"Complete — {result.total_violations:,} violations "
                    f"across {result.rules_run} rules "
                    f"in {result.validation_time_seconds:.1f}s"
                ),
                "result":   result,
            }
            # Also store under the internal result.job_id so the report
            # endpoint always resolves correctly regardless of which ID is used
            if result.job_id != run_id:
                _job_store_ref[result.job_id] = _job_store_ref[run_id]

            # Evict old done/error entries — keep at most 20 to prevent
            # memory accumulation from DataFrame objects held in "result".
            _evict_job_store(_job_store_ref, keep=20, exclude={run_id, result.job_id})

        # ── 8. Purge old DB runs + their report files ─────────────────────────
        try:
            _purge_old_runs_and_reports(report_dir)
        except Exception as purge_exc:
            logger.warning(f"[Pipeline] Purge old runs failed (non-fatal): {purge_exc}")

        logger.info(
            f"[Pipeline] DONE  run_id={run_id}  slot={schedule_slot}  "
            f"violations={result.total_violations}  "
            f"time={result.validation_time_seconds:.1f}s"
        )

    except Exception as exc:
        logger.exception(f"[Pipeline] FAILED  run_id={run_id}  slot={schedule_slot}: {exc}")
        RunRepository.set_error(run_id=run_id, error_message=str(exc))
        if _job_store_ref is not None:
            _job_store_ref[run_id] = {
                "status":   "error",
                "progress": 0,
                "message":  str(exc),
                "result":   None,
            }


def _update(run_id: str, progress: int, message: str) -> None:
    if _job_store_ref and run_id in _job_store_ref:
        _job_store_ref[run_id]["progress"] = progress
        _job_store_ref[run_id]["message"]  = message


def _evict_job_store(store: dict, keep: int, exclude: set) -> None:
    """Remove oldest done/error entries from store, keeping at most `keep` total.
    Never removes entries in `exclude` (the current run's IDs).
    Scheduled runs are persisted to DB/disk, so eviction is safe.
    """
    evictable = [
        k for k, v in store.items()
        if k not in exclude and v.get("status") in ("done", "error")
    ]
    excess = len(store) - keep
    if excess > 0:
        for k in evictable[:excess]:
            del store[k]


def _purge_old_runs_and_reports(report_dir: Path) -> None:
    """Delete DB runs older than retention_days and their report files."""
    import time
    deleted_runs = RunRepository.purge_old_runs()
    # Clean up any report files older than retention_days regardless of DB purge count
    retention_days = getattr(settings, "retention_days", 30)
    for f in report_dir.glob("EMR_Validation_Report_*.xlsx"):
        age_days = (time.time() - f.stat().st_mtime) / 86400
        if age_days > retention_days:
            try:
                f.unlink()
                logger.info(f"[Pipeline] Deleted old report: {f.name}")
            except Exception:
                pass
    if deleted_runs:
        logger.info(f"[Pipeline] Purged {deleted_runs} old DB run(s)")
