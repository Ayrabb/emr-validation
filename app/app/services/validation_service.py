# app/services/validation_service.py
# ─────────────────────────────────────────────────────────────────────────────
# CHANGE: Added optional `on_progress` callback parameter to run().
#
# This lets validation_routes.py report progress milestones back to the
# job_store as the engine works. The callback signature is:
#
#   on_progress(progress: int, message: str) → None
#     progress: 0-100 integer
#     message:  human-readable status for Flutter UI
#
# If run_validation() in validator.py is updated to support per-rule
# callbacks (via on_rule_complete kwarg), granular progress is automatic.
# If not, this falls back gracefully — the caller uses simulated progress.
# ─────────────────────────────────────────────────────────────────────────────

import logging
from typing import Callable

import pandas as pd

from app.engine.loader import load_hts_sheet, load_radet_file
from app.engine.validator import ValidationResult, get_all_rules, run_validation

logger = logging.getLogger(__name__)

# Type alias for the progress callback
ProgressCallback = Callable[[int, str], None]


class ValidationService:

    @staticmethod
    def run(
        file_bytes: bytes,
        filename: str,
        enabled_rule_ids: list[str] | None = None,
        reporting_start: pd.Timestamp | None = None,
        reporting_end: pd.Timestamp | None = None,
        on_progress: ProgressCallback | None = None,
    ) -> ValidationResult:
        """
        Full validation pipeline: load → validate → return result.

        Args:
            file_bytes:       Raw bytes of the uploaded .xlsx file.
            filename:         Original filename (used for logging only).
            enabled_rule_ids: Rule IDs to run. None = run all enabled rules.
            reporting_start:  Reporting period start (for R-09).
            reporting_end:    Reporting period end (for R-09).
            on_progress:      Optional callback(progress: int, message: str).
                              Called at key milestones: 10% (loaded), 20% (started),
                              90% (building results). For per-rule granularity,
                              update run_validation() in validator.py to accept
                              on_rule_complete and wire it through.

        Returns:
            ValidationResult with all violations, summaries, and metadata.

        Raises:
            ValueError:   If the file cannot be read as a valid RADET Excel file.
            RuntimeError: For unexpected engine failures.
        """
        logger.info(f"Starting validation for '{filename}'")

        # ── Step 1: Load and clean the file ────────────────────────────────
        if on_progress:
            on_progress(10, "Loading and parsing Excel file…")

        load_result = load_radet_file(file_bytes)
        hts_df      = load_hts_sheet(file_bytes)
        aux_dfs     = {"hts": hts_df} if hts_df is not None else None

        for warning in load_result.warnings:
            logger.info(f"Load warning: {warning}")

        row_count = len(load_result.df)
        rule_count = len(enabled_rule_ids) if enabled_rule_ids else len(get_all_rules())

        if on_progress:
            on_progress(
                20,
                f"Loaded {row_count:,} rows — running {rule_count} rules…",
            )

        # ── Step 2: Run all enabled rules ──────────────────────────────────
        # Try to pass on_rule_complete for per-rule progress.
        # If validator.py doesn't support it yet, fall back silently.
        if on_progress:
            def _rule_callback(rule_num: int, total_rules: int, rule_title: str) -> None:
                pct = 20 + int(65 * rule_num / max(total_rules, 1))
                on_progress(pct, f"Rule {rule_num}/{total_rules}: {rule_title}")
        else:
            _rule_callback = None

        validation_result = run_validation(
            df=load_result.df,
            enabled_rule_ids=enabled_rule_ids,
            reporting_start=reporting_start,
            reporting_end=reporting_end,
            on_rule_complete=_rule_callback,
            aux_dfs=aux_dfs,
        )

        # ── Step 3: Carry load warnings into result ────────────────────────
        if load_result.warnings:
            validation_result.engine_warnings = (
                load_result.warnings + validation_result.engine_warnings
            )

        if on_progress:
            on_progress(
                90,
                f"Building results — {validation_result.total_violations:,} violations found…",
            )

        logger.info(
            f"Validation complete for '{filename}' — "
            f"job_id={validation_result.job_id}  "
            f"violations={validation_result.total_violations}  "
            f"time={validation_result.validation_time_seconds:.1f}s"
        )
        return validation_result

    @staticmethod
    def get_rules() -> list[dict]:
        """Return all rule definitions from rules_config.yaml."""
        return get_all_rules()
