# rule registry + execution loop
# app/engine/validator.py
# ─────────────────────────────────────────────────────────────────────────────
# PURPOSE:
#   The rule engine. Sits between loader.py (which produces a clean DataFrame)
#   and reporter.py (which generates the Excel report). It:
#
#     1. Loads rules_config.yaml to know which rules exist and are enabled
#     2. Checks that required columns are present before running each rule
#     3. Runs each enabled rule function from custom_rules.RULE_REGISTRY
#     4. Aggregates all violations into a single DataFrame
#     5. Builds two summary structures:
#          - Rule summary    : one row per rule (violation count, % affected)
#          - Facility summary: violations grouped by facility + state + rule
#     6. Returns a ValidationResult dataclass consumed by the API route
#
# PERFORMANCE:
#   All rule functions are vectorised (no Python loops over rows).
#   On a 39,000-row combined RADET file running all active rules, total
#   validation time is expected to be under 20 seconds on standard hardware.
#   The engine runs rules sequentially and collects results in a list,
#   then calls pd.concat() once — never appending inside a loop.
# ─────────────────────────────────────────────────────────────────────────────

import time
import logging
import uuid
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
import yaml

from app.rules.custom_rules import RULE_REGISTRY
from app.utils.excel_helpers import check_required_columns

logger = logging.getLogger(__name__)

# Path to the YAML config — resolved relative to this file's location
_RULES_CONFIG_PATH = Path(__file__).parent.parent / "rules" / "rules_config.yaml"


# ─────────────────────────────────────────────────────────────────────────────
# OUTPUT DATACLASSES
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class RuleSummary:
    """One row per rule in the Summary sheet and API response."""
    rule_id:           str
    rule_title:        str
    category:          str
    severity:          str
    enabled:           bool
    total_violations:  int
    pct_rows_affected: float   # percentage of total rows that failed this rule
    skipped:           bool    # True if rule was skipped due to missing columns
    skip_reason:       str     # empty string if not skipped


@dataclass
class FacilitySummary:
    """One row per facility × state × rule combination."""
    facility_name:     str
    state:             str     # State (replaces LGA in facility grouping)
    rule_id:           str
    rule_title:        str
    severity:          str
    violation_count:   int
    pct_facility_rows: float   # % of rows FOR THIS FACILITY that failed


@dataclass
class ValidationResult:
    """
    Everything the API route and reporter need.
    Returned by run_validation().
    """
    job_id:                  str
    total_rows:              int
    total_violations:        int
    rules_run:               int
    rules_skipped:           int
    validation_time_seconds: float

    rule_summaries:          list[RuleSummary]
    facility_summaries:      list[FacilitySummary]

    # Full violations DataFrame — all failing rows with standard output schema
    # Columns: _row_number, patient_id, patient_uid, facility_name, lga_of_residence,
    #          rule_id, rule_title, severity, failing_field,
    #          current_value, expected_condition
    violations_df:           pd.DataFrame

    # Warnings from the engine itself (not rule violations)
    # e.g. "Column X missing — Rule R-04 skipped"
    engine_warnings:         list[str]


# ─────────────────────────────────────────────────────────────────────────────
# LOAD RULES CONFIG
# ─────────────────────────────────────────────────────────────────────────────

def load_rules_config(config_path: Path = _RULES_CONFIG_PATH) -> list[dict]:
    """
    Load and return the list of rule definitions from rules_config.yaml.

    Each dict has keys: id, title, category, severity, enabled, columns,
    description, condition, action.

    Raises FileNotFoundError if the YAML file is missing.
    """
    if not config_path.exists():
        raise FileNotFoundError(
            f"rules_config.yaml not found at {config_path}. "
            "Make sure the file exists at app/rules/rules_config.yaml."
        )
    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)

    rules = config.get("rules", [])
    logger.info("Loaded %d rules from config.", len(rules))
    return rules


# ─────────────────────────────────────────────────────────────────────────────
# MAIN VALIDATION FUNCTION
# ─────────────────────────────────────────────────────────────────────────────


def run_validation(
    df: pd.DataFrame,
    enabled_rule_ids: list[str] | None = None,
    reporting_start: pd.Timestamp | None = None,
    reporting_end: pd.Timestamp | None = None,
    config_path: Path = _RULES_CONFIG_PATH,
    on_rule_complete=None,
    aux_dfs: dict[str, pd.DataFrame] | None = None,
) -> ValidationResult:
    """
    Run all enabled validation rules against a cleaned RADET DataFrame.

    Args:
        df:               Cleaned DataFrame from loader.load_radet_file().
                          Must already have canonical column names and
                          normalised date columns.

        enabled_rule_ids: Optional list of rule IDs to run (e.g. ["R-01","R-08"]).
                          If None, runs all rules where enabled=true in YAML.
                          The Flutter sidebar passes this when the user toggles rules.

        reporting_start:  Start of the reporting period (for R-09 new ART check).
                          Defaults to January 1st of the current year.

        reporting_end:    End of the reporting period (for R-09).
                          Defaults to December 31st of the current year.

        config_path:      Path to rules_config.yaml. Override for testing.

    Returns:
        ValidationResult with all violations, summaries, and metadata.
    """
    start_time = time.perf_counter()
    job_id     = str(uuid.uuid4())[:8]
    total_rows = len(df)
    engine_warnings: list[str] = []

    # ── Step 1: Load rule definitions from YAML ───────────────────────────────
    all_rules = load_rules_config(config_path)
    rule_meta: dict[str, dict] = {r["id"]: r for r in all_rules}

    # ── Step 2: Determine which rules to run ─────────────────────────────────
    if enabled_rule_ids is not None:
        unknown = [rid for rid in enabled_rule_ids if rid not in rule_meta]
        if unknown:
            engine_warnings.append(
                f"Unknown rule IDs requested and ignored: {unknown}"
            )
        rules_to_run = [
            rule_meta[rid] for rid in enabled_rule_ids
            if rid in rule_meta
        ]
    else:
        rules_to_run = [r for r in all_rules if r.get("enabled", False)]

    logger.info(
        "Job %s: Running %d rules on %d rows.",
        job_id, len(rules_to_run), total_rows,
    )

    # ── Step 3: Execute each rule ─────────────────────────────────────────────
    violation_frames: list[pd.DataFrame] = []
    rule_summaries:   list[RuleSummary]  = []
    rules_run     = 0
    rules_skipped = 0
    total_rules   = len(rules_to_run)

    for rule_num, rule_def in enumerate(rules_to_run, start=1):
        rule_id    = rule_def["id"]
        rule_title = rule_def["title"]
        severity   = rule_def["severity"]
        category   = rule_def["category"]
        required_cols: list[str] = rule_def.get("columns", [])

        # ── Pre-flight: check required columns exist ──────────────────────────
        present, missing = check_required_columns(df, required_cols)

        if missing:
            skip_reason = f"Required column(s) not found in file: {missing}"
            engine_warnings.append(f"Rule {rule_id} skipped — {skip_reason}")
            rule_summaries.append(RuleSummary(
                rule_id=rule_id, rule_title=rule_title,
                category=category, severity=severity, enabled=True,
                total_violations=0, pct_rows_affected=0.0,
                skipped=True, skip_reason=skip_reason,
            ))
            rules_skipped += 1
            logger.warning("Rule %s skipped: %s", rule_id, skip_reason)
            continue

        # ── Get the rule function from the registry ───────────────────────────
        rule_fn = RULE_REGISTRY.get(rule_id)
        if rule_fn is None:
            skip_reason = (
                f"No implementation found in RULE_REGISTRY for {rule_id}. "
                "Add the function to custom_rules.py."
            )
            engine_warnings.append(f"Rule {rule_id} skipped — {skip_reason}")
            rule_summaries.append(RuleSummary(
                rule_id=rule_id, rule_title=rule_title,
                category=category, severity=severity, enabled=True,
                total_violations=0, pct_rows_affected=0.0,
                skipped=True, skip_reason=skip_reason,
            ))
            rules_skipped += 1
            continue

        # ── Run the rule ──────────────────────────────────────────────────────
        rule_start    = time.perf_counter()
        is_cross_sheet = rule_def.get("cross_sheet", False)
        try:
            violations = rule_fn(df, aux_dfs=aux_dfs) if is_cross_sheet else rule_fn(df)

        except Exception as exc:
            skip_reason = f"Runtime error: {exc}"
            logger.exception("Rule %s raised an exception: %s", rule_id, exc)
            engine_warnings.append(
                f"Rule {rule_id} encountered an error and was skipped — "
                f"{skip_reason}"
            )
            rule_summaries.append(RuleSummary(
                rule_id=rule_id, rule_title=rule_title,
                category=category, severity=severity, enabled=True,
                total_violations=0, pct_rows_affected=0.0,
                skipped=True, skip_reason=skip_reason,
            ))
            rules_skipped += 1
            continue

        rule_elapsed    = time.perf_counter() - rule_start
        violation_count = len(violations)
        pct = round(violation_count / total_rows * 100, 2) if total_rows > 0 else 0.0

        if on_rule_complete is not None:
            on_rule_complete(rule_num, total_rules, rule_title)

        logger.info(
            "  %s: %5d violations  (%.1f%% of rows)  %.1fms",
            rule_id, violation_count, pct, rule_elapsed * 1000,
        )

        if not violations.empty:
            # ── Normalise title + severity from YAML ──────────────────────────
            # Rule functions hardcode their own title and severity inside
            # _build_output. If the YAML is updated (e.g. a title rename or a
            # severity change like Warning → Error), the violations_df would
            # still carry the old values — causing Dashboard / Rule Summary to
            # disagree with By Facility / All Violations in the Excel report.
            #
            # Solution: YAML is the single source of truth. Overwrite both
            # columns here so every downstream consumer (reporter, API) is
            # consistent regardless of what the rule function wrote.
            violations["rule_title"] = rule_title
            violations["severity"]   = severity
            violation_frames.append(violations)

        rule_summaries.append(RuleSummary(
            rule_id=rule_id, rule_title=rule_title,
            category=category, severity=severity, enabled=True,
            total_violations=violation_count,
            pct_rows_affected=pct,
            skipped=False, skip_reason="",
        ))
        rules_run += 1

    # ── Step 4: Combine all violation frames ──────────────────────────────────
    if violation_frames:
        all_violations = pd.concat(violation_frames, ignore_index=True)
    else:
        all_violations = pd.DataFrame(columns=[
            "_row_number", "patient_id", "patient_uid", "facility_name", "lga_of_residence",
            "rule_id", "rule_title", "severity",
            "failing_field", "current_value", "expected_condition",
        ])

    total_violations = len(all_violations)

    # ── Step 5: Build facility summary ────────────────────────────────────────
    facility_summaries = _build_facility_summaries(all_violations, df)

    # ── Step 6: Log final stats ───────────────────────────────────────────────
    elapsed = round(time.perf_counter() - start_time, 2)
    logger.info(
        "Job %s complete — %d violations across %d rules in %.2fs  (%d skipped)",
        job_id, total_violations, rules_run, elapsed, rules_skipped,
    )

    return ValidationResult(
        job_id=job_id,
        total_rows=total_rows,
        total_violations=total_violations,
        rules_run=rules_run,
        rules_skipped=rules_skipped,
        validation_time_seconds=elapsed,
        rule_summaries=rule_summaries,
        facility_summaries=facility_summaries,
        violations_df=all_violations,
        engine_warnings=engine_warnings,
    )


# ─────────────────────────────────────────────────────────────────────────────
# FACILITY SUMMARY BUILDER
# ─────────────────────────────────────────────────────────────────────────────

def _build_facility_summaries(
    violations: pd.DataFrame,
    source_df: pd.DataFrame,
) -> list[FacilitySummary]:
    """
    Group violations by facility_name × state × rule_id.

    pct_facility_rows is the percentage of rows for that specific
    facility+state combination that failed this rule — not the full file's
    row count. This helps facilities understand their own data quality
    relative to their own patient count.

    Args:
        violations: Combined violations DataFrame from all rules.
        source_df:  The original cleaned DataFrame (to get per-facility counts).

    Returns:
        List of FacilitySummary objects sorted by:
          facility_name → violation_count descending
    """
    if violations.empty:
        return []

    # ── Determine state column ────────────────────────────────────────────────
    # Prefer `state` if present; fall back to `lga_of_residence` so existing
    # files without a state column do not crash. Log which is being used.
    if "state" in source_df.columns and "state" in violations.columns:
        geo_col = "state"
    else:
        geo_col = "lga_of_residence"
        logger.warning(
            "_build_facility_summaries: 'state' column not found — "
            "falling back to 'lga_of_residence' for facility grouping."
        )

    # ── Count total rows per facility + geo in source data ───────────────────
    facility_totals = (
        source_df
        .groupby(["facility_name", geo_col])
        .size()
        .reset_index(name="facility_total")
    )

    # ── Count violations per facility + geo + rule ────────────────────────────
    grouped = (
        violations
        .groupby(["facility_name", geo_col, "rule_id", "rule_title", "severity"])
        .size()
        .reset_index(name="violation_count")
    )

    # ── Merge in facility totals to compute percentage ────────────────────────
    grouped = grouped.merge(
        facility_totals,
        on=["facility_name", geo_col],
        how="left",
    )
    grouped["pct_facility_rows"] = (
        grouped["violation_count"] / grouped["facility_total"] * 100
    ).round(2)

    # ── Sort: facility alphabetically, then most violations first ────────────
    grouped = grouped.sort_values(
        ["facility_name", "violation_count"],
        ascending=[True, False],
    )

    return [
        FacilitySummary(
            facility_name=row.facility_name,
            state=getattr(row, geo_col),
            rule_id=row.rule_id,
            rule_title=row.rule_title,
            severity=row.severity,
            violation_count=int(row.violation_count),
            pct_facility_rows=float(row.pct_facility_rows),
        )
        for row in grouped.itertuples()
    ]


# ─────────────────────────────────────────────────────────────────────────────
# CONVENIENCE: get_all_rules
#   Called by GET /rules endpoint to return the rule list to Flutter.
# ─────────────────────────────────────────────────────────────────────────────

def get_all_rules(config_path: Path = _RULES_CONFIG_PATH) -> list[dict]:
    """
    Return all rule definitions from YAML for the /rules API endpoint.
    Flutter uses this to build the sidebar rule toggle list.
    """
    return load_rules_config(config_path)


# ─────────────────────────────────────────────────────────────────────────────
# QUICK DIAGNOSTIC — run directly to test the full engine
#
#   Usage from project root:
#       python -m app.engine.validator path/to/radet.xlsx
#
#   Optional flags:
#       --rules R-01,R-08,R-15     run only these rules
#       --start 2025-10-01         reporting period start (for R-09)
#       --end   2025-12-31         reporting period end
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    import argparse

    logging.basicConfig(
        level=logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    parser = argparse.ArgumentParser(
        description="RADET Validation Engine — diagnostic run"
    )
    parser.add_argument("file", help="Path to the RADET .xlsx file")
    parser.add_argument(
        "--rules", default=None,
        help="Comma-separated rule IDs to run (default: all enabled)"
    )
    parser.add_argument("--start", default=None, help="Reporting start date YYYY-MM-DD")
    parser.add_argument("--end",   default=None, help="Reporting end date YYYY-MM-DD")
    args = parser.parse_args()

    from app.engine.loader import load_radet_file

    print(f"\n{'='*65}")
    print(f"  RADET Validation Engine — Diagnostic")
    print(f"  File: {args.file}")
    print(f"{'='*65}\n")

    load_result = load_radet_file(args.file)
    print(f"  Loaded: {load_result.total_rows:,} rows  |  "
          f"Column coverage: {load_result.column_coverage_pct}%\n")

    rule_ids  = [r.strip() for r in args.rules.split(",")] if args.rules else None
    rep_start = pd.Timestamp(args.start) if args.start else None
    rep_end   = pd.Timestamp(args.end)   if args.end   else None

    result = run_validation(
        load_result.df,
        enabled_rule_ids=rule_ids,
        reporting_start=rep_start,
        reporting_end=rep_end,
    )

    print(f"  Job ID              : {result.job_id}")
    print(f"  Total rows          : {result.total_rows:,}")
    print(f"  Total violations    : {result.total_violations:,}")
    print(f"  Rules run           : {result.rules_run}")
    print(f"  Rules skipped       : {result.rules_skipped}")
    print(f"  Validation time     : {result.validation_time_seconds}s")

    if result.engine_warnings:
        print(f"\n  Engine warnings:")
        for w in result.engine_warnings:
            print(f"    ⚠  {w}")

    print(f"\n  {'Rule':<6} {'Violations':>10} {'% Rows':>7}  {'Skipped'}  Title")
    print(f"  {'-'*75}")
    for rs in result.rule_summaries:
        skipped_tag = f"[SKIPPED: {rs.skip_reason[:30]}]" if rs.skipped else ""
        print(
            f"  {rs.rule_id:<6} {rs.total_violations:>10,} "
            f"{rs.pct_rows_affected:>6.1f}%  "
            f"{'yes' if rs.skipped else 'no ':>7}  "
            f"{rs.rule_title[:40]} {skipped_tag}"
        )

    print(f"\n  Facility / State Summary (top 10):")
    print(f"  {'Facility':<30} {'State':<15} {'Rule':<6} {'Count':>6} {'%Fac':>6}")
    print(f"  {'-'*70}")
    for fs in result.facility_summaries[:10]:
        print(
            f"  {fs.facility_name[:28]:<30} "
            f"{fs.state[:13]:<15} "
            f"{fs.rule_id:<6} "
            f"{fs.violation_count:>6} "
            f"{fs.pct_facility_rows:>5.1f}%"
        )
    print()
