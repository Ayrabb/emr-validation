# app/engine/reporter.py
# ─────────────────────────────────────────────────────────────────────────────
# PURPOSE:
#   Generates a formatted 4-sheet Excel validation report using XlsxWriter.
#   Called lazily — only when the user clicks "Download Report" in Flutter,
#   not on every validation run.
#
# SHEETS:
#   1. Dashboard    — metadata header: file info, totals, date, quick stats
#   2. Rule Summary — one row per rule: ID, title, severity, count, %
#   3. By Facility  — one row per facility × rule: facility, LGA, rule, count
#   4. All Violations — every individual violation row with full details
#
# FIX — "Removed Records: Formula from sheet4.xml":
#   The previous version used ws.write() throughout, which lets xlsxwriter
#   auto-detect cell types. When a string value like "4" (EAC sessions)
#   or "4.0" is passed, xlsxwriter silently coerces it to a number and writes
#   it without a type marker. Excel then misreads the cell as a formula token.
#
#   Fix: use ws.write_string() for every text cell and ws.write_number() for
#   every numeric cell — no auto-detection, no coercion, no corruption.
#
#   Also set three Workbook constructor flags to disable all auto-detection:
#     strings_to_numbers:  False
#     strings_to_formulas: False
#     strings_to_urls:     False
# ─────────────────────────────────────────────────────────────────────────────

import logging
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import pandas as pd
import xlsxwriter

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# SAFE TYPE HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _safe_str(value) -> str:
    """Convert any value to a clean string safe for write_string()."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value)


def _safe_int(value) -> int:
    """Convert to int safely, returning 0 on failure."""
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


def _safe_float(value) -> float:
    """Convert to float safely, returning 0.0 on failure."""
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


# ─────────────────────────────────────────────────────────────────────────────
# MAIN ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

def generate_report(validation_result, output_path: Path) -> Path:
    """
    Generate a formatted Excel report from a completed validation run.

    Args:
        validation_result: ValidationResult from run_validation().
        output_path:       Where to save the .xlsx file.

    Returns:
        The output_path (for convenience).
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    logger.info(f"Generating report: {output_path}")

    workbook = xlsxwriter.Workbook(
        str(output_path),
        {
            "strings_to_urls":     False,  # no hyperlink auto-detection
            "strings_to_numbers":  False,  # no string→number coercion
            "strings_to_formulas": False,  # no string→formula coercion
            "nan_inf_to_errors":   True,   # NaN/Inf → Excel error, not crash
        },
    )

    fmt = _create_formats(workbook)

    _write_dashboard(workbook, fmt, validation_result)
    _write_rule_summary(workbook, fmt, validation_result)
    _write_facility_summary(workbook, fmt, validation_result)
    _write_all_violations(workbook, fmt, validation_result)

    workbook.close()
    size_kb = output_path.stat().st_size / 1024
    logger.info(f"Report saved: {output_path}  ({size_kb:.1f} KB)")
    return output_path


# ─────────────────────────────────────────────────────────────────────────────
# FORMAT FACTORY
# ─────────────────────────────────────────────────────────────────────────────

def _create_formats(workbook: xlsxwriter.Workbook) -> dict:
    return {
        "title": workbook.add_format({
            "bold": True, "font_size": 16, "font_color": "#1B3A5C",
            "bottom": 2, "bottom_color": "#1B3A5C",
        }),
        "subtitle": workbook.add_format({
            "bold": True, "font_size": 11, "font_color": "#555555",
        }),
        "header": workbook.add_format({
            "bold": True, "font_size": 10,
            "bg_color": "#1B3A5C", "font_color": "#FFFFFF",
            "border": 1, "border_color": "#CCCCCC",
            "text_wrap": True, "valign": "vcenter",
        }),
        "text": workbook.add_format({
            "font_size": 10, "border": 1, "border_color": "#E0E0E0",
            "text_wrap": True, "valign": "vcenter",
        }),
        "number": workbook.add_format({
            "font_size": 10, "border": 1, "border_color": "#E0E0E0",
            "num_format": "#,##0", "valign": "vcenter",
        }),
        "percent": workbook.add_format({
            "font_size": 10, "border": 1, "border_color": "#E0E0E0",
            "num_format": "0.0%", "valign": "vcenter",
        }),
        "error": workbook.add_format({
            "font_size": 10, "border": 1, "border_color": "#E0E0E0",
            "bg_color": "#FDE8E8", "font_color": "#B91C1C",
            "bold": True, "valign": "vcenter",
        }),
        "warning": workbook.add_format({
            "font_size": 10, "border": 1, "border_color": "#E0E0E0",
            "bg_color": "#FEF3C7", "font_color": "#92400E",
            "bold": True, "valign": "vcenter",
        }),
        "metric_label": workbook.add_format({
            "font_size": 11, "bold": True, "font_color": "#374151",
            "right": 1, "right_color": "#E0E0E0", "valign": "vcenter",
        }),
        "metric_value_text": workbook.add_format({
            "font_size": 11, "font_color": "#1B3A5C", "valign": "vcenter",
        }),
        "text_alt": workbook.add_format({
            "font_size": 10, "border": 1, "border_color": "#E0E0E0",
            "bg_color": "#F9FAFB", "text_wrap": True, "valign": "vcenter",
        }),
        "number_alt": workbook.add_format({
            "font_size": 10, "border": 1, "border_color": "#E0E0E0",
            "bg_color": "#F9FAFB", "num_format": "#,##0", "valign": "vcenter",
        }),
        "percent_alt": workbook.add_format({
            "font_size": 10, "border": 1, "border_color": "#E0E0E0",
            "bg_color": "#F9FAFB", "num_format": "0.0%", "valign": "vcenter",
        }),
        "error_alt": workbook.add_format({
            "font_size": 10, "border": 1, "border_color": "#E0E0E0",
            "bg_color": "#FDE8E8", "font_color": "#B91C1C",
            "bold": True, "valign": "vcenter",
        }),
        "warning_alt": workbook.add_format({
            "font_size": 10, "border": 1, "border_color": "#E0E0E0",
            "bg_color": "#FEF3C7", "font_color": "#92400E",
            "bold": True, "valign": "vcenter",
        }),
    }


def _alt(fmt: dict, key: str, is_alt: bool):
    """Return alternating-row variant of a format key if it exists."""
    alt_key = f"{key}_alt"
    return fmt[alt_key] if (is_alt and alt_key in fmt) else fmt[key]


def _severity_fmt(fmt: dict, severity: str, alt: bool = False):
    """Return error (red) format for Service Gap severities, warning (yellow) for Data Quality."""
    key = "error" if "service gap" in severity.strip().lower() else "warning"
    return _alt(fmt, key, alt)


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 1: DASHBOARD
# ─────────────────────────────────────────────────────────────────────────────

def _write_dashboard(workbook, fmt, result) -> None:
    ws = workbook.add_worksheet("Dashboard")
    ws.hide_gridlines(2)
    ws.set_tab_color("#1B3A5C")
    ws.set_column("A:A", 3)       # indent
    ws.set_column("B:B", 9)       # Rule ID / metric labels
    ws.set_column("C:C", 34)      # Rule Title / metric values
    ws.set_column("D:D", 22)      # Severity (rules)
    ws.set_column("E:E", 26)      # Flags / severity breakdown labels
    ws.set_column("F:F", 12)      # % Rows / severity breakdown counts
    ws.set_column("G:G", 12)      # (gap) / severity breakdown %
    ws.set_column("H:H", 20)      # State
    ws.set_column("I:I", 14)      # Service Gap
    ws.set_column("J:J", 14)      # SG/DQ
    ws.set_column("K:K", 14)      # Data Quality
    ws.set_column("L:L", 14)      # Total Flags
    ws.set_column("M:M", 12)      # % of Total

    row = 1
    ws.write_string(row, 1, "EMR Validation Report", fmt["title"])
    row += 1
    ws.write_string(row, 1, "HIV Programme Data Quality Assessment", fmt["subtitle"])
    row += 2

    metrics = [
        ("Job ID",          _safe_str(result.job_id)),
        ("Generated",       datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Total Data Rows", f"{result.total_rows:,}"),
        ("Total Flags",     f"{result.total_violations:,}"),
        ("Rules Executed",  str(result.rules_run)),
        ("Rules Skipped",   str(result.rules_skipped)),
        ("Validation Time", f"{result.validation_time_seconds}s"),
    ]
    for label, value in metrics:
        ws.write_string(row, 1, label, fmt["metric_label"])
        ws.write_string(row, 2, value, fmt["metric_value_text"])
        row += 1

    # ── Severity Breakdown (right side, with % of total flags) ───────────
    severity_counts: dict = defaultdict(int)
    for rs in result.rule_summaries:
        if not rs.skipped:
            severity_counts[rs.severity.strip()] += rs.total_violations

    _SEVERITY_ORDER = ["Service Gap", "Service Gap/Data Quality", "Data Quality"]
    known   = [s for s in _SEVERITY_ORDER if s in severity_counts]
    unknown = sorted(s for s in severity_counts if s not in _SEVERITY_ORDER)
    ordered_severities = known + unknown

    total_flags = result.total_violations
    flag_rate = (
        total_flags / result.total_rows * 100
        if result.total_rows > 0 else 0.0
    )

    right_row = 5
    ws.write_string(right_row, 4, "Severity Breakdown",  fmt["subtitle"])
    ws.write_string(right_row, 6, "% of Total Flags",    fmt["subtitle"])
    right_row += 1
    for sev in ordered_severities:
        count   = severity_counts[sev]
        sev_pct = count / total_flags if total_flags > 0 else 0.0
        ws.write_string(right_row, 4, sev,     _severity_fmt(fmt, sev))
        ws.write_number(right_row, 5, count,   fmt["number"])
        ws.write_number(right_row, 6, sev_pct, fmt["percent"])
        right_row += 1
    ws.write_string(right_row, 4, "Flag Rate",           fmt["metric_label"])
    ws.write_string(right_row, 5, f"{flag_rate:.1f}%",   fmt["metric_value_text"])
    right_row += 2

    if result.engine_warnings:
        row += 2
        ws.write_string(row, 1, "Engine Warnings", fmt["subtitle"])
        row += 1
        for w in result.engine_warnings:
            ws.write_string(row, 1, f"  - {_safe_str(w)}", fmt["text"])
            row += 1

    table_start = max(row, right_row) + 2

    _SEV_COLS    = ["Service Gap", "Service Gap/Data Quality", "Data Quality"]
    _SEV_HEADERS = ["Service Gap", "SG/DQ", "Data Quality"]

    # ── LEFT: Top Offending Rules (cols B–F, indices 1–5) ────────────────
    rr = table_start
    ws.write_string(rr, 1, "Top Offending Rules", fmt["subtitle"])
    rr += 1

    top_rules = sorted(
        [rs for rs in result.rule_summaries if rs.total_violations > 0],
        key=lambda x: x.total_violations,
        reverse=True,
    )[:10]

    if top_rules:
        for c, h in enumerate(["Rule", "Title", "Severity", "Flags", "% Rows"]):
            ws.write_string(rr, 1 + c, h, fmt["header"])
        rr += 1
        for i, rs in enumerate(top_rules):
            alt = i % 2 == 1
            ws.write_string(rr, 1, _safe_str(rs.rule_id),    _alt(fmt, "text", alt))
            ws.write_string(rr, 2, _safe_str(rs.rule_title), _alt(fmt, "text", alt))
            ws.write_string(rr, 3, _safe_str(rs.severity),   _severity_fmt(fmt, rs.severity, alt))
            ws.write_number(rr, 4, _safe_int(rs.total_violations),          _alt(fmt, "number", alt))
            ws.write_number(rr, 5, _safe_float(rs.pct_rows_affected) / 100, _alt(fmt, "percent", alt))
            rr += 1

    # ── RIGHT: Flags by State (cols H–M, indices 7–12) ───────────────────
    # Column order: State | Service Gap | SG/DQ | Data Quality | Total Flags | % of Total
    sr = table_start
    ws.write_string(sr, 7, "Flags by State", fmt["subtitle"])
    sr += 1

    vdf = result.violations_df
    if not vdf.empty and "state" in vdf.columns:
        for c, h in enumerate(["State"] + _SEV_HEADERS + ["Total Flags", "% of Total"]):
            ws.write_string(sr, 7 + c, h, fmt["header"])
        sr += 1

        state_totals = vdf.groupby("state").size().sort_values(ascending=False)
        state_sev_df = (
            vdf.groupby(["state", "severity"])
            .size()
            .unstack(fill_value=0)
        )

        for i, (state_name, count) in enumerate(state_totals.items()):
            alt = i % 2 == 1
            pct = count / total_flags if total_flags > 0 else 0.0
            ws.write_string(sr, 7, _safe_str(state_name), _alt(fmt, "text", alt))
            for c_off, sev in enumerate(_SEV_COLS):
                sev_count = (
                    int(state_sev_df.loc[state_name, sev])
                    if state_name in state_sev_df.index and sev in state_sev_df.columns
                    else 0
                )
                ws.write_number(sr, 8 + c_off, sev_count, _severity_fmt(fmt, sev, alt))
            ws.write_number(sr, 11, _safe_int(count), _alt(fmt, "number", alt))
            ws.write_number(sr, 12, pct,               _alt(fmt, "percent", alt))
            sr += 1

        # Totals row
        ws.write_string(sr, 7, "TOTAL", fmt["metric_label"])
        for c_off, sev in enumerate(_SEV_COLS):
            sev_total = int(state_sev_df[sev].sum()) if sev in state_sev_df.columns else 0
            ws.write_number(sr, 8 + c_off, sev_total, _severity_fmt(fmt, sev))
        ws.write_number(sr, 11, _safe_int(total_flags), fmt["number"])
        ws.write_number(sr, 12, 1.0,                    fmt["percent"])
        sr += 1
    else:
        ws.write_string(sr, 7, "No state data available.", fmt["text"])
        sr += 1


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 2: RULE SUMMARY
# ─────────────────────────────────────────────────────────────────────────────

def _write_rule_summary(workbook, fmt, result) -> None:
    ws = workbook.add_worksheet("Rule Summary")
    ws.set_tab_color("#047857")

    headers = ["Rule ID", "Title", "Category", "Severity",
               "Violations", "% Rows Affected", "Skipped", "Skip Reason"]
    widths  = [8, 35, 20, 20, 12, 16, 9, 35]

    for c, (h, w) in enumerate(zip(headers, widths)):
        ws.set_column(c, c, w)
        ws.write_string(0, c, h, fmt["header"])

    ws.freeze_panes(1, 0)

    n = len(result.rule_summaries)
    if n > 0:
        ws.autofilter(0, 0, n, len(headers) - 1)

    for i, rs in enumerate(result.rule_summaries):
        r   = i + 1
        alt = i % 2 == 1
        ws.write_string(r, 0, _safe_str(rs.rule_id),    _alt(fmt, "text", alt))
        ws.write_string(r, 1, _safe_str(rs.rule_title), _alt(fmt, "text", alt))
        ws.write_string(r, 2, _safe_str(rs.category),   _alt(fmt, "text", alt))
        ws.write_string(r, 3, _safe_str(rs.severity),   _severity_fmt(fmt, rs.severity, alt))
        ws.write_number(r, 4, _safe_int(rs.total_violations),           _alt(fmt, "number", alt))
        ws.write_number(r, 5, _safe_float(rs.pct_rows_affected) / 100,  _alt(fmt, "percent", alt))
        ws.write_string(r, 6, "Yes" if rs.skipped else "No",            _alt(fmt, "text", alt))
        ws.write_string(r, 7, _safe_str(rs.skip_reason),                _alt(fmt, "text", alt))


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 3: BY FACILITY
# ─────────────────────────────────────────────────────────────────────────────

def _write_facility_summary(workbook, fmt, result) -> None:
    ws = workbook.add_worksheet("By Facility")
    ws.set_tab_color("#D97706")

    headers = ["Facility", "State", "Rule ID", "Rule Title",
               "Severity", "Violations", "% Facility Rows"]
    widths  = [30, 18, 8, 35, 20, 12, 16]

    for c, (h, w) in enumerate(zip(headers, widths)):
        ws.set_column(c, c, w)
        ws.write_string(0, c, h, fmt["header"])

    ws.freeze_panes(1, 0)

    n = len(result.facility_summaries)
    if n > 0:
        ws.autofilter(0, 0, n, len(headers) - 1)

    for i, fs in enumerate(result.facility_summaries):
        r   = i + 1
        alt = i % 2 == 1
        ws.write_string(r, 0, _safe_str(fs.facility_name),    _alt(fmt, "text", alt))
        ws.write_string(r, 1, _safe_str(fs.state),            _alt(fmt, "text", alt))
        ws.write_string(r, 2, _safe_str(fs.rule_id),          _alt(fmt, "text", alt))
        ws.write_string(r, 3, _safe_str(fs.rule_title),       _alt(fmt, "text", alt))
        ws.write_string(r, 4, _safe_str(fs.severity),         _severity_fmt(fmt, fs.severity, alt))
        ws.write_number(r, 5, _safe_int(fs.violation_count),              _alt(fmt, "number", alt))
        ws.write_number(r, 6, _safe_float(fs.pct_facility_rows) / 100,    _alt(fmt, "percent", alt))


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 4: ALL VIOLATIONS
# ─────────────────────────────────────────────────────────────────────────────

def _write_all_violations(workbook, fmt, result) -> None:
    ws = workbook.add_worksheet("All Violations")
    ws.set_tab_color("#DC2626")

    headers = [
        "Row #", "Patient ID", "Patient UID", "Facility", "LGA",
        "Rule ID", "Rule Title", "Severity",
        "Failing Field", "Current Value", "Expected Condition",
    ]
    widths = [7, 16, 22, 28, 16, 8, 30, 20, 18, 22, 35]

    for c, (h, w) in enumerate(zip(headers, widths)):
        ws.set_column(c, c, w)
        ws.write_string(0, c, h, fmt["header"])

    ws.freeze_panes(1, 0)

    df = result.violations_df
    if df.empty:
        ws.write_string(1, 0, "No violations found.", fmt["text"])
        return

    df = df.sort_values(
        ["facility_name", "rule_id", "_row_number"],
        ascending=True,
    ).reset_index(drop=True)

    total_rows = len(df)
    # autofilter over the exact written range — never use an estimate
    ws.autofilter(0, 0, total_rows, len(headers) - 1)

    for i in range(total_rows):
        r   = i + 1
        alt = i % 2 == 1
        row = df.iloc[i]

        severity = _safe_str(row.get("severity", ""))

        # Row number — numeric
        ws.write_number(r, 0, _safe_int(row.get("_row_number", 0)),
                        _alt(fmt, "number", alt))

        # All text columns — explicit write_string, never write()
        ws.write_string(r,  1, _safe_str(row.get("patient_id",        "")), _alt(fmt, "text", alt))
        ws.write_string(r,  2, _safe_str(row.get("patient_uid",       "")), _alt(fmt, "text", alt))
        ws.write_string(r,  3, _safe_str(row.get("facility_name",     "")), _alt(fmt, "text", alt))
        ws.write_string(r,  4, _safe_str(row.get("lga_of_residence",  "")), _alt(fmt, "text", alt))
        ws.write_string(r,  5, _safe_str(row.get("rule_id",           "")), _alt(fmt, "text", alt))
        ws.write_string(r,  6, _safe_str(row.get("rule_title",        "")), _alt(fmt, "text", alt))
        ws.write_string(r,  7, severity,                                     _severity_fmt(fmt, severity, alt))
        ws.write_string(r,  8, _safe_str(row.get("failing_field",     "")), _alt(fmt, "text", alt))
        ws.write_string(r,  9, _safe_str(row.get("current_value",     "")), _alt(fmt, "text", alt))
        ws.write_string(r, 10, _safe_str(row.get("expected_condition","")), _alt(fmt, "text", alt))

    logger.info(f"All Violations sheet: {total_rows:,} rows written")
