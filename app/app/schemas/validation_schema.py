# app/schemas/validation_schema.py
# ─────────────────────────────────────────────────────────────────────────────
# CHANGES FROM v1:
#   - Added AsyncJobStartResponse  (returned by POST /validate)
#   - Modified JobStatusResponse   (added optional `result` field populated when done)
#   - All other models UNCHANGED
# ─────────────────────────────────────────────────────────────────────────────

from __future__ import annotations

import math
from datetime import date, datetime
from typing import Optional

from pydantic import BaseModel, ConfigDict, Field, field_validator


# ═══════════════════════════════════════════════════════════════════════════════
# REQUEST MODELS
# ═══════════════════════════════════════════════════════════════════════════════

class ValidateRequest(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    enabled_rules: Optional[list[str]] = Field(default=None)
    reporting_start: Optional[date] = Field(default=None)
    reporting_end: Optional[date] = Field(default=None)

    @field_validator("enabled_rules", mode="before")
    @classmethod
    def parse_comma_string(cls, v):
        if isinstance(v, str):
            return [r.strip() for r in v.split(",") if r.strip()]
        return v

    @field_validator("enabled_rules", mode="after")
    @classmethod
    def validate_rule_ids(cls, v):
        if v is None:
            return v
        import re
        pattern = re.compile(r"^R-\d{2}$")
        invalid = [r for r in v if not pattern.match(r)]
        if invalid:
            raise ValueError(f"Invalid rule ID format: {invalid}. Expected R-01 through R-32")
        return v

    @field_validator("reporting_end", mode="after")
    @classmethod
    def end_after_start(cls, v, info):
        start = info.data.get("reporting_start")
        if start and v and v < start:
            raise ValueError("reporting_end must be on or after reporting_start.")
        return v


# ═══════════════════════════════════════════════════════════════════════════════
# RESPONSE MODELS
# ═══════════════════════════════════════════════════════════════════════════════

class ViolationRecord(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    row_number:         int = Field(description="Excel row number (2 = first data row)")
    patient_id:         str
    patient_uid:        str = ""
    facility_name:      str
    lga_of_residence:   str
    rule_id:            str
    rule_title:         str
    severity:           str
    failing_field:      str
    current_value:      str
    expected_condition: str


class RuleSummaryResponse(BaseModel):
    rule_id:           str
    rule_title:        str
    category:          str
    severity:          str
    total_violations:  int
    pct_rows_affected: float
    skipped:           bool  = False
    skip_reason:       str   = ""


class FacilitySummaryResponse(BaseModel):
    facility_name:     str
    lga_of_residence:  str
    rule_id:           str
    rule_title:        str
    severity:          str
    violation_count:   int
    pct_facility_rows: float


class ValidationResponse(BaseModel):
    """
    Full validation result.
    Returned by GET /validate/{job_id}/status when status == 'done'.
    Previously returned directly by POST /validate (sync version).
    """
    model_config = ConfigDict()

    job_id:                  str
    total_rows:              int
    total_violations:        int
    rules_run:               int
    rules_skipped:           int
    validation_time_seconds: float
    engine_warnings:         list[str]                   = Field(default_factory=list)
    rule_summaries:          list[RuleSummaryResponse]   = Field(default_factory=list)
    facility_summaries:      list[FacilitySummaryResponse] = Field(default_factory=list)
    violations:              list[ViolationRecord]        = Field(default_factory=list)
    total_pages:             int  = 1
    page:                    int  = 1
    page_size:               int  = 200


class ViolationsPageResponse(BaseModel):
    """Response from GET /validate/{job_id}/violations?page=N"""
    job_id:        str
    page:          int
    page_size:     int
    total_pages:   int
    total_records: int
    violations:    list[ViolationRecord]


# ─── NEW: Immediate response from POST /validate ───────────────────────────────

class AsyncJobStartResponse(BaseModel):
    """
    Returned immediately (< 1s) by POST /validate.
    Flutter uses job_id to poll GET /validate/{job_id}/status.
    """
    job_id: str  = Field(description="Poll /validate/{job_id}/status for progress")
    status: str  = Field(default="pending", description="Always 'pending' on creation")


# ─── UPDATED: JobStatusResponse now includes optional full result ──────────────

class JobStatusResponse(BaseModel):
    """
    Response from GET /validate/{job_id}/status.

    Flutter polls this every 2 seconds.
    - While running: progress 0→100, result == null
    - When done:     progress == 100, result contains full ValidationResponse
    - On error:      status == 'error', message contains details
    """
    job_id:   str  = Field(description="Job ID")
    status:   str  = Field(description="pending | processing | done | error")
    progress: int  = Field(default=0, ge=0, le=100, description="0-100")
    message:  str  = Field(default="")
    result:   Optional[ValidationResponse] = Field(
        default=None,
        description="Full result — non-null only when status == 'done'",
    )


class ReportMetaResponse(BaseModel):
    job_id:       str
    file_name:    str
    file_size_kb: float
    generated_at: datetime
    total_sheets: int = 4


class RuleListResponse(BaseModel):
    total: int
    rules: list[RuleDefinition]


class RuleDefinition(BaseModel):
    rule_id:     str       = Field(alias="id", serialization_alias="id")
    title:       str
    description: str
    category:    str
    severity:    str
    enabled:     bool
    columns:     list[str]

    model_config = ConfigDict(populate_by_name=True)


class LastRunInfo(BaseModel):
    run_id:        str
    schedule_slot: str
    run_time:      Optional[str]
    status:        str


class HealthResponse(BaseModel):
    status:      str
    version:     str
    last_run:    Optional[LastRunInfo] = None
    next_run_at: Optional[str]        = None


# ═══════════════════════════════════════════════════════════════════════════════
# CONVERTER HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def rule_summary_to_response(rs) -> RuleSummaryResponse:
    return RuleSummaryResponse(
        rule_id=rs.rule_id,
        rule_title=rs.rule_title,
        category=rs.category,
        severity=rs.severity,
        total_violations=rs.total_violations,
        pct_rows_affected=rs.pct_rows_affected,
        skipped=rs.skipped,
        skip_reason=rs.skip_reason,
    )


def facility_summary_to_response(fs) -> FacilitySummaryResponse:
    return FacilitySummaryResponse(
        facility_name=fs.facility_name,
        lga_of_residence=fs.state,
        rule_id=fs.rule_id,
        rule_title=fs.rule_title,
        severity=fs.severity,
        violation_count=fs.violation_count,
        pct_facility_rows=fs.pct_facility_rows,
    )


def _nan_safe_str(val) -> str:
    """Convert a DataFrame cell to str, returning '' for None or float NaN (NaN != NaN)."""
    if val is None:
        return ""
    if isinstance(val, float) and val != val:
        return ""
    return str(val)


def violations_df_to_records(
    violations_df,
    page: int = 1,
    page_size: int = 200,
) -> tuple[list[ViolationRecord], int, int]:
    total_records = len(violations_df)
    total_pages   = max(1, math.ceil(total_records / page_size))
    page          = max(1, min(page, total_pages))

    start  = (page - 1) * page_size
    end    = start + page_size
    page_df = violations_df.iloc[start:end]

    records = [
        ViolationRecord(
            row_number=int(row["_row_number"]),
            patient_id=str(row["patient_id"]),
            patient_uid=_nan_safe_str(row.get("patient_uid")),
            facility_name=str(row["facility_name"]),
            lga_of_residence=str(row["lga_of_residence"]),
            rule_id=str(row["rule_id"]),
            rule_title=str(row["rule_title"]),
            severity=str(row["severity"]),
            failing_field=str(row["failing_field"]),
            current_value=str(row["current_value"]),
            expected_condition=str(row["expected_condition"]),
        )
        for _, row in page_df.iterrows()
    ]

    return records, total_pages, total_records


def build_validation_response(
    result,
    page: int = 1,
    page_size: int = 200,
    override_job_id: str | None = None,   # ← NEW PARAMETER
) -> ValidationResponse:
    """
    Build the full ValidationResponse from a ValidationResult.
 
    Args:
        result:           ValidationResult from run_validation().
        page:             Which page of violations to include (default 1).
        page_size:        Violations per page (default 200).
        override_job_id:  If provided, use this job_id instead of result.job_id.
                          Used by get_job_status() to ensure ValidationResponse.job_id
                          matches the API job_id (and therefore the report filename),
                          not the internal validator job_id.
    """
    violations, total_pages, total_records = violations_df_to_records(
        result.violations_df, page=page, page_size=page_size
    )
 
    return ValidationResponse(
        job_id=override_job_id if override_job_id is not None else result.job_id,  # ← THE FIX
        total_rows=result.total_rows,
        total_violations=result.total_violations,
        rules_run=result.rules_run,
        rules_skipped=result.rules_skipped,
        validation_time_seconds=result.validation_time_seconds,
        engine_warnings=result.engine_warnings,
        rule_summaries=[
            rule_summary_to_response(rs)
            for rs in result.rule_summaries
        ],
        facility_summaries=[
            facility_summary_to_response(fs)
            for fs in result.facility_summaries
        ],
        violations=violations,
        total_pages=total_pages,
        page=page,
        page_size=page_size,
    )


# ═══════════════════════════════════════════════════════════════════════════════
# COLUMN PREVIEW MODELS
# ═══════════════════════════════════════════════════════════════════════════════

class ColumnMatchItem(BaseModel):
    file_header: str
    canonical:   str
    score:       float


class ColumnPreviewResponse(BaseModel):
    sheet_name:         str
    total_in_file:      int
    total_required:     int
    coverage_pct:       float
    matched:            list[ColumnMatchItem]
    missing_canonicals: list[str]
    extra_file_headers: list[str]
