# app/db/models.py
# SQLAlchemy ORM models.  One table for run metadata, one for individual
# violations.  The violations table is the persistent fallback when the
# in-memory _job_store has been cleared (e.g. after a service restart).

import json
from datetime import datetime
from sqlalchemy import (
    Column, Integer, String, Float, Text, DateTime, Date,
    ForeignKey, Index,
)
from sqlalchemy.orm import relationship, declarative_base

Base = declarative_base()


class ValidationRun(Base):
    """One record per scheduled pipeline invocation."""
    __tablename__ = "validation_runs"

    id                      = Column(Integer, primary_key=True, autoincrement=True)
    run_id                  = Column(String(8),   unique=True, index=True, nullable=False)
    schedule_slot           = Column(String(5),   nullable=False)   # "06:00"|"12:00"|"18:00"
    run_time                = Column(DateTime,    nullable=False, default=datetime.utcnow)
    blob_date               = Column(Date,        nullable=True)
    status                  = Column(String(20),  nullable=False, default="pending")
    # status lifecycle: "pending" → "running" → "done" | "error"

    total_rows              = Column(Integer,     nullable=True)
    total_violations        = Column(Integer,     nullable=True)
    rules_run               = Column(Integer,     nullable=True)
    validation_time_seconds = Column(Float,       nullable=True)
    error_message           = Column(Text,        nullable=True)
    report_path             = Column(String(500), nullable=True)   # absolute disk path

    # JSON: {"rule_summaries": [...], "facility_summaries": [...]}
    # Stored so the Results dashboard never needs to load violation rows.
    result_json             = Column(Text,        nullable=True)

    errors = relationship("FacilityError", back_populates="run",
                          cascade="all, delete-orphan")

    __table_args__ = (
        Index("ix_vr_status_runtime", "status", "run_time"),
    )

    def set_result_json(self, rule_summaries: list, facility_summaries: list) -> None:
        self.result_json = json.dumps({
            "rule_summaries":     rule_summaries,
            "facility_summaries": facility_summaries,
        })

    def get_result_json(self) -> dict:
        if not self.result_json:
            return {"rule_summaries": [], "facility_summaries": []}
        return json.loads(self.result_json)


class FacilityError(Base):
    """One record per individual patient-level violation.
    Mirrors the All Violations sheet exactly — used by
    GET /validate/{run_id}/violations for paginated, filtered reads
    without requiring the in-memory violations_df.
    """
    __tablename__ = "facility_errors"

    id                 = Column(Integer, primary_key=True, autoincrement=True)
    run_id_fk          = Column(Integer, ForeignKey("validation_runs.id",
                                                    ondelete="CASCADE"), nullable=False)
    row_number         = Column(Integer,      nullable=True)
    patient_id         = Column(String(100),  nullable=True)
    patient_uid        = Column(String(100),  nullable=True)
    facility_name      = Column(String(200),  nullable=True)
    state              = Column(String(100),  nullable=True)
    lga_of_residence   = Column(String(100),  nullable=True)
    rule_id            = Column(String(10),   nullable=False, index=True)
    rule_title         = Column(String(200),  nullable=True)
    severity           = Column(String(10),   nullable=True)
    failing_field      = Column(String(100),  nullable=True)
    current_value      = Column(Text,         nullable=True)
    expected_condition = Column(Text,         nullable=True)

    run = relationship("ValidationRun", back_populates="errors")

    __table_args__ = (
        Index("ix_fe_run_rule",     "run_id_fk", "rule_id"),
        Index("ix_fe_run_facility", "run_id_fk", "facility_name"),
        Index("ix_fe_run_lga",      "run_id_fk", "lga_of_residence"),
        Index("ix_fe_run_state",    "run_id_fk", "state"),
    )
