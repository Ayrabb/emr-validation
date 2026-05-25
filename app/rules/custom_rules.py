# app/rules/custom_rules.py
# ─────────────────────────────────────────────────────────────────────────────
# Vectorised pandas validation rule functions.
#
# Severity values match rules_config.yaml exactly:
#   "Data Quality"            — documentation / data entry error
#   "Service Gap/Data Quality" — missed care intervention + documentation gap
#   "Service Gap"             — missed care intervention
#
# OUTPUT SCHEMA (every rule returns a DataFrame with these columns):
#   _row_number        int  — Excel row number (2 = first data row)
#   patient_id         str  — Patient identifier
#   patient_uid        str  — NDR Patient Identifier (national unique ID)
#   facility_name      str  — Facility
#   state              str  — State
#   lga_of_residence   str  — LGA
#   rule_id            str  — "R-01" … "R-32"
#   rule_title         str  — Human-readable title (must match YAML)
#   severity           str  — "Data Quality" | "Service Gap/Data Quality" | "Service Gap"
#   failing_field      str  — Column(s) that caused the violation
#   current_value      str  — What was found
#   expected_condition str  — What should have been there
# ─────────────────────────────────────────────────────────────────────────────

import logging
import pandas as pd

logger = logging.getLogger(__name__)

CONTEXT = ["_row_number", "patient_id", "patient_uid", "facility_name", "state", "lga_of_residence"]
ACTIVE_STATUSES = {"ACTIVE", "ACTIVE RESTART"}

# ── Severity constants ────────────────────────────────────────────────────────
DQ  = "Data Quality"
SGD = "Service Gap/Data Quality"
SG  = "Service Gap"


# ── Helpers ───────────────────────────────────────────────────────────────────

def _build_output(failing_df, rule_id, rule_title, severity, failing_field, value_col, expected):
    if failing_df.empty:
        return pd.DataFrame(columns=CONTEXT + [
            "rule_id", "rule_title", "severity",
            "failing_field", "current_value", "expected_condition",
        ])
    present_context = [c for c in CONTEXT if c in failing_df.columns]
    out = failing_df[present_context].copy()
    if "state" not in out.columns:
        out["state"] = None
    if "patient_uid" not in out.columns:
        out["patient_uid"] = None
    out["rule_id"]            = rule_id
    out["rule_title"]         = rule_title
    out["severity"]           = severity
    out["failing_field"]      = failing_field
    if isinstance(value_col, list):
        out["current_value"] = (
            failing_df[value_col].astype(str)
            .apply(lambda row: " | ".join(row.values), axis=1)
        )
    else:
        out["current_value"] = failing_df[value_col].astype(str)
    out["expected_condition"] = expected
    return out.reset_index(drop=True)


def _str_contains(series, pattern):
    return series.fillna("").str.strip().str.upper().str.contains(
        pattern.upper(), regex=False, na=False)


def _str_in(series, values):
    upper_vals = {v.upper() for v in values}
    return series.fillna("").str.strip().str.upper().isin(upper_vals)


def _is_missing(series):
    return (
            series.isna()
            | (series.astype(str).str.strip() == "")
            | (series.astype(str).str.strip().str.upper() == "NAN")
    )


def _active_mask(df, col="current_art_status"):
    if col not in df.columns:
        return pd.Series(False, index=df.index)
    return _str_in(df[col], ACTIVE_STATUSES)


def _empty_rule(rule_id, title, severity, field):
    return _build_output(pd.DataFrame(), rule_id, title, severity, field, field, "")


def _date_before_any(df, date_col, baseline_cols, active_only=True):
    base_mask = _active_mask(df) if active_only else pd.Series(True, index=df.index)
    check_date = df[date_col].dt.normalize()
    mask   = pd.Series(False, index=df.index)
    labels = pd.Series("", index=df.index, dtype="object")
    for bc in baseline_cols:
        baseline = df[bc].dt.normalize()
        bad = base_mask & check_date.notna() & baseline.notna() & (check_date < baseline)
        labels.loc[bad & (labels == "")]  = bc
        labels.loc[bad & (labels != bc)] = labels.loc[bad & (labels != bc)] + ", " + bc
        mask = mask | bad
    return mask, labels


def _valid_current_viral_load(value):
    v = str(value).strip()
    if not v or v.upper() == "NAN":
        return True
    # Accept comma-formatted numbers (thousand separators) e.g. 6,5130 → 65130
    try:
        return float(v.replace(",", "")) >= 0
    except ValueError:
        return False


def _current_fy_bounds():
    """Nigerian fiscal year: Oct 1 – Sep 30. Returns (fy_start, fy_end) Timestamps."""
    today = pd.Timestamp.today().normalize()
    if today.month >= 10:
        return pd.Timestamp(today.year, 10, 1), pd.Timestamp(today.year + 1, 9, 30)
    return pd.Timestamp(today.year - 1, 10, 1), pd.Timestamp(today.year, 9, 30)


# ═══════════════════════════════════════════════════════════════════════════════
# RULES
# ═══════════════════════════════════════════════════════════════════════════════

def rule_r01(df):
    """R-01: Client Verification Outcome Required for Active Clients."""
    col_s, col_c = "current_art_status", "client_verification_outcome"
    if col_s not in df.columns or col_c not in df.columns:
        logger.warning("R-01: Required columns missing — skipping.")
        return _empty_rule("R-01", "Client Verification Outcome Required for Active Clients", DQ, col_c)
    mask = _str_in(df[col_s], {"Active", "Active Restart"}) & _is_missing(df[col_c])
    return _build_output(df[mask], "R-01",
                         "Client Verification Outcome Required for Active Clients", DQ,
                         col_c, col_s,
                         "Active/Active Restart clients must have a valid Client Verification Outcome")


def rule_r02(df):
    """R-02: TB Screening Required for Recent Drug Pickup (semiannual window)."""
    RULE_ID, RULE_TITLE = "R-02", "TB Screening Required for Recent Drug Pickup"
    COL_P, COL_T = "last_pickup_date", "tb_screening_date"
    if COL_P not in df.columns or COL_T not in df.columns:
        return _empty_rule(RULE_ID, RULE_TITLE, DQ, f"{COL_P} / {COL_T}")
    today = pd.Timestamp.today().normalize()
    if 4 <= today.month <= 9:
        period_start = pd.Timestamp(today.year, 4, 1)
        period_end   = pd.Timestamp(today.year, 9, 30)
    elif today.month >= 10:
        period_start = pd.Timestamp(today.year,     10, 1)
        period_end   = pd.Timestamp(today.year + 1,  3, 31)
    else:
        period_start = pd.Timestamp(today.year - 1, 10, 1)
        period_end   = pd.Timestamp(today.year,      3, 31)
    pickup = df[COL_P].dt.normalize()
    tb     = df[COL_T].dt.normalize()
    # tb_screening_date > last_pickup_date is valid (sick-visit) — do NOT flag
    mask = pickup.notna() & (pickup >= period_start) & (pickup <= period_end) & tb.isna()
    failing = df[mask].copy()
    failing["_current"] = (
            "Pickup: " + pickup[mask].dt.strftime("%Y-%m-%d") + " | TB Screen: [missing]")
    return _build_output(failing, RULE_ID, RULE_TITLE, DQ,
                         f"{COL_P} / {COL_T}", "_current",
                         f"TB Screening Date required when pickup is within the current semiannual period "
                         f"({period_start.date()} to {period_end.date()})")


def rule_r03(df):
    """R-03: TB Sample Collection Required for Presumptive TB."""
    col_s, col_d, col_a = "tb_status", "tb_sample_collection_date", "current_art_status"
    if any(c not in df.columns for c in [col_s, col_d, col_a]):
        return _empty_rule("R-03", "TB Sample Collection Required for Presumptive TB", SGD, col_d)
    mask = _active_mask(df) & _str_contains(df[col_s], "PRESUMPTIVE") & df[col_d].isna()
    failing = df[mask].copy()
    if not failing.empty:
        failing["_current"] = (
                "Status: " + df[col_a][mask].astype(str)
                + " | TB Status: " + df[col_s][mask].astype(str)
                + " | Sample Collection: [missing]")
    return _build_output(failing, "R-03",
                         "TB Sample Collection Required for Presumptive TB", SGD,
                         col_d, "_current",
                         "Active/Active Restart clients must have TB Sample Collection Date when TB Status = Presumptive TB")


def rule_r04(df):
    """R-04: TB Treatment Start Date Required for Positive Result."""
    col_r, col_d, col_s = "tb_diagnostic_result", "tb_treatment_start_date", "current_art_status"
    if col_r not in df.columns or col_d not in df.columns:
        return _empty_rule("R-04", "TB Treatment Start Date Required for Positive Result", SGD, col_d)
    is_active   = _active_mask(df) if col_s in df.columns else pd.Series(True, index=df.index)
    is_positive = (
            _str_contains(df[col_r], "POSITIVE") |
            _str_contains(df[col_r], "MTB DETECTED") |
            _str_contains(df[col_r], "MTB_DETECTED")
    )
    mask = is_active & is_positive & df[col_d].isna()
    return _build_output(df[mask], "R-04",
                         "TB Treatment Start Date Required for Positive Result", SGD,
                         col_d, col_r,
                         "TB Treatment Start Date must not be NULL when TB Diagnostic Result is Positive")


def rule_r05(df):
    """R-05: TPT Start Date Required for Negative TB Results."""
    col_r, col_d, col_s = "tb_diagnostic_result", "tpt_start_date", "current_art_status"
    if any(c not in df.columns for c in [col_r, col_d, col_s]):
        return _empty_rule("R-05", "TPT Start Date Required for Negative TB Results", SGD, col_d)
    is_active = _active_mask(df)
    norm = (df[col_r].fillna("").astype(str).str.strip().str.upper()
            .str.replace(r"[-_]+", " ", regex=True).str.replace(r"\s+", " ", regex=True))
    is_negative = norm.isin({"NEGATIVE", "MTB NOT DETECTED",
                             "X RAY NOT SUGGESTIVE", "XRAY NOT SUGGESTIVE"})
    mask = is_active & is_negative & df[col_d].isna()
    failing = df[mask].copy()
    if not failing.empty:
        failing["_current"] = (
                "Status: " + df[col_s][mask].astype(str)
                + " | TB Result: " + df[col_r][mask].astype(str)
                + " | TPT Start: [missing]")
    return _build_output(failing, "R-05",
                         "TPT Start Date Required for Negative TB Results", SGD,
                         col_d, "_current",
                         "Active/Active Restart clients must have TPT Start Date when TB result is Negative / MTB Not Detected / X-Ray Not Suggestive")


def rule_r06(df):
    """R-06: TPT Start Date Required — No Signs or Symptoms."""
    col_s, col_d, col_a = "tb_status", "tpt_start_date", "current_art_status"
    if col_s not in df.columns or col_d not in df.columns:
        return _empty_rule("R-06", "TPT Start Date Required — No Signs or Symptoms", SGD, col_d)
    is_active = _active_mask(df) if col_a in df.columns else pd.Series(True, index=df.index)
    is_no_signs = _str_contains(df[col_s], "NO SIGN") | _str_contains(df[col_s], "NO SYMPTOM")
    mask = is_active & is_no_signs & df[col_d].isna()
    return _build_output(df[mask], "R-06",
                         "TPT Start Date Required — No Signs or Symptoms", SGD,
                         col_d, col_s,
                         "TPT Start Date must not be NULL when TB Status = No Signs or Symptoms of TB")


def rule_r07(df):
    """R-07: Unsuppressed client (VL >= 1000) must have EAC Commencement Date."""
    col_s, col_v, col_e = "current_art_status", "current_viral_load", "eac_commencement_date"
    for col in [col_s, col_v, col_e]:
        if col not in df.columns:
            return _empty_rule("R-07", "Unsppressed Client Without EAC Commencement", SG, col_e)
    is_active  = _active_mask(df)
    vl_numeric = pd.to_numeric(df[col_v], errors="coerce")
    mask       = is_active & (vl_numeric >= 1000) & df[col_e].isna()
    failing = df[mask].copy()
    failing["_current"] = (
            "Status: " + df[col_s][mask].astype(str) + " | VL: " + vl_numeric[mask].astype(str))
    return _build_output(failing, "R-07",
                         "Unsppressed Client Without EAC Commencement", SG,      # title matches YAML exactly
                         col_e, "_current",
                         "EAC Commencement Date must not be NULL when Status is Active/Active Restart and VL >= 1,000")


def rule_r08(df, reporting_start=None, reporting_end=None):
    """R-08: Biometric Enrollment Date required for Active clients aged >= 5."""
    col_s, col_b = "current_art_status", "biometrics_enrolled_date"
    if col_s not in df.columns or col_b not in df.columns:
        return _empty_rule("R-08", "Biometric Enrollment Required for All Active ART Clients", SG, col_b)
    is_active   = _active_mask(df)
    bio_missing = _is_missing(df[col_b])
    age_ok      = (pd.to_numeric(df["age"], errors="coerce") >= 5) if "age" in df.columns else pd.Series(True, index=df.index)
    mask = is_active & bio_missing & age_ok
    failing = df[mask].copy()
    if not failing.empty:
        failing["_current"] = (
                "Status: " + df[col_s][mask].astype(str) + " | Biometrics Enrolled Date: [missing]")
    return _build_output(failing, "R-08",
                         "Biometric Enrollment Required for All Active ART Clients", SG,
                         col_b, "_current",
                         "Biometric Enrollment Date must not be NULL for Active or Active Restart clients (age >= 5)")


def rule_r09(df):
    """R-09: Case Manager Required for Active Clients."""
    col_s, col_c = "current_art_status", "case_manager"
    if col_s not in df.columns or col_c not in df.columns:
        return _empty_rule("R-09", "Case Manager Required for Active Clients", DQ, col_c)
    mask = _active_mask(df) & (_is_missing(df[col_c]))
    return _build_output(df[mask], "R-09",
                         "Case Manager Required for Active Clients", DQ,
                         col_c, col_s,
                         "Case Manager must not be NULL for Active or Active Restart clients")


def rule_r10(df):
    """R-10: Date of Birth must not be after ART Start or Last Pickup Date."""
    RULE_ID, TITLE = "R-10", "Date of Birth Must Not Be After ART Start or Last Pickup Date"
    cols = ["current_art_status", "dob", "art_start_date", "last_pickup_date"]
    if any(c not in df.columns for c in cols):
        return _empty_rule(RULE_ID, TITLE, DQ, "dob")
    is_active = _active_mask(df)
    dob    = df["dob"].dt.normalize()
    art    = df["art_start_date"].dt.normalize()
    pickup = df["last_pickup_date"].dt.normalize()
    mask   = (is_active & dob.notna() & art.notna() & (dob > art)) | \
             (is_active & dob.notna() & pickup.notna() & (dob > pickup))
    failing = df[mask].copy()
    if not failing.empty:
        failing["_current"] = (
                "DOB: " + dob[mask].astype(str)
                + " | ART Start: " + art[mask].astype(str)
                + " | Last Pickup: " + pickup[mask].astype(str))
    return _build_output(failing, RULE_ID, TITLE, DQ,
                         "dob / art_start_date / last_pickup_date", "_current",
                         "Date of Birth must not be after ART Start Date or Last Pickup Date")


def rule_r11(df):
    """R-11: Last Pickup Date must not be earlier than registration, enrollment, or ART Start."""
    RULE_ID, TITLE = "R-11", "Last Pickup Date Must Not Be Earlier Than Registration, Enrollment, or ART Start Date"
    cols = ["current_art_status", "last_pickup_date", "date_of_registration", "enrollment_date", "art_start_date"]
    if any(c not in df.columns for c in cols):
        return _empty_rule(RULE_ID, TITLE, DQ, "last_pickup_date")
    mask, labels = _date_before_any(df, "last_pickup_date",
                                    ["date_of_registration", "enrollment_date", "art_start_date"])
    pickup = df["last_pickup_date"].dt.normalize()
    failing = df[mask].copy()
    if not failing.empty:
        failing["_current"] = "Pickup: " + pickup[mask].astype(str) + " | Earlier than: " + labels[mask]
    return _build_output(failing, RULE_ID, TITLE, DQ,
                         "last_pickup_date / date_of_registration / enrollment_date / art_start_date", "_current",
                         "Last Pickup Date must be on or after registration, enrollment, and ART Start Date")


def rule_r12(df):
    """R-12: TB Treatment Completion Date must be >= TB Treatment Start Date.

    Retreatment clients are excluded — their treatment episodes can span earlier
    dates that legitimately pre-date the current episode's start date.
    """
    RULE_ID, TITLE = "R-12", "TB Treatment Completion Date Must Be After Start Date"
    col_e, col_s = "tb_treatment_completion_date", "tb_treatment_start_date"
    if col_e not in df.columns or col_s not in df.columns:
        return _empty_rule(RULE_ID, TITLE, DQ, col_e)
    end   = df[col_e].dt.normalize()
    start = df[col_s].dt.normalize()
    mask  = end.notna() & start.notna() & (end < start)
    # Exclude Retreatment TB type (New, Relapsed, etc. are still checked)
    if "tb_type" in df.columns:
        mask = mask & ~_str_in(df["tb_type"], {"RETREATMENT"})
    failing = df[mask].copy()
    if not failing.empty:
        tb_type_str = (" | TB Type: " + df["tb_type"][mask].astype(str)) if "tb_type" in df.columns else ""
        failing["_current"] = (
            "Completion: " + end[mask].astype(str)
            + " | Start: " + start[mask].astype(str)
            + tb_type_str
        )
    return _build_output(failing, RULE_ID, TITLE, DQ,
                         "tb_treatment_completion_date / tb_treatment_start_date", "_current",
                         "TB Treatment Completion Date must be on or after TB Treatment Start Date "
                         "(Retreatment clients excluded)")


def rule_r13(df):
    """R-13: TPT Completion Date must be >= TPT Start Date."""
    RULE_ID, TITLE = "R-13", "TPT Completion Date Must Be After TPT Start Date"
    col_e, col_s = "tpt_completion_date", "tpt_start_date"
    if col_e not in df.columns or col_s not in df.columns:
        return _empty_rule(RULE_ID, TITLE, DQ, col_e)
    end = df[col_e].dt.normalize(); start = df[col_s].dt.normalize()
    is_active = _active_mask(df) if "current_art_status" in df.columns else pd.Series(True, index=df.index)
    mask = is_active & end.notna() & start.notna() & (end < start)
    failing = df[mask].copy()
    if not failing.empty:
        failing["_current"] = "Completion: " + end[mask].astype(str) + " | Start: " + start[mask].astype(str)
    return _build_output(failing, RULE_ID, TITLE, DQ,
                         "tpt_completion_date / tpt_start_date", "_current",
                         "TPT Completion Date must be on or after TPT Start Date")


_VALID_CD4_SENTINELS = {">=200", "<=200", "<200", ">200"}

def _is_valid_cd4(value):
    v = str(value).strip()
    if v.upper() in {s.upper() for s in _VALID_CD4_SENTINELS}:
        return True
    try:
        n = float(v)
        return 0.0 <= n <= 3000.0
    except (ValueError, TypeError):
        return False


def rule_r14(df):
    """R-14: Last CD4 Count must be numeric 0–3,000 or valid sentinel. NULL values not flagged."""
    RULE_ID, TITLE = "R-14", "Last CD4 Count Must Be a Valid Value"
    col, col_s = "last_cd4_count", "current_art_status"
    if col not in df.columns or col_s not in df.columns:
        return _empty_rule(RULE_ID, TITLE, DQ, col)
    is_active    = _active_mask(df)
    non_null     = df[col].notna()
    invalid_cd4  = df[col][non_null].apply(lambda v: not _is_valid_cd4(v))
    mask = is_active & non_null & df.index.isin(invalid_cd4[invalid_cd4].index)
    failing = df[mask].copy()
    if not failing.empty:
        failing["_current"] = (
                "Status: " + df[col_s][mask].astype(str) + " | Last CD4 Count: " + df[col][mask].astype(str))
    return _build_output(failing, RULE_ID, TITLE, DQ,
                         col, "_current",
                         "Must be numeric 0–3,000 or one of: >=200 | <=200 | <200 | >200")


def rule_r15(df):
    """R-15: Cause of Death required when Current ART Status = Died.

    Scoped to the current Nigerian FY (Oct 1 – Sep 30) via date_current_art_status.
    Records with a missing or out-of-FY date_current_art_status are excluded.
    """
    RULE_ID, TITLE = "R-15", "Cause of Death Required When Status Is Died"
    col_s, col_c, col_d = "current_art_status", "cause_of_death", "date_current_art_status"
    if col_s not in df.columns or col_c not in df.columns:
        return _empty_rule(RULE_ID, TITLE, DQ, col_c)
    mask = _str_in(df[col_s], {"DIED"}) & _is_missing(df[col_c])
    # Restrict to records whose ART status date falls in the current FY
    if col_d in df.columns:
        fy_start, fy_end = _current_fy_bounds()
        dca  = df[col_d].dt.normalize()
        mask = mask & dca.notna() & (dca >= fy_start) & (dca <= fy_end)
    else:
        logger.warning("R-15: column '%s' not present — FY filter skipped.", col_d)
    failing = df[mask].copy()
    if not failing.empty:
        failing["_current"] = df[col_s][mask].astype(str)
    return _build_output(failing, RULE_ID, TITLE, DQ,
                         col_c, "_current",
                         "Cause of Death must not be NULL when Current ART Status = Died "
                         "(current FY only — date_current_art_status in Oct–Sep)")


def rule_r16(df):
    """R-16: Sex required for Active/Active Restart clients."""
    RULE_ID, TITLE = "R-16", "Sex Required for Active Clients"
    if any(c not in df.columns for c in ["current_art_status", "sex"]):
        return _empty_rule(RULE_ID, TITLE, DQ, "sex")
    mask = _active_mask(df) & _is_missing(df["sex"])
    return _build_output(df[mask], RULE_ID, TITLE, DQ,
                         "sex", "current_art_status",
                         "Sex must not be blank for Active/Active Restart clients")


def rule_r17(df):
    """R-17: Pregnancy Status should only appear for female clients."""
    RULE_ID, TITLE = "R-17", "Pregnancy Status Only Applies to Female Clients"
    if any(c not in df.columns for c in ["current_art_status", "sex", "pregnancy_status"]):
        return _empty_rule(RULE_ID, TITLE, DQ, "pregnancy_status")
    sex  = df["sex"].fillna("").astype(str).str.strip().str.upper()
    mask = _active_mask(df) & ~_is_missing(df["pregnancy_status"]) & (sex != "FEMALE")
    failing = df[mask].copy()
    if not failing.empty:
        failing["_current"] = (
                "Sex: " + df["sex"][mask].astype(str)
                + " | Pregnancy Status: " + df["pregnancy_status"][mask].astype(str))
    return _build_output(failing, RULE_ID, TITLE, DQ,
                         "sex / pregnancy_status", "_current",
                         "Pregnancy Status must only be populated when Sex is Female")


def rule_r18(df):
    """R-18: Date of Birth and Age required for Active/Active Restart clients."""
    RULE_ID, TITLE = "R-18", "Date of Birth and Age Required for Active Clients"
    if any(c not in df.columns for c in ["current_art_status", "dob", "age"]):
        return _empty_rule(RULE_ID, TITLE, DQ, "dob / age")
    mask = _active_mask(df) & (df["dob"].isna() | _is_missing(df["age"]))
    failing = df[mask].copy()
    if not failing.empty:
        failing["_current"] = (
                "DOB: " + df["dob"][mask].astype(str) + " | Age: " + df["age"][mask].astype(str))
    return _build_output(failing, RULE_ID, TITLE, DQ,
                         "dob / age", "_current",
                         "Date of Birth and Age must not be blank for Active/Active Restart clients")


def rule_r19(df):
    """
    R-19: ART Start Date must not be earlier than Date of Registration or Enrollment Date.

    Excludes Transfer-in clients — their ART may have started at the previous facility
    before they were registered here, so an earlier ART date is clinically valid.

    entry_point column is checked if present; if absent the exclusion is skipped
    so the rule still runs on files that don't have that column.
    """
    RULE_ID, TITLE = "R-19", "ART Start Date Must Not Be Earlier Than Date of Registration or Enrollment"
    base_cols = ["current_art_status", "art_start_date", "date_of_registration", "enrollment_date"]
    if any(c not in df.columns for c in base_cols):
        logger.warning("R-19: Required columns missing — skipping.")
        return _empty_rule(RULE_ID, TITLE, DQ, "art_start_date")

    mask, labels = _date_before_any(df, "art_start_date",
                                    ["date_of_registration", "enrollment_date"], active_only=True)

    # Exclude Transfer-in clients if entry_point column is available
    if "entry_point" in df.columns:
        is_transfer_in = _str_contains(df["entry_point"], "TRANSFER")
        mask = mask & ~is_transfer_in
        logger.debug("R-19: entry_point column found — Transfer-in clients excluded.")
    else:
        logger.debug("R-19: entry_point column not present — Transfer-in exclusion skipped.")

    art = df["art_start_date"].dt.normalize()
    failing = df[mask].copy()
    if not failing.empty:
        failing["_current"] = (
                "ART Start: " + art[mask].astype(str) + " | Earlier than: " + labels[mask])
    return _build_output(failing, RULE_ID, TITLE, DQ,
                         "art_start_date / date_of_registration / enrollment_date", "_current",
                         "ART Start Date must be on or after Date of Registration and Enrollment Date "
                         "(Transfer-in clients are excluded)")


def rule_r20(df):
    """R-20: Last Pickup Date and Months of ARV Refill required for Active clients."""
    RULE_ID, TITLE = "R-20", "Last Pickup Date and Months of ARV Refill Required"
    cols = ["current_art_status", "last_pickup_date", "months_arv_refill"]
    if any(c not in df.columns for c in cols):
        return _empty_rule(RULE_ID, TITLE, DQ, "last_pickup_date / months_arv_refill")
    mask = _active_mask(df) & (df["last_pickup_date"].isna() | _is_missing(df["months_arv_refill"]))
    failing = df[mask].copy()
    if not failing.empty:
        failing["_current"] = (
                "Last Pickup: " + df["last_pickup_date"][mask].astype(str)
                + " | Refill: " + df["months_arv_refill"][mask].astype(str))
    return _build_output(failing, RULE_ID, TITLE, DQ,
                         "last_pickup_date / months_arv_refill", "_current",
                         "Last Pickup Date and Months of ARV Refill must not be blank for Active/Active Restart clients")


def rule_r21(df):
    """R-21: Current Viral Load must be a non-negative integer or decimal value.

    Only flags Active/Active Restart clients where VL is PRESENT but not a valid
    non-negative number.  Text entries (TND, Tiertermin, Target Not Detected, etc.)
    are invalid.  NULL / blank VL is not flagged by this rule.
    """
    RULE_ID, TITLE = "R-21", "Current Viral Load Must Be Valid"
    if any(c not in df.columns for c in ["current_art_status", "current_viral_load"]):
        return _empty_rule(RULE_ID, TITLE, DQ, "current_viral_load")
    present  = ~_is_missing(df["current_viral_load"])
    # reindex so the boolean Series aligns with df.index (present rows only → fill rest False)
    is_invalid = (
        df["current_viral_load"][present]
        .apply(lambda v: not _valid_current_viral_load(v))
        .reindex(df.index, fill_value=False)
    )
    mask = _active_mask(df) & present & is_invalid
    failing = df[mask].copy()
    if not failing.empty:
        failing["_current"] = (
            "Status: " + df["current_art_status"][mask].astype(str)
            + " | VL: " + df["current_viral_load"][mask].astype(str)
        )
    return _build_output(failing, RULE_ID, TITLE, DQ,
                         "current_viral_load", "_current",
                         "Current Viral Load must be a non-negative integer or decimal value — "
                         "text entries (TND, Tiertermin, Target Not Detected, etc.) are not accepted")


def rule_r22(df):
    """R-22: Current ART Status must not be blank."""
    RULE_ID, TITLE = "R-22", "Current ART Status Required for All Records"
    if "current_art_status" not in df.columns:
        return _empty_rule(RULE_ID, TITLE, DQ, "current_art_status")
    mask = _is_missing(df["current_art_status"])
    return _build_output(df[mask], RULE_ID, TITLE, DQ,
                         "current_art_status", "current_art_status",
                         "Current ART Status must not be blank")


def rule_r23(df):
    """R-23: Previous ART status confirmed date must not be after current ART status date.

    Scoped to the current Nigerian FY (Oct 1 – Sep 30) via date_current_art_status.
    Only records whose date_current_art_status falls within the current FY are evaluated.
    """
    RULE_ID, TITLE = "R-23", "Previous ART Status Date Must Not Be After Current ART Status Date"
    cols = ["date_current_art_status", "confirmed_date_previous_art_status"]
    if any(c not in df.columns for c in cols):
        return _empty_rule(RULE_ID, TITLE, DQ, "confirmed_date_previous_art_status")
    prev = df["confirmed_date_previous_art_status"].dt.normalize()
    curr = df["date_current_art_status"].dt.normalize()
    # Restrict to current FY on date_current_art_status
    fy_start, fy_end = _current_fy_bounds()
    in_fy = curr.notna() & (curr >= fy_start) & (curr <= fy_end)
    mask  = in_fy & prev.notna() & (prev > curr)
    failing = df[mask].copy()
    if not failing.empty:
        failing["_current"] = (
            "Previous Confirmed: " + prev[mask].astype(str)
            + " | Current ART Status Date: " + curr[mask].astype(str)
        )
    return _build_output(failing, RULE_ID, TITLE, DQ,
                         "confirmed_date_previous_art_status / date_current_art_status", "_current",
                         "Confirmed Date of Previous ART Status must not be after Date of Current ART Status "
                         "(current FY only — date_current_art_status in Oct–Sep)")


def rule_r24(df):
    """R-24: TB Screening Date must not be earlier than ART Start or Last Pickup Date."""
    RULE_ID, TITLE = "R-24", "TB Screening Date Must Not Be Earlier Than ART Start or Last Pickup Date"
    cols = ["current_art_status", "tb_screening_date", "art_start_date", "last_pickup_date"]
    if any(c not in df.columns for c in cols):
        return _empty_rule(RULE_ID, TITLE, DQ, "tb_screening_date")
    mask, labels = _date_before_any(df, "tb_screening_date",
                                    ["art_start_date", "last_pickup_date"])
    failing = df[mask].copy()
    if not failing.empty:
        failing["_current"] = (
                "TB Screening: " + df["tb_screening_date"][mask].astype(str)
                + " | Earlier than: " + labels[mask])
    return _build_output(failing, RULE_ID, TITLE, DQ,
                         "tb_screening_date / art_start_date / last_pickup_date", "_current",
                         "TB Screening Date must be on or after ART Start Date and Last Pickup Date")


def rule_r25(df):
    """R-25: TB Screening Type and TB Status required when TB Screening Date exists."""
    RULE_ID, TITLE = "R-25", "TB Screening Type and TB Status Required When Screening Date Exists"
    cols = ["current_art_status", "tb_screening_date", "tb_screening_type", "tb_status"]
    if any(c not in df.columns for c in cols):
        return _empty_rule(RULE_ID, TITLE, DQ, "tb_screening_type / tb_status")
    mask = (_active_mask(df) & df["tb_screening_date"].notna()
            & (_is_missing(df["tb_screening_type"]) | _is_missing(df["tb_status"])))
    failing = df[mask].copy()
    if not failing.empty:
        failing["_current"] = (
                "TB Screening Date: " + df["tb_screening_date"][mask].astype(str)
                + " | Type: " + df["tb_screening_type"][mask].astype(str)
                + " | Status: " + df["tb_status"][mask].astype(str))
    return _build_output(failing, RULE_ID, TITLE, DQ,
                         "tb_screening_type / tb_status / tb_screening_date", "_current",
                         "TB Screening Type and TB Status must not be blank when TB Screening Date is present")


def rule_r26(df):
    """R-26: TB Treatment Start Date must not be earlier than diagnostic result received date."""
    RULE_ID, TITLE = "R-26", "TB Treatment Start Date Must Not Be Earlier Than Diagnostic Result Received Date"
    cols = ["current_art_status", "tb_treatment_start_date", "tb_diagnostic_result_received_date"]
    if any(c not in df.columns for c in cols):
        return _empty_rule(RULE_ID, TITLE, DQ, "tb_treatment_start_date")
    start       = df["tb_treatment_start_date"].dt.normalize()
    result_date = df["tb_diagnostic_result_received_date"].dt.normalize()
    mask = _active_mask(df) & start.notna() & result_date.notna() & (start < result_date)
    failing = df[mask].copy()
    if not failing.empty:
        failing["_current"] = (
                "Treatment Start: " + start[mask].astype(str)
                + " | Result Received: " + result_date[mask].astype(str))
    return _build_output(failing, RULE_ID, TITLE, DQ,
                         "tb_treatment_start_date / tb_diagnostic_result_received_date", "_current",
                         "TB Treatment Start Date must be on or after TB Diagnostic Result Received Date")

def rule_r27(df):
    """R-27: TPT Completion Status required when TPT Completion Date exists."""
    RULE_ID, TITLE = "R-27", "TPT Completion Status Required When TPT Completion Date Exists"
    cols = ["current_art_status", "tpt_completion_date", "tpt_completion_status"]
    if any(c not in df.columns for c in cols):
        return _empty_rule(RULE_ID, TITLE, DQ, "tpt_completion_status")
    mask = _active_mask(df) & df["tpt_completion_date"].notna() & _is_missing(df["tpt_completion_status"])
    return _build_output(df[mask], RULE_ID, TITLE, DQ,
                         "tpt_completion_status / tpt_completion_date", "tpt_completion_date",
                         "TPT Completion Status must not be blank when TPT Completion Date is present")


def rule_r28(df):
    """R-28: EAC dates must not be earlier than ART Start Date."""
    RULE_ID, TITLE = "R-28", "EAC Dates Must Not Be Earlier Than ART Start Date"
    date_cols = ["eac_commencement_date", "eac_last_session_date",
                 "extended_eac_completion_date",
                 "post_eac_vl_sample_collection_date", "post_eac_vl_result_date"]
    cols = ["current_art_status", "art_start_date"] + date_cols
    if any(c not in df.columns for c in cols):
        return _empty_rule(RULE_ID, TITLE, DQ, " / ".join(date_cols))
    is_active = _active_mask(df)
    art    = df["art_start_date"].dt.normalize()
    mask   = pd.Series(False, index=df.index)
    labels = pd.Series("", index=df.index, dtype="object")
    for col in date_cols:
        d   = df[col].dt.normalize()
        bad = is_active & art.notna() & d.notna() & (d < art)
        labels.loc[bad & (labels == "")] = col
        labels.loc[bad & (labels != col)] = labels.loc[bad & (labels != col)] + ", " + col
        mask = mask | bad
    failing = df[mask].copy()
    if not failing.empty:
        failing["_current"] = (
                "ART Start: " + art[mask].astype(str) + " | Earlier EAC date(s): " + labels[mask])
    return _build_output(failing, RULE_ID, TITLE, DQ,
                         " / ".join(date_cols + ["art_start_date"]), "_current",
                         "EAC dates must be on or after ART Start Date")


def rule_r29(df):
    """R-29: Last EAC Session date must not be earlier than EAC Commencement date."""
    RULE_ID, TITLE = "R-29", "Last EAC Session Date Must Not Be Earlier Than EAC Commencement Date"
    cols = ["current_art_status", "eac_commencement_date", "eac_last_session_date"]
    if any(c not in df.columns for c in cols):
        return _empty_rule(RULE_ID, TITLE, DQ, "eac_last_session_date")
    start = df["eac_commencement_date"].dt.normalize()
    last  = df["eac_last_session_date"].dt.normalize()
    mask  = _active_mask(df) & start.notna() & last.notna() & (last < start)
    failing = df[mask].copy()
    if not failing.empty:
        failing["_current"] = (
                "EAC Start: " + start[mask].astype(str) + " | Last Session: " + last[mask].astype(str))
    return _build_output(failing, RULE_ID, TITLE, DQ,
                         "eac_last_session_date / eac_commencement_date", "_current",
                         "Date of Last EAC Session Completed must be on or after Date of Commencement of EAC")


def rule_r30(df):
    """R-30: DSD dates must not be earlier than ART Start Date."""
    RULE_ID, TITLE = "R-30", "DSD Dates Must Not Be Earlier Than ART Start Date"
    date_cols = ["devolvement_date", "current_dsd_date", "dsd_return_date"]
    cols = ["current_art_status", "art_start_date"] + date_cols
    if any(c not in df.columns for c in cols):
        return _empty_rule(RULE_ID, TITLE, DQ, " / ".join(date_cols))
    is_active = _active_mask(df)
    art    = df["art_start_date"].dt.normalize()
    mask   = pd.Series(False, index=df.index)
    labels = pd.Series("", index=df.index, dtype="object")
    for col in date_cols:
        d   = df[col].dt.normalize()
        bad = is_active & art.notna() & d.notna() & (d < art)
        labels.loc[bad & (labels == "")] = col
        labels.loc[bad & (labels != col)] = labels.loc[bad & (labels != col)] + ", " + col
        mask = mask | bad
    failing = df[mask].copy()
    if not failing.empty:
        failing["_current"] = (
                "ART Start: " + art[mask].astype(str) + " | Earlier DSD date(s): " + labels[mask])
    return _build_output(failing, RULE_ID, TITLE, DQ,
                         " / ".join(date_cols + ["art_start_date"]), "_current",
                         "DSD dates must be on or after ART Start Date")


def rule_r31(df: pd.DataFrame, aux_dfs: dict | None = None) -> pd.DataFrame:
    """R-31: Every HIV+ HTS result in the current FY must have a matching patient in RADET."""
    RULE_ID = "R-31"
    TITLE   = "HIV-Positive HTS Client Not Enrolled in RADET"
    _EMPTY  = pd.DataFrame(columns=CONTEXT + [
        "rule_id", "rule_title", "severity",
        "failing_field", "current_value", "expected_condition",
    ])

    if not aux_dfs:
        logger.warning("R-31: aux_dfs not provided — HTS sheet not loaded; rule skipped.")
        return _EMPTY

    hts_df = aux_dfs.get("hts")
    if hts_df is None or hts_df.empty:
        logger.warning("R-31: HTS sheet is empty or missing; rule skipped.")
        return _EMPTY

    required_hts_cols = ["patient_id", "final_hiv_test_result", "date_of_hiv_testing"]
    missing_hts = [c for c in required_hts_cols if c not in hts_df.columns]
    if missing_hts:
        logger.warning("R-31: HTS missing required columns %s — rule skipped.", missing_hts)
        return _EMPTY

    fy_start, fy_end = _current_fy_bounds()

    # ── Filter HTS to HIV+ results within the current fiscal year ────────────
    test_date  = hts_df["date_of_hiv_testing"]
    hiv_result = hts_df["final_hiv_test_result"].fillna("").str.strip().str.upper()

    in_fy       = test_date.notna() & (test_date >= fy_start) & (test_date <= fy_end)
    is_positive = hiv_result == "POSITIVE"

    hts_positives = hts_df[in_fy & is_positive].copy()
    if hts_positives.empty:
        return _EMPTY

    # ── Build DATIM code lookups from RADET (facility name + state) ─────────
    # Uses RADET's canonical datim_id column to map each DATIM code to the
    # facility's name and state. This ensures R-31 output carries the
    # facility's state (e.g. Borno), not the patient's residential state
    # (e.g. Gombe) — which would break _build_facility_summaries grouping.
    datim_to_facility: dict[str, str] = {}
    datim_to_state:    dict[str, str] = {}
    radet_datim_col = next(
        (c for c in ["datim_id", "datimid"] if c in df.columns), None
    )
    if radet_datim_col:
        _ref = df[[radet_datim_col] + [c for c in ["facility_name", "state"] if c in df.columns]].dropna(subset=[radet_datim_col]).drop_duplicates(subset=[radet_datim_col])
        if "facility_name" in _ref.columns:
            datim_to_facility = _ref.set_index(radet_datim_col)["facility_name"].to_dict()
        if "state" in _ref.columns:
            datim_to_state = _ref.set_index(radet_datim_col)["state"].to_dict()

    # ── Set of all patient IDs already recorded in RADET (case-insensitive) ──
    radet_ids = set(df["patient_id"].dropna().str.strip().str.upper())

    # ── Find positives whose patient ID has no matching RADET record ─────────
    pid_upper  = hts_positives["patient_id"].fillna("").str.strip().str.upper()
    not_linked = hts_positives[~pid_upper.isin(radet_ids)].copy()

    if not_linked.empty:
        return _EMPTY

    # ── Build output DataFrame ────────────────────────────────────────────────
    out = pd.DataFrame(index=not_linked.index)
    out["_row_number"] = not_linked["_row_number"] if "_row_number" in not_linked.columns else range(len(not_linked))
    out["patient_id"]  = not_linked["patient_id"]
    out["patient_uid"] = not_linked.get("patient_uid", pd.Series("", index=not_linked.index))

    if "datim_code" in not_linked.columns and datim_to_facility:
        out["facility_name"] = not_linked["datim_code"].map(
            lambda c: datim_to_facility.get(str(c).strip(), str(c).strip())
        )
    elif "datim_code" in not_linked.columns:
        out["facility_name"] = not_linked["datim_code"]
    else:
        out["facility_name"] = ""

    # Use facility's state from RADET (via DATIM lookup) so the By Facility
    # grouping in _build_facility_summaries matches RADET's facility totals.
    if "datim_code" in not_linked.columns and datim_to_state:
        out["state"] = not_linked["datim_code"].map(
            lambda c: datim_to_state.get(str(c).strip(), "")
        )
    else:
        out["state"] = pd.Series("", index=not_linked.index)
    out["lga_of_residence"] = (
        not_linked["lga_of_residence"]
        if "lga_of_residence" in not_linked.columns
        else pd.Series("", index=not_linked.index)
    )

    out["rule_id"]    = RULE_ID
    out["rule_title"] = TITLE
    out["severity"]   = SG
    out["failing_field"] = "patientId (HTS → RADET)"
    out["current_value"] = (
        "HIV+ on "
        + not_linked["date_of_hiv_testing"].dt.strftime("%Y-%m-%d").fillna("unknown date")
    )
    out["expected_condition"] = (
        "Every HIV+ result from HTS in the current FY must have a matching patient in RADET"
    )

    return out.reset_index(drop=True)


def rule_r32(df: pd.DataFrame, aux_dfs: dict | None = None) -> pd.DataFrame:
    """R-32: Every RADET patient newly starting ART this FY (non-Transfer-in) must have a
    matching HIV+ HTS record in the current FY."""
    RULE_ID = "R-32"
    TITLE   = "New ART Patient in RADET Missing HIV+ HTS Record"
    _EMPTY  = pd.DataFrame(columns=CONTEXT + [
        "rule_id", "rule_title", "severity",
        "failing_field", "current_value", "expected_condition",
    ])

    if not aux_dfs:
        logger.warning("R-32: aux_dfs not provided — HTS sheet not loaded; rule skipped.")
        return _EMPTY

    hts_df = aux_dfs.get("hts")
    if hts_df is None or hts_df.empty:
        logger.warning("R-32: HTS sheet is empty or missing; rule skipped.")
        return _EMPTY

    required_hts_cols = ["patient_id", "final_hiv_test_result", "date_of_hiv_testing"]
    missing_hts = [c for c in required_hts_cols if c not in hts_df.columns]
    if missing_hts:
        logger.warning("R-32: HTS missing required columns %s — rule skipped.", missing_hts)
        return _EMPTY

    if "art_start_date" not in df.columns:
        logger.warning("R-32: RADET missing art_start_date column — rule skipped.")
        return _EMPTY

    fy_start, fy_end = _current_fy_bounds()

    # ── Filter RADET to new ART starters this FY, excluding Transfer-in ──────
    art_start = df["art_start_date"].dt.normalize()
    in_fy     = art_start.notna() & (art_start >= fy_start) & (art_start <= fy_end)

    if "entry_point" in df.columns:
        is_transfer_in = _str_contains(df["entry_point"], "TRANSFER")
        new_starters   = df[in_fy & ~is_transfer_in].copy()
    else:
        new_starters = df[in_fy].copy()

    if new_starters.empty:
        return _EMPTY

    # ── Build set of HTS patient IDs with HIV+ result in current FY ──────────
    test_date  = hts_df["date_of_hiv_testing"]
    hiv_result = hts_df["final_hiv_test_result"].fillna("").str.strip().str.upper()
    hts_in_fy  = test_date.notna() & (test_date >= fy_start) & (test_date <= fy_end)
    hts_pos_ids = set(
        hts_df.loc[hts_in_fy & (hiv_result == "POSITIVE"), "patient_id"]
        .dropna().str.strip().str.upper()
    )

    # ── Find RADET new starters with no matching HTS HIV+ record ─────────────
    pid_upper  = new_starters["patient_id"].fillna("").str.strip().str.upper()
    not_linked = new_starters[~pid_upper.isin(hts_pos_ids)].copy()

    if not_linked.empty:
        return _EMPTY

    # ── Build output DataFrame (facility/state come directly from RADET) ──────
    out = pd.DataFrame(index=not_linked.index)
    out["_row_number"]    = not_linked.get("_row_number", pd.Series(range(len(not_linked)), index=not_linked.index))
    out["patient_id"]     = not_linked["patient_id"]
    out["patient_uid"]    = not_linked.get("patient_uid",    pd.Series("", index=not_linked.index))
    out["facility_name"]  = not_linked.get("facility_name",  pd.Series("", index=not_linked.index))
    out["state"]          = not_linked.get("state",          pd.Series("", index=not_linked.index))
    out["lga_of_residence"] = not_linked.get("lga_of_residence", pd.Series("", index=not_linked.index))

    out["rule_id"]    = RULE_ID
    out["rule_title"] = TITLE
    out["severity"]   = SG
    out["failing_field"] = "patientId (RADET → HTS)"
    out["current_value"] = (
        "ART started "
        + not_linked["art_start_date"].dt.strftime("%Y-%m-%d").fillna("unknown date")
    )
    out["expected_condition"] = (
        "Every non-Transfer-in patient starting ART in the current FY must have "
        "a matching HIV+ HTS result in the current FY"
    )

    return out.reset_index(drop=True)


# ── RULE REGISTRY ─────────────────────────────────────────────────────────────
RULE_REGISTRY: dict[str, callable] = {
    "R-01": rule_r01, "R-02": rule_r02, "R-03": rule_r03,
    "R-04": rule_r04, "R-05": rule_r05, "R-06": rule_r06,
    "R-07": rule_r07, "R-08": rule_r08, "R-09": rule_r09,
    "R-10": rule_r10, "R-11": rule_r11, "R-12": rule_r12,
    "R-13": rule_r13, "R-14": rule_r14, "R-15": rule_r15,
    "R-16": rule_r16, "R-17": rule_r17, "R-18": rule_r18,
    "R-19": rule_r19, "R-20": rule_r20, "R-21": rule_r21,
    "R-22": rule_r22, "R-23": rule_r23, "R-24": rule_r24,
    "R-25": rule_r25, "R-26": rule_r26,
    "R-27": rule_r27, "R-28": rule_r28,
    "R-29": rule_r29, "R-30": rule_r30,
    "R-31": rule_r31, "R-32": rule_r32,
}