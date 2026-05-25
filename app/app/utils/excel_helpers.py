# app/utils/excel_helpers.py
# ─────────────────────────────────────────────────────────────────────────────
# PURPOSE:
#   1. Fuzzy-match actual Excel column headers → canonical (standard) names
#   2. Normalise date columns safely (handles mixed types, NaT, bad values)
#   3. Report which required columns are missing before any rule runs
#
# WHY THIS EXISTS:
#   RADET files from different facilities often have slightly different column
#   names. For example:
#       "ART Start Date"  vs  "ART Start Date (yyyy-mm-dd)"  vs  "ARTStartDate"
#   Without this helper, the engine would crash on every naming variation.
#   With it, all variations map to one clean internal name used by all rules.
# ─────────────────────────────────────────────────────────────────────────────

import pandas as pd
from rapidfuzz import fuzz, process
import logging

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# 1. CANONICAL COLUMN MAP
#    Left side  = the clean internal name your rules will use
#    Right side = list of known Excel header variations across facilities
#
#    When you add a new rule that needs a new column, add it here.
# ─────────────────────────────────────────────────────────────────────────────

CANONICAL_COLUMNS: dict[str, list[str]] = {
    # ── Identifiers ──────────────────────────────────────────────────────────
    "patient_id": [
        "Patient ID",
        "PatientID",
        "Patient Id",
        "patient_id",
    ],
    "patient_uid": [
        "NDR Patient Identifier",
        "NDR Patient ID",
        "NDR Identifier",
        "patient_uid",
    ],
    "facility_name": [
        "Facility Name",
        "FacilityName",
        "Facility",
        "facility_name",
    ],
    "lga_of_residence": [
        "LGA Of Residence",
        "LGA of Residence",
        "LGAOfResidence",
        "lga_of_residence",
        "LGA Residence",
    ],
    "state": [
        "State",
        "state",
    ],
    "lga": [
        "L.G.A",
        "LGA",
        "lga",
    ],
    "datim_id": [
        "DatimId",
        "datimCode",
        "DATIM ID",
        "Datim Id",
    ],

    # ── Demographics ──────────────────────────────────────────────────────────
    "dob": [
        "Date of Birth (yyyy-mm-dd)",
        "Date of Birth",
        "DOB",
        "DateOfBirth",
        "dob",
    ],
    "age": [
        "Age",
        "age",
    ],
    "sex": [
        "Sex",
        "Gender",
        "sex",
    ],

    # ── ART ───────────────────────────────────────────────────────────────────
    "art_start_date": [
        "ART Start Date (yyyy-mm-dd)",
        "ART Start Date",
        "ARTStartDate",
        "art_start_date",
        "ART Initiation Date",
        "Date of ART Start",
    ],
    "current_art_status": [
        "Current ART Status",
        "CurrentARTStatus",
        "current_art_status",
        "ART Status",
    ],
    "last_pickup_date": [
        "Last Pickup Date (yyyy-mm-dd)",
        "Last Pickup Date",
        "LastPickupDate",
        "last_pickup_date",
        "Last Drug Pickup Date",
    ],
    "months_arv_refill": [
        "Months of ARV Refill",
        "MonthsOfARVRefill",
        "months_arv_refill",
        "ARV Refill Months",
        "Months ARV Refill",
    ],
    "cause_of_death": [
        "Cause of Death",
        "CauseOfDeath",
        "cause_of_death",
    ],

    # ── CD4 ───────────────────────────────────────────────────────────────────
    "last_cd4_date": [
        "Date of Last CD4 Count",
        "DateOfLastCD4Count",
        "last_cd4_date",
        "Last CD4 Date",
        "CD4 Date",
    ],
    "last_cd4_count": [
        "Last CD4 Count",
        "LastCD4Count",
        "last_cd4_count",
        "CD4 Count",
    ],

    # ── Viral Load ────────────────────────────────────────────────────────────
    "current_viral_load": [
        "Current Viral Load (c/ml)",
        "Current Viral Load",
        "CurrentViralLoad",
        "current_viral_load",
        "Viral Load",
        "VL Result",
    ],

    # ── TB ────────────────────────────────────────────────────────────────────
    "tb_screening_date": [
        "Date of TB Screening (yyyy-mm-dd)",
        "Date of TB Screening",
        "TBScreeningDate",
        "tb_screening_date",
    ],
    "tb_status": [
        "TB status",
        "TB Status",
        "TBStatus",
        "tb_status",
    ],
    "tb_sample_collection_date": [
        "Date of TB Sample Collection (yyyy-mm-dd)",
        "Date of TB Sample Collection",
        "TBSampleCollectionDate",
        "tb_sample_collection_date",
    ],
    "tb_diagnostic_result": [
        "TB Diagnostic Result",
        "TBDiagnosticResult",
        "tb_diagnostic_result",
        "TB Result",
    ],
    "tb_treatment_start_date": [
        "Date of Start of TB Treatment (yyyy-mm-dd)",
        "Date of Start of TB Treatment",
        "TBTreatmentStartDate",
        "tb_treatment_start_date",
        "TB Treatment Start",
    ],
    "tb_treatment_completion_date": [
        "Date of Completion of TB Treatment (yyyy-mm-dd)",
        "Date of Completion of TB Treatment",
        "TBTreatmentCompletionDate",
        "tb_treatment_completion_date",
    ],
    "tb_type": [
        "TB Type (new, relapsed etc)",
        "TB Type",
        "TBType",
        "tb_type",
        "Type of TB",
        "TB Treatment Type",
    ],

    # ── TPT ───────────────────────────────────────────────────────────────────
    "tpt_start_date": [
        "Date of TPT Start (yyyy-mm-dd)",
        "Date TPT Start",
        "TPT Start Date",
        "TPTStartDate",
        "tpt_start_date",
        "Date of Start of TPT",
    ],
    "tpt_completion_date": [
        "TPT Completion date (yyyy-mm-dd)",
        "TPT Completion Date",
        "TPTCompletionDate",
        "tpt_completion_date",
    ],

    # ── EAC ───────────────────────────────────────────────────────────────────
    "eac_commencement_date": [
        "Date of commencement of EAC (yyyy-mm-dd)",
        "Date of Commencement of EAC",
        "EACCommencementDate",
        "eac_commencement_date",
        "EAC Start Date",
    ],
    "eac_sessions": [
        "Number of EAC Sessions Completed",
        "EAC Sessions Completed",
        "EACSessions",
        "eac_sessions",
        "Number of EAC Sessions",
    ],

    # ── Biometrics ────────────────────────────────────────────────────────────
    "biometrics_enrolled_date": [
        "Date Biometrics Enrolled (yyyy-mm-dd)",
        "Date Biometric Enrolled",
        "Date Biometrics Enrolled",
        "BiometricsEnrolledDate",
        "biometrics_enrolled_date",
        "Biometric Enrollment Date",
    ],


    "pregnancy_status": [
        "Pregnancy Status",
        "PregnancyStatus",
        "pregnancy_status",
    ],
    "date_of_registration": [
        "Date of Registration",
        "Date of Registration (yyyy-mm-dd)",
        "Registration Date",
        "date_of_registration",
    ],
    "enrollment_date": [
        "Enrollment Date",
        "Date of Enrollment",
        "Enrollment Date (yyyy-mm-dd)",
        "enrollment_date",
    ],
    "date_current_art_status": [
        "Date of Current ART Status",
        "Date of Current ART Status (yyyy-mm-dd)",
        "Current ART Status Date",
        "date_current_art_status",
    ],
    "confirmed_date_previous_art_status": [
        "Confirmed Date of Previous ART Status",
        "Confirmed Date of Previous ART Status (yyyy-mm-dd)",
        "Previous ART Status Confirmed Date",
        "confirmed_date_previous_art_status",
    ],
    "tb_screening_type": [
        "TB Screening Type",
        "TBScreeningType",
        "tb_screening_type",
    ],
    "tb_diagnostic_result_received_date": [
        "Date of TB Diagnostic Result Received (yyyy-mm-dd)",
        "Date of TB Diagnostic Result Received",
        "TB Diagnostic Result Received Date",
        "tb_diagnostic_result_received_date",
    ],
    "tpt_completion_status": [
        "TPT Completion Status",
        "TPTCompletionStatus",
        "tpt_completion_status",
    ],
    "eac_last_session_date": [
        "Date of Last EAC Session Completed",
        "Date of Last EAC Session Completed (yyyy-mm-dd)",
        "Last EAC Session Completed Date",
        "eac_last_session_date",
    ],
    "extended_eac_completion_date": [
        "Date of Extended EAC Completion (yyyy-mm-dd)",
        "Date of Extended EAC Completion",
        "Extended EAC Completion Date",
        "extended_eac_completion_date",
    ],
    "post_eac_vl_sample_collection_date": [
        "Date of Repeat Viral Load - Post EAC VL Sample Collected (yyyy-mm-dd)",
        "Date of Repeat Viral Load - Post EAC VL Sample collected",
        "Post EAC VL Sample Collection Date",
        "post_eac_vl_sample_collection_date",
    ],
    "post_eac_vl_result_date": [
        "Date of Repeat Viral Load Result - POST EAC VL",
        "Date of Repeat Viral load result - POST EAC VL",
        "Post EAC VL Result Date",
        "post_eac_vl_result_date",
    ],
    "devolvement_date": [
        "Date of Devolvement",
        "Date of Devolvement (yyyy-mm-dd)",
        "devolvement_date",
    ],
    "current_dsd_date": [
        "Date of Current DSD",
        "Date of Current DSD (yyyy-mm-dd)",
        "current_dsd_date",
    ],
    "dsd_return_date": [
        "Date of Return of DSD Client to Facility (yyyy-mm-dd)",
        "Date of Return of DSD Client to Facility",
        "DSD Return Date",
        "dsd_return_date",
    ],

    # ── Case Manager ──────────────────────────────────────────────────────────
    "case_manager": [
        "Case Manager",
        "CaseManager",
        "case_manager",
    ],
    "client_verification_outcome": [
        "Client Verification Outcome",
        "ClientVerificationOutcome",
        "Client Verification",
        "Verification Outcome",
        "CV Outcome",
    ],

    # ── Entry Point — used by R-19 to exclude Transfer-in clients ────────────
    "entry_point": [
        "Care Entry Point",
        "Entry Point",
        "EntryPoint",
        "entry_point",
        "Point of Entry",
        "ART Entry Point",
        "Patient Entry Point",
    ],
}


# ─────────────────────────────────────────────────────────────────────────────
# 1b. COLUMN BLOCKLIST
#    Some actual Excel column names are superficially similar to the wrong
#    canonical because they share multiple words. For example, both VL and TB
#    columns contain "Sample Collection Date", scoring 86%+ with token_sort_ratio.
#    If the VL column is iterated first it steals tb_sample_collection_date,
#    leaving the real TB column unmapped — causing R-04 false positives.
#
#    Format: { "actual_excel_header": {"blocked_canonical_1", ...} }
# ─────────────────────────────────────────────────────────────────────────────

COLUMN_BLOCKLIST: dict[str, set[str]] = {
    "Date of Viral Load Sample Collection (yyyy-mm-dd)":
        {"tb_sample_collection_date"},
    "Date of Current ViralLoad Result Sample (yyyy-mm-dd)":
        {"tb_sample_collection_date"},
    "Date of Repeat Viral Load - Post EAC VL Sample collected (yyyy-mm-dd)":
        {"tb_sample_collection_date"},
    "Date of Cervical Cancer Screening (yyyy-mm-dd)":
        {"tb_screening_date"},
    "Date of Return of DSD Client to Facility (yyyy-mm-dd)":
        {"tb_treatment_completion_date"},
    "Date Biometrics Recapture (yyyy-mm-dd)":
        {"biometrics_enrolled_date"},
}


# ─────────────────────────────────────────────────────────────────────────────
# 2. DATE COLUMNS
#    These are all columns that should be treated as dates.
#    The loader will call normalise_date_columns() on all of these.
# ─────────────────────────────────────────────────────────────────────────────

DATE_COLUMNS: list[str] = [
    "dob",
    "art_start_date",
    "last_pickup_date",
    "last_cd4_date",
    "tb_screening_date",
    "tb_sample_collection_date",
    "tb_treatment_start_date",
    "tb_treatment_completion_date",
    "tpt_start_date",
    "tpt_completion_date",
    "eac_commencement_date",
    "biometrics_enrolled_date",
    "date_of_registration",
    "enrollment_date",
    "date_current_art_status",
    "confirmed_date_previous_art_status",
    "tb_diagnostic_result_received_date",
    "eac_last_session_date",
    "extended_eac_completion_date",
    "post_eac_vl_sample_collection_date",
    "post_eac_vl_result_date",
    "devolvement_date",
    "current_dsd_date",
    "dsd_return_date",
]


# ─────────────────────────────────────────────────────────────────────────────
# 3. FUZZY COLUMN MATCHER
#
#    HOW IT WORKS:
#    - Takes the actual headers from the uploaded Excel file
#    - For each canonical column, tries to find a match in the actual headers
#    - Uses fuzzy string matching (rapidfuzz) with a similarity threshold
#    - Returns a rename_map: { "actual header": "canonical_name" }
#
#    SIMILARITY THRESHOLD = 80
#    This means the actual header must be at least 80% similar to a known
#    variation. 100 = exact match. Lower = more lenient but risks false matches.
#    80 is a safe balance for RADET column name variations.
# ─────────────────────────────────────────────────────────────────────────────

def build_column_rename_map(
        actual_headers: list[str],
        threshold: int = 80,
) -> dict[str, str]:
    """
    Match actual Excel headers to canonical column names using fuzzy matching.

    Strategy (safer than flat-list scoring):
      For each actual header, score it against EACH canonical's own variation
      list separately. Take the canonical with the highest score overall.
      This prevents "Date of Registration" from accidentally matching
      "tb_screening_date" just because both contain date-related words.

    Args:
        actual_headers: The raw column names from the uploaded Excel file.
        threshold:      Minimum similarity score (0–100). Default 80.

    Returns:
        A dict mapping actual header → canonical name.
        e.g. { "ART Start Date (yyyy-mm-dd)": "art_start_date" }
    """
    rename_map: dict[str, str] = {}

    # Pre-build: exact lookup (case-insensitive) for fast-path
    exact_lookup: dict[str, str] = {}
    for canonical, variations in CANONICAL_COLUMNS.items():
        for variation in variations:
            exact_lookup[variation.lower().strip()] = canonical

    # Track which canonical names have already been claimed
    # so one canonical cannot be assigned to two different actual columns
    claimed: dict[str, str] = {}  # canonical → actual header that claimed it

    for actual_header in actual_headers:
        lower_actual = actual_header.lower().strip()

        # ── Fast path: exact match (case-insensitive) ─────────────────────────
        if lower_actual in exact_lookup:
            canonical = exact_lookup[lower_actual]
            if canonical not in claimed:
                rename_map[actual_header] = canonical
                claimed[canonical] = actual_header
                logger.debug(f"Exact match: '{actual_header}' → '{canonical}'")
            else:
                logger.warning(
                    f"Column '{actual_header}' matches '{canonical}' "
                    f"but it was already claimed by '{claimed[canonical]}'. Skipping."
                )
            continue

        # ── Fuzzy path: score against EACH canonical's own variation list ─────
        # For each canonical, find its best-scoring variation against actual_header.
        # Then pick the canonical with the highest best-score overall.
        best_canonical: str | None = None
        best_score: int = 0
        best_variation: str = ""

        for canonical, variations in CANONICAL_COLUMNS.items():
            result = process.extractOne(
                actual_header,
                variations,
                scorer=fuzz.token_sort_ratio,
            )
            if result is None:
                continue
            matched_variation, score, _ = result
            if score > best_score:
                best_score = score
                best_canonical = canonical
                best_variation = matched_variation

        if best_canonical is not None and best_score >= threshold:
            # Check blocklist — certain headers must never map to certain canonicals
            blocked = COLUMN_BLOCKLIST.get(actual_header, set())
            if best_canonical in blocked:
                logger.info(
                    f"Blocklist: '{actual_header}' → '{best_canonical}' "
                    f"blocked ({best_score}%). Finding next best match."
                )
                # Re-score excluding the blocked canonical
                best_canonical, best_score, best_variation = None, 0, ""
                for canonical, variations in CANONICAL_COLUMNS.items():
                    if canonical in blocked:
                        continue
                    result = process.extractOne(
                        actual_header, variations, scorer=fuzz.token_sort_ratio
                    )
                    if result is None:
                        continue
                    mv, sc, _ = result
                    if sc > best_score:
                        best_score = sc
                        best_canonical = canonical
                        best_variation = mv
                if best_canonical is None or best_score < threshold:
                    logger.debug(
                        f"No valid match for '{actual_header}' after blocklist filter."
                    )
                    continue

            if best_canonical not in claimed:
                rename_map[actual_header] = best_canonical
                claimed[best_canonical] = actual_header
                logger.info(
                    f"Fuzzy match ({best_score}%): "
                    f"'{actual_header}' → '{best_canonical}' "
                    f"(via '{best_variation}')"
                )
            else:
                logger.debug(
                    f"Skipping '{actual_header}' → '{best_canonical}' "
                    f"(already claimed by '{claimed[best_canonical]}')"
                )
        else:
            if best_score > 0:
                logger.debug(
                    f"No match for '{actual_header}' "
                    f"(best: '{best_canonical}' at {best_score}% — "
                    f"below threshold {threshold}%)"
                )

    return rename_map


# ─────────────────────────────────────────────────────────────────────────────
# 4. APPLY RENAME MAP TO DATAFRAME
#
#    Takes the rename_map from build_column_rename_map() and applies it
#    to the DataFrame. Columns not in the map are kept with their original
#    name — they won't interfere with validation.
# ─────────────────────────────────────────────────────────────────────────────

def normalise_column_names(df: pd.DataFrame, threshold: int = 80) -> pd.DataFrame:
    """
    Rename DataFrame columns from actual Excel headers to canonical names.

    Args:
        df:         The raw DataFrame from pd.read_excel().
        threshold:  Fuzzy match threshold (default 80).

    Returns:
        DataFrame with columns renamed to canonical names where matched.
    """
    actual_headers = df.columns.tolist()
    rename_map = build_column_rename_map(actual_headers, threshold)

    if rename_map:
        df = df.rename(columns=rename_map)
        logger.info(f"Renamed {len(rename_map)} columns to canonical names.")
    else:
        logger.warning("No columns could be matched to canonical names.")

    return df


# ─────────────────────────────────────────────────────────────────────────────
# 5. DATE NORMALISER
#
#    WHY THIS IS NEEDED:
#    Excel date columns are messy in the real world:
#      - Some cells are proper datetime objects (from Excel date cells)
#      - Some are strings: "2024-01-15", "15/01/2024", "Jan 15 2024"
#      - Some are Excel serial numbers: 45306
#      - Some are empty (NaN / NaT)
#      - Some are garbage: "N/A", "unknown", "0000-00-00"
#
#    pd.to_datetime(errors='coerce') handles all of these:
#      - Valid dates   → proper datetime
#      - Bad values    → NaT (Not a Time — pandas equivalent of NULL for dates)
#      - Empty cells   → NaT
#
#    After this, your rules can safely compare dates without crashing.
# ─────────────────────────────────────────────────────────────────────────────

def normalise_date_columns(
        df: pd.DataFrame,
        date_columns: list[str] | None = None,
) -> pd.DataFrame:
    """
    Convert all date columns to pandas datetime (NaT for invalid/missing).

    Args:
        df:           DataFrame with canonical column names already applied.
        date_columns: List of canonical column names to normalise.
                      Defaults to DATE_COLUMNS defined at top of this file.

    Returns:
        DataFrame with date columns converted to datetime64[ns].
    """
    if date_columns is None:
        date_columns = DATE_COLUMNS

    for col in date_columns:
        if col not in df.columns:
            # Column missing entirely — skip silently.
            # The column validator will catch this separately.
            continue

        original_dtype = df[col].dtype

        df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=False)

        # Count how many values became NaT due to bad data (not just blanks)
        nat_count = df[col].isna().sum()
        if nat_count > 0:
            logger.debug(
                f"Column '{col}': {nat_count} values are NaT "
                f"(original dtype was {original_dtype})"
            )

    return df


# ─────────────────────────────────────────────────────────────────────────────
# 6. COLUMN PRESENCE VALIDATOR
#
#    Before any rule runs, check that the columns it needs actually exist
#    in the DataFrame after renaming. Returns two lists:
#      - present:  columns that exist ✅
#      - missing:  columns that don't ❌
#
#    The validation service uses this to:
#      a) Warn the user which columns are missing
#      b) Skip rules whose required columns are absent (instead of crashing)
# ─────────────────────────────────────────────────────────────────────────────

def check_required_columns(
        df: pd.DataFrame,
        required: list[str],
) -> tuple[list[str], list[str]]:
    """
    Check which required canonical columns are present in the DataFrame.

    Args:
        df:       DataFrame with canonical column names already applied.
        required: List of canonical column names to check.

    Returns:
        (present, missing) — two lists.
    """
    present = [col for col in required if col in df.columns]
    missing = [col for col in required if col not in df.columns]

    if missing:
        logger.warning(f"Missing columns: {missing}")

    return present, missing


# ─────────────────────────────────────────────────────────────────────────────
# 7. CONVENIENCE FUNCTION — get all required columns across active rules
#    Used by the loader to do one pass column check on upload.
# ─────────────────────────────────────────────────────────────────────────────

ALL_REQUIRED_COLUMNS: list[str] = list(CANONICAL_COLUMNS.keys())


def validate_upload_columns(df: pd.DataFrame) -> dict:
    """
    Run a full column presence check after normalising column names.
    Called once when a file is first uploaded, before any rule runs.

    Returns a dict with:
      - matched_count:   how many required columns were found
      - total_required:  total number of required columns
      - missing:         list of canonical names that couldn't be matched
      - coverage_pct:    percentage of required columns found
    """
    present, missing = check_required_columns(df, ALL_REQUIRED_COLUMNS)

    coverage = round(len(present) / len(ALL_REQUIRED_COLUMNS) * 100, 1)

    result = {
        "matched_count": len(present),
        "total_required": len(ALL_REQUIRED_COLUMNS),
        "missing": missing,
        "coverage_pct": coverage,
    }

    logger.info(
        f"Column coverage: {len(present)}/{len(ALL_REQUIRED_COLUMNS)} "
        f"({coverage}%) — Missing: {missing if missing else 'none'}"
    )

    return result