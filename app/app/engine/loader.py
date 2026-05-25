# Excel reader (test with the RADET template)

# app/engine/loader.py
# ─────────────────────────────────────────────────────────────────────────────
# PURPOSE:
#   Load a RADET Excel file into a clean, validated pandas DataFrame that is
#   ready for the rule engine. Specifically it:
#
#     1. Reads the Excel file with dtype hints (avoids costly type inference
#        on 91 columns × 39,000–100,000 rows)
#     2. Renames columns to canonical names via fuzzy matching
#     3. Normalises all date columns to datetime (bad values → NaT)
#     4. Forces string columns that pandas misreads as float (common in RADET
#        files where TB/screening columns are mostly null in some facilities)
#     5. Cleans whitespace from all string columns
#     6. Checks column coverage and returns a LoadResult with warnings
#
# IMPORTANT ABOUT dtype HINTS:
#   Looking at the actual RADET template, pandas infers many TB columns as
#   float64 because those cells are mostly empty. If we leave them as float,
#   string comparisons like df["tb_status"].str.contains("Presumptive") will
#   crash. We explicitly tell pandas to read them as strings (dtype=str →
#   pandas stores as object, which handles both strings and NaN).
# ─────────────────────────────────────────────────────────────────────────────

import io
import logging
import time
import warnings
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

# Import our helpers — adjust path if running tests from project root
from app.utils.excel_helpers import (
    normalise_column_names,
    normalise_date_columns,
    validate_upload_columns,
    DATE_COLUMNS,
)

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# 1. LOAD RESULT
#    A clean container returned by load_radet_file().
#    The validation service reads df for rules and warnings for the API response.
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class LoadResult:
    df: pd.DataFrame                    # Clean DataFrame ready for rules
    total_rows: int                     # Number of data rows (excludes header)
    total_columns: int                  # Number of columns in file
    column_coverage_pct: float          # % of required columns found
    missing_columns: list[str]          # Canonical names not found in file
    warnings: list[str]                 # Non-fatal issues found during load
    load_time_seconds: float            # How long the load took


# ─────────────────────────────────────────────────────────────────────────────
# 2. DTYPE HINTS
#
#    We tell pandas exactly what type to use for each column BEFORE reading.
#    This serves two purposes:
#      a) Performance: pandas skips type inference on 91 columns
#      b) Correctness: prevents TB/screening columns from becoming float64
#
#    HOW TO READ THIS MAP:
#      "str"   → read as string (object). Use for IDs, status fields,
#                 text fields, and any column that might be mostly null.
#                 NaN cells stay NaN — str dtype does NOT fill them with "nan".
#      "Int64" → nullable integer (capital I). Unlike int64, this allows NaN.
#                 Use for numeric columns that have missing values.
#      We do NOT put date columns here — pd.read_excel() reads them better
#      without hints, and our normaliser handles any edge cases after loading.
#
#    These are the ORIGINAL Excel column names (before renaming).
# ─────────────────────────────────────────────────────────────────────────────

DTYPE_HINTS: dict[str, str] = {
    # ── Identifiers (always string) ───────────────────────────────────────────
    "Patient ID":               "str",
    "NDR Patient Identifier":   "str",
    "Hospital Number":          "str",
    "Unique Id":                "str",
    "Household Unique No":      "str",
    "OVC Unique ID":            "str",
    "DatimId":                  "str",
    "State":                    "str",
    "L.G.A":                    "str",
    "LGA Of Residence":         "str",
    "Facility Name":            "str",

    # ── Demographics ──────────────────────────────────────────────────────────
    "Sex":                      "str",
    "Target group":             "str",
    "Pregnancy Status":         "str",
    "Care Entry Point":         "str",

    # ── ART ───────────────────────────────────────────────────────────────────
    "Regimen Line at ART Start":            "str",
    "Regimen at ART Start":                 "str",
    "Current Regimen Line":                 "str",
    "Current ART Regimen":                  "str",
    "Clinical Staging at Last Visit":       "str",
    "Current ART Status":                   "str",
    "Client Verification Outcome":          "str",
    "Cause of Death":                       "str",
    "VA Cause of Death":                    "str",
    "Previous ART Status":                  "str",
    "ART Enrollment Setting":               "str",
    "Viral Load Indication":                "str",
    "Viral Load Eligibility Status":        "str",
    "Model devolved to":                    "str",
    "Current DSD model":                    "str",
    "Current DSD Outlet":                   "str",

    # ── CD4 ───────────────────────────────────────────────────────────────────
    # IMPORTANT: Last CD4 Count must be str — it contains sentinel strings
    # like "<200", ">=200", "<=200" mixed with numeric values.
    # If read as float, "<200" becomes NaN and real data is lost.
    "Last CD4 Count":           "str",

    # ── ARV Refill — nullable integer (has missing values in some rows) ───────
    "Months of ARV Refill":     "str",

    # ── TB — ALL as str ───────────────────────────────────────────────────────
    # These are inferred as float64 by pandas when mostly null (seen in template).
    # In production RADET files (39K rows) they contain mixed strings/nulls.
    # Reading as str ensures .str.contains() and .str.upper() work safely.
    "TB Screening Type":        "str",
    "TB status":                "str",
    "TB Diagnostic Test Type":  "str",
    "TB Diagnostic Result":     "str",
    "TB Type (new, relapsed etc)": "str",
    "TB Treatment Outcome":     "str",
    "Additional TB Diagnosis Result using XRAY (for client with negative lab results with CAD score of 40 & above)": "str",

    # ── TPT ───────────────────────────────────────────────────────────────────
    "TPT Type":                 "str",
    "TPT Completion status":    "str",

    # ── EAC ───────────────────────────────────────────────────────────────────
    # Sessions is a number but has NaN rows — use nullable Int64
    "Number of EAC Sessions Completed":     "str",
    "Number of Fingers Captured":           "str",
    "Number of Fingers Recaptured":         "str",

    # ── Other ─────────────────────────────────────────────────────────────────
    "Screening for Chronic Conditions":     "str",
    "Co-morbidities":                       "str",
    "Cervical Cancer Screening Type":       "str",
    "Cervical Cancer Screening Method":     "str",
    "Result of Cervical Cancer Screening":  "str",
    "Precancerous Lesions Treatment Methods": "str",
    "Case Manager":             "str",
}


# ─────────────────────────────────────────────────────────────────────────────
# 3. STRING COLUMNS TO CLEAN
#    After loading, strip leading/trailing whitespace from all these columns.
#    This prevents rules from failing because of invisible spaces.
#    e.g. "Active " (with trailing space) would NOT match "Active" without this.
# ─────────────────────────────────────────────────────────────────────────────

STRING_CANONICAL_COLUMNS: list[str] = [
    "patient_id",
    "facility_name",
    "lga_of_residence",
    "state",
    "lga",
    "sex",
    "pregnancy_status",
    "current_art_status",
    "cause_of_death",
    "last_cd4_count",
    "tb_status",
    "tb_screening_type",
    "tb_diagnostic_result",
    "tb_treatment_outcome",
    "tpt_completion_status",
    "case_manager",
]


# ─────────────────────────────────────────────────────────────────────────────
# 4. MAIN LOADER FUNCTION
#
#    Accepts either:
#      - A file path (str or Path) — used in tests and CLI
#      - Raw bytes (bytes or BytesIO) — used when FastAPI receives an upload
#
#    Returns a LoadResult containing the clean DataFrame and metadata.
# ─────────────────────────────────────────────────────────────────────────────

def load_radet_file(
    source: str | Path | bytes | io.BytesIO,
    sheet_name: int | str = 0,
    fuzzy_threshold: int = 80,
) -> LoadResult:
    """
    Load a RADET Excel file and return a clean DataFrame ready for validation.

    Args:
        source:          File path, raw bytes, or BytesIO object.
        sheet_name:      Which sheet to read. Default 0 = first sheet.
        fuzzy_threshold: Column matching sensitivity (0–100). Default 80.

    Returns:
        LoadResult with df, metadata, warnings.

    Raises:
        ValueError:  If the file cannot be read as Excel or has no data rows.
        RuntimeError: For unexpected errors during loading.
    """
    load_warnings: list[str] = []
    start_time = time.perf_counter()

    # ── Step 1: Prepare the source ────────────────────────────────────────────
    # FastAPI gives us bytes from the upload. pandas needs a file-like object.
    if isinstance(source, bytes):
        source = io.BytesIO(source)

    # ── Step 2: Read the Excel file ───────────────────────────────────────────
    # We pass dtype hints to prevent pandas from misreading sparse columns.
    # We do NOT pass parse_dates here — our normaliser handles dates more
    # robustly than pandas' built-in date parser (which crashes on bad values).
    logger.info("Reading Excel file...")
    try:
        with warnings.catch_warnings():
            warnings.filterwarnings(
                "ignore",
                message="Workbook contains no default style",
                category=UserWarning,
                module="openpyxl",
            )
            df = pd.read_excel(
                source,
                sheet_name=sheet_name,
                dtype=DTYPE_HINTS,      # prevents TB columns becoming float64
                engine="openpyxl",      # required for .xlsx files
                na_values=[             # treat these as NaN/null
                    "", "N/A", "n/a", "NA", "na", "NULL", "null",
                    "None", "none", "0000-00-00", "-",
                ],
                keep_default_na=True,   # also keep pandas default NA values
            )
    except Exception as e:
        raise ValueError(
            f"Could not read the uploaded file as an Excel workbook. "
            f"Make sure it is a valid .xlsx file. Detail: {e}"
        )

    logger.info(f"Raw shape: {df.shape[0]} rows × {df.shape[1]} columns")

    # ── Step 3: Basic sanity checks ───────────────────────────────────────────
    if df.empty:
        raise ValueError(
            "The uploaded file is empty — no data rows found. "
            "Check that the correct sheet is being read."
        )

    if df.shape[0] < 2:
        load_warnings.append(
            f"Only {df.shape[0]} data row(s) found. "
            "This may be a template or test file, not a production RADET upload."
        )

    total_rows = df.shape[0]
    total_columns = df.shape[1]

    # ── Step 4: Rename columns to canonical names ─────────────────────────────
    # This is where "ART Start Date (yyyy-mm-dd)" becomes "art_start_date" etc.
    logger.info("Normalising column names...")
    df = normalise_column_names(df, threshold=fuzzy_threshold)
    logger.info(f"Columns after rename ({len(df.columns)}): {list(df.columns)}")

    # ── Step 5: Normalise date columns ────────────────────────────────────────
    # Convert all date columns to datetime64. Bad values become NaT.
    # This runs AFTER renaming so we can use canonical names.
    logger.info("Normalising date columns...")
    df = normalise_date_columns(df, date_columns=DATE_COLUMNS)

    # ── Step 6: Clean whitespace from string columns ──────────────────────────
    # Prevents "Active " (trailing space) from failing status checks.
    logger.info("Cleaning whitespace from string columns...")
    for col in STRING_CANONICAL_COLUMNS:
        if col in df.columns and df[col].dtype == object:
            df[col] = df[col].str.strip()

    # ── Step 7: Normalise Current ART Status casing ───────────────────────────
    # Rules use case-insensitive matching, but let's also store a clean version.
    # We keep original casing in the column — rules do their own .str.upper().
    # Just strip extra whitespace (already done above) — no further change.

    # ── Step 8: Validate column coverage ─────────────────────────────────────
    # Check how many of the 29 required columns are present.
    logger.info("Checking column coverage...")
    coverage_result = validate_upload_columns(df)

    if coverage_result["missing"]:
        for missing_col in coverage_result["missing"]:
            load_warnings.append(
                f"Required column '{missing_col}' not found in this file. "
                f"Rules that depend on it will be skipped."
            )

    # ── Step 9: Add a 1-based row number column ───────────────────────────────
    # Rules use this to report "Row 42 failed Rule R-08".
    # +2 because: +1 for 0-based index → 1-based, +1 for the header row in Excel.
    df["_row_number"] = df.index + 2

    # ── Step 10: Log summary ──────────────────────────────────────────────────
    elapsed = round(time.perf_counter() - start_time, 2)
    logger.info(
        f"Load complete in {elapsed}s — "
        f"{total_rows:,} rows, "
        f"{coverage_result['matched_count']}/{coverage_result['total_required']} "
        f"columns matched ({coverage_result['coverage_pct']}%)"
    )

    return LoadResult(
        df=df,
        total_rows=total_rows,
        total_columns=total_columns,
        column_coverage_pct=coverage_result["coverage_pct"],
        missing_columns=coverage_result["missing"],
        warnings=load_warnings,
        load_time_seconds=elapsed,
    )


# ─────────────────────────────────────────────────────────────────────────────
# 5. QUICK DIAGNOSTIC — run this directly to test the loader against a file
#
#    Usage from project root:
#        python -m app.engine.loader path/to/your_radet.xlsx
#
#    This prints a summary of what the loader found without running any rules.
# ─────────────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
# 5. HTS SHEET LOADER — for cross-sheet rule R-31
#
#    Reads the HTS sheet from the same uploaded Excel file.
#    Sheet name varies by file type:
#      "HTS"         — single-facility per-download file
#      "CombinedHTS" — multi-facility blob (combined export)
#
#    Returns a cleaned DataFrame with canonical snake_case column names,
#    parsed date columns, and a _row_number column.
#    Returns None if no HTS sheet is present (R-31 will skip gracefully).
# ─────────────────────────────────────────────────────────────────────────────

HTS_DTYPE_HINTS: dict[str, str] = {
    "datimCode":          "str",
    "patientId":          "str",
    "sex":                "str",
    "maritalStatus":      "str",
    "LGAOfResidence":     "str",
    "StateOfResidence":   "str",
    "firstTimeVisit":     "str",
    "entrypoint":         "str",
    "indexClient":        "str",
    "previouslyTested":   "str",
    "targetGroup":        "str",
    "referredFrom":       "str",
    "testingSetting":     "str",
    "modality":           "str",
    "counselingType":     "str",
    "pregnancyStatus":    "str",
    "indexType":          "str",
    "previoustestresult": "str",
    "finalHIVTestResult": "str",
    "prepOffered":        "str",
    "prepAccepted":       "str",
}

_HTS_CANONICAL_MAP: dict[str, str] = {
    "datimCode":          "datim_code",
    "patientId":          "patient_id",
    "sex":                "sex",
    "age":                "age",
    "maritalStatus":      "marital_status",
    "LGAOfResidence":     "lga_of_residence",
    "StateOfResidence":   "state_of_residence",
    "dateVisit":          "date_visit",
    "firstTimeVisit":     "first_time_visit",
    "entrypoint":         "entry_point",
    "indexClient":        "index_client",
    "previouslyTested":   "previously_tested",
    "targetGroup":        "target_group",
    "referredFrom":       "referred_from",
    "testingSetting":     "testing_setting",
    "modality":           "modality",
    "counselingType":     "counseling_type",
    "pregnancyStatus":    "pregnancy_status",
    "indexType":          "index_type",
    "PreviousTestDate":   "previous_test_date",
    "previoustestresult": "previous_test_result",
    "htscount":           "hts_count",
    "finalHIVTestResult": "final_hiv_test_result",
    "dateOfHIVTesting":   "date_of_hiv_testing",
    "prepOffered":        "prep_offered",
    "prepAccepted":       "prep_accepted",
}

_HTS_SHEET_CANDIDATES = ["HTS", "CombinedHTS"]
_HTS_DATE_COLUMNS     = ["date_of_hiv_testing", "previous_test_date", "date_visit"]


def load_hts_sheet(
    source: str | Path | bytes | io.BytesIO,
) -> "pd.DataFrame | None":
    """
    Load the HTS sheet from an Excel file for cross-sheet validation (R-31).

    Tries sheet names in order: 'HTS' (single-facility file) then 'CombinedHTS'
    (multi-facility combined blob file). Returns None if neither sheet exists.

    Args:
        source: File path, raw bytes, or BytesIO. When passing the same bytes
                that were given to load_radet_file(), a fresh BytesIO is created
                internally so the two calls are fully independent.

    Returns:
        Cleaned DataFrame with canonical column names and parsed dates, or None.
    """
    # Ensure we have a seekable source that can be re-read independently
    if isinstance(source, bytes):
        excel_source: str | Path | io.BytesIO = io.BytesIO(source)
    elif isinstance(source, io.BytesIO):
        source.seek(0)
        excel_source = source
    else:
        excel_source = source  # str or Path — openpyxl re-opens each time

    hts_df = None
    found_sheet = None
    for sheet in _HTS_SHEET_CANDIDATES:
        try:
            with warnings.catch_warnings():
                warnings.filterwarnings(
                    "ignore",
                    message="Workbook contains no default style",
                    category=UserWarning,
                    module="openpyxl",
                )
                hts_df = pd.read_excel(
                    excel_source,
                    sheet_name=sheet,
                    dtype=HTS_DTYPE_HINTS,
                    engine="openpyxl",
                    na_values=[
                        "", "N/A", "n/a", "NA", "na", "NULL", "null",
                        "None", "none", "0000-00-00", "-",
                    ],
                    keep_default_na=True,
                )
            found_sheet = sheet
            logger.info("HTS sheet '%s' loaded: %d rows.", sheet, len(hts_df))
            break
        except Exception:
            # Sheet not present or unreadable — try next candidate
            if isinstance(excel_source, io.BytesIO):
                excel_source.seek(0)
            continue

    if hts_df is None or hts_df.empty:
        logger.info(
            "No HTS sheet found (tried: %s) — R-31 will be skipped.",
            _HTS_SHEET_CANDIDATES,
        )
        return None

    # Rename to canonical snake_case names
    hts_df = hts_df.rename(columns=_HTS_CANONICAL_MAP)

    # Parse date columns
    for col in _HTS_DATE_COLUMNS:
        if col in hts_df.columns:
            hts_df[col] = pd.to_datetime(
                hts_df[col], errors="coerce"
            ).dt.normalize()

    # Add row number (matches Excel row: +2 for header + 0-based index)
    hts_df["_row_number"] = hts_df.index + 2

    # Strip whitespace from key string columns
    for col in ["patient_id", "final_hiv_test_result", "lga_of_residence",
                "state_of_residence", "datim_code"]:
        if col in hts_df.columns and hts_df[col].dtype == object:
            hts_df[col] = hts_df[col].str.strip()

    logger.info(
        "HTS sheet ready — %d rows from sheet '%s'.", len(hts_df), found_sheet
    )
    return hts_df


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python -m app.engine.loader <path_to_radet.xlsx>")
        sys.exit(1)

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    file_path = sys.argv[1]

    print(f"\n{'='*60}")
    print(f"  RADET Loader Diagnostic")
    print(f"  File: {file_path}")
    print(f"{'='*60}\n")

    try:
        result = load_radet_file(file_path)
    except ValueError as e:
        print(f"ERROR: {e}")
        sys.exit(1)

    print(f"  Rows loaded:       {result.total_rows:,}")
    print(f"  Columns in file:   {result.total_columns}")
    print(f"  Column coverage:   {result.column_coverage_pct}%")
    print(f"  Load time:         {result.load_time_seconds}s")

    if result.missing_columns:
        print(f"\n  Missing columns ({len(result.missing_columns)}):")
        for col in result.missing_columns:
            print(f"    - {col}")
    else:
        print(f"\n  All required columns found ✅")

    if result.warnings:
        print(f"\n  Warnings ({len(result.warnings)}):")
        for w in result.warnings:
            print(f"    ⚠ {w}")
    else:
        print(f"  No warnings ✅")

    print(f"\n  Sample of loaded DataFrame (first 3 rows, key columns):")
    key_cols = [
        c for c in [
            "patient_id", "facility_name", "lga_of_residence",
            "art_start_date", "current_art_status", "current_viral_load",
        ]
        if c in result.df.columns
    ]
    print(result.df[key_cols].head(3).to_string())
    print()
