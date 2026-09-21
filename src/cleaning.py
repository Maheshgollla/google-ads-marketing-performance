"""
Google Ads Marketing Performance
--------------------------------
Phase 3: Data Cleaning

This module is the single source of truth for Phase 3 cleaning.

Design principles:
- Never modify the raw CSV.
- Never overwrite raw data.
- Transform copies of DataFrames.
- Do not perform business-metric calculations in Phase 3.
- Do not automatically impute missing values without an approved rule.
- Preserve raw keyword values for auditability.
- Produce a cleaned CSV in data/processed/.
- Produce cleaning audit files in reports/cleaning/.

Phase 5 calculations such as CTR, CPC, CPL, CVR, CPA, ROAS
and ROI are intentionally NOT performed here.
"""

from pathlib import Path
from typing import Dict, Tuple

import pandas as pd
from openpyxl import load_workbook


# ============================================================================
# PROJECT PATHS
# ============================================================================

PROJECT_ROOT = Path(__file__).resolve().parents[1]

RAW_DATA_PATH = (
    PROJECT_ROOT
    / "data"
    / "raw"
    / "GoogleAds_DataAnalytics_Sales_Uncleaned.csv"
)

PROCESSED_DATA_DIR = PROJECT_ROOT / "data" / "processed"
CLEANED_DATA_PATH = PROCESSED_DATA_DIR / "google_ads_cleaned.csv"

CLEANING_REPORT_DIR = PROJECT_ROOT / "reports" / "cleaning"
EXCEL_WORKBOOK_PATH = PROJECT_ROOT / "excel" / "marketing_campaign_analysis.xlsx"


# ============================================================================
# PHASE 3 BASELINE
# ============================================================================

EXPECTED_RAW_ROWS = 2600
EXPECTED_RAW_COLUMNS = 13


EXPECTED_RAW_COLUMNS_LIST = [
    "Ad_ID",
    "Campaign_Name",
    "Clicks",
    "Impressions",
    "Cost",
    "Leads",
    "Conversions",
    "Conversion Rate",
    "Sale_Amount",
    "Ad_Date",
    "Location",
    "Device",
    "Keyword",
]


EXPECTED_CLEAN_COLUMNS = [
    "ad_id",
    "campaign_name",
    "clicks",
    "impressions",
    "cost",
    "leads",
    "conversions",
    "conversion_rate",
    "sale_amount",
    "ad_date",
    "location",
    "device",
    "keyword_raw",
    "keyword_clean",
]


# ============================================================================
# COLUMN STANDARDIZATION
# ============================================================================

COLUMN_MAPPING: Dict[str, str] = {
    "Ad_ID": "ad_id",
    "Campaign_Name": "campaign_name",
    "Clicks": "clicks",
    "Impressions": "impressions",
    "Cost": "cost",
    "Leads": "leads",
    "Conversions": "conversions",
    "Conversion Rate": "conversion_rate",
    "Sale_Amount": "sale_amount",
    "Ad_Date": "ad_date",
    "Location": "location",
    "Device": "device",
    "Keyword": "keyword",
}


# ============================================================================
# APPROVED CATEGORICAL STANDARDIZATION
# ============================================================================

CAMPAIGN_MAPPING: Dict[str, str] = {
    "Data Analytics Course": "Data Analytics Course",
    "Data Analytcis Course": "Data Analytics Course",
    "DataAnalyticsCourse": "Data Analytics Course",
    "Data Anlytics Corse": "Data Analytics Course",
    "Data Analytics Corse": "Data Analytics Course",
}


LOCATION_MAPPING: Dict[str, str] = {
    "hyderabad": "Hyderabad",
    "HYDERABAD": "Hyderabad",
    "Hyderbad": "Hyderabad",
    "hydrebad": "Hyderabad",
}


DEVICE_MAPPING: Dict[str, str] = {
    "desktop": "Desktop",
    "Desktop": "Desktop",
    "DESKTOP": "Desktop",
    "mobile": "Mobile",
    "Mobile": "Mobile",
    "MOBILE": "Mobile",
    "tablet": "Tablet",
    "Tablet": "Tablet",
    "TABLET": "Tablet",
}


# ============================================================================
# CLEANING DICTIONARY
# ============================================================================

CLEANING_DICTIONARY = pd.DataFrame(
    [
        {
            "raw_field": "Ad_ID",
            "clean_field": "ad_id",
            "transformation": "Rename",
            "rule": "Preserve identifier values unchanged.",
        },
        {
            "raw_field": "Campaign_Name",
            "clean_field": "campaign_name",
            "transformation": "Rename + standardize",
            "rule": "Map approved campaign variants to Data Analytics Course.",
        },
        {
            "raw_field": "Clicks",
            "clean_field": "clicks",
            "transformation": "Rename + numeric conversion",
            "rule": "Convert to numeric; retain missing values.",
        },
        {
            "raw_field": "Impressions",
            "clean_field": "impressions",
            "transformation": "Rename + numeric conversion",
            "rule": "Convert to numeric; retain missing values.",
        },
        {
            "raw_field": "Cost",
            "clean_field": "cost",
            "transformation": "Rename + currency conversion",
            "rule": "Remove currency symbols/commas and convert to numeric; retain missing values.",
        },
        {
            "raw_field": "Leads",
            "clean_field": "leads",
            "transformation": "Rename + numeric conversion",
            "rule": "Convert to numeric; retain missing values.",
        },
        {
            "raw_field": "Conversions",
            "clean_field": "conversions",
            "transformation": "Rename + numeric conversion",
            "rule": "Convert to numeric; retain missing values; do not infer from Sale_Amount.",
        },
        {
            "raw_field": "Conversion Rate",
            "clean_field": "conversion_rate",
            "transformation": "Rename only",
            "rule": "Retain existing values and missing values. Recalculation deferred to Phase 5.",
        },
        {
            "raw_field": "Sale_Amount",
            "clean_field": "sale_amount",
            "transformation": "Rename + currency conversion",
            "rule": "Remove currency symbols/commas and convert to numeric; retain missing values.",
        },
        {
            "raw_field": "Ad_Date",
            "clean_field": "ad_date",
            "transformation": "Rename + explicit date parsing",
            "rule": "Parse known YYYY-MM-DD, YYYY/MM/DD and DD-MM-YYYY representations.",
        },
        {
            "raw_field": "Location",
            "clean_field": "location",
            "transformation": "Rename + standardize",
            "rule": "Map approved Hyderabad variants to Hyderabad.",
        },
        {
            "raw_field": "Device",
            "clean_field": "device",
            "transformation": "Rename + case normalization",
            "rule": "Normalize approved device variants to Mobile, Desktop or Tablet.",
        },
        {
            "raw_field": "Keyword",
            "clean_field": "keyword_raw / keyword_clean",
            "transformation": "Preserve + audit field",
            "rule": "Do not arbitrarily merge analytical keyword phrases. Preserve raw values.",
        },
    ]
)


# ============================================================================
# LOADING
# ============================================================================

def load_raw_data(path: Path = RAW_DATA_PATH) -> pd.DataFrame:
    """
    Load the protected raw CSV.

    This function only reads the raw file.
    It never modifies or overwrites it.
    """
    if not path.exists():
        raise FileNotFoundError(
            f"Raw data file not found: {path}"
        )

    return pd.read_csv(path)


# ============================================================================
# STEP 3.1 — RAW BASELINE
# ============================================================================

def capture_raw_baseline(df: pd.DataFrame) -> Dict:
    """
    Capture the Phase 3 raw-data baseline before transformations.
    """
    baseline = {
        "rows": len(df),
        "columns": len(df.columns),
        "total_missing_cells": int(df.isna().sum().sum()),
        "duplicate_rows": int(df.duplicated().sum()),
        "duplicate_ad_id": (
            int(df["Ad_ID"].duplicated().sum())
            if "Ad_ID" in df.columns
            else None
        ),
        "columns_list": list(df.columns),
        "dtypes": {
            column: str(dtype)
            for column, dtype in df.dtypes.items()
        },
        "missing_by_column": {
            column: int(value)
            for column, value in df.isna().sum().items()
        },
    }

    return baseline


def validate_raw_baseline(df: pd.DataFrame) -> None:
    """
    Validate the known Phase 2 raw-data baseline.
    """
    if len(df) != EXPECTED_RAW_ROWS:
        raise ValueError(
            f"Unexpected raw row count: {len(df)}. "
            f"Expected {EXPECTED_RAW_ROWS}."
        )

    if len(df.columns) != EXPECTED_RAW_COLUMNS:
        raise ValueError(
            f"Unexpected raw column count: {len(df.columns)}. "
            f"Expected {EXPECTED_RAW_COLUMNS}."
        )

    if list(df.columns) != EXPECTED_RAW_COLUMNS_LIST:
        raise ValueError(
            "Raw column names do not match the expected Phase 3 baseline."
        )

    if df["Ad_ID"].duplicated().sum() != 0:
        raise ValueError("Duplicate Ad_ID values detected in raw data.")

    if df.duplicated().sum() != 0:
        raise ValueError("Duplicate rows detected in raw data.")


# ============================================================================
# STEP 3.2 — COLUMN NAMES
# ============================================================================

def standardize_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """
    Standardize raw column names.

    Returns a copy and does not modify the input DataFrame.
    """
    cleaned_df = df.copy()

    missing_columns = set(COLUMN_MAPPING) - set(cleaned_df.columns)

    if missing_columns:
        raise ValueError(
            f"Expected raw columns are missing: {sorted(missing_columns)}"
        )

    cleaned_df = cleaned_df.rename(columns=COLUMN_MAPPING)

    return cleaned_df


def validate_clean_column_names(df: pd.DataFrame) -> None:
    """
    Validate the standardized analytical column structure.
    """
    if list(df.columns) != EXPECTED_CLEAN_COLUMNS:
        raise ValueError(
            "Cleaned column structure does not match the expected schema.\n"
            f"Expected: {EXPECTED_CLEAN_COLUMNS}\n"
            f"Actual: {list(df.columns)}"
        )


# ============================================================================
# STEP 3.4 — CAMPAIGN STANDARDIZATION
# ============================================================================

def standardize_campaign(df: pd.DataFrame) -> pd.DataFrame:
    """
    Standardize campaign names using the approved mapping.
    """
    cleaned_df = df.copy()

    cleaned_df["campaign_name"] = (
        cleaned_df["campaign_name"]
        .map(CAMPAIGN_MAPPING)
    )

    return cleaned_df


# ============================================================================
# STEP 3.5 — LOCATION STANDARDIZATION
# ============================================================================

def standardize_location(df: pd.DataFrame) -> pd.DataFrame:
    """
    Standardize location values using the approved mapping.
    """
    cleaned_df = df.copy()

    cleaned_df["location"] = (
        cleaned_df["location"]
        .map(LOCATION_MAPPING)
    )

    return cleaned_df


# ============================================================================
# STEP 3.6 — DEVICE STANDARDIZATION
# ============================================================================

def standardize_device(df: pd.DataFrame) -> pd.DataFrame:
    """
    Standardize device values to Mobile, Desktop and Tablet.
    """
    cleaned_df = df.copy()

    cleaned_df["device"] = (
        cleaned_df["device"]
        .map(DEVICE_MAPPING)
    )

    return cleaned_df


# ============================================================================
# STEP 3.7 — KEYWORD HANDLING
# ============================================================================

def preserve_and_clean_keywords(df: pd.DataFrame) -> pd.DataFrame:
    """
    Preserve original keyword values and create a clean analytical field.

    No arbitrary keyword merging or spelling correction is performed.
    The Phase 2 audit identified keyword variation, but did not establish
    sufficient evidence for collapsing the six analytical phrases.

    Therefore:
        keyword_raw   = original keyword
        keyword_clean = normalized copy of original keyword
    """
    cleaned_df = df.copy()

    cleaned_df["keyword_raw"] = cleaned_df["keyword"]

    cleaned_df["keyword_clean"] = (
        cleaned_df["keyword"]
        .astype("string")
        .str.strip()
    )

    cleaned_df = cleaned_df.drop(columns=["keyword"])

    return cleaned_df


def validate_keyword_handling(
    raw_df: pd.DataFrame,
    cleaned_df: pd.DataFrame,
) -> None:
    """
    Validate Step 3.7 keyword preservation and cleaning.

    Rules:
    - keyword_raw must exactly preserve the original Keyword values.
    - keyword_clean may only differ from the raw value because
      surrounding whitespace has been removed.
    - No keyword merging or spelling correction is permitted.
    """

    raw_keywords = raw_df["Keyword"].astype("string")

    cleaned_raw_keywords = (
        cleaned_df["keyword_raw"].astype("string")
    )

    if not raw_keywords.equals(cleaned_raw_keywords):
        raise ValueError(
            "keyword_raw does not exactly preserve the original Keyword values."
        )

    expected_clean_keywords = raw_keywords.str.strip()

    actual_clean_keywords = (
        cleaned_df["keyword_clean"].astype("string")
    )

    if not expected_clean_keywords.equals(actual_clean_keywords):
        raise ValueError(
            "keyword_clean contains transformations beyond whitespace normalization."
        )

    if len(cleaned_df) != len(raw_df):
        raise ValueError(
            "Keyword handling changed the number of records."
        )


# ============================================================================
# STEP 3.8 — CURRENCY CLEANING
# ============================================================================

def clean_currency_column(
    series: pd.Series,
    column_name: str,
) -> pd.Series:
    """
    Convert currency-formatted strings to numeric values.

    Examples:
        '$231.88' -> 231.88
        '$1,892'  -> 1892.00

    Missing values remain missing.
    """
    cleaned = (
        series.astype("string")
        .str.strip()
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
    )

    numeric = pd.to_numeric(cleaned, errors="coerce")

    # Values that were non-null before parsing but became NaN
    # indicate a parsing failure.
    parsing_failures = series.notna() & numeric.isna()

    if parsing_failures.any():
        count = int(parsing_failures.sum())

        raise ValueError(
            f"{column_name}: {count} non-null values could not be "
            "converted to numeric."
        )

    return numeric.astype("float64")


def clean_currency_fields(df: pd.DataFrame) -> pd.DataFrame:
    """
    Clean Cost and Sale_Amount currency fields.
    """
    cleaned_df = df.copy()

    cleaned_df["cost"] = clean_currency_column(
        cleaned_df["cost"],
        "cost",
    )

    cleaned_df["sale_amount"] = clean_currency_column(
        cleaned_df["sale_amount"],
        "sale_amount",
    )

    return cleaned_df


# ============================================================================
# STEP 3.9 — DATE STANDARDIZATION
# ============================================================================

def parse_ad_dates(series: pd.Series) -> pd.Series:
    """
    Explicitly parse the known raw date representations.

    Supported representations:
        YYYY-MM-DD
        YYYY/MM/DD
        DD-MM-YYYY

    The function avoids ambiguous generic date parsing.
    """
    raw = series.astype("string").str.strip()

    result = pd.Series(
        pd.NaT,
        index=series.index,
        dtype="datetime64[ns]",
    )

    # YYYY-MM-DD
    mask_iso_dash = raw.str.fullmatch(r"\d{4}-\d{2}-\d{2}", na=False)

    if mask_iso_dash.any():
        result.loc[mask_iso_dash] = pd.to_datetime(
            raw.loc[mask_iso_dash],
            format="%Y-%m-%d",
            errors="coerce",
        )

    # YYYY/MM/DD
    mask_iso_slash = raw.str.fullmatch(r"\d{4}/\d{2}/\d{2}", na=False)

    if mask_iso_slash.any():
        result.loc[mask_iso_slash] = pd.to_datetime(
            raw.loc[mask_iso_slash],
            format="%Y/%m/%d",
            errors="coerce",
        )

    # DD-MM-YYYY
    mask_day_first = raw.str.fullmatch(r"\d{2}-\d{2}-\d{4}", na=False)

    if mask_day_first.any():
        result.loc[mask_day_first] = pd.to_datetime(
            raw.loc[mask_day_first],
            format="%d-%m-%Y",
            errors="coerce",
        )

    # Any non-null value that remains NaT is an unsupported/invalid date.
    invalid_dates = raw.notna() & result.isna()

    if invalid_dates.any():
        examples = raw.loc[invalid_dates].drop_duplicates().head(10).tolist()

        raise ValueError(
            "Unsupported or invalid Ad_Date values found: "
            f"{examples}"
        )

    return result


def standardize_dates(df: pd.DataFrame) -> pd.DataFrame:
    """
    Standardize Ad_Date into pandas datetime.
    """
    cleaned_df = df.copy()

    cleaned_df["ad_date"] = parse_ad_dates(
        cleaned_df["ad_date"]
    )

    return cleaned_df


# ============================================================================
# NUMERIC FIELD CLEANING
# ============================================================================

NUMERIC_COLUMNS = [
    "clicks",
    "impressions",
    "leads",
    "conversions",
]


def convert_numeric_fields(df: pd.DataFrame) -> pd.DataFrame:
    """
    Convert count fields to numeric.

    Missing values are retained.
    No imputation is performed.
    """
    cleaned_df = df.copy()

    for column in NUMERIC_COLUMNS:
        original_non_null = cleaned_df[column].notna()

        cleaned_df[column] = pd.to_numeric(
            cleaned_df[column],
            errors="coerce",
        )

        parsing_failures = (
            original_non_null
            & cleaned_df[column].isna()
        )

        if parsing_failures.any():
            count = int(parsing_failures.sum())

            raise ValueError(
                f"{column}: {count} non-null values could not "
                "be converted to numeric."
            )

    # Conversion Rate is intentionally not recalculated.
    # It is only converted/retained as numeric.
    cleaned_df["conversion_rate"] = pd.to_numeric(
        cleaned_df["conversion_rate"],
        errors="coerce",
    )

    return cleaned_df


# ============================================================================
# STEP 3.10 — MISSING-VALUE TREATMENT
# ============================================================================

def apply_missing_value_policy(df: pd.DataFrame) -> pd.DataFrame:
    """
    Apply the Phase 3 missing-value policy.

    IMPORTANT:
    No automatic statistical or business imputation is performed.

    Specifically:
    - conversion_rate missing values remain missing.
    - sale_amount missing values remain missing.
    - conversions missing values remain missing.
    - clicks missing values remain missing.
    - impressions missing values remain missing.
    - leads missing values remain missing.
    - cost missing values remain missing.

    Relationships such as:
        conversions > 0 and sale_amount missing
        sale_amount > 0 and conversions missing

    are preserved for investigation rather than inferred.
    """
    cleaned_df = df.copy()

    # Intentionally no fillna() operations here.

    return cleaned_df


# ============================================================================
# BUSINESS-RELATIONSHIP AUDIT
# ============================================================================

def create_missing_value_audit(df: pd.DataFrame) -> pd.DataFrame:
    """
    Create a row-level missing-value treatment audit.

    This documents important relationships identified during Phase 2.
    """
    audit = pd.DataFrame(
        {
            "sale_amount_missing": df["sale_amount"].isna(),
            "conversions_missing": df["conversions"].isna(),
        }
    )

    audit["conversions_positive_sale_amount_missing"] = (
        df["conversions"].gt(0)
        & df["sale_amount"].isna()
    )

    audit["sale_amount_positive_conversions_missing"] = (
        df["sale_amount"].gt(0)
        & df["conversions"].isna()
    )

    summary = pd.DataFrame(
        {
            "condition": [
                "Sale_Amount missing",
                "Conversions missing",
                "Conversions > 0 and Sale_Amount missing",
                "Sale_Amount > 0 and Conversions missing",
                "Conversion Rate missing",
                "Clicks missing",
                "Impressions missing",
                "Leads missing",
                "Cost missing",
            ],
            "record_count": [
                int(df["sale_amount"].isna().sum()),
                int(df["conversions"].isna().sum()),
                int(
                    (
                        df["conversions"].gt(0)
                        & df["sale_amount"].isna()
                    ).sum()
                ),
                int(
                    (
                        df["sale_amount"].gt(0)
                        & df["conversions"].isna()
                    ).sum()
                ),
                int(df["conversion_rate"].isna().sum()),
                int(df["clicks"].isna().sum()),
                int(df["impressions"].isna().sum()),
                int(df["leads"].isna().sum()),
                int(df["cost"].isna().sum()),
            ],
        }
    )

    return summary


# ============================================================================
# VALIDATION
# ============================================================================

def validate_no_unexpected_categories(df: pd.DataFrame) -> None:
    """
    Ensure approved categorical transformations succeeded.
    """
    expected_campaign = {"Data Analytics Course"}
    expected_location = {"Hyderabad"}
    expected_device = {"Mobile", "Desktop", "Tablet"}

    actual_campaign = set(
        df["campaign_name"].dropna().unique()
    )

    actual_location = set(
        df["location"].dropna().unique()
    )

    actual_device = set(
        df["device"].dropna().unique()
    )

    if not actual_campaign.issubset(expected_campaign):
        raise ValueError(
            f"Unexpected campaign values: "
            f"{sorted(actual_campaign - expected_campaign)}"
        )

    if not actual_location.issubset(expected_location):
        raise ValueError(
            f"Unexpected location values: "
            f"{sorted(actual_location - expected_location)}"
        )

    if not actual_device.issubset(expected_device):
        raise ValueError(
            f"Unexpected device values: "
            f"{sorted(actual_device - expected_device)}"
        )


def validate_dates(df: pd.DataFrame) -> None:
    """
    Validate standardized dates.
    """
    if not pd.api.types.is_datetime64_any_dtype(df["ad_date"]):
        raise TypeError("ad_date is not a datetime dtype.")

    nat_count = int(df["ad_date"].isna().sum())

    if nat_count != 0:
        raise ValueError(
            f"Invalid/missing standardized dates detected: {nat_count}"
        )

    min_date = df["ad_date"].min()
    max_date = df["ad_date"].max()

    expected_min = pd.Timestamp("2024-11-01")
    expected_max = pd.Timestamp("2024-11-30")

    if min_date != expected_min:
        raise ValueError(
            f"Unexpected minimum date: {min_date}. "
            f"Expected {expected_min}."
        )

    if max_date != expected_max:
        raise ValueError(
            f"Unexpected maximum date: {max_date}. "
            f"Expected {expected_max}."
        )


def validate_numeric_fields(df: pd.DataFrame) -> None:
    """
    Validate numeric analytical fields.
    """
    numeric_columns = [
        "clicks",
        "impressions",
        "cost",
        "leads",
        "conversions",
        "conversion_rate",
        "sale_amount",
    ]

    for column in numeric_columns:
        if not pd.api.types.is_numeric_dtype(df[column]):
            raise TypeError(
                f"{column} is not numeric after cleaning."
            )


def validate_currency_values(df: pd.DataFrame) -> None:
    """
    Validate that currency fields no longer contain currency symbols
    or comma-formatted strings.
    """
    for column in ["cost", "sale_amount"]:
        if df[column].dtype == "object":
            raise TypeError(
                f"{column} remains object dtype after currency cleaning."
            )

        if df[column].astype("string").str.contains(
            r"[$,]",
            regex=True,
            na=False,
        ).any():
            raise ValueError(
                f"{column} still contains currency symbols or commas."
            )


def validate_identifier_integrity(
    raw_df: pd.DataFrame,
    cleaned_df: pd.DataFrame,
) -> None:
    """
    Ensure row count and Ad_ID values remain unchanged.
    """
    if len(raw_df) != len(cleaned_df):
        raise ValueError(
            "Row count changed during cleaning."
        )

    if raw_df["Ad_ID"].tolist() != cleaned_df["ad_id"].tolist():
        raise ValueError(
            "Ad_ID values changed during cleaning."
        )

    if cleaned_df["ad_id"].duplicated().any():
        raise ValueError(
            "Duplicate ad_id values detected after cleaning."
        )


def validate_cleaned_data(
    raw_df: pd.DataFrame,
    cleaned_df: pd.DataFrame,
) -> None:
    """
    Run all Phase 3 validation checks.
    """
    validate_clean_column_names(cleaned_df)
    validate_identifier_integrity(raw_df, cleaned_df)
    validate_no_unexpected_categories(cleaned_df)
    validate_keyword_handling(
        raw_df,
        cleaned_df,
    )
    validate_dates(cleaned_df)
    validate_numeric_fields(cleaned_df)
    validate_currency_values(cleaned_df)

    if len(cleaned_df) != EXPECTED_RAW_ROWS:
        raise ValueError(
            "Cleaned row count does not match raw baseline."
        )

    if cleaned_df.duplicated().any():
        raise ValueError(
            "Duplicate rows detected after cleaning."
        )


# ============================================================================
# BEFORE / AFTER RECONCILIATION
# ============================================================================

def create_before_after_summary(
    raw_df: pd.DataFrame,
    cleaned_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Create the Phase 3 before/after reconciliation table.
    """
    summary = pd.DataFrame(
        [
            {
                "check": "Rows",
                "raw": len(raw_df),
                "cleaned": len(cleaned_df),
            },
            {
                "check": "Columns",
                "raw": len(raw_df.columns),
                "cleaned": len(cleaned_df.columns),
            },
            {
                "check": "Duplicate rows",
                "raw": int(raw_df.duplicated().sum()),
                "cleaned": int(cleaned_df.duplicated().sum()),
            },
            {
                "check": "Duplicate Ad_ID",
                "raw": int(raw_df["Ad_ID"].duplicated().sum()),
                "cleaned": int(cleaned_df["ad_id"].duplicated().sum()),
            },
            {
                "check": "Campaign unique values",
                "raw": raw_df["Campaign_Name"].nunique(dropna=True),
                "cleaned": cleaned_df["campaign_name"].nunique(
                    dropna=True
                ),
            },
            {
                "check": "Location unique values",
                "raw": raw_df["Location"].nunique(dropna=True),
                "cleaned": cleaned_df["location"].nunique(
                    dropna=True
                ),
            },
            {
                "check": "Device unique values",
                "raw": raw_df["Device"].nunique(dropna=True),
                "cleaned": cleaned_df["device"].nunique(
                    dropna=True
                ),
            },
            {
                "check": "Keyword unique values",
                "raw": raw_df["Keyword"].nunique(dropna=True),
                "cleaned": cleaned_df["keyword_clean"].nunique(
                    dropna=True
                ),
            },
            {
                "check": "Missing cells",
                "raw": int(raw_df.isna().sum().sum()),
                "cleaned": int(cleaned_df.isna().sum().sum()),
            },
        ]
    )

    return summary


# ============================================================================
# AUDIT OUTPUTS
# ============================================================================

def save_cleaning_audits(
    raw_df: pd.DataFrame,
    cleaned_df: pd.DataFrame,
) -> None:
    """
    Save Phase 3 audit artifacts.
    """
    CLEANING_REPORT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    CLEANING_DICTIONARY.to_csv(
        CLEANING_REPORT_DIR / "cleaning_dictionary.csv",
        index=False,
    )

    create_before_after_summary(
        raw_df,
        cleaned_df,
    ).to_csv(
        CLEANING_REPORT_DIR / "before_after_reconciliation.csv",
        index=False,
    )

    create_missing_value_audit(
        cleaned_df,
    ).to_csv(
        CLEANING_REPORT_DIR / "missing_value_treatment.csv",
        index=False,
    )

    category_audit = pd.DataFrame(
        {
            "dimension": [
                "campaign_name",
                "location",
                "device",
                "keyword_clean",
            ],
            "unique_values_after_cleaning": [
                cleaned_df["campaign_name"].nunique(
                    dropna=True
                ),
                cleaned_df["location"].nunique(
                    dropna=True
                ),
                cleaned_df["device"].nunique(
                    dropna=True
                ),
                cleaned_df["keyword_clean"].nunique(
                    dropna=True
                ),
            ],
        }
    )

    category_audit.to_csv(
        CLEANING_REPORT_DIR / "before_after_categories.csv",
        index=False,
    )


# ============================================================================
# STEP 3.14 — EXCEL AUDIT
# ============================================================================

def save_excel_audits(
    raw_df: pd.DataFrame,
    cleaned_df: pd.DataFrame,
    raw_baseline: Dict,
    path: Path = EXCEL_WORKBOOK_PATH,
) -> None:
    """
    Add the Phase 3 audit sheets to the existing Excel workbook.

    Existing Phase 2 sheets are preserved. Only the six Phase 3 sheets
    managed by this function are replaced on repeat runs.
    """
    phase_3_sheets = [
        "Cleaning_Dictionary",
        "Cleaning_Summary",
        "Before_After_Categories",
        "Missing_Value_Treatment",
        "Date_Cleaning_Audit",
        "Currency_Cleaning_Audit",
    ]

    workbook = load_workbook(path)

    for sheet_name in phase_3_sheets:
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]

    def add_dataframe(sheet_name: str, dataframe: pd.DataFrame) -> None:
        worksheet = workbook.create_sheet(sheet_name)

        for column_index, column_name in enumerate(dataframe.columns, start=1):
            worksheet.cell(
                row=1,
                column=column_index,
                value=column_name,
            )

        for row_index, row_values in enumerate(
            dataframe.itertuples(index=False, name=None),
            start=2,
        ):
            for column_index, value in enumerate(row_values, start=1):
                if pd.isna(value):
                    value = None

                worksheet.cell(
                    row=row_index,
                    column=column_index,
                    value=value,
                )

    cleaning_summary = pd.DataFrame(
        {
            "Metric": [
                "Rows",
                "Columns",
                "Missing cells",
                "Duplicate rows",
                "Duplicate Ad_ID",
            ],
            "Raw": [
                raw_baseline["rows"],
                raw_baseline["columns"],
                raw_baseline["total_missing_cells"],
                raw_baseline["duplicate_rows"],
                raw_baseline["duplicate_ad_id"],
            ],
            "Cleaned": [
                len(cleaned_df),
                len(cleaned_df.columns),
                int(cleaned_df.isna().sum().sum()),
                int(cleaned_df.duplicated().sum()),
                int(cleaned_df["ad_id"].duplicated().sum()),
            ],
        }
    )

    before_after_categories = pd.DataFrame(
        {
            "Dimension": [
                "Campaign",
                "Location",
                "Device",
                "Keyword",
            ],
            "Raw Unique Values": [
                raw_df["Campaign_Name"].nunique(dropna=True),
                raw_df["Location"].nunique(dropna=True),
                raw_df["Device"].nunique(dropna=True),
                raw_df["Keyword"].nunique(dropna=True),
            ],
            "Cleaned Unique Values": [
                cleaned_df["campaign_name"].nunique(dropna=True),
                cleaned_df["location"].nunique(dropna=True),
                cleaned_df["device"].nunique(dropna=True),
                cleaned_df["keyword_clean"].nunique(dropna=True),
            ],
        }
    )

    missing_value_treatment = create_missing_value_audit(cleaned_df)

    raw_dates = raw_df["Ad_Date"].astype("string").str.strip()
    date_formats = {
        "YYYY-MM-DD": (r"\d{4}-\d{2}-\d{2}", "%Y-%m-%d"),
        "YYYY/MM/DD": (r"\d{4}/\d{2}/\d{2}", "%Y/%m/%d"),
        "DD-MM-YYYY": (r"\d{2}-\d{2}-\d{4}", "%d-%m-%Y"),
    }
    parsed_dates = parse_ad_dates(raw_df["Ad_Date"])
    date_cleaning_rows = []

    for raw_format, (pattern, parsing_format) in date_formats.items():
        format_mask = raw_dates.str.fullmatch(pattern, na=False)
        format_count = int(format_mask.sum())
        format_parsed = parsed_dates.loc[format_mask]
        parse_failures = int(format_mask.sum() - format_parsed.notna().sum())

        date_cleaning_rows.append(
            {
                "Raw date format": raw_format,
                "Record count": format_count,
                "Parsing format": parsing_format,
                "Parsing result": "Success" if parse_failures == 0 else "Failure",
                "Minimum cleaned date": format_parsed.min(),
                "Maximum cleaned date": format_parsed.max(),
                "NaT count": parse_failures,
            }
        )

    unsupported_mask = ~raw_dates.str.fullmatch(
        r"\d{4}-\d{2}-\d{2}|\d{4}/\d{2}/\d{2}|\d{2}-\d{2}-\d{4}",
        na=False,
    )
    unsupported_count = int(unsupported_mask.sum())
    date_cleaning_rows.append(
        {
            "Raw date format": "Other/unsupported",
            "Record count": unsupported_count,
            "Parsing format": "Not applicable",
            "Parsing result": "None" if unsupported_count == 0 else "Review",
            "Minimum cleaned date": None,
            "Maximum cleaned date": None,
            "NaT count": unsupported_count,
        }
    )

    date_cleaning_audit = pd.DataFrame(date_cleaning_rows)

    currency_cleaning_rows = []
    for raw_column, clean_column in [
        ("Cost", "cost"),
        ("Sale_Amount", "sale_amount"),
    ]:
        raw_values = raw_df[raw_column]
        cleaned_values = cleaned_df[clean_column]
        currency_cleaning_rows.append(
            {
                "Field": clean_column,
                "Raw Non-Null Count": int(raw_values.notna().sum()),
                "Raw Missing Count": int(raw_values.isna().sum()),
                "Cleaned Missing Count": int(cleaned_values.isna().sum()),
                "Cleaned Data Type": str(cleaned_values.dtype),
                "Remaining Currency Symbols": int(
                    cleaned_values.astype("string")
                    .str.contains(r"\$", regex=True, na=False)
                    .sum()
                ),
                "Remaining Commas": int(
                    cleaned_values.astype("string")
                    .str.contains(",", regex=False, na=False)
                    .sum()
                ),
                "Negative Values": int((cleaned_values < 0).sum()),
                "Parsing Failures": int(
                    raw_values.notna().sum() - cleaned_values.notna().sum()
                ),
            }
        )

    currency_cleaning_audit = pd.DataFrame(currency_cleaning_rows)

    add_dataframe("Cleaning_Dictionary", CLEANING_DICTIONARY)
    add_dataframe("Cleaning_Summary", cleaning_summary)
    add_dataframe("Before_After_Categories", before_after_categories)
    add_dataframe("Missing_Value_Treatment", missing_value_treatment)
    add_dataframe("Date_Cleaning_Audit", date_cleaning_audit)
    add_dataframe("Currency_Cleaning_Audit", currency_cleaning_audit)

    workbook.save(path)

    expected_sheets = [
        "Audit_Readme",
        "Data_Profile",
        "Missing_Values",
        "Duplicates",
        "Categories",
        "Date_Audit",
        "Numeric_Audit",
        "Outliers",
        "Data_Quality_Issue_Log",
        *phase_3_sheets,
    ]
    final_workbook = load_workbook(path, read_only=True)

    if final_workbook.sheetnames != expected_sheets:
        raise ValueError(
            "Excel workbook sheets do not match the expected Phase 2 and "
            "Phase 3 audit sheets."
        )


# ============================================================================
# STEP 3.13 — SAVE CLEANED DATA
# ============================================================================

def save_cleaned_data(
    df: pd.DataFrame,
    path: Path = CLEANED_DATA_PATH,
) -> None:
    """
    Save the final cleaned dataset.

    The destination is always under data/processed/.
    """
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    df.to_csv(
        path,
        index=False,
    )


# ============================================================================
# COMPLETE PHASE 3 PIPELINE
# ============================================================================

def clean_data(
    raw_path: Path = RAW_DATA_PATH,
) -> Tuple[pd.DataFrame, pd.DataFrame, Dict]:
    """
    Execute the complete Phase 3 cleaning pipeline.

    Returns:
        raw_df
        cleaned_df
        raw_baseline
    """
    # ------------------------------------------------------------------------
    # Step 3.1 — Load and freeze baseline
    # ------------------------------------------------------------------------

    raw_df = load_raw_data(raw_path)

    validate_raw_baseline(raw_df)

    raw_baseline = capture_raw_baseline(raw_df)

    # ------------------------------------------------------------------------
    # Create independent working copy
    # ------------------------------------------------------------------------

    cleaned_df = raw_df.copy()

    # ------------------------------------------------------------------------
    # Step 3.2 — Column names
    # ------------------------------------------------------------------------

    cleaned_df = standardize_column_names(
        cleaned_df
    )

    # ------------------------------------------------------------------------
    # Step 3.4 — Campaign
    # ------------------------------------------------------------------------

    cleaned_df = standardize_campaign(
        cleaned_df
    )

    # ------------------------------------------------------------------------
    # Step 3.5 — Location
    # ------------------------------------------------------------------------

    cleaned_df = standardize_location(
        cleaned_df
    )

    # ------------------------------------------------------------------------
    # Step 3.6 — Device
    # ------------------------------------------------------------------------

    cleaned_df = standardize_device(
        cleaned_df
    )

    # ------------------------------------------------------------------------
    # Step 3.7 — Keywords
    # ------------------------------------------------------------------------

    cleaned_df = preserve_and_clean_keywords(
        cleaned_df
    )

    # ------------------------------------------------------------------------
    # Step 3.8 — Currency
    # ------------------------------------------------------------------------

    cleaned_df = clean_currency_fields(
        cleaned_df
    )

    # ------------------------------------------------------------------------
    # Numeric fields
    # ------------------------------------------------------------------------

    cleaned_df = convert_numeric_fields(
        cleaned_df
    )

    # ------------------------------------------------------------------------
    # Step 3.9 — Dates
    # ------------------------------------------------------------------------

    cleaned_df = standardize_dates(
        cleaned_df
    )

    # ------------------------------------------------------------------------
    # Step 3.10 — Missing values
    # ------------------------------------------------------------------------

    cleaned_df = apply_missing_value_policy(
        cleaned_df
    )

    # ------------------------------------------------------------------------
    # Final validation
    # ------------------------------------------------------------------------

    validate_cleaned_data(
        raw_df,
        cleaned_df,
    )

    # ------------------------------------------------------------------------
    # Audit outputs
    # ------------------------------------------------------------------------

    save_cleaning_audits(
        raw_df,
        cleaned_df,
    )

    save_excel_audits(
        raw_df,
        cleaned_df,
        raw_baseline,
    )

    # ------------------------------------------------------------------------
    # Save cleaned dataset
    # ------------------------------------------------------------------------

    save_cleaned_data(
        cleaned_df,
    )

    return raw_df, cleaned_df, raw_baseline


# ============================================================================
# COMMAND-LINE ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    raw_df, cleaned_df, baseline = clean_data()

    print("=" * 70)
    print("PHASE 3 DATA CLEANING COMPLETED")
    print("=" * 70)

    print("\nRAW BASELINE")
    print(f"Rows:              {baseline['rows']}")
    print(f"Columns:           {baseline['columns']}")
    print(f"Missing cells:     {baseline['total_missing_cells']}")
    print(f"Duplicate rows:    {baseline['duplicate_rows']}")
    print(f"Duplicate Ad_ID:   {baseline['duplicate_ad_id']}")

    print("\nCLEANED DATA")
    print(f"Rows:              {len(cleaned_df)}")
    print(f"Columns:           {len(cleaned_df.columns)}")
    print(
        f"Missing cells:     "
        f"{int(cleaned_df.isna().sum().sum())}"
    )

    print("\nCLEANED COLUMNS")
    for column in cleaned_df.columns:
        print(f"- {column}")

    print("\nOUTPUT")
    print(f"Cleaned CSV:       {CLEANED_DATA_PATH}")
    print(f"Cleaning reports:  {CLEANING_REPORT_DIR}")

    print("\nValidation: PASSED")