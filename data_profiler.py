"""
data_profiler.py
Backend logic for the "Data Profiler" sidebar page in LinkHarvest.

Responsibilities:
- Load CSV/Excel files into a DataFrame
- Detect data-quality anomalies (missing values, duplicates, blank rows,
  extra spaces, special characters, mixed data types)
- Clean the data based on a user-selected subset of those anomaly filters
- Save the cleaned result back to CSV/Excel

No UI code lives here. Keep this importable and testable on its own.
"""

import os
import re
import pandas as pd


# Characters that are allowed inside normal text and should NOT be flagged
# as "special characters" (basic punctuation, currency, etc.)
_ALLOWED_PUNCTUATION = set(" .,-_/()&'\":;!?%$#@+*=[]{}|\\<>~`^")


def load_data(filepath: str) -> pd.DataFrame:
    """
    Load a CSV or Excel file into a DataFrame.
    Raises ValueError with a consistent, readable message on any failure.
    """
    if not os.path.exists(filepath):
        raise ValueError(f"Could not read file: file not found: {filepath}")

    ext = os.path.splitext(filepath)[1].lower()

    try:
        if ext == ".csv":
            df = pd.read_csv(filepath)
        elif ext in (".xlsx", ".xls"):
            df = pd.read_excel(filepath)
        else:
            raise ValueError(
                f"Could not read file: unsupported file type '{ext}'. "
                "Please provide a .csv, .xlsx, or .xls file."
            )
    except ValueError as e:
        # Re-raise our own "Could not read file:" errors as-is, but wrap
        # pandas' own exceptions (EmptyDataError is a ValueError subclass,
        # ParserError, etc.) into the same consistent format.
        msg = str(e)
        if msg.startswith("Could not read file:"):
            raise
        raise ValueError(f"Could not read file: {msg}")
    except Exception as e:
        raise ValueError(f"Could not read file: {e}")

    if df.empty and df.columns.empty:
        raise ValueError("Could not read file: the file contains no data.")

    return df


def _is_special_char_string(value: str) -> bool:
    """True if the string contains a character outside letters/digits/allowed punctuation."""
    for ch in value:
        if ch.isalnum() or ch.isspace() or ch in _ALLOWED_PUNCTUATION:
            continue
        return True
    return False


def analyze_data(df: pd.DataFrame) -> dict:
    """
    Run all anomaly detectors against the DataFrame and return a dict of
    {anomaly_key: count}. Used to populate the KPI cards, the checkbox
    labels ("Missing Values (142 cases)"), and the donut chart.
    """
    results = {}

    # 1. Missing values -> count of individual NaN/None cells
    results["missing_values"] = int(df.isna().sum().sum())

    # 1b. Of those, how many sit in numeric columns specifically (these are
    #     the ones eligible for the "fill with column mean" option below —
    #     it never applies to text/categorical columns).
    numeric_cols = df.select_dtypes(include="number").columns
    results["missing_values_numeric"] = int(df[numeric_cols].isna().sum().sum()) if len(numeric_cols) else 0

    # 2. Duplicate rows (excluding the first occurrence of each dupe set)
    results["duplicate_rows"] = int(df.duplicated(keep="first").sum())

    # 3. Blank rows -> every cell in the row is NaN or an empty/whitespace string
    def _row_is_blank(row):
        for v in row:
            if pd.isna(v):
                continue
            if isinstance(v, str) and v.strip() == "":
                continue
            return False
        return True

    results["blank_rows"] = int(df.apply(_row_is_blank, axis=1).sum())

    # 4. Extra spaces -> string cells with leading/trailing whitespace or
    #    internal double-spaces
    extra_spaces = 0
    special_chars = 0
    for col in df.select_dtypes(include="object").columns:
        for v in df[col].dropna():
            if not isinstance(v, str):
                continue
            if v != v.strip() or "  " in v:
                extra_spaces += 1
            if _is_special_char_string(v):
                special_chars += 1
    results["extra_spaces"] = extra_spaces
    results["special_characters"] = special_chars

    # 5. Mixed data types -> columns that contain more than one non-null
    #    Python type (e.g. a numeric column with a stray 'abc' string)
    mixed_type_cols = 0
    for col in df.columns:
        non_null = df[col].dropna()
        types_seen = {type(v) for v in non_null}
        # int/float mixing is normal (e.g. 3 and 3.0), don't flag that
        types_seen -= {int, float} if types_seen <= {int, float} else set()
        if len(types_seen) > 1:
            mixed_type_cols += 1
    results["mixed_data_types"] = mixed_type_cols

    results["total_records"] = int(len(df))
    results["total_columns"] = int(len(df.columns))
    total_anomalies = sum(
        v for k, v in results.items()
        # "missing_values_numeric" is a subset of "missing_values", not a
        # separate anomaly type, so it must not be double-counted here.
        if k not in ("total_records", "total_columns", "missing_values_numeric")
    )
    results["total_anomalies"] = total_anomalies

    # Health score: share of "clean" cells across the whole sheet.
    total_cells = max(results["total_records"] * results["total_columns"], 1)
    flawed_cells = min(
        results["missing_values"] + results["extra_spaces"] + results["special_characters"],
        total_cells,
    )
    results["health_pct"] = round(100 * (1 - flawed_cells / total_cells))

    return results


def clean_data(df: pd.DataFrame, filters: dict) -> pd.DataFrame:
    """
    Apply only the anomaly fixes whose key is True in `filters`.
    filters keys mirror analyze_data()'s keys, e.g.:
        {"missing_values": True, "duplicate_rows": True, "blank_rows": False,
         "extra_spaces": True, "special_characters": False}

    Ambiguous cases (e.g. a stray non-numeric value sitting in a numeric
    column -- "mixed_data_types") are never auto-fixed; they're left in
    place so the user can review them manually.
    """
    cleaned = df.copy()

    if filters.get("blank_rows"):
        def _row_is_blank(row):
            for v in row:
                if pd.isna(v):
                    continue
                if isinstance(v, str) and v.strip() == "":
                    continue
                return False
            return True
        cleaned = cleaned[~cleaned.apply(_row_is_blank, axis=1)]

    if filters.get("duplicate_rows"):
        cleaned = cleaned.drop_duplicates(keep="first")

    if filters.get("extra_spaces"):
        for col in cleaned.select_dtypes(include="object").columns:
            cleaned[col] = cleaned[col].apply(
                lambda v: re.sub(r"\s+", " ", v.strip()) if isinstance(v, str) else v
            )

    if filters.get("special_characters"):
        def _strip_special(v):
            if not isinstance(v, str):
                return v
            return "".join(
                ch for ch in v
                if ch.isalnum() or ch.isspace() or ch in _ALLOWED_PUNCTUATION
            )
        for col in cleaned.select_dtypes(include="object").columns:
            cleaned[col] = cleaned[col].apply(_strip_special)

    if filters.get("fill_missing_numeric_mean"):
        # Numeric-only: fill missing cells with that column's mean. Never
        # touches text/categorical columns — an ambiguous string column has
        # no valid "average" to fall back on, so those are left for the
        # missing_values / manual-review path below.
        for col in cleaned.select_dtypes(include="number").columns:
            if cleaned[col].isna().any():
                mean_val = cleaned[col].mean()
                if pd.notna(mean_val):
                    cleaned[col] = cleaned[col].fillna(round(mean_val, 2))

    if filters.get("missing_values"):
        # Drop rows that still contain any missing value after the above fixes
        # (numeric columns may already be filled by fill_missing_numeric_mean
        # above, so this only removes rows still missing a non-numeric value —
        # unless that option wasn't used, in which case it behaves as before).
        cleaned = cleaned.dropna(how="any")

    return cleaned.reset_index(drop=True)


def save_data(df: pd.DataFrame, filepath: str) -> None:
    """Save a DataFrame to CSV or Excel based on the file extension."""
    ext = os.path.splitext(filepath)[1].lower()
    try:
        if ext == ".csv":
            df.to_csv(filepath, index=False)
        elif ext in (".xlsx", ".xls"):
            df.to_excel(filepath, index=False)
        else:
            raise ValueError(f"Unsupported export type '{ext}'. Use .csv or .xlsx.")
    except ValueError:
        raise
    except Exception as e:
        raise ValueError(f"Could not save file: {e}")


# Human-readable labels for the checkbox list / donut chart legend.
ANOMALY_LABELS = {
    "missing_values": "Missing Values",
    "duplicate_rows": "Duplicate Rows",
    "blank_rows": "Blank Rows",
    "extra_spaces": "Extra Spaces",
    "special_characters": "Special Characters",
    "mixed_data_types": "Mixed Data Types",
}
