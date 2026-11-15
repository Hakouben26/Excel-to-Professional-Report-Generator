
import pandas as pd
import numpy as np


# ==========================================
# 1. EXCEL FILE READING
# ==========================================

def get_excel_sheets(file):
    """
    Return all sheet names from an Excel file.
    """
    try:
        excel_file = pd.ExcelFile(file)
        return excel_file.sheet_names

    except Exception as error:
        raise ValueError(f"Unable to read Excel file: {error}")


def load_excel_file(file, sheet_name=0):
    """
    Load a selected Excel worksheet into a DataFrame.
    """
    try:
        df = pd.read_excel(
            file,
            sheet_name=sheet_name,
            engine="openpyxl"
        )

        if df.empty:
            raise ValueError("The selected worksheet is empty.")

        return df

    except Exception as error:
        raise ValueError(f"Excel loading failed: {error}")


# ==========================================
# 2. DATA CLEANING
# ==========================================

def clean_dataframe(df):
    """
    Clean a DataFrame without automatically
    replacing missing values.
    """

    cleaned_df = df.copy()

    original_rows = len(cleaned_df)
    original_columns = len(cleaned_df.columns)

    # Remove completely empty rows
    cleaned_df = cleaned_df.dropna(how="all")

    # Remove completely empty columns
    cleaned_df = cleaned_df.dropna(axis=1, how="all")

    # Clean column names
    cleaned_df.columns = [
        str(column).strip()
        for column in cleaned_df.columns
    ]

    # Remove leading and trailing spaces from text
    for column in cleaned_df.select_dtypes(
        include=["object", "string"]
    ).columns:

        cleaned_df[column] = cleaned_df[column].apply(
            lambda value: value.strip()
            if isinstance(value, str)
            else value
        )

    # Count duplicate rows before removing them
    duplicate_count = int(cleaned_df.duplicated().sum())

    # Remove duplicate rows
    cleaned_df = cleaned_df.drop_duplicates()

    # Reset index
    cleaned_df = cleaned_df.reset_index(drop=True)

    cleaning_report = {
        "original_rows": original_rows,
        "final_rows": len(cleaned_df),
        "original_columns": original_columns,
        "final_columns": len(cleaned_df.columns),
        "duplicates_removed": duplicate_count,
        "empty_rows_removed": original_rows - len(df.dropna(how="all"))
    }

    return cleaned_df, cleaning_report


# ==========================================
# 3. DATA QUALITY ANALYSIS
# ==========================================

def analyze_data_quality(df):
    """
    Analyze missing values, duplicates,
    data types and dataset quality.
    """

    total_cells = df.size

    missing_values = df.isnull().sum()

    total_missing = int(missing_values.sum())

    missing_percentage = (
        (total_missing / total_cells) * 100
        if total_cells > 0
        else 0
    )

    duplicate_rows = int(df.duplicated().sum())

    quality_report = {
        "total_rows": len(df),
        "total_columns": len(df.columns),
        "total_cells": total_cells,
        "missing_values": total_missing,
        "missing_percentage": round(missing_percentage, 2),
        "duplicate_rows": duplicate_rows,
        "numeric_columns": len(
            df.select_dtypes(include=np.number).columns
        ),
        "categorical_columns": len(
            df.select_dtypes(include=["object", "string", "category"]).columns
        ),
        "datetime_columns": len(
            df.select_dtypes(include=["datetime", "datetimetz"]).columns
        )
    }

    missing_details = pd.DataFrame({
        "Column": missing_values.index,
        "Missing Values": missing_values.values,
        "Missing Percentage": (
            missing_values.values / len(df) * 100
            if len(df) > 0
            else 0
        )
    })

    missing_details = missing_details[
        missing_details["Missing Values"] > 0
    ].sort_values(
        by="Missing Values",
        ascending=False
    )

    return quality_report, missing_details


# ==========================================
# 4. NUMERICAL ANALYSIS
# ==========================================

def analyze_numerical_columns(df):
    """
    Generate descriptive statistics
    for numerical columns.
    """

    numerical_df = df.select_dtypes(include=np.number)

    if numerical_df.empty:
        return pd.DataFrame()

    statistics = numerical_df.describe().T

    statistics["missing"] = numerical_df.isnull().sum()

    return statistics.round(2)


# ==========================================
# 5. CATEGORICAL ANALYSIS
# ==========================================

def analyze_categorical_columns(df):
    """
    Analyze categorical columns.
    """

    categorical_df = df.select_dtypes(
        include=["object", "string", "category"]
    )

    results = []

    for column in categorical_df.columns:

        results.append({
            "Column": column,
            "Unique Values": int(df[column].nunique()),
            "Missing Values": int(df[column].isnull().sum()),
            "Most Frequent Value": (
                df[column].mode().iloc[0]
                if not df[column].mode().empty
                else "N/A"
            )
        })

    return pd.DataFrame(results)


# ==========================================
# 6. DATASET OVERVIEW
# ==========================================

def generate_dataset_overview(df):
    """
    Generate a complete dataset overview.
    """

    quality_report, missing_details = analyze_data_quality(df)

    numerical_statistics = analyze_numerical_columns(df)

    categorical_statistics = analyze_categorical_columns(df)

    overview = {
        "quality": quality_report,
        "missing_details": missing_details,
        "numerical_statistics": numerical_statistics,
        "categorical_statistics": categorical_statistics,
        "columns": list(df.columns),
        "data_types": df.dtypes.astype(str).to_dict()
    }

    return overview
