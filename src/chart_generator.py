
from pathlib import Path
import re

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


def _safe_filename(value):
    """Convert a column name into a safe filename."""
    filename = re.sub(r"[^a-zA-Z0-9_-]+", "_", str(value))
    return filename.strip("_").lower() or "chart"


def _save_chart(fig, output_path):
    """Save a chart as a high-resolution PNG image."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    fig.tight_layout()
    fig.savefig(
        output_path,
        dpi=180,
        bbox_inches="tight",
        facecolor="white"
    )
    plt.close(fig)

    return str(output_path)


# ==========================================
# 1. NUMERICAL DISTRIBUTION
# ==========================================

def create_numeric_distribution_chart(
    df,
    column,
    output_dir="reports/charts"
):
    """Create a histogram for a numerical column."""

    values = pd.to_numeric(df[column], errors="coerce").dropna()

    if values.empty:
        return None

    fig, ax = plt.subplots(figsize=(8, 4.5))

    ax.hist(values, bins=20, edgecolor="black", alpha=0.8)

    ax.set_title(f"Distribution of {column}", fontsize=14)
    ax.set_xlabel(column)
    ax.set_ylabel("Frequency")
    ax.grid(axis="y", alpha=0.25)

    filename = f"distribution_{_safe_filename(column)}.png"
    output_path = Path(output_dir) / filename

    return _save_chart(fig, output_path)


# ==========================================
# 2. CATEGORICAL ANALYSIS
# ==========================================

def create_categorical_bar_chart(
    df,
    column,
    output_dir="reports/charts",
    top_n=10
):
    """Create a bar chart showing the most frequent categories."""

    values = df[column].dropna().astype(str)

    if values.empty:
        return None

    counts = values.value_counts().head(top_n)

    fig, ax = plt.subplots(figsize=(8, 4.5))

    counts.sort_values().plot(
        kind="barh",
        ax=ax,
        edgecolor="black"
    )

    ax.set_title(f"Top Categories in {column}", fontsize=14)
    ax.set_xlabel("Number of Records")
    ax.set_ylabel(column)
    ax.grid(axis="x", alpha=0.25)

    filename = f"categories_{_safe_filename(column)}.png"
    output_path = Path(output_dir) / filename

    return _save_chart(fig, output_path)


# ==========================================
# 3. CORRELATION HEATMAP
# ==========================================

def create_correlation_heatmap(
    df,
    output_dir="reports/charts"
):
    """Create a correlation heatmap for numerical columns."""

    numeric_df = df.select_dtypes(include=np.number)

    if numeric_df.shape[1] < 2:
        return None

    correlation = numeric_df.corr()

    fig, ax = plt.subplots(figsize=(9, 7))

    image = ax.imshow(
        correlation,
        cmap="coolwarm",
        vmin=-1,
        vmax=1,
        aspect="auto"
    )

    ax.set_xticks(range(len(correlation.columns)))
    ax.set_yticks(range(len(correlation.columns)))

    ax.set_xticklabels(
        correlation.columns,
        rotation=45,
        ha="right"
    )
    ax.set_yticklabels(correlation.columns)

    ax.set_title("Numerical Features Correlation", fontsize=14)

    fig.colorbar(image, ax=ax, label="Correlation")

    output_path = Path(output_dir) / "correlation_heatmap.png"

    return _save_chart(fig, output_path)


# ==========================================
# 4. AUTOMATIC CHART GENERATION
# ==========================================

def generate_report_charts(
    df,
    output_dir="reports/charts",
    max_numeric=3,
    max_categorical=2
):
    """
    Automatically generate charts for a dataset.

    Returns a list containing chart titles,
    types, columns and saved image paths.
    """

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    charts = []

    numeric_columns = list(
        df.select_dtypes(include=np.number).columns
    )

    categorical_columns = list(
        df.select_dtypes(
            include=["object", "string", "category", "bool"]
        ).columns
    )

    # Numerical distribution charts
    for column in numeric_columns[:max_numeric]:

        path = create_numeric_distribution_chart(
            df,
            column,
            output_dir
        )

        if path:
            charts.append({
                "title": f"Distribution of {column}",
                "type": "histogram",
                "column": str(column),
                "path": path
            })

    # Categorical charts
    for column in categorical_columns[:max_categorical]:

        path = create_categorical_bar_chart(
            df,
            column,
            output_dir
        )

        if path:
            charts.append({
                "title": f"Top Categories in {column}",
                "type": "bar",
                "column": str(column),
                "path": path
            })

    # Correlation heatmap
    if len(numeric_columns) >= 2:

        path = create_correlation_heatmap(
            df,
            output_dir
        )

        if path:
            charts.append({
                "title": "Numerical Features Correlation",
                "type": "heatmap",
                "column": "multiple",
                "path": path
            })

    return charts
