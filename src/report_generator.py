
from pathlib import Path
from datetime import datetime

import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# ==========================================
# 1. DOCUMENT DESIGN
# ==========================================

def set_cell_background(cell, color):
    """Apply a background color to a table cell."""

    tc_pr = cell._tc.get_or_add_tcPr()

    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), color)

    tc_pr.append(shading)


def format_table(table, header=True):
    """Apply professional formatting to a Word table."""

    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for row_index, row in enumerate(table.rows):

        for cell in row.cells:

            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(9)

            if header and row_index == 0:
                set_cell_background(cell, "1F4E78")

                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(255, 255, 255)


def add_table(document, headers, rows):
    """Create a formatted Word table."""

    table = document.add_table(
        rows=1,
        cols=len(headers)
    )

    for index, header in enumerate(headers):
        table.rows[0].cells[index].text = str(header)

    for row_data in rows:

        cells = table.add_row().cells

        for index, value in enumerate(row_data):
            cells[index].text = str(value)

    format_table(table)

    document.add_paragraph()

    return table


def add_heading(document, text, level=1):
    """Add a styled section heading."""

    heading = document.add_heading(text, level=level)

    for run in heading.runs:
        run.font.name = "Arial"
        run.font.color.rgb = RGBColor(31, 78, 121)

    return heading


# ==========================================
# 2. EXECUTIVE SUMMARY
# ==========================================

def generate_executive_summary(
    df,
    quality_report,
    cleaning_report=None
):
    """Generate a concise summary from the dataset."""

    rows = len(df)
    columns = len(df.columns)

    numeric_columns = df.select_dtypes(
        include="number"
    ).columns

    categorical_columns = df.select_dtypes(
        include=["object", "string", "category"]
    ).columns

    summary = [
        f"The dataset contains {rows:,} records "
        f"and {columns} columns.",

        f"The dataset has {quality_report['missing_values']:,} "
        f"missing values "
        f"({quality_report['missing_percentage']}%).",

        f"It contains {quality_report['numeric_columns']} "
        f"numerical columns and "
        f"{quality_report['categorical_columns']} "
        f"categorical columns."
    ]

    if cleaning_report:

        summary.append(
            f"Data cleaning removed "
            f"{cleaning_report['duplicates_removed']} duplicate rows."
        )

    if len(numeric_columns) > 0:

        column = numeric_columns[0]

        mean_value = df[column].mean()

        if pd.notna(mean_value):

            summary.append(
                f"The average value of {column} is "
                f"{mean_value:,.2f}."
            )

    if len(categorical_columns) > 0:

        column = categorical_columns[0]

        if not df[column].dropna().empty:

            top_value = df[column].mode().iloc[0]

            summary.append(
                f"The most frequent value in {column} "
                f"is {top_value}."
            )

    return summary


# ==========================================
# 3. NUMERICAL STATISTICS
# ==========================================

def get_numerical_statistics(df):
    """Prepare numerical statistics for the report."""

    numeric_df = df.select_dtypes(include="number")

    if numeric_df.empty:
        return []

    statistics = numeric_df.describe().T

    rows = []

    for column, values in statistics.iterrows():

        rows.append([
            column,
            f"{values.get('count', 0):,.0f}",
            f"{values.get('mean', 0):,.2f}",
            f"{values.get('std', 0):,.2f}",
            f"{values.get('min', 0):,.2f}",
            f"{values.get('max', 0):,.2f}"
        ])

    return rows


# ==========================================
# 4. CATEGORICAL STATISTICS
# ==========================================

def get_categorical_statistics(df):
    """Prepare categorical analysis for the report."""

    categorical_df = df.select_dtypes(
        include=["object", "string", "category", "bool"]
    )

    rows = []

    for column in categorical_df.columns:

        values = df[column].dropna()

        if values.empty:
            most_frequent = "N/A"
        else:
            most_frequent = values.mode().iloc[0]

        rows.append([
            column,
            df[column].nunique(),
            df[column].isna().sum(),
            str(most_frequent)
        ])

    return rows


# ==========================================
# 5. DATA QUALITY DETAILS
# ==========================================

def get_missing_value_rows(missing_details):
    """Convert missing value details into table rows."""

    if missing_details is None or missing_details.empty:
        return []

    rows = []

    for _, row in missing_details.iterrows():

        rows.append([
            row["Column"],
            int(row["Missing Values"]),
            f"{row['Missing Percentage']:.2f}%"
        ])

    return rows


# ==========================================
# 6. WORD REPORT GENERATION
# ==========================================

def create_word_report(
    df,
    overview,
    cleaning_report=None,
    charts=None,
    output_path="reports/Excel_Analysis_Report.docx",
    report_title="Excel Data Analysis Report"
):
    """
    Generate a professional Word report.

    Parameters:
        df: cleaned DataFrame
        overview: output of generate_dataset_overview()
        cleaning_report: output of clean_dataframe()
        charts: output of generate_report_charts()
        output_path: destination of the Word report
    """

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    document = Document()

    # Page margins
    for section in document.sections:
        section.top_margin = Inches(0.65)
        section.bottom_margin = Inches(0.65)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    # Default font
    styles = document.styles

    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(10)

    # ======================================
    # COVER PAGE
    # ======================================

    document.add_paragraph("\n\n\n")

    title = document.add_heading(report_title, 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for run in title.runs:
        run.font.color.rgb = RGBColor(31, 78, 121)
        run.font.size = Pt(28)

    subtitle = document.add_paragraph(
        "Automated Data Analysis and Business Insights"
    )
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

    document.add_paragraph("\n")

    date_paragraph = document.add_paragraph(
        f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    )
    date_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    document.add_paragraph("\n\n")

    cover_info = document.add_paragraph(
        f"Dataset size: {len(df):,} rows × {len(df.columns)} columns"
    )
    cover_info.alignment = WD_ALIGN_PARAGRAPH.CENTER

    document.add_page_break()

    # ======================================
    # EXECUTIVE SUMMARY
    # ======================================

    add_heading(document, "1. Executive Summary")

    summary_items = generate_executive_summary(
        df,
        overview["quality"],
        cleaning_report
    )

    for item in summary_items:
        document.add_paragraph(item, style="List Bullet")

    # ======================================
    # KEY PERFORMANCE INDICATORS
    # ======================================

    add_heading(document, "2. Dataset KPIs")

    quality = overview["quality"]

    kpi_rows = [
        ["Total Rows", f"{quality['total_rows']:,}"],
        ["Total Columns", quality["total_columns"]],
        ["Missing Values", quality["missing_values"]],
        ["Missing Percentage", f"{quality['missing_percentage']}%"],
        ["Duplicate Rows", quality["duplicate_rows"]],
        ["Numerical Columns", quality["numeric_columns"]],
        ["Categorical Columns", quality["categorical_columns"]]
    ]

    if cleaning_report:

        kpi_rows.extend([
            ["Removed Duplicates", cleaning_report["duplicates_removed"]],
            ["Removed Empty Rows", cleaning_report["empty_rows_removed"]]
        ])

    add_table(
        document,
        ["Metric", "Value"],
        kpi_rows
    )

    # ======================================
    # DATA QUALITY
    # ======================================

    add_heading(document, "3. Data Quality Analysis")

    missing_rows = get_missing_value_rows(
        overview["missing_details"]
    )

    if missing_rows:

        document.add_paragraph(
            "The following columns contain missing values:"
        )

        add_table(
            document,
            ["Column", "Missing Values", "Missing Percentage"],
            missing_rows
        )

    else:

        document.add_paragraph(
            "No missing values were detected in the dataset."
        )

    # ======================================
    # NUMERICAL ANALYSIS
    # ======================================

    add_heading(document, "4. Numerical Analysis")

    numerical_rows = get_numerical_statistics(df)

    if numerical_rows:

        add_table(
            document,
            ["Column", "Count", "Mean", "Std", "Min", "Max"],
            numerical_rows
        )

    else:

        document.add_paragraph(
            "No numerical columns were found."
        )

    # ======================================
    # CATEGORICAL ANALYSIS
    # ======================================

    add_heading(document, "5. Categorical Analysis")

    categorical_rows = get_categorical_statistics(df)

    if categorical_rows:

        add_table(
            document,
            [
                "Column",
                "Unique Values",
                "Missing Values",
                "Most Frequent Value"
            ],
            categorical_rows
        )

    else:

        document.add_paragraph(
            "No categorical columns were found."
        )

    # ======================================
    # CHARTS AND VISUALIZATIONS
    # ======================================

    add_heading(document, "6. Data Visualizations")

    if charts:

        for chart in charts:

            chart_path = Path(chart["path"])

            if not chart_path.exists():
                continue

            document.add_heading(
                chart["title"],
                level=2
            )

            document.add_picture(
                str(chart_path),
                width=Inches(6.0)
            )

    else:

        document.add_paragraph(
            "No visualizations were generated for this dataset."
        )

    # ======================================
    # CONCLUSION
    # ======================================

    add_heading(document, "7. Conclusion")

    conclusion = (
        "This report provides an overview of the dataset, "
        "its quality, numerical characteristics, categorical "
        "distribution and generated visualizations. "
        "Further business analysis can be performed depending "
        "on the dataset's objectives."
    )

    document.add_paragraph(conclusion)

    # ======================================
    # FOOTER
    # ======================================

    for section in document.sections:

        footer = section.footer

        paragraph = footer.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        run = paragraph.add_run(
            "Generated automatically by Excel Report Generator"
        )

        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(120, 120, 120)

    # Save document
    document.save(output_path)

    return str(output_path)
