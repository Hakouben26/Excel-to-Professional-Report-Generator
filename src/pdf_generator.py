
from pathlib import Path
from datetime import datetime

import pandas as pd

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    Image
)


# ==========================================
# 1. PDF DESIGN
# ==========================================

def create_pdf_styles():

    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle(
        name="ReportTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=25,
        leading=32,
        textColor=colors.HexColor("#1F4E78"),
        alignment=TA_CENTER,
        spaceAfter=20
    ))

    styles.add(ParagraphStyle(
        name="SectionHeading",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=15,
        leading=20,
        textColor=colors.HexColor("#1F4E78"),
        spaceBefore=14,
        spaceAfter=10
    ))

    styles.add(ParagraphStyle(
        name="BodyTextCustom",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=9,
        leading=14,
        spaceAfter=7
    ))

    styles.add(ParagraphStyle(
        name="SmallText",
        parent=styles["BodyText"],
        fontSize=8,
        leading=11
    ))

    return styles


# ==========================================
# 2. PDF TABLE GENERATOR
# ==========================================

def create_pdf_table(headers, rows):

    data = [headers] + rows

    table = Table(
        data,
        repeatRows=1,
        hAlign="LEFT"
    )

    table.setStyle(TableStyle([

        ("BACKGROUND", (0, 0), (-1, 0),
         colors.HexColor("#1F4E78")),

        ("TEXTCOLOR", (0, 0), (-1, 0),
         colors.white),

        ("FONTNAME", (0, 0), (-1, 0),
         "Helvetica-Bold"),

        ("FONTNAME", (0, 1), (-1, -1),
         "Helvetica"),

        ("FONTSIZE", (0, 0), (-1, -1), 8),

        ("GRID", (0, 0), (-1, -1),
         0.4, colors.HexColor("#D9E2F3")),

        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.white, colors.HexColor("#F2F6FA")]),

        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),

        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),

        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),

    ]))

    return table


# ==========================================
# 3. PDF HEADER AND FOOTER
# ==========================================

def add_page_number(canvas, document):

    canvas.saveState()

    page_width, page_height = A4

    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.grey)

    canvas.drawString(
        2 * cm,
        1 * cm,
        "Excel Professional Report Generator"
    )

    canvas.drawRightString(
        page_width - 2 * cm,
        1 * cm,
        f"Page {document.page}"
    )

    canvas.restoreState()


# ==========================================
# 4. EXECUTIVE SUMMARY
# ==========================================

def create_pdf_executive_summary(
    df,
    quality_report,
    cleaning_report=None
):

    summary = []

    summary.append(
        f"The dataset contains {len(df):,} records "
        f"and {len(df.columns)} columns."
    )

    summary.append(
        f"Detected missing values: "
        f"{quality_report['missing_values']:,} "
        f"({quality_report['missing_percentage']}%)."
    )

    summary.append(
        f"Numerical columns: "
        f"{quality_report['numeric_columns']}."
    )

    summary.append(
        f"Categorical columns: "
        f"{quality_report['categorical_columns']}."
    )

    if cleaning_report:

        summary.append(
            f"Duplicate rows removed during cleaning: "
            f"{cleaning_report['duplicates_removed']}."
        )

    numeric_columns = df.select_dtypes(
        include="number"
    ).columns

    if len(numeric_columns) > 0:

        column = numeric_columns[0]
        mean_value = df[column].mean()

        if pd.notna(mean_value):

            summary.append(
                f"Average {column}: {mean_value:,.2f}."
            )

    return summary


# ==========================================
# 5. PDF REPORT GENERATION
# ==========================================

def create_pdf_report(
    df,
    overview,
    cleaning_report=None,
    charts=None,
    output_path="reports/Excel_Analysis_Report.pdf",
    report_title="Excel Data Analysis Report"
):

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    styles = create_pdf_styles()

    document = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=2 * cm,
        leftMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm
    )

    story = []

    quality = overview["quality"]

    # ======================================
    # COVER PAGE
    # ======================================

    story.append(Spacer(1, 5 * cm))

    story.append(
        Paragraph(report_title, styles["ReportTitle"])
    )

    story.append(
        Paragraph(
            "Automated Data Analysis and Business Insights",
            styles["BodyTextCustom"]
        )
    )

    story.append(Spacer(1, 1.5 * cm))

    story.append(
        Paragraph(
            f"<b>Generated on:</b> "
            f"{datetime.now().strftime('%Y-%m-%d %H:%M')}",
            styles["BodyTextCustom"]
        )
    )

    story.append(
        Paragraph(
            f"<b>Dataset size:</b> "
            f"{len(df):,} rows × {len(df.columns)} columns",
            styles["BodyTextCustom"]
        )
    )

    story.append(PageBreak())

    # ======================================
    # EXECUTIVE SUMMARY
    # ======================================

    story.append(
        Paragraph("1. Executive Summary", styles["SectionHeading"])
    )

    summary = create_pdf_executive_summary(
        df,
        quality,
        cleaning_report
    )

    for item in summary:

        story.append(
            Paragraph(
                f"&bull; {item}",
                styles["BodyTextCustom"]
            )
        )

    story.append(Spacer(1, 0.3 * cm))

    # ======================================
    # DATASET KPIs
    # ======================================

    story.append(
        Paragraph("2. Dataset KPIs", styles["SectionHeading"])
    )

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
            ["Removed Duplicates",
             cleaning_report["duplicates_removed"]],

            ["Removed Empty Rows",
             cleaning_report["empty_rows_removed"]]
        ])

    story.append(
        create_pdf_table(
            ["Metric", "Value"],
            kpi_rows
        )
    )

    story.append(Spacer(1, 0.5 * cm))

    # ======================================
    # DATA QUALITY
    # ======================================

    story.append(
        Paragraph("3. Data Quality Analysis", styles["SectionHeading"])
    )

    missing_details = overview["missing_details"]

    if missing_details is not None and not missing_details.empty:

        missing_rows = []

        for _, row in missing_details.iterrows():

            missing_rows.append([
                str(row["Column"]),
                int(row["Missing Values"]),
                f"{row['Missing Percentage']:.2f}%"
            ])

        story.append(
            create_pdf_table(
                ["Column", "Missing Values", "Missing Percentage"],
                missing_rows
            )
        )

    else:

        story.append(
            Paragraph(
                "No missing values were detected.",
                styles["BodyTextCustom"]
            )
        )

    story.append(Spacer(1, 0.5 * cm))

    # ======================================
    # NUMERICAL ANALYSIS
    # ======================================

    story.append(
        Paragraph("4. Numerical Analysis", styles["SectionHeading"])
    )

    numeric_df = df.select_dtypes(include="number")

    if not numeric_df.empty:

        statistics = numeric_df.describe().T

        numerical_rows = []

        for column, values in statistics.iterrows():

            numerical_rows.append([
                str(column),
                f"{values['count']:,.0f}",
                f"{values['mean']:,.2f}",
                f"{values['std']:,.2f}",
                f"{values['min']:,.2f}",
                f"{values['max']:,.2f}"
            ])

        story.append(
            create_pdf_table(
                ["Column", "Count", "Mean", "Std", "Min", "Max"],
                numerical_rows
            )
        )

    else:

        story.append(
            Paragraph(
                "No numerical columns were found.",
                styles["BodyTextCustom"]
            )
        )

    story.append(Spacer(1, 0.5 * cm))

    # ======================================
    # CATEGORICAL ANALYSIS
    # ======================================

    story.append(
        Paragraph("5. Categorical Analysis", styles["SectionHeading"])
    )

    categorical_df = df.select_dtypes(
        include=["object", "string", "category", "bool"]
    )

    if not categorical_df.empty:

        categorical_rows = []

        for column in categorical_df.columns:

            values = df[column].dropna()

            most_frequent = (
                str(values.mode().iloc[0])
                if not values.empty
                else "N/A"
            )

            categorical_rows.append([
                str(column),
                int(df[column].nunique()),
                int(df[column].isna().sum()),
                most_frequent
            ])

        story.append(
            create_pdf_table(
                [
                    "Column",
                    "Unique Values",
                    "Missing Values",
                    "Most Frequent"
                ],
                categorical_rows
            )
        )

    else:

        story.append(
            Paragraph(
                "No categorical columns were found.",
                styles["BodyTextCustom"]
            )
        )

    story.append(Spacer(1, 0.5 * cm))

    # ======================================
    # VISUALIZATIONS
    # ======================================

    story.append(
        Paragraph("6. Data Visualizations", styles["SectionHeading"])
    )

    if charts:

        for chart in charts:

            chart_path = Path(chart["path"])

            if not chart_path.exists():
                continue

            story.append(
                Paragraph(
                    str(chart["title"]),
                    styles["Heading3"]
                )
            )

            image = Image(str(chart_path))

            max_width = 16 * cm
            max_height = 9 * cm

            scale = min(
                max_width / image.imageWidth,
                max_height / image.imageHeight,
                1
            )

            image.drawWidth = image.imageWidth * scale
            image.drawHeight = image.imageHeight * scale

            story.append(image)
            story.append(Spacer(1, 0.4 * cm))

    else:

        story.append(
            Paragraph(
                "No visualizations were generated.",
                styles["BodyTextCustom"]
            )
        )

    # ======================================
    # CONCLUSION
    # ======================================

    story.append(
        Paragraph("7. Conclusion", styles["SectionHeading"])
    )

    story.append(
        Paragraph(
            "This report presents the dataset overview, data quality, "
            "statistical analysis and visualizations. "
            "The findings can support further business analysis "
            "and informed decision-making.",
            styles["BodyTextCustom"]
        )
    )

    # Generate PDF
    document.build(
        story,
        onFirstPage=add_page_number,
        onLaterPages=add_page_number
    )

    return str(output_path)
