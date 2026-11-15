
from pathlib import Path
import hashlib
import io

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

from src.data_processor import (
    get_excel_sheets,
    load_excel_file,
    clean_dataframe,
    analyze_data_quality,
    analyze_numerical_columns,
    analyze_categorical_columns,
    generate_dataset_overview,
)

from src.chart_generator import generate_report_charts
from src.report_generator import create_word_report
from src.pdf_generator import create_pdf_report


# =========================================================
# PAGE CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="NEXA Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================================================
# THEME & CUSTOM CSS
# =========================================================

st.markdown("""
<style>

/* Global */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background-color: #F8FAFC;
    color: #0F172A;
}

.block-container {
    padding-top: 2rem;
    padding-bottom: 3rem;
    max-width: 1450px;
}

/* Hide Streamlit branding */
#MainMenu {
    visibility: hidden;
}

footer {
    visibility: hidden;
}

header[data-testid="stHeader"] {
    background: transparent;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #111827;
    border-right: 1px solid #263244;
}

section[data-testid="stSidebar"] > div {
    background: #111827;
}

section[data-testid="stSidebar"] * {
    color: #F9FAFB !important;
}

section[data-testid="stSidebar"] .stRadio label {
    background: transparent;
    border-radius: 9px;
    padding: 9px 12px;
}

section[data-testid="stSidebar"] .stRadio label:hover {
    background: #253247;
}

section[data-testid="stSidebar"] [data-testid="stFileUploader"] {
    background: #1F2937;
    border: 1px dashed #64748B;
    border-radius: 12px;
    padding: 12px;
}

/* Main headings */
h1, h2, h3, h4 {
    color: #0F172A !important;
    font-weight: 700 !important;
}

p, label, li {
    color: #334155;
}

.small-muted {
    color: #64748B;
    font-size: 13px;
}

/* KPI Cards */
.kpi-card {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 14px;
    padding: 20px;
    min-height: 125px;
    box-shadow: 0 2px 8px rgba(15, 23, 42, 0.035);
}

.kpi-label {
    color: #64748B;
    font-size: 12px;
    font-weight: 700;
    letter-spacing: 0.5px;
    margin-bottom: 12px;
}

.kpi-value {
    color: #0F172A;
    font-size: 25px;
    font-weight: 800;
    line-height: 1.3;
}

.kpi-caption {
    color: #64748B;
    font-size: 12px;
    margin-top: 7px;
}

/* Section cards */
.content-card {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 14px;
    padding: 20px;
    margin-bottom: 18px;
}

/* Buttons */
.stButton > button {
    background: #7C3AED;
    color: #FFFFFF !important;
    border: none;
    border-radius: 9px;
    padding: 9px 18px;
    font-weight: 600;
    transition: 0.2s;
}

.stButton > button:hover {
    background: #6D28D9;
    border: none;
    color: #FFFFFF !important;
}

/* Download buttons */
.stDownloadButton > button {
    background: #FFFFFF;
    color: #6D28D9 !important;
    border: 1px solid #C4B5FD;
    border-radius: 9px;
    font-weight: 600;
}

.stDownloadButton > button:hover {
    background: #F5F3FF;
    border-color: #7C3AED;
}

/* Inputs */
.stTextInput input,
.stSelectbox div[data-baseweb="select"],
.stNumberInput input,
.stTextArea textarea {
    background-color: #FFFFFF !important;
    color: #0F172A !important;
    border-radius: 8px;
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    gap: 12px;
    border-bottom: 1px solid #E2E8F0;
}

.stTabs [data-baseweb="tab"] {
    color: #475569 !important;
    font-weight: 600;
}

.stTabs [aria-selected="true"] {
    color: #7C3AED !important;
}

/* Dataframes */
[data-testid="stDataFrame"] {
    border: 1px solid #E2E8F0;
    border-radius: 10px;
    overflow: hidden;
}

/* Alerts */
div[data-testid="stAlert"] {
    border-radius: 10px;
}

/* Metric */
[data-testid="stMetric"] {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 15px;
}

[data-testid="stMetricLabel"] {
    color: #64748B !important;
}

[data-testid="stMetricValue"] {
    color: #0F172A !important;
}

/* Divider */
hr {
    border-color: #E2E8F0;
}

/* Responsive */
@media (max-width: 768px) {
    .block-container {
        padding: 1rem;
    }

    .kpi-value {
        font-size: 20px;
    }
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# CONSTANTS
# =========================================================

PRIMARY = "#7C3AED"
SECONDARY = "#0F172A"
MUTED = "#64748B"
PALETTE = [
    "#7C3AED",
    "#2563EB",
    "#0891B2",
    "#059669",
    "#D97706",
    "#DB2777",
    "#64748B",
]


# =========================================================
# SESSION STATE
# =========================================================

defaults = {
    "df_original": None,
    "df_clean": None,
    "cleaning_report": None,
    "selected_sheet": None,
    "file_signature": None,
    "word_bytes": None,
    "pdf_bytes": None,
    "report_signature": None,
    "report_title": "Business Data Analysis Report",
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def format_number(value):
    try:
        if pd.isna(value):
            return "N/A"
        return f"{value:,.2f}"
    except (TypeError, ValueError):
        return str(value)


def format_size(size):
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    return f"{size / (1024 * 1024):.2f} MB"


def section_header(title, description=None):
    st.markdown(f"## {title}")
    if description:
        st.markdown(
            f'<p class="small-muted">{description}</p>',
            unsafe_allow_html=True
        )


def kpi_card(label, value, caption=""):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-caption">{caption}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def style_chart(fig, height=380):
    fig.update_layout(
        height=height,
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(
            family="Inter, sans-serif",
            color="#334155",
            size=12
        ),
        margin=dict(l=20, r=20, t=45, b=30),
        title_font=dict(
            size=16,
            color="#0F172A"
        ),
        legend=dict(
            bgcolor="rgba(255,255,255,0)",
            font=dict(color="#475569")
        ),
        xaxis=dict(
            showgrid=False,
            zeroline=False,
            linecolor="#E2E8F0",
            tickfont=dict(color="#475569")
        ),
        yaxis=dict(
            showgrid=True,
            gridcolor="#E2E8F0",
            zeroline=False,
            tickfont=dict(color="#475569")
        ),
    )
    return fig


def dataframe_signature(uploaded_file, sheet_name):
    file_bytes = uploaded_file.getvalue()
    file_hash = hashlib.sha256(file_bytes).hexdigest()
    return f"{file_hash}_{sheet_name}"


def get_numeric_columns(df):
    return df.select_dtypes(include="number").columns.tolist()


def get_categorical_columns(df):
    return df.select_dtypes(
        include=["object", "category", "string"]
    ).columns.tolist()


def build_report_files(df, overview, cleaning_report, charts, title):
    """Create Word and PDF files and keep their bytes in session state."""

    reports_dir = Path("reports")
    reports_dir.mkdir(exist_ok=True)

    word_path = reports_dir / "Excel_Analysis_Report.docx"
    pdf_path = reports_dir / "Excel_Analysis_Report.pdf"

    create_word_report(
        df,
        overview,
        cleaning_report,
        charts,
        output_path=str(word_path),
        report_title=title,
    )

    create_pdf_report(
        df,
        overview,
        cleaning_report,
        charts,
        output_path=str(pdf_path),
        report_title=title,
    )

    st.session_state.word_bytes = word_path.read_bytes()
    st.session_state.pdf_bytes = pdf_path.read_bytes()


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown(
        """
        <div style="padding: 12px 0 24px 0;">
            <div style="font-size:28px;font-weight:800;color:#A78BFA;">
                ◈ NEXA
            </div>
            <div style="font-size:11px;color:#94A3B8;letter-spacing:2px;">
                ANALYTICS PLATFORM
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("### Navigation")

    page = st.radio(
        "Navigation",
        [
            "Overview",
            "Data Explorer",
            "Data Cleaning",
            "Analytics",
            "Reports",
        ],
        label_visibility="collapsed",
    )

    st.markdown("---")

    st.markdown("### Upload Dataset")

    uploaded_file = st.file_uploader(
        "Upload your Excel file",
        type=["xlsx"],
        help="Upload an Excel workbook (.xlsx)."
    )

    if uploaded_file is not None:

        st.caption(
            f"File: {uploaded_file.name}"
        )

        st.caption(
            f"Size: {format_size(uploaded_file.size)}"
        )

        try:
            sheets = get_excel_sheets(uploaded_file)

            selected_sheet = st.selectbox(
                "Select worksheet",
                sheets,
                key="worksheet_selector"
            )

            signature = dataframe_signature(
                uploaded_file,
                selected_sheet
            )

            if signature != st.session_state.file_signature:

                with st.spinner("Loading and processing dataset..."):

                    uploaded_file.seek(0)

                    original_df = load_excel_file(
                        uploaded_file,
                        sheet_name=selected_sheet
                    )

                    cleaned_df, cleaning_report = clean_dataframe(
                        original_df.copy()
                    )

                    st.session_state.df_original = original_df
                    st.session_state.df_clean = cleaned_df
                    st.session_state.cleaning_report = cleaning_report
                    st.session_state.selected_sheet = selected_sheet
                    st.session_state.file_signature = signature

                    st.session_state.word_bytes = None
                    st.session_state.pdf_bytes = None
                    st.session_state.report_signature = None

                st.success("Dataset loaded successfully!")

        except Exception as error:
            st.error(f"Unable to load Excel file: {error}")

    else:
        st.markdown(
            """
            <div style="
                background:#1F2937;
                border-radius:10px;
                padding:14px;
                color:#CBD5E1;
                font-size:12px;
            ">
                Upload an Excel workbook to start analyzing your data.
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("---")

    st.markdown(
        """
        <div style="font-size:11px;color:#94A3B8;">
            NEXA ANALYTICS v1.0<br>
            Built with Python & Streamlit
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# MAIN HEADER
# =========================================================

st.markdown(
    """
    <div style="margin-bottom:24px;">
        <div style="
            color:#7C3AED;
            font-size:12px;
            font-weight:700;
            letter-spacing:2px;
        ">
            DATA INTELLIGENCE
        </div>
        <div style="
            color:#0F172A;
            font-size:30px;
            font-weight:800;
            margin-top:5px;
        ">
            NEXA Analytics
        </div>
        <div style="
            color:#64748B;
            font-size:14px;
            margin-top:5px;
        ">
            Transform raw spreadsheets into meaningful business insights.
        </div>
    </div>
    """,
    unsafe_allow_html=True
)


# =========================================================
# CHECK DATASET
# =========================================================

if st.session_state.df_clean is None:

    st.markdown(
        """
        <div style="
            background:#FFFFFF;
            border:1px solid #E2E8F0;
            border-radius:16px;
            padding:45px 25px;
            text-align:center;
            margin-top:25px;
        ">
            <div style="font-size:48px;">📊</div>
            <div style="
                color:#0F172A;
                font-size:23px;
                font-weight:800;
                margin-top:15px;
            ">
                Welcome to NEXA Analytics
            </div>
            <div style="
                color:#64748B;
                font-size:14px;
                margin-top:10px;
            ">
                Upload an Excel workbook from the sidebar to begin.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.stop()


# =========================================================
# GLOBAL DATA
# =========================================================

original_df = st.session_state.df_original
df = st.session_state.df_clean
cleaning_report = st.session_state.cleaning_report

overview = generate_dataset_overview(df)

numeric_cols = get_numeric_columns(df)
categorical_cols = get_categorical_columns(df)


# =========================================================
# OVERVIEW PAGE
# =========================================================

if page == "Overview":

    section_header(
        "Business Overview",
        "A high-level view of your dataset and its key indicators."
    )

    total_rows = len(df)
    total_columns = len(df.columns)
    total_missing = int(df.isna().sum().sum())
    duplicate_count = int(df.duplicated().sum())

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        kpi_card(
            "TOTAL ROWS",
            f"{total_rows:,}",
            "Records after cleaning"
        )

    with col2:
        kpi_card(
            "TOTAL COLUMNS",
            f"{total_columns:,}",
            "Dataset variables"
        )

    with col3:
        kpi_card(
            "MISSING VALUES",
            f"{total_missing:,}",
            "Remaining empty cells"
        )

    with col4:
        kpi_card(
            "DUPLICATES",
            f"{duplicate_count:,}",
            "Duplicate records"
        )

    st.markdown("")

    left, right = st.columns([1.4, 1])

    with left:

        st.markdown('<div class="content-card">', unsafe_allow_html=True)

        st.markdown("### Dataset Preview")

        st.caption("First 10 rows of the cleaned dataset.")

        st.dataframe(
            df.head(10),
            use_container_width=True,
            hide_index=True
        )

        st.markdown("</div>", unsafe_allow_html=True)

    with right:

        st.markdown('<div class="content-card">', unsafe_allow_html=True)

        st.markdown("### Column Distribution")

        type_data = pd.DataFrame({
            "Type": ["Numerical", "Categorical", "Other"],
            "Count": [
                len(numeric_cols),
                len(categorical_cols),
                max(
                    0,
                    len(df.columns)
                    - len(numeric_cols)
                    - len(categorical_cols)
                )
            ]
        })

        fig = px.pie(
            type_data,
            names="Type",
            values="Count",
            hole=0.65,
            color_discrete_sequence=PALETTE,
        )

        fig.update_traces(
            textposition="inside",
            textinfo="label+value"
        )

        fig = style_chart(fig, 300)

        st.plotly_chart(
            fig,
            use_container_width=True
        )

        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("### Numerical Summary")

    if numeric_cols:

        st.dataframe(
            df[numeric_cols].describe().T,
            use_container_width=True
        )

    else:
        st.info("No numerical columns were detected.")

    st.markdown("### Dataset Information")

    info_col1, info_col2 = st.columns(2)

    with info_col1:
        st.write("**Original file:**", uploaded_file.name if uploaded_file else "Excel workbook")
        st.write("**Worksheet:**", st.session_state.selected_sheet)

    with info_col2:
        st.write("**Original rows:**", f"{len(original_df):,}")
        st.write("**Cleaned rows:**", f"{len(df):,}")


# =========================================================
# DATA EXPLORER PAGE
# =========================================================

elif page == "Data Explorer":

    section_header(
        "Data Explorer",
        "Inspect, search and export your original and processed data."
    )

    tab1, tab2, tab3 = st.tabs([
        "Cleaned Dataset",
        "Original Dataset",
        "Column Information"
    ])

    with tab1:

        st.markdown("### Cleaned Data")

        st.caption(
            f"{len(df):,} rows × {len(df.columns):,} columns"
        )

        st.dataframe(
            df,
            use_container_width=True,
            height=450
        )

        csv_data = df.to_csv(
            index=False
        ).encode("utf-8-sig")

        st.download_button(
            "Download Cleaned CSV",
            data=csv_data,
            file_name="cleaned_dataset.csv",
            mime="text/csv",
        )

    with tab2:

        st.markdown("### Original Data")

        st.caption(
            f"{len(original_df):,} rows × {len(original_df.columns):,} columns"
        )

        st.dataframe(
            original_df,
            use_container_width=True,
            height=450
        )

        original_csv = original_df.to_csv(
            index=False
        ).encode("utf-8-sig")

        st.download_button(
            "Download Original CSV",
            data=original_csv,
            file_name="original_dataset.csv",
            mime="text/csv",
        )

    with tab3:

        st.markdown("### Column Metadata")

        metadata = pd.DataFrame({
            "Column": df.columns,
            "Data Type": df.dtypes.astype(str).values,
            "Missing Values": df.isna().sum().values,
            "Missing (%)": (
                df.isna().mean().values * 100
            ).round(2),
            "Unique Values": df.nunique().values,
        })

        st.dataframe(
            metadata,
            use_container_width=True,
            hide_index=True
        )


# =========================================================
# DATA CLEANING PAGE
# =========================================================

elif page == "Data Cleaning":

    section_header(
        "Data Cleaning",
        "Understand the quality of your data and the cleaning operations."
    )

    original_rows = cleaning_report.get("original_rows", len(original_df))
    final_rows = cleaning_report.get("final_rows", len(df))
    removed_duplicates = cleaning_report.get("duplicates_removed", 0)
    removed_empty_rows = cleaning_report.get("empty_rows_removed", 0)
    removed_empty_columns = cleaning_report.get("empty_columns_removed", 0)

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        kpi_card("ORIGINAL ROWS", f"{original_rows:,}")

    with c2:
        kpi_card("FINAL ROWS", f"{final_rows:,}")

    with c3:
        kpi_card("DUPLICATES REMOVED", f"{removed_duplicates:,}")

    with c4:
        kpi_card(
            "EMPTY ROWS REMOVED",
            f"{removed_empty_rows:,}"
        )

    st.markdown("")

    left, right = st.columns(2)

    with left:

        st.markdown('<div class="content-card">', unsafe_allow_html=True)

        st.markdown("### Missing Values by Column")

        missing = df.isna().sum()

        missing = missing[missing > 0].sort_values(
            ascending=False
        )

        if not missing.empty:

            missing_df = missing.reset_index()
            missing_df.columns = ["Column", "Missing Values"]

            fig = px.bar(
                missing_df,
                x="Column",
                y="Missing Values",
                color_discrete_sequence=[PRIMARY],
            )

            fig = style_chart(fig)

            st.plotly_chart(
                fig,
                use_container_width=True
            )

        else:
            st.success("No missing values found in the cleaned dataset.")

        st.markdown("</div>", unsafe_allow_html=True)

    with right:

        st.markdown('<div class="content-card">', unsafe_allow_html=True)

        st.markdown("### Cleaning Operations")

        operations = [
            ("Empty rows removed", removed_empty_rows),
            ("Empty columns removed", removed_empty_columns),
            ("Duplicate rows removed", removed_duplicates),
        ]

        for label, value in operations:

            st.markdown(
                f"""
                <div style="
                    display:flex;
                    justify-content:space-between;
                    padding:12px 0;
                    border-bottom:1px solid #E2E8F0;
                ">
                    <span style="color:#475569;">{label}</span>
                    <strong style="color:#0F172A;">{value:,}</strong>
                </div>
                """,
                unsafe_allow_html=True
            )

        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("### Data Quality Summary")

    quality = analyze_data_quality(df)

    if isinstance(quality, pd.DataFrame):

        st.dataframe(
            quality,
            use_container_width=True
        )

    elif isinstance(quality, dict):

        quality_df = pd.DataFrame(
            list(quality.items()),
            columns=["Metric", "Value"]
        )

        st.dataframe(
            quality_df,
            use_container_width=True,
            hide_index=True
        )

    else:

        st.write(quality)

    st.info(
        "Note: Missing values are reported, not automatically imputed. "
        "This avoids introducing assumptions into the original data."
    )


# =========================================================
# ANALYTICS PAGE
# =========================================================

elif page == "Analytics":

    section_header(
        "Data Analytics",
        "Explore distributions, category patterns and numerical relationships."
    )

    if numeric_cols:

        st.markdown("### Numerical Analysis")

        selected_numeric = st.selectbox(
            "Select a numerical column",
            numeric_cols
        )

        series = df[selected_numeric].dropna()

        a, b, c = st.columns(3)

        with a:
            kpi_card(
                "AVERAGE",
                format_number(series.mean())
            )

        with b:
            kpi_card(
                "MEDIAN",
                format_number(series.median())
            )

        with c:
            kpi_card(
                "STANDARD DEVIATION",
                format_number(series.std())
            )

        left, right = st.columns(2)

        with left:

            st.markdown('<div class="content-card">', unsafe_allow_html=True)

            st.markdown("### Distribution")

            fig = px.histogram(
                df,
                x=selected_numeric,
                nbins=30,
                color_discrete_sequence=[PRIMARY],
            )

            fig = style_chart(fig)

            st.plotly_chart(
                fig,
                use_container_width=True
            )

            st.markdown("</div>", unsafe_allow_html=True)

        with right:

            st.markdown('<div class="content-card">', unsafe_allow_html=True)

            st.markdown("### Box Plot")

            fig = px.box(
                df,
                y=selected_numeric,
                color_discrete_sequence=["#2563EB"],
            )

            fig = style_chart(fig)

            st.plotly_chart(
                fig,
                use_container_width=True
            )

            st.markdown("</div>", unsafe_allow_html=True)

    else:

        st.info("No numerical columns available for numerical analysis.")

    st.markdown("---")

    if categorical_cols:

        st.markdown("### Categorical Analysis")

        selected_category = st.selectbox(
            "Select a categorical column",
            categorical_cols
        )

        category_counts = (
            df[selected_category]
            .fillna("Missing")
            .astype(str)
            .value_counts()
            .head(15)
            .reset_index()
        )

        category_counts.columns = [
            "Category",
            "Count"
        ]

        fig = px.bar(
            category_counts,
            x="Category",
            y="Count",
            color="Count",
            color_continuous_scale=["#DDD6FE", "#7C3AED"],
            text="Count",
        )

        fig.update_traces(
            textposition="outside"
        )

        fig = style_chart(fig)

        st.plotly_chart(
            fig,
            use_container_width=True
        )

        st.markdown("### Category Frequency")

        st.dataframe(
            category_counts,
            use_container_width=True,
            hide_index=True
        )

    else:

        st.info("No categorical columns available.")

    st.markdown("---")

    if len(numeric_cols) >= 2:

        st.markdown("### Correlation Analysis")

        correlation = df[numeric_cols].corr()

        fig = px.imshow(
            correlation,
            text_auto=".2f",
            color_continuous_scale="RdBu_r",
            zmin=-1,
            zmax=1,
            aspect="auto",
        )

        fig.update_layout(
            height=550,
            paper_bgcolor="#FFFFFF",
            font=dict(
                family="Inter",
                color="#334155"
            ),
            margin=dict(l=20, r=20, t=40, b=20),
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )

    else:

        st.info(
            "At least two numerical columns are required for correlation analysis."
        )


# =========================================================
# REPORTS PAGE
# =========================================================

elif page == "Reports":

    section_header(
        "Professional Reports",
        "Generate downloadable Word and PDF reports from your Excel data."
    )

    st.markdown('<div class="content-card">', unsafe_allow_html=True)

    st.markdown("### Report Configuration")

    report_title = st.text_input(
        "Report Title",
        value=st.session_state.report_title,
        max_chars=100
    )

    st.session_state.report_title = report_title

    st.markdown(
        """
        <div style="color:#64748B;font-size:13px;margin-top:5px;">
            Your reports include dataset overview, cleaning results,
            statistical summaries, charts and conclusions.
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("</div>", unsafe_allow_html=True)

    report_signature = (
        f"{st.session_state.file_signature}_{report_title}"
    )

    if st.session_state.report_signature != report_signature:

        st.session_state.word_bytes = None
        st.session_state.pdf_bytes = None

    st.markdown("### Generate Your Reports")

    col1, col2 = st.columns(2)

    with col1:

        st.markdown('<div class="content-card">', unsafe_allow_html=True)

        st.markdown("### Word Document")

        st.markdown(
            '<p class="small-muted">Editable professional report in DOCX format.</p>',
            unsafe_allow_html=True
        )

        if st.button(
            "Generate Word Report",
            use_container_width=True
        ):

            try:

                with st.spinner("Generating Word report..."):

                    charts = generate_report_charts(
                        df,
                        output_dir="reports/charts"
                    )

                    build_report_files(
                        df,
                        overview,
                        cleaning_report,
                        charts,
                        report_title
                    )

                    st.session_state.report_signature = report_signature

                st.success("Word report generated successfully!")

            except Exception as error:

                st.error(f"Word report error: {error}")

        if st.session_state.word_bytes:

            st.download_button(
                "Download Word Report",
                data=st.session_state.word_bytes,
                file_name="NEXA_Analysis_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

        st.markdown("</div>", unsafe_allow_html=True)

    with col2:

        st.markdown('<div class="content-card">', unsafe_allow_html=True)

        st.markdown("### PDF Document")

        st.markdown(
            '<p class="small-muted">Ready-to-share report in PDF format.</p>',
            unsafe_allow_html=True
        )

        if st.button(
            "Generate PDF Report",
            use_container_width=True
        ):

            try:

                with st.spinner("Generating PDF report..."):

                    charts = generate_report_charts(
                        df,
                        output_dir="reports/charts"
                    )

                    build_report_files(
                        df,
                        overview,
                        cleaning_report,
                        charts,
                        report_title
                    )

                    st.session_state.report_signature = report_signature

                st.success("PDF report generated successfully!")

            except Exception as error:

                st.error(f"PDF report error: {error}")

        if st.session_state.pdf_bytes:

            st.download_button(
                "Download PDF Report",
                data=st.session_state.pdf_bytes,
                file_name="NEXA_Analysis_Report.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("---")

    st.markdown("### Report Contents")

    features = [
        "Executive Summary",
        "Dataset Overview and KPIs",
        "Data Quality Analysis",
        "Numerical Statistical Summary",
        "Categorical Statistical Summary",
        "Charts and Visualizations",
        "Business Insights and Conclusion",
    ]

    for feature in features:
        st.markdown(f"- ✓ {feature}")


# =========================================================
# FOOTER
# =========================================================

st.markdown("---")

st.markdown(
    """
    <div style="
        text-align:center;
        color:#94A3B8;
        font-size:12px;
        padding:10px 0;
    ">
        NEXA Analytics · Excel to Professional Reports · Built with Python
    </div>
    """,
    unsafe_allow_html=True
)
