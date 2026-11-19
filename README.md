# NEXA Analytics — Excel to Professional Reports

Transform Excel spreadsheets into meaningful business insights and professional reports with an interactive Python application.

NEXA Analytics helps users explore datasets, assess data quality, clean records, visualize information, and generate downloadable Word and PDF reports — all through a modern Streamlit interface.

## Overview

NEXA Analytics is a data analysis and reporting application designed to simplify spreadsheet analysis for individuals, analysts, and businesses.

Upload an Excel workbook, select a worksheet, explore your data, and generate professional reports without manually preparing everything in Excel.

## Features

### 1. Excel Data Import
- Upload `.xlsx` workbooks.
- Select worksheets from multi-sheet Excel files.
- Preview original and processed datasets.

### 2. Data Cleaning
- Remove duplicate records.
- Remove completely empty rows and columns.
- Trim unnecessary spaces from text values.
- Detect missing values.
- Generate a data cleaning summary.

### 3. Data Exploration
- Interactive dataset preview.
- Column data types.
- Missing-value counts and percentages.
- Unique-value analysis.
- Numerical statistical summaries.
- Download processed data as CSV.

### 4. Data Analytics
- Numerical distribution analysis.
- Histograms and box plots.
- Categorical frequency analysis.
- Correlation heatmaps.
- Interactive Plotly visualizations.

### 5. Professional Report Generation
Generate downloadable reports in two formats:

**Microsoft Word (.docx)**
- Report title and introduction.
- Executive summary.
- Dataset overview.
- Data quality results.
- Statistical analysis.
- Charts and visualizations.
- Conclusion.

**PDF (.pdf)**
- Structured report layout.
- Dataset statistics.
- Data quality analysis.
- Embedded charts.
- Professional formatting.

## Technology Stack

| Technology | Purpose |
|---|---|
| Python | Core application |
| Pandas | Data processing and analysis |
| NumPy | Numerical operations |
| Streamlit | Interactive web interface |
| Plotly | Interactive visualizations |
| Matplotlib | Report chart generation |
| OpenPyXL | Excel workbook processing |
| python-docx | Word report generation |
| ReportLab | PDF report generation |

## Project Structure

```text
Excel-to-Professional-Report-Generator/
│
├── .streamlit/
│   └── config.toml
│
├── assets/
│   └── overview.png
│
├── sample_data/
│   ├── generate_sample.py
│   └── business_sales.xlsx
│
├── src/
│   ├── __init__.py
│   ├── data_processor.py
│   ├── chart_generator.py
│   ├── report_generator.py
│   └── pdf_generator.py
│
├── reports/
│
├── app.py
├── requirements.txt
├── README.md
└── .gitignore
```

## Installation

### Prerequisites
- Python 3.10 or later.
- pip.
- Git (optional, for cloning the repository).

### 1. Clone the repository

```bash
git clone https://github.com/Hakouben26/Excel-to-Professional-Report-Generator.git
cd Excel-to-Professional-Report-Generator
```

### 2. Create a virtual environment

Windows:

```powershell
python -m venv .venv
.venv\Scripts\activate
```

Linux / macOS:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Run the application

```bash
streamlit run app.py
```

The application will open in your browser.

## Usage

1. Launch NEXA Analytics.
2. Upload an Excel `.xlsx` workbook.
3. Select the worksheet you want to analyze.
4. Explore the dataset using the Overview and Data Explorer pages.
5. Review detected data-quality issues in Data Cleaning.
6. Explore numerical and categorical patterns in Analytics.
7. Open Reports and enter a report title.
8. Generate and download your Word or PDF report.

## Sample Dataset

A sample business sales workbook is included in `sample_data/` to help test the application.

It contains multiple worksheets and example records for exploring the application's data processing and reporting capabilities.

To regenerate the sample workbook:

```bash
python sample_data/generate_sample.py
```

## Data Cleaning Policy

NEXA Analytics performs conservative data cleaning:

- Duplicate records are removed.
- Completely empty rows and columns are removed.
- Leading and trailing spaces in text are trimmed.
- Missing values are identified and reported.

**Important:** Missing values are not automatically imputed. This avoids introducing assumptions that could change the meaning of the original dataset.

## Limitations

- Current Excel import supports `.xlsx` files.
- Missing-value imputation is not currently automated.
- Statistical results depend on the quality and structure of the uploaded dataset.
- Generated reports are based on the selected worksheet.

## Future Improvements

- Support for CSV and additional Excel formats.
- Custom missing-value treatment options.
- More advanced business KPIs.
- Multilingual reports (English, French, and Arabic).
- More customizable report templates.
- Additional export formats.

## Author

Developed as a portfolio project demonstrating practical skills in:

- Data Analysis
- Python Development
- Data Cleaning
- Data Visualization
- Automated Reporting
- Streamlit Application Development

## License

This project is available under the MIT License.
