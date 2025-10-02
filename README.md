# Marketing KPI Dashboard Generator

A Streamlit web application that generates professional PowerPoint dashboards from marketing KPI data.

## Features

- 📁 **Excel Upload**: Upload your existing KPI data in Excel format
- ✏️ **Manual Entry**: Add KPIs manually through an intuitive form
- 🎯 **Auto Status Calculation**: Automatically calculates Green/Yellow/Red status based on performance
- 📊 **PowerPoint Generation**: Creates professional presentations with Klaviyo-inspired styling
- 🏢 **Customizable**: Add your company name to the presentation

## Installation

1. Clone or download the files
2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the application:
```bash
streamlit run marketing_kpi_app.py
```

2. Open your browser to `http://localhost:8501`

3. Choose your input method:
   - **Upload Excel**: Use the format from your sample file
   - **Manual Entry**: Add KPIs one by one using the form

4. Set your company name in the sidebar

5. Generate and download your PowerPoint presentation!

## Excel Format

Your Excel file should have these columns:
- Campaign Type
- KPI Name  
- Benchmark
- Actual
- Direction (HigherIsBetter or LowerIsBetter)

## Status Logic

- **Green**: Performance meets or exceeds benchmark
- **Yellow**: Performance is close to benchmark (90-99% for HigherIsBetter, 1-10% above for LowerIsBetter)
- **Red**: Performance significantly below benchmark

## Requirements

- Python 3.7+
- Streamlit
- pandas
- python-pptx
- openpyxl
