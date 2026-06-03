# Excel Report Automation with Python

This portfolio project automates a small business sales reporting workflow with Python and Excel. It loads monthly Excel files, cleans and validates the sales data, calculates sales summaries, and exports a formatted Excel report.

The project is designed to show practical beginner-to-intermediate data automation skills that could be useful for a freelancer, operations assistant, analyst, or small business owner.

## Business Problem

Many small businesses track sales in separate monthly Excel files. This can make reporting slow and error-prone because someone has to manually combine files, clean inconsistent values, check data quality, calculate metrics, and format a report.

This project demonstrates how that workflow can be automated with Python while still producing Excel outputs that non-technical users can review.

## Solution Overview

The pipeline:

1. Reads monthly sales workbooks from `data/raw/`.
2. Combines the files into one dataset.
3. Cleans column names, text fields, dates, and numeric values.
4. Calculates revenue for each order line.
5. Runs data quality checks and creates a validation report.
6. Builds sales summary tables and key performance indicators.
7. Exports a final formatted Excel report.

## Key Features

- Generates sample monthly Excel sales files for demo use.
- Loads and combines multiple monthly Excel workbooks.
- Standardizes column names into `snake_case`.
- Trims text fields and standardizes text casing where appropriate.
- Converts order dates and numeric sales fields into usable data types.
- Calculates revenue from quantity, unit price, and discount.
- Creates a data quality report with summary and issue detail sheets.
- Calculates basic sales KPIs and summary tables.
- Exports a formatted final Excel workbook.
- Includes pytest tests for core cleaning, validation, and KPI logic.

## Project Structure

```text
excel-report-automation-python/
├── data/
│   ├── raw/                 # Monthly source Excel files
│   ├── processed/           # Cleaned dataset output
│   └── output/              # Reports and summary workbooks
├── docs/
│   └── project_notes.md     # Brief technical notes
├── src/
│   ├── analyze_data.py      # KPI and summary table calculations
│   ├── clean_data.py        # Data cleaning and revenue calculation
│   ├── export_report.py     # Final formatted Excel report export
│   ├── generate_sample_data.py
│   ├── load_data.py         # Load and combine monthly Excel files
│   ├── main.py              # Full pipeline entry point
│   └── validate_data.py     # Data quality checks
├── tests/
│   ├── test_analyze_data.py
│   ├── test_clean_data.py
│   └── test_validate_data.py
├── requirements.txt
└── README.md
```

## Technologies Used

- Python
- pandas
- openpyxl
- pytest
- Excel workbooks (`.xlsx`)

## Setup Instructions

Clone the repository, then create and activate a virtual environment.

```bash
python -m venv .venv
```

On Windows:

```bash
.venv\Scripts\activate
```

On macOS or Linux:

```bash
source .venv/bin/activate
```

Install the dependencies:

```bash
pip install -r requirements.txt
```

## Generate Sample Data

The repository includes sample raw Excel files, but you can regenerate them with:

```bash
python -m src.generate_sample_data
```

This creates 12 monthly sales files in `data/raw/` for the year 2025. The sample data includes a few intentional data quality issues so the validation report has realistic examples to catch.

## Run the Full Pipeline

Run:

```bash
python -m src.main
```

The pipeline loads the raw monthly files, cleans the data, validates it, creates sales summaries, and exports the final report.

## Run Tests

Run:

```bash
pytest
```

The tests use small in-memory pandas DataFrames where possible. They focus on the most important data pipeline logic rather than testing Excel file input and output heavily.

## Output Files Generated

After running the full pipeline, these files are created or updated:

- `data/processed/clean_sales_data.xlsx`
- `data/output/data_quality_report.xlsx`
- `data/output/sales_summary.xlsx`
- `data/output/final_sales_report.xlsx`

## Data Quality Checks Included

The validation step checks for:

- Missing required columns.
- Missing values in key fields.
- Duplicate `order_id` values.
- Invalid or missing `order_date` values.
- Invalid `quantity` values.
- Invalid `unit_price` values.
- Invalid `discount` values outside the expected 0 to 1 range.
- Missing required text fields.
- Missing, non-positive, or incorrectly calculated `revenue` values.

## Example Use Case

A small retail business or freelance client may receive one sales spreadsheet per month from their point-of-sale system. Instead of manually combining those files and building a report every time, this project shows how Python can automate the repeatable work:

- combine the monthly files;
- clean inconsistent names, dates, and numbers;
- flag data issues before reporting;
- calculate sales performance metrics;
- produce an Excel workbook that can be sent to the owner or manager.

## Possible Future Improvements

- Add command-line arguments for input and output folders.
- Add more tests around Excel export formatting.
- Add configuration for custom required columns.
- Support multiple years of source files.
- Add richer logging for production use.
- Add optional charts inside the Excel report.

