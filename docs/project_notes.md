# Project Notes

## Pipeline Stages

1. `src.generate_sample_data` creates demo monthly Excel files in `data/raw/`.
2. `src.load_data` reads all raw `.xlsx` files and combines them into one DataFrame.
3. `src.clean_data` standardizes column names, trims and formats text fields, converts dates and numeric values, and calculates revenue.
4. `src.validate_data` creates validation summary and issue detail tables.
5. `src.analyze_data` calculates KPIs and sales summary tables.
6. `src.export_report` builds the final formatted Excel workbook.
7. `src.main` runs the full workflow in order.

## Testing Scope

The tests focus on core business logic with small in-memory pandas DataFrames:

- column name normalization;
- text cleaning;
- date and numeric conversion;
- revenue calculation;
- validation issue detection;
- basic sales KPI calculation.

Excel file input and output are covered lightly by the runnable pipeline rather than heavily mocked unit tests.

## Current Limitations

- The project assumes the raw input files use the expected sales columns.
- The final Excel report is static and does not include interactive dashboard behavior.
- The sample data is simulated and intended for portfolio demonstration, not real business analysis.
