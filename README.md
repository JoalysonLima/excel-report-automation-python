# Excel Report Automation with Python

This portfolio project automates an Excel reporting workflow with Python.

## Clean the monthly sales data

Install the dependencies, then run:

```bash
python src/clean_data.py
```

The cleaned dataset is saved to `data/processed/clean_sales_data.xlsx`.

## Analyze the cleaned sales data

Run:

```bash
python src/analyze_data.py
```

The sales summary workbook is saved to `data/output/sales_summary.xlsx`.

## Run the full pipeline

Run:

```bash
python -m src.main
```

This creates the cleaned dataset, data quality report, sales summary workbook, and final formatted Excel report.
