"""Run the full Excel report automation pipeline."""

from src.analyze_data import (
    OUTPUT_FILE as SALES_SUMMARY_FILE,
    build_summary_tables,
    calculate_kpis,
    print_kpis,
    save_summary_tables,
)
from src.clean_data import OUTPUT_FILE as CLEAN_DATA_FILE
from src.clean_data import clean_sales_data, save_cleaned_data
from src.export_report import OUTPUT_FILE as FINAL_REPORT_FILE
from src.export_report import build_final_report
from src.load_data import load_monthly_files
from src.validate_data import OUTPUT_FILE as DATA_QUALITY_FILE
from src.validate_data import (
    print_validation_summary,
    run_validations,
    save_quality_report,
)


def main() -> None:
    """Run the complete sales reporting workflow from raw files to final report."""
    print("Starting Excel report automation pipeline...")
    print("")

    print("Step 1: Loading raw monthly Excel files...")
    raw_data = load_monthly_files()
    if raw_data.empty:
        print("No monthly sales data found. Pipeline stopped.")
        return
    print(f"Loaded {len(raw_data)} rows from raw monthly files.")
    print("")

    print("Step 2: Cleaning and standardizing sales data...")
    clean_data = clean_sales_data(raw_data)
    save_cleaned_data(clean_data)
    print(f"Saved cleaned data to: {CLEAN_DATA_FILE}")
    print("")

    print("Step 3: Validating cleaned data...")
    validation_summary, issue_details = run_validations(clean_data)
    save_quality_report(validation_summary, issue_details)
    print_validation_summary(clean_data, validation_summary, issue_details)
    print("")

    print("Step 4: Calculating sales metrics and summary tables...")
    kpis = calculate_kpis(clean_data)
    summary_tables = build_summary_tables(clean_data)
    save_summary_tables(summary_tables)
    print_kpis(kpis)
    print(f"Saved sales summary to: {SALES_SUMMARY_FILE}")
    print("")

    print("Step 5: Exporting final formatted Excel report...")
    build_final_report()
    print(f"Saved final sales report to: {FINAL_REPORT_FILE}")
    print("")

    print("Pipeline complete.")
    print(f"Cleaned data: {CLEAN_DATA_FILE}")
    print(f"Data quality report: {DATA_QUALITY_FILE}")
    print(f"Sales summary: {SALES_SUMMARY_FILE}")
    print(f"Final report: {FINAL_REPORT_FILE}")


if __name__ == "__main__":
    main()
