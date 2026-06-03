"""Validate the cleaned sales dataset and create a data quality report."""

from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
INPUT_FILE = PROJECT_ROOT / "data" / "processed" / "clean_sales_data.xlsx"
OUTPUT_FILE = PROJECT_ROOT / "data" / "output" / "data_quality_report.xlsx"

KEY_COLUMNS = [
    "order_id",
    "order_date",
    "customer_name",
    "product",
    "category",
    "quantity",
    "unit_price",
    "discount",
    "region",
    "sales_rep",
    "revenue",
]

TEXT_COLUMNS = [
    "customer_name",
    "product",
    "category",
    "region",
    "sales_rep",
]


def load_cleaned_data(input_file: Path = INPUT_FILE) -> pd.DataFrame:
    """Load the cleaned sales data from an Excel file."""
    if not input_file.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")

    return pd.read_excel(input_file)


def is_missing(series: pd.Series) -> pd.Series:
    """Return True for blank or missing values in a column."""
    return series.isna() | series.astype(str).str.strip().eq("")


def add_issue_rows(
    issues: list[dict[str, object]],
    dataframe: pd.DataFrame,
    mask: pd.Series,
    check_name: str,
    message: str,
) -> None:
    """Add detailed issue rows for records that failed a validation check."""
    failed_rows = dataframe.loc[mask]

    for row_index, row in failed_rows.iterrows():
        issues.append(
            {
                "row_number": row_index + 2,
                "check_name": check_name,
                "order_id": row.get("order_id", ""),
                "problem": message,
            }
        )


def add_summary_row(
    summary_rows: list[dict[str, object]],
    check_name: str,
    issue_count: int,
    details: str,
) -> None:
    """Add one row to the validation summary."""
    summary_rows.append(
        {
            "check_name": check_name,
            "status": "PASS" if issue_count == 0 else "FAIL",
            "issue_count": issue_count,
            "details": details,
        }
    )


def validate_required_columns(
    dataframe: pd.DataFrame,
    summary_rows: list[dict[str, object]],
) -> None:
    """Check that all expected columns exist in the dataset."""
    missing_columns = [
        column_name for column_name in KEY_COLUMNS if column_name not in dataframe.columns
    ]

    add_summary_row(
        summary_rows,
        "Required columns",
        len(missing_columns),
        "Missing columns: " + ", ".join(missing_columns)
        if missing_columns
        else "All required columns are present.",
    )


def validate_missing_values(
    dataframe: pd.DataFrame,
    summary_rows: list[dict[str, object]],
    issues: list[dict[str, object]],
) -> None:
    """Check missing values in key columns."""
    for column_name in KEY_COLUMNS:
        if column_name not in dataframe.columns:
            continue

        mask = is_missing(dataframe[column_name])
        issue_count = int(mask.sum())

        add_summary_row(
            summary_rows,
            f"Missing values: {column_name}",
            issue_count,
            f"{issue_count} missing value(s) found in {column_name}.",
        )
        add_issue_rows(
            issues,
            dataframe,
            mask,
            f"Missing values: {column_name}",
            f"Missing value in {column_name}.",
        )


def validate_duplicate_order_ids(
    dataframe: pd.DataFrame,
    summary_rows: list[dict[str, object]],
    issues: list[dict[str, object]],
) -> None:
    """Check for duplicated order_id values."""
    if "order_id" not in dataframe.columns:
        return

    mask = dataframe["order_id"].duplicated(keep=False)
    issue_count = int(mask.sum())

    add_summary_row(
        summary_rows,
        "Duplicated order_id",
        issue_count,
        f"{issue_count} row(s) have a duplicated order_id.",
    )
    add_issue_rows(
        issues,
        dataframe,
        mask,
        "Duplicated order_id",
        "Duplicated order_id value.",
    )


def validate_dates(
    dataframe: pd.DataFrame,
    summary_rows: list[dict[str, object]],
    issues: list[dict[str, object]],
) -> None:
    """Check for invalid or missing order_date values."""
    if "order_date" not in dataframe.columns:
        return

    parsed_dates = pd.to_datetime(dataframe["order_date"], errors="coerce")
    mask = parsed_dates.isna()
    issue_count = int(mask.sum())

    add_summary_row(
        summary_rows,
        "Invalid or missing order_date",
        issue_count,
        f"{issue_count} row(s) have an invalid or missing order_date.",
    )
    add_issue_rows(
        issues,
        dataframe,
        mask,
        "Invalid or missing order_date",
        "Invalid or missing order_date.",
    )


def validate_positive_number(
    dataframe: pd.DataFrame,
    column_name: str,
    summary_rows: list[dict[str, object]],
    issues: list[dict[str, object]],
) -> None:
    """Check that a numeric column has values greater than zero."""
    if column_name not in dataframe.columns:
        return

    values = pd.to_numeric(dataframe[column_name], errors="coerce")
    mask = values.isna() | (values <= 0)
    issue_count = int(mask.sum())

    add_summary_row(
        summary_rows,
        f"Invalid {column_name}",
        issue_count,
        f"{issue_count} row(s) have missing or non-positive {column_name}.",
    )
    add_issue_rows(
        issues,
        dataframe,
        mask,
        f"Invalid {column_name}",
        f"{column_name} is missing or less than or equal to zero.",
    )


def validate_discount(
    dataframe: pd.DataFrame,
    summary_rows: list[dict[str, object]],
    issues: list[dict[str, object]],
) -> None:
    """Check that discount values are between 0 and 1."""
    if "discount" not in dataframe.columns:
        return

    values = pd.to_numeric(dataframe["discount"], errors="coerce")
    mask = values.isna() | (values < 0) | (values > 1)
    issue_count = int(mask.sum())

    add_summary_row(
        summary_rows,
        "Invalid discount",
        issue_count,
        f"{issue_count} row(s) have missing discount or discount outside 0 to 1.",
    )
    add_issue_rows(
        issues,
        dataframe,
        mask,
        "Invalid discount",
        "Discount is missing, less than 0, or greater than 1.",
    )


def validate_text_columns(
    dataframe: pd.DataFrame,
    summary_rows: list[dict[str, object]],
    issues: list[dict[str, object]],
) -> None:
    """Check required text fields for missing values."""
    for column_name in TEXT_COLUMNS:
        if column_name not in dataframe.columns:
            continue

        mask = is_missing(dataframe[column_name])
        issue_count = int(mask.sum())

        add_summary_row(
            summary_rows,
            f"Missing {column_name}",
            issue_count,
            f"{issue_count} row(s) are missing {column_name}.",
        )
        add_issue_rows(
            issues,
            dataframe,
            mask,
            f"Missing {column_name}",
            f"Missing {column_name}.",
        )


def validate_revenue(
    dataframe: pd.DataFrame,
    summary_rows: list[dict[str, object]],
    issues: list[dict[str, object]],
) -> None:
    """Check for missing, non-positive, or incorrectly calculated revenue."""
    required_columns = ["quantity", "unit_price", "discount", "revenue"]
    if any(column_name not in dataframe.columns for column_name in required_columns):
        return

    quantity = pd.to_numeric(dataframe["quantity"], errors="coerce")
    unit_price = pd.to_numeric(dataframe["unit_price"], errors="coerce")
    discount = pd.to_numeric(dataframe["discount"], errors="coerce")
    revenue = pd.to_numeric(dataframe["revenue"], errors="coerce")

    expected_revenue = quantity * unit_price * (1 - discount)
    revenue_difference = (revenue - expected_revenue).abs()
    mask = revenue.isna() | (revenue <= 0) | (revenue_difference > 0.01)
    issue_count = int(mask.sum())

    add_summary_row(
        summary_rows,
        "Invalid revenue",
        issue_count,
        f"{issue_count} row(s) have missing, non-positive, or mismatched revenue.",
    )
    add_issue_rows(
        issues,
        dataframe,
        mask,
        "Invalid revenue",
        "Revenue is missing, less than or equal to zero, or does not match the expected calculation.",
    )


def run_validations(dataframe: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Run all validation checks and return summary and issue details."""
    summary_rows: list[dict[str, object]] = []
    issues: list[dict[str, object]] = []

    validate_required_columns(dataframe, summary_rows)
    validate_missing_values(dataframe, summary_rows, issues)
    validate_duplicate_order_ids(dataframe, summary_rows, issues)
    validate_dates(dataframe, summary_rows, issues)
    validate_positive_number(dataframe, "quantity", summary_rows, issues)
    validate_positive_number(dataframe, "unit_price", summary_rows, issues)
    validate_discount(dataframe, summary_rows, issues)
    validate_text_columns(dataframe, summary_rows, issues)
    validate_revenue(dataframe, summary_rows, issues)

    summary = pd.DataFrame(summary_rows)
    issue_details = pd.DataFrame(issues)

    return summary, issue_details


def save_quality_report(
    summary: pd.DataFrame,
    issue_details: pd.DataFrame,
    output_file: Path = OUTPUT_FILE,
) -> None:
    """Save the data quality report to an Excel file."""
    output_file.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Validation Summary", index=False)
        issue_details.to_excel(writer, sheet_name="Issue Details", index=False)


def print_validation_summary(
    dataframe: pd.DataFrame,
    summary: pd.DataFrame,
    issue_details: pd.DataFrame,
) -> None:
    """Print a clear validation summary in the terminal."""
    failed_checks = summary[summary["status"] == "FAIL"]

    print("Data Validation Summary")
    print("=======================")
    print(f"Rows checked: {len(dataframe)}")
    print(f"Checks run: {len(summary)}")
    print(f"Failed checks: {len(failed_checks)}")
    print(f"Issue rows found: {len(issue_details)}")
    print("")

    for _, row in summary.iterrows():
        print(f"{row['status']} - {row['check_name']}: {row['issue_count']}")

    print("")
    print(f"Saved data quality report to: {OUTPUT_FILE}")


def main() -> None:
    """Load the cleaned data, validate it, and save a quality report."""
    dataframe = load_cleaned_data()
    summary, issue_details = run_validations(dataframe)

    save_quality_report(summary, issue_details)
    print_validation_summary(dataframe, summary, issue_details)


if __name__ == "__main__":
    main()
