"""Analyze cleaned sales data and export summary tables."""

from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
INPUT_FILE = PROJECT_ROOT / "data" / "processed" / "clean_sales_data.xlsx"
OUTPUT_FILE = PROJECT_ROOT / "data" / "output" / "sales_summary.xlsx"


def load_cleaned_data(input_file: Path = INPUT_FILE) -> pd.DataFrame:
    """Load the cleaned sales data from an Excel file."""
    if not input_file.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")

    return pd.read_excel(input_file)


def calculate_kpis(dataframe: pd.DataFrame) -> dict[str, float | int]:
    """Calculate the main sales KPIs."""
    total_revenue = dataframe["revenue"].sum()
    total_orders = dataframe["order_id"].nunique()
    total_quantity_sold = dataframe["quantity"].sum()
    average_order_value = total_revenue / total_orders if total_orders else 0

    return {
        "total_revenue": total_revenue,
        "total_orders": total_orders,
        "total_quantity_sold": total_quantity_sold,
        "average_order_value": average_order_value,
    }


def summarize_revenue_by_month(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Return total revenue by month."""
    monthly_data = dataframe.copy()
    monthly_data["month"] = pd.to_datetime(monthly_data["order_date"]).dt.to_period("M")

    summary = (
        monthly_data.groupby("month", as_index=False)["revenue"]
        .sum()
        .sort_values("month")
    )
    summary["month"] = summary["month"].astype(str)

    return summary


def summarize_revenue_by_column(
    dataframe: pd.DataFrame,
    column_name: str,
) -> pd.DataFrame:
    """Return total revenue grouped by a selected column."""
    return (
        dataframe.groupby(column_name, as_index=False)["revenue"]
        .sum()
        .sort_values("revenue", ascending=False)
    )


def summarize_top_products(dataframe: pd.DataFrame, limit: int = 10) -> pd.DataFrame:
    """Return the top products by revenue."""
    return summarize_revenue_by_column(dataframe, "product").head(limit)


def build_summary_tables(dataframe: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Build all summary tables for the sales analysis workbook."""
    return {
        "Revenue by Month": summarize_revenue_by_month(dataframe),
        "Revenue by Category": summarize_revenue_by_column(dataframe, "category"),
        "Revenue by Region": summarize_revenue_by_column(dataframe, "region"),
        "Revenue by Sales Rep": summarize_revenue_by_column(dataframe, "sales_rep"),
        "Top Products": summarize_top_products(dataframe),
    }


def save_summary_tables(
    summary_tables: dict[str, pd.DataFrame],
    output_file: Path = OUTPUT_FILE,
) -> None:
    """Save summary tables to an Excel workbook with separate sheets."""
    output_file.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for sheet_name, summary_table in summary_tables.items():
            summary_table.to_excel(writer, sheet_name=sheet_name, index=False)


def print_kpis(kpis: dict[str, float | int]) -> None:
    """Print the main sales KPIs clearly in the terminal."""
    print("Sales Analysis Summary")
    print("======================")
    print(f"Total revenue: ${kpis['total_revenue']:,.2f}")
    print(f"Total orders: {kpis['total_orders']:,}")
    print(f"Total quantity sold: {kpis['total_quantity_sold']:,}")
    print(f"Average order value: ${kpis['average_order_value']:,.2f}")


def main() -> None:
    """Load cleaned sales data, analyze it, and save summary tables."""
    dataframe = load_cleaned_data()
    kpis = calculate_kpis(dataframe)
    summary_tables = build_summary_tables(dataframe)

    save_summary_tables(summary_tables)
    print_kpis(kpis)
    print("")
    print(f"Saved sales summary to: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
