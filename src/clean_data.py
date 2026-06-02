"""Clean the combined monthly sales data and save the processed dataset."""

from pathlib import Path
import re

import pandas as pd

from load_data import load_monthly_files


PROCESSED_DATA_DIRECTORY = Path(__file__).resolve().parents[1] / "data" / "processed"
OUTPUT_FILE = PROCESSED_DATA_DIRECTORY / "clean_sales_data.xlsx"

TEXT_COLUMNS_TO_TITLE_CASE = [
    "customer_name",
    "product",
    "category",
    "region",
    "sales_rep",
]
NUMERIC_COLUMNS = ["quantity", "unit_price", "discount"]


def to_snake_case(column_name: object) -> str:
    """Return a lowercase snake_case version of a column name."""
    name = str(column_name).strip()
    name = re.sub(r"[^a-zA-Z0-9]+", "_", name)
    return name.strip("_").lower()


def clean_sales_data(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Clean the combined monthly sales data."""
    cleaned_dataframe = dataframe.copy()
    cleaned_dataframe.columns = [
        to_snake_case(column_name) for column_name in cleaned_dataframe.columns
    ]

    text_columns = cleaned_dataframe.select_dtypes(include="object").columns
    for column_name in text_columns:
        cleaned_dataframe[column_name] = cleaned_dataframe[column_name].str.strip()

    if "order_id" in cleaned_dataframe.columns:
        cleaned_dataframe["order_id"] = cleaned_dataframe["order_id"].str.upper()

    for column_name in TEXT_COLUMNS_TO_TITLE_CASE:
        if column_name in cleaned_dataframe.columns:
            cleaned_dataframe[column_name] = cleaned_dataframe[column_name].str.title()

    if "order_date" in cleaned_dataframe.columns:
        cleaned_dataframe["order_date"] = pd.to_datetime(
            cleaned_dataframe["order_date"], errors="coerce"
        )

    for column_name in NUMERIC_COLUMNS:
        if column_name in cleaned_dataframe.columns:
            cleaned_dataframe[column_name] = pd.to_numeric(
                cleaned_dataframe[column_name], errors="coerce"
            )

    cleaned_dataframe["revenue"] = (
        cleaned_dataframe["quantity"]
        * cleaned_dataframe["unit_price"]
        * (1 - cleaned_dataframe["discount"])
    )

    return cleaned_dataframe


def save_cleaned_data(
    dataframe: pd.DataFrame, output_file: Path = OUTPUT_FILE
) -> None:
    """Save the cleaned dataset to an Excel file."""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    dataframe.to_excel(output_file, index=False)


def main() -> None:
    """Load, clean, and save the monthly sales data."""
    dataframe = load_monthly_files()

    if dataframe.empty:
        print("No monthly sales data found.")
        return

    cleaned_dataframe = clean_sales_data(dataframe)
    save_cleaned_data(cleaned_dataframe)

    print(f"Cleaned rows: {len(cleaned_dataframe)}")
    print(f"Saved cleaned data to: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
