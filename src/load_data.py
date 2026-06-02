"""Load and combine monthly sales Excel files."""

from pathlib import Path

import pandas as pd


RAW_DATA_DIRECTORY = Path(__file__).resolve().parents[1] / "data" / "raw"


def load_monthly_files(raw_data_directory: Path = RAW_DATA_DIRECTORY) -> pd.DataFrame:
    """Load all monthly Excel files and combine them into one DataFrame."""
    excel_files = sorted(raw_data_directory.glob("*.xlsx"))
    dataframes = []

    for excel_file in excel_files:
        try:
            dataframe = pd.read_excel(excel_file)
        except Exception as error:
            raise RuntimeError(f"Could not read {excel_file.name}") from error

        dataframe["source_file"] = excel_file.name
        dataframes.append(dataframe)

    if not dataframes:
        return pd.DataFrame()

    return pd.concat(dataframes, ignore_index=True)


def print_summary(dataframe: pd.DataFrame) -> None:
    """Print basic information about the combined data."""
    source_file_counts = dataframe["source_file"].value_counts().sort_index()

    print(f"Files loaded: {len(source_file_counts)}")
    print(f"Total rows: {len(dataframe)}")
    print(f"Columns: {list(dataframe.columns)}")
    print("Rows per source file:")

    for source_file, row_count in source_file_counts.items():
        print(f"  {source_file}: {row_count}")


def main() -> None:
    """Load monthly Excel files and print a summary."""
    dataframe = load_monthly_files()

    if dataframe.empty:
        print(f"No .xlsx files found in: {RAW_DATA_DIRECTORY}")
        return

    print_summary(dataframe)


if __name__ == "__main__":
    main()
