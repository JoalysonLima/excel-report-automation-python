"""Generate simulated monthly sales Excel files for the project."""

from calendar import monthrange
from datetime import date, timedelta
from pathlib import Path
import random

from openpyxl import Workbook


OUTPUT_DIRECTORY = Path(__file__).resolve().parents[1] / "data" / "raw"
YEAR = 2025
ROWS_PER_MONTH = 80
RANDOM_SEED = 42

CUSTOMERS = [
    "Acme Retail",
    "Bluebird Market",
    "City Corner Shop",
    "Evergreen Supplies",
    "Harbor Stores",
    "Northstar Traders",
    "Oak & Pine Goods",
    "Summit Wholesale",
]

PRODUCTS = [
    ("Laptop Stand", "Office", 34.90),
    ("Wireless Mouse", "Electronics", 24.50),
    ("Mechanical Keyboard", "Electronics", 79.99),
    ("Desk Lamp", "Office", 42.00),
    ("Notebook Set", "Stationery", 12.75),
    ("Gel Pen Pack", "Stationery", 8.50),
    ("Water Bottle", "Lifestyle", 18.25),
    ("Travel Backpack", "Lifestyle", 54.90),
]

REGIONS = ["North", "South", "East", "West"]
SALES_REPS = ["Alice Carter", "Ben Turner", "Chloe Martin", "Daniel Reed"]
DISCOUNTS = [0, 0, 0, 0.05, 0.10, 0.15]

HEADERS = [
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
]


def random_date_in_month(year: int, month: int) -> date:
    """Return a random date within the requested month."""
    last_day = monthrange(year, month)[1]
    first_day = date(year, month, 1)
    return first_day + timedelta(days=random.randint(0, last_day - 1))


def create_sales_rows(year: int, month: int) -> list[list[object]]:
    """Create sales rows with a few intentional data quality issues."""
    rows = []

    for row_number in range(1, ROWS_PER_MONTH + 1):
        product, category, unit_price = random.choice(PRODUCTS)
        order_id = f"ORD-{year}{month:02d}-{row_number:04d}"

        rows.append(
            [
                order_id,
                random_date_in_month(year, month),
                random.choice(CUSTOMERS),
                product,
                category,
                random.randint(1, 12),
                unit_price,
                random.choice(DISCOUNTS),
                random.choice(REGIONS),
                random.choice(SALES_REPS),
            ]
        )

    # These controlled issues are intentionally predictable for later validation.
    rows[5][0] = rows[4][0]  # Duplicated order ID
    rows[11][2] = None  # Missing customer name
    rows[17][8] = rows[17][8].lower()  # Inconsistent region casing
    rows[23][3] = f"  {rows[23][3]}  "  # Extra spaces around product name
    rows[29][5] = 0  # Invalid quantity
    rows[35][6] = -10.00  # Invalid unit price
    rows[41][7] = 1.25  # Invalid discount
    rows[47][9] = rows[47][9].upper()  # Inconsistent sales rep casing
    rows[53][8] = None  # Missing region

    return rows


def write_monthly_file(year: int, month: int) -> Path:
    """Write one month of simulated sales data to an Excel file."""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Sales"
    worksheet.append(HEADERS)

    for row in create_sales_rows(year, month):
        worksheet.append(row)

    for cell in worksheet["B"][1:]:
        cell.number_format = "yyyy-mm-dd"

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    output_file = OUTPUT_DIRECTORY / f"sales_{year}_{month:02d}.xlsx"
    workbook.save(output_file)
    return output_file


def main() -> None:
    """Generate one Excel sales file for every month of the configured year."""
    random.seed(RANDOM_SEED)
    OUTPUT_DIRECTORY.mkdir(parents=True, exist_ok=True)

    print(f"Generating monthly sales files in: {OUTPUT_DIRECTORY}")
    for month in range(1, 13):
        output_file = write_monthly_file(YEAR, month)
        print(f"Created {output_file.name}")

    print("Done. Generated 12 monthly Excel files.")


if __name__ == "__main__":
    main()
