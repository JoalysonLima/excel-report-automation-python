import pandas as pd

from src.clean_data import clean_sales_data, to_snake_case


def test_to_snake_case_normalizes_column_names():
    assert to_snake_case(" Order Date ") == "order_date"
    assert to_snake_case("Unit Price ($)") == "unit_price"
    assert to_snake_case("Sales-Rep Name") == "sales_rep_name"


def test_clean_sales_data_standardizes_fields_and_calculates_revenue():
    dataframe = pd.DataFrame(
        {
            " Order ID ": [" ord-001 "],
            "Order Date": ["2025-01-15"],
            "Customer Name": [" jane doe "],
            "Product": [" widget pro "],
            "Category": [" office supplies "],
            "Quantity": ["3"],
            "Unit Price": ["10.50"],
            "Discount": ["0.10"],
            "Region": [" north "],
            "Sales Rep": [" alex smith "],
        }
    )

    cleaned = clean_sales_data(dataframe)

    assert list(cleaned.columns) == [
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
    assert cleaned.loc[0, "order_id"] == "ORD-001"
    assert cleaned.loc[0, "customer_name"] == "Jane Doe"
    assert cleaned.loc[0, "product"] == "Widget Pro"
    assert cleaned.loc[0, "category"] == "Office Supplies"
    assert cleaned.loc[0, "region"] == "North"
    assert cleaned.loc[0, "sales_rep"] == "Alex Smith"
    assert pd.api.types.is_datetime64_any_dtype(cleaned["order_date"])
    assert pd.api.types.is_numeric_dtype(cleaned["quantity"])
    assert pd.api.types.is_numeric_dtype(cleaned["unit_price"])
    assert pd.api.types.is_numeric_dtype(cleaned["discount"])
    assert cleaned.loc[0, "revenue"] == 28.35
