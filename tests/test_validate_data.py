import pandas as pd

from src.validate_data import run_validations


def test_run_validations_detects_data_quality_issues():
    dataframe = pd.DataFrame(
        {
            "order_id": ["ORD-001", "ORD-001", "ORD-003", "ORD-004"],
            "order_date": [
                pd.Timestamp("2025-01-01"),
                pd.Timestamp("2025-01-02"),
                "not a date",
                pd.Timestamp("2025-01-04"),
            ],
            "customer_name": ["Jane Doe", "", "Sam Lee", "Ava Chen"],
            "product": ["Widget", "Widget", "Gadget", "Service"],
            "category": ["Office", "Office", "Tech", "Support"],
            "quantity": [2, 0, 3, 1],
            "unit_price": [10.0, 15.0, -5.0, 100.0],
            "discount": [0.1, 0.2, 1.2, 0.0],
            "region": ["North", "North", "South", "East"],
            "sales_rep": ["Alex", "Alex", "Riley", "Morgan"],
            "revenue": [18.0, 24.0, 10.0, 90.0],
        }
    )

    summary, issue_details = run_validations(dataframe)
    issue_counts = dict(zip(summary["check_name"], summary["issue_count"]))

    assert issue_counts["Missing values: customer_name"] == 1
    assert issue_counts["Duplicated order_id"] == 2
    assert issue_counts["Invalid or missing order_date"] == 1
    assert issue_counts["Invalid quantity"] == 1
    assert issue_counts["Invalid unit_price"] == 1
    assert issue_counts["Invalid discount"] == 1
    assert issue_counts["Invalid revenue"] == 3
    assert not issue_details.empty
    assert {
        "Missing values: customer_name",
        "Duplicated order_id",
        "Invalid or missing order_date",
        "Invalid quantity",
        "Invalid unit_price",
        "Invalid discount",
        "Invalid revenue",
    }.issubset(set(issue_details["check_name"]))
