import pandas as pd

from src.analyze_data import calculate_kpis


def test_calculate_kpis_returns_basic_sales_metrics():
    dataframe = pd.DataFrame(
        {
            "order_id": ["ORD-001", "ORD-002", "ORD-002"],
            "quantity": [2, 3, 5],
            "revenue": [20.0, 30.0, 50.0],
        }
    )

    kpis = calculate_kpis(dataframe)

    assert kpis == {
        "total_revenue": 100.0,
        "total_orders": 2,
        "total_quantity_sold": 10,
        "average_order_value": 50.0,
    }
