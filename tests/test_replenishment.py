import pandas as pd

from demand_forecasting.replenishment import recommend_replenishment, round_to_pack


def test_round_to_pack():
    assert round_to_pack(0, 12) == 0
    assert round_to_pack(1, 12) == 12
    assert round_to_pack(25, 12) == 36


def test_recommend_replenishment_generates_order():
    forecast = pd.DataFrame(
        {
            "date": pd.date_range("2025-01-01", periods=5),
            "product_id": ["P001"] * 5,
            "product_name": ["Leche"] * 5,
            "forecast_units": [10, 10, 10, 10, 10],
        }
    )
    inventory = pd.DataFrame(
        {
            "product_id": ["P001"],
            "product_name": ["Leche"],
            "category": ["Lacteos"],
            "stock_on_hand": [5],
            "on_order": [0],
            "lead_time_days": [2],
            "pack_size": [12],
            "max_stock": [100],
        }
    )
    recommendations = recommend_replenishment(forecast, inventory, safety_stock_days=3)
    assert recommendations.iloc[0]["suggested_order_qty"] == 60
    assert recommendations.iloc[0]["priority"] == "alta"
