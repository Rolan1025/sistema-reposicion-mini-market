import pandas as pd

from demand_forecasting.data import daily_product_series, generate_synthetic_sales
from demand_forecasting.evaluation import rolling_origin_evaluation


def test_multi_horizon_evaluation_has_expected_prediction_counts():
    sales = generate_synthetic_sales("2024-01-01", "2024-06-30")
    series = daily_product_series(sales)
    predictions, by_horizon, by_product, ranking = rolling_origin_evaluation(
        series,
        horizons=(7, 14),
    )

    products = series["product_id"].nunique()
    models = 3
    expected_rows = products * models * (7 + 14)
    assert len(predictions) == expected_rows
    assert set(by_horizon["horizon_days"]) == {7, 14}
    assert set(by_horizon["n_predictions"]) == {products * 7, products * 14}
    assert not by_product.empty
    assert ranking.iloc[0]["rank"] == 1
