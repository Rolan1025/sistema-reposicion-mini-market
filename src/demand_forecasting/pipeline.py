from __future__ import annotations

from pathlib import Path

import pandas as pd

from demand_forecasting.data import daily_product_series, load_inventory_csv, load_sales_csv
from demand_forecasting.features import build_supervised_dataset
from demand_forecasting.forecasting import forecast_demand
from demand_forecasting.models import load_model_bundle, save_model_bundle, train_and_select_model
from demand_forecasting.replenishment import recommend_replenishment


def train_pipeline(
    sales_csv: str | Path,
    model_path: str | Path,
    metrics_path: str | Path | None = None,
    test_days: int = 30,
) -> tuple[Path, pd.DataFrame]:
    sales = load_sales_csv(sales_csv)
    series = daily_product_series(sales)
    X, y, _, metadata = build_supervised_dataset(series)
    bundle = train_and_select_model(X, y, metadata, test_days=test_days)
    saved_path = save_model_bundle(bundle, model_path)
    if metrics_path is not None:
        metrics_path = Path(metrics_path)
        metrics_path.parent.mkdir(parents=True, exist_ok=True)
        bundle.metrics.to_csv(metrics_path, index=False)
    return saved_path, bundle.metrics


def predict_pipeline(
    sales_csv: str | Path,
    inventory_csv: str | Path,
    model_path: str | Path,
    horizon_days: int,
    forecast_path: str | Path | None = None,
    recommendations_path: str | Path | None = None,
    safety_stock_days: int = 3,
    review_period_days: int = 1,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    sales = load_sales_csv(sales_csv)
    inventory = load_inventory_csv(inventory_csv)
    history = daily_product_series(sales)
    bundle = load_model_bundle(model_path)
    forecast = forecast_demand(history, bundle, horizon_days=horizon_days)
    recommendations = recommend_replenishment(
        forecast,
        inventory,
        safety_stock_days=safety_stock_days,
        review_period_days=review_period_days,
    )

    if forecast_path is not None:
        forecast_path = Path(forecast_path)
        forecast_path.parent.mkdir(parents=True, exist_ok=True)
        forecast.to_csv(forecast_path, index=False)
    if recommendations_path is not None:
        recommendations_path = Path(recommendations_path)
        recommendations_path.parent.mkdir(parents=True, exist_ok=True)
        recommendations.to_csv(recommendations_path, index=False)

    return forecast, recommendations

