"""Backtesting temporal para saber cómo habría funcionado cada modelo.

Backtesting significa ocultar deliberadamente el tramo final del histórico,
predecirlo usando solo el pasado y después comparar contra las ventas reales.
"""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path

import numpy as np
import pandas as pd

from demand_forecasting.data import daily_product_series, load_sales_csv
from demand_forecasting.features import build_supervised_dataset
from demand_forecasting.forecasting import forecast_demand
from demand_forecasting.metrics import regression_metrics
from demand_forecasting.models import ModelBundle, candidate_models


DEFAULT_HORIZONS = (7, 14, 30)


def _fit_bundle(history: pd.DataFrame, model_name: str, random_state: int = 42) -> ModelBundle:
    """Entrena un modelo candidato usando únicamente el histórico recibido."""
    X, y, feature_columns, metadata = build_supervised_dataset(history)
    model = candidate_models(random_state)[model_name]
    model.fit(X, y)
    return ModelBundle(
        model_name=model_name,
        model=model,
        feature_columns=feature_columns,
        metrics=pd.DataFrame(),
        trained_until=str(metadata["date"].max().date()),
    )


def _seasonal_naive_forecast(history: pd.DataFrame, horizon_days: int) -> pd.DataFrame:
    """Pronostica repitiendo el valor observado siete días antes."""
    current = history.copy().sort_values(["product_id", "date"])
    current["date"] = pd.to_datetime(current["date"])
    products = current[["product_id", "product_name", "category"]].drop_duplicates()
    last_date = current["date"].max()
    rows: list[dict] = []

    for step in range(1, horizon_days + 1):
        target_date = last_date + pd.Timedelta(days=step)
        day_rows = []
        for product in products.itertuples(index=False):
            product_history = current[current["product_id"] == product.product_id].sort_values("date")
            values = product_history["units_sold"].astype(float).to_numpy()
            prediction = float(values[-7] if len(values) >= 7 else values.mean())
            day_rows.append(
                {
                    "date": target_date,
                    "product_id": product.product_id,
                    "product_name": product.product_name,
                    "category": product.category,
                    "forecast_units": max(prediction, 0),
                    "unit_price": float(product_history["unit_price"].iloc[-1]),
                    "model_name": "seasonal_naive_7d",
                }
            )
        day = pd.DataFrame(day_rows)
        rows.extend(day_rows)
        current = pd.concat(
            [
                current,
                day.rename(columns={"forecast_units": "units_sold"})[
                    ["date", "product_id", "product_name", "category", "units_sold", "unit_price"]
                ],
            ],
            ignore_index=True,
        )
    return pd.DataFrame(rows)


def rolling_origin_evaluation(
    series: pd.DataFrame,
    horizons: tuple[int, ...] = DEFAULT_HORIZONS,
    random_state: int = 42,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Evalúa baseline, Ridge y Random Forest en 7, 14 y 30 días.

    Para cada horizonte se mueve el punto de corte hacia atrás, se entrena con
    el pasado y se compara el pronóstico con el tramo real ocultado.
    """
    series = series.copy()
    series["date"] = pd.to_datetime(series["date"])
    max_date = series["date"].max()
    prediction_frames = []

    for horizon in horizons:
        # En 14 días, por ejemplo, las últimas 14 fechas quedan fuera del train.
        origin = max_date - pd.Timedelta(days=horizon)
        train_history = series[series["date"] <= origin].copy()
        actual = series[series["date"] > origin].copy()

        model_forecasts = {"seasonal_naive_7d": _seasonal_naive_forecast(train_history, horizon)}
        for model_name in candidate_models(random_state):
            bundle = _fit_bundle(train_history, model_name, random_state)
            model_forecasts[model_name] = forecast_demand(train_history, bundle, horizon_days=horizon)

        actual_cols = actual[
            ["date", "product_id", "product_name", "category", "units_sold"]
        ].rename(columns={"units_sold": "actual_units"})
        for model_name, forecast in model_forecasts.items():
            joined = forecast.merge(
                actual_cols,
                on=["date", "product_id", "product_name", "category"],
                how="inner",
            )
            joined["horizon_days"] = horizon
            joined["forecast_step"] = (joined["date"] - origin).dt.days
            joined["absolute_error"] = (joined["actual_units"] - joined["forecast_units"]).abs()
            prediction_frames.append(joined)

    predictions = pd.concat(prediction_frames, ignore_index=True)

    horizon_rows = []
    product_rows = []
    for (horizon, model_name), group in predictions.groupby(["horizon_days", "model_name"]):
        horizon_rows.append(
            {
                "horizon_days": int(horizon),
                "model": model_name,
                "n_predictions": len(group),
                **regression_metrics(group["actual_units"], group["forecast_units"]),
            }
        )
    for (horizon, model_name, product_id, product_name, category), group in predictions.groupby(
        ["horizon_days", "model_name", "product_id", "product_name", "category"]
    ):
        product_rows.append(
            {
                "horizon_days": int(horizon),
                "model": model_name,
                "product_id": product_id,
                "product_name": product_name,
                "category": category,
                "n_predictions": len(group),
                **regression_metrics(group["actual_units"], group["forecast_units"]),
            }
        )

    metrics_by_horizon = pd.DataFrame(horizon_rows).sort_values(["horizon_days", "smape", "mae"])
    metrics_by_product = pd.DataFrame(product_rows).sort_values(
        ["horizon_days", "model", "product_id"]
    )
    # El ranking promedia los tres horizontes. Menor SMAPE ocupa mejor puesto.
    ranking = (
        metrics_by_horizon.groupby("model", as_index=False)
        .agg(
            mean_mae=("mae", "mean"),
            mean_rmse=("rmse", "mean"),
            mean_mape=("mape", "mean"),
            mean_smape=("smape", "mean"),
        )
        .sort_values(["mean_smape", "mean_mae"])
        .reset_index(drop=True)
    )
    ranking["rank"] = np.arange(1, len(ranking) + 1)
    return predictions, metrics_by_horizon, metrics_by_product, ranking


def evaluate_sales_csv(
    sales_csv: str | Path,
    output_dir: str | Path,
    horizons: tuple[int, ...] = DEFAULT_HORIZONS,
) -> dict[str, Path]:
    """Ejecuta la evaluación y guarda sus cuatro archivos CSV de evidencia."""
    sales = load_sales_csv(sales_csv)
    series = daily_product_series(sales)
    predictions, by_horizon, by_product, ranking = rolling_origin_evaluation(series, horizons)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    outputs = {
        "predictions": output_dir / "backtest_predictions.csv",
        "by_horizon": output_dir / "metrics_by_horizon.csv",
        "by_product": output_dir / "metrics_by_product.csv",
        "ranking": output_dir / "model_ranking.csv",
    }
    predictions.to_csv(outputs["predictions"], index=False)
    by_horizon.to_csv(outputs["by_horizon"], index=False)
    by_product.to_csv(outputs["by_product"], index=False)
    ranking.to_csv(outputs["ranking"], index=False)
    return outputs
