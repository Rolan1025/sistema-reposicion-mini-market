from __future__ import annotations

import pandas as pd

from demand_forecasting.features import build_future_feature_row
from demand_forecasting.models import ModelBundle


def forecast_demand(
    history: pd.DataFrame,
    bundle: ModelBundle,
    horizon_days: int = 14,
) -> pd.DataFrame:
    history = history.copy()
    history["date"] = pd.to_datetime(history["date"])
    history = history.sort_values(["product_id", "date"]).reset_index(drop=True)

    products = history[["product_id", "product_name", "category"]].drop_duplicates()
    last_prices = history.groupby("product_id")["unit_price"].last().to_dict()
    current = history.copy()
    last_date = current["date"].max()
    rows = []

    for step in range(1, horizon_days + 1):
        target_date = last_date + pd.Timedelta(days=step)
        day_rows = []
        for product in products.itertuples(index=False):
            row = build_future_feature_row(
                current,
                product.product_id,
                product.product_name,
                product.category,
                target_date,
                last_prices.get(product.product_id, 0),
                bundle.feature_columns,
            )
            prediction = float(max(bundle.model.predict(row)[0], 0))
            day_rows.append(
                {
                    "date": target_date,
                    "product_id": product.product_id,
                    "product_name": product.product_name,
                    "category": product.category,
                    "forecast_units": round(prediction, 2),
                    "unit_price": last_prices.get(product.product_id, 0),
                    "model_name": bundle.model_name,
                }
            )

        forecast_day = pd.DataFrame(day_rows)
        rows.extend(day_rows)
        current = pd.concat(
            [
                current,
                forecast_day.rename(columns={"forecast_units": "units_sold"})[
                    ["date", "product_id", "product_name", "category", "units_sold", "unit_price"]
                ],
            ],
            ignore_index=True,
        )

    return pd.DataFrame(rows)

