from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import joblib
import numpy as np
import pandas as pd
from sklearn.ensemble import RandomForestRegressor
from sklearn.linear_model import Ridge
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler

from demand_forecasting.metrics import metrics_frame, regression_metrics


@dataclass
class ModelBundle:
    model_name: str
    model: object
    feature_columns: list[str]
    metrics: pd.DataFrame
    trained_until: str


def candidate_models(random_state: int = 42) -> dict[str, object]:
    return {
        "ridge": Pipeline(
            steps=[
                ("scaler", StandardScaler()),
                ("regressor", Ridge(alpha=1.0, random_state=random_state)),
            ]
        ),
        "random_forest": RandomForestRegressor(
            n_estimators=160,
            max_depth=12,
            min_samples_leaf=3,
            random_state=random_state,
            n_jobs=-1,
        ),
    }


def seasonal_naive_predictions(metadata_test: pd.DataFrame, full_metadata: pd.DataFrame) -> np.ndarray:
    lookup = full_metadata[["date", "product_id", "units_sold"]].copy()
    lookup["date"] = pd.to_datetime(lookup["date"]) + pd.Timedelta(days=7)
    shifted = metadata_test[["date", "product_id"]].merge(
        lookup,
        on=["date", "product_id"],
        how="left",
    )
    fallback = full_metadata.groupby("product_id")["units_sold"].transform("mean")
    fallback_by_product = full_metadata.assign(fallback=fallback).groupby("product_id")["fallback"].last()
    shifted["units_sold"] = shifted["units_sold"].fillna(shifted["product_id"].map(fallback_by_product))
    return shifted["units_sold"].fillna(0).to_numpy()


def train_and_select_model(
    X: pd.DataFrame,
    y: pd.Series,
    metadata: pd.DataFrame,
    test_days: int = 30,
    random_state: int = 42,
) -> ModelBundle:
    cutoff = metadata["date"].max() - pd.Timedelta(days=test_days)
    train_mask = metadata["date"] <= cutoff
    test_mask = metadata["date"] > cutoff
    if train_mask.sum() == 0 or test_mask.sum() == 0:
        raise ValueError("Not enough rows after feature engineering for train/test split.")

    X_train, y_train = X.loc[train_mask], y.loc[train_mask]
    X_test, y_test = X.loc[test_mask], y.loc[test_mask]
    metadata_test = metadata.loc[test_mask].copy()

    rows = []
    trained_models = {}

    baseline_pred = seasonal_naive_predictions(metadata_test, metadata)
    rows.append({"model": "seasonal_naive_7d", **regression_metrics(y_test, baseline_pred)})

    for name, model in candidate_models(random_state).items():
        model.fit(X_train, y_train)
        pred = np.maximum(model.predict(X_test), 0)
        rows.append({"model": name, **regression_metrics(y_test, pred)})
        trained_models[name] = model

    metrics = metrics_frame(rows)
    best_name = metrics.iloc[0]["model"]

    if best_name == "seasonal_naive_7d":
        best_name = metrics[metrics["model"] != "seasonal_naive_7d"].iloc[0]["model"]

    final_model = candidate_models(random_state)[best_name]
    final_model.fit(X, y)
    return ModelBundle(
        model_name=best_name,
        model=final_model,
        feature_columns=list(X.columns),
        metrics=metrics,
        trained_until=str(metadata["date"].max().date()),
    )


def save_model_bundle(bundle: ModelBundle, path: str | Path) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    joblib.dump(bundle, path)
    return path


def load_model_bundle(path: str | Path) -> ModelBundle:
    return joblib.load(path)

