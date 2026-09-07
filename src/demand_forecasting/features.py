"""Construcción de variables que el modelo usa para aprender patrones.

Una variable o feature es una pista entregada al algoritmo: día de la semana,
precio, promoción, ventas anteriores o promedio de las últimas semanas.
"""

from __future__ import annotations

import numpy as np
import pandas as pd


LAG_DAYS = [1, 2, 7, 14, 28]
ROLLING_WINDOWS = [7, 14, 28]


def add_calendar_features(df: pd.DataFrame) -> pd.DataFrame:
    """Convierte la fecha en pistas de calendario y ciclos semanales/mensuales."""
    out = df.copy()
    out["date"] = pd.to_datetime(out["date"])
    out["day_of_week"] = out["date"].dt.dayofweek
    out["day_of_month"] = out["date"].dt.day
    out["week_of_year"] = out["date"].dt.isocalendar().week.astype(int)
    out["month"] = out["date"].dt.month
    out["is_weekend"] = out["day_of_week"].isin([5, 6]).astype(int)
    out["is_payday"] = out["day_of_month"].isin([1, 15, 30]).astype(int)
    # Seno y coseno hacen que domingo y lunes queden próximos en un ciclo.
    out["dow_sin"] = np.sin(2 * np.pi * out["day_of_week"] / 7)
    out["dow_cos"] = np.cos(2 * np.pi * out["day_of_week"] / 7)
    out["month_sin"] = np.sin(2 * np.pi * out["month"] / 12)
    out["month_cos"] = np.cos(2 * np.pi * out["month"] / 12)
    return out


def add_lag_features(df: pd.DataFrame) -> pd.DataFrame:
    """Añade ventas pasadas y resúmenes móviles sin usar información futura."""
    out = df.sort_values(["product_id", "date"]).copy()
    grouped = out.groupby("product_id")["units_sold"]
    for lag in LAG_DAYS:
        # lag_7 significa: unidades vendidas por este producto hace siete días.
        out[f"lag_{lag}"] = grouped.shift(lag)
    for window in ROLLING_WINDOWS:
        # shift(1) excluye el día objetivo y evita fuga de información.
        out[f"rolling_mean_{window}"] = grouped.shift(1).rolling(window).mean().reset_index(level=0, drop=True)
        out[f"rolling_std_{window}"] = grouped.shift(1).rolling(window).std().reset_index(level=0, drop=True)
    return out


def build_supervised_dataset(series: pd.DataFrame) -> tuple[pd.DataFrame, pd.Series, list[str], pd.DataFrame]:
    """Construye X (pistas) e y (unidades que el modelo debe aprender)."""
    featured = add_lag_features(add_calendar_features(series))
    if "promo" not in featured.columns:
        featured["promo"] = 0
    featured = featured.dropna().reset_index(drop=True)
    featured["unit_price"] = featured["unit_price"].fillna(featured["unit_price"].median())

    base_features = [
        "unit_price",
        "promo",
        "day_of_week",
        "day_of_month",
        "week_of_year",
        "month",
        "is_weekend",
        "is_payday",
        "dow_sin",
        "dow_cos",
        "month_sin",
        "month_cos",
    ]
    lag_features = [f"lag_{lag}" for lag in LAG_DAYS]
    rolling_features = []
    for window in ROLLING_WINDOWS:
        rolling_features.extend([f"rolling_mean_{window}", f"rolling_std_{window}"])

    # Los modelos necesitan números: one-hot encoding convierte cada producto
    # y categoría en columnas de 0/1 sin inventar un orden entre ellos.
    encoded = pd.get_dummies(featured, columns=["product_id", "category"], prefix=["pid", "cat"])
    dummy_features = [col for col in encoded.columns if col.startswith("pid_") or col.startswith("cat_")]
    feature_columns = base_features + lag_features + rolling_features + dummy_features

    X = encoded[feature_columns].astype(float)
    y = encoded["units_sold"].astype(float)
    metadata = featured[["date", "product_id", "product_name", "category", "units_sold"]].copy()
    return X, y, feature_columns, metadata


def build_future_feature_row(
    history: pd.DataFrame,
    product_id: str,
    product_name: str,
    category: str,
    target_date: pd.Timestamp,
    unit_price: float,
    feature_columns: list[str],
) -> pd.DataFrame:
    """Crea las pistas de un producto para una fecha futura concreta.

    La promoción futura se fija en cero porque el prototipo no recibe un
    calendario de promociones planificadas.
    """
    product_history = history[history["product_id"] == product_id].sort_values("date")
    if product_history.empty:
        raise ValueError(f"No history found for product_id={product_id}")

    values = {
        "unit_price": unit_price,
        "promo": 0,
        "day_of_week": target_date.dayofweek,
        "day_of_month": target_date.day,
        "week_of_year": int(target_date.isocalendar().week),
        "month": target_date.month,
        "is_weekend": int(target_date.dayofweek in [5, 6]),
        "is_payday": int(target_date.day in [1, 15, 30]),
        "dow_sin": np.sin(2 * np.pi * target_date.dayofweek / 7),
        "dow_cos": np.cos(2 * np.pi * target_date.dayofweek / 7),
        "month_sin": np.sin(2 * np.pi * target_date.month / 12),
        "month_cos": np.cos(2 * np.pi * target_date.month / 12),
    }

    sales = product_history["units_sold"].astype(float).to_numpy()
    for lag in LAG_DAYS:
        values[f"lag_{lag}"] = sales[-lag] if len(sales) >= lag else sales.mean()
    for window in ROLLING_WINDOWS:
        recent = sales[-window:] if len(sales) >= window else sales
        values[f"rolling_mean_{window}"] = float(np.mean(recent))
        values[f"rolling_std_{window}"] = float(np.std(recent))

    values[f"pid_{product_id}"] = 1
    values[f"cat_{category}"] = 1
    row = pd.DataFrame([{col: values.get(col, 0) for col in feature_columns}])
    return row.astype(float)
