from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd


REQUIRED_SALES_COLUMNS = {
    "date",
    "product_id",
    "product_name",
    "category",
    "units_sold",
    "unit_price",
}

REQUIRED_INVENTORY_COLUMNS = {
    "product_id",
    "product_name",
    "stock_on_hand",
    "on_order",
    "lead_time_days",
    "pack_size",
}


@dataclass(frozen=True)
class ProductSeed:
    product_id: str
    product_name: str
    category: str
    base_demand: float
    unit_price: float
    perishability_days: int
    pack_size: int
    lead_time_days: int


PRODUCT_CATALOG = [
    ProductSeed("P001", "Leche entera 1L", "Lacteos", 18, 1.25, 7, 12, 2),
    ProductSeed("P002", "Pan tajado", "Panaderia", 24, 1.10, 5, 10, 1),
    ProductSeed("P003", "Huevos x12", "Basicos", 16, 2.40, 21, 12, 3),
    ProductSeed("P004", "Arroz 1kg", "Despensa", 10, 1.70, 365, 24, 4),
    ProductSeed("P005", "Atun lata", "Despensa", 8, 1.95, 730, 24, 5),
    ProductSeed("P006", "Yogur individual", "Lacteos", 20, 0.85, 10, 20, 2),
    ProductSeed("P007", "Gaseosa 1.5L", "Bebidas", 14, 1.60, 180, 12, 3),
    ProductSeed("P008", "Manzana roja kg", "Frutas", 13, 2.20, 8, 8, 2),
    ProductSeed("P009", "Papel higienico x4", "Aseo", 7, 2.90, 365, 12, 6),
    ProductSeed("P010", "Cafe molido 250g", "Bebidas", 6, 3.40, 365, 10, 5),
]


def load_sales_csv(path: str | Path) -> pd.DataFrame:
    df = pd.read_csv(path, parse_dates=["date"])
    missing = REQUIRED_SALES_COLUMNS - set(df.columns)
    if missing:
        raise ValueError(f"Missing required sales columns: {sorted(missing)}")
    df = df.copy()
    df["product_id"] = df["product_id"].astype(str)
    df["units_sold"] = pd.to_numeric(df["units_sold"], errors="coerce").fillna(0)
    df["unit_price"] = pd.to_numeric(df["unit_price"], errors="coerce").fillna(0)
    df["revenue"] = df["units_sold"] * df["unit_price"]
    return df.sort_values(["date", "product_id"]).reset_index(drop=True)


def load_inventory_csv(path: str | Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    missing = REQUIRED_INVENTORY_COLUMNS - set(df.columns)
    if missing:
        raise ValueError(f"Missing required inventory columns: {sorted(missing)}")
    df = df.copy()
    df["product_id"] = df["product_id"].astype(str)
    numeric_cols = ["stock_on_hand", "on_order", "lead_time_days", "pack_size"]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df


def daily_product_series(sales: pd.DataFrame) -> pd.DataFrame:
    sales = sales.copy()
    sales["date"] = pd.to_datetime(sales["date"])
    grouped = (
        sales.groupby(["date", "product_id", "product_name", "category"], as_index=False)
        .agg(units_sold=("units_sold", "sum"), unit_price=("unit_price", "mean"))
        .sort_values(["product_id", "date"])
    )

    all_dates = pd.date_range(grouped["date"].min(), grouped["date"].max(), freq="D")
    products = grouped[["product_id", "product_name", "category"]].drop_duplicates()
    grid = products.merge(pd.DataFrame({"date": all_dates}), how="cross")

    complete = grid.merge(
        grouped,
        on=["date", "product_id", "product_name", "category"],
        how="left",
    )
    complete["units_sold"] = complete["units_sold"].fillna(0)
    complete["unit_price"] = complete.groupby("product_id")["unit_price"].ffill().bfill()
    return complete.sort_values(["product_id", "date"]).reset_index(drop=True)


def time_based_split(df: pd.DataFrame, test_days: int = 30) -> tuple[pd.DataFrame, pd.DataFrame]:
    cutoff = df["date"].max() - pd.Timedelta(days=test_days)
    train = df[df["date"] <= cutoff].copy()
    test = df[df["date"] > cutoff].copy()
    if train.empty or test.empty:
        raise ValueError("Not enough data for the requested test window.")
    return train, test


def generate_synthetic_sales(
    start_date: str = "2024-01-01",
    end_date: str = "2025-12-31",
    random_state: int = 42,
) -> pd.DataFrame:
    rng = np.random.default_rng(random_state)
    dates = pd.date_range(start_date, end_date, freq="D")
    rows = []

    for product in PRODUCT_CATALOG:
        trend = rng.normal(0.0008, 0.0003)
        for idx, date in enumerate(dates):
            dow = date.dayofweek
            month = date.month
            weekend_factor = 1.22 if dow in (5, 6) else 1.0
            payday_factor = 1.18 if date.day in (1, 15, 30) else 1.0
            season_factor = 1.0
            if product.category == "Bebidas" and month in (6, 7, 8, 12):
                season_factor = 1.25
            if product.category in ("Lacteos", "Panaderia") and dow == 0:
                season_factor *= 0.9
            if product.category == "Frutas" and month in (3, 4, 5):
                season_factor = 1.15

            promo = int(rng.random() < 0.08)
            promo_factor = 1.30 if promo else 1.0
            noise = rng.normal(1.0, 0.18)
            demand = product.base_demand * weekend_factor * payday_factor * season_factor
            demand *= promo_factor * (1 + trend * idx) * max(noise, 0.25)
            units = max(0, rng.poisson(max(demand, 0.1)))
            price = round(product.unit_price * (0.92 if promo else 1.0), 2)
            rows.append(
                {
                    "date": date.date().isoformat(),
                    "product_id": product.product_id,
                    "product_name": product.product_name,
                    "category": product.category,
                    "units_sold": int(units),
                    "unit_price": price,
                    "promo": promo,
                }
            )

    return pd.DataFrame(rows)


def generate_inventory_snapshot(random_state: int = 42) -> pd.DataFrame:
    rng = np.random.default_rng(random_state + 1)
    rows = []
    for product in PRODUCT_CATALOG:
        rows.append(
            {
                "product_id": product.product_id,
                "product_name": product.product_name,
                "category": product.category,
                "stock_on_hand": int(rng.integers(product.pack_size, product.pack_size * 5)),
                "on_order": int(rng.choice([0, product.pack_size])),
                "lead_time_days": product.lead_time_days,
                "pack_size": product.pack_size,
                "perishability_days": product.perishability_days,
                "max_stock": int(product.base_demand * 21),
            }
        )
    return pd.DataFrame(rows)


def write_sample_data(sales_path: str | Path, inventory_path: str | Path) -> tuple[Path, Path]:
    sales_path = Path(sales_path)
    inventory_path = Path(inventory_path)
    sales_path.parent.mkdir(parents=True, exist_ok=True)
    inventory_path.parent.mkdir(parents=True, exist_ok=True)
    generate_synthetic_sales().to_csv(sales_path, index=False)
    generate_inventory_snapshot().to_csv(inventory_path, index=False)
    return sales_path, inventory_path

