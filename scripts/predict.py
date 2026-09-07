"""Comando para crear forecast, recomendaciones y el snapshot SQLite."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from demand_forecasting.config import (
    DEFAULT_INVENTORY_CSV,
    DEFAULT_MODEL_PATH,
    DEFAULT_SALES_CSV,
    REPORTS_DIR,
)
from demand_forecasting.pipeline import predict_pipeline
from demand_forecasting.storage import write_report_tables


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create demand forecasts and replenishment recommendations.")
    parser.add_argument("--sales-csv", default=str(DEFAULT_SALES_CSV))
    parser.add_argument("--inventory-csv", default=str(DEFAULT_INVENTORY_CSV))
    parser.add_argument("--model-path", default=str(DEFAULT_MODEL_PATH))
    parser.add_argument("--horizon-days", type=int, default=14)
    parser.add_argument("--safety-stock-days", type=int, default=3)
    parser.add_argument("--review-period-days", type=int, default=1)
    parser.add_argument("--forecast-path", default=str(REPORTS_DIR / "forecast.csv"))
    parser.add_argument("--recommendations-path", default=str(REPORTS_DIR / "recommendations.csv"))
    parser.add_argument("--sqlite-path", default=str(REPORTS_DIR / "mini_market_reports.db"))
    return parser.parse_args()


def main() -> None:
    """Carga el modelo ya entrenado y produce las salidas operativas."""
    args = parse_args()
    forecast, recommendations = predict_pipeline(
        sales_csv=args.sales_csv,
        inventory_csv=args.inventory_csv,
        model_path=args.model_path,
        horizon_days=args.horizon_days,
        forecast_path=args.forecast_path,
        recommendations_path=args.recommendations_path,
        safety_stock_days=args.safety_stock_days,
        review_period_days=args.review_period_days,
    )
    metrics_path = REPORTS_DIR / "metrics.csv"
    metrics = pd.read_csv(metrics_path) if metrics_path.exists() else None
    db_path = write_report_tables(args.sqlite_path, forecast, recommendations, metrics)
    print(f"Forecast rows: {len(forecast)}")
    print(f"Recommendations rows: {len(recommendations)}")
    print(f"SQLite report database: {db_path}")
    print(recommendations.head(10).to_string(index=False))


if __name__ == "__main__":
    main()
