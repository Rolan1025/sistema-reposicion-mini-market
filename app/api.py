from __future__ import annotations

import sys
from pathlib import Path

from fastapi import FastAPI, Query

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from demand_forecasting.config import DEFAULT_INVENTORY_CSV, DEFAULT_MODEL_PATH, DEFAULT_SALES_CSV, REPORTS_DIR
from demand_forecasting.data import write_sample_data
from demand_forecasting.pipeline import predict_pipeline, train_pipeline

app = FastAPI(title="Mini-market demand forecasting API", version="0.1.0")


@app.get("/health")
def health() -> dict:
    return {"status": "ok"}


@app.post("/sample-data")
def sample_data() -> dict:
    sales_path, inventory_path = write_sample_data(DEFAULT_SALES_CSV, DEFAULT_INVENTORY_CSV)
    return {"sales_csv": str(sales_path), "inventory_csv": str(inventory_path)}


@app.post("/train")
def train(test_days: int = Query(30, ge=7, le=120)) -> dict:
    if not DEFAULT_SALES_CSV.exists() or not DEFAULT_INVENTORY_CSV.exists():
        write_sample_data(DEFAULT_SALES_CSV, DEFAULT_INVENTORY_CSV)
    model_path, metrics = train_pipeline(
        DEFAULT_SALES_CSV,
        DEFAULT_MODEL_PATH,
        REPORTS_DIR / "metrics.csv",
        test_days=test_days,
    )
    return {"model_path": str(model_path), "metrics": metrics.to_dict(orient="records")}


@app.get("/predict")
def predict(
    horizon_days: int = Query(14, ge=1, le=60),
    safety_stock_days: int = Query(3, ge=0, le=30),
    review_period_days: int = Query(1, ge=1, le=14),
) -> dict:
    if not DEFAULT_SALES_CSV.exists() or not DEFAULT_INVENTORY_CSV.exists():
        write_sample_data(DEFAULT_SALES_CSV, DEFAULT_INVENTORY_CSV)
    if not DEFAULT_MODEL_PATH.exists():
        train_pipeline(DEFAULT_SALES_CSV, DEFAULT_MODEL_PATH, REPORTS_DIR / "metrics.csv")
    forecast, recommendations = predict_pipeline(
        DEFAULT_SALES_CSV,
        DEFAULT_INVENTORY_CSV,
        DEFAULT_MODEL_PATH,
        horizon_days=horizon_days,
        forecast_path=REPORTS_DIR / "forecast.csv",
        recommendations_path=REPORTS_DIR / "recommendations.csv",
        safety_stock_days=safety_stock_days,
        review_period_days=review_period_days,
    )
    return {
        "forecast": forecast.to_dict(orient="records"),
        "recommendations": recommendations.to_dict(orient="records"),
    }
