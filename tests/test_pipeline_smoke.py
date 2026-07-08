from pathlib import Path

from demand_forecasting.data import write_sample_data
from demand_forecasting.pipeline import predict_pipeline, train_pipeline


def test_pipeline_smoke(tmp_path: Path):
    sales_path = tmp_path / "sales.csv"
    inventory_path = tmp_path / "inventory.csv"
    model_path = tmp_path / "model.joblib"
    metrics_path = tmp_path / "metrics.csv"
    forecast_path = tmp_path / "forecast.csv"
    recommendations_path = tmp_path / "recommendations.csv"

    write_sample_data(sales_path, inventory_path)
    _, metrics = train_pipeline(sales_path, model_path, metrics_path, test_days=30)
    forecast, recommendations = predict_pipeline(
        sales_path,
        inventory_path,
        model_path,
        horizon_days=7,
        forecast_path=forecast_path,
        recommendations_path=recommendations_path,
    )

    assert model_path.exists()
    assert metrics_path.exists()
    assert not metrics.empty
    assert len(forecast) > 0
    assert len(recommendations) > 0
    assert forecast_path.exists()
    assert recommendations_path.exists()

