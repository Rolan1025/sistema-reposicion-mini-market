import pytest

from demand_forecasting.metrics import mae, mape, smape


def test_metrics_are_reasonable():
    y_true = [10, 20, 30]
    y_pred = [12, 18, 33]
    assert mae(y_true, y_pred) == pytest.approx(2.333333, rel=1e-4)
    assert mape(y_true, y_pred) == pytest.approx(13.333333, rel=1e-4)
    assert smape(y_true, y_pred) > 0


def test_mape_ignores_zero_actuals():
    assert mape([0, 10], [5, 12]) == pytest.approx(20.0)
