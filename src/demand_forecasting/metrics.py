import numpy as np
import pandas as pd


def mae(y_true, y_pred) -> float:
    return float(np.mean(np.abs(np.asarray(y_true) - np.asarray(y_pred))))


def rmse(y_true, y_pred) -> float:
    return float(np.sqrt(np.mean((np.asarray(y_true) - np.asarray(y_pred)) ** 2)))


def mape(y_true, y_pred) -> float:
    y_true = np.asarray(y_true, dtype=float)
    y_pred = np.asarray(y_pred, dtype=float)
    mask = y_true != 0
    if not mask.any():
        return 0.0
    return float(np.mean(np.abs((y_true[mask] - y_pred[mask]) / y_true[mask])) * 100)


def smape(y_true, y_pred) -> float:
    y_true = np.asarray(y_true, dtype=float)
    y_pred = np.asarray(y_pred, dtype=float)
    denominator = (np.abs(y_true) + np.abs(y_pred)) / 2
    mask = denominator != 0
    if not mask.any():
        return 0.0
    return float(np.mean(np.abs(y_true[mask] - y_pred[mask]) / denominator[mask]) * 100)


def regression_metrics(y_true, y_pred) -> dict[str, float]:
    y_pred = np.maximum(np.asarray(y_pred, dtype=float), 0)
    return {
        "mae": mae(y_true, y_pred),
        "rmse": rmse(y_true, y_pred),
        "mape": mape(y_true, y_pred),
        "smape": smape(y_true, y_pred),
    }


def metrics_frame(rows: list[dict]) -> pd.DataFrame:
    columns = ["model", "mae", "rmse", "mape", "smape"]
    return pd.DataFrame(rows).sort_values(["smape", "mae"])[columns].reset_index(drop=True)

