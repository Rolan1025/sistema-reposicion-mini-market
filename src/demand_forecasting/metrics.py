"""Métricas que comparan ventas reales con ventas pronosticadas.

Todas representan error: cuanto menor sea el valor, mejor se ajusta el modelo.
MAE y RMSE se expresan en unidades; MAPE y SMAPE se expresan en porcentaje.
"""

import numpy as np
import pandas as pd


def mae(y_true, y_pred) -> float:
    """Error absoluto medio: unidades que el modelo falla, en promedio.

    Ejemplo: MAE=6 significa que una predicción diaria se aleja unas seis
    unidades del valor real, sin importar si predijo de más o de menos.
    """
    return float(np.mean(np.abs(np.asarray(y_true) - np.asarray(y_pred))))


def rmse(y_true, y_pred) -> float:
    """Raíz del error cuadrático medio: penaliza más los errores grandes.

    Dos modelos pueden tener MAE parecido, pero el que comete algunos fallos
    muy grandes normalmente tendrá un RMSE mayor.
    """
    return float(np.sqrt(np.mean((np.asarray(y_true) - np.asarray(y_pred)) ** 2)))


def mape(y_true, y_pred) -> float:
    """Error porcentual absoluto medio respecto al valor real.

    Se omiten ventas reales iguales a cero porque no se puede dividir por cero.
    Por esa razón MAPE puede ser inestable en productos de muy baja demanda.
    """
    y_true = np.asarray(y_true, dtype=float)
    y_pred = np.asarray(y_pred, dtype=float)
    mask = y_true != 0
    if not mask.any():
        return 0.0
    return float(np.mean(np.abs((y_true[mask] - y_pred[mask]) / y_true[mask])) * 100)


def smape(y_true, y_pred) -> float:
    """Error porcentual simétrico entre el valor real y el pronosticado.

    Usa ambos valores en el denominador, por lo que suele ser más equilibrado
    que MAPE cuando se comparan productos con escalas diferentes.
    """
    y_true = np.asarray(y_true, dtype=float)
    y_pred = np.asarray(y_pred, dtype=float)
    denominator = (np.abs(y_true) + np.abs(y_pred)) / 2
    mask = denominator != 0
    if not mask.any():
        return 0.0
    return float(np.mean(np.abs(y_true[mask] - y_pred[mask]) / denominator[mask]) * 100)


def regression_metrics(y_true, y_pred) -> dict[str, float]:
    """Calcula las cuatro métricas y evita predicciones negativas."""
    y_pred = np.maximum(np.asarray(y_pred, dtype=float), 0)
    return {
        "mae": mae(y_true, y_pred),
        "rmse": rmse(y_true, y_pred),
        "mape": mape(y_true, y_pred),
        "smape": smape(y_true, y_pred),
    }


def metrics_frame(rows: list[dict]) -> pd.DataFrame:
    """Ordena los modelos: primero menor SMAPE y luego menor MAE."""
    columns = ["model", "mae", "rmse", "mape", "smape"]
    return pd.DataFrame(rows).sort_values(["smape", "mae"])[columns].reset_index(drop=True)
