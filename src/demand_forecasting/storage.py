from __future__ import annotations

import sqlite3
from pathlib import Path

import pandas as pd


def write_report_tables(
    db_path: str | Path,
    forecast: pd.DataFrame,
    recommendations: pd.DataFrame,
    metrics: pd.DataFrame | None = None,
) -> Path:
    db_path = Path(db_path)
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as conn:
        forecast.to_sql("forecast", conn, if_exists="replace", index=False)
        recommendations.to_sql("recommendations", conn, if_exists="replace", index=False)
        if metrics is not None:
            metrics.to_sql("model_metrics", conn, if_exists="replace", index=False)
    return db_path

