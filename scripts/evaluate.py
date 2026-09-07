"""Comando que genera la evaluación detallada de 7, 14 y 30 días."""

from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from demand_forecasting.config import DEFAULT_SALES_CSV, REPORTS_DIR
from demand_forecasting.evaluation import evaluate_sales_csv


def main() -> None:
    """Ejecuta el backtesting y guarda sus CSV dentro de reports/."""
    outputs = evaluate_sales_csv(DEFAULT_SALES_CSV, REPORTS_DIR)
    print("Evaluación multihorizonte completada:")
    for name, path in outputs.items():
        print(f"- {name}: {path}")


if __name__ == "__main__":
    main()
