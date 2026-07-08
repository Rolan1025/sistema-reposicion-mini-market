from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from demand_forecasting.config import DEFAULT_MODEL_PATH, DEFAULT_SALES_CSV, REPORTS_DIR
from demand_forecasting.pipeline import train_pipeline


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Train demand forecasting models.")
    parser.add_argument("--sales-csv", default=str(DEFAULT_SALES_CSV))
    parser.add_argument("--model-path", default=str(DEFAULT_MODEL_PATH))
    parser.add_argument("--metrics-path", default=str(REPORTS_DIR / "metrics.csv"))
    parser.add_argument("--test-days", type=int, default=30)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    model_path, metrics = train_pipeline(
        sales_csv=args.sales_csv,
        model_path=args.model_path,
        metrics_path=args.metrics_path,
        test_days=args.test_days,
    )
    print(f"Model saved to: {model_path}")
    print(metrics.to_string(index=False))


if __name__ == "__main__":
    main()
