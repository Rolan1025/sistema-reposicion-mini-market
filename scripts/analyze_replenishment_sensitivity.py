from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from demand_forecasting.data import load_inventory_csv
from demand_forecasting.replenishment import recommend_replenishment


def main() -> None:
    forecast = pd.read_csv(ROOT / "reports" / "forecast.csv", parse_dates=["date"])
    inventory = load_inventory_csv(ROOT / "data" / "raw" / "inventory_sample.csv")
    rows = []
    for safety_days in (1, 3, 5):
        for review_days in (1, 3):
            recommendations = recommend_replenishment(
                forecast,
                inventory,
                safety_stock_days=safety_days,
                review_period_days=review_days,
            )
            rows.append(
                {
                    "safety_stock_days": safety_days,
                    "review_period_days": review_days,
                    "products_with_order": int((recommendations["suggested_order_qty"] > 0).sum()),
                    "high_priority_products": int((recommendations["priority"] == "alta").sum()),
                    "total_suggested_units": int(recommendations["suggested_order_qty"].sum()),
                    "mean_coverage_days": round(float(recommendations["coverage_days"].mean()), 2),
                }
            )
    output = ROOT / "reports" / "replenishment_sensitivity.csv"
    pd.DataFrame(rows).to_csv(output, index=False)
    print(output)


if __name__ == "__main__":
    main()
