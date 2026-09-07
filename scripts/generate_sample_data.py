"""Restaura los CSV sintéticos reproducibles usados en la demostración.

Advertencia: sobrescribe sales_sample.csv e inventory_sample.csv.
"""

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from demand_forecasting.config import DEFAULT_INVENTORY_CSV, DEFAULT_SALES_CSV
from demand_forecasting.data import write_sample_data


def main() -> None:
    """Genera ventas e inventario con las rutas predeterminadas."""
    sales_path, inventory_path = write_sample_data(DEFAULT_SALES_CSV, DEFAULT_INVENTORY_CSV)
    print(f"Sales sample written to: {sales_path}")
    print(f"Inventory sample written to: {inventory_path}")


if __name__ == "__main__":
    main()
