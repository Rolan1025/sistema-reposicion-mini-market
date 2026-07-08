from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DATA_DIR = PROJECT_ROOT / "data"
RAW_DATA_DIR = DATA_DIR / "raw"
PROCESSED_DATA_DIR = DATA_DIR / "processed"
MODELS_DIR = PROJECT_ROOT / "models"
REPORTS_DIR = PROJECT_ROOT / "reports"

DEFAULT_SALES_CSV = RAW_DATA_DIR / "sales_sample.csv"
DEFAULT_INVENTORY_CSV = RAW_DATA_DIR / "inventory_sample.csv"
DEFAULT_MODEL_PATH = MODELS_DIR / "demand_model.joblib"

DEFAULT_HORIZONS = [7, 14, 30]
RANDOM_STATE = 42

