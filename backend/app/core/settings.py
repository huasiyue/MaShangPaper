from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[3]
BACKEND_DIR = PROJECT_ROOT / "backend"
TEMP_DIR = BACKEND_DIR / "temp"
ASSETS_DIR = TEMP_DIR / "assets"
DOCS_DATA_DIR = PROJECT_ROOT / "docs" / "data"

DEFAULT_SCHOOL_ID = "sdfmu"
DEFAULT_THESIS_TYPE = "thesis"
SUPPORTED_SCHOOLS = ("sdfmu", "yzu", "sdfmu_ai")

TEMP_DIR.mkdir(parents=True, exist_ok=True)
ASSETS_DIR.mkdir(parents=True, exist_ok=True)
