import os
from pathlib import Path

BASE_DIR = Path(__file__).parent

INPUT_PHOTOS_DIR = BASE_DIR / "input_photos"
PROCESSED_PHOTOS_DIR = BASE_DIR / "processed_photos"
FAILED_PHOTOS_DIR = BASE_DIR / "failed_photos"
OUTPUT_DIR = BASE_DIR / "output"
DB_PATH = BASE_DIR / "db.sqlite"

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")

IDENTIFY_MODEL = "claude-opus-4-7"
LISTING_MODEL = "claude-sonnet-4-6"

BATCH_SIZE = 25
CONFIDENCE_THRESHOLD = 50
COMP_DELAY_SECONDS = 2.5
SKIP_COMPS_UNDER_VALUE = 15
MAX_RETRIES = 3

SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".gif"}
