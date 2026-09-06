"""Shared constants and logging configuration for the gap analysis tool."""
import logging
from pathlib import Path

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("gap_analysis")

# --- Data fetching / caching ---
CACHE_DIR = Path(".cache")
CACHE_MAX_AGE_SECONDS = 24 * 60 * 60  # 1 day
MAX_FETCH_RETRIES = 3
RETRY_BACKOFF_SECONDS = 2

# --- Report / chart theme colors ---
ACCENT_TEAL = "004D40"
ACCENT_GREEN = "2E7D32"
ACCENT_RED = "C62828"
GRID_GRAY = "E0E0E0"
