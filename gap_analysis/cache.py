"""Fetching historical price data from Yahoo Finance, with local caching and retries."""
import hashlib
import time
from pathlib import Path
from typing import Optional

import pandas as pd
import yfinance as yf

from . import config
from .config import logger


def _cache_path(ticker: str, start: str, end: str) -> Path:
    key = hashlib.md5(f"{ticker}_{start}_{end}".encode()).hexdigest()[:12]
    return config.CACHE_DIR / f"{ticker}_{key}.csv"


def _load_from_cache(path: Path) -> Optional[pd.DataFrame]:
    if not path.exists():
        return None
    age = time.time() - path.stat().st_mtime
    if age > config.CACHE_MAX_AGE_SECONDS:
        return None
    try:
        data = pd.read_csv(path, index_col=0, parse_dates=True)
    except Exception:
        return None
    return data if not data.empty else None


def fetch_data(ticker: str, start: str, end: str, use_cache: bool = True) -> Optional[pd.DataFrame]:
    cache_path = _cache_path(ticker, start, end)

    if use_cache:
        cached = _load_from_cache(cache_path)
        if cached is not None:
            logger.info("Using cached data for %s (from %s).", ticker, cache_path.name)
            return cached

    data = None
    for attempt in range(1, config.MAX_FETCH_RETRIES + 1):
        logger.info(
            "Fetching historical data for %s from Yahoo Finance (attempt %d/%d)...",
            ticker, attempt, config.MAX_FETCH_RETRIES
        )
        try:
            data = yf.download(ticker, start=start, end=end, auto_adjust=False, actions=False)
        except Exception as exc:
            logger.warning("Attempt %d failed for '%s': %s", attempt, ticker, exc)
            data = None

        if data is not None and isinstance(data.columns, pd.MultiIndex):
            data.columns = data.columns.droplevel(1)

        if data is not None and not data.empty:
            break

        if attempt < config.MAX_FETCH_RETRIES:
            wait = config.RETRY_BACKOFF_SECONDS * attempt
            logger.info("Retrying in %ds...", wait)
            time.sleep(wait)

    if data is None or data.empty:
        logger.error("No data returned for ticker '%s' after %d attempts. Check the symbol or date range.",
                     ticker, config.MAX_FETCH_RETRIES)
        return None

    if use_cache:
        try:
            config.CACHE_DIR.mkdir(exist_ok=True)
            data.to_csv(cache_path)
        except Exception as exc:
            logger.warning("Could not write cache for '%s': %s", ticker, exc)

    return data
