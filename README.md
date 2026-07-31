## 📋 Usage

> Requires **Python 3.11+** (pandas 3.x drops support for older versions).

**Windows:** double-click `run_gappers.bat` — it sets up the venv, installs dependencies, and then asks you interactively for ticker(s), dates, threshold, and direction (type several tickers separated by spaces or commas, e.g. `AAPL TSLA NVDA`).

You can also skip the prompts by passing arguments directly:
```bat
run_gappers.bat --ticker TSLA --start 2026-01-01 --end 2026-06-01
run_gappers.bat --ticker AAPL,TSLA,NVDA
```

**Manually / any OS:**
Run with default settings (AAPL, Jan–May 2026):
```bash
python gappers.py
```

Or specify your own parameters:
```bash
python gappers.py --ticker TSLA --start 2026-01-01 --end 2026-06-01 --threshold 0.02
```

Analyze several tickers in one run (each gets its own sheet, plus a Summary comparison sheet):
```bash
python gappers.py --ticker AAPL,TSLA,NVDA --output multi_report.xlsx
```

**Arguments:**
- `--ticker` — stock symbol, or comma-separated list for multiple (default: AAPL)
- `--start` / `--end` — date range (YYYY-MM-DD)
- `--threshold` — minimum gap % to include (default: 0.015 = 1.5%)
- `--direction` — which gaps to detect: `up` (default), `down`, or `both`
- `--output` — custom output filename
- `--no-cache` — skip the local cache and always fetch fresh data

## 💾 Caching & retries

Downloaded price data is cached locally in `.cache/` for 24 hours, so re-running
the same ticker/date range doesn't hit Yahoo Finance again. Failed downloads are
retried automatically (up to 3 attempts with backoff) before giving up on a ticker.

## 🧪 Testing

```bash
pip install -r requirements-dev.txt
pytest test_gappers.py -v
```

Tests run automatically on every push via GitHub Actions (`.github/workflows/tests.yml`).