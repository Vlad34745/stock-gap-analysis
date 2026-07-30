## 📋 Usage

**Windows:** double-click `run_gappers.bat` (creates a venv, installs dependencies, and runs with defaults), or pass your own arguments:
```bat
run_gappers.bat --ticker TSLA --start 2026-01-01 --end 2026-06-01
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

Analyze several tickers in one run (each gets its own sheet in the same file):
```bash
python gappers.py --ticker AAPL,TSLA,NVDA --output multi_report.xlsx
```

**Arguments:**
- `--ticker` — stock symbol, or comma-separated list for multiple (default: AAPL)
- `--start` / `--end` — date range (YYYY-MM-DD)
- `--threshold` — minimum gap % to include (default: 0.015 = 1.5%)
- `--direction` — which gaps to detect: `up` (default), `down`, or `both`
- `--output` — custom output filename

## 🧪 Testing

```bash
pip install -r requirements-dev.txt
pytest test_gappers.py -v
```