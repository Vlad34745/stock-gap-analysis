"""
Stock Gap & Price Action Analysis Tool — entry point
-----------------------------------------------------
The implementation lives in the gap_analysis/ package (data fetching +
caching, gap-detection logic, chart styling, Excel report building, and
the CLI), split into focused modules instead of one large file. This
script is kept as the entry point so `python gappers.py ...` and
run_gappers.bat keep working exactly as before.

By default only gap-ups (Open > previous Close) are detected. Use
--direction to also include gap-downs or both.

Multiple tickers can be analyzed in one run (comma-separated); each
gets its own sheet in a single output workbook, plus a Summary
comparison sheet.

Usage:
    python gappers.py --ticker TSLA --start 2026-01-01 --end 2026-05-15
    python gappers.py --ticker AAPL --threshold 0.02
    python gappers.py --ticker AAPL --direction both
    python gappers.py --ticker AAPL,TSLA,NVDA --output multi_report.xlsx
    python gappers.py --ticker AAPL --csv
"""
from gap_analysis.cli import main

if __name__ == "__main__":
    main()
