"""
Stock Gap & Price Action Analysis Tool
----------------------------------------
Implementation package. See the root gappers.py for the CLI entry point
and README.md for usage.

Modules:
    config       - shared constants and logging setup
    cache        - fetch_data() with local caching + retry-with-backoff
    analysis     - calculate_gaps() core gap-detection logic
    chart_style  - shared chart styling helpers (theme colors, fonts)
    report       - build_report() / build_summary_sheet() Excel output
    cli          - parse_args() / main() orchestration

The most commonly used pieces are re-exported here for convenience:
    from gap_analysis import calculate_gaps, fetch_data, build_report, build_summary_sheet
"""
from .analysis import calculate_gaps
from .cache import fetch_data, _cache_path, _load_from_cache
from .report import build_report, build_summary_sheet
from .cli import parse_args, main

__all__ = [
    "calculate_gaps", "fetch_data", "_cache_path", "_load_from_cache",
    "build_report", "build_summary_sheet", "parse_args", "main",
]
