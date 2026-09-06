"""Command-line interface: argument parsing and the main run orchestration."""
import argparse
from pathlib import Path

import pandas as pd

from .cache import fetch_data
from .analysis import calculate_gaps
from .report import build_report, build_summary_sheet
from .config import logger


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Stock gap and price action analyzer.")
    parser.add_argument(
        "--ticker", default="AAPL",
        help="Ticker symbol, e.g. TSLA, NVDA, BTC-USD. Comma-separated for multiple, e.g. AAPL,TSLA,NVDA"
    )
    parser.add_argument("--start", default="2026-01-01", help="Start date (YYYY-MM-DD)")
    parser.add_argument("--end", default="2026-05-15", help="End date (YYYY-MM-DD)")
    parser.add_argument(
        "--threshold", type=float, default=0.015,
        help="Minimum opening gap to include, as a decimal (0.015 = 1.5%%)"
    )
    parser.add_argument(
        "--direction", choices=["up", "down", "both"], default="up",
        help="Which gap direction to detect: 'up' (default), 'down', or 'both'"
    )
    parser.add_argument("--output", default=None, help="Output .xlsx filename")
    parser.add_argument(
        "--no-cache", action="store_true",
        help="Skip local cache and always re-fetch fresh data from Yahoo Finance"
    )
    parser.add_argument(
        "--csv", action="store_true",
        help="Also save the raw gap events (all tickers combined) to a .csv file next to the .xlsx report"
    )
    args = parser.parse_args()

    if args.threshold <= 0:
        parser.error("--threshold must be a positive number (e.g. 0.015 for 1.5%)")

    try:
        start_date = pd.to_datetime(args.start)
        end_date = pd.to_datetime(args.end)
    except ValueError as exc:
        parser.error(f"Invalid date format, expected YYYY-MM-DD: {exc}")
    else:
        if start_date >= end_date:
            parser.error(f"--start ({args.start}) must be before --end ({args.end})")

    return args


def main() -> None:
    args = parse_args()
    tickers = [t.strip().upper() for t in args.ticker.split(",") if t.strip()]

    wb = None
    tickers_with_data = []
    stats = []
    csv_frames = []
    for ticker in tickers:
        data = fetch_data(ticker, args.start, args.end, use_cache=not args.no_cache)
        if data is None:
            continue
        gappers = calculate_gaps(data, args.threshold, args.direction)

        if gappers.empty:
            logger.warning("No gap events found for '%s'. Skipping.", ticker)
            continue

        wb = build_report(gappers, ticker, wb)
        tickers_with_data.append(ticker)

        fill_rate = float(gappers["Gap_Filled"].mean())
        stats.append({
            "ticker": ticker,
            "count": len(gappers),
            "avg_gap": float(gappers["Gap_Pct"].mean()),
            "fill_rate": fill_rate,
            "avg_day2": float(gappers["Day2_Move_Pct"].dropna().mean()) if gappers["Day2_Move_Pct"].notna().any() else 0.0,
            "avg_day3": float(gappers["Day3_Move_Pct"].dropna().mean()) if gappers["Day3_Move_Pct"].notna().any() else 0.0,
        })

        if args.csv:
            labeled = gappers.copy()
            labeled.insert(0, "Ticker", ticker)
            csv_frames.append(labeled)

    if wb is None:
        logger.warning("No gap events found for any ticker. No report generated.")
        return

    if len(tickers_with_data) > 1:
        build_summary_sheet(wb, stats)

    if args.output:
        output_filename = args.output
    elif len(tickers_with_data) == 1:
        output_filename = f"{tickers_with_data[0]}_Gap_Analysis_Project.xlsx"
    else:
        output_filename = "Multi_Ticker_Gap_Analysis_Project.xlsx"

    wb.save(output_filename)
    logger.info("Success! Report saved as: %s (%d ticker sheet(s))", output_filename, len(tickers_with_data))

    if args.csv and csv_frames:
        csv_filename = str(Path(output_filename).with_suffix(".csv"))
        pd.concat(csv_frames).to_csv(csv_filename)
        logger.info("Raw gap events also saved as: %s", csv_filename)

    print("\n=== Gap Analysis Summary ===")
    for s in stats:
        print(
            f"{s['ticker']:<8} {s['count']:>3} gaps | "
            f"avg gap {s['avg_gap']:+.2%} | fill rate {s['fill_rate']:.1%} | "
            f"avg Day2 {s['avg_day2']:+.2%} | avg Day3 {s['avg_day3']:+.2%}"
        )
