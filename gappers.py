"""
Stock Gap & Price Action Analysis Tool
----------------------------------------
Fetches historical market data, detects opening gaps above a threshold,
and exports a formatted Excel report with subsequent price action analysis.

Usage:
    python gappers.py --ticker TSLA --start 2026-01-01 --end 2026-05-15
    python gappers.py --ticker AAPL --threshold 0.02
"""

import argparse
import logging
import sys

import yfinance as yf
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Stock gap and price action analyzer.")
    parser.add_argument("--ticker", default="AAPL", help="Ticker symbol, e.g. TSLA, NVDA, BTC-USD")
    parser.add_argument("--start", default="2026-01-01", help="Start date (YYYY-MM-DD)")
    parser.add_argument("--end", default="2026-05-15", help="End date (YYYY-MM-DD)")
    parser.add_argument(
        "--threshold", type=float, default=0.015,
        help="Minimum opening gap to include, as a decimal (0.015 = 1.5%%)"
    )
    parser.add_argument("--output", default=None, help="Output .xlsx filename")
    return parser.parse_args()


def fetch_data(ticker: str, start: str, end: str) -> pd.DataFrame:
    logger.info("Fetching historical data for %s from Yahoo Finance...", ticker)
    try:
        data = yf.download(ticker, start=start, end=end, auto_adjust=False, actions=False)
    except Exception as exc:
        logger.error("Failed to fetch data for '%s': %s", ticker, exc)
        sys.exit(1)

    if isinstance(data.columns, pd.MultiIndex):
        data.columns = data.columns.droplevel(1)

    if data.empty:
        logger.error("No data returned for ticker '%s'. Check the symbol or date range.", ticker)
        sys.exit(1)

    return data


def calculate_gaps(data: pd.DataFrame, threshold: float) -> pd.DataFrame:
    data = data.copy()
    data["Prev_Close"] = data["Close"].shift(1)
    data["Gap_Pct"] = (data["Open"] - data["Prev_Close"]) / data["Prev_Close"]
    data["Day2_Move_Pct"] = (data["Close"].shift(-1) - data["Close"]) / data["Close"]
    data["Day3_Move_Pct"] = (data["Close"].shift(-2) - data["Close"]) / data["Close"]

    gappers = data[data["Gap_Pct"] > threshold].dropna(subset=["Prev_Close"])
    logger.info("Found %d gap events above %.2f%% threshold.", len(gappers), threshold * 100)
    return gappers


def build_report(gappers: pd.DataFrame, ticker: str) -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Gap Analysis Summary"
    ws.views.sheetView[0].showGridLines = True

    font_title = Font(name="Segoe UI", size=16, bold=True, color="004D40")
    font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Segoe UI", size=11, bold=True)

    fill_header = PatternFill(start_color="004D40", end_color="004D40", fill_type="solid")
    fill_zebra = PatternFill(start_color="F0F7F4", end_color="F0F7F4", fill_type="solid")
    fill_green = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
    fill_red = PatternFill(start_color="FFEBEE", end_color="FFEBEE", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"), right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"), bottom=Side(style="thin", color="CCCCCC"),
    )

    ws["A1"] = f"Historical Stock Gap Analysis: {ticker}"
    ws["A1"].font = font_title
    ws.merge_cells("A1:G1")
    ws.row_dimensions[1].height = 30

    headers = ["Date", "Open Price", "Prev Close", "Gap %", "Day 1 Close", "Day 2 Move", "Day 3 Move"]
    for idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    ws.row_dimensions[4].height = 25

    current_row = 5
    for date, row in gappers.iterrows():
        is_zebra = current_row % 2 == 0
        row_fill = fill_zebra if is_zebra else PatternFill(fill_type=None)

        ws.cell(row=current_row, column=1, value=date.strftime("%Y-%m-%d")).alignment = \
            Alignment(horizontal="center", vertical="center")
        ws.cell(row=current_row, column=2, value=float(row["Open"])).number_format = "$#,##0.00"
        ws.cell(row=current_row, column=3, value=float(row["Prev_Close"])).number_format = "$#,##0.00"

        c_gap = ws.cell(row=current_row, column=4, value=float(row["Gap_Pct"]))
        c_gap.number_format = "0.00%"
        c_gap.font = font_bold

        ws.cell(row=current_row, column=5, value=float(row["Close"])).number_format = "$#,##0.00"

        for col_idx, col_name in [(6, "Day2_Move_Pct"), (7, "Day3_Move_Pct")]:
            val = float(row[col_name])
            c_move = ws.cell(row=current_row, column=col_idx, value=val)
            c_move.number_format = "0.00%"
            if val > 0:
                c_move.fill = fill_green
                c_move.font = Font(name="Segoe UI", size=11, color="1B5E20", bold=True)
            else:
                c_move.fill = fill_red
                c_move.font = Font(name="Segoe UI", size=11, color="B71C1C")

        for col in range(1, 8):
            cell = ws.cell(row=current_row, column=col)
            cell.border = thin_border
            if col != 1:
                cell.alignment = Alignment(horizontal="right", vertical="center")
            if col not in (6, 7) and is_zebra:
                cell.fill = row_fill

        ws.row_dimensions[current_row].height = 20
        current_row += 1

    summary_row = current_row + 1
    ws.cell(row=summary_row, column=1, value="Average").font = font_bold
    ws.cell(row=summary_row, column=1).alignment = Alignment(horizontal="left", vertical="center")

    for col_idx, letter in [(4, "D"), (6, "F"), (7, "G")]:
        c_avg = ws.cell(row=summary_row, column=col_idx, value=f"=AVERAGE({letter}5:{letter}{summary_row - 2})")
        c_avg.font = font_bold
        c_avg.number_format = "0.00%"
        c_avg.alignment = Alignment(horizontal="right", vertical="center")
        c_avg.border = Border(top=Side(style="thin", color="000000"), bottom=Side(style="double", color="000000"))

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 13)

    return wb


def main() -> None:
    args = parse_args()
    data = fetch_data(args.ticker, args.start, args.end)
    gappers = calculate_gaps(data, args.threshold)

    if gappers.empty:
        logger.warning("No gap events found for the given parameters. No report generated.")
        return

    wb = build_report(gappers, args.ticker)
    output_filename = args.output or f"{args.ticker}_Gap_Analysis_Project.xlsx"
    wb.save(output_filename)
    logger.info("Success! Report saved as: %s", output_filename)


if __name__ == "__main__":
    main()