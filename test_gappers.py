"""
Unit tests for the gap_analysis package (entry point: gappers.py).

Run with:
    pip install pytest
    pytest test_gappers.py -v
"""
import numpy as np
import pandas as pd
import pytest

from gap_analysis import calculate_gaps, build_report, build_summary_sheet, _cache_path, _load_from_cache


@pytest.fixture
def sample_data() -> pd.DataFrame:
    """
    5 trading days of synthetic OHLC data with one clear gap-up
    (day 3) and one clear gap-down (day 5).
    """
    dates = pd.date_range("2026-01-01", periods=5, freq="D")
    return pd.DataFrame(
        {
            "Open":  [100, 101, 105, 104, 96],
            "Close": [100, 100, 104, 100, 97],
            "High":  [101, 102, 106, 105, 98],
            "Low":   [99, 100, 104, 99, 95],
        },
        index=dates,
    )


def test_calculate_gaps_detects_gap_up(sample_data):
    result = calculate_gaps(sample_data, threshold=0.02, direction="up")
    # Day 3: Open=105 vs Prev_Close=100 -> +5% gap, above 2% threshold
    assert len(result) == 1
    assert result.iloc[0]["Gap_Pct"] == pytest.approx(0.05)


def test_calculate_gaps_detects_gap_down(sample_data):
    result = calculate_gaps(sample_data, threshold=0.02, direction="down")
    # Day 5: Open=96 vs Prev_Close=100 -> -4% gap
    assert len(result) == 1
    assert result.iloc[0]["Gap_Pct"] == pytest.approx(-0.04)


def test_calculate_gaps_both_directions(sample_data):
    result = calculate_gaps(sample_data, threshold=0.02, direction="both")
    assert len(result) == 2


def test_calculate_gaps_respects_threshold(sample_data):
    # With a very high threshold, nothing should qualify
    result = calculate_gaps(sample_data, threshold=0.5, direction="both")
    assert result.empty


def test_calculate_gaps_first_row_excluded(sample_data):
    # First row has no Prev_Close, must never appear in results
    result = calculate_gaps(sample_data, threshold=0.0, direction="both")
    assert sample_data.index[0] not in result.index


def test_calculate_gaps_up_only_ignores_down_moves(sample_data):
    # direction="up" must not include the gap-down day even though
    # its magnitude exceeds the threshold
    result = calculate_gaps(sample_data, threshold=0.02, direction="up")
    assert all(result["Gap_Pct"] > 0)


def test_build_report_handles_trailing_nan_rows(sample_data):
    """
    The last row(s) of a gap event list can have NaN Day2/Day3 moves
    because there's no future price data. build_report must not crash
    and must not write raw NaN into the workbook.
    """
    gappers = calculate_gaps(sample_data, threshold=0.02, direction="both")
    wb = build_report(gappers, ticker="TEST")
    ws = wb.active

    # Collect all cell values in the Day2/Day3 columns (G, H) below the header
    for row in ws.iter_rows(min_row=5, max_col=8, min_col=7):
        for cell in row:
            if cell.value is not None and not isinstance(cell.value, str):
                assert not (isinstance(cell.value, float) and np.isnan(cell.value)), (
                    "Raw NaN was written to the workbook instead of being handled"
                )


def test_calculate_gaps_detects_fill(sample_data):
    """
    Day 3 gap-up: Open=105, Prev_Close=100, Low=104 -> never traded back
    down to 100, so this gap should NOT be marked as filled.
    """
    result = calculate_gaps(sample_data, threshold=0.02, direction="up")
    assert bool(result.iloc[0]["Gap_Filled"]) is False


def test_calculate_gaps_detects_fill_when_price_returns():
    """A gap-up where price dips back to (or below) Prev_Close counts as filled."""
    dates = pd.date_range("2026-01-01", periods=2, freq="D")
    data = pd.DataFrame(
        {"Open": [100, 110], "Close": [100, 105], "High": [101, 112], "Low": [99, 99]},
        index=dates,
    )
    result = calculate_gaps(data, threshold=0.02, direction="up")
    assert len(result) == 1
    assert bool(result.iloc[0]["Gap_Filled"]) is True


def test_build_report_empty_input_does_not_crash():
    empty = pd.DataFrame(columns=["Open", "Close", "High", "Low", "Prev_Close",
                                   "Gap_Pct", "Gap_Filled", "Day2_Move_Pct", "Day3_Move_Pct"])
    wb = build_report(empty, ticker="EMPTY")
    assert wb.active["A1"].value == "Historical Stock Gap Analysis: EMPTY"


def test_build_report_multi_ticker_creates_separate_sheets(sample_data):
    """Passing an existing workbook should append a new sheet, not overwrite."""
    gappers = calculate_gaps(sample_data, threshold=0.02, direction="both")

    wb = build_report(gappers, ticker="AAPL")
    wb = build_report(gappers, ticker="TSLA", wb=wb)

    assert wb.sheetnames == ["AAPL", "TSLA"]
    assert wb["AAPL"]["A1"].value == "Historical Stock Gap Analysis: AAPL"
    assert wb["TSLA"]["A1"].value == "Historical Stock Gap Analysis: TSLA"


def test_build_report_sanitizes_unsafe_sheet_name(sample_data):
    gappers = calculate_gaps(sample_data, threshold=0.02, direction="both")
    wb = build_report(gappers, ticker="BRK/B")  # '/' is illegal in sheet names
    assert wb.active.title == "BRKB"


def test_build_report_filled_column_matches_gap_filled_flag(sample_data):
    gappers = calculate_gaps(sample_data, threshold=0.02, direction="both")
    wb = build_report(gappers, ticker="TEST")
    ws = wb.active

    for i, (_, row) in enumerate(gappers.iterrows()):
        cell_value = ws.cell(row=5 + i, column=5).value  # column E = "Filled?"
        expected = "Yes" if bool(row["Gap_Filled"]) else "No"
        assert cell_value == expected


def test_build_report_adds_chart_when_data_present(sample_data):
    gappers = calculate_gaps(sample_data, threshold=0.02, direction="both")
    wb = build_report(gappers, ticker="TEST")
    ws = wb.active
    assert len(ws._charts) == 1


def test_build_report_no_chart_when_empty():
    empty = pd.DataFrame(columns=["Open", "Close", "High", "Low", "Prev_Close",
                                   "Gap_Pct", "Gap_Filled", "Day2_Move_Pct", "Day3_Move_Pct"])
    wb = build_report(empty, ticker="EMPTY")
    assert len(wb.active._charts) == 0


def test_build_summary_sheet_is_first_and_has_correct_values(sample_data):
    gappers = calculate_gaps(sample_data, threshold=0.02, direction="both")
    wb = build_report(gappers, ticker="AAPL")
    wb = build_report(gappers, ticker="TSLA", wb=wb)

    stats = [
        {"ticker": "AAPL", "count": 2, "avg_gap": 0.01, "fill_rate": 0.5, "avg_day2": 0.02, "avg_day3": 0.03},
        {"ticker": "TSLA", "count": 1, "avg_gap": -0.01, "fill_rate": 1.0, "avg_day2": 0.0, "avg_day3": 0.0},
    ]
    build_summary_sheet(wb, stats)

    assert wb.sheetnames[0] == "Summary"
    ws = wb["Summary"]
    assert ws.cell(row=4, column=1).value == "AAPL"
    assert ws.cell(row=4, column=2).value == 2
    assert ws.cell(row=5, column=1).value == "TSLA"
    assert len(ws._charts) == 1


def test_cache_roundtrip(tmp_path, monkeypatch, sample_data):
    from gap_analysis import config as config_module
    monkeypatch.setattr(config_module, "CACHE_DIR", tmp_path)

    path = _cache_path("AAPL", "2026-01-01", "2026-01-10")
    sample_data.to_csv(path)

    cached = _load_from_cache(path)
    assert cached is not None
    assert list(cached.columns) == list(sample_data.columns)
    assert len(cached) == len(sample_data)


def test_cache_ignored_when_missing():
    path = _cache_path("NON_EXISTENT_TICKER_XYZ", "2026-01-01", "2026-01-10")
    assert _load_from_cache(path) is None