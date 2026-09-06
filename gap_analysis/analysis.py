"""Core gap-detection logic: computing gaps, follow-through, and gap-fill status."""
import numpy as np
import pandas as pd

from .config import logger


def calculate_gaps(data: pd.DataFrame, threshold: float, direction: str = "up") -> pd.DataFrame:
    data = data.copy()
    data["Prev_Close"] = data["Close"].shift(1)
    data["Gap_Pct"] = (data["Open"] - data["Prev_Close"]) / data["Prev_Close"]
    data["Day2_Move_Pct"] = (data["Close"].shift(-1) - data["Close"]) / data["Close"]
    data["Day3_Move_Pct"] = (data["Close"].shift(-2) - data["Close"]) / data["Close"]

    # A gap "fills" when price trades back to the previous close on the same day:
    # for a gap-up, that means the day's Low dipped back down to Prev_Close;
    # for a gap-down, it means the day's High climbed back up to Prev_Close.
    gap_up_filled = data["Low"] <= data["Prev_Close"]
    gap_down_filled = data["High"] >= data["Prev_Close"]
    data["Gap_Filled"] = np.where(data["Gap_Pct"] > 0, gap_up_filled, gap_down_filled)

    if direction == "up":
        mask = data["Gap_Pct"] > threshold
    elif direction == "down":
        mask = data["Gap_Pct"] < -threshold
    else:  # both
        mask = data["Gap_Pct"].abs() > threshold

    gappers = data[mask].dropna(subset=["Prev_Close"])
    logger.info(
        "Found %d gap events (direction=%s) above %.2f%% threshold.",
        len(gappers), direction, threshold * 100
    )
    return gappers
