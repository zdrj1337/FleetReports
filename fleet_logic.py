"""
Fleet payment logic.

The two ride-hailing platforms export driver data in different shapes:
Bolt and Uber use different column names and a different structure. This module
reads each export and *normalises* both into one common table, so the rest of
the program can work with a single, consistent format.

Common schema (one row per driver per platform):
    driver, platform, gross, platform_commission, cash, tips, bonus
"""

import os
import pandas as pd

COMMON_COLS = ["driver", "platform", "gross", "platform_commission", "cash", "tips", "bonus"]

# The fleet's own commission, charged on top of the platform's fee.
FIRM_RATE = 0.07

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def resolve_data_dir(period: str, base: str | None = None) -> str:
    """
    Return the data directory for *period* (e.g. '2026-W10').

    Checks the new hierarchical layout (data/{year}/{week}/) first, then
    falls back to the flat legacy layout (data/{period}/) so old folders
    still work.
    """
    if base is None:
        base = _BASE_DIR
    year, week = period.split("-")
    hierarchical = os.path.join(base, "data", year, week)
    if os.path.isdir(hierarchical):
        return hierarchical
    return os.path.join(base, "data", period)


def load_bolt(path: str) -> pd.DataFrame:
    """Read a Bolt fleet export and map it to the common schema."""
    df = pd.read_csv(path)
    out = pd.DataFrame({
        "driver":               df["Driver"],
        "platform":             "Bolt",
        "gross":                df["Gross earnings"],
        "platform_commission":  df["Bolt commission"],
        "cash":                 df["Cash gross earnings"],
        "tips":                 df["Tips"],
        "bonus":                df["Bonus"],
    })
    return out[COMMON_COLS]


def load_uber(path: str) -> pd.DataFrame:
    """Read an Uber fleet export (different columns) and map it to the same schema."""
    df = pd.read_csv(path)
    out = pd.DataFrame({
        "driver":               df["Driver"],
        "platform":             "Uber",
        "gross":                df["Trip earnings"],
        "platform_commission":  df["Uber service fee"],
        "cash":                 df["Cash collected"],
        "tips":                 df["Tips"],
        "bonus":                df["Promotions"],   # Uber calls bonuses "Promotions"
    })
    return out[COMMON_COLS]


def load_period(data_dir: str) -> pd.DataFrame:
    """
    Load both platform files for one period and return a single long table,
    one row per driver per platform.
    """
    uber = load_uber(os.path.join(data_dir, "uber.csv"))
    bolt = load_bolt(os.path.join(data_dir, "bolt.csv"))
    combined = pd.concat([uber, bolt], ignore_index=True)
    return combined.sort_values(["driver", "platform"]).reset_index(drop=True)


def drivers_in(combined: pd.DataFrame) -> list[str]:
    """Return the sorted list of unique driver names in the combined table."""
    return sorted(combined["driver"].unique())


def compute_summary(combined: pd.DataFrame) -> pd.DataFrame:
    """
    Roll the long table up to one row per driver and compute the payment figures.

    This is the same math the Excel reports encode as formulas, done in Python so
    the web page can show the numbers without opening the generated files. Columns:
        driver, gross, platform_commission, cash, tips, bonus,
        firm_commission, net, bank_transfer
    """
    g = (combined.groupby("driver", as_index=False)[
            ["gross", "platform_commission", "cash", "tips", "bonus"]]
         .sum()
         .sort_values("driver")
         .reset_index(drop=True))
    g["firm_commission"] = (g["gross"] * FIRM_RATE).round(2)
    g["net"] = (g["gross"] - g["platform_commission"] - g["firm_commission"]).round(2)
    g["bank_transfer"] = (g["net"] + g["tips"] + g["bonus"] - g["cash"]).round(2)
    return g
