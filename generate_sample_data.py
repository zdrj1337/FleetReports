"""
generate_sample_data.py — synthetic Bolt / Uber weekly exports for testing.

Creates  data/{year}/{week}/bolt.csv  and  data/{year}/{week}/uber.csv
for every period in DEFAULT_PERIODS (2025-W40 → 2026-W20).

Usage
-----
    python generate_sample_data.py              # generates all default periods
    python generate_sample_data.py 2026-W05     # (re)generates a single period
"""

import csv
import os
import random
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# name → (bolt_weekly_base_RON, uber_weekly_base_RON)   0 = not on that platform
DRIVERS: dict[str, tuple[float, float]] = {
    "Andrei Popescu":    (630, 480),
    "Ioana Munteanu":    (410, 355),
    "Vlad Georgescu":    (520,   0),
    "Cristina Dinu":     (285, 215),
    "Mihai Stoica":      (  0, 505),
    "Alexandru Ionescu": (555, 460),
    "Elena Radu":        (375, 325),
    "Bogdan Constantin": (  0, 435),
    "Raluca Popa":       (450,   0),
    "Dragos Nistor":     (490, 415),
    "Madalina Florea":   (  0, 295),
    "Cosmin Barbu":      (535, 470),
    "Teodora Luca":      (315,   0),
    "Gabriel Stan":      (605, 525),
    "Diana Matei":       (  0, 380),
}

DEFAULT_PERIODS = (
    [f"2025-W{w:02d}" for w in range(40, 53)] +
    [f"2026-W{w:02d}" for w in range(1,  21)]
)

BOLT_HEADER = [
    "Driver", "Gross earnings", "In-app gross earnings",
    "Cash gross earnings", "Bolt commission", "Tips", "Bonus",
]
UBER_HEADER = [
    "Driver", "Trip earnings", "Uber service fee",
    "Promotions", "Tips", "Cash collected",
]


def _seasonal(year: int, week: int) -> float:
    """Gross multiplier based on season."""
    if year == 2025 and week >= 49:
        return 1.18          # Christmas rush
    if year == 2026 and week <= 2:
        return 0.88          # post-holiday slump
    if year == 2026 and 14 <= week <= 16:
        return 1.10          # Easter / spring spike
    return 1.0


def _vary(rng: random.Random, base: float, pct: float = 0.14) -> float:
    return round(base * rng.uniform(1 - pct, 1 + pct), 2)


def _bolt_row(rng: random.Random, name: str, base: float, year: int, week: int) -> list:
    gross      = _vary(rng, base * _seasonal(year, week))
    commission = round(gross * 0.25, 2)
    cash       = round(gross * rng.uniform(0.28, 0.52), 2)
    in_app     = round(gross - cash, 2)
    tips       = round(max(0, _vary(rng, gross * 0.023, pct=0.6)), 2)
    bonus      = round(rng.choice([0, 0, 0, 0, 10, 15, 20, 25, 30]), 2)
    return [name, gross, in_app, cash, commission, tips, bonus]


def _uber_row(rng: random.Random, name: str, base: float, year: int, week: int) -> list:
    gross = _vary(rng, base * _seasonal(year, week))
    fee   = round(gross * 0.25, 2)
    cash  = round(gross * rng.uniform(0.22, 0.48), 2)
    tips  = round(max(0, _vary(rng, gross * 0.021, pct=0.6)), 2)
    bonus = round(rng.choice([0, 0, 0, 0, 15, 20, 25]), 2)
    return [name, gross, fee, bonus, tips, cash]


def write_period(period: str) -> int:
    """Write bolt.csv + uber.csv for one period. Returns number of active drivers."""
    year_str, week_str = period.split("-")
    year = int(year_str)
    week = int(week_str[1:])

    # Per-period deterministic seed → stable data even when regenerating one period
    rng = random.Random(year * 100 + week)

    out_dir = os.path.join(BASE_DIR, "data", year_str, week_str)
    os.makedirs(out_dir, exist_ok=True)

    bolt_rows: list[list] = []
    uber_rows: list[list] = []

    for name, (bolt_base, uber_base) in DRIVERS.items():
        if rng.random() < 0.10:   # ~10 % chance driver skips this week
            continue
        if bolt_base > 0:
            bolt_rows.append(_bolt_row(rng, name, bolt_base, year, week))
        if uber_base > 0:
            uber_rows.append(_uber_row(rng, name, uber_base, year, week))

    bolt_rows.sort(key=lambda r: r[0])
    uber_rows.sort(key=lambda r: r[0])

    with open(os.path.join(out_dir, "bolt.csv"), "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows([BOLT_HEADER] + bolt_rows)

    with open(os.path.join(out_dir, "uber.csv"), "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows([UBER_HEADER] + uber_rows)

    active = len({r[0] for r in bolt_rows} | {r[0] for r in uber_rows})
    return active


def main() -> None:
    periods = sys.argv[1:] if len(sys.argv) > 1 else DEFAULT_PERIODS
    print(f"Generating {len(periods)} period(s)...")
    for p in periods:
        n = write_period(p)
        print(f"  {p}  ->  {n:2d} drivers")
    print("Done.")


if __name__ == "__main__":
    main()
