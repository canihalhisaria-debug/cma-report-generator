from __future__ import annotations

import hashlib
from typing import Iterable


YEARS = ["FY1", "FY2", "FY3", "FY4", "FY5"]


def safe_div(num: float, den: float) -> float:
    return num / den if den else 0.0


def pct(value: float) -> float:
    return value / 100.0


def growth_series(base: float, growth_rate: float, years: int = 5) -> list[float]:
    series = []
    current = base
    for _ in range(years):
        series.append(current)
        current *= 1 + growth_rate
    return series


def rolling_opening_closing(closing_values: Iterable[float]) -> tuple[list[float], list[float]]:
    closings = list(closing_values)
    openings = [0.0] + closings[:-1]
    return openings, closings


def profile_seed(*parts: str) -> int:
    key = "|".join((p or "").strip().lower() for p in parts)
    digest = hashlib.sha256(key.encode("utf-8")).hexdigest()
    return int(digest[:8], 16)


def bounded_variation(seed: int, low: float, high: float) -> float:
    span = high - low
    normalized = (seed % 10000) / 10000
    return low + (span * normalized)


def benchmark_status(name: str, value: float) -> tuple[str, str]:
    rules = {
        "Current Ratio": (1.33, 2.5),
        "Quick Ratio": (1.0, 2.0),
        "Debt Equity Ratio": (0.0, 2.0),
        "Total Indebtedness Ratio": (0.0, 4.0),
        "Debt Asset Ratio": (0.0, 0.7),
        "Gross Profit Ratio": (0.15, 0.45),
        "Operating Profit Ratio": (0.08, 0.25),
        "Net Profit Ratio": (0.05, 0.2),
        "Interest Coverage Ratio": (1.5, 99.0),
        "DSCR": (1.25, 99.0),
        "Inventory Turnover": (4.0, 12.0),
        "Debtors Turnover": (4.0, 12.0),
        "Creditors Turnover": (4.0, 12.0),
        "Working Capital Turnover": (3.0, 8.0),
    }
    low, high = rules.get(name, (None, None))
    if low is None:
        return ("NA", "NA")
    ok = low <= value <= high
    return (f"{low:.2f} - {high:.2f}", "OK" if ok else "Alert")
