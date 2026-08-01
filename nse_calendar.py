from __future__ import annotations

from datetime import date, timedelta
from functools import lru_cache
from typing import FrozenSet

# Official NSE trading holidays (equity segment).
# Weekends are excluded automatically; only weekday closures are listed.
NSE_HOLIDAYS: FrozenSet[date] = frozenset(
    {
        # 2024
        date(2024, 1, 26),
        date(2024, 3, 8),
        date(2024, 3, 25),
        date(2024, 3, 29),
        date(2024, 4, 11),
        date(2024, 4, 17),
        date(2024, 4, 21),
        date(2024, 5, 1),
        date(2024, 6, 17),
        date(2024, 7, 17),
        date(2024, 8, 15),
        date(2024, 10, 2),
        date(2024, 11, 1),
        date(2024, 11, 15),
        date(2024, 12, 25),
        # 2025
        date(2025, 2, 26),
        date(2025, 3, 14),
        date(2025, 3, 31),
        date(2025, 4, 10),
        date(2025, 4, 14),
        date(2025, 4, 18),
        date(2025, 5, 1),
        date(2025, 8, 15),
        date(2025, 8, 27),
        date(2025, 10, 2),
        date(2025, 10, 22),
        date(2025, 11, 5),
        date(2025, 12, 25),
        # 2026
        date(2026, 1, 15),
        date(2026, 1, 26),
        date(2026, 3, 3),
        date(2026, 3, 26),
        date(2026, 3, 31),
        date(2026, 4, 3),
        date(2026, 4, 14),
        date(2026, 5, 1),
        date(2026, 5, 28),
        date(2026, 6, 26),  # Muharram — Friday holiday; use Thursday close
        date(2026, 9, 14),
        date(2026, 10, 2),
        date(2026, 10, 20),
        date(2026, 11, 10),
        date(2026, 11, 24),
        date(2026, 12, 25),
        # 2027 (extend as NSE publishes)
        date(2027, 1, 26),
        date(2027, 3, 22),
        date(2027, 3, 26),
        date(2027, 4, 14),
        date(2027, 5, 1),
        date(2027, 8, 15),
        date(2027, 10, 2),
        date(2027, 12, 25),
    }
)


def is_trading_day(day: date) -> bool:
    return day.weekday() < 5 and day not in NSE_HOLIDAYS


@lru_cache(maxsize=4096)
def previous_trading_day(day: date) -> date:
    current = day - timedelta(days=1)
    while not is_trading_day(current):
        current -= timedelta(days=1)
    return current


@lru_cache(maxsize=4096)
def resolve_trading_date(target: date) -> date:
    """Map a calendar date to the NSE session used for its close price."""
    if is_trading_day(target):
        return target
    return previous_trading_day(target)
