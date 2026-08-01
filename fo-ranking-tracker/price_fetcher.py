from __future__ import annotations

import time
from datetime import date, datetime, timedelta, timezone
from typing import Dict, Iterable, List, Optional, Tuple

import requests

YAHOO_CHART_URL = "https://query1.finance.yahoo.com/v8/finance/chart/{symbol}"


def to_nse_symbol(symbol: str) -> str:
    cleaned = str(symbol).strip().upper()
    if not cleaned:
        return ""
    if cleaned.endswith(".NS"):
        return cleaned
    if cleaned.endswith("-EQ"):
        cleaned = cleaned[:-3]
    return f"{cleaned}.NS"


def get_fridays(start: date, end: date) -> List[date]:
    if start > end:
        return []

    current = start
    while current.weekday() != 4:
        current += timedelta(days=1)

    fridays: List[date] = []
    while current <= end:
        fridays.append(current)
        current += timedelta(days=7)
    return fridays


def _to_unix(d: date) -> int:
    return int(datetime(d.year, d.month, d.day, tzinfo=timezone.utc).timestamp())


def _close_on_or_before(
    history: List[Tuple[date, float]],
    target: date,
) -> Optional[float]:
    if not history:
        return None

    eligible = [(d, price) for d, price in history if d <= target]
    if not eligible:
        return None
    return eligible[-1][1]


def _fetch_history(symbol: str, start: date, end: date) -> List[Tuple[date, float]]:
    period1 = _to_unix(start)
    period2 = _to_unix(end + timedelta(days=1))

    params = {
        "period1": period1,
        "period2": period2,
        "interval": "1d",
        "events": "history",
    }
    headers = {"User-Agent": "Mozilla/5.0"}

    try:
        response = requests.get(
            YAHOO_CHART_URL.format(symbol=symbol),
            params=params,
            headers=headers,
            timeout=20,
        )
        response.raise_for_status()
        payload = response.json()
    except Exception:
        return []

    results = payload.get("chart", {}).get("result") or []
    if not results:
        return []

    result = results[0]
    timestamps = result.get("timestamp") or []
    quotes = (result.get("indicators", {}).get("quote") or [{}])[0]
    closes = quotes.get("close") or []

    history: List[Tuple[date, float]] = []
    for ts, close in zip(timestamps, closes):
        if close is None:
            continue
        day = datetime.fromtimestamp(ts, tz=timezone.utc).date()
        history.append((day, float(close)))

    history.sort(key=lambda item: item[0])
    return history


class PriceFetcher:
    def __init__(self, symbols: Iterable[str], start: date, end: date):
        self.symbols = [to_nse_symbol(s) for s in symbols if str(s).strip()]
        self.start = start
        self.end = end
        self._history: Dict[str, List[Tuple[date, float]]] = {}

    def load(self) -> None:
        if not self.symbols:
            return

        fetch_start = self.start - timedelta(days=10)
        fetch_end = self.end

        for index, symbol in enumerate(self.symbols):
            self._history[symbol] = _fetch_history(symbol, fetch_start, fetch_end)
            # Light throttle to avoid Yahoo rate limits.
            if index < len(self.symbols) - 1:
                time.sleep(0.15)

    def price_on(self, raw_symbol: str, target: date) -> Optional[float]:
        symbol = to_nse_symbol(raw_symbol)
        history = self._history.get(symbol)
        if history is None:
            return None
        return _close_on_or_before(history, target)
