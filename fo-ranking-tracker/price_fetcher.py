from __future__ import annotations

import time
from datetime import date, datetime, timedelta, timezone
from typing import Dict, Iterable, List, Optional, Tuple
from zoneinfo import ZoneInfo

import requests

from nse_calendar import is_trading_day, previous_trading_day, resolve_trading_date

YAHOO_CHART_URL = "https://query1.finance.yahoo.com/v8/finance/chart/{symbol}"
IST = ZoneInfo("Asia/Kolkata")
PRICE_DECIMALS = 2


def _round_price(price: float) -> float:
    return round(price, PRICE_DECIMALS)


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


def trading_date_for(target: date) -> date:
    return resolve_trading_date(target)


def _to_unix(d: date) -> int:
    return int(datetime(d.year, d.month, d.day, tzinfo=timezone.utc).timestamp())


def _today_ist() -> date:
    return datetime.now(IST).date()


def today_ist() -> date:
    """Current calendar date in India (NSE timezone)."""
    return _today_ist()


def _market_session_open_ist() -> bool:
    now = datetime.now(IST)
    if now.weekday() >= 5:
        return False
    session_open = now.replace(hour=9, minute=15, second=0, microsecond=0)
    session_close = now.replace(hour=15, minute=30, second=0, microsecond=0)
    return session_open <= now < session_close


def _today_bar_is_final() -> bool:
    """True only after NSE cash session has finished for today (15:30 IST)."""
    now = datetime.now(IST)
    today = now.date()
    if not is_trading_day(today):
        return False
    session_close = now.replace(hour=15, minute=30, second=0, microsecond=0)
    return now >= session_close


def _last_completed_trading_day() -> date:
    today = _today_ist()
    if is_trading_day(today) and not _market_session_open_ist():
        return today
    return previous_trading_day(today)


def _close_for_trading_date(
    history: List[Tuple[date, float]],
    trading_date: date,
) -> Optional[float]:
    if not history:
        return None

    by_date = {d: price for d, price in history}

    if trading_date in by_date:
        return by_date[trading_date]

    # Walk back to the nearest earlier session with data (splits/listing gaps).
    current = trading_date
    earliest = history[0][0]
    while current >= earliest:
        if current in by_date:
            return by_date[current]
        current = previous_trading_day(current)
    return None


def _close_on_or_before(
    history: List[Tuple[date, float]],
    target: date,
) -> Optional[float]:
    if not history:
        return None

    trading_date = resolve_trading_date(target)
    today = _today_ist()

    # During live session, "today" means last completed NSE close.
    if trading_date >= today and _market_session_open_ist():
        trading_date = _last_completed_trading_day()

    return _close_for_trading_date(history, trading_date)


def _fetch_history(symbol: str, start: date, end: date) -> List[Tuple[date, float]]:
    period1 = _to_unix(start)
    period2 = _to_unix(end + timedelta(days=2))

    params = {
        "period1": period1,
        "period2": period2,
        "interval": "1d",
        "events": "history",
        "includeAdjustedClose": "false",
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
    indicators = result.get("indicators") or {}

    quote_block = (indicators.get("quote") or [{}])[0]
    closes = quote_block.get("close") or []

    history: List[Tuple[date, float]] = []
    today = _today_ist()
    for index, ts in enumerate(timestamps):
        if index >= len(closes) or closes[index] is None:
            continue

        day = datetime.fromtimestamp(ts, tz=IST).date()

        # Yahoo sometimes forward-fills holidays (e.g. 26-Jun-2026 Muharram).
        if not is_trading_day(day):
            continue

        # Never use today's bar until the NSE session close is published.
        if day == today and not _today_bar_is_final():
            continue

        close_price = _round_price(float(closes[index]))
        history.append((day, close_price))

    deduped: Dict[date, float] = {}
    for day, price in history:
        deduped[day] = price

    return sorted(deduped.items(), key=lambda item: item[0])


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
            if index < len(self.symbols) - 1:
                time.sleep(0.15)

    def price_on(self, raw_symbol: str, target: date) -> Optional[float]:
        symbol = to_nse_symbol(raw_symbol)
        history = self._history.get(symbol)
        if history is None:
            return None
        return _close_on_or_before(history, target)

    def trading_date_used(self, target: date) -> date:
        trading_date = resolve_trading_date(target)
        today = _today_ist()
        if trading_date >= today and _market_session_open_ist():
            return _last_completed_trading_day()
        return trading_date
