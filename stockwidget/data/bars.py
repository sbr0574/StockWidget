"""Lightweight, unadjusted history adapters. No Qt or optional data libraries.

Minute timestamps are normalized to the instrument's exchange wall clock.
Missing prices are rejected rather than turned into zero-price candles.
"""

from dataclasses import dataclass, replace
from datetime import datetime, timedelta, timezone
import json
import math
import re

import requests

from stockwidget.data.network_errors import request_error_message
from stockwidget.data.quotes import _em_secid, _sina_code, _EM_HEADERS, _SINA_HEADERS
from stockwidget.core.window_rules import quote_timestamp


MA_PERIODS = (5, 10, 20, 30, 60)
DAILY_HISTORY = 100  # 30 visible candles + 59 warm-up candles for MA60.


@dataclass(frozen=True)
class Bar:
    time: str
    open: float
    high: float
    low: float
    close: float
    volume: float = 0
    amount: float | None = None
    average: float | None = None


@dataclass(frozen=True)
class BarResult:
    bars: tuple[Bar, ...] = ()
    source: str = ""
    message: str = ""
    cached: bool = False
    stale: bool = False


def us_offset(utc_time):
    """US DST boundaries in UTC, without a tzdata dependency on Windows."""
    year = utc_time.year
    march = datetime(year, 3, 8, 7, tzinfo=timezone.utc)
    start = march + timedelta(days=(6 - march.weekday()) % 7)
    november = datetime(year, 11, 1, 6, tzinfo=timezone.utc)
    end = november + timedelta(days=(6 - november.weekday()) % 7)
    return -4 if start <= utc_time < end else -5


def market_now(instrument, now=None):
    now = now or datetime.now(timezone.utc)
    if now.tzinfo is None:
        now = now.replace(tzinfo=timezone.utc)
    now = now.astimezone(timezone.utc)
    hours = us_offset(now) if instrument.get("market") == "us" else 8
    return now.astimezone(timezone(timedelta(hours=hours)))


def history_day(instrument, now):
    """One bootstrap per exchange day, including the next futures night day."""
    local = market_now(instrument, now)
    if instrument.get("type") == "期" or not instrument.get("market"):
        return trading_date(Bar(local.isoformat(), 0, 0, 0, 0), instrument)
    day = local.date()
    while day.weekday() >= 5:
        day -= timedelta(days=1)
    return day.isoformat()


def quote_clock(quote, instrument):
    stamp = quote_timestamp(quote)
    if stamp is None:
        return None
    return market_now(instrument, datetime.fromtimestamp(stamp, timezone.utc)).replace(tzinfo=None)


def _minute_sessions(instrument):
    market = instrument.get("market")
    if instrument.get("type") == "期" or not market:
        return ((0, 150), (540, 615), (630, 690), (810, 900), (1260, 1440))
    if market == "us":
        return ((570, 960),)
    if market == "hk":
        return ((570, 720), (780, 960))
    if market in {"sh", "sz", "bj"}:
        return ((570, 690), (780, 900))
    return ((0, 1440),)  # Unknown exchange: retain its observed clock.


def minute_gaps(bars, instrument, current):
    """Detect missing completed minutes; lunch/overnight breaks aren't gaps.

    Providers may omit the opening 09:30/21:00 point. Accept their first full
    minute, and never infer previous night availability for a futures product.
    """
    day = trading_date(Bar(current.isoformat(), 0, 0, 0, 0), instrument)
    observed = {datetime.fromisoformat(bar.time).replace(second=0, microsecond=0)
                for bar in bars if len(bar.time) > 10 and trading_date(bar, instrument) == day}
    if observed:
        last = max(observed)
        if last.date() < current.date() and last.hour >= 21 and current.hour < 3:
            candidate = last + timedelta(minutes=1)
            while candidate < current:
                if candidate not in observed:
                    return True
                candidate += timedelta(minutes=1)
    dates = {current.date(), *(stamp.date() for stamp in observed)}
    for date in dates:
        midnight = datetime.combine(date, datetime.min.time())
        for start, end in _minute_sessions(instrument):
            # Night end times vary: inspect the current night segment rather
            # than assuming every product traded until 02:30 the night before.
            segment = midnight + timedelta(minutes=start)
            stop = midnight + timedelta(minutes=end)
            if not (segment <= current <= stop):
                if start == 0 or start >= 1260:
                    continue
                if stop > current:
                    continue
            if instrument.get("market") == "gb":
                segment = min(observed, default=current) - timedelta(minutes=1)
            candidate = segment + timedelta(minutes=1)
            while candidate < min(current, stop):
                if trading_date(Bar(candidate.isoformat(), 0, 0, 0, 0), instrument) == day and candidate not in observed:
                    return True
                candidate += timedelta(minutes=1)
    return False


def merge_quote(bars, instrument, view, quote, previous=None, *, gaps=False):
    """Supplement a minute/current-day candle using accepted raw quotes.

    Volume in this schema is shares (contracts for futures), like live quotes.
    A missing span must not be assigned wholesale to the newest minute.
    """
    current = quote_clock(quote, instrument)
    price = _number(quote.get("current_price"), 0)
    opening = _number(quote.get("opening_price"), 0)
    if current is None or not price or not opening:
        return tuple(bars)
    stamp = current.replace(second=0, microsecond=0).strftime("%Y-%m-%d %H:%M:%S")
    day = trading_date(Bar(stamp, 0, 0, 0, 0), instrument)
    daily = view == "daily"
    key = day if daily else stamp
    if bars and key < bars[-1].time:
        return tuple(bars)  # Stale/out-of-order quotes cannot roll back history.
    total_volume = max(0, _number(quote.get("deals_vol"), 0))
    total_amount = max(0, _number(quote.get("deals_amt"), 0))
    old = bars[-1] if bars and bars[-1].time == key else None
    if daily:
        high = max(price, opening, _number(quote.get("high_price"), price))
        low = min(price, opening, _number(quote.get("low_price"), price) or price)
        if old:
            high, low = max(high, old.high), min(low, old.low)
            total_volume = max(total_volume, old.volume)
        bar = Bar(key, opening, high, low, price, total_volume, total_amount or None)
    else:
        minute = current.hour * 60 + current.minute
        futures = instrument.get("type") == "期" or not instrument.get("market")
        if current.weekday() >= 5 and not (futures and current.weekday() == 5 and minute <= 150):
            return tuple(bars)
        if not any(start <= minute <= end for start, end in _minute_sessions(instrument)):
            return tuple(bars)
        prefix = [bar for bar in bars if bar.time < key and trading_date(bar, instrument) == day]
        old_volume, old_amount = (old.volume, old.amount or 0) if old else (0, 0)
        previous_clock = quote_clock(previous, instrument) if previous else None
        same_minute = previous_clock and previous_clock.replace(second=0, microsecond=0) == current.replace(second=0, microsecond=0)
        volume = (max(old_volume, total_volume - sum(bar.volume for bar in prefix)) if not gaps
                  else old_volume + (max(0, total_volume - _number(previous.get("deals_vol"), 0)) if same_minute else 0))
        amount = None
        if total_amount:
            if not gaps and all(bar.amount is not None for bar in prefix):
                amount = max(old_amount, total_amount - sum(bar.amount for bar in prefix))
            elif same_minute:
                amount = old_amount + max(0, total_amount - _number(previous.get("deals_amt"), 0))
        # Index turnover/volume is the components' price, not index points.
        index = instrument.get("type") == "指"
        average = total_amount / total_volume if total_volume and total_amount and instrument.get("type") not in {"期", "指"} else None
        bar = Bar(key, old.open if old else price, max(old.high, price) if old else price,
                  min(old.low, price) if old else price, price, volume, amount,
                  None if index else average or (old.average if old else None))
    merged = (*bars[:-1], bar) if old else (*bars, bar)
    return select_bars(merged, "daily" if daily else "five_day", instrument)


class HistorySeries:
    """Historical bootstrap plus live observations, without a polling timer."""
    def __init__(self, instrument, view):
        self.instrument = dict(instrument)
        self.view = "daily" if view == "daily" else "five_day"
        self.result = BarResult()
        self.loaded_day = None
        self.repair_needed = False
        self.retry_after = 0
        self._quote = None

    @property
    def reference_price(self):
        """Latest plotted session's previous close, never a guessed opening."""
        if not self.result.bars:
            return None
        day = trading_date(self.result.bars[-1], self.instrument)
        clock = quote_clock(self._quote, self.instrument) if self._quote else None
        if clock and trading_date(Bar(clock.isoformat(), 0, 0, 0, 0), self.instrument) == day:
            previous = _number(self._quote.get("prev_close"), 0)
            if previous > 0:
                return previous
        previous = next((bar for bar in reversed(self.result.bars)
                         if trading_date(bar, self.instrument) < day), None)
        if previous and previous.close > 0:
            if len(previous.time) == 10:
                return previous.close
            market = self.instrument.get("market")
            closing = 960 if market in {"us", "hk"} else 900 if market in {"sh", "sz", "bj"} else None
            stamp = datetime.fromisoformat(previous.time)
            if (self.instrument.get("type") != "期" and closing
                    and stamp.hour * 60 + stamp.minute in {closing - 1, closing}):
                return previous.close
        return None

    def needs_download(self, now):
        if self.loaded_day is not None and self.loaded_day != history_day(self.instrument, now):
            return True
        if now.timestamp() < self.retry_after:
            return False
        return (self.loaded_day != history_day(self.instrument, now)
                or not self.result.bars or self.repair_needed)

    def set_history(self, result, day, now):
        if not result.bars and self.result.bars:
            result = replace(result, bars=self.result.bars, source=self.result.source, cached=True, stale=True)
        elif self.result.bars and result.source == self.result.source:
            # Repairs must keep minutes observed live while the endpoint was
            # delayed/offline. Fresh downloaded candles correct sampled OHLC.
            first, second = ((result.bars, self.result.bars) if result.stale else (self.result.bars, result.bars))
            merged = {bar.time: bar for bar in (*first, *second)}
            clock = quote_clock(self._quote, self.instrument) if self._quote else None
            if self.view != "daily" and clock:
                stamp = clock.replace(second=0, microsecond=0).strftime("%Y-%m-%d %H:%M:%S")
                live = next((bar for bar in reversed(self.result.bars) if bar.time == stamp), None)
                if live and stamp in merged:
                    bar = merged[stamp]
                    merged[stamp] = replace(bar, high=max(bar.high, live.high), low=min(bar.low, live.low))
            result = replace(result, bars=select_bars(tuple(sorted(merged.values(), key=lambda bar: bar.time)),
                                                     self.view, self.instrument))
        self.result, self.loaded_day = result, day
        self.repair_needed = result.stale or not result.bars
        latest, self._quote = self._quote, None
        if latest:
            self.update_quote(latest)
        # Empty/short/delayed responses cannot trigger a request on every tick.
        self.retry_after = now.timestamp() + 300 if self.repair_needed else 0

    def update_quote(self, quote):
        if self._quote == quote:
            return
        current = quote_clock(quote, self.instrument)
        previous_clock = quote_clock(self._quote, self.instrument) if self._quote else None
        if (current is None or not _number(quote.get("current_price"), 0)
                or not _number(quote.get("opening_price"), 0) or (previous_clock and current < previous_clock)):
            return
        if self.view != "daily" and self.loaded_day is not None and not self.repair_needed:
            self.repair_needed |= minute_gaps(self.result.bars, self.instrument, current.replace(second=0, microsecond=0))
        if self.loaded_day is not None:
            self.result = replace(self.result, bars=merge_quote(self.result.bars, self.instrument, self.view,
                                                              quote, self._quote, gaps=self.repair_needed))
        self._quote = dict(quote)


def _number(value, default=None):
    try:
        value = float(value)
        return value if math.isfinite(value) else default
    except (TypeError, ValueError, OverflowError):
        return default


def make_bar(time, opening, high, low, close, volume=0, amount=None, average=None):
    parsed = datetime.fromisoformat(str(time).replace("/", "-"))
    stamp = parsed.strftime("%Y-%m-%d %H:%M:%S") if " " in str(time) or "T" in str(time) else parsed.strftime("%Y-%m-%d")
    prices = [_number(value) for value in (opening, high, low, close)]
    if any(value is None for value in prices):
        raise ValueError("Missing OHLC")
    opening, high, low, close = prices
    # trends2 can report zero open/high/low for US opening minutes.
    if close == 0:
        raise ValueError("Missing close")
    opening = opening or close
    high = high or max(opening, close)
    low = low or min(opening, close)
    if high < max(opening, close) or low > min(opening, close) or high < low:
        raise ValueError("Invalid OHLC")
    return Bar(stamp, opening, high, low, close, _number(volume, 0),
               _number(amount), _number(average))


def _ordered(bars):
    return tuple(sorted({bar.time: bar for bar in bars}.values(), key=lambda bar: bar.time))


def parse_eastmoney(payload, *, daily=False, instrument=None):
    data = payload.get("data") if isinstance(payload, dict) else None
    if not isinstance(data, dict):
        return ()
    rows = data.get("klines" if daily else "trends") or []
    bars = []
    for row in rows:
        try:
            parts = row.split(",")
            stamp = parts[0]
            if not daily and (instrument or {}).get("market") == "us":
                beijing = datetime.fromisoformat(stamp).replace(tzinfo=timezone(timedelta(hours=8)))
                stamp = market_now(instrument, beijing).strftime("%Y-%m-%d %H:%M:%S")
            volume = _number(parts[5], 0)
            if (instrument or {}).get("market") in {"sh", "sz", "bj"} and (instrument or {}).get("type") != "期":
                volume *= 100  # Eastmoney history uses lots; quotes use shares.
            bars.append(make_bar(stamp, parts[1], parts[3], parts[4], parts[2],
                                 volume, parts[6], None if daily or (instrument or {}).get("type") == "指" else parts[7]))
        except (ValueError, TypeError, IndexError, AttributeError):
            continue
    return _ordered(bars)


def parse_sina(text, *, daily=False):
    """Accept JSON/JSONP and legacy bare object keys; never evaluate JavaScript."""
    start, stop = text.find("["), text.rfind("]")
    if start < 0 or stop < start:
        return ()
    body = text[start:stop + 1]
    body = re.sub(r'([{,]\s*)([A-Za-z_][A-Za-z_0-9]*)(\s*:)', r'\1"\2"\3', body)
    rows = json.loads(body)
    bars = []
    for row in rows:
        try:
            if isinstance(row, dict):
                get = lambda full, short: row.get(full, row.get(short))
                bar = make_bar(row.get("day", row.get("date", row.get("d"))),
                               get("open", "o"), get("high", "h"), get("low", "l"),
                               get("close", "c"), get("volume", "v"), get("amount", "a"))
            else:
                bar = make_bar(*row[:6])  # futures: time, open, high, low, close, volume
            if daily:
                bar = replace(bar, time=bar.time[:10])
            bars.append(bar)
        except (ValueError, TypeError, IndexError, AttributeError):
            continue
    return _ordered(bars)


def trading_date(bar, instrument):
    date = datetime.fromisoformat(bar.time)
    if instrument.get("type") == "期" or not instrument.get("market"):
        if date.hour >= 20:
            date += timedelta(days=1)
        while date.weekday() >= 5:
            date += timedelta(days=1)
    return date.strftime("%Y-%m-%d")


def select_bars(bars, view, instrument):
    if view == "daily":
        return tuple(bars[-DAILY_HISTORY:])
    dates = sorted({trading_date(bar, instrument) for bar in bars})
    selected = set(dates[-(1 if view == "intraday" else 5):])
    return tuple(bar for bar in bars if trading_date(bar, instrument) in selected)


def moving_average(bars, period):
    """Compute on full history, with None until a complete window exists."""
    if period <= 0:
        raise ValueError("period must be positive")
    result, total = [], 0.0
    for i, bar in enumerate(bars):
        total += bar.close
        if i >= period:
            total -= bars[i - period].close
        result.append(total / period if i >= period - 1 else None)
    return result


def intraday_average(bars, instrument):
    """Session VWAP, with minute-price estimates where turnover is unavailable.

    Index turnover belongs to its components, so average index points instead.
    Futures turnover can include a contract multiplier; never divide it by lots.
    """
    result = []
    day = None
    index = instrument.get("type") == "指"
    futures = instrument.get("type") == "期" or not instrument.get("market")
    volume = weighted = prices = count = 0
    for bar in bars:
        current_day = trading_date(bar, instrument)
        if current_day != day:
            day = current_day
            volume = weighted = prices = count = 0
        prices += bar.close
        count += 1
        quantity = max(0, bar.volume)
        amount = _number(bar.amount)
        volume += quantity
        weighted += (amount if not futures and amount is not None and amount > 0 and quantity > 0
                     else bar.close * quantity)
        supplied = _number(bar.average)
        if index:
            value = prices / count
        elif supplied is not None and supplied > 0:
            value = supplied
        else:
            value = weighted / volume if volume else prices / count
        result.append(value)
    return result


def _em_candidates(instrument):
    secid = _em_secid(instrument)
    if instrument.get("market") == "us" and secid:
        return [f"{market}.{secid.split('.', 1)[1]}" for market in (105, 106, 107)]
    return [secid] if secid else []


def request_eastmoney_bars(instrument, view):
    daily = view == "daily"
    for secid in _em_candidates(instrument):
        params = {"secid": secid, "fields1": "f1,f2,f3,f4,f5,f6,f7,f8,f9,f10,f11,f12,f13",
                  "fields2": "f51,f52,f53,f54,f55,f56,f57,f58"}
        if daily:
            params.update(klt=101, fqt=0, beg=0, end=20500101, lmt=DAILY_HISTORY)
        else:
            params.update(ndays=5, iscr=0)
        response = requests.get(
            "https://push2his.eastmoney.com/api/qt/stock/" + ("kline/get" if daily else "trends2/get"),
            params=params, headers=_EM_HEADERS, timeout=5)
        response.raise_for_status()
        bars = parse_eastmoney(response.json(), daily=daily, instrument=instrument)
        if bars:
            return select_bars(bars, view, instrument)
    return ()


def request_sina_bars(instrument, view):
    market = instrument.get("market", "")
    futures = instrument.get("type") == "期" or not market
    if futures:
        method = "getDailyKLine" if view == "daily" else "getFewMinLine"
        url = "https://stock2.finance.sina.com.cn/futures/api/jsonp.php/=/InnerFuturesNewService." + method
        params = {"symbol": str(instrument.get("code", "")).upper(),
                  "type": market_now(instrument).strftime("%Y_%m_%d") if view == "daily" else "1"}
    elif market in {"sh", "sz", "bj"}:
        url = "https://quotes.sina.cn/cn/api/jsonp_v2.php/=/CN_MarketDataService.getKLineData"
        params = {"symbol": _sina_code(instrument), "scale": 240 if view == "daily" else 1,
                  "ma": "no", "datalen": DAILY_HISTORY if view == "daily" else 1970}
    else:
        # HK/US daily use encoded JS; no JS engine dependency for charting.
        # Unsupported adapters return empty so the other provider can be tried.
        return ()
    response = requests.get(url, params=params, headers=_SINA_HEADERS, timeout=5)
    response.raise_for_status()
    return select_bars(parse_sina(response.text, daily=view == "daily"), view, instrument)


def fetch_bars(instrument, view, source):
    if view not in {"intraday", "five_day", "daily"}:
        raise ValueError("Unknown history view")
    source = source if source in {"sina", "eastmoney"} else "sina"
    errors = []
    partial = None
    for provider in (source, "sina" if source == "eastmoney" else "eastmoney"):
        try:
            request = request_sina_bars if provider == "sina" else request_eastmoney_bars
            bars = request(instrument, view)
            if bars:
                dates = {trading_date(bar, instrument) for bar in bars}
                message = ""
                if view == "five_day" and len(dates) < 5:
                    message = f"接口仅返回 {len(dates)} 个交易日"
                    candidate = BarResult(tuple(bars), provider, message)
                    if partial is None or len(dates) > len({trading_date(bar, instrument) for bar in partial.bars}):
                        partial = candidate
                    continue  # Try the backup before settling for partial history.
                return BarResult(tuple(bars), provider, message)
        except requests.RequestException as exc:
            errors.append(f"{'新浪' if provider == 'sina' else '东财'}：{request_error_message(exc)}")
        except (ValueError, TypeError, KeyError, AttributeError):
            errors.append(f"{'新浪' if provider == 'sina' else '东财'}：历史数据格式不可用")
    return partial or BarResult(message="暂无可用历史数据" + (" · " + "；".join(errors) if errors else ""))
