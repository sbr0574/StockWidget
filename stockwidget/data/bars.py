"""Lightweight, unadjusted history adapters. No Qt or optional data libraries.

Minute timestamps are normalized to the instrument's exchange wall clock.
Missing prices are rejected rather than turned into zero-price candles.
"""

from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
import json
import math
import re

import requests

from stockwidget.data.network_errors import request_error_message
from stockwidget.data.quotes import _em_secid, _sina_code, _EM_HEADERS, _SINA_HEADERS


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
            bars.append(make_bar(stamp, parts[1], parts[3], parts[4], parts[2],
                                 parts[5], parts[6], None if daily else parts[7]))
        except (ValueError, TypeError, IndexError, AttributeError):
            continue
    return _ordered(bars)


def parse_sina(text):
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
    return select_bars(parse_sina(response.text), view, instrument)


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
