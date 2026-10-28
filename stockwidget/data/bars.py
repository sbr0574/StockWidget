"""Lightweight, unadjusted history adapters. No Qt or optional data libraries.

Minute timestamps are normalized to the instrument's exchange wall clock.
Missing prices are rejected rather than turned into zero-price candles.
"""

from dataclasses import dataclass, replace
from datetime import datetime, timedelta, timezone
import json
import math
import re
from zoneinfo import ZoneInfo

import requests

from stockwidget.data.network_errors import request_error_message
from stockwidget.data.quotes import _sina_code, _SINA_HEADERS, request_quote
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
    source: str = "sina"
    message: str = ""
    cached: bool = False
    stale: bool = False


def market_now(instrument, now=None):
    now = now or datetime.now(timezone.utc)
    if now.tzinfo is None:
        now = now.replace(tzinfo=timezone.utc)
    market = instrument.get("market")
    zone = ("America/New_York" if market == "us" else
            instrument.get("timezone", "Asia/Shanghai")
            if market == "gb" else "Asia/Shanghai")
    return now.astimezone(ZoneInfo(zone))


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
    if instrument.get("market") == "gb" and not instrument.get("timezone"):
        return None  # Unknown exchange clock must not create a guessed daily candle.
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
        # Minute endpoints commonly begin at 09:31. The accepted quote still
        # supplies the actual 09:30 opening price, even before that first bar
        # arrives. Keep it without inventing a minute's volume or extrema.
        opening_stamp = (current.replace(hour=9, minute=30, second=0, microsecond=0)
                         .strftime("%Y-%m-%d %H:%M:%S")
                         if not futures and instrument.get("market") in {"sh", "sz", "bj", "hk", "us"}
                         else None)
        if opening_stamp and key >= opening_stamp:
            opening_bar = next((bar for bar in bars if bar.time == opening_stamp), None)
            if opening_bar:
                corrected = replace(opening_bar, open=opening, high=max(opening_bar.high, opening),
                                    low=min(opening_bar.low, opening))
                bars = tuple(corrected if bar.time == opening_stamp else bar for bar in bars)
                old = corrected if key == opening_stamp else old
            elif key > opening_stamp:
                bars = tuple(sorted((*bars, Bar(opening_stamp, opening, opening, opening, opening, 0, 0)),
                                    key=lambda bar: bar.time))
        prefix = [bar for bar in bars if bar.time < key and trading_date(bar, instrument) == day]
        old_volume, old_amount = (old.volume, old.amount or 0) if old else (0, 0)
        previous_clock = quote_clock(previous, instrument) if previous else None
        same_minute = previous_clock and previous_clock.replace(second=0, microsecond=0) == current.replace(second=0, microsecond=0)
        volume = (max(old_volume, total_volume - sum(bar.volume for bar in prefix)) if not gaps or key == opening_stamp
                  else old_volume + (max(0, total_volume - _number(previous.get("deals_vol"), 0)) if same_minute else 0))
        amount = None
        if total_amount:
            if (not gaps or key == opening_stamp) and all(bar.amount is not None for bar in prefix):
                amount = max(old_amount, total_amount - sum(bar.amount for bar in prefix))
            elif same_minute:
                amount = old_amount + max(0, total_amount - _number(previous.get("deals_amt"), 0))
        # Index turnover/volume is the components' price, not index points.
        index = instrument.get("type") == "指"
        average = total_amount / total_volume if total_volume and total_amount and instrument.get("type") not in {"期", "指"} else None
        bar_open = opening if key == opening_stamp else old.open if old else price
        bar = Bar(key, bar_open, max(old.high if old else bar_open, price),
                  min(old.low if old else bar_open, price), price, volume, amount,
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
        return self.reference_price_for(self.result.bars)

    def reference_price_for(self, bars):
        """Resolve the same session baseline for cached or live bootstrap bars."""
        if not bars:
            return None
        day = trading_date(bars[-1], self.instrument)
        clock = quote_clock(self._quote, self.instrument) if self._quote else None
        if clock and trading_date(Bar(clock.isoformat(), 0, 0, 0, 0), self.instrument) == day:
            previous = _number(self._quote.get("prev_close"), 0)
            if previous > 0:
                return previous
        previous = next((bar for bar in reversed(bars)
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
        elif self.result.bars:
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
    # A minute point may contain a close without separate OHLC observations.
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


def intraday_timeline(bars, instrument, now=None):
    """Current exchange day's minute slots through close, without fake bars.

    Breaks consume no chart width. Past days/unknown exchanges keep observed
    timestamps; futures include a night session only when it was observed.
    """
    now = now or datetime.now(timezone.utc)
    if not bars or trading_date(bars[-1], instrument) != history_day(instrument, now):
        return ()
    market = instrument.get("market")
    futures = instrument.get("type") == "期" or not market
    day = datetime.fromisoformat(trading_date(bars[-1], instrument))
    clocks = tuple(datetime.fromisoformat(bar.time) for bar in bars)
    sessions = []
    if futures:
        night = next((stamp for stamp in clocks if stamp.hour >= 20 or stamp.hour < 3), None)
        if night:
            product = re.match(r"[a-z]+", str(instrument.get("code", "")).lower())
            product = product[0] if product else ""
            # Regular hours: shfe.com.cn/services/calenderandholidays/tradinghours/
            if product in {"au", "ag", "sc"}:
                duration = 330
            elif product in {"cu", "bc", "al", "ao", "zn", "pb", "ni", "sn", "ss", "ad"}:
                duration = 240
            elif product in {"rb", "hc", "fu", "bu", "ru", "br", "sp", "op", "nr", "lu"}:
                duration = 120
            else:
                return ()
            night = night.replace(hour=21, minute=0, second=0, microsecond=0) - timedelta(days=night.hour < 3)
            sessions.append((night, night + timedelta(minutes=duration)))
        ranges = ((540, 615), (630, 690), (810, 900))
    elif market in {"sh", "sz", "bj", "hk", "us"}:
        local = market_now(instrument, now)
        if day.date() != local.date() or local.weekday() >= 5:
            return ()
        ranges = _minute_sessions(instrument)
    else:
        return ()
    sessions += [(day + timedelta(minutes=start), day + timedelta(minutes=end)) for start, end in ranges]
    times = tuple((start + timedelta(minutes=i)).strftime("%Y-%m-%d %H:%M:%S")
                  for start, end in sessions for i in range(int((end - start).total_seconds() / 60) + 1))
    available = set(times)
    if any(stamp.strftime("%Y-%m-%d %H:%M:00") not in available for stamp in clocks):
        return ()
    return times


def mini_trend_points(bars, instrument, now=None):
    """Keep the latest close of each three-minute bucket on the full session axis.

    Minute histories remain intact for detailed charts and the shared cache.
    The unfinished bucket follows the latest quote, including across breaks.
    """
    timeline = intraday_timeline(bars, instrument, now)
    slots = {time: index for index, time in enumerate(timeline)}
    count = max(1, len(timeline or bars))
    samples = {}
    for index, bar in enumerate(bars):
        stamp = datetime.fromisoformat(bar.time)
        bucket = stamp.replace(minute=stamp.minute // 3 * 3, second=0, microsecond=0)
        samples[bucket] = ((slots.get(bar.time, index) + .5) / count, bar.close)
    points = tuple(samples.values())
    if bars and bars[0].time[11:16] == "09:30" and instrument.get("market") in {"sh", "sz", "bj", "hk", "us"}:
        points = ((0.0, bars[0].open), *points)
    return points


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


def volume_directions(bars, view, instrument):
    """Minute volumes follow consecutive prices within the same trading day.

    Flat minutes retain the last direction; an unknown first direction is
    neutral. Daily candles continue to use their own open/close direction.
    """
    if view == "daily":
        return [1 if bar.close >= bar.open else -1 for bar in bars]
    result = []
    day = previous = None
    direction = 0
    for bar in bars:
        current_day = trading_date(bar, instrument)
        if current_day != day:
            day, previous, direction = current_day, bar.open, 0
        if bar.close != previous:
            direction = 1 if bar.close > previous else -1
        result.append(direction)
        previous = bar.close
    return result


class _SinaDailyBits:
    """Read the little-endian six-bit stream used by Sina's KLC K2 daily data."""
    def __init__(self, text):
        alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
        self.values = [alphabet.index(char) for char in text]
        self.position = 0

    def read(self, width=1, signed=False):
        if not 0 <= width <= 64 or self.position + width > len(self.values) * 6:
            raise ValueError("Invalid Sina daily bit stream")
        value = 0
        for bit in range(width):
            index, shift = divmod(self.position, 6)
            value |= ((self.values[index] >> shift) & 1) << bit
            self.position += 1
        return value - (1 << width) if signed and width and value & (1 << (width - 1)) else value

    def delta(self):
        sign, value = self.read(), 1
        while self.read():
            value += 1
            if value > 64:
                raise ValueError("Invalid Sina daily delta")
        return value if sign else -value


def decode_sina_daily(text):
    """Decode Sina's K2 OHLC/volume locally, without executing downloaded JS.

    Format reference: Sina finance SDK utils.util, xh5_S_KLC_D, format 3466.
    Only the quoted data assignment is consumed; JavaScript is never evaluated.
    """
    text = re.sub(r"/\*.*?\*/", "", text, flags=re.DOTALL)
    match = re.fullmatch(r'\s*var\s+[\w$]+\s*=\s*"([A-Za-z0-9+/]+)"\s*;?\s*', text)
    if not match:
        raise ValueError("Invalid Sina daily assignment")
    bits = _SinaDailyBits(match.group(1))
    if bits.read(12) != 3466 or bits.read(6) ^ 63:
        raise ValueError("Unsupported Sina daily format")
    state = dict(b_avp=1, b_ph=0, b_phx=0, b_sep=0, wd=62, d=0,
                 p_p=6, p_v=0, p_a=0, p_e=0, p_t=0,
                 l_o=3, l_h=3, l_l=3, l_c=3, l_v=5, l_a=5, l_e=3, l_t=0,
                 u_p=0, u_v=0, u_a=0)

    def precision(value):
        sign = 1 if value >= 0 else -1
        whole, fraction = divmod(abs(value), 3)
        return sign * (whole + (fraction == 1)), sign * (whole + (fraction == 2))

    def scale(value, old, new=0, *, rounding=False):
        old = precision(old) if isinstance(old, int) else old
        new = precision(new)
        value *= 2 ** (new[0] - old[0]) * 5 ** (new[1] - old[1])
        return round(value) if rounding else value

    def date_delta():
        value = bits.read(3)
        if value == 1:
            state["d"] = bits.read(18, True)
            return 0
        return value or bits.read(6)

    rows, previous_amount = [], 0
    for _ in range(100000):
        day_delta, controls = 1, 0
        if bits.read():
            if bits.read():
                if bits.read():
                    controls += 1
                    was_average = state["b_avp"]
                    if bits.read():
                        for flag in ("b_avp", "b_ph", "b_phx"):
                            state[flag] ^= bits.read()
                        was_separate = state["b_sep"]
                        state["b_sep"] ^= bits.read()
                        if bits.read():
                            state["wd"] = bits.read(7) or 62
                        if was_separate != state["b_sep"]:
                            if was_separate:
                                state["u_p"] = state["u_c"]
                            else:
                                for field in "ohlc":
                                    state["u_" + field] = state["u_p"]
                    for field in "pvaet"[:3 + 2 * state["b_ph"]]:
                        if bits.read():
                            old = state["p_" + field]
                            state["p_" + field] += bits.delta()
                            new = state["p_" + field]
                            state["u_" + field] = scale(state.get("u_" + field, 0), old, new, rounding=True)
                            if state["b_sep"] and field == "p":
                                for price in "ohlc":
                                    state["u_" + price] = scale(state["u_" + price], old, new, rounding=True)
                    if not state["b_avp"] and was_average:
                        state["u_a"] = scale(previous_amount, 0, state["p_a"], rounding=True)
                if bits.read():
                    controls += 1
                    for index in range(7 + state["b_ph"] + state["b_phx"]):
                        if bits.read():
                            if index == 6:
                                day_delta = date_delta()
                            else:
                                state["l_" + "ohlcva*et"[index]] += bits.delta()
                if bits.read():
                    controls += 1
                    width = state["l_o"] + (bits.delta() if bits.read() else 0)
                    change = bits.read(3 * width, True)
                    if not state["b_sep"]:
                        state["u_p"] += change
                if not controls:
                    return _ordered(rows)
            elif bits.read():
                if bits.read():
                    if bits.read():
                        day_delta = date_delta()
                    else:
                        state["l_v"] += bits.delta()
                elif state["b_ph"] and bits.read():
                    field = "t" if state["b_phx"] and bits.read() else "e"
                    state["l_" + field] += bits.delta()
                else:
                    state["l_a"] += bits.delta()
            else:
                state["l_" + "ohlc"[bits.read(2)]] += bits.delta()
        values = {}
        for index, field in enumerate("ohlcvaet"[:6 + state["b_ph"] + state["b_phx"]]):
            values[field] = bits.read(3 * state["l_" + field], bool((191 if state["b_sep"] else 185) >> index & 1))
        for _ in range(day_delta):
            state["d"] += 1
            while not state["wd"] & (1 << ((state["d"] % 7 + 10) % 7)):
                state["d"] += 1
        day = (datetime(1970, 1, 1) + timedelta(days=7657 + state["d"])).strftime("%Y-%m-%d")
        if state["b_sep"]:
            prices = []
            for field in "ohlc":
                state["u_" + field] += values[field]
                prices.append(scale(state["u_" + field], state["p_p"]))
            implied = sum(state["u_" + field] for field in "ohlc") / 4
        else:
            opening = state["u_p"] + values["o"]
            state["u_p"] = opening + values["c"]
            prices = [scale(value, state["p_p"]) for value in (
                opening, opening + values["h"], opening - values["l"], state["u_p"])]
            implied = opening + (values["h"] - values["l"] + values["c"]) / 4
        state["u_v"] += values["v"]
        if state["b_avp"]:
            price_precision, volume_precision = precision(state["p_p"]), precision(state["p_v"])
            combined = tuple(a + b for a, b in zip(price_precision, volume_precision))
            amount = scale(math.floor(implied * state["u_v"] + .5), combined, state["p_a"], rounding=True)
            previous_amount = scale(amount + values["a"], state["p_a"])
        else:
            state["u_a"] += values["a"]
            previous_amount = scale(state["u_a"], state["p_a"])
        try:
            if any(price <= 0 for price in prices):
                continue
            rows.append(make_bar(day, *prices, scale(state["u_v"], state["p_v"]), previous_amount))
        except (ValueError, TypeError):
            # Old source records can have missing opens or inconsistent OHLC.
            # Keep decoding subsequent deltas; never repair prices by guessing.
            continue
    raise ValueError("Sina daily history exceeds limit")


def _sina_payload(text):
    text = re.sub(r"^\s*/\*.*?\*/\s*", "", text, flags=re.DOTALL).strip()
    if not text.startswith(("[", "{", '"', "null")):
        match = re.search(r'(?:=|\()\s*([\[{"]|null\b)', text)
        if not match:
            raise ValueError("Invalid Sina history response")
        text = text[match.start(1):]
    value, end = json.JSONDecoder().raw_decode(text)
    if text[end:].strip().strip(");").strip():
        raise ValueError("Invalid Sina history suffix")
    return value


def _parse_us_minutes(text):
    payload = _sina_payload(text)
    if not isinstance(payload, str):
        raise ValueError("Invalid Sina US minutes")
    bars = []
    for session in payload.split():
        day = None
        for index, row in enumerate(session.split(";")):
            fields = row.split(",")
            if index == 0:
                day, fields = fields[0], fields[1:]
                datetime.strptime(day, "%Y-%m-%d")
            if len(fields) < 4 or not re.fullmatch(r"\d{2}:\d{2}:\d{2}", fields[0]):
                raise ValueError("Invalid Sina US minute row")
            price, volume, average = _number(fields[3]), _number(fields[1]), _number(fields[2])
            bars.append(make_bar(day + " " + fields[0], price, price, price, price, volume,
                                 average=average if average and average > 0 else None))
    return _ordered(bars)


def request_sina_bars(instrument, view):
    market, code = instrument.get("market", ""), str(instrument.get("code", ""))
    daily = view == "daily"
    futures = instrument.get("type") == "期" or not market
    if futures:
        method = "getDailyKLine" if daily else "getFewMinLine"
        url = "https://stock2.finance.sina.com.cn/futures/api/jsonp.php/=/InnerFuturesNewService." + method
        params = {"symbol": code.upper(), "type": market_now(instrument).strftime("%Y_%m_%d") if daily else "1"}
    elif market in {"sh", "sz", "bj"}:
        url = "https://quotes.sina.cn/cn/api/jsonp_v2.php/=/CN_MarketDataService.getKLineData"
        params = {"symbol": _sina_code(instrument), "scale": 240 if daily else 1,
                  "ma": "no", "datalen": DAILY_HISTORY if daily else 1970}
    elif market in {"hk", "us"} and daily:
        symbol = code.upper().replace("_", ".")
        if market == "us" and instrument.get("type") == "指":
            symbol = "." + symbol.lstrip(".")
        url = (f"https://finance.sina.com.cn/stock/hkstock/{symbol}/klc2_kl.js" if market == "hk"
               else f"https://finance.sina.com.cn/staticdata/us/{symbol}")
        params = {}
    elif market == "us":
        symbol = code.upper().replace("_", ".")
        if instrument.get("type") == "指":
            symbol = "." + symbol.lstrip(".")
        url = "https://stock.finance.sina.com.cn/usstock/api/jsonp_v2.php/=/US_MinlineNService.getMinline"
        params = {"symbol": symbol, "day": 5}
    elif market == "hk":
        url = "https://stock.finance.sina.com.cn/hkstock/api/openapi.php/HK_StockService.getHKMinline"
        params = {"symbol": code.upper(), "callback": "sw"}
    elif market == "gb" and daily:
        url = "https://gi.finance.sina.com.cn/hq/daily"
        params = {"symbol": code.upper(), "num": DAILY_HISTORY}
    else:
        return ()
    response = requests.get(url, params=params, headers=_SINA_HEADERS, timeout=5)
    try:
        response.raise_for_status()
        text = response.text
    finally:
        response.close()
    if market in {"hk", "us"} and daily:
        bars = decode_sina_daily(text)
    elif market == "us":
        bars = _parse_us_minutes(text)
    elif market == "hk":
        payload = _sina_payload(text)
        result = payload["result"]
        if result.get("status", {}).get("code") != 0 or not isinstance(result.get("data"), list):
            raise ValueError("Invalid Sina HK minutes")
        quote = request_quote({"hk" + code.lower(): instrument}).get("hk" + code.lower(), {})
        day = str(quote.get("date", "")).replace("/", "-")
        datetime.strptime(day, "%Y-%m-%d")
        bars = _ordered(make_bar(day + " " + row["m"], *([row["p"]] * 4), row["v"], row.get("a"),
                                 row.get("avg_p")) for row in result["data"])
    elif market == "gb":
        payload = _sina_payload(text)
        bars = parse_sina(json.dumps(payload["result"]["data"]), daily=True)
    else:
        bars = parse_sina(text, daily=daily)
    return select_bars(bars, view, instrument)


def fetch_bars(instrument, view):
    if view not in {"intraday", "five_day", "daily"}:
        raise ValueError("Unknown history view")
    if instrument.get("market") == "gb" and view != "daily":
        return BarResult(message="新浪暂未提供该指数的分钟历史")
    try:
        bars = request_sina_bars(instrument, view)
        if bars:
            dates = {trading_date(bar, instrument) for bar in bars}
            message = f"接口仅返回 {len(dates)} 个交易日" if view == "five_day" and len(dates) < 5 else ""
            return BarResult(tuple(bars), "sina", message)
    except requests.RequestException as error:
        return BarResult(message="暂无可用历史数据 · 新浪：" + request_error_message(error))
    except (ValueError, TypeError, KeyError, AttributeError, IndexError, OverflowError):
        return BarResult(message="暂无可用历史数据 · 新浪：历史数据格式不可用")
    return BarResult(message="暂无可用历史数据")
