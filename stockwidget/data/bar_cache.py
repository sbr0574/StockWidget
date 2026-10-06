"""Versioned history cache; minute views share one five-session download.

Exchange clocks include US DST and futures night sessions. This deliberately
uses weekday/session boundaries rather than pretending to have a holiday
calendar: holidays can cause one extra refresh, never indefinite stale data.
"""

from dataclasses import asdict, replace
from datetime import datetime, timedelta, timezone
import hashlib
import json
from pathlib import Path
import tempfile
import os
import threading

from stockwidget.core.config_store import history_cache_dir
from stockwidget.data.bars import BarResult, fetch_bars, make_bar, market_now, select_bars, trading_date


def cache_state(instrument, now):
    local = market_now(instrument, now)
    day = local.date()
    minutes = local.hour * 60 + local.minute
    futures = instrument.get("type") == "期" or not instrument.get("market")
    if futures:
        active = ((local.weekday() < 5 and (540 <= minutes < 920 or minutes >= 1260))
                  or (local.weekday() in {1, 2, 3, 4, 5} and minutes < 160))
        phase = "night" if minutes >= 1260 or minutes < 160 else "day" if minutes >= 540 else "before"
    else:
        opening = 570
        closing = 970 if instrument.get("market") in {"hk", "us"} else 910
        active = local.weekday() < 5 and opening <= minutes < closing
        phase = "before" if minutes < opening else "open" if minutes < closing else "closed"
    # Unknown global exchanges: poll conservatively, do not freeze a guessed close.
    if instrument.get("market") == "gb":
        active, phase = True, "unknown"
    if local.weekday() >= 5 and not active:
        while day.weekday() >= 5:
            day -= timedelta(days=1)
        phase = "closed"
    return f"{day.isoformat()}:{phase}", active


class BarCache:
    def __init__(self, directory=None, *, clock=None, fetcher=None):
        self.directory = Path(directory or history_cache_dir())
        self.clock = clock or (lambda: datetime.now(timezone.utc))
        self.fetcher = fetcher or fetch_bars
        self._lock = threading.Lock()
        self._retry_after = {}

    @staticmethod
    def key(instrument, view, source):
        identity = {name: str(instrument.get(name, "")).strip().lower()
                    for name in ("market", "code", "type")}
        identity.update(version=1, source=source, view="daily" if view == "daily" else "five_day",
                        adjustment="none", daily_count=100)
        return hashlib.sha256(json.dumps(identity, sort_keys=True).encode()).hexdigest()

    def _read(self, path):
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
            result = BarResult(tuple(make_bar(row["time"], row["open"], row["high"], row["low"], row["close"],
                                              row.get("volume", 0), row.get("amount"), row.get("average"))
                                     for row in payload["bars"]),
                               payload["source"], payload.get("message", ""), cached=True)
            payload["saved_at"] = float(payload["saved_at"])
            return payload, result
        except (OSError, ValueError, KeyError, TypeError):
            return None, None

    def _write(self, path, payload):
        temporary = None
        try:
            self.directory.mkdir(parents=True, exist_ok=True)
            with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=self.directory,
                                             prefix=path.stem, suffix=".tmp", delete=False) as file:
                temporary = file.name
                json.dump(payload, file, ensure_ascii=False, allow_nan=False)
            os.replace(temporary, path)
        except (OSError, ValueError):
            # Disk permissions/full disk must not stop displaying network data.
            pass
        finally:
            if temporary and os.path.exists(temporary):
                try:
                    os.unlink(temporary)
                except OSError:
                    pass

    def get(self, instrument, view, source):
        if view not in {"intraday", "five_day", "daily"}:
            raise ValueError("Unknown history view")
        with self._lock:
            now = self.clock()
            stamp, active = cache_state(instrument, now)
            path = self.directory / (self.key(instrument, view, source) + ".json")
            payload, cached = self._read(path)
            ttl = (300 if view == "daily" else 60) if active else 43200
            if payload and not cached.bars:
                ttl = min(ttl, 300)  # Temporary outages/unsupported responses aren't permanent.
            if payload:
                age = now.timestamp() - float(payload.get("saved_at", 0))
                if payload.get("stamp") == stamp and payload.get("active") == active and 0 <= age < ttl:
                    return replace(cached, bars=select_bars(cached.bars, view, instrument))
                if cached.bars and now.timestamp() < self._retry_after.get(path.stem, 0):
                    return replace(cached, bars=select_bars(cached.bars, view, instrument), stale=True,
                                   message="缓存数据，稍后重试更新")
            request_view = "daily" if view == "daily" else "five_day"
            result = self.fetcher(instrument, request_view, source)
            if not result.bars and cached and cached.bars:
                self._retry_after[path.stem] = now.timestamp() + 300
                # Do not overwrite a good download with an outage or empty response.
                return replace(cached, bars=select_bars(cached.bars, view, instrument), stale=True,
                               message="缓存数据，更新失败 · " + result.message)
            self._retry_after.pop(path.stem, None)
            if request_view == "five_day" and cached and cached.bars and cached.source == result.source:
                # Especially useful for short futures minute windows: accumulate
                # observed sessions without mixing providers or adjustment bases.
                merged = {bar.time: bar for bar in (*cached.bars, *result.bars)}
                bars = select_bars(tuple(sorted(merged.values(), key=lambda bar: bar.time)), "five_day", instrument)
                result = replace(result, bars=bars)
                if len({trading_date(bar, instrument) for bar in bars}) >= 5 and result.message.startswith("接口仅返回"):
                    result = replace(result, message="")
            self._write(path, {"saved_at": now.timestamp(), "stamp": stamp, "active": active,
                               "source": result.source, "message": result.message,
                               "bars": [asdict(bar) for bar in result.bars]})
            return replace(result, bars=select_bars(result.bars, view, instrument))
