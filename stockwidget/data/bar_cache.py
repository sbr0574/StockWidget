"""Versioned history cache; minute views share one five-session download.

Exchange clocks include US DST and futures night sessions. This deliberately
uses weekday/session boundaries rather than pretending to have a holiday
calendar: holidays can cause one extra refresh, never indefinite stale data.
"""

from contextlib import closing
from dataclasses import replace
from datetime import datetime, timedelta, timezone
import hashlib
import json
import math
from pathlib import Path
import re
import sqlite3
import threading

from stockwidget.core.config_store import history_cache_dir
from stockwidget.data.bars import BarResult, fetch_bars, history_day, make_bar, market_now, select_bars, trading_date


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
        self._generation = 0
        self.database = self.directory / "history.sqlite3"

    @property
    def generation(self):
        with self._lock:
            return self._generation

    @staticmethod
    def key(instrument, view, source):
        identity = {name: str(instrument.get(name, "")).strip().lower()
                    for name in ("market", "code", "type")}
        identity.update(version=2, source=source, view="daily" if view == "daily" else "five_day",
                        adjustment="none", daily_count=100)
        return hashlib.sha256(json.dumps(identity, sort_keys=True).encode()).hexdigest()

    def _connect(self):
        self.directory.mkdir(parents=True, exist_ok=True)
        for attempt in range(2):
            db = sqlite3.connect(self.database, timeout=.25)
            try:
                db.execute("PRAGMA foreign_keys = ON")
                version = db.execute("PRAGMA user_version").fetchone()[0]
                if version not in (0, 1):
                    raise sqlite3.DatabaseError("Unsupported history cache version")
                db.execute("""CREATE TABLE IF NOT EXISTS entries (
                    id INTEGER PRIMARY KEY, cache_key TEXT NOT NULL UNIQUE, saved_at REAL NOT NULL, stamp TEXT NOT NULL,
                    active INTEGER NOT NULL, day TEXT NOT NULL, source TEXT NOT NULL, message TEXT NOT NULL
                )""")
                db.execute("""CREATE TABLE IF NOT EXISTS bars (
                    entry_id INTEGER NOT NULL REFERENCES entries(id) ON DELETE CASCADE,
                    time TEXT NOT NULL, open REAL NOT NULL, high REAL NOT NULL, low REAL NOT NULL,
                    close REAL NOT NULL, volume REAL NOT NULL, amount REAL, average REAL,
                    PRIMARY KEY (entry_id, time)
                ) WITHOUT ROWID""")
                if version == 0:
                    db.execute("PRAGMA user_version = 1")
                return db
            except sqlite3.DatabaseError as error:
                db.close()
                if attempt or getattr(error, "sqlite_errorcode", None) not in {sqlite3.SQLITE_CORRUPT, sqlite3.SQLITE_NOTADB}:
                    raise
                # Only a confirmed corrupt cache is recreated; locked/read-only
                # databases and newer schemas must not be deleted.
                self._remove_database()

    def _remove_database(self):
        for suffix in ("", "-journal", "-wal", "-shm"):
            Path(str(self.database) + suffix).unlink(missing_ok=True)

    def _read(self, key):
        if self.database.exists():
            try:
                with closing(self._connect()) as db, db:
                    row = db.execute("SELECT id, saved_at, stamp, active, day, source, message FROM entries WHERE cache_key = ?",
                                     (key,)).fetchone()
                    if row:
                        payload = dict(zip(("saved_at", "stamp", "active", "day", "source", "message"), row[1:]))
                        bars = tuple(make_bar(*values) for values in db.execute(
                            "SELECT time, open, high, low, close, volume, amount, average FROM bars WHERE entry_id = ? ORDER BY time",
                            (row[0],)))
                        if not math.isfinite(payload["saved_at"]):
                            raise ValueError("Invalid cache timestamp")
                        return payload, BarResult(bars, payload["source"], payload["message"], cached=True)
            except (OSError, sqlite3.Error, ValueError, TypeError) as error:
                if getattr(error, "sqlite_errorcode", None) in {sqlite3.SQLITE_CORRUPT, sqlite3.SQLITE_NOTADB}:
                    try:
                        self._remove_database()
                    except OSError:
                        pass
        return None, None

    def _write(self, key, payload, bars):
        try:
            rows = [(bar.time, bar.open, bar.high, bar.low, bar.close, bar.volume, bar.amount, bar.average)
                    for bar in bars]
            if not math.isfinite(payload["saved_at"]) or any(
                    value is not None and not math.isfinite(value) for row in rows for value in row[1:]):
                return False
            with closing(self._connect()) as db, db:
                db.execute("DELETE FROM entries WHERE cache_key = ?", (key,))
                entry = db.execute("""INSERT INTO entries (cache_key, saved_at, stamp, active, day, source, message)
                                      VALUES (?, ?, ?, ?, ?, ?, ?)""",
                                   (key, payload["saved_at"], payload["stamp"], payload["active"],
                                    payload.get("day", ""), payload["source"], payload.get("message", "")))
                db.executemany("INSERT INTO bars VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                               ((entry.lastrowid, *row) for row in rows))
            return True
        except (OSError, sqlite3.Error, ValueError, TypeError, KeyError):
            return False  # Failed transactions retain the previous cache.

    def clear(self):
        """Clear only history files, and invalidate downloads already in flight."""
        with self._lock:
            self._generation += 1
            self._retry_after.clear()
            try:
                self._remove_database()
                if self.directory.exists():
                    for path in self.directory.iterdir():
                        if re.fullmatch(r"[0-9a-f]{64}(?:\.json|.*\.tmp)", path.name):
                            path.unlink()
                return True
            except OSError:
                return False

    def get(self, instrument, view, source, *, repair=False, generation=None):
        if view not in {"intraday", "five_day", "daily"}:
            raise ValueError("Unknown history view")
        with self._lock:
            if generation is not None and generation != self._generation:
                return BarResult()
            generation = self._generation
            now = self.clock()
            stamp, active = cache_state(instrument, now)
            key = self.key(instrument, view, source)
            payload, cached = self._read(key)
            if source == "eastmoney" and cached and cached.bars and cached.source != source:
                # Older versions could save a Sina fallback under an Eastmoney key.
                payload = cached = None
            ttl = 86400 if view == "daily" else 60 if active else 43200
            if payload and not cached.bars:
                ttl = 300  # Temporary outages/unsupported responses aren't permanent.
            if payload:
                age = now.timestamp() - float(payload.get("saved_at", 0))
                valid = (payload.get("day") == history_day(instrument, now) if view == "daily"
                         else payload.get("stamp") == stamp and payload.get("active") == active)
                if not repair and valid and age >= 0 and ((view == "daily" and cached.bars) or age < ttl):
                    return replace(cached, bars=select_bars(cached.bars, view, instrument))
                if cached.bars and now.timestamp() < self._retry_after.get(key, 0):
                    return replace(cached, bars=select_bars(cached.bars, view, instrument), stale=True,
                                   message="缓存数据，稍后重试更新")
        request_view = "daily" if view == "daily" else "five_day"
        result = self.fetcher(instrument, request_view, source)
        with self._lock:
            if generation != self._generation:
                return result  # A clear must not be undone by a late download.
            if not result.bars and cached and cached.bars:
                self._retry_after[key] = now.timestamp() + 300
                # Do not overwrite a good download with an outage or empty response.
                return replace(cached, bars=select_bars(cached.bars, view, instrument), stale=True,
                               message="缓存数据，更新失败 · " + result.message)
            self._retry_after.pop(key, None)
            if request_view == "five_day" and cached and cached.bars and cached.source == result.source:
                # Especially useful for short futures minute windows: accumulate
                # observed sessions without mixing providers or adjustment bases.
                merged = {bar.time: bar for bar in (*cached.bars, *result.bars)}
                bars = select_bars(tuple(sorted(merged.values(), key=lambda bar: bar.time)), "five_day", instrument)
                result = replace(result, bars=bars)
                if len({trading_date(bar, instrument) for bar in bars}) >= 5 and result.message.startswith("接口仅返回"):
                    result = replace(result, message="")
            self._write(key, {"saved_at": now.timestamp(), "stamp": stamp, "active": active,
                               "day": history_day(instrument, now),
                               "source": result.source, "message": result.message}, result.bars)
            return replace(result, bars=select_bars(result.bars, view, instrument))
