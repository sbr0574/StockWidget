"""Daily hiding rules and provider quote times, independent of Qt."""

from datetime import datetime, timedelta, timezone
import math


QUOTE_TIMEZONE = timezone(timedelta(hours=8))  # 新浪日期/时间为北京时间。
MAX_HIDE_TIMES = 3
STALE_SECONDS = 30


def normalize_hide_times(values):
    if not isinstance(values, (list, tuple)):
        return []
    times = []
    for value in values:
        if not isinstance(value, str):
            continue
        try:
            parsed = datetime.strptime(value.strip(), "%H:%M").strftime("%H:%M")
        except ValueError:
            continue
        if parsed not in times:
            times.append(parsed)
        if len(times) == MAX_HIDE_TIMES:
            break
    return sorted(times)


def quote_timestamp(entry):
    """Return an absolute quote time; a missing/invalid time stays unknown."""
    timestamp = entry.get("timestamp")
    if timestamp is not None:
        try:
            value = float(timestamp)
            if math.isfinite(value) and value > 0:
                datetime.fromtimestamp(value, QUOTE_TIMEZONE)
                return value
        except (TypeError, ValueError, OverflowError, OSError):
            pass
        return None
    date = str(entry.get("date", "") or "").strip().replace("/", "-")
    time = str(entry.get("time", "") or "").strip()
    if not date or not time:
        return None
    try:
        value = datetime.fromisoformat(f"{date}T{time}")
        if value.tzinfo is None:
            value = value.replace(tzinfo=QUOTE_TIMEZONE)
        return value.timestamp()
    except (ValueError, OverflowError, OSError):
        return None


def all_quotes_stale(data, expected_codes, now):
    """Only hide when every selected instrument has a known, stale time."""
    if not expected_codes or set(data) != set(expected_codes):
        return False
    for entry in data.values():
        timestamp = quote_timestamp(entry)
        if timestamp is None or now - timestamp <= STALE_SECONDS:
            return False
    return True


def due_hide_times(times, previous, now):
    """Catch delayed timer ticks, without replaying prior days' schedules."""
    due = []
    for value in times:
        hour, minute = map(int, value.split(":"))
        target = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        if previous < target <= now or now.strftime("%H:%M") == value:
            due.append(value)
    return due
