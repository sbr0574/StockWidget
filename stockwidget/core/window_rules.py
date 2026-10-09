"""窗口位置、屏幕边界、定时隐藏与行情过期判断；不依赖 Qt。"""

from datetime import datetime, timedelta, timezone
import math


def screen_containing(x, y, rects):
    """返回包含点 (x, y) 的第一个屏幕矩形；找不到返回 None。"""
    for r in rects:
        left, top, w, h = r
        if left <= x < left + w and top <= y < top + h:
            return r
    return None


def clamp_point(x, y, rect, widget_w=0, widget_h=0):
    """把左上角 (x, y) 夹取到 rect 内，保证 widget 不超出该屏幕。"""
    left, top, w, h = rect
    right = left + w - max(1, widget_w)
    bottom = top + h - max(1, widget_h)
    return max(left, min(x, right)), max(top, min(y, bottom))


def best_screen(x, y, width, height, rects):
    """优先选择窗口重叠最多的屏幕，屏外位置选择调整距离最短的屏幕。"""
    def score(rect):
        left, top, w, h = rect
        overlap = (max(0, min(x + width, left + w) - max(x, left))
                   * max(0, min(y + height, top + h) - max(y, top)))
        cx, cy = clamp_point(x, y, rect, width, height)
        return overlap, -((cx - x) ** 2 + (cy - y) ** 2)

    return max(rects, key=score) if rects else None


def adjacent_popup_size(anchor, bounds, preferred, minimum, gap=8):
    """小屏幕优先缩小图表，让它仍能放在行情旁；保留可读的最小尺寸。"""
    left, top, anchor_w, anchor_h = anchor
    screen_x, screen_y, screen_w, screen_h = bounds
    width, height = min(preferred[0], screen_w), min(preferred[1], screen_h)
    candidates = (
        (min(width, screen_x + screen_w - left - anchor_w - gap), height),
        (min(width, left - screen_x - gap), height),
        (width, min(height, screen_y + screen_h - top - anchor_h - gap)),
        (width, min(height, top - screen_y - gap)),
    )
    viable = [size for size in candidates if size[0] >= minimum[0] and size[1] >= minimum[1]]
    return max(viable, key=lambda size: size[0] * size[1]) if viable else (width, height)


def adjacent_popup_position(anchor, bounds, width, height, *, prefer_above=False, gap=8):
    """把图表放在行情旁，空间不足时换边并限制在当前屏幕可用区域。"""
    left, top, anchor_w, anchor_h = anchor
    right = (left + anchor_w + gap, top)
    left_side = (left - width - gap, top)
    below = (left, top + anchor_h + gap)
    above = (left, top - height - gap)
    candidates = (above, below, right, left_side) if prefer_above else (right, left_side, below, above)
    for position in candidates:
        x, y = clamp_point(*position, bounds, width, height)
        if ((position == right and x >= right[0])
                or (position == left_side and x <= left_side[0])
                or (position == below and y >= below[1])
                or (position == above and y <= above[1])):
            return x, y
    return clamp_point(*candidates[0], bounds, width, height)


def resolve_restore_position(saved, rects, primary, widget_w=0, widget_h=0,
                             margin_x=40, margin_y=80):
    """根据保存位置决定窗口恢复位置。

    - ``saved``: (x, y) 或 None（无保存位置）。
    - ``rects``: 所有屏幕可用区域 [(left, top, w, h), ...]。
    - ``primary``: 主屏可用区域 (left, top, w, h)。

    规则：保存位置落在任一屏幕内则原位恢复（并夹取到该屏幕内）；
    否则（如副屏已断开）回退到主屏右下角默认位置。
    """
    if saved is not None:
        x, y = int(saved[0]), int(saved[1])
        target = screen_containing(x, y, rects) or primary
        return clamp_point(x, y, target, widget_w, widget_h)

    left, top, w, h = primary
    x = left + w - widget_w - margin_x
    y = top + h - widget_h - margin_y
    return max(left, x), max(top, y)


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
