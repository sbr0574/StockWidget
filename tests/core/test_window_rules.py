"""屏幕几何、定时隐藏和行情时间的纯规则。"""

from datetime import datetime, timedelta
import os
import unittest

from stockwidget.core.window_rules import (
    best_screen,
    clamp_point,
    resolve_restore_position,
    screen_containing,
    QUOTE_TIMEZONE,
    all_quotes_stale,
    due_hide_times,
    normalize_hide_times,
    quote_timestamp,
)


PRIMARY = (0, 0, 1920, 1080)


SECONDARY = (-1920, 0, 1920, 1080)


RECTS = [PRIMARY, SECONDARY]


WIDTH = 200


HEIGHT = 100


class ScreenContainingTests(unittest.TestCase):
    def test_primary(self):
        self.assertEqual(screen_containing(100, 100, RECTS), PRIMARY)

    def test_secondary_negative_coords(self):
        self.assertEqual(screen_containing(-100, 100, RECTS), SECONDARY)

    def test_none_when_outside(self):
        self.assertIsNone(screen_containing(99999, 99999, RECTS))


class ClampPointTests(unittest.TestCase):
    def test_inside_unchanged(self):
        self.assertEqual(clamp_point(100, 100, PRIMARY, WIDTH, HEIGHT), (100, 100))

    def test_outside_right_bottom(self):
        self.assertEqual(clamp_point(5000, 5000, PRIMARY, WIDTH, HEIGHT), (1720, 980))

    def test_outside_left_top(self):
        self.assertEqual(clamp_point(-5000, -5000, PRIMARY, WIDTH, HEIGHT), (0, 0))


class BestScreenTests(unittest.TestCase):
    def test_largest_overlap_allows_crossing_monitor_seams(self):
        self.assertEqual(best_screen(-80, 200, 200, 100, RECTS), PRIMARY)
        self.assertEqual(best_screen(-180, 200, 200, 100, RECTS), SECONDARY)

    def test_outside_uses_nearest_monitor_including_negative_coordinates(self):
        self.assertEqual(best_screen(-5000, 200, 200, 100, RECTS), SECONDARY)
        self.assertEqual(best_screen(3000, 200, 200, 100, RECTS), PRIMARY)

    def test_staggered_monitors_and_empty_screen_list(self):
        upper = (0, -1080, 1920, 1080)
        self.assertEqual(best_screen(200, -100, 200, 150, [PRIMARY, upper]), upper)
        self.assertIsNone(best_screen(0, 0, 200, 100, []))


class ResolveRestorePositionTests(unittest.TestCase):
    def test_restore_on_secondary(self):
        self.assertEqual(
            resolve_restore_position((-1500, 200), RECTS, PRIMARY, WIDTH, HEIGHT),
            (-1500, 200),
        )

    def test_restore_on_primary(self):
        self.assertEqual(
            resolve_restore_position((500, 300), RECTS, PRIMARY, WIDTH, HEIGHT),
            (500, 300),
        )

    def test_secondary_disconnected_falls_back_to_primary(self):
        self.assertEqual(
            resolve_restore_position((-1500, 200), [PRIMARY], PRIMARY, WIDTH, HEIGHT),
            (0, 200),
        )

    def test_clamp_into_secondary_when_near_edge(self):
        self.assertEqual(
            resolve_restore_position((-150, 200), RECTS, PRIMARY, WIDTH, HEIGHT),
            (-200, 200),
        )

    def test_none_saved_uses_primary_default(self):
        self.assertEqual(
            resolve_restore_position(None, RECTS, PRIMARY, WIDTH, HEIGHT),
            (1680, 900),
        )


os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


NOW = datetime(2026, 9, 30, 15, 0, 0, tzinfo=QUOTE_TIMEZONE)


CODES = {"sh600000": {"market": "sh", "code": "600000", "checked": True},
         "sz000001": {"market": "sz", "code": "000001", "checked": True}}


def entry(age=31, *, timestamp=False):
    updated = NOW - timedelta(seconds=age)
    result = {"name": "测试", "current_price": 10, "prev_close": 9,
              "opening_price": 9, "high_price": 11, "low_price": 8,
              "deals_vol": 100, "deals_amt": 1000,
              "purchaser_price": [9] * 5, "seller_price": [10] * 5,
              "purchaser_vol": [100] * 5, "seller_vol": [100] * 5}
    if timestamp:
        result["timestamp"] = updated.timestamp()
    else:
        result.update(date=updated.strftime("%Y/%m/%d"), time=updated.strftime("%H:%M:%S"))
    return result


class HideRuleTests(unittest.TestCase):
    def test_normalize_unique_minutes_and_limit_three(self):
        self.assertEqual(normalize_hide_times([None, "25:00", "9:05", "09:05", "15:00",
                                              "00:00", "20:00", "10:00:01"]),
                         ["00:00", "09:05", "15:00"])
        for invalid in (None, "15:00", 15, {}):
            self.assertEqual(normalize_hide_times(invalid), [])

    def test_sina_beijing_time_and_eastmoney_epoch_ignore_computer_timezone(self):
        self.assertEqual(quote_timestamp(entry()), NOW.timestamp() - 31)
        self.assertEqual(quote_timestamp(entry(timestamp=True)), NOW.timestamp() - 31)
        self.assertEqual(quote_timestamp({"date": "2026-09-30", "time": "07:00:00+00:00"}),
                         NOW.timestamp())
        for value in (None, 0, "-", "bad", float("nan"), float("inf"), 1e99):
            self.assertIsNone(quote_timestamp({"timestamp": value}))
        for value in ({}, {"date": "2026-09-30"}, {"date": "bad", "time": "15:00:00"}):
            self.assertIsNone(quote_timestamp(value))

    def test_strict_threshold_all_selected_missing_future_and_previous_day(self):
        for second in (30, 29, -10):
            self.assertFalse(all_quotes_stale({"a": entry(), "b": entry(second)}, ["a", "b"],
                                             NOW.timestamp()))
        self.assertTrue(all_quotes_stale({"a": entry(31), "b": entry(86400, timestamp=True)},
                                        ["a", "b"], NOW.timestamp()))
        for data in ({}, {"a": entry()}, {"a": entry(), "b": {}},
                     {"a": entry(), "b": {"date": "2026-09-30", "time": "invalid"}}):
            self.assertFalse(all_quotes_stale(data, ["a", "b"], NOW.timestamp()))
        self.assertFalse(all_quotes_stale({}, [], NOW.timestamp()))

    def test_schedule_minute_delayed_ticks_midnight_and_clock_backwards(self):
        self.assertEqual(due_hide_times(["15:00"], NOW - timedelta(seconds=1), NOW), ["15:00"])
        self.assertEqual(due_hide_times(["15:00"], NOW, NOW + timedelta(seconds=59)), ["15:00"])
        self.assertEqual(due_hide_times(["15:00"], NOW - timedelta(seconds=1),
                                       NOW + timedelta(minutes=2)), ["15:00"])
        midnight = NOW.replace(hour=0)
        self.assertEqual(due_hide_times(["00:00", "23:59"], midnight - timedelta(seconds=1),
                                       midnight), ["00:00"])
        self.assertEqual(due_hide_times(["15:00"], NOW + timedelta(hours=1),
                                       NOW - timedelta(hours=1)), [])
