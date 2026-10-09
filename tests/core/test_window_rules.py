"""屏幕几何、定时隐藏和行情时间的纯规则。"""

from datetime import datetime, timedelta
import unittest

from stockwidget.core.window_rules import (
    adjacent_popup_position,
    adjacent_popup_size,
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
    def test_primary_negative_coordinates_and_outside(self):
        for point, expected in (((100, 100), PRIMARY), ((-100, 100), SECONDARY),
                                ((99999, 99999), None)):
            with self.subTest(point=point):
                self.assertEqual(screen_containing(*point, RECTS), expected)


class ClampPointTests(unittest.TestCase):
    def test_inside_and_opposite_screen_boundaries(self):
        for point, expected in (((100, 100), (100, 100)), ((5000, 5000), (1720, 980)),
                                ((-5000, -5000), (0, 0))):
            with self.subTest(point=point):
                self.assertEqual(clamp_point(*point, PRIMARY, WIDTH, HEIGHT), expected)


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


class AdjacentPopupTests(unittest.TestCase):
    def test_small_screen_shrinks_popup_to_fit_beside_quotes(self):
        self.assertEqual(adjacent_popup_size((30, 30, 167, 138), (0, 0, 640, 376),
                                             (560, 400), (400, 300)), (435, 376))
        self.assertEqual(adjacent_popup_size((100, 100, 200, 100), PRIMARY,
                                             (560, 400), (400, 300)), (560, 400))
        self.assertEqual(adjacent_popup_size((0, 0, 640, 360), (0, 0, 640, 360),
                                             (560, 400), (400, 300)), (560, 360))

    def test_changes_sides_and_clamps_without_covering_anchor(self):
        cases = (
            ((100, 100, 200, 100), PRIMARY, (308, 100)),
            ((1700, 900, 200, 100), PRIMARY, (1132, 680)),
            ((300, 100, 200, 100), (0, 0, 800, 800), (240, 208)),
            ((300, 600, 200, 100), (0, 0, 800, 800), (240, 192)),
            ((-1700, 100, 200, 100), SECONDARY, (-1492, 100)),
        )
        for anchor, bounds, expected in cases:
            with self.subTest(anchor=anchor, bounds=bounds):
                self.assertEqual(adjacent_popup_position(anchor, bounds, 560, 400), expected)

    def test_taskbar_prefers_above_and_stays_on_screen_near_right_edge(self):
        self.assertEqual(adjacent_popup_position((1890, 1060, 1, 1), (0, 0, 1920, 1040),
                                                 560, 400, prefer_above=True), (1360, 640))
        self.assertEqual(adjacent_popup_position((100, 0, 1, 1), PRIMARY,
                                                 560, 400, prefer_above=True), (100, 9))

    def test_no_space_and_offscreen_anchor_keep_popup_accessible(self):
        self.assertEqual(adjacent_popup_position((0, 0, 800, 800), (0, 0, 800, 800), 560, 400), (240, 0))
        x, y = adjacent_popup_position((5000, 5000, 200, 100), PRIMARY, 560, 400)
        self.assertTrue(0 <= x <= 1360 and 0 <= y <= 680)


class ResolveRestorePositionTests(unittest.TestCase):
    def test_restore_clamps_saved_positions_and_handles_disconnected_screens(self):
        cases = (
            ((-1500, 200), RECTS, (-1500, 200)),
            ((500, 300), RECTS, (500, 300)),
            ((-1500, 200), [PRIMARY], (0, 200)),
            ((-150, 200), RECTS, (-200, 200)),
            (None, RECTS, (1680, 900)),
        )
        for saved, screens, expected in cases:
            with self.subTest(saved=saved, screens=screens):
                self.assertEqual(resolve_restore_position(saved, screens, PRIMARY, WIDTH, HEIGHT), expected)


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
