"""Mini chart projection and shared asynchronous history behavior."""

from datetime import datetime, timedelta, timezone
import tempfile
import threading
import time
from unittest.mock import Mock

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QFontMetrics

from stockwidget.core.quote_presentation import MiniChartCell
from stockwidget.data.bar_cache import BarCache
from stockwidget.data.bars import Bar, BarResult
from stockwidget.data.quotes import _new_entry
from stockwidget.ui.floating.taskbar import render_taskbar
from tests.support import SettingsTestCase


class MiniChartTests(SettingsTestCase):
    def make_window(self, quote_time="09:32:00", **options):
        codes = {f"sh{600000 + i}": {"code": str(600000 + i), "market": "sh", "type": "沪", "name": f"股{i}"}
                 for i in range(7)}
        dialog, window = self._make_dialog(
            codes=codes, watchlist={key: {**value, "checked": True} for key, value in codes.items()},
            visible_metrics=["name", "kline", "price", "change_pct"], **options)
        history = window.history
        history.cache = BarCache(self.enterContext(tempfile.TemporaryDirectory()))
        history.cache.clock = lambda: datetime(2026, 9, 30, 1, 32, tzinfo=timezone.utc)
        start = datetime(2026, 9, 30, 9, 31)
        minutes = tuple(Bar((start + timedelta(minutes=i)).isoformat(" "), 100, 102, 98, 100 + i, 1)
                        for i in range(2))
        history.cache.get = Mock(side_effect=lambda instrument, view, **kwargs: BarResult(minutes, "sina"))
        data = {key: _new_entry(value["name"], 100, 100, 101 + i, 103 + i, 98 - i, 100, 10000,
                               [0] * 5, [0] * 5, [0] * 5, [0] * 5, "2026-09-30", quote_time)
                for i, (key, value) in enumerate(codes.items())}
        window.quotes.accept_result((True, data, None))
        window.show()
        return dialog, window

    def wait_until(self, predicate):
        deadline = time.monotonic() + 3
        while not predicate() and time.monotonic() < deadline:
            self.app.processEvents()
            time.sleep(.005)
        self.assertTrue(predicate())

    def ready(self, history, calls):
        self.wait_until(lambda: history.cache.get.call_count >= calls and not history._busy)
        self.app.processEvents()

    @staticmethod
    def charts(model):
        column = model._headers.index("迷你图")
        return [model.index(row, column).data(Qt.UserRole) for row in range(model.rowCount())]

    def test_amplitude_limits_follow_visible_page_split_and_numeric_sort(self):
        _, window = self.make_window(mini_chart_type="amplitude", float_paging_enabled=True,
                                     float_max_rows=2, float_split_enabled=True)
        left, right = self.charts(window.model), self.charts(window.right_model)
        for cells, expected in ((left, (-3, 4)), (right, (-5, 6))):
            self.assertEqual(len(cells), 2)
            for cell in cells:
                self.assertIsInstance(cell, MiniChartCell)
                for actual, limit in zip(cell.limits, expected):
                    self.assertAlmostEqual(actual, limit)
        window.quotes.change_page("float", 1)
        for cell in self.charts(window.model):
            self.assertAlmostEqual(cell.limits[0], -7)
            self.assertAlmostEqual(cell.limits[1], 8)
        self.assertEqual(len(self.charts(window.right_model)), 1)
        self.assertAlmostEqual(self.charts(window.right_model)[0].limits[0], -8)
        window.quotes.set_sort("涨幅", Qt.DescendingOrder)
        self.assertEqual(window.quotes.float_page, 0)
        self.assertEqual(window.quotes.instrument_at("float", 0)["code"], "600006")
        self.assertAlmostEqual(self.charts(window.model)[0].limits[0], -8)
        self.assertAlmostEqual(self.charts(window.right_model)[0].limits[0], -6)
        self.assertEqual(window.history.cache.get.call_count, 0)

    def test_mini_trend_bootstrap_shares_popup_and_live_updates_without_polling(self):
        _, window = self.make_window(chart_enabled=True, float_paging_enabled=True, float_max_rows=1)
        window.set_view_options(mini_chart_type="trend")
        self.ready(window.history, 1)
        self.assertEqual(len(self.charts(window.model)[0].points), 2)
        window.history.open_row("float", 0)
        self.app.processEvents()
        self.assertEqual(window.history.cache.get.call_count, 1)
        data = dict(window.quotes._latest_quotes)
        key = next(iter(data))
        data[key] = {**data[key], "current_price": 102.5, "high_price": 103, "time": "09:32:20"}
        window.quotes.accept_result((True, data, None))
        self.app.processEvents()
        self.assertEqual(self.charts(window.model)[0].points[-1][1], 102.5)
        self.assertEqual(window.history.dialog.chart.bars[-1].close, 102.5)
        self.assertEqual(window.history.cache.get.call_count, 1)
        window.quotes.change_page("float", 1)
        self.ready(window.history, 2)
        self.assertEqual(window.history.cache.get.call_args.args[0]["code"], "600001")

    def test_clear_and_hide_reject_late_minute_bootstrap(self):
        for action in ("clear", "hide"):
            with self.subTest(action=action):
                _, window = self.make_window(float_paging_enabled=True, float_max_rows=1)
                history = window.history
                release = threading.Event()
                started = threading.Event()
                fetch = history.cache.get.side_effect

                def delayed(*args, **kwargs):
                    started.set()
                    release.wait(3)
                    return fetch(*args, **kwargs)

                history.cache.get.side_effect = delayed
                window.set_view_options(mini_chart_type="trend")
                self.wait_until(started.is_set)
                if action == "clear":
                    history.clear_cache()
                else:
                    window.hide_widget()
                release.set()
                self.wait_until(lambda: not history._busy)
                self.assertTrue(all(not series.result.bars for series in history._series.values()))
                self.assertIsNone(history.dialog)

    def test_watchlist_change_ignores_old_trend_and_reuses_new_bootstrap_for_popup(self):
        _, window = self.make_window(chart_enabled=True, float_paging_enabled=True, float_max_rows=1)
        history = window.history
        window.set_view_options(mini_chart_type="trend")
        self.ready(history, 1)
        window.set_watchlist({key: {**value, "name": "更新后的名称"} for key, value in window.watchlist.items()})
        self.assertEqual(self.charts(window.model)[0].points, ())
        data = {key: _new_entry("新源", 100, 100, 106, 108, 98, 1, 100,
                         [0] * 5, [0] * 5, [0] * 5, [0] * 5, "2026-09-30", "09:32:00")
                for key in window.checked_codes}
        window.quotes.accept_result((True, data, None))
        self.ready(history, 2)
        self.assertEqual(history.cache.get.call_args.args[1], "five_day")
        self.assertEqual(self.charts(window.model)[0].points[-1][1], 106)
        window.history.open_row("float", 0)
        self.app.processEvents()
        self.assertEqual(history.cache.get.call_count, 2)

    def test_visible_fallback_float_requests_trends_with_taskbar_only_display_mode(self):
        _, window = self.make_window(taskbar_enabled=True, taskbar_sync_metrics=False, taskbar_metrics=["name", "price"],
                                     float_paging_enabled=True, float_max_rows=1)
        window.display_mode = "taskbar"
        window.hide()
        window.set_view_options(mini_chart_type="trend")
        self.app.processEvents()
        self.assertEqual(window.history.cache.get.call_count, 0)
        window.show()  # Native embedding failed and the float was shown.
        self.ready(window.history, 1)
        self.assertEqual(window.display_mode, "taskbar")
        self.assertTrue(self.charts(window.model)[0].points)

    def test_previous_session_trend_uses_that_sessions_reference_before_open(self):
        _, window = self.make_window(quote_time="09:25:00", float_paging_enabled=True, float_max_rows=1)
        history = window.history
        history.cache.clock = lambda: datetime(2026, 9, 30, 1, 25, tzinfo=timezone.utc)
        history.cache.get.side_effect = None
        history.cache.get.return_value = BarResult((Bar("2026-09-28 15:00:00", 80, 80, 80, 80),
                                                    Bar("2026-09-29 15:00:00", 95, 100, 95, 100)), "sina")
        data = {key: {**quote, "opening_price": 0, "time": "09:25:00"}
                for key, quote in window.quotes._latest_quotes.items()}
        window.quotes.accept_result((True, data, None))
        window.set_view_options(mini_chart_type="trend")
        self.ready(history, 1)
        chart = self.charts(window.model)[0]
        self.assertEqual(chart.reference, 80)
        self.assertEqual(chart.points[-1][1], 100)

    def test_mini_repair_preserves_open_popups_five_day_view(self):
        _, window = self.make_window(chart_enabled=True, float_paging_enabled=True, float_max_rows=1)
        window.set_view_options(mini_chart_type="trend")
        history = window.history
        self.ready(history, 1)
        history.open_row("float", 0)
        history.dialog.set_view("five_day")
        series = history._series_for(history.instrument, "five_day")
        series.repair_needed = True
        series.retry_after = 0
        history._refresh_minis()
        self.ready(history, 2)
        self.assertEqual(history.dialog.view, "five_day")
        self.assertEqual(history.dialog.chart.view, "five_day")

    def test_taskbar_keeps_ten_character_width_and_its_own_split_ranges(self):
        _, window = self.make_window(mini_chart_type="amplitude", taskbar_sync_metrics=False,
                                     taskbar_metrics=["kline"], taskbar_sync_split=False,
                                     taskbar_split_enabled=True, taskbar_sync_paging=False, taskbar_rows=2)
        cells = self.charts(window.taskbar_model)
        self.assertAlmostEqual(cells[0].limits[0], -3)
        self.assertAlmostEqual(cells[2].limits[0], -5)
        font, *_ = window.get_taskbar_appearance()
        font = QFont(font)
        from stockwidget.core.view_options import taskbar_font_pixels
        for dpi in (96, 144, 192):
            height = round(44 * dpi / 96)
            font.setPixelSize(taskbar_font_pixels(height, dpi, 2, window.font.pointSizeF()))
            image = render_taskbar(window, height, dpi=dpi, max_width=80)
            expected = QFontMetrics(font).horizontalAdvance("0" * 10) * 2 + round(9 * dpi / 96)
            self.assertEqual(image.width(), expected)
