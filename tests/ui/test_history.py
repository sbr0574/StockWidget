"""Settings, row identity, gesture safety, asynchronous selection and plotting."""

from unittest.mock import Mock, patch
from dataclasses import replace
from datetime import datetime, timedelta, timezone
from types import SimpleNamespace
import os
import sys
import time
import unittest

from PySide6.QtCore import QPoint, QRect, Qt
from PySide6.QtGui import QCursor, QPalette
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication

from stockwidget.data.bars import Bar, BarResult
from stockwidget.data.quotes import _new_entry
from stockwidget.ui.floating.taskbar import TaskbarController, render_taskbar
from stockwidget.ui.controls.history_chart import HistoryChart
from tests.data.test_bars import sample_bars
from tests.support import SettingsTestCase


class HistoryTestCase(SettingsTestCase):
    def setUp(self):
        super().setUp()
        self._cursor_position = QCursor.pos()

    def tearDown(self):
        super().tearDown()
        # QTest mouse gestures can move the global cursor even offscreen.
        # Restore it so later hover tests don't inherit a deleted window's hit.
        QCursor.setPos(self._cursor_position)
        self.app.processEvents()

    def wait_until(self, predicate):
        deadline = time.monotonic() + 2
        while not predicate() and time.monotonic() < deadline:
            self.app.processEvents()
            time.sleep(.01)
        self.app.processEvents()
        self.assertTrue(predicate())

    def wait_ready(self, window):
        self.wait_until(lambda: not window.history._busy)

    def make_window(self, **settings):
        codes = {f"sh{600000 + i}": {"code": str(600000 + i), "market": "sh", "type": "沪", "name": f"标的{i}"}
                 for i in range(7)}
        dialog, window = self._make_dialog(watchlist={key: {**value, "checked": True} for key, value in codes.items()},
                                           codes=codes, **settings)
        data = {key: _new_entry(value["name"], 10, 9, i + 10, i + 12, 8, 100, 1000,
                               [0] * 5, [0] * 5, [0] * 5, [0] * 5, "2026-09-30", "10:00:00")
                for i, (key, value) in enumerate(codes.items())}
        window.quotes.accept_result((True, data, None))
        daily = sample_bars()
        shift = datetime(2026, 9, 30) - datetime.fromisoformat(daily[-1].time)
        daily = tuple(replace(bar, time=(datetime.fromisoformat(bar.time) + shift).strftime("%Y-%m-%d")) for bar in daily)
        start = datetime(2026, 9, 30, 9, 31)
        minutes = tuple(Bar((start + timedelta(minutes=i)).isoformat(" "), 10, 12, 8, 10, 3, 30, 10) for i in range(30))
        window.history.cache.get = Mock(side_effect=lambda instrument, view, source, **kwargs:
                                        BarResult(daily if view == "daily" else minutes, source))
        window.history.cache.clock = lambda: datetime(2026, 9, 30, 2, tzinfo=timezone.utc)
        window.show()
        self.app.processEvents()
        return dialog, window


class HistoryUITests(HistoryTestCase):
    def test_chart_settings_use_common_description_colors_and_group_geometry(self):
        dialog, window = self.make_window(chart_enabled=True)
        dialog.show()
        dialog.ui.settings_pages.setCurrentWidget(dialog.ui.data)
        for mode in ("light", "dark"):
            window.set_view_options(color_mode=mode)
            self.app.processEvents()
            group = dialog.ui.gb_chart
            title = dialog.ui.chart_enabled_title
            position = title.mapTo(group, QPoint())
            self.assertGreaterEqual(position.x(), 10)
            self.assertGreaterEqual(position.y(), 25)
            common = dialog.ui.cb_head_description.palette().color(QPalette.WindowText)
            for label in (dialog.ui.chart_enabled_description, dialog.ui.chart_display_mode_description,
                          dialog.ui.chart_indicators_description):
                self.assertEqual(label.palette().color(QPalette.WindowText), common)
                self.assertEqual(label.maximumWidth(), 310)
            self.assertEqual(dialog.ui.data_scroll.horizontalScrollBar().maximum(), 0)

    def test_four_sizes_and_indicator_preferences_preserve_view_without_refetch(self):
        dialog, window = self.make_window(chart_enabled=True)
        window.move(self.app.primaryScreen().availableGeometry().topLeft() + QPoint(20, 20))
        window.history.open_row("float", 0)
        self.wait_ready(window)
        popup = window.history.dialog
        popup.set_view("daily")
        self.wait_ready(window)
        dialog.ui.cb_chart_average.setChecked(False)
        dialog.ui.cb_chart_volume.setChecked(False)
        dialog.ui.cb_chart_ma60.setChecked(False)
        sizes = []
        for mode in ("large", "medium", "small"):
            combo = dialog.ui.cmb_chart_display_mode
            combo.setCurrentIndex(combo.findData(mode))
            self.wait_until(lambda: popup.rect().contains(popup.chart.geometry()))
            sizes.append(popup.size())
            self.assertEqual(popup.view, "daily")
            self.assertFalse(popup.chart.show_average)
            self.assertFalse(popup.chart.show_volume)
            self.assertNotIn(60, popup.chart.periods)
            self.assertTrue(popup.rect().contains(popup.chart.geometry()))
            self.assertTrue(self.app.primaryScreen().availableGeometry().contains(popup.geometry()))
            self.assertTrue(all(popup.rect().contains(button.mapTo(popup, QPoint()) + button.rect().bottomRight())
                                for button in popup.view_buttons.values()))
        self.assertLess(sizes[2].width(), sizes[1].width())
        self.assertLess(sizes[1].width(), sizes[0].width())
        self.assertLessEqual(sizes[2].width(), 320)
        self.assertLessEqual(sizes[2].height(), 230)
        self.assertEqual(window.history.cache.get.call_count, 2)
        cfg = window.current_config()
        self.assertEqual(cfg["chart_display_mode"], "small")
        self.assertEqual(cfg["chart_ma_periods"], [5, 10, 20, 30])

    def test_live_quotes_update_chart_at_quote_cadence_and_repair_only_missing_minutes(self):
        _, window = self.make_window(chart_enabled=True)
        history = window.history
        history.open_row("float", 0)
        self.wait_ready(window)
        data = dict(window.quotes._latest_quotes)
        def tick(time, price):
            data["sh600000"] = {**data["sh600000"], "time": time, "current_price": price}
            window.quotes.accept_result((True, data, None))
            self.app.processEvents()
        for time, price in (("10:00:01", 11), ("10:00:50", 12), ("10:01:10", 10.5)):
            tick(time, price)
            self.assertEqual(history.dialog.chart.bars[-1].close, price)
        self.assertEqual(history.cache.get.call_count, 1)
        history.dialog.set_view("five_day")
        history.dialog.close()
        tick("10:01:30", 11.25)
        history.open_row("float", 0)
        self.wait_ready(window)
        self.assertEqual(history.dialog.chart.bars[-1].close, 11.25)
        self.assertEqual(history.cache.get.call_count, 1)
        tick("10:03:10", 12.25)  # 10:02 wasn't observed: repair the history.
        self.wait_ready(window)
        self.assertEqual(history.cache.get.call_count, 2)
        self.assertTrue(history.cache.get.call_args.kwargs["repair"])
        self.assertEqual(history.dialog.chart.bars[-1].close, 12.25)
        tick("10:03:20", 12.5)  # An incomplete response is throttled.
        self.assertEqual(history.cache.get.call_count, 2)
        self.assertEqual(history.dialog.chart.bars[-1].close, 12.5)

    def test_daily_quotes_update_current_candle_and_first_load_next_day_fetches_once(self):
        _, window = self.make_window(chart_enabled=True)
        history = window.history
        history.cache.get.side_effect = None
        history.cache.get.return_value = BarResult((Bar("2026-09-29", 9, 10, 8, 9),
                                                    Bar("2026-09-30", 10, 12, 8, 10)), "sina")
        history.open_row("float", 0)
        self.wait_ready(window)
        history.dialog.set_view("daily")
        self.wait_ready(window)
        calls = history.cache.get.call_count
        data = dict(window.quotes._latest_quotes)
        data["sh600000"] = {**data["sh600000"], "time": "10:01:20", "current_price": 12.5,
                            "high_price": 12.5, "deals_vol": 120, "deals_amt": 1500}
        window.quotes.accept_result((True, data, None))
        candle = history.dialog.chart.bars[-1]
        self.assertEqual((candle.time, candle.open, candle.high, candle.low, candle.close, candle.volume),
                         ("2026-09-30", 10, 12.5, 8, 12.5, 120))
        history.reload()
        self.assertEqual(history.cache.get.call_count, calls)
        history.cache.clock = lambda: datetime(2026, 10, 1, 2, tzinfo=timezone.utc)
        data["sh600000"] = {**data["sh600000"], "date": "2026-10-01", "time": "10:00:00",
                            "opening_price": 13, "current_price": 13.25, "high_price": 14, "low_price": 12}
        window.quotes.accept_result((True, data, None))
        self.wait_ready(window)
        self.assertEqual(history.cache.get.call_count, calls + 1)
        self.assertEqual(history.dialog.chart.bars[-1].time, "2026-10-01")
        window.quotes.accept_result((True, data, None))
        self.assertEqual(history.cache.get.call_count, calls + 1)

    def test_quotes_arriving_during_bootstrap_are_used_and_failed_old_requests_do_not_update_chart(self):
        import threading
        _, window = self.make_window(chart_enabled=True)
        history = window.history
        release = threading.Event()
        fetch = history.cache.get.side_effect
        def delayed(*args, **kwargs):
            release.wait(2)
            return fetch(*args, **kwargs)
        history.cache.get.side_effect = delayed
        try:
            history.open_row("float", 0)
            data = dict(window.quotes._latest_quotes)
            data["sh600000"] = {**data["sh600000"], "time": "10:00:20", "current_price": 11.125}
            window.quotes.accept_result((True, data, None))
            release.set()
            self.wait_ready(window)
            self.assertEqual(history.dialog.chart.bars[-1].close, 11.125)
            self.assertIn("收 11.12", history.dialog.chart.tooltip_text(len(history.dialog.chart.bars) - 1))
            window.quotes.accept_result((False, {}, "请求失败"))
            window.quotes.invalidate()
            self.assertIsNone(window.quotes.quote_for(history.instrument))
            data["sh600000"] = {**data["sh600000"], "current_price": 99}
            window.quotes.accept_result((True, data, None, window.quotes._quote_generation - 1))
            self.assertEqual(history.dialog.chart.bars[-1].close, 11.125)
            self.assertEqual(history.cache.get.call_count, 1)
        finally:
            release.set()
            self.wait_ready(window)

    def test_chart_tooltip_prices_match_live_quote_precision(self):
        from shiboken6 import delete
        chart = HistoryChart()
        try:
            bars = tuple(replace(bar, open=10.1, high=12.1234, low=9.1, close=11.1234, volume=2_000_000)
                         for bar in sample_bars())
            for market, type_, expected, unit_mode, volume in (
                    ("sh", "沪", "11.12", "cn", "2.00万手"),
                    ("sz", "基", "11.123", "en", "20.00k手"),
                    ("us", "美", "11.123", "auto", "2.00M股"),
                    ("hk", "港", "11.12", "cn", "200.00万股"),
                    ("", "期", "11.12", "en", "2.00M合约")):
                chart.set_options([5, 10, 20, 30, 60], True, True, unit_mode)
                chart.set_data(bars, "daily", {"market": market, "type": type_})
                text = chart.tooltip_text(29)
                self.assertIn(f"收 {expected}", text)
                self.assertIn(f"MA60: {expected}", text)
                self.assertIn(f"成交量 {volume}", text)
                chart.set_options([], False, False)
                self.assertNotIn("MA60", chart.tooltip_text(29))
                self.assertNotIn("成交量", chart.tooltip_text(29))
                chart.set_options([5, 10, 20, 30, 60], True, True)
        finally:
            delete(chart)

    def test_index_zero_axis_and_volume_separator_in_compact_charts(self):
        from shiboken6 import delete

        chart = HistoryChart()
        try:
            chart.setMinimumSize(220, 110)
            for market, code, price in (("sh", "000001", 3842), ("sz", "399001", 12887), ("sz", "399006", 3135)):
                instrument = {"market": market, "code": code, "type": "指"}
                bars = (Bar("2026-09-29 15:00:00", price, price, price, price, 100, 1500, 15),
                        Bar("2026-09-30 09:31:00", price, price + 2, price, price + 2, 150, 2250, 15))
                for view in ("intraday", "five_day"):
                    for width, height in ((220, 110), (308, 154), (408, 224), (548, 324)):
                        with self.subTest(code=code, view=view, size=(width, height)):
                            chart.resize(width, height)
                            chart.set_data(bars[-1:] if view == "intraday" else bars, view, instrument,
                                           reference_price=price)
                            image = chart.grab().toImage()
                            low, high = chart.price_range()
                            self.assertGreater(low, price * .98)
                            self.assertLess(high, price * 1.02)
                            self.assertTrue(chart._plot.top() < chart._reference_y < chart._plot.bottom())
                            self.assertNotIn("均价", chart.tooltip_text(0))
                            # The separator extends into the otherwise empty axis margin.
                            scale = image.devicePixelRatio()
                            y = round((chart._plot.bottom() + chart._volume_plot.top()) / 2 * scale)
                            background = chart.palette().color(QPalette.Base)
                            self.assertTrue(any(image.pixelColor(round(2 * scale), row) != background
                                                for row in range(y - 1, y + 2)))
            chart.set_options([], True, False)
            chart.grab()
            self.assertTrue(chart._volume_plot.isEmpty())
            self.assertIsNotNone(chart._reference_y)
            chart.set_data(sample_bars(), "daily", instrument, reference_price=price)
            chart.grab()
            self.assertIsNone(chart._reference_y)
        finally:
            delete(chart)

    def test_volume_units_update_from_settings_without_fetch_and_normal_footer_is_hidden(self):
        dialog, window = self.make_window(chart_enabled=True, unit_mode="cn")
        history = window.history
        minutes = history.cache.get.side_effect({}, "five_day", "sina").bars
        history.cache.get.side_effect = None
        history.cache.get.return_value = BarResult((replace(minutes[0], volume=2_000_000), *minutes[1:]), "sina")
        history.open_row("float", 0)
        self.wait_ready(window)
        popup = history.dialog
        self.assertTrue(popup.status.isHidden())
        self.assertEqual(popup.chart.reference_price, 9)
        for mode, volume in (("cn", "2.00万手"), ("en", "20.00k手"), ("auto", "2.00万手")):
            with self.subTest(mode=mode):
                dialog.ui.cmb_unit_mode.setCurrentIndex(dialog.ui.cmb_unit_mode.findData(mode))
                self.assertIn(f"成交量 {volume}", popup.chart.tooltip_text(0))
                self.assertEqual(history.cache.get.call_count, 1)
        series = next(iter(history._series.values()))
        series.result = replace(series.result, stale=True, message="历史数据更新失败")
        history.reload()
        self.assertFalse(popup.status.isHidden())
        self.assertIn("更新失败", popup.status.toolTip())

    def test_compact_chart_time_labels_keep_endpoints_without_overlap(self):
        from shiboken6 import delete

        chart = HistoryChart()
        try:
            chart.setMinimumSize(220, 110)
            daily = sample_bars()
            start = datetime(2026, 9, 30, 9, 30)
            minutes = tuple(replace(bar, time=(start + timedelta(minutes=i)).strftime("%Y-%m-%d %H:%M:%S"))
                            for i, bar in enumerate(daily))
            for view in ("intraday", "five_day", "daily"):
                for width in (280, 308, 360, 435, 560):
                    with self.subTest(view=view, width=width):
                        chart.resize(width, 220)
                        chart.set_data(daily if view == "daily" else minutes, view, {"market": "sh", "type": "沪"})
                        self.assertFalse(chart.grab().isNull())
                        labels = chart._time_labels()
                        begin, stop = (5, 10) if view == "daily" else (11, 16) if view == "intraday" else (5, 16)
                        self.assertIn(labels[0][1], (chart.bars[0].time[begin:stop], chart.bars[0].time[5:10]))
                        self.assertIn(labels[-1][1], (chart.bars[-1].time[begin:stop], chart.bars[-1].time[5:10]))
                        for (previous, _), (current, _) in zip(labels, labels[1:]):
                            self.assertGreaterEqual(current.left() - previous.right(), 8)
                        self.assertTrue(all(chart.rect().contains(rect.toAlignedRect()) for rect, _ in labels))
        finally:
            delete(chart)

    def test_display_mode_setting_persists_and_switches_open_chart(self):
        dialog, window = self.make_window()
        combo = dialog.ui.cmb_chart_display_mode
        self.assertEqual(combo.currentData(), "window")
        self.assertFalse(combo.isEnabled())
        dialog.ui.cb_chart_enabled.setChecked(True)
        combo.setCurrentIndex(combo.findData("large"))
        self.assertEqual(window.current_config()["chart_display_mode"], "large")
        window.history.open_row("float", 0)
        self.wait_ready(window)
        popup = window.history.dialog
        popup.set_view("daily")
        self.wait_ready(window)
        dialog.ui.cb_chart_ma60.setChecked(False)
        combo.setCurrentIndex(combo.findData("window"))
        self.wait_ready(window)
        self.assertTrue(popup.isVisible())
        self.assertEqual(popup.view, "daily")
        self.assertNotIn(60, popup.chart.periods)
        self.assertEqual(len(popup.chart.bars), 30)
        combo.setCurrentIndex(combo.findData("large"))
        self.wait_ready(window)
        self.wait_until(lambda: QApplication.activePopupWidget() is popup)
        dialog.ui.cb_chart_enabled.setChecked(False)
        self.assertFalse(popup.isVisible())
        self.assertFalse(combo.isEnabled())
        self.assertEqual(window.current_config()["chart_display_mode"], "large")
        dialog.ui.cb_chart_enabled.setChecked(True)
        self.assertEqual(combo.currentData(), "large")
        window.reset_settings()
        self.assertFalse(dialog.ui.cb_chart_enabled.isChecked())
        self.assertEqual(combo.currentData(), "window")

    def test_floating_view_buttons_outside_click_and_escape(self):
        _, window = self.make_window(chart_enabled=True, chart_display_mode="large")
        window.history.open_row("float", 0)
        self.wait_ready(window)
        popup = window.history.dialog
        QTest.mouseClick(popup.chart, Qt.LeftButton, pos=popup.chart.rect().center())
        self.assertTrue(popup.isVisible())
        QTest.mouseClick(popup.view_buttons["daily"], Qt.LeftButton)
        self.wait_ready(window)
        self.assertTrue(popup.isVisible())
        self.assertEqual(popup.view, "daily")
        self.assertTrue(popup.view_buttons["daily"].isChecked())
        self.assertTrue(all(not button.isChecked() for view, button in popup.view_buttons.items() if view != "daily"))
        window.set_view_options(chart_ma_periods=[5, 10, 20, 30])
        self.assertNotIn(60, popup.chart.periods)
        QTest.mouseClick(popup, Qt.RightButton, pos=QPoint(-12, -12))
        self.assertFalse(popup.isVisible())
        self.assertTrue(window.widget_visible)
        window.history.open_row("float", 1)
        self.wait_ready(window)
        self.assertIn("标的1", popup.title.text())
        QTest.keyClick(popup, Qt.Key_Escape)
        self.assertFalse(popup.isVisible())
        window.history.open_row("float", 1)
        QTest.mouseClick(popup.close_button, Qt.LeftButton)
        self.assertFalse(popup.isVisible())

    def test_popup_hide_discards_inflight_and_pending_results(self):
        _, window = self.make_window(chart_enabled=True, chart_display_mode="large")
        history = window.history
        with patch.object(history, "_start") as start:
            history.open_row("float", 0)
            request = start.call_args.args[0]
            history._busy = True
            history.open_row("float", 1)
            self.assertIsNotNone(history._pending)
            history.dialog.hide()
            self.assertIsNone(history._pending)
            history._accept((*request, BarResult(sample_bars(), "sina")))
            start.assert_called_once()
            self.assertFalse(history.dialog.isVisible())
            self.assertEqual(history.dialog.chart.bars, ())

    def test_floating_and_taskbar_anchors_keep_chart_on_target_screen(self):
        _, window = self.make_window(chart_enabled=True, chart_display_mode="large")
        bounds = self.app.primaryScreen().availableGeometry()
        window.move(bounds.topLeft() + QPoint(20, 20))
        self.app.processEvents()
        window.history.open_row("float", 0)
        self.wait_ready(window)
        popup = window.history.dialog
        self.assertTrue(bounds.contains(popup.geometry()))
        self.assertFalse(popup.geometry().intersects(window.geometry()))
        popup.close()
        position = QPoint(bounds.right() - 10, bounds.bottom() + 10)
        window.history.request_row("taskbar", 0, global_pos=position)
        with patch("stockwidget.ui.history.QCursor.pos", return_value=bounds.topLeft()):
            self.wait_until(popup.isVisible)
        self.wait_ready(window)
        self.assertTrue(bounds.contains(popup.geometry()))
        self.assertLess(popup.geometry().bottom(), position.y())
        self.assertGreater(popup.geometry().right(), bounds.center().x())

    def test_setting_persists_and_disable_closes_and_blocks_chart(self):
        dialog, window = self.make_window()
        self.assertFalse(dialog.ui.cb_chart_enabled.isChecked())
        window.history.open_row("float", 0)
        self.assertIsNone(window.history.dialog)
        window.history.cache.get.assert_not_called()
        dialog.ui.cb_chart_enabled.setChecked(True)
        self.assertTrue(window.current_config()["chart_enabled"])
        window.history.open_row("float", 0)
        QTest.qWait(30)
        self.assertTrue(window.history.dialog.isVisible())
        dialog.ui.cb_chart_enabled.setChecked(False)
        self.assertFalse(window.history.dialog.isVisible())
        calls = window.history.cache.get.call_count
        window.history.open_row("float", 1)
        self.assertEqual(window.history.cache.get.call_count, calls)

    def test_sorted_paged_split_rows_keep_instrument_identity(self):
        _, window = self.make_window(chart_enabled=True, float_paging_enabled=True, float_max_rows=2,
                                     float_split_enabled=True, taskbar_rows=1, taskbar_sync_split=False,
                                     taskbar_split_enabled=True, taskbar_sync_paging=False, taskbar_page_mode="manual")
        window.quotes.set_sort("现价", Qt.DescendingOrder)
        self.assertEqual(window.quotes.instrument_at("float", 0)["code"], "600006")
        self.assertEqual(window.quotes.instrument_at("float", 0, 1)["code"], "600004")
        window.quotes.change_page("float", 1)
        self.assertEqual(window.quotes.instrument_at("float", 0, 1)["code"], "600000")
        self.assertIsNone(window.quotes.instrument_at("float", 1, 1))
        window.quotes.change_page("taskbar", 1)
        self.assertEqual(window.quotes.instrument_at("taskbar", 0, 1)["code"], "600003")
        self.assertIsNone(window.quotes.instrument_at("float", -1))

    def test_click_opens_but_drag_and_double_click_do_not(self):
        for mode in ("window", "large", "medium", "small"):
            with self.subTest(mode=mode):
                _, window = self.make_window(chart_enabled=True, chart_display_mode=mode)
                target = window.table.viewport()
                pos = window.table.visualRect(window.model.index(0, 0)).center()
                with patch.object(window.history, "_open") as open_chart:
                    QTest.mouseClick(target, Qt.LeftButton, pos=pos)
                    self.assertTrue(window.history.click_timer.isActive())
                    self.wait_until(lambda: open_chart.call_count > 0)
                    open_chart.assert_called_once()
                    open_chart.reset_mock()
                    QTest.mousePress(target, Qt.LeftButton, pos=pos)
                    QTest.mouseMove(target, pos + QPoint(50, 20))
                    QTest.mouseRelease(target, Qt.LeftButton, pos=pos + QPoint(50, 20))
                    QTest.qWait(self.app.doubleClickInterval() + 30)
                    open_chart.assert_not_called()
                    QTest.mouseClick(target, Qt.LeftButton, pos=pos)
                    QTest.mouseDClick(target, Qt.LeftButton, pos=pos)
                    QTest.mouseRelease(target, Qt.LeftButton, pos=pos)
                    QTest.qWait(self.app.doubleClickInterval() + 30)
                    open_chart.assert_not_called()
                    self.assertFalse(window.widget_visible)

    def test_data_source_and_views_follow_settings_and_ma60_is_warmed(self):
        _, window = self.make_window(chart_enabled=True)
        window.history.open_row("float", 0)
        self.wait_ready(window)
        history = window.history.dialog
        self.assertEqual(window.history.cache.get.call_args.args[2], "sina")
        window.set_data_source("eastmoney")
        self.wait_ready(window)
        self.assertEqual(window.history.cache.get.call_args.args[2], "eastmoney")
        history.set_view("daily")
        self.wait_ready(window)
        self.assertEqual(len(history.chart.bars), 30)
        self.assertEqual(history.chart.averages[60][0], 41.5)
        window.set_view_options(chart_ma_periods=[5, 10, 20, 30])
        self.assertNotIn(60, history.chart.periods)
        for view in ("intraday", "five_day", "daily"):
            history.set_view(view)
            self.wait_ready(window)
            self.assertFalse(history.chart.grab().isNull())

    def test_late_result_cannot_replace_current_selection_or_reopen(self):
        _, window = self.make_window(chart_enabled=True)
        history = window.history
        with patch.object(history, "_start") as start:
            history.open_row("float", 0)
            request = start.call_args.args[0]
            history.open_row("float", 1)
            history._accept((*request, BarResult(sample_bars(), "sina")))
            self.assertEqual(history.dialog.chart.bars, ())
            self.assertIn("标的1", history.dialog.windowTitle())
            history.dialog.close()
            history._accept((history._generation, history.instrument, "daily", "sina", "2026-10-06", False, BarResult(sample_bars(), "sina")))
            self.assertFalse(history.dialog.isVisible())

    def test_taskbar_hit_regions_match_split_rows_and_empty_states(self):
        _, window = self.make_window(chart_enabled=True, taskbar_rows=4,
                                     taskbar_sync_split=False, taskbar_split_enabled=True)
        hits = []
        image = render_taskbar(window, 96, max_width=1500, hit_regions=hits)
        self.assertFalse(image.isNull())
        right = [(rect, row) for rect, row, block in hits if block == 1]
        self.assertTrue(right)
        self.assertEqual(window.quotes.instrument_at("taskbar", right[0][1], 1)["code"], "600004")
        controller = SimpleNamespace(source=window, _pager_rect=QRect(), _hit_regions=hits)
        point = right[0][0].center()
        with patch.object(window.history, "request_row") as request:
            TaskbarController._click(controller, point.x(), point.y())
            request.assert_called_once_with("taskbar", right[0][1], 1, global_pos=None)
            request.reset_mock()
            TaskbarController._click(controller, point.x(), point.y(), QPoint(800, 1000))
            request.assert_called_once_with("taskbar", right[0][1], 1, global_pos=QPoint(800, 1000))
        window.history.cache.get.return_value = BarResult(message="暂无可用历史数据")
        window.history.cache.get.side_effect = None
        window.history.open_row("float", 0)
        QTest.qWait(50)
        self.assertIn("暂无可用历史数据", window.history.dialog.status.text())
        self.assertFalse(window.history.dialog.chart.grab().isNull())

    def test_native_data_release_opens_chart_only_below_drag_threshold(self):
        from shiboken6 import delete

        _, window = self.make_window(chart_enabled=True, taskbar_enabled=True)
        controller = TaskbarController(window, Mock())
        try:
            render_taskbar(window, 96, max_width=1500, hit_regions=controller._hit_regions)
            rect, row, block = controller._hit_regions[0]
            local = rect.center()
            start = QPoint(500, 500)
            with patch.object(controller, "apply_mode"), patch.object(window.history, "request_row") as request:
                for movement, opens in ((QPoint(1, 1), True), (QPoint(50, 20), False)):
                    with self.subTest(movement=movement):
                        request.reset_mock()
                        controller._dispatch_pointer("press", local.x(), local.y(), start, False)
                        controller._dispatch_pointer("release", local.x(), local.y(), start + movement, False)
                        if opens:
                            request.assert_called_once_with("taskbar", row, block, global_pos=start + movement)
                        else:
                            request.assert_not_called()
        finally:
            controller.close()
            delete(controller)


@unittest.skipUnless(sys.platform == "win32" and os.environ.get("STOCKWIDGET_TEST_WINDOWS") == "1",
                     "Requires an opt-in Windows desktop session")
class WindowsHistoryPopupTests(HistoryTestCase):
    def test_native_outside_press_closes_popup_without_hiding_quotes(self):
        import ctypes
        from ctypes import wintypes as w

        from stockwidget.platform.window import get_user32

        class MouseInput(ctypes.Structure):
            _fields_ = [("dx", w.LONG), ("dy", w.LONG), ("data", w.DWORD),
                        ("flags", w.DWORD), ("time", w.DWORD), ("extra", ctypes.c_size_t)]

        class Input(ctypes.Structure):
            _fields_ = [("kind", w.DWORD), ("mouse", MouseInput)]

        _, window = self.make_window(chart_enabled=True, chart_display_mode="small")
        window.move(self.app.primaryScreen().availableGeometry().topLeft() + QPoint(20, 20))
        self.app.processEvents()
        window.history.open_row("float", 0)
        self.wait_ready(window)
        popup = window.history.dialog
        user = get_user32()
        user.SendInput.argtypes = [w.UINT, ctypes.POINTER(Input), ctypes.c_int]
        user.SendInput.restype = w.UINT
        # Actual OS input exercises popup capture outside its native client area.
        position = popup.mapFromGlobal(window.geometry().center())
        self.assertFalse(popup.rect().contains(position))
        QCursor.setPos(window.geometry().center())
        self.app.processEvents()
        inputs = (Input * 2)(Input(0, MouseInput(0, 0, 0, 0x0002, 0, 0)),
                             Input(0, MouseInput(0, 0, 0, 0x0004, 0, 0)))
        self.assertEqual(user.SendInput(2, inputs, ctypes.sizeof(Input)), 2)
        self.wait_until(lambda: not popup.isVisible())
        QTest.qWait(self.app.doubleClickInterval() + 30)
        self.assertFalse(popup.isVisible())
        self.assertTrue(window.widget_visible)
        self.assertTrue(window.isVisible())

    def test_native_taskbar_click_opens_popup_above_hidden_quote_window(self):
        import ctypes
        from ctypes import wintypes as w

        _, window = self.make_window(chart_enabled=True, chart_display_mode="small",
                                     taskbar_enabled=True, display_mode="taskbar")
        controller = TaskbarController(window, Mock())
        try:
            controller.apply_mode()
            self.app.processEvents()
            self.assertTrue(controller._active, window.taskbar_status)
            self.assertFalse(window.isVisible())
            rect, row, block = controller._hit_regions[0]
            point = rect.center()
            position = self.app.primaryScreen().availableGeometry().bottomRight() + QPoint(-50, 15)
            user, hwnd = controller.native.user, controller.native.hwnd
            user.SendMessageW.argtypes = [w.HWND, w.UINT, w.WPARAM, w.LPARAM]
            user.SendMessageW.restype = ctypes.c_ssize_t
            packed = (point.y() << 16) | point.x()
            with patch("stockwidget.ui.floating.taskbar.QCursor.pos", return_value=position):
                user.SendMessageW(hwnd, 0x201, 1, packed)
                user.SendMessageW(hwnd, 0x202, 0, packed)
                self.wait_until(lambda: window.history.dialog is not None and window.history.dialog.isVisible())
            self.wait_ready(window)
            popup = window.history.dialog
            self.assertIn(window.quotes.instrument_at("taskbar", row, block)["name"], popup.title.text())
            self.assertLess(popup.geometry().bottom(), position.y())
            self.assertTrue(self.app.primaryScreen().availableGeometry().contains(popup.geometry()))
            popup.close()
            self.assertTrue(window.widget_visible)
            self.assertFalse(window.isVisible())
            self.assertEqual(window.display_mode, "taskbar")
        finally:
            controller.close()
