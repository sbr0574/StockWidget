"""Settings, row identity, gesture safety, asynchronous selection and plotting."""

from unittest.mock import Mock, patch
from types import SimpleNamespace
import time

from PySide6.QtCore import QPoint, QRect, Qt
from PySide6.QtTest import QTest

from stockwidget.data.bars import BarResult
from stockwidget.data.quotes import _new_entry
from stockwidget.ui.floating.taskbar import TaskbarController, render_taskbar
from tests.data.test_bars import sample_bars
from tests.support import SettingsTestCase


class HistoryUITests(SettingsTestCase):
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
        window.history.cache.get = Mock(return_value=BarResult(sample_bars(), "sina"))
        window.show()
        self.app.processEvents()
        return dialog, window

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
        _, window = self.make_window(chart_enabled=True)
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
        history.view_combo.setCurrentIndex(2)
        self.wait_ready(window)
        self.assertEqual(len(history.chart.bars), 30)
        self.assertEqual(history.chart.averages[60][0], 41.5)
        history.ma_boxes[-1].setChecked(False)
        self.assertNotIn(60, history.chart.periods)
        for index in (0, 1, 2):
            history.view_combo.setCurrentIndex(index)
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
            history._accept((history._generation, history.instrument, "daily", "sina", BarResult(sample_bars(), "sina")))
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
            request.assert_called_once_with("taskbar", right[0][1], 1)
        window.history.cache.get.return_value = BarResult(message="暂无可用历史数据")
        window.history.open_row("float", 0)
        QTest.qWait(50)
        self.assertIn("暂无可用历史数据", window.history.dialog.status.text())
        self.assertFalse(window.history.dialog.chart.grab().isNull())
