"""Settings, row identity, gesture safety, asynchronous selection and plotting."""

from unittest.mock import Mock, patch
from dataclasses import replace
from datetime import datetime, timedelta
from types import SimpleNamespace
import os
import sys
import time
import unittest

from PySide6.QtCore import QPoint, QRect, Qt
from PySide6.QtGui import QCursor
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication

from stockwidget.data.bars import BarResult
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
        window.history.cache.get = Mock(return_value=BarResult(sample_bars(), "sina"))
        window.show()
        self.app.processEvents()
        return dialog, window


class HistoryUITests(HistoryTestCase):
    def test_compact_chart_time_labels_keep_endpoints_without_overlap(self):
        from shiboken6 import delete

        chart = HistoryChart()
        try:
            daily = sample_bars()
            start = datetime(2026, 9, 30, 9, 30)
            minutes = tuple(replace(bar, time=(start + timedelta(minutes=i)).strftime("%Y-%m-%d %H:%M:%S"))
                            for i, bar in enumerate(daily))
            for view in ("intraday", "five_day", "daily"):
                for width in (360, 435, 560):
                    with self.subTest(view=view, width=width):
                        chart.resize(width, 220)
                        chart.set_data(daily if view == "daily" else minutes, view, {"market": "sh", "type": "沪"})
                        self.assertFalse(chart.grab().isNull())
                        labels = chart._time_labels()
                        stop = 10 if view == "daily" else 16
                        self.assertEqual(labels[0][1], chart.bars[0].time[5:stop])
                        self.assertEqual(labels[-1][1], chart.bars[-1].time[5:stop])
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
        combo.setCurrentIndex(combo.findData("floating"))
        self.assertEqual(window.current_config()["chart_display_mode"], "floating")
        window.history.open_row("float", 0)
        self.wait_ready(window)
        popup = window.history.dialog
        popup.view_combo.setCurrentIndex(2)
        self.wait_ready(window)
        popup.ma_boxes[-1].setChecked(False)
        combo.setCurrentIndex(combo.findData("window"))
        self.wait_ready(window)
        self.assertTrue(popup.isVisible())
        self.assertEqual(popup.view, "daily")
        self.assertNotIn(60, popup.chart.periods)
        self.assertEqual(len(popup.chart.bars), 30)
        combo.setCurrentIndex(combo.findData("floating"))
        self.wait_ready(window)
        self.assertIs(QApplication.activePopupWidget(), popup)
        dialog.ui.cb_chart_enabled.setChecked(False)
        self.assertFalse(popup.isVisible())
        self.assertFalse(combo.isEnabled())
        self.assertEqual(window.current_config()["chart_display_mode"], "floating")
        dialog.ui.cb_chart_enabled.setChecked(True)
        self.assertEqual(combo.currentData(), "floating")
        window.reset_settings()
        self.assertFalse(dialog.ui.cb_chart_enabled.isChecked())
        self.assertEqual(combo.currentData(), "window")

    def test_floating_controls_dropdown_outside_click_and_escape(self):
        _, window = self.make_window(chart_enabled=True, chart_display_mode="floating")
        window.history.open_row("float", 0)
        self.wait_ready(window)
        popup = window.history.dialog
        QTest.mouseClick(popup.chart, Qt.LeftButton, pos=popup.chart.rect().center())
        self.assertTrue(popup.isVisible())
        popup.view_combo.showPopup()
        QTest.qWait(200)  # Native combo popups briefly ignore the opening release.
        dropdown = popup.view_combo.view()
        QTest.mouseClick(dropdown.viewport(), Qt.LeftButton,
                         pos=dropdown.visualRect(dropdown.model().index(2, 0)).center())
        self.wait_ready(window)
        self.assertTrue(popup.isVisible())
        self.assertEqual(popup.view, "daily")
        QTest.mouseClick(popup.ma_boxes[-1], Qt.LeftButton)
        self.assertNotIn(60, popup.chart.periods)
        QTest.mouseClick(popup, Qt.RightButton, pos=QPoint(-12, -12))
        self.assertFalse(popup.isVisible())
        self.assertFalse(window.history.timer.isActive())
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
        _, window = self.make_window(chart_enabled=True, chart_display_mode="floating")
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
            self.assertFalse(history.timer.isActive())

    def test_floating_and_taskbar_anchors_keep_chart_on_target_screen(self):
        _, window = self.make_window(chart_enabled=True, chart_display_mode="floating")
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
        for mode in ("window", "floating"):
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
            request.assert_called_once_with("taskbar", right[0][1], 1, global_pos=None)
            request.reset_mock()
            TaskbarController._click(controller, point.x(), point.y(), QPoint(800, 1000))
            request.assert_called_once_with("taskbar", right[0][1], 1, global_pos=QPoint(800, 1000))
        window.history.cache.get.return_value = BarResult(message="暂无可用历史数据")
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

        _, window = self.make_window(chart_enabled=True, chart_display_mode="floating")
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
        self.assertFalse(window.history.timer.isActive())

    def test_native_taskbar_click_opens_popup_above_hidden_quote_window(self):
        import ctypes
        from ctypes import wintypes as w

        _, window = self.make_window(chart_enabled=True, chart_display_mode="floating",
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
