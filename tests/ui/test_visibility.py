"""隐藏 / 恢复、快捷键、置顶和窗口销毁的生命周期。"""

from datetime import datetime, timedelta
from unittest.mock import Mock, patch
import os
import sys
import unittest

from PySide6.QtCore import QPoint, QTime, Qt
from PySide6.QtGui import QColor, QPalette
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QScrollArea, QWidget, QPushButton
from shiboken6 import delete

from stockwidget.core.window_rules import QUOTE_TIMEZONE
from stockwidget.platform.hotkeys import HotkeyResult
from stockwidget.platform.taskbar import TaskbarArea
from stockwidget.platform.window import apply_click_through, ensure_topmost, get_user32
from stockwidget.ui.floating.presenter import QuotePresenter
from stockwidget.ui.floating.taskbar import taskbar_message, render_taskbar, TaskbarController
from stockwidget.ui.floating.widget import FloatLabel
from stockwidget.ui.settings.dialog import SettingsDialog
from stockwidget.ui.watchlist.add_panel import AddCodePanel

from tests.support import QtTestCase


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


class HidingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        cls.app.setQuitOnLastWindowClosed(False)

    def setUp(self):
        self.enterContext(patch.object(QuotePresenter, "refresh"))
        self.enterContext(patch("stockwidget.ui.floating.widget.GlobalHotkeyManager"))
        self.enterContext(patch("stockwidget.ui.floating.widget.apply_click_through"))
        self.enterContext(patch("stockwidget.ui.floating.interaction.time.time", return_value=NOW.timestamp()))
        self.monotonic = self.enterContext(patch("stockwidget.ui.floating.interaction.time.monotonic", return_value=100))
        self.win = FloatLabel({"watchlist": CODES, "taskbar_enabled": True,
                               "scheduled_hide_times": ["15:00", "16:00"]}, CODES)
        self.win.show()
        self.app.processEvents()

    def tearDown(self):
        delete(self.win)
        self.app.processEvents()

    def refresh(self, data=None):
        self.win.quotes.accept_result((True, data if data is not None else {code: entry() for code in CODES}, None))

    def restore_and_countdown(self):
        self.win.hide_widget()
        self.win.toggle_win()
        self.win.set_hide_options(auto_hide_enabled=True)
        self.refresh()

    def test_initial_automatic_hide_all_surfaces_restore_mode_and_position(self):
        for mode in ("float", "taskbar", "both"):
            for source in ("sina", "eastmoney"):
                with self.subTest(mode=mode, source=source):
                    self.win.hide_controller._manual_show = False
                    self.win.widget_visible = True
                    self.win.display_mode = mode
                    self.win.set_data_source(source)
                    self.win.move(120, 170)
                    self.win.set_hide_options(auto_hide_enabled=True)
                    notices = []
                    handler = lambda: notices.append(self.win.widget_visible)
                    self.win.widget_visibility_changed.connect(handler)
                    self.refresh({code: entry(timestamp=source == "eastmoney") for code in CODES})
                    self.assertFalse(self.win.widget_visible)
                    self.assertIn(False, notices)
                    self.assertTrue(self.win.isHidden())
                    self.assertEqual((self.win.pos(), self.win.display_mode), (QPoint(120, 170), mode))
                    self.win.toggle_win()
                    self.assertTrue(self.win.widget_visible)
                    self.assertEqual((self.win.pos(), self.win.display_mode), (QPoint(120, 170), mode))
                    self.win.widget_visibility_changed.disconnect(handler)

    def test_manual_restore_countdown_five_seconds_not_restarted_by_refresh(self):
        self.restore_and_countdown()
        self.assertTrue(self.win.widget_visible)
        self.assertIn("5秒", self.win.hide_notice.text())
        self.assertTrue(self.win.hide_controller.countdown_timer.isActive())
        self.monotonic.return_value = 102
        self.refresh()
        self.win.hide_controller.check_countdown()
        self.assertIn("3秒", self.win.hide_notice.text())
        self.assertEqual(self.win.model.rowCount(), 2)
        self.assertEqual(taskbar_message(self.win), self.win.hide_notice.text())
        for dpi in (96, 144, 192):
            self.assertFalse(render_taskbar(self.win, round(48 * dpi / 96), dpi=dpi).isNull())
        self.monotonic.return_value = 104.999
        self.win.hide_controller.check_countdown()
        self.assertTrue(self.win.widget_visible)
        self.monotonic.return_value = 105
        self.win.hide_controller.check_countdown()
        self.assertFalse(self.win.widget_visible)
        self.assertEqual(self.win.hide_notice.text(), "")
        self.assertFalse(self.win.hide_controller.countdown_timer.isActive())
        self.win.toggle_win()
        self.refresh()
        self.assertIn("5秒", self.win.hide_notice.text())

    def test_refresh_fresh_incomplete_missing_time_or_failure_cancels_countdown(self):
        missing_time = {key: value for key, value in entry().items() if key not in ("date", "time")}
        for payload in ((True, {"sh600000": entry(), "sz000001": entry(30)}, None),
                        (True, {"sh600000": entry()}, None),
                        (True, {"sh600000": entry(), "sz000001": missing_time}, None),
                        (True, {}, None), (False, None, "网络请求失败")):
            with self.subTest(payload=payload):
                self.restore_and_countdown()
                self.win.quotes.accept_result(payload)
                self.assertTrue(self.win.widget_visible)
                self.assertEqual(self.win.hide_notice.text(), "")
                self.assertFalse(self.win.hide_controller.countdown_timer.isActive())

    def test_disabling_auto_switching_source_and_watchlist_cancel_pending_hide(self):
        self.restore_and_countdown()
        self.win.set_hide_options(auto_hide_enabled=False)
        self.assertEqual(self.win.hide_notice.text(), "")
        self.assertEqual(self.win.scheduled_hide_times, ["15:00", "16:00"])
        self.restore_and_countdown()
        generation = self.win.quotes._quote_generation
        self.win.set_data_source("eastmoney")
        self.assertEqual(self.win.hide_notice.text(), "")
        self.win.quotes.accept_result((True, {code: entry() for code in CODES}, None, generation))
        self.assertIsNone(self.win.hide_controller._deadline)
        self.refresh()
        self.assertIsNotNone(self.win.hide_controller._deadline)
        self.win.set_watchlist({})
        self.assertIsNone(self.win.hide_controller._deadline)
        self.win.reset_settings()
        self.assertFalse(self.win.auto_hide_enabled)
        self.assertFalse(self.win.scheduled_hide_enabled)
        self.assertEqual(self.win.scheduled_hide_times, [])

    def test_scheduled_hide_overrides_countdown_once_a_day_even_when_hidden(self):
        self.restore_and_countdown()
        self.win.set_hide_options(scheduled_hide_enabled=True)
        controller = self.win.hide_controller
        controller._last_check = NOW - timedelta(seconds=1)
        controller.check_schedule(NOW)
        self.assertFalse(self.win.widget_visible)
        self.assertEqual(self.win.hide_notice.text(), "")
        self.win.toggle_win()
        controller.check_schedule(NOW + timedelta(seconds=20))
        self.assertTrue(self.win.widget_visible)
        controller.check_schedule(NOW + timedelta(days=1))
        self.assertFalse(self.win.widget_visible)
        controller.check_schedule(NOW + timedelta(days=1, hours=1))
        self.win.toggle_win()
        controller.check_schedule(NOW + timedelta(days=1, hours=1, seconds=10))
        self.assertTrue(self.win.widget_visible)

    def test_schedule_timer_independent_of_refresh_and_preserves_config(self):
        self.win.set_hide_options(scheduled_hide_enabled=True, auto_hide_enabled=True)
        self.win.hide_widget()
        self.assertFalse(self.win.timer.isActive())
        self.assertTrue(self.win.hide_controller.schedule_timer.isActive())
        saved = self.win.current_config()
        other = FloatLabel(saved, CODES)
        try:
            self.assertTrue(other.scheduled_hide_enabled)
            self.assertTrue(other.auto_hide_enabled)
            self.assertEqual(other.scheduled_hide_times, ["15:00", "16:00"])
        finally:
            delete(other)
        self.win.set_hide_options(scheduled_hide_enabled=False)
        self.assertFalse(self.win.hide_controller.schedule_timer.isActive())
        self.assertTrue(self.win.auto_hide_enabled)
        self.assertEqual(self.win.scheduled_hide_times, ["15:00", "16:00"])

    def test_master_hide_switch_cancels_both_rules_and_preserves_options_on_reload(self):
        self.win.set_hide_options(scheduled_hide_enabled=True, auto_hide_enabled=True)
        self.restore_and_countdown()
        self.assertIsNotNone(self.win.hide_controller._deadline)
        self.win.set_hide_options(hide_enabled=False)
        self.assertEqual(self.win.hide_notice.text(), "")
        self.assertFalse(self.win.hide_controller.schedule_timer.isActive())
        self.assertFalse(self.win.hide_controller.countdown_timer.isActive())
        self.refresh()
        self.win.hide_controller._last_check = NOW - timedelta(seconds=1)
        self.win.hide_controller.check_schedule(NOW)
        self.assertTrue(self.win.widget_visible)
        saved = self.win.current_config()
        other = FloatLabel(saved, CODES)
        try:
            self.assertFalse(other.hide_enabled)
            self.assertTrue(other.auto_hide_enabled)
            self.assertTrue(other.scheduled_hide_enabled)
            self.assertEqual(other.scheduled_hide_times, ["15:00", "16:00"])
            self.assertFalse(other.hide_controller.schedule_timer.isActive())
            other.set_hide_options(hide_enabled=True)
            self.assertTrue(other.hide_controller.schedule_timer.isActive())
        finally:
            delete(other)

    def test_old_hide_switches_enable_the_new_master_when_loading_legacy_config(self):
        for scheduled, automatic in ((False, False), (True, False), (False, True), (True, True)):
            with self.subTest(scheduled=scheduled, automatic=automatic):
                other = FloatLabel({"scheduled_hide_enabled": scheduled,
                                    "auto_hide_enabled": automatic}, CODES)
                try:
                    self.assertEqual(other.hide_enabled, scheduled or automatic)
                    self.assertEqual(other.scheduled_hide_enabled, scheduled)
                    self.assertEqual(other.auto_hide_enabled, automatic)
                finally:
                    delete(other)

    def test_display_mode_menu_restore_also_gets_countdown(self):
        self.win.set_hide_options(auto_hide_enabled=True)
        self.win.hide_widget()
        self.win.set_display_mode("both")
        self.refresh()
        self.assertTrue(self.win.widget_visible)
        self.assertIn("5秒", self.win.hide_notice.text())

    def test_native_taskbar_hides_and_restores_with_auto_and_scheduled_rules(self):
        self.enterContext(patch("stockwidget.ui.floating.taskbar.QApplication.platformName", return_value="windows"))
        self.enterContext(patch("stockwidget.ui.floating.taskbar.find_taskbar", return_value=TaskbarArea(1, 1280, 48, 900)))
        self.enterContext(patch("stockwidget.ui.floating.widget.sys.platform", "win32"))
        native = Mock()
        self.enterContext(patch("stockwidget.ui.floating.taskbar.NativeTaskbarWindow", return_value=native))
        controller = TaskbarController(self.win, Mock())
        try:
            self.win.set_hide_options(auto_hide_enabled=True, scheduled_hide_enabled=True)
            for mode in ("taskbar", "both"):
                self.win.set_display_mode(mode)
                self.win.hide_controller._manual_show = False
                controller.apply_mode()
                self.assertTrue(controller._active)
                self.refresh()
                self.assertFalse(controller._active)
                native.hide.assert_called()
                self.win.toggle_win()
                self.refresh()
                self.assertTrue(controller._active)
                self.assertIn("5秒", self.win.hide_notice.text())
                self.win.hide_controller._last_check = NOW - timedelta(seconds=1)
                self.win.hide_controller._fired.clear()
                self.win.hide_controller.check_schedule(NOW)
                self.assertFalse(controller._active)
                self.assertFalse(self.win.widget_visible)
                self.assertEqual(self.win.display_mode, mode)
                self.win.toggle_win()
                self.assertTrue(controller._active)
        finally:
            controller.close()
            delete(controller)

    def test_scheduled_hide_cancels_drag_and_restores_original_position(self):
        origin = self.win.pos()
        start = origin + QPoint(5, 5)
        self.win.begin_drag(start)
        self.win.move_drag(start + QPoint(40, 25))
        self.assertNotEqual(self.win.pos(), origin)
        self.win.set_hide_options(scheduled_hide_enabled=True)
        self.win.hide_controller._last_check = NOW - timedelta(seconds=1)
        self.win.hide_controller.check_schedule(NOW)
        self.assertFalse(self.win.widget_visible)
        self.assertEqual(self.win.pos(), origin)
        self.assertIsNone(self.win._drag_pos)
        self.win.toggle_win()
        self.assertEqual(self.win.pos(), origin)

    def test_settings_add_delete_limit_duplicates_toggle_theme_and_layout(self):
        with patch.object(SettingsDialog, "_start_github_check"):
            dialog = SettingsDialog(self.win, self.win)
        original = self.app.palette()
        try:
            dialog.show()
            dialog.ui.settings_pages.setCurrentWidget(dialog.ui.general)
            dialog.ui.gb_scheduled_hide.setChecked(True)
            for hour in (17, 18):
                dialog.ui.hide_time_edit.setTime(QTime(hour, 0))
                dialog.ui.btn_add_hide_time.click()
            self.assertEqual(self.win.scheduled_hide_times, ["15:00", "16:00", "17:00"])
            self.assertFalse(dialog.ui.btn_add_hide_time.isEnabled())
            self.app.processEvents()
            times = dialog.ui.list_hide_times
            self.assertTrue(all(times.viewport().rect().contains(times.visualItemRect(times.item(i)))
                                for i in range(3)))
            dialog.ui.list_hide_times.setCurrentRow(1)
            item = dialog.ui.list_hide_times.currentItem()
            rect = dialog.ui.list_hide_times.visualItemRect(item)
            point = dialog.ui.list_hide_times.itemDelegate().delete_rect(rect).center()
            QTest.mouseClick(dialog.ui.list_hide_times.viewport(), Qt.LeftButton, pos=point)
            self.assertEqual(self.win.scheduled_hide_times, ["15:00", "17:00"])
            self.assertTrue(dialog.ui.btn_add_hide_time.isEnabled())
            dialog.ui.hide_time_edit.setTime(QTime(15, 0))
            self.assertFalse(dialog.ui.btn_add_hide_time.isEnabled())
            dialog.ui.cb_auto_hide.setChecked(True)
            self.assertTrue(self.win.auto_hide_enabled)
            self.win.set_data_source("eastmoney")
            self.assertTrue(dialog.ui.cb_auto_hide.isEnabled())
            for enabled in (False, True):
                dialog.ui.gb_scheduled_hide.setChecked(enabled)
                for foreground, background in (("#eeeeee", "#222222"), ("#222222", "#eeeeee")):
                    palette = QPalette(original)
                    palette.setColor(QPalette.WindowText, QColor(foreground))
                    palette.setColor(QPalette.Window, QColor(background))
                    self.app.setPalette(palette)
                    self.app.processEvents()
                    self.assertEqual(dialog.ui.hide_time_edit.isEnabled(), enabled)
                    self.assertEqual(dialog.ui.list_hide_times.isEnabled(), enabled)
                    self.assertEqual(dialog.ui.cb_auto_hide.isEnabled(), enabled)
                    self.assertEqual(self.win.hide_enabled, enabled)
                    self.assertTrue(self.win.auto_hide_enabled)
                    self.assertIn("程序将在" if enabled else "程序未启用自动隐藏", dialog.ui.label_hide_status.text())
            page = dialog.ui.general
            self.assertTrue(page.findChildren(QScrollArea))
            content = dialog.ui.general_content
            for control in (dialog.ui.gb_icon, dialog.ui.cb_auto_start_row, dialog.ui.label_hide_status,
                            dialog.ui.gb_scheduled_hide, dialog.ui.list_hide_times,
                            dialog.ui.btn_add_hide_time, dialog.ui.hide_time_edit):
                bounds = control.rect().translated(control.mapTo(content, QPoint()))
                self.assertTrue(content.rect().contains(bounds), (control.objectName(), bounds, content.rect()))
            self.assertLess(dialog.ui.cb_auto_start_row.geometry().bottom(), dialog.ui.cmb_color_mode_row.geometry().top())
        finally:
            self.app.setPalette(original)
            delete(dialog)


class WidgetTopmostTests(QtTestCase):

    def setUp(self):
        with patch.object(QuotePresenter, "refresh"), patch(
            "stockwidget.ui.floating.widget.GlobalHotkeyManager"
        ):
            self.window = FloatLabel({}, {})
        for name, kwargs in (
            ("ensure_topmost", {}),
            ("apply_click_through", {}),
            ("QApplication.activeWindow", {"return_value": None}),
            ("QApplication.activePopupWidget", {"return_value": None}),
        ):
            mock = self.enterContext(patch("stockwidget.ui.floating.widget." + name, **kwargs))
            if name == "ensure_topmost":
                self.native_top = mock
        self.window.force_top = True
        self.window.show()
        self.native_top.reset_mock()

    def tearDown(self):
        delete(self.window)

    def test_default_and_saved_topmost_configuration(self):
        self.assertTrue(self.window.float_on_top)
        self.assertTrue(self.window.windowFlags() & Qt.WindowStaysOnTopHint)
        for cfg in ({"force_top": True}, {"float_on_top": False, "force_top": True}):
            with self.subTest(cfg=cfg), patch.object(QuotePresenter, "refresh"), patch(
                "stockwidget.ui.floating.widget.GlobalHotkeyManager"
            ), patch("stockwidget.ui.floating.widget.force_top_supported", return_value=True):
                restored = FloatLabel(cfg, {})
                try:
                    expected = cfg.get("float_on_top", True)
                    self.assertEqual(restored.float_on_top, expected)
                    self.assertEqual(bool(restored.windowFlags() & Qt.WindowStaysOnTopHint), expected)
                    self.assertEqual(restored.force_top, expected)
                    self.assertEqual(restored._keep_top_timer.isActive(), expected)
                finally:
                    delete(restored)

    def test_disabling_topmost_stops_force_top_and_preserves_visible_state(self):
        window = self.window
        self.app.processEvents()
        window.move(120, 140)
        geometry = window.geometry()
        changed = Mock()
        visibility = Mock()
        window.set_on_change(changed)
        window.widget_visibility_changed.connect(visibility)
        window.hide_controller._deadline = 12345
        window._keep_top_timer.start()
        for enabled in (False, True, False):
            with self.subTest(enabled=enabled):
                window.set_float_on_top(enabled)
                self.assertEqual(bool(window.windowFlags() & Qt.WindowStaysOnTopHint), enabled)
                self.assertTrue(window.isVisible())
                self.assertTrue(window.widget_visible)
                self.assertEqual(window.geometry(), geometry)
                self.assertEqual(window.hide_controller._deadline, 12345)
                self.assertFalse(window.force_top)
                self.assertFalse(window._keep_top_timer.isActive())
                self.assertTrue(window.timer.isActive())
                self.assertFalse(window.testAttribute(Qt.WA_ShowWithoutActivating))
                window._keep_top_timer.timeout.emit()
        visibility.assert_not_called()
        self.native_top.assert_not_called()
        self.assertEqual(changed.call_count, 3)
        with patch("stockwidget.ui.floating.widget.force_top_supported", return_value=True):
            window.set_force_top(True)
        self.assertFalse(window.force_top)

    def test_hidden_topmost_toggle_and_restore_keep_mode_and_position(self):
        window = self.window
        for mode in ("float", "taskbar", "both"):
            with self.subTest(mode=mode):
                window.display_mode = mode
                window.show()
                window.hide_widget()
                position = window.pos()
                for enabled in (False, True, False):
                    window.set_float_on_top(enabled)
                    self.assertFalse(window.isVisible())
                    self.assertFalse(window.widget_visible)
                    self.assertEqual(window.pos(), position)
                    self.assertEqual(window.display_mode, mode)
                window.toggle_win()
                self.assertTrue(window.widget_visible)
                self.assertEqual(window.pos(), position)
                self.assertEqual(window.display_mode, mode)
                self.assertFalse(window.windowFlags() & Qt.WindowStaysOnTopHint)
                self.assertEqual(window.isVisible(), mode != "taskbar")

    def test_topmost_setting_roundtrips_and_resets_to_enabled(self):
        self.window.set_float_on_top(False)
        cfg = self.window.current_config()
        self.assertFalse(cfg["float_on_top"])
        self.assertFalse(cfg["force_top"])
        with patch.object(QuotePresenter, "refresh"), patch(
            "stockwidget.ui.floating.widget.GlobalHotkeyManager"
        ):
            restored = FloatLabel(cfg, {})
            try:
                self.assertFalse(restored.float_on_top)
                self.assertFalse(restored.windowFlags() & Qt.WindowStaysOnTopHint)
                restored.reset_settings()
                self.assertTrue(restored.float_on_top)
                self.assertTrue(restored.windowFlags() & Qt.WindowStaysOnTopHint)
                self.assertFalse(restored.force_top)
            finally:
                delete(restored)

    def test_timer_keeps_raising_with_click_through_enabled(self):
        self.window.click_through = True
        self.window._keep_top_timer.timeout.emit()
        self.native_top.assert_called_once_with(self.window)

    def test_click_through_toggle_immediately_restores_topmost(self):
        for enabled in (True, False):
            with self.subTest(enabled=enabled):
                self.native_top.reset_mock()
                self.window.set_click_through(enabled)
                self.native_top.assert_called_once_with(self.window)

    def test_hidden_or_disabled_window_is_not_raised(self):
        self.window.hide()
        self.window._ensure_on_top()
        self.native_top.assert_not_called()
        self.window.force_top = False
        self.window.show()
        self.window.set_click_through(True)
        self.window._ensure_on_top()
        self.native_top.assert_not_called()

    def test_show_restores_topmost_and_timer_with_click_through(self):
        self.window.click_through = True
        self.window.hide()
        self.window.show()
        self.native_top.assert_called_once_with(self.window)
        self.assertTrue(self.window._keep_top_timer.isActive())

    def test_settings_and_popups_are_not_covered(self):
        other = QWidget()
        try:
            self.window.click_through = True
            for name in ("activeWindow", "activePopupWidget"):
                with self.subTest(name=name), patch(
                    "stockwidget.ui.floating.widget.QApplication." + name, return_value=other
                ):
                    self.window._ensure_on_top()
                    self.native_top.assert_not_called()
        finally:
            delete(other)


@unittest.skipUnless(
    sys.platform == "win32" and os.environ.get("STOCKWIDGET_TEST_WINDOWS") == "1",
    "Requires an opt-in Windows desktop session",
)
class WindowsTopmostIntegrationTests(QtTestCase):

    def test_float_topmost_toggle_changes_native_style_without_hiding_settings(self):
        from ctypes import wintypes

        from stockwidget.ui.settings.dialog import SettingsDialog

        self.assertEqual(self.app.platformName(), "windows")
        api = get_user32()
        api.GetForegroundWindow.argtypes = []
        api.GetForegroundWindow.restype = wintypes.HWND
        with patch.object(QuotePresenter, "refresh"), patch(
            "stockwidget.ui.floating.widget.GlobalHotkeyManager"
        ):
            window = FloatLabel({}, {})
        with patch.object(SettingsDialog, "_start_github_check"):
            dialog = SettingsDialog(window, window)
        try:
            window.show()
            self.app.processEvents()
            window.move(30, 30)
            window.set_force_top(True)
            window.set_click_through(True)
            dialog.show()
            dialog.activateWindow()
            self.app.processEvents()
            geometry = window.geometry()
            foreground = api.GetForegroundWindow()
            for enabled in (False, True, False):
                with self.subTest(enabled=enabled):
                    dialog.ui.cb_float_on_top.setChecked(enabled)
                    self.app.processEvents()
                    style = api.GetWindowLongW(int(window.winId()), -20)
                    self.assertEqual(bool(style & 0x8), enabled)  # WS_EX_TOPMOST
                    self.assertTrue(style & 0x20)  # WS_EX_TRANSPARENT
                    self.assertTrue(style & 0x80000)  # WS_EX_LAYERED
                    self.assertEqual(window.geometry(), geometry)
                    self.assertTrue(window.isVisible())
                    self.assertTrue(dialog.isVisible())
                    self.assertEqual(api.GetForegroundWindow(), foreground)
        finally:
            delete(dialog)
            delete(window)

    def test_native_topmost_preserves_click_through_geometry_and_foreground(self):
        import ctypes
        from ctypes import wintypes

        self.assertEqual(self.app.platformName(), "windows")
        api = get_user32()
        api.GetForegroundWindow.argtypes = []
        api.GetForegroundWindow.restype = wintypes.HWND
        api.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
        api.GetWindowRect.restype = wintypes.BOOL
        window = QWidget()
        window.setWindowFlags(Qt.Tool | Qt.FramelessWindowHint)
        window.setAttribute(Qt.WA_TranslucentBackground)
        window.setAttribute(Qt.WA_ShowWithoutActivating)
        window.setGeometry(30, 30, 100, 40)
        try:
            window.show()
            self.app.processEvents()
            hwnd = int(window.winId())
            foreground = api.GetForegroundWindow()
            before = wintypes.RECT()
            self.assertTrue(api.GetWindowRect(hwnd, ctypes.byref(before)))
            for enabled in (True, False, True):
                with self.subTest(click_through=enabled):
                    # 先降为普通窗口，验证原生调用确实恢复 TOPMOST。
                    self.assertTrue(api.SetWindowPos(hwnd, -2, 0, 0, 0, 0, 0x13))
                    apply_click_through(window, enabled)
                    self.assertEqual(api.GetForegroundWindow(), foreground)
                    self.assertTrue(ensure_topmost(window))
                    self.app.processEvents()
                    style = api.GetWindowLongW(hwnd, -20)
                    self.assertTrue(style & 0x8)  # WS_EX_TOPMOST
                    self.assertEqual(bool(style & 0x20), enabled)  # WS_EX_TRANSPARENT
                    self.assertTrue(style & 0x80000)  # WS_EX_LAYERED
                    self.assertEqual(api.GetForegroundWindow(), foreground)
                    after = wintypes.RECT()
                    self.assertTrue(api.GetWindowRect(hwnd, ctypes.byref(after)))
                    self.assertEqual(bytes(after), bytes(before))
        finally:
            delete(window)


class WidgetHotkeyTests(QtTestCase):

    def setUp(self):
        with patch.object(QuotePresenter, "refresh"), patch(
            "stockwidget.ui.floating.widget.GlobalHotkeyManager"
        ):
            self.window = FloatLabel({}, {})
        self.manager = self.window._hotkeys
        self.manager.reset_mock()
        self.saved = Mock()
        self.window.set_on_change(self.saved)

    def tearDown(self):
        self.window.close()
        self.window.deleteLater()
        self.app.processEvents()

    def test_disabled_hotkeys_save_without_registration(self):
        for update in (self.window.update_hotkey, self.window.update_click_through_hotkey):
            self.assertTrue(update(" Ctrl+Alt+X "))
        self.assertEqual(self.window.hotkey, "Ctrl+Alt+X")
        self.assertEqual(self.window.hotkey_click_through, "Ctrl+Alt+X")
        self.manager.register.assert_not_called()
        self.assertEqual(self.saved.call_count, 2)

    def test_failed_update_keeps_requested_key_and_registers_other_hotkey(self):
        self.window.hotkey_enabled = True
        self.window.hotkey_click_through_enabled = True
        self.manager.register.side_effect = [
            HotkeyResult(False, "conflict"), HotkeyResult(True),
        ]

        result = self.window.update_hotkey("Ctrl+Alt+X")

        self.assertEqual(result.reason, "conflict")
        self.assertEqual(self.window.hotkey, "Ctrl+Alt+X")
        self.assertEqual(
            [c.args[0] for c in self.manager.register.call_args_list],
            ["Ctrl+Alt+X", "Ctrl+Alt+C"],
        )
        self.assertFalse(self.window.hotkey_results["hotkey"])
        self.assertTrue(self.window.hotkey_results["hotkey_click_through"])
        self.saved.assert_called_once_with()

    def test_failed_enable_preserves_other_hotkey(self):
        self.window.hotkey_enabled = True
        self.manager.register.side_effect = [
            HotkeyResult(True), HotkeyResult(False, "conflict"),
        ]

        self.assertFalse(self.window.set_click_through_hotkey_enabled(True))

        self.assertTrue(self.window.hotkey_click_through_enabled)
        self.assertTrue(self.window.hotkey_enabled)
        self.assertEqual(
            [c.args[0] for c in self.manager.register.call_args_list],
            ["Ctrl+Alt+F", "Ctrl+Alt+C"],
        )
        self.saved.assert_called_once_with()

    def test_unchanged_failed_key_can_retry_and_disabling_clears_failure(self):
        self.manager.register.side_effect = [HotkeyResult(False, "conflict"), HotkeyResult(True)]
        self.assertFalse(self.window.set_hotkey_enabled(True))
        self.assertTrue(self.window.update_hotkey(self.window.hotkey))
        self.assertEqual(self.manager.register.call_count, 2)
        self.saved.assert_called_once_with()
        self.assertTrue(self.window.set_hotkey_enabled(False))
        self.assertNotIn("hotkey", self.window.hotkey_results)

    def test_other_failure_does_not_change_successful_update_result(self):
        self.window.hotkey_enabled = self.window.hotkey_click_through_enabled = True
        self.manager.register.side_effect = [HotkeyResult(False, "conflict"), HotkeyResult(True)]
        self.assertTrue(self.window.update_click_through_hotkey("Ctrl+Alt+X"))
        self.assertFalse(self.window.hotkey_results["hotkey"])

    def test_successful_enable_saves_once_and_repeated_value_is_noop(self):
        self.manager.register.return_value = HotkeyResult(True)
        self.assertTrue(self.window.set_hotkey_enabled(True))
        self.assertTrue(self.window.set_hotkey_enabled(True))
        self.manager.register.assert_called_once()
        self.saved.assert_called_once_with()


class WidgetLifecycleTests(QtTestCase):

    def test_deleting_float_window_cancels_pending_resize(self):
        with patch.object(QuotePresenter, "refresh"), patch("sys.excepthook") as errors:
            window = FloatLabel({}, {})
            window._defer_fit()
            delete(window)
            self.app.processEvents()
        errors.assert_not_called()

    def test_clearing_watchlist_discards_pending_quotes_and_errors(self):
        with patch.object(QuotePresenter, "refresh"):
            window = FloatLabel({"watchlist": {"au0": {"checked": True}}}, {})
        try:
            window.model.set_rows_headers([["800"]], ["现价"], [["text"]])
            window.quotes._refresh_thread = Mock()
            window.quotes._refresh_thread.is_alive.return_value = True
            window.set_watchlist({})
            self.assertEqual(window.model.rowCount(), 0)
            window.quotes.accept_result((True, {"au0": {"current_price": 800}}, None))
            self.assertEqual(window.model.rowCount(), 0)
            window.quotes.accept_result((False, None, "网络请求失败"))
            self.assertIn("添加自选股", window.message_label.text())
        finally:
            delete(window)
            self.app.processEvents()

    def test_deleting_add_panel_cancels_pending_focus(self):
        parent = QWidget()
        anchor = QPushButton(parent)
        panel = AddCodePanel("搜索", parent)
        with patch("sys.excepthook") as errors:
            panel.show_for(anchor)
            delete(parent)
            self.app.processEvents()
        errors.assert_not_called()
