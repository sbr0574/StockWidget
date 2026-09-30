"""Taskbar data projection, refresh lifecycle and opt-in native Windows checks."""

import os
import sys
import unittest
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication
from shiboken6 import delete

from stockwidget.platform.taskbar import TaskbarArea, NativeTaskbarWindow, find_taskbar
from stockwidget.ui.taskbar import TaskbarController, render_taskbar
from stockwidget.ui.widget import FloatLabel


class TaskbarTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.refresh = self.enterContext(patch.object(FloatLabel, "_refresh_from_function"))
        self.enterContext(patch("stockwidget.ui.widget.GlobalHotkeyManager"))
        self.enterContext(patch("stockwidget.ui.widget.apply_click_through"))
        self.window = FloatLabel({"taskbar_enabled": True, "taskbar_sync_metrics": False, "taskbar_sync_paging": False}, {})
        self.window.visible_metrics = ["name", "price"]
        self.window.set_view_options(taskbar_metrics=["name", "price"])
        self.controller = None
        self.window._clear_message()

    def tearDown(self):
        if self.controller:
            self.controller.close()
            delete(self.controller)
        delete(self.window)
        self.app.processEvents()

    def set_rows(self, rows):
        self.window._project_columns([dict(zip(("名称", "现价"), row)) for row in rows],
                                     [{"名称": "text", "现价": "up"} for _ in rows])

    def test_render_uses_only_first_two_rows(self):
        self.set_rows([["黄金", "123.45"], ["白银", "67.89"], ["不会出现", "100"]])
        before = render_taskbar(self.window, 44)
        self.set_rows([["黄金", "123.45"], ["白银", "67.89"], ["第三行非常长" * 20, "9999999"]])
        self.assertEqual(before, render_taskbar(self.window, 44))
        self.set_rows([["黄金", "1250.45"], ["白银", "67.89"]])
        self.assertNotEqual(before, render_taskbar(self.window, 44))

    def test_sorted_projection_changes_taskbar_and_preserves_main_rows(self):
        w = self.window
        w.visible_metrics = ["name", "price"]
        rows = [{"名称": name, "现价": price} for name, price in (("甲", "10"), ("乙乙", "30"), ("丙丙丙", "20"))]
        colors = [{"名称": "text", "现价": "up"} for _ in rows]
        values = [{"现价": value} for value in (10, 30, 20)]
        w._project_columns(rows, colors, values)
        original = render_taskbar(w, 44)
        w.sort_header, w.sort_order = "现价", Qt.DescendingOrder
        w._project_columns(rows, colors, values)
        self.assertEqual(w.model.rowCount(), 3)
        self.assertEqual(w.model.index(0, 0).data(), "乙乙")
        self.assertEqual(w.model.index(1, 0).data(), "丙丙丙")
        self.assertNotEqual(original, render_taskbar(w, 44))

    def test_render_fits_dpi_and_shows_errors_instead_of_stale_quotes(self):
        self.set_rows([["名称很长" * 30, "123.45"], ["白银", "67.89"]])
        before = render_taskbar(self.window, 66, 144, 220)
        self.assertLessEqual(before.width(), 220)
        self.assertEqual(before.height(), 66)
        self.window._show_message("网络请求失败", is_error=True)
        self.assertNotEqual(before, render_taskbar(self.window, 66, 144, 220))
        self.window._clear_message()
        self.assertEqual(before, render_taskbar(self.window, 66, 144, 220))

    def test_hiding_float_keeps_refresh_only_when_taskbar_enabled(self):
        self.window.show()
        self.window.hide()
        self.assertFalse(self.window.timer.isActive())
        with patch("stockwidget.ui.widget.sys.platform", "win32"):
            self.window.set_display_mode("taskbar")
        self.assertTrue(self.window.timer.isActive())
        self.window.show()
        self.window.hide()
        self.assertTrue(self.window.timer.isActive())
        self.window.set_display_mode("float")
        self.assertFalse(self.window.timer.isActive())

    def test_invalid_and_non_windows_modes_fall_back_to_float(self):
        with patch("stockwidget.ui.widget.sys.platform", "linux"):
            self.window.set_display_mode("taskbar")
        self.assertEqual(self.window.display_mode, "float")
        self.window.set_display_mode("invalid")
        self.assertEqual(self.window.display_mode, "float")

    def test_config_and_reset(self):
        with patch("stockwidget.ui.widget.sys.platform", "win32"):
            self.window.set_display_mode("both")
        self.window.set_taskbar_offset(150)
        self.assertEqual(self.window.current_config()["display_mode"], "both")
        self.assertEqual(self.window.current_config()["taskbar_offset"], 150)
        self.window.reset_settings()
        self.assertEqual(self.window.display_mode, "float")
        self.assertEqual(self.window.taskbar_offset, 0)

    def test_kline_and_empty_watchlist_render(self):
        self.window.visible_metrics = ["kline"]
        self.window.set_view_options(taskbar_metrics=["kline"])
        self.window._project_columns(
            [{"K线": {"k": (10, 12, 13, 9, 10)}}, {"K线": {"k": (12, 11, 14, 10, 12)}}],
            [{"K线": "up"}, {"K线": "down"}])
        image = render_taskbar(self.window, 44)
        self.assertFalse(image.isNull())
        self.window._process_data((True, {}, None))
        empty = render_taskbar(self.window, 44)
        self.assertNotEqual(image, empty)

    def test_settings_and_tray_mode_controls_stay_in_sync(self):
        from PySide6.QtGui import QIcon
        from stockwidget.ui.settings_dialog import SettingsDialog
        from stockwidget.ui.tray import TrayIcon

        self.window.set_view_options(taskbar_enabled=False)
        with patch.object(SettingsDialog, "_start_github_check"), patch("sys.platform", "win32"):
            dialog = SettingsDialog(self.window, self.window)
            tray = TrayIcon(QIcon(), "Test", on_toggle=Mock(), on_open_settings=Mock(),
                            on_quit=Mock(), on_click_through=Mock(), click_through_getter=lambda: False,
                            on_display_mode=self.window.set_display_mode,
                            display_mode_getter=lambda: self.window.display_mode,
                            taskbar_enabled_getter=lambda: self.window.view_options.taskbar_enabled)
            try:
                self.assertFalse(tray._mode_actions["taskbar"].isEnabled())
                dialog.taskbar_settings.setChecked(True)
                self.assertEqual(self.window.display_mode, "taskbar")
                tray.sync_click_through()
                self.assertTrue(tray._mode_actions["taskbar"].isChecked())
                tray._mode_actions["both"].trigger()
                self.assertTrue(dialog.taskbar_settings.dual.isChecked())
                self.assertTrue(dialog.taskbar_settings.isChecked())
                self.window.set_display_mode("float")
                self.assertTrue(dialog.taskbar_settings.isChecked())
                self.window.reset_settings()
                self.assertFalse(dialog.taskbar_settings.isChecked())
            finally:
                delete(tray)
                delete(dialog)

    def test_controller_recovers_from_missing_taskbar_and_stops_cleanly(self):
        self.set_rows([["黄金", "123.45"], ["白银", "67.89"]])
        self.controller = TaskbarController(self.window, Mock())
        area = TaskbarArea(123, 1280, 48, 900)
        native = Mock()
        with patch("stockwidget.ui.taskbar.QApplication.platformName", return_value="windows"), \
             patch("stockwidget.ui.taskbar.find_taskbar", return_value=None) as find, \
             patch("stockwidget.ui.taskbar.NativeTaskbarWindow", return_value=native), \
             patch("stockwidget.ui.widget.sys.platform", "win32"):
            self.window.set_display_mode("taskbar")
            self.assertTrue(self.window.isVisible())
            self.assertIn("自动重试", self.window.taskbar_status)
            find.return_value = area
            self.controller.refresh()
            self.assertFalse(self.window.isVisible())
            self.assertTrue(self.window.timer.isActive())
            self.assertTrue(self.controller._active)
            self.window.set_display_mode("float")
            self.assertTrue(self.window.isVisible())
            native.hide.assert_called()
            self.assertFalse(self.controller.timer.isActive())
            self.controller.close()
            native.close.assert_called_once()


@unittest.skipUnless(sys.platform == "win32" and os.environ.get("STOCKWIDGET_TEST_WINDOWS") == "1",
                     "Requires opt-in Windows desktop session")
class NativeTaskbarTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_native_drag_keeps_capture_while_hidden_then_releases_into_float(self):
        import ctypes
        from ctypes import wintypes as w
        from PySide6.QtCore import QPoint

        with patch.object(FloatLabel, "_refresh_from_function"), patch("stockwidget.ui.widget.GlobalHotkeyManager"):
            source = FloatLabel({"pos": {"x": 100, "y": 100}, "taskbar_enabled": True, "taskbar_sync_paging": False}, {})
        controller = TaskbarController(source, Mock())
        try:
            source.set_display_mode("taskbar")
            self.assertTrue(controller._active, source.taskbar_status)
            user, hwnd = controller.native.user, controller.native.hwnd
            # Synthetic messages do not hold the physical mouse button.
            self.enterContext(patch.object(controller.native, "poll_pointer", return_value=None))
            user.SendMessageW.argtypes = [w.HWND, w.UINT, w.WPARAM, w.LPARAM]
            user.SendMessageW.restype = ctypes.c_ssize_t
            press = QPoint(700, 770)
            with patch("stockwidget.ui.taskbar.QCursor.pos", return_value=press), \
                 patch("stockwidget.ui.taskbar.cursor_over_taskbar", return_value=True):
                user.SendMessageW(hwnd, 0x201, 1, (20 << 16) | 100)
                self.app.processEvents()
            self.assertEqual(user.GetCapture(), hwnd)
            outside = QPoint(710, 500)
            with patch("stockwidget.ui.taskbar.QCursor.pos", return_value=outside), \
                 patch("stockwidget.ui.taskbar.cursor_over_taskbar", return_value=False):
                xy = ((-250 & 0xffff) << 16) | 110
                user.SendMessageW(hwnd, 0x200, 1, xy)
                self.app.processEvents()
                self.assertTrue(source.isVisible())
                self.assertFalse(controller._active)
                self.assertTrue(user.IsWindow(hwnd))
                self.assertEqual(user.GetCapture(), hwnd)
                user.SendMessageW(hwnd, 0x202, 0, xy)
                self.app.processEvents()
            self.assertEqual(source.display_mode, "float")
            self.assertTrue(source.widget_visible)
            self.assertTrue(source.isVisible())
            self.assertIsNone(user.GetCapture())
            self.assertIsNone(controller.native.hwnd)
        finally:
            controller.close()
            delete(controller)
            delete(source)
            self.app.processEvents()

    def test_native_mouse_messages_page_then_double_click_hides_and_restores_taskbar(self):
        import ctypes
        from ctypes import wintypes as w

        with patch.object(FloatLabel, "_refresh_from_function"), patch("stockwidget.ui.widget.GlobalHotkeyManager"):
            source = FloatLabel({"pos": {"x": 100, "y": 100}, "taskbar_enabled": True, "taskbar_sync_paging": False}, {})
        source._clear_message()
        source._last_full_rows = [{"名称": f"测试{i}", "现价": str(i), "涨幅": "0%"} for i in range(5)]
        source._last_color_roles = [{h: "text" for h in row} for row in source._last_full_rows]
        source.set_view_options(taskbar_rows=3, taskbar_page_mode="manual")
        source._reproject_cached_data()
        controller = TaskbarController(source, Mock())
        try:
            source.set_display_mode("taskbar")
            self.app.processEvents()
            self.assertTrue(controller._active, source.taskbar_status)
            user, hwnd = controller.native.user, controller.native.hwnd
            user.SendMessageW.argtypes = [w.HWND, w.UINT, w.WPARAM, w.LPARAM]
            user.SendMessageW.restype = ctypes.c_ssize_t
            data_xy = (20 << 16) | (controller._pager_rect.right() + 20)
            user.SendMessageW(hwnd, 0x201, 1, data_xy)
            user.SendMessageW(hwnd, 0x202, 0, data_xy)
            self.app.processEvents()
            self.assertFalse(source.isVisible())
            rect = controller._pager_rect
            page_xy = ((rect.bottom() - 2) << 16) | rect.center().x()
            user.SendMessageW(hwnd, 0x201, 1, page_xy)
            user.SendMessageW(hwnd, 0x202, 0, page_xy)
            self.app.processEvents()
            self.assertEqual(source.taskbar_page, 1)
            self.assertEqual(source.taskbar_model.rowCount(), 2)
            user.SendMessageW(hwnd, 0x203, 1, data_xy)
            user.SendMessageW(hwnd, 0x202, 0, data_xy)
            self.app.processEvents()
            self.assertEqual(source.display_mode, "taskbar")
            self.assertFalse(source.widget_visible)
            self.assertFalse(source.isVisible())
            self.assertFalse(controller._active)
            self.assertIsNone(controller.native.hwnd)
            source.toggle_win()
            self.assertTrue(controller._active)
            self.assertFalse(source.isVisible())
        finally:
            controller.close()
            delete(controller)
            delete(source)
            self.app.processEvents()

    def test_embed_recreate_and_cleanup_without_modifying_taskbar(self):
        import ctypes
        from ctypes import wintypes as w
        from PySide6.QtGui import QColor, QImage

        area = find_taskbar()
        self.assertIsNotNone(area)
        native = NativeTaskbarWindow(Mock(), Mock())
        user = native.user
        user.GetForegroundWindow.argtypes = []
        user.GetForegroundWindow.restype = w.HWND
        foreground = user.GetForegroundWindow()
        before = w.RECT()
        user.GetWindowRect(area.hwnd, ctypes.byref(before))
        image = QImage(160, 40, QImage.Format_ARGB32_Premultiplied)
        image.fill(QColor(0, 0, 0, 1))
        try:
            for _ in range(2):
                native.present(area, 160, 40, bytes(image.constBits()))
                self.app.processEvents()
                self.assertEqual(user.GetParent(native.hwnd), area.hwnd)
                self.assertTrue(user.GetWindowLongW(native.hwnd, -16) & 0x40000000)
                self.assertEqual(user.GetForegroundWindow(), foreground)
                # Simulate loss of our child window without restarting the user's Explorer.
                user.DestroyWindow(native.hwnd)
                self.app.processEvents()
                self.assertIsNone(native.hwnd)
            after = w.RECT()
            user.GetWindowRect(area.hwnd, ctypes.byref(after))
            self.assertEqual(bytes(before), bytes(after))
        finally:
            native.close()
