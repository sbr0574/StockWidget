"""任务栏绘制、输入、停靠 / 拖出和可选 Windows 原生集成。"""

from unittest.mock import Mock, patch
import os
import sys
import unittest

from PySide6.QtCore import Qt, QEvent, QPoint, QPointF
from PySide6.QtGui import QColor, QKeyEvent, QMouseEvent, QPalette
from PySide6.QtWidgets import QApplication
from shiboken6 import delete

from stockwidget.platform.taskbar import TaskbarArea, NativeTaskbarWindow, find_taskbar
from stockwidget.ui.floating.presenter import QuotePresenter
from stockwidget.ui.floating.taskbar import TaskbarController, render_taskbar, taskbar_message
from stockwidget.ui.floating.widget import FloatLabel
from stockwidget.ui.menus import build_quote_menu

from tests.support import QtTestCase, PagingTestCase


class TaskbarTests(QtTestCase):

    def setUp(self):
        self.refresh = self.enterContext(patch.object(QuotePresenter, "refresh"))
        self.enterContext(patch("stockwidget.ui.floating.widget.GlobalHotkeyManager"))
        self.enterContext(patch("stockwidget.ui.floating.widget.apply_click_through"))
        self.enterContext(patch("stockwidget.ui.settings.groups.find_taskbar", return_value=None))
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
        self.window.quotes.project_rows([dict(zip(("名称", "现价"), row)) for row in rows],
                                     [{"名称": "text", "现价": "up"} for _ in rows])

    def test_startup_shows_loading_for_configured_taskbar_metrics(self):
        for synced, floating, independent, expected in (
            (True, ["price"], [], "加载中…"),
            (False, [], ["price"], "加载中…"),
            (True, [], ["price"], "请选择任务栏显示指标"),
            (False, ["price"], [], "请选择任务栏显示指标"),
        ):
            with self.subTest(synced=synced, metrics=(floating, independent)):
                window = FloatLabel({"taskbar_enabled": True, "display_mode": "taskbar",
                                     "visible_metrics": floating, "name_visible": False,
                                     "taskbar_metrics": independent,
                                     "taskbar_sync_metrics": synced}, {})
                try:
                    self.assertEqual(taskbar_message(window), expected)
                    self.assertFalse(render_taskbar(window, 44).isNull())
                finally:
                    delete(window)

    def test_quote_menu_exit_uses_the_application_quit_callback(self):
        quit_app = Mock()
        self.window.set_quit_callback(lambda: quit_app())
        menu = build_quote_menu(self.window)
        try:
            next(action for action in menu.actions() if action.text() == "退出").trigger()
            quit_app.assert_called_once_with()
        finally:
            delete(menu)

    def test_automatic_taskbar_color_tracks_theme_without_overwriting_manual_color(self):
        original = self.app.palette()
        self.window.set_view_options(taskbar_sync_appearance=False, taskbar_auto_color=True,
                                     taskbar_color="#123456", taskbar_unicolor=False)
        self.set_rows([["黄金", "123.45"]])
        with patch("stockwidget.ui.controls.style.QGuiApplication.styleHints") as hints:
            hints.return_value.colorScheme.return_value = Qt.ColorScheme.Unknown
            try:
                images = []
                for background, expected in (("#202020", "#ffffff"), ("#eeeeee", "#000000")):
                    palette = QPalette(original)
                    palette.setColor(QPalette.Window, QColor(background))
                    self.app.setPalette(palette)
                    self.app.processEvents()
                    self.assertEqual(self.window.get_taskbar_appearance()[1].name(), expected)
                    for column in range(self.window.taskbar_model.columnCount()):
                        color = self.window.taskbar_model.index(0, column).data(Qt.ForegroundRole)
                        self.assertEqual(color.name(), expected)
                    images.append(render_taskbar(self.window, 44))
                self.assertNotEqual(*images)
                self.assertEqual(self.window.current_config()["taskbar_color"], "#123456")
                self.assertFalse(self.window.current_config()["taskbar_unicolor"])
                self.window.set_view_options(taskbar_sync_appearance=True)
                self.assertEqual(self.window.get_taskbar_appearance()[1], self.window.fg)
                self.window.set_view_options(taskbar_sync_appearance=False, taskbar_auto_color=False)
                self.assertEqual(self.window.get_taskbar_appearance()[1].name(), "#123456")
                self.assertFalse(self.window.get_taskbar_appearance()[3])
            finally:
                self.app.setPalette(original)

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
        w.quotes.project_rows(rows, colors, values)
        original = render_taskbar(w, 44)
        w.quotes.sort_header, w.quotes.sort_order = "现价", Qt.DescendingOrder
        w.quotes.project_rows(rows, colors, values)
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
        with patch("stockwidget.ui.floating.widget.sys.platform", "win32"):
            self.window.set_display_mode("taskbar")
        self.assertTrue(self.window.timer.isActive())
        self.window.show()
        self.window.hide()
        self.assertTrue(self.window.timer.isActive())
        self.window.set_display_mode("float")
        self.assertFalse(self.window.timer.isActive())

    def test_invalid_and_non_windows_modes_fall_back_to_float(self):
        for platform in ("linux", "darwin"):
            with self.subTest(platform=platform), patch("sys.platform", platform):
                self.window.set_display_mode("taskbar")
                menu = build_quote_menu(self.window)
                try:
                    labels = [action.text() for action in menu.actions()]
                    self.assertNotIn("任务栏行情", labels)
                    self.assertNotIn("显示位置", labels)
                    self.assertIn("分栏", labels)
                    self.assertEqual(self.window.display_mode, "float")
                finally:
                    delete(menu)
        self.assertEqual(self.window.display_mode, "float")
        self.window.set_display_mode("invalid")
        self.assertEqual(self.window.display_mode, "float")

    def test_config_and_reset(self):
        with patch("stockwidget.ui.floating.widget.sys.platform", "win32"):
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
        self.window.quotes.project_rows(
            [{"K线": {"k": (10, 12, 13, 9, 10)}}, {"K线": {"k": (12, 11, 14, 10, 12)}}],
            [{"K线": "up"}, {"K线": "down"}])
        image = render_taskbar(self.window, 44)
        self.assertFalse(image.isNull())
        self.window.quotes.accept_result((True, {}, None))
        empty = render_taskbar(self.window, 44)
        self.assertNotEqual(image, empty)

    def test_settings_and_tray_mode_controls_stay_in_sync(self):
        from PySide6.QtGui import QIcon
        from stockwidget.ui.settings.dialog import SettingsDialog
        from stockwidget.ui.menus import TrayIcon

        self.window.set_view_options(taskbar_enabled=False)
        with patch.object(SettingsDialog, "_start_github_check"), patch("sys.platform", "win32"):
            dialog = SettingsDialog(self.window, self.window)
            tray = TrayIcon(QIcon(), "Test", on_toggle=Mock(), on_open_settings=Mock(),
                            on_quit=Mock(), source=self.window)
            try:
                actions = {action.text(): action for action in tray.contextMenu().actions()}
                mode = actions["任务栏行情"]
                self.assertFalse(mode.isChecked())
                dialog.taskbar_settings.enabled.setChecked(True)
                self.assertEqual(self.window.display_mode, "taskbar")
                tray.sync_settings()
                self.assertTrue(mode.isChecked())
                mode.trigger()
                self.assertFalse(dialog.taskbar_settings.enabled.isChecked())
                mode.trigger()
                dialog.taskbar_settings.dual.setChecked(True)
                self.window.set_display_mode("both")
                self.assertTrue(dialog.taskbar_settings.dual.isChecked())
                self.assertTrue(dialog.taskbar_settings.enabled.isChecked())
                actions["分栏"].trigger()
                self.assertTrue(dialog.float_split_settings.isChecked())
                self.window.set_display_mode("float")
                self.assertTrue(dialog.taskbar_settings.enabled.isChecked())
                self.window.set_view_options(taskbar_dual_open=False, taskbar_sync_split=False,
                                             float_split_enabled=False, taskbar_split_enabled=False)
                self.window.set_display_mode("taskbar")
                tray.sync_settings()
                self.assertFalse(actions["分栏"].isChecked())
                actions["分栏"].trigger()
                self.assertTrue(self.window.view_options.taskbar_split_enabled)
                self.assertFalse(self.window.view_options.float_split_enabled)
                menu = build_quote_menu(self.window, "taskbar")
                try:
                    split = next(action for action in menu.actions() if action.text() == "分栏")
                    self.assertTrue(split.isChecked())
                    split.trigger()
                    self.assertFalse(self.window.view_options.taskbar_split_enabled)
                    self.assertFalse(self.window.view_options.float_split_enabled)
                finally:
                    delete(menu)
                self.window.reset_settings()
                self.assertFalse(dialog.taskbar_settings.enabled.isChecked())
                tray.sync_settings()
                self.assertFalse(mode.isChecked())
                with patch("stockwidget.ui.menus.QDesktopServices.openUrl") as open_url:
                    actions["使用帮助"].trigger()
                    actions["问题反馈"].trigger()
                self.assertEqual([call.args[0].toString() for call in open_url.call_args_list],
                                 ["https://github.com/sbr0574/StockWidget#readme", "https://github.com/sbr0574/StockWidget/issues"])
            finally:
                delete(tray)
                delete(dialog)

    def test_controller_recovers_from_missing_taskbar_and_stops_cleanly(self):
        self.set_rows([["黄金", "123.45"], ["白银", "67.89"]])
        self.controller = TaskbarController(self.window, Mock())
        area = TaskbarArea(123, 1280, 48, 900)
        native = Mock()
        with patch("stockwidget.ui.floating.taskbar.QApplication.platformName", return_value="windows"), \
             patch("stockwidget.ui.floating.taskbar.find_taskbar", return_value=None) as find, \
             patch("stockwidget.ui.floating.taskbar.NativeTaskbarWindow", return_value=native), \
             patch("stockwidget.ui.floating.widget.sys.platform", "win32"):
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
class NativeTaskbarTests(QtTestCase):
    def test_float_preview_keeps_native_capture_and_restores_on_leave_and_escape(self):
        import ctypes
        from ctypes import wintypes as w

        with patch.object(QuotePresenter, "refresh"), patch("stockwidget.ui.floating.widget.GlobalHotkeyManager"):
            source = FloatLabel({"pos": {"x": 100, "y": 100}, "taskbar_enabled": True}, {})
        controller = TaskbarController(source, Mock())
        try:
            source.show()
            self.app.processEvents()
            self.enterContext(patch.object(NativeTaskbarWindow, "poll_pointer", return_value=None))
            start = source.pos() + QPoint(5, 5)
            source.begin_drag(start)
            with patch("stockwidget.ui.floating.taskbar.cursor_over_taskbar", return_value=True):
                source.move_drag(start + QPoint(80, 80))
                self.app.processEvents()
            self.assertTrue(controller._active, source.taskbar_status)
            self.assertFalse(source.isVisible())
            self.assertTrue(source.widget_visible)
            user, hwnd = controller.native.user, controller.native.hwnd
            self.assertEqual(user.GetCapture(), hwnd)
            user.SendMessageW.argtypes = [w.HWND, w.UINT, w.WPARAM, w.LPARAM]
            user.SendMessageW.restype = ctypes.c_ssize_t
            xy = ((-250 & 0xffff) << 16) | 110
            with patch("stockwidget.ui.floating.taskbar.QCursor.pos", return_value=QPoint(300, 300)), \
                 patch("stockwidget.ui.floating.taskbar.cursor_over_taskbar", return_value=False):
                user.SendMessageW(hwnd, 0x200, 1, xy)
                self.app.processEvents()
                self.assertTrue(source.isVisible())
                self.assertTrue(source.widget_visible)
                self.assertFalse(controller._active)
                self.assertEqual(user.GetCapture(), hwnd)
                user.SendMessageW(hwnd, 0x202, 0, xy)
                self.app.processEvents()
            self.assertEqual(source.display_mode, "float")
            self.assertIsNone(user.GetCapture())
            self.assertIsNone(controller.native.hwnd)
            origin = source.pos()
            source.begin_drag(origin + QPoint(5, 5))
            with patch("stockwidget.ui.floating.taskbar.cursor_over_taskbar", return_value=True):
                source.move_drag(origin + QPoint(90, 90))
                self.app.processEvents()
                self.assertFalse(source.isVisible())
                user.SendMessageW(controller.native.hwnd, 0x100, 0x1B, 0)
                self.app.processEvents()
            self.assertEqual(source.pos(), origin)
            self.assertTrue(source.isVisible())
            self.assertFalse(source.taskbar_preview_active)
            self.assertIsNone(user.GetCapture())
            self.assertIsNone(controller.native.hwnd)
        finally:
            controller.close()
            delete(controller)
            delete(source)
            self.app.processEvents()


    def test_native_drag_keeps_capture_while_hidden_then_releases_into_float(self):
        import ctypes
        from ctypes import wintypes as w
        from PySide6.QtCore import QPoint

        with patch.object(QuotePresenter, "refresh"), patch("stockwidget.ui.floating.widget.GlobalHotkeyManager"):
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
            with patch("stockwidget.ui.floating.taskbar.QCursor.pos", return_value=press), \
                 patch("stockwidget.ui.floating.taskbar.cursor_over_taskbar", return_value=True):
                user.SendMessageW(hwnd, 0x201, 1, (20 << 16) | 100)
                self.app.processEvents()
            self.assertEqual(user.GetCapture(), hwnd)
            outside = QPoint(710, 500)
            with patch("stockwidget.ui.floating.taskbar.QCursor.pos", return_value=outside), \
                 patch("stockwidget.ui.floating.taskbar.cursor_over_taskbar", return_value=False):
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

        with patch.object(QuotePresenter, "refresh"), patch("stockwidget.ui.floating.widget.GlobalHotkeyManager"):
            source = FloatLabel({"pos": {"x": 100, "y": 100}, "taskbar_enabled": True, "taskbar_sync_paging": False}, {})
        source._clear_message()
        source.quotes._last_full_rows = [{"名称": f"测试{i}", "现价": str(i), "涨幅": "0%"} for i in range(5)]
        source.quotes._last_color_roles = [{h: "text" for h in row} for row in source.quotes._last_full_rows]
        source.set_view_options(taskbar_rows=3, taskbar_page_mode="manual")
        source.quotes.reproject()
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
            self.assertEqual(source.quotes.taskbar_page, 1)
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


class NativeInputTests(unittest.TestCase):
    def test_polling_finishes_background_drag_once_and_supports_escape(self):
        native = NativeTaskbarWindow.__new__(NativeTaskbarWindow)
        native.hwnd, native._pointer_down, native._ignore_release = 1, True, False
        native._last_pointer_position = (100, 20)
        native.user = Mock()
        native.user.GetCapture.return_value = 1
        native.user.GetCursorPos.side_effect = lambda pointer: (
            setattr(pointer._obj, "x", 110), setattr(pointer._obj, "y", -150), True)[-1]
        native.user.ScreenToClient.return_value = True
        native.user.GetAsyncKeyState.side_effect = lambda key: 0x8000 if key == 1 else 0
        self.assertEqual(native.poll_pointer(), ("move", 110, -150))
        self.assertIsNone(native.poll_pointer())
        native.user.GetAsyncKeyState.return_value = 0
        native.user.GetAsyncKeyState.side_effect = None
        self.assertEqual(native.poll_pointer(), ("release", 110, -150))
        self.assertIsNone(native.poll_pointer())
        self.assertTrue(native._ignore_release)
        native._pointer_down = True
        native.user.GetAsyncKeyState.side_effect = lambda key: 0x8000 if key == 0x1B else 0
        self.assertEqual(native.poll_pointer(), ("cancel", 110, -150))

    def test_double_click_sequence_does_not_dispatch_two_single_clicks(self):
        native = NativeTaskbarWindow.__new__(NativeTaskbarWindow)
        native._ignore_release = False
        native._pointer_down = False
        native.hwnd = 1
        native._on_pointer = Mock()
        native.user = Mock()
        xy = (20 << 16) | 100
        for message in (0x201, 0x202, 0x203, 0x202):
            native._wndproc(1, message, 0, xy)
        self.assertEqual([call.args for call in native._on_pointer.call_args_list],
                         [("press", 100, 20), ("release", 100, 20), ("double_click", 100, 20)])

    def test_native_capture_and_cancel_emit_pointer_events_only(self):
        native = NativeTaskbarWindow.__new__(NativeTaskbarWindow)
        native.hwnd, native._pointer_down, native._ignore_release = 1, False, False
        native._last_frame = None
        native._on_pointer, native.user = Mock(), Mock()
        native.user.GetCapture.return_value = 1
        native._wndproc(1, 0x201, 1, (20 << 16) | 100)
        native.user.SetCapture.assert_called_once_with(1)
        native._wndproc(1, 0x200, 1, ((-100 & 0xffff) << 16) | 120)
        native.hide()
        native.user.DestroyWindow.assert_not_called()
        native._wndproc(1, 0x215, 0, 0)
        self.assertFalse(native._pointer_down)
        self.assertEqual([call.args[0] for call in native._on_pointer.call_args_list], ["press", "move", "cancel"])


class DockingInteractionTests(PagingTestCase):
    def test_native_taskbar_drag_uses_full_size_when_float_is_a_strip(self):
        w = self.win
        controller = self.make_controller()
        w.move(80, 80)
        w.set_position_options(boundary_check_enabled=True, edge_hide_enabled=True)
        w.set_display_mode("both")
        self.app.processEvents()
        screen = self.app.primaryScreen().geometry()
        with patch("stockwidget.ui.floating.interaction.QCursor.pos", return_value=QPoint(99999, 99999)):
            w.move(screen.right() - w.width() + 1, 100)
            self.assertEqual(w.width(), 5)
            full = w.position_controller.full_geometry()
            controller._dpi = 192
            start = QPoint(400, screen.bottom())
            controller._dispatch_pointer("press", 80, 10, start, True)
            self.assertEqual(w._drag_pos.x(), 40)
            self.assertEqual(w.geometry(), full)
            controller._dispatch_pointer("move", 80, 10, start - QPoint(50, 50), False)
            controller._dispatch_pointer("cancel", 80, 10, start, False)
            self.assertEqual(w.position_controller.full_geometry(), full)
            self.assertEqual(w.width(), 5)
            self.assertEqual(w.display_mode, "both")
            self.assertTrue(controller._active)

    def test_taskbar_double_click_hides_and_toggle_restores_the_same_location(self):
        controller = self.make_controller()
        self.win.set_display_mode("taskbar")
        controller._click(100, 20)
        self.app.processEvents()
        self.assertEqual(self.win.display_mode, "taskbar")
        self.assertFalse(self.win.isVisible())

        controller._double_click(100, 20)
        self.app.processEvents()
        self.assertEqual(self.win.display_mode, "taskbar")
        self.assertFalse(self.win.widget_visible)
        self.assertFalse(self.win.isVisible())
        self.assertFalse(controller._active)
        self.populate()
        self.win.set_view_options(taskbar_rows=3)
        controller.refresh()
        self.assertFalse(controller._active)
        self.assertFalse(self.win.widget_visible)
        self.assertFalse(self.win.timer.isActive())
        self.win.toggle_win()
        self.assertEqual(self.win.display_mode, "taskbar")
        self.assertTrue(controller._active)
        self.assertFalse(self.win.isVisible())

    def test_taskbar_menu_switch_enables_docking_and_can_restore_the_float(self):
        controller = self.make_controller()
        self.win.set_view_options(taskbar_enabled=False)
        self.win.move(100, 100)
        target, origin = self.win.table.viewport(), self.win.pos()
        start = target.mapToGlobal(QPoint(5, 5))
        self.over.reset_mock()
        self.over.return_value = True
        self.mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
        self.mouse(target, QEvent.MouseMove, start + QPoint(0, 600), Qt.NoButton, Qt.LeftButton)
        self.mouse(target, QEvent.MouseButtonRelease, start + QPoint(0, 600), Qt.LeftButton, Qt.NoButton)
        self.assertEqual(self.win.pos(), origin + QPoint(0, 600))
        self.assertFalse(controller._active)
        self.assertEqual(self.win.display_mode, "float")
        self.assertFalse(self.win.taskbar_preview_active)
        self.over.assert_not_called()
        self.win.set_display_mode("taskbar")
        self.assertEqual(self.win.display_mode, "float")
        menu = self.win.build_context_menu()
        try:
            actions = {action.text(): action for action in menu.actions()}
            self.assertNotIn("显示位置", actions)
            self.assertFalse(actions["任务栏行情"].isChecked())
            actions["任务栏行情"].trigger()
            self.assertTrue(self.win.view_options.taskbar_enabled)
            self.assertEqual(self.win.display_mode, "taskbar")
            self.assertFalse(self.win.isVisible())
            self.assertTrue(controller._active)
            actions["分栏"].trigger()
            self.assertTrue(self.win.view_options.float_split_enabled)
            with patch("stockwidget.ui.floating.widget.apply_click_through"):
                actions["鼠标穿透"].trigger()
                self.assertTrue(self.win.click_through)
        finally:
            delete(menu)
        self.win.set_view_options(taskbar_enabled=True)
        self.win.set_display_mode("taskbar")
        self.assertTrue(controller._active)
        self.win.set_view_options(taskbar_enabled=False)
        self.assertEqual(self.win.display_mode, "float")
        self.assertTrue(self.win.isVisible())
        self.assertFalse(controller._active)

    def test_disabling_taskbar_during_docking_cancels_preview_and_keeps_float(self):
        controller = self.make_controller()
        origin = self.win.pos()
        start = origin + QPoint(5, 5)
        self.over.return_value = True
        self.win.begin_drag(start)
        self.win.move_drag(start + QPoint(0, 200))
        self.assertTrue(controller._active)
        self.win.set_view_options(taskbar_enabled=False)
        self.win.finish_drag()
        self.assertFalse(controller._active)
        self.assertFalse(controller._dragging)
        self.assertFalse(self.win.taskbar_preview_active)
        self.assertEqual(self.win.display_mode, "float")
        self.assertEqual(self.win.pos(), origin)
        self.assertFalse(controller.timer.isActive())

    def test_taskbar_page_click_only_changes_taskbar_page(self):
        controller = self.make_controller()
        self.win.set_view_options(taskbar_rows=3, taskbar_page_mode="manual")
        self.win.set_display_mode("taskbar")
        rect = controller._pager_rect
        controller._click(rect.center().x(), rect.bottom() - 2)
        self.app.processEvents()
        self.assertEqual(self.win.quotes.taskbar_page, 1)
        controller._double_click(rect.center().x(), rect.center().y())
        self.app.processEvents()
        self.assertEqual(self.win.display_mode, "taskbar")
        self.assertFalse(self.win.widget_visible)
        self.assertFalse(self.win.isVisible())

    def test_dual_mode_drag_never_checks_docking_or_hides_taskbar(self):
        controller = self.make_controller()
        self.win.set_view_options(float_split_enabled=True)
        self.win.set_display_mode("both")
        self.win.move(100, 100)
        target = self.win.right_table.viewport()
        origin = self.win.pos()
        start = target.mapToGlobal(QPoint(5, 5))
        self.over.reset_mock()
        self.native.hide.reset_mock()
        self.mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
        for delta, over in ((QPoint(90, 30), False), (QPoint(90, 650), True), (QPoint(200, 90), False)):
            self.over.return_value = over
            self.mouse(target, QEvent.MouseMove, start + delta, Qt.NoButton, Qt.LeftButton)
            self.assertEqual(self.win.pos(), origin + delta)
            self.assertEqual(self.win.display_mode, "both")
            self.assertTrue(self.win.isVisible())
            self.assertTrue(controller._active)
            self.assertFalse(self.win.taskbar_preview_active)
        self.mouse(target, QEvent.MouseButtonRelease, start + delta, Qt.LeftButton, Qt.NoButton)
        self.assertEqual(self.win.display_mode, "both")
        self.over.assert_not_called()
        self.native.hide.assert_not_called()
        # Esc cancellation also uses the same free-drag path.
        origin = self.win.pos()
        start = target.mapToGlobal(QPoint(5, 5))
        self.mouse(target, QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
        self.mouse(target, QEvent.MouseMove, start + QPoint(50, 30), Qt.NoButton, Qt.LeftButton)
        QApplication.sendEvent(self.win, QKeyEvent(QEvent.KeyPress, Qt.Key_Escape, Qt.NoModifier))
        self.assertEqual(self.win.pos(), origin)
        self.assertTrue(controller._active)
        self.native.hide.assert_not_called()

    def test_taskbar_pager_drag_uses_same_transition_without_flipping_page(self):
        controller = self.make_controller()
        self.win.set_view_options(taskbar_page_mode="manual")
        self.win.set_display_mode("taskbar")
        rect = controller._pager_rect
        press = QPoint(700, 760)
        outside = press - QPoint(0, 150)
        controller._dispatch_pointer("press", rect.center().x(), rect.bottom() - 2, press, True)
        controller._dispatch_pointer("move", rect.center().x(), -130, outside, False)
        controller._dispatch_pointer("release", rect.center().x(), -130, outside, False)
        self.assertEqual(self.win.quotes.taskbar_page, 0)
        self.assertEqual(self.win.display_mode, "float")
        self.assertTrue(self.win.isVisible())

    def test_native_dual_drag_preserves_taskbar_and_cancel_cleans_capture(self):
        controller = self.make_controller()
        self.win.set_display_mode("both")
        origin = self.win.pos()
        press = QPoint(700, 760)
        self.native.hide.reset_mock()
        controller._dispatch_pointer("press", 100, 20, press, True)
        controller._dispatch_pointer("move", 110, -130, press - QPoint(0, 150), False)
        self.assertTrue(controller._active)
        self.assertEqual(self.win.display_mode, "both")
        QApplication.sendEvent(self.win, QKeyEvent(QEvent.KeyPress, Qt.Key_Escape, Qt.NoModifier))
        self.assertEqual(self.win.pos(), origin)
        self.assertFalse(controller.input_timer.isActive())
        self.assertIsNone(controller._native_press)
        self.native.release_pointer.assert_called()
        self.native.hide.assert_not_called()

    def test_dual_mode_double_click_hides_both_and_toggle_restores_both(self):
        controller = self.make_controller()
        self.win.set_display_mode("both")
        controller._double_click(100, 20)
        self.app.processEvents()
        self.assertFalse(self.win.isVisible())
        self.assertFalse(controller._active)
        self.assertFalse(self.win.widget_visible)
        self.win.toggle_win()
        self.assertTrue(self.win.isVisible())
        self.assertTrue(controller._active)
        self.win.mouseDoubleClickEvent(QMouseEvent(QEvent.MouseButtonDblClick, QPointF(5, 5), QPointF(5, 5),
                                                   Qt.LeftButton, Qt.LeftButton, Qt.NoModifier))
        self.assertFalse(self.win.isVisible())
        self.assertFalse(controller._active)
        self.win.toggle_win()
        self.assertEqual(self.win.display_mode, "both")

    def test_native_drag_out_uses_shared_drag_and_can_redock_or_cancel(self):
        controller = self.make_controller()
        self.win.set_display_mode("taskbar")
        original = self.win.pos()
        press, outside = QPoint(700, 760), QPoint(710, 500)
        controller._dispatch_pointer("press", 100, 15, press, True)
        controller._dispatch_pointer("move", 110, -245, outside, False)
        self.assertTrue(self.win.isVisible())
        self.assertFalse(controller._active)
        self.assertEqual(self.win.pos(), outside - self.win._drag_pos)
        controller._dispatch_pointer("cancel", 0, 0, outside, False)
        self.assertEqual(self.win.pos(), original)
        self.assertEqual(self.win.display_mode, "taskbar")
        self.assertTrue(controller._active)
        self.assertFalse(self.win.isVisible())

        controller._dispatch_pointer("press", 100, 15, press, True)
        controller._dispatch_pointer("move", 110, -245, outside, False)
        controller._dispatch_pointer("move", 100, 15, press, True)
        self.assertTrue(controller._active)
        controller._dispatch_pointer("release", 100, 15, press, True)
        self.assertEqual(self.win.display_mode, "taskbar")
        self.assertFalse(self.win.isVisible())

        controller._dispatch_pointer("press", 100, 15, press, True)
        controller._dispatch_pointer("move", 110, -245, outside, False)
        controller._dispatch_pointer("release", 110, -245, outside, False)
        self.assertEqual(self.win.display_mode, "float")
        self.assertTrue(self.win.isVisible())
        self.assertFalse(controller._active)
        self.assertIsNone(self.win._drag_pos)

    def test_native_small_movement_and_stationary_page_click_never_undock(self):
        controller = self.make_controller()
        self.win.set_view_options(taskbar_rows=3, taskbar_page_mode="manual")
        self.win.set_display_mode("taskbar")
        original = self.win.pos()
        press = QPoint(700, 760)
        controller._dispatch_pointer("press", 100, 15, press, True)
        controller._dispatch_pointer("move", 101, 15, press + QPoint(1, 0), True)
        controller._dispatch_pointer("release", 101, 15, press + QPoint(1, 0), True)
        self.assertEqual(self.win.pos(), original)
        self.assertEqual(self.win.display_mode, "taskbar")
        rect = controller._pager_rect
        controller._dispatch_pointer("press", rect.center().x(), rect.bottom() - 2, press, True)
        controller._dispatch_pointer("move", rect.center().x() + 1, rect.bottom() - 2, press + QPoint(1, 0), True)
        controller._dispatch_pointer("release", rect.center().x() + 1, rect.bottom() - 2, press + QPoint(1, 0), True)
        self.assertEqual(self.win.quotes.taskbar_page, 1)
        self.assertEqual(self.win.display_mode, "taskbar")
        self.assertFalse(self.win.isVisible())

    def test_shared_menu_has_independent_metrics_and_common_sort_and_hide(self):
        self.make_controller()
        self.win.set_display_mode("taskbar")
        menu = self.win.build_context_menu("taskbar")
        try:
            actions = {action.text(): action for action in menu.actions()}
            self.assertIn("排序", actions)
            self.assertIn("设置…", actions)
            metric_actions = {action.text(): action for action in actions["显示指标"].menu().actions()}
            self.assertTrue(metric_actions["成交量"].isEnabled())
            metric_actions["涨幅"].trigger()
            self.assertEqual(self.win.view_options.taskbar_metrics, ["name", "price"])
            self.assertEqual(self.win.visible_metrics, ["name", "price", "change_pct"])
            actions["隐藏"].trigger()
            self.assertFalse(self.win.widget_visible)
        finally:
            delete(menu)

    def test_drag_preview_leave_and_commit(self):
        controller = self.make_controller()
        original = self.win.pos()
        self.win._drag_start_window_pos = original
        controller.drag_started()
        self.assertFalse(controller._active)
        self.over.return_value = True
        controller.drag_moved()
        self.assertTrue(controller._active)
        self.assertFalse(self.win.isVisible())
        self.assertTrue(self.win.widget_visible)
        self.assertTrue(self.win.taskbar_preview_active)
        self.assertTrue(self.win.timer.isActive())
        self.native.capture_pointer.assert_called_once()
        self.assertEqual(self.win.display_mode, "float")
        self.over.return_value = False
        controller.drag_moved()
        self.assertFalse(controller._active)
        self.assertTrue(self.win.isVisible())
        self.assertTrue(self.win.widget_visible)
        self.over.return_value = True
        controller.drag_moved()
        self.win.move(40, 750)
        controller.drag_finished(True)
        self.assertEqual(self.win.display_mode, "taskbar")
        self.assertFalse(self.win.isVisible())
        self.assertEqual(self.win.pos(), original)

    def test_drag_cancel_and_dual_commit(self):
        controller = self.make_controller()
        self.win.set_view_options(taskbar_dual_open=True)
        origin = self.win.pos()
        self.win._drag_start_window_pos = origin
        controller.drag_started()
        self.over.return_value = True
        controller.drag_moved()
        self.win.move(0, 700)
        controller.drag_finished(False)
        self.assertEqual(self.win.display_mode, "float")
        self.assertEqual(self.win.pos(), origin)
        controller.drag_started()
        controller.drag_finished(True)
        self.assertEqual(self.win.display_mode, "both")
        self.assertTrue(self.win.isVisible())
        self.assertTrue(controller._active)

    def test_drag_events_preview_dock_and_escape_without_a_controller_shortcut(self):
        controller = self.make_controller()
        self.win.move(100, 100)
        origin = self.win.pos()
        target = self.win.table.viewport()

        def mouse(kind, global_pos, button, buttons):
            local = target.mapFromGlobal(global_pos)
            event = QMouseEvent(kind, QPointF(local), QPointF(global_pos), button, buttons, Qt.NoModifier)
            QApplication.sendEvent(target, event)

        start = target.mapToGlobal(QPoint(3, 3))
        mouse(QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
        self.over.return_value = True
        mouse(QEvent.MouseMove, start + QPoint(0, 100), Qt.NoButton, Qt.LeftButton)
        self.assertTrue(controller._active)
        self.assertFalse(self.win.isVisible())
        QApplication.sendEvent(self.win, QKeyEvent(QEvent.KeyPress, Qt.Key_Escape, Qt.NoModifier))
        self.assertEqual(self.win.pos(), origin)
        self.assertTrue(self.win.isVisible())
        self.assertFalse(controller._active)
        mouse(QEvent.MouseButtonRelease, start, Qt.LeftButton, Qt.NoButton)
        self.assertEqual(self.win.display_mode, "float")

        mouse(QEvent.MouseButtonPress, start, Qt.LeftButton, Qt.LeftButton)
        mouse(QEvent.MouseMove, start + QPoint(0, 100), Qt.NoButton, Qt.LeftButton)
        mouse(QEvent.MouseButtonRelease, start + QPoint(0, 100), Qt.LeftButton, Qt.NoButton)
        self.assertEqual(self.win.display_mode, "taskbar")
        self.assertFalse(self.win.isVisible())
        self.assertEqual(self.win.pos(), origin)

    def test_failed_drag_attachment_never_hides_float(self):
        controller = self.make_controller()
        self.native.present.side_effect = OSError("模拟失败")
        self.win._drag_start_window_pos = self.win.pos()
        controller.drag_started()
        self.over.return_value = True
        controller.drag_moved()
        controller.drag_finished(True)
        self.assertEqual(self.win.display_mode, "float")
        self.assertTrue(self.win.isVisible())
