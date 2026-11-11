"""Windows 任务栏枚举、屏幕匹配与原生窗口迁移的 API 回归。"""

from unittest.mock import Mock, patch
from types import SimpleNamespace
import ctypes
import os
import sys
import unittest

from PySide6.QtCore import QPoint
from PySide6.QtGui import QCursor
from PySide6.QtTest import QTest
from shiboken6 import delete

from stockwidget.data.quotes import _new_entry
from stockwidget.platform.taskbar import NativeTaskbarWindow, TaskbarArea, _api, _taskbar_windows, cursor_over_taskbar, find_taskbar
from stockwidget.app import App
from stockwidget.ui.floating.presenter import QuotePresenter
from stockwidget.ui.floating.taskbar import TaskbarController
from stockwidget.ui.floating.widget import FloatLabel
from stockwidget.ui.settings.dialog import SettingsDialog
from tests.support import QtTestCase, SettingsTestCase


@unittest.skipUnless(sys.platform == "win32" and os.environ.get("STOCKWIDGET_TEST_WINDOWS") == "1",
                     "Requires an opt-in Windows desktop session")
class NativeSettingsWindowTests(SettingsTestCase):
    def test_settings_has_an_unowned_non_topmost_taskbar_window_with_application_icon(self):
        from ctypes import wintypes
        from PySide6.QtCore import Qt
        from PySide6.QtGui import QIcon, QPixmap
        from stockwidget.platform.window import get_user32

        original = self.qt_app.windowIcon()
        self.addCleanup(self.qt_app.setWindowIcon, original)
        pixmap = QPixmap(20, 20)
        pixmap.fill(Qt.blue)
        self.qt_app.setWindowIcon(QIcon(pixmap))
        dialog, window = self._make_dialog(float_on_top=True)
        window.show()
        dialog.show()
        self.qt_app.processEvents()
        user = get_user32()
        user.GetWindow.argtypes = [wintypes.HWND, wintypes.UINT]
        user.GetWindow.restype = wintypes.HWND
        hwnd = int(dialog.winId())
        self.assertFalse(user.GetWindow(hwnd, 4))  # GW_OWNER
        extended_style = user.GetWindowLongW(hwnd, -20)
        self.assertFalse(extended_style & 0x00000080)  # WS_EX_TOOLWINDOW
        self.assertFalse(extended_style & 0x00000008)  # WS_EX_TOPMOST
        self.assertFalse(dialog.windowIcon().isNull())
        dialog.showMinimized()
        self.qt_app.processEvents()
        self.assertTrue(dialog.isMinimized())
        self.assertTrue(window.isVisible())


@unittest.skipUnless(sys.platform == "win32" and os.environ.get("STOCKWIDGET_TEST_WINDOWS") == "1",
                     "Requires an opt-in Windows desktop session")
class NativeTaskbarRefreshTests(QtTestCase):
    def test_collapsed_float_keeps_native_taskbar_quotes_updating_without_hover(self):
        codes = {"sh600000": {"market": "sh", "code": "600000", "checked": True}}

        def payload(price):
            quote = _new_entry("测试", 10, 10, price, price, 9, 100, 1000,
                               [0] * 5, [0] * 5, [0] * 5, [0] * 5, "2026-11-11", "10:00:00")
            return True, {"sh600000": quote}, None

        with patch.object(QuotePresenter, "refresh"), patch("stockwidget.ui.floating.widget.GlobalHotkeyManager"):
            source = FloatLabel({"pos": {"x": 100, "y": 100}, "taskbar_enabled": True,
                                 "display_mode": "both", "boundary_check_enabled": True,
                                 "edge_hide_enabled": True, "watchlist": codes,
                                 "visible_metrics": ["price"], "taskbar_sync_metrics": True}, {})
        controller = TaskbarController(source, Mock())
        try:
            source.timer.timeout.disconnect()
            source.timer.timeout.connect(source.quotes.refresh)
            source.timer.setInterval(80)
            with patch("stockwidget.ui.floating.presenter.fetch_quote_result", return_value=payload(10)) as fetch:
                controller.apply_mode()
                for _ in range(40):
                    QTest.qWait(25)
                    if source.quotes.quote_for(codes["sh600000"]) is not None:
                        break
                self.assertTrue(controller._active, source.taskbar_status)
                self.assertIsNotNone(source.quotes.quote_for(codes["sh600000"]))
                bounds = source.screen().availableGeometry()
                source.move(bounds.left(), bounds.top() + 100)
                self.app.processEvents()
                self.assertTrue(source.position_controller.collapsed)
                self.assertTrue(source.timer.isActive())
                before = controller.native._last_frame
                fetch.return_value = payload(111)
                fetch.reset_mock()
                for _ in range(40):
                    QTest.qWait(25)
                    if controller.native._last_frame != before:
                        break
                fetch.assert_called()
                self.assertNotEqual(before, controller.native._last_frame)
                self.assertEqual(source.quotes.quote_for(codes["sh600000"])["current_price"], 111)
                self.assertTrue(controller.native.user.IsWindowVisible(controller.native.hwnd))
                self.assertTrue(source.position_controller.collapsed)
                self.assertEqual(source.width(), 5)
                source.hide_widget()
                self.assertFalse(source.timer.isActive())
                self.assertFalse(controller._active)
                fetch.reset_mock()
                QTest.qWait(200)
                fetch.assert_not_called()
        finally:
            controller.close()
            delete(controller)
            delete(source)
            self.app.processEvents()


@unittest.skipUnless(sys.platform == "win32" and os.environ.get("STOCKWIDGET_TEST_WINDOWS") == "1",
                     "Requires an opt-in Windows desktop session")
class NativeTaskbarDragTests(QtTestCase):
    def setUp(self):
        from ctypes import wintypes as w

        self.enterContext(patch.object(QuotePresenter, "refresh"))
        self.enterContext(patch("stockwidget.ui.floating.widget.GlobalHotkeyManager"))
        self.user, _, _ = _api()
        self.user.mouse_event.argtypes = [w.DWORD, w.DWORD, w.DWORD, w.DWORD, ctypes.c_size_t]
        self.user.mouse_event.restype = None
        self.user.WindowFromPoint.argtypes = [w.POINT]
        self.user.WindowFromPoint.restype = w.HWND
        self.pointer = w.POINT()
        self.assertTrue(self.user.GetCursorPos(ctypes.byref(self.pointer)))
        self.source = FloatLabel({"pos": {"x": 100, "y": 100}, "taskbar_enabled": True}, {})
        self.controller = TaskbarController(self.source, Mock())
        self.controller.apply_mode()
        self.source.raise_()
        self.source.activateWindow()
        self.app.processEvents()
        area = find_taskbar(window=int(self.source.winId()))
        self.assertIsNotNone(area)
        rect = self.rect(area.hwnd)
        self.bar_point = QPoint((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
        self.outside = QPoint(self.bar_point.x(), rect.top - 200)

    def tearDown(self):
        self.user.mouse_event(4, 0, 0, 0, 0)
        self.app.processEvents()
        self.controller.close()
        delete(self.controller)
        delete(self.source)
        self.user.SetCursorPos(self.pointer.x, self.pointer.y)
        self.app.processEvents()

    def rect(self, hwnd):
        from ctypes import wintypes as w

        rect = w.RECT()
        self.assertTrue(self.user.GetWindowRect(hwnd, ctypes.byref(rect)))
        return rect

    def move_pointer(self, point):
        self.assertTrue(self.user.SetCursorPos(point.x(), point.y()))
        QTest.qWait(120)

    def press(self, hwnd):
        from ctypes import wintypes as w

        rect = self.rect(hwnd)
        point = QPoint((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
        self.move_pointer(point)
        self.assertEqual(self.user.WindowFromPoint(w.POINT(point.x(), point.y())), hwnd)
        self.user.mouse_event(2, 0, 0, 0, 0)
        QTest.qWait(80)

    def test_one_physical_drag_can_enter_leave_and_reenter_the_taskbar(self):
        self.press(int(self.source.winId()))
        self.assertIsNotNone(self.source._drag_pos)
        for cycle in range(3):
            self.move_pointer(self.bar_point)
            self.assertTrue(self.controller._active, (cycle, self.source.taskbar_status, self.source.pos(),
                                                     self.source._dragging, self.user.GetCapture()))
            self.assertFalse(self.source.isVisible())
            self.assertTrue(self.source._dragging)
            self.assertEqual(self.user.GetCapture(), self.controller.native.hwnd)
            self.move_pointer(self.outside)
            self.assertTrue(self.source.isVisible())
            self.assertTrue(self.source._dragging)
            self.assertFalse(self.controller._active)
        self.user.mouse_event(4, 0, 0, 0, 0)
        QTest.qWait(120)
        self.assertEqual(self.source.display_mode, "float")
        self.assertTrue(self.source.isVisible())
        self.assertIsNone(self.source._drag_pos)
        self.assertIsNone(self.user.GetCapture())

    def test_dual_mode_fixes_the_taskbar_and_freely_drags_the_float_over_it(self):
        self.source.set_display_mode("both")
        self.app.processEvents()
        native = self.controller.native
        parent, origin = native.parent, self.source.pos()
        self.press(native.hwnd)
        self.assertIsNone(self.source._drag_pos)
        points = [self.outside, self.bar_point]
        for hwnd in _taskbar_windows(self.user):
            if hwnd != parent:
                rect = self.rect(hwnd)
                points.append(QPoint((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2))
        for point in points:
            self.move_pointer(point)
            self.assertFalse(self.controller._dragging)
            self.assertEqual(native.parent, parent)
        self.user.mouse_event(4, 0, 0, 0, 0)
        QTest.qWait(120)
        self.assertEqual(native.parent, parent)
        self.assertEqual(self.source.pos(), origin)
        self.assertIsNone(self.source._drag_pos)
        self.press(int(self.source.winId()))
        for point in points:
            self.move_pointer(point)
            self.assertTrue(self.source.isVisible())
            self.assertTrue(self.source.widget_visible)
            self.assertTrue(self.controller._active)
            self.assertFalse(self.controller._dragging)
            self.assertFalse(self.source.taskbar_preview_active)
            self.assertTrue(self.source.geometry().contains(QCursor.pos()))
            self.assertEqual(native.parent, parent)
        self.user.mouse_event(4, 0, 0, 0, 0)
        QTest.qWait(120)
        self.assertEqual(self.source.display_mode, "both")
        self.assertTrue(self.source.isVisible())


class TaskbarMonitorTests(unittest.TestCase):
    def setUp(self):
        self.user = Mock()
        self.enterContext(patch("stockwidget.platform.taskbar._api", return_value=(self.user, None, None)))
        self.user.FindWindowW.return_value = 100
        def find_child(parent, previous, class_name, title):
            if parent is None and class_name == "Shell_SecondaryTrayWnd":
                return 200 if previous is None else None
            return 101 if parent == 100 and class_name == "TrayNotifyWnd" else None
        self.user.FindWindowExW.side_effect = find_child
        self.user.MonitorFromWindow.side_effect = lambda hwnd, flags: 2 if hwnd in (200, 900) else 1
        def monitor_info(monitor, pointer):
            pointer._obj.device = rf"\\.\DISPLAY{monitor}"
            return True
        self.user.GetMonitorInfoW.side_effect = monitor_info
        def client_rect(hwnd, pointer):
            pointer._obj.right = 1920
            pointer._obj.bottom = 72 if hwnd == 200 else 48
            return True
        self.user.GetClientRect.side_effect = client_rect
        def window_rect(hwnd, pointer):
            left, top, right, bottom = {100: (0, 1032, 1920, 1080), 101: (1700, 1032, 1920, 1080),
                                       200: (-1920, -72, 0, 0)}[hwnd]
            pointer._obj.left, pointer._obj.top = left, top
            pointer._obj.right, pointer._obj.bottom = right, bottom
            return True
        self.user.GetWindowRect.side_effect = window_rect
        self.user.GetDpiForWindow.side_effect = lambda hwnd: 144 if hwnd == 200 else 96

    def test_primary_secondary_and_removed_displays_choose_the_expected_taskbar(self):
        for kwargs, expected in (({}, 100), ({"window": 900}, 200),
                                 ({"device_name": r"\\.\DISPLAY2"}, 200),
                                 ({"device_name": "removed", "window": 900}, 200),
                                 ({"device_name": "removed"}, 100), ({"hwnd": 200}, 200)):
            with self.subTest(kwargs=kwargs):
                area = find_taskbar(**kwargs)
                self.assertEqual(area.hwnd, expected)
                self.assertEqual(area.dpi, 144 if expected == 200 else 96)
                self.assertEqual(area.device_name, rf"\\.\DISPLAY{2 if expected == 200 else 1}")
                self.assertGreater(area.right, 0)
                self.assertLess(area.right, area.width)
        self.assertIsNone(find_taskbar(hwnd=999))
        self.user.FindWindowW.return_value = None
        self.user.FindWindowExW.return_value = None
        self.user.FindWindowExW.side_effect = None
        self.assertIsNone(find_taskbar())

    def test_pointer_hit_checks_secondary_taskbars_with_negative_native_coordinates(self):
        with patch("stockwidget.platform.taskbar.sys.platform", "win32"):
            for position, expected in (((500, 1050), 100), ((-500, -40), 200), ((-500, 200), None)):
                with self.subTest(position=position):
                    def cursor(pointer):
                        pointer._obj.x, pointer._obj.y = position
                        return True
                    self.user.GetCursorPos.side_effect = cursor
                    self.assertEqual(cursor_over_taskbar(), expected)

    def test_reparent_preserves_the_existing_input_owner_and_pointer_capture(self):
        native = NativeTaskbarWindow.__new__(NativeTaskbarWindow)
        native.user = self.user
        native.hwnd, native.parent, native._last_frame = 50, 100, (1, 1, b"frame")
        native._pointer_down = True
        self.user.IsWindow.return_value = True
        self.user.GetParent.side_effect = [100, 200]
        with patch.object(ctypes, "set_last_error", create=True), \
                patch.object(ctypes, "get_last_error", return_value=0, create=True):
            native._attach(TaskbarArea(200, 1920, 72, 1700, 144))
        self.assertEqual((native.hwnd, native.parent), (50, 200))
        self.assertTrue(native._pointer_down)
        self.user.SetParent.assert_called_once_with(50, 200)
        self.user.CreateWindowExW.assert_not_called()
        self.user.DestroyWindow.assert_not_called()
        self.user.ReleaseCapture.assert_not_called()


@unittest.skipUnless(sys.platform == "win32" and os.environ.get("STOCKWIDGET_TEST_WINDOWS") == "1",
                     "Requires opt-in Windows desktop session with multiple taskbars")
class NativeMultiscreenTaskbarTests(QtTestCase):
    def test_native_cross_screen_capture_enable_target_and_settings_placement(self):
        from ctypes import wintypes as w

        user, _, _ = _api()
        areas = [find_taskbar(hwnd=hwnd) for hwnd in _taskbar_windows(user)]
        if len(areas) < 2:
            self.skipTest("Requires at least two Windows taskbars")
        areas = areas[:2]
        screens = [next((s for s in self.app.screens()
                         if (s.geometry().x(), s.geometry().y()) == area.monitor_rect[:2]), None)
                   for area in areas]
        self.assertTrue(all(screens))
        with patch.object(QuotePresenter, "refresh"), \
                patch("stockwidget.ui.floating.widget.GlobalHotkeyManager"):
            source = FloatLabel({"taskbar_enabled": True}, {})
        controller = TaskbarController(source, Mock())
        dialog = None
        try:
            source.move(screens[1].geometry().topLeft() + QPoint(100, 100))
            source.show()
            self.app.processEvents()
            source.set_display_mode("taskbar")
            self.assertTrue(controller._active, source.taskbar_status)
            self.assertEqual(source.taskbar_screen, areas[1].device_name)
            self.assertIs(controller.screen(), screens[1])
            source.set_display_mode("float")
            source.move(screens[0].geometry().topLeft() + QPoint(100, 100))
            self.app.processEvents()
            source.set_display_mode("taskbar")
            self.assertEqual(source.taskbar_screen, areas[0].device_name)
            origin = source.pos()
            user.SendMessageW.argtypes = [w.HWND, w.UINT, w.WPARAM, w.LPARAM]
            user.SendMessageW.restype = ctypes.c_ssize_t
            self.enterContext(patch.object(NativeTaskbarWindow, "poll_pointer", return_value=None))

            for target_index, accepted in ((1, True), (0, False)):
                previous_screen = source.taskbar_screen
                previous_bar = controller.native.parent
                hwnd = controller.native.hwnd
                for message, bar, position in (
                    (0x201, previous_bar, source.pos() + QPoint(5, 5)),
                    (0x200, areas[target_index].hwnd, screens[target_index].geometry().topLeft() + QPoint(300, 300)),
                ):
                    with patch("stockwidget.ui.floating.taskbar.QCursor.pos", return_value=position), \
                            patch("stockwidget.ui.floating.taskbar.cursor_over_taskbar", return_value=bar):
                        user.SendMessageW(hwnd, message, 1, (20 << 16) | 100)
                        self.app.processEvents()
                self.assertEqual(controller.native.hwnd, hwnd)
                self.assertEqual(user.GetParent(hwnd), areas[target_index].hwnd)
                self.assertEqual(user.GetCapture(), hwnd)
                self.assertTrue(source._dragging)
                self.assertFalse(source.isVisible())
                self.assertEqual(source.taskbar_screen, areas[target_index].device_name)
                with patch("stockwidget.ui.floating.taskbar.QCursor.pos", return_value=position), \
                        patch("stockwidget.ui.floating.taskbar.cursor_over_taskbar", return_value=areas[target_index].hwnd):
                    user.SendMessageW(hwnd, 0x202 if accepted else 0x100, 0 if accepted else 0x1B, (20 << 16) | 100)
                    self.app.processEvents()
                self.assertIsNone(user.GetCapture())
                self.assertEqual(source.taskbar_screen, areas[target_index].device_name if accepted else previous_screen)
                self.assertEqual(source.display_mode, "taskbar")
                self.assertEqual(source.pos(), origin)

            source.set_display_mode("both")
            owner = SimpleNamespace(win=source, taskbar=controller, settings_dlg=None)
            with patch.object(SettingsDialog, "_start_github_check"), \
                    patch("stockwidget.app.SettingsDialog", side_effect=lambda win, app: SettingsDialog(win)):
                App.open_settings(owner)
            dialog = owner.settings_dlg
            self.app.processEvents()
            self.assertIs(dialog.windowHandle().screen(), screens[0])
            source.set_display_mode("taskbar")
            App.open_settings(owner)
            self.app.processEvents()
            self.assertIs(dialog.windowHandle().screen(), screens[1])
            self.assertTrue(screens[1].availableGeometry().contains(dialog.frameGeometry().center()))
        finally:
            if dialog is not None:
                delete(dialog)
            controller.close()
            delete(controller)
            delete(source)
            self.app.processEvents()
