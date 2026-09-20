"""穿透与强制置顶可同时开启；隐藏和编辑设置时仍应避让。"""

import os
import sys
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QWidget
from shiboken6 import delete

from stockwidget.platform.click_through import apply_click_through
from stockwidget.platform.windows import ensure_topmost, get_user32
from stockwidget.ui.widget import FloatLabel


class WidgetTopmostTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        with patch.object(FloatLabel, "_refresh_from_function"), patch(
            "stockwidget.ui.widget.GlobalHotkeyManager"
        ):
            self.window = FloatLabel({}, {})
        for name, kwargs in (
            ("ensure_topmost", {}),
            ("apply_click_through", {}),
            ("QApplication.activeWindow", {"return_value": None}),
            ("QApplication.activePopupWidget", {"return_value": None}),
        ):
            mock = self.enterContext(patch("stockwidget.ui.widget." + name, **kwargs))
            if name == "ensure_topmost":
                self.native_top = mock
        self.window.force_top = True
        self.window.show()
        self.native_top.reset_mock()

    def tearDown(self):
        delete(self.window)

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
                    "stockwidget.ui.widget.QApplication." + name, return_value=other
                ):
                    self.window._ensure_on_top()
                    self.native_top.assert_not_called()
        finally:
            delete(other)


@unittest.skipUnless(
    sys.platform == "win32" and os.environ.get("STOCKWIDGET_TEST_WINDOWS") == "1",
    "Requires an opt-in Windows desktop session",
)
class WindowsTopmostIntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

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
