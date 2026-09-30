"""穿透与强制置顶可同时开启；隐藏和编辑设置时仍应避让。"""

import os
import sys
import unittest
from unittest.mock import Mock, patch

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

    def test_default_and_saved_topmost_configuration(self):
        self.assertTrue(self.window.float_on_top)
        self.assertTrue(self.window.windowFlags() & Qt.WindowStaysOnTopHint)
        for cfg in ({"force_top": True}, {"float_on_top": False, "force_top": True}):
            with self.subTest(cfg=cfg), patch.object(FloatLabel, "_refresh_from_function"), patch(
                "stockwidget.ui.widget.GlobalHotkeyManager"
            ), patch("stockwidget.ui.widget.force_top_supported", return_value=True):
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
        with patch("stockwidget.ui.widget.force_top_supported", return_value=True):
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
        with patch.object(FloatLabel, "_refresh_from_function"), patch(
            "stockwidget.ui.widget.GlobalHotkeyManager"
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

    def test_float_topmost_toggle_changes_native_style_without_hiding_settings(self):
        from ctypes import wintypes

        from stockwidget.ui.settings_dialog import SettingsDialog

        self.assertEqual(self.app.platformName(), "windows")
        api = get_user32()
        api.GetForegroundWindow.argtypes = []
        api.GetForegroundWindow.restype = wintypes.HWND
        with patch.object(FloatLabel, "_refresh_from_function"), patch(
            "stockwidget.ui.widget.GlobalHotkeyManager"
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
