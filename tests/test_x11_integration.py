"""真实 X11 集成测试，仅在显式指定的隔离 Xvfb 会话运行。

QT_QPA_PLATFORM=xcb STOCKWIDGET_TEST_X11=1 xvfb-run -a python -m unittest tests.test_x11_integration
"""

import ctypes
import os
import sys
import unittest
from unittest.mock import Mock

from PySide6.QtCore import Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QPushButton
from shiboken6 import delete

from stockwidget.platform.click_through import apply_click_through
from stockwidget.platform.hotkeys import (
    GlobalHotkeyManager, X11_CONTROL_MASK, X11_MOD1_MASK, X11_LOCK_MASK,
)
from stockwidget.ui.hotkey_sequence_edit import HotkeySequenceEdit
from PySide6.QtGui import QKeySequence


@unittest.skipUnless(sys.platform == "linux" and os.environ.get("STOCKWIDGET_TEST_X11") == "1",
                     "requires an explicitly enabled isolated Xvfb session")
class X11IntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        if cls.app.platformName() != "xcb":
            raise RuntimeError("X11 integration tests require QT_QPA_PLATFORM=xcb")

    def setUp(self):
        self.managers = []
        self.manager = self._manager()
        self.lib = self.manager._ensure_x11()
        self.assertIsNotNone(self.lib)
        self.injector = self.lib.XOpenDisplay(None)
        self.xtest = ctypes.CDLL("libXtst.so.6")
        self.xtest.XTestFakeKeyEvent.argtypes = [
            ctypes.c_void_p, ctypes.c_uint, ctypes.c_int, ctypes.c_ulong]
        self.xtest.XTestFakeButtonEvent.argtypes = [
            ctypes.c_void_p, ctypes.c_uint, ctypes.c_int, ctypes.c_ulong]
        self.xtest.XTestFakeMotionEvent.argtypes = [
            ctypes.c_void_p, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_ulong]

    def _manager(self):
        manager = GlobalHotkeyManager()
        self.managers.append(manager)
        return manager

    def tearDown(self):
        self.lib.XCloseDisplay(self.injector)
        for manager in self.managers:
            manager._unregister_all_x11()
            delete(manager)

    def _send_hotkey(self, names=(b"Control_L", b"Alt_L", b"f")):
        keys = [self.lib.XKeysymToKeycode(self.injector, self.lib.XStringToKeysym(name))
                for name in names]
        for key in keys:
            self.xtest.XTestFakeKeyEvent(self.injector, key, True, 0)
        for key in reversed(keys):
            self.xtest.XTestFakeKeyEvent(self.injector, key, False, 0)
        self.lib.XSync(self.injector, False)

    def test_registered_key_reaches_qt_callback(self):
        callback = Mock()
        self.assertTrue(self.manager.register("Ctrl+Alt+F", callback))
        self._send_hotkey()
        for _ in range(100):
            if callback.called:
                break
            QTest.qWait(10)
        callback.assert_called_once_with()

    def test_formerly_blocked_combinations_reach_callbacks(self):
        for combination, names in (
            ("Ctrl+G", (b"Control_L", b"g")),
            ("Alt+G", (b"Alt_L", b"g")),
            ("Ctrl+Shift+G", (b"Control_L", b"Shift_L", b"g")),
            ("Meta+G", (b"Super_L", b"g")),
        ):
            with self.subTest(combination=combination):
                callback = Mock()
                self.assertTrue(self.manager.register(combination, callback))
                self._send_hotkey(names)
                for _ in range(100):
                    if callback.called:
                        break
                    QTest.qWait(10)
                callback.assert_called_once_with()
                self.manager._unregister_all_x11()

    def test_all_bare_function_keys_reach_callbacks(self):
        for number in range(1, 13):
            with self.subTest(number=number):
                callback = Mock()
                self.assertTrue(self.manager.register(f"F{number}", callback))
                self._send_hotkey((f"F{number}".encode(),))
                for _ in range(100):
                    if callback.called:
                        break
                    QTest.qWait(10)
                callback.assert_called_once_with()
                self.manager._unregister_all_x11()

    def test_repeating_registered_key_in_focused_editor_keeps_sequence(self):
        editor = HotkeySequenceEdit()
        editor.setMaximumSequenceLength(1)
        editor.setKeySequence(QKeySequence("Ctrl+Alt+F"))
        editor.show()
        editor.activateWindow()
        editor.setFocus()
        QTest.qWait(50)
        try:
            self.assertTrue(editor.hasFocus())
            callback = Mock()
            self.assertTrue(self.manager.register("Ctrl+Alt+F", callback))
            self._send_hotkey()
            QTest.qWait(1100)
            self.assertEqual(editor.keySequence().toString(), "Ctrl+Alt+F")
            callback.assert_called_once_with()
            self.assertTrue(editor.hasFocus())
        finally:
            delete(editor)

    def test_partial_conflict_is_reported_and_successful_grabs_are_released(self):
        blocker = self.manager
        key = self.lib.XKeysymToKeycode(self.injector, self.lib.XStringToKeysym(b"f"))
        mask = X11_CONTROL_MASK | X11_MOD1_MASK | X11_LOCK_MASK
        self.lib.XGrabKey(blocker._x11_display, key, mask, blocker._x11_root, False, 1, 1)
        self.lib.XSync(blocker._x11_display, False)
        try:
            contender = self._manager()
            result = contender.register("Ctrl+Alt+F", Mock())
            self.assertEqual(result.reason, "conflict")
        finally:
            self.lib.XUngrabKey(blocker._x11_display, key, mask, blocker._x11_root)
            self.lib.XSync(blocker._x11_display, False)
        verifier = self._manager()
        self.assertTrue(verifier.register("Ctrl+Alt+F", Mock()))

    def test_mouse_reaches_underlying_window_and_returns_after_disable(self):
        lower, upper = QPushButton("lower"), QPushButton("upper")
        lower_clicks, upper_clicks = Mock(), Mock()
        lower.clicked.connect(lower_clicks)
        upper.clicked.connect(upper_clicks)
        for window in (lower, upper):
            window.setWindowFlags(Qt.Window | Qt.FramelessWindowHint)
            window.setGeometry(100, 100, 200, 100)
            window.show()
        upper.raise_()
        QTest.qWait(50)

        def click():
            self.xtest.XTestFakeMotionEvent(self.injector, -1, 150, 150, 0)
            self.xtest.XTestFakeButtonEvent(self.injector, 1, True, 0)
            self.xtest.XTestFakeButtonEvent(self.injector, 1, False, 0)
            self.lib.XSync(self.injector, False)
            QTest.qWait(50)

        try:
            click()
            self.assertEqual(upper_clicks.call_count, 1)
            apply_click_through(upper, True)
            upper.resize(220, 110)
            click()
            self.assertEqual(lower_clicks.call_count, 1)
            self.assertEqual(upper_clicks.call_count, 1)
            apply_click_through(upper, False)
            click()
            self.assertEqual(upper_clicks.call_count, 2)
            self.assertEqual(lower_clicks.call_count, 1)
        finally:
            delete(upper)
            delete(lower)
