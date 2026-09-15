"""在所有平台验证 Xlib 调用约定、异步错误处理和事件分发。"""

import ctypes
import os
import unittest
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from stockwidget.platform import click_through
from stockwidget.platform.hotkeys import (
    GlobalHotkeyManager, _XErrorHandler, _XErrorEvent, _XEvent,
    X_KEY_PRESS, X11_CONTROL_MASK, X11_MOD1_MASK, X11_LOCK_MASK,
)


class X11HotkeyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.manager = GlobalHotkeyManager()
        self.lib = Mock()
        self.lib.XStringToKeysym.return_value = 102
        self.lib.XKeysymToKeycode.return_value = 41
        self.lib.XGrabKey.return_value = 1  # 不是冲突标志
        self.lib.XPending.return_value = 0
        self.previous_errors = []
        self.previous = _XErrorHandler(
            lambda display, event: self.previous_errors.append(event.contents.error_code) or 0)
        self.handler = ctypes.cast(self.previous, ctypes.c_void_p).value
        self.original_handler = self.handler

        def set_handler(value):
            old = self.handler
            self.handler = value.value if isinstance(value, ctypes.c_void_p) else value
            return old

        self.lib.XSetErrorHandler.side_effect = set_handler
        self.manager._x11_lib = self.lib
        self.manager._x11_display = 1234
        self.manager._x11_root = 1
        self.manager._x11_notifier = Mock()

    def tearDown(self):
        self.manager.deleteLater()
        self.app.processEvents()

    def _error(self, code, display=1234):
        event = _XErrorEvent()
        event.error_code = code
        event.request_code = 33
        _XErrorHandler(self.handler)(display, ctypes.byref(event))

    def test_nonzero_return_registers_and_duplicate_is_rejected(self):
        self.assertTrue(self.manager._register_x11("Ctrl+Alt+F", Mock()))
        self.assertEqual(self.lib.XGrabKey.call_count, 8)
        self.assertEqual(self.handler, self.original_handler)
        self.assertEqual(self.manager._register_x11("Alt+Ctrl+F", Mock()).reason, "conflict")
        self.assertEqual(self.lib.XGrabKey.call_count, 8)

    def test_bare_function_keys_register_but_other_bare_keys_do_not(self):
        for number in range(1, 13):
            with self.subTest(number=number):
                self.lib.XGrabKey.reset_mock()
                self.assertTrue(self.manager._register_x11(f"F{number}", Mock()))
                self.assertEqual(self.lib.XGrabKey.call_count, 8)
                self.assertEqual(self.lib.XGrabKey.call_args_list[0].args[2], 0)
                self.manager._unregister_all_x11()
        for key in ("A", "1", "Space", "Delete", "F0", "F13", "F01", "Ctrl", ""):
            with self.subTest(key=key):
                self.lib.XGrabKey.reset_mock()
                self.assertEqual(self.manager._register_x11(key, Mock()).reason, "invalid")
                self.lib.XGrabKey.assert_not_called()

    def test_async_conflict_releases_partial_grabs_and_restores_handler(self):
        def sync(*_):
            if self.lib.XGrabKey.call_count and not self.lib.XUngrabKey.call_count:
                self._error(10)

        self.lib.XSync.side_effect = sync
        result = self.manager._register_x11("Ctrl+Alt+F", Mock())
        self.assertEqual(result.reason, "conflict")
        self.assertEqual(self.lib.XUngrabKey.call_count, 8)
        self.assertFalse(self.manager._x11_grabs)
        self.assertEqual(self.handler, self.original_handler)
        self.manager._x11_notifier.setEnabled.assert_not_called()

    def test_other_connection_errors_are_forwarded(self):
        def sync(*_):
            if self.lib.XGrabKey.call_count:
                self._error(3, display=5678)

        self.lib.XSync.side_effect = sync
        self.assertTrue(self.manager._register_x11("Ctrl+Alt+F", Mock()))
        self.assertEqual(self.previous_errors, [3])
        self.assertEqual(self.handler, self.original_handler)

    def test_xlib_queue_dispatches_keypress_and_ignores_release(self):
        callback = Mock()
        self.assertTrue(self.manager._register_x11("Ctrl+Alt+F", callback))
        self.lib.XPending.side_effect = [2, 1, 0]
        event_types = iter([X_KEY_PRESS, 3])

        def next_event(display, pointer):
            self.assertEqual(display, 1234)
            event = ctypes.cast(pointer, ctypes.POINTER(_XEvent)).contents
            event.xkey.type = next(event_types)
            event.xkey.keycode = 41
            event.xkey.state = X11_CONTROL_MASK | X11_MOD1_MASK | X11_LOCK_MASK

        self.lib.XNextEvent.side_effect = next_event
        self.manager._read_x11_events()
        callback.assert_called_once_with()
        self.lib.XPending.side_effect = None
        self.manager._unregister_all_x11()
        self.assertEqual(self.lib.XUngrabKey.call_count, 8)
        self.manager._x11_notifier.setEnabled.assert_called_with(False)


class X11ClickThroughTests(unittest.TestCase):
    def test_shape_abi_enable_disable_and_retry_after_missing_display(self):
        lib = Mock()
        lib.XOpenDisplay.side_effect = [None, 1234]
        # 使用 ctypes 函数校验参数数量和类型，防止 Mock 掩盖 ABI 错误。
        rectangles_calls, mask_calls = [], []
        rect_type = ctypes.CFUNCTYPE(None, ctypes.c_void_p, ctypes.c_ulong,
                                    ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                    ctypes.c_void_p, ctypes.c_int, ctypes.c_int, ctypes.c_int)
        mask_type = ctypes.CFUNCTYPE(None, ctypes.c_void_p, ctypes.c_ulong,
                                    ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                    ctypes.c_ulong, ctypes.c_int)
        ext = Mock()
        ext.XShapeCombineRectangles = rect_type(lambda *args: rectangles_calls.append(args))
        ext.XShapeCombineMask = mask_type(lambda *args: mask_calls.append(args))
        widget = Mock()
        widget.winId.return_value = 42
        with patch.object(click_through, "_x11_xlib", None), patch.object(
            click_through, "_x11_xext", None
        ), patch.object(click_through, "_x11_shape_display", None), patch.object(
            click_through.ctypes, "CDLL", side_effect=lambda name: ext if "Xext" in name else lib
        ), patch.object(click_through.QGuiApplication, "sync"):
            click_through._click_through_x11(widget, True)
            self.assertFalse(rectangles_calls)
            click_through._click_through_x11(widget, True)
            click_through._click_through_x11(widget, False)
        self.assertEqual(rectangles_calls, [(1234, 42, 2, 0, 0, None, 0, 0, 0)])
        self.assertEqual(mask_calls, [(1234, 42, 2, 0, 0, 0, 0)])
        self.assertEqual(lib.XSync.call_count, 2)
