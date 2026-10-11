"""Opt-in Cocoa check of the actual NSWindow mouse-event state and restoration."""

import os
import sys
import unittest
from unittest.mock import Mock, patch

from PySide6.QtCore import QPoint
from shiboken6 import delete

from stockwidget.ui.floating.presenter import QuotePresenter
from stockwidget.ui.floating.widget import FloatLabel
from tests.support import QtTestCase


@unittest.skipUnless(sys.platform == "darwin" and os.environ.get("STOCKWIDGET_TEST_MACOS") == "1",
                     "Requires opt-in macOS Cocoa desktop session")
class CocoaClickThroughTests(QtTestCase):
    def test_native_mouse_transparency_survives_topmost_hide_and_restore(self):
        import ctypes
        import ctypes.util

        self.assertEqual(self.app.platformName(), "cocoa")
        objc = ctypes.CDLL(ctypes.util.find_library("objc"))
        objc.sel_registerName.argtypes = [ctypes.c_char_p]
        objc.sel_registerName.restype = ctypes.c_void_p
        send_pointer = ctypes.CFUNCTYPE(ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p)(
            ("objc_msgSend", objc))
        send_bool = ctypes.CFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)(
            ("objc_msgSend", objc))

        def ignores_mouse(widget):
            native_window = send_pointer(int(widget.winId()), objc.sel_registerName(b"window"))
            self.assertTrue(native_window)
            return send_bool(native_window, objc.sel_registerName(b"ignoresMouseEvents"))

        with patch.object(QuotePresenter, "refresh"), \
                patch("stockwidget.ui.floating.widget.GlobalHotkeyManager", return_value=Mock()):
            window = FloatLabel({}, {})
            try:
                window.move(self.app.primaryScreen().availableGeometry().topLeft() + QPoint(100, 100))
                window.show()
                self.app.processEvents()
                geometry = window.geometry()
                for enabled in (True, False, True):
                    window.set_click_through(enabled)
                    self.app.processEvents()
                    self.assertEqual(ignores_mouse(window), enabled)
                    for on_top in (False, True):
                        window.set_float_on_top(on_top)
                        self.app.processEvents()
                        self.assertEqual(ignores_mouse(window), enabled)
                        self.assertEqual(window.geometry(), geometry)
                    window.hide_widget()
                    window.set_float_on_top(False)
                    self.assertFalse(window.isVisible())
                    window.toggle_win()
                    self.app.processEvents()
                    self.assertEqual(ignores_mouse(window), enabled)
                    self.assertEqual(window.geometry(), geometry)
            finally:
                delete(window)
                self.app.processEvents()
