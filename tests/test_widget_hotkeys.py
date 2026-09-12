"""快捷键配置共用接口的成功、停用和失败回滚行为。"""

import os
import unittest
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from stockwidget.platform.hotkeys import HotkeyResult
from stockwidget.ui.widget import FloatLabel


class WidgetHotkeyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        with patch.object(FloatLabel, "_refresh_from_function"), patch(
            "stockwidget.ui.widget.GlobalHotkeyManager"
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

    def test_failed_update_restores_both_registered_hotkeys(self):
        self.window.hotkey_enabled = True
        self.window.hotkey_click_through_enabled = True
        self.manager.register.side_effect = [
            HotkeyResult(False, "conflict"), HotkeyResult(True), HotkeyResult(True),
        ]

        result = self.window.update_hotkey("Ctrl+Alt+X")

        self.assertEqual(result.reason, "conflict")
        self.assertEqual(self.window.hotkey, "Ctrl+Alt+F")
        self.assertEqual(
            [c.args[0] for c in self.manager.register.call_args_list],
            ["Ctrl+Alt+X", "Ctrl+Alt+F", "Ctrl+Alt+C"],
        )
        self.saved.assert_not_called()

    def test_failed_enable_preserves_other_hotkey(self):
        self.window.hotkey_enabled = True
        self.manager.register.side_effect = [
            HotkeyResult(True), HotkeyResult(False, "conflict"), HotkeyResult(True),
        ]

        self.assertFalse(self.window.set_click_through_hotkey_enabled(True))

        self.assertFalse(self.window.hotkey_click_through_enabled)
        self.assertTrue(self.window.hotkey_enabled)
        self.assertEqual(
            [c.args[0] for c in self.manager.register.call_args_list],
            ["Ctrl+Alt+F", "Ctrl+Alt+C", "Ctrl+Alt+F"],
        )
        self.saved.assert_not_called()

    def test_successful_enable_saves_once_and_repeated_value_is_noop(self):
        self.manager.register.return_value = HotkeyResult(True)
        self.assertTrue(self.window.set_hotkey_enabled(True))
        self.assertTrue(self.window.set_hotkey_enabled(True))
        self.manager.register.assert_called_once()
        self.saved.assert_called_once_with()
