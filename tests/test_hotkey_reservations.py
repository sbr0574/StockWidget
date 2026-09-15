"""普通组合交给原生接口判断；保留规则不得跨平台或按整类修饰键误拦截。"""

import os
import unittest
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication
from shiboken6 import delete

from stockwidget.platform.hotkeys import (
    GlobalHotkeyManager, HotkeyResult, is_reserved, _parse_hotkey,
    _parse_hotkey_macos, MOD_NOREPEAT,
)


class HotkeyReservationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_bare_function_keys_register_without_modifiers_on_windows(self):
        manager = GlobalHotkeyManager()
        manager._filter = Mock()  # 只验证原生注册参数，不安装 Windows 事件过滤器。
        user32 = Mock()
        user32.RegisterHotKey.return_value = 1
        try:
            with patch("stockwidget.platform.hotkeys.ctypes.WinDLL", return_value=user32, create=True):
                for number in range(1, 13):
                    with self.subTest(number=number):
                        self.assertEqual(_parse_hotkey(f" f{number} "), (0, 0x6F + number))
                        self.assertTrue(manager._register_windows(f"F{number}", Mock()))
                        self.assertEqual(user32.RegisterHotKey.call_args.args[2:],
                                         (MOD_NOREPEAT, 0x6F + number))
        finally:
            delete(manager)

    def test_macos_requires_modifiers_for_every_function_key(self):
        manager = GlobalHotkeyManager()
        try:
            for number in range(1, 13):
                with self.subTest(number=number):
                    self.assertIsNone(_parse_hotkey_macos(f"F{number}"))
                    self.assertEqual(manager._register_macos(f"F{number}", Mock()).reason, "invalid")
                    self.assertIsNotNone(_parse_hotkey_macos(f"Ctrl+F{number}"))
        finally:
            delete(manager)

    def test_other_bare_keys_remain_invalid(self):
        for key in ("A", "1", "Space", "Tab", "Delete", "F0", "F13", "F01", "Ctrl", ""):
            with self.subTest(key=key):
                self.assertIsNone(_parse_hotkey(key))
                self.assertIsNone(_parse_hotkey_macos(key))

    def test_common_combinations_reach_each_platform_registration(self):
        combinations = (
            "Ctrl+C", "Ctrl+9", "Ctrl+Space", "Alt+F", "Alt+1",
            "Ctrl+Shift+F", "Ctrl+Shift+S", "Meta+G", "Win+Alt+G",
            "Alt+Tab", "Alt+F4", "Alt+Esc", "Alt+Space", "Ctrl+Esc",
            "Ctrl+Shift+Esc", "Ctrl+Alt+Shift+Delete",
        )
        manager = GlobalHotkeyManager()
        callback = Mock()
        try:
            for system, backend in (("win32", "_register_windows"),
                                    ("darwin", "_register_macos"), ("linux", "_register_x11")):
                with patch("stockwidget.platform.hotkeys.sys.platform", system), patch(
                    "stockwidget.platform.hotkeys.session_type", return_value="x11"
                ), patch.object(manager, backend, return_value=HotkeyResult(True)) as register:
                    for combination in combinations:
                        with self.subTest(system=system, combination=combination):
                            self.assertTrue(manager.register(combination, callback))
                            register.assert_called_with(combination, callback)
                    # 放行只代表允许尝试注册，实际占用仍保留原生接口的失败结果。
                    conflict = HotkeyResult(False, "conflict")
                    register.return_value = conflict
                    self.assertIs(manager.register("Ctrl+Shift+F", callback), conflict)
        finally:
            delete(manager)

    def test_secure_sequence_is_exact_and_windows_only(self):
        for system in ("win32", "linux", "darwin"):
            with patch("stockwidget.platform.hotkeys.sys.platform", system):
                for combination in ("Ctrl+Alt+Del", "Alt+Control+Delete", " ctrl + alt + DEL "):
                    with self.subTest(system=system, combination=combination):
                        self.assertEqual(is_reserved(combination), system == "win32")
                self.assertFalse(is_reserved("Ctrl+Alt+Shift+Delete"))
                self.assertFalse(is_reserved("Ctrl+Delete"))
                self.assertFalse(is_reserved(""))

    def test_windows_secure_sequence_does_not_reach_native_registration(self):
        manager = GlobalHotkeyManager()
        try:
            with patch("stockwidget.platform.hotkeys.sys.platform", "win32"), patch.object(
                manager, "_register_windows"
            ) as register:
                self.assertEqual(manager.register("Ctrl+Alt+Delete", Mock()).reason, "reserved")
                register.assert_not_called()
        finally:
            delete(manager)
