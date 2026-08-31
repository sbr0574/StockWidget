import threading
import time
import unittest
from unittest.mock import patch

from stockwidget.ui.settings_dialog import SettingsDialog, _build_settings_stylesheet


class _SignalRecorder:
    def __init__(self):
        self.event = threading.Event()
        self.value = None

    def emit(self, value):
        self.value = value
        self.event.set()


class SettingsDialogTests(unittest.TestCase):
    def test_macos_button_styles_are_platform_scoped(self):
        regular = _build_settings_stylesheet(dark=False)
        macos = _build_settings_stylesheet(dark=False, macos=True)

        self.assertNotIn("QPushButton#btn_add", regular)
        self.assertIn("QPushButton#btn_add", macos)
        self.assertIn("QPushButton#btn_icon_default", macos)
        self.assertIn("rgb(10, 132, 255)", macos)

    def test_source_toggle_only_updates_for_selected_button(self):
        calls = []
        owner = type(
            "Owner",
            (),
            {"win": type("Window", (), {"set_data_source": calls.append})()},
        )()

        SettingsDialog._on_source_toggled(owner, "sina", False)
        SettingsDialog._on_source_toggled(owner, "eastmoney", True)

        self.assertEqual(calls, ["eastmoney"])

    def test_github_check_does_not_block_dialog_thread(self):
        owner = type("Owner", (), {"github_check_finished": _SignalRecorder()})()

        def slow_check(timeout):
            time.sleep(0.2)
            return False

        with patch(
            "stockwidget.ui.settings_dialog.github_available",
            side_effect=slow_check,
        ):
            started = time.perf_counter()
            SettingsDialog._start_github_check(owner)
            elapsed = time.perf_counter() - started
            self.assertLess(elapsed, 0.1)
            self.assertTrue(owner.github_check_finished.event.wait(1))

        self.assertTrue(owner.github_check_finished.value)


if __name__ == "__main__":
    unittest.main()
