"""行情后台错误必须保留来源与请求代数，并同步到浮窗和任务栏。"""

from http.client import RemoteDisconnected
import os
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import requests
from PySide6.QtWidgets import QApplication
from shiboken6 import delete

from stockwidget.ui.settings_dialog import SettingsDialog
from stockwidget.ui.taskbar import taskbar_message, render_taskbar
from stockwidget.ui.widget import FloatLabel


CODES = {"sh600000": {"market": "sh", "code": "600000", "checked": True}}


class QuoteWorkerTests(unittest.TestCase):
    def test_errors_identify_source_and_keep_generation(self):
        for source, label in (("sina", "新浪"), ("eastmoney", "东财")):
            for error, text in ((requests.ConnectTimeout(), "连接超时"),
                                (requests.ConnectionError(RemoteDisconnected()), "远端关闭连接"),
                                (requests.exceptions.JSONDecodeError("bad", "", 0), "响应不是有效 JSON"),
                                (ValueError("sensitive response"), "行情处理失败（ValueError）")):
                with self.subTest(source=source, error=type(error).__name__):
                    signal = Mock()
                    owner = SimpleNamespace(data_ready=signal)
                    with patch("stockwidget.ui.widget.request_quote", side_effect=error):
                        FloatLabel._fetch_data_worker(owner, CODES, source, 7)
                    signal.emit.assert_called_once_with((False, None, f"{label}：{text}", 7))

    def test_successful_quote_payload_is_unchanged(self):
        signal = Mock()
        data = {"sh600000": {"current_price": 9.48}}
        with patch("stockwidget.ui.widget.request_quote", return_value=data) as fetch:
            FloatLabel._fetch_data_worker(SimpleNamespace(data_ready=signal), CODES, "eastmoney", 3)
        fetch.assert_called_once_with(CODES, source="eastmoney")
        signal.emit.assert_called_once_with((True, data, None, 3))


class RequestMessageIntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_quote_errors_show_on_both_surfaces_and_ignore_stale_results(self):
        with patch.object(FloatLabel, "_refresh_from_function"), patch(
            "stockwidget.ui.widget.GlobalHotkeyManager"
        ), patch("stockwidget.ui.widget.apply_click_through"):
            window = FloatLabel({"watchlist": CODES, "taskbar_enabled": True}, CODES)
        try:
            window._project_columns([{"现价": "9.48"}], [{"现价": "text"}])
            with patch("stockwidget.ui.widget.request_quote", side_effect=requests.ConnectionError(
                RemoteDisconnected("closed")
            )):
                window._fetch_data_worker(CODES, "eastmoney", window._quote_generation)
            self.assertEqual(window.message_label.text(), "东财：远端关闭连接")
            self.assertEqual(taskbar_message(window), "东财：远端关闭连接")
            for dpi in (96, 144, 192):
                image = render_taskbar(window, round(48 * dpi / 96), dpi=dpi)
                self.assertFalse(image.isNull())
                self.assertLessEqual(image.width(), 480)
            window._process_data((False, None, "新浪：连接超时", window._quote_generation - 1))
            self.assertEqual(window.message_label.text(), "东财：远端关闭连接")
        finally:
            delete(window)
            self.app.processEvents()

    def test_manual_update_failure_displays_each_source_reason(self):
        owner = SimpleNamespace(ui=SimpleNamespace(btn_check_update=Mock()), app=None, _setup_about=Mock())
        reasons = "GitHub：连接超时\nGitee：HTTP 403：访问被拒绝"
        with patch("stockwidget.ui.settings_dialog.QMessageBox.warning") as warning:
            SettingsDialog._on_update_check_finished(owner, (False, None, reasons))
        warning.assert_called_once_with(owner, "检查更新", "检查更新失败。\n" + reasons)
        owner.ui.btn_check_update.setEnabled.assert_called_once_with(True)
