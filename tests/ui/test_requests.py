"""后台行情请求代数、错误提示和失效结果的界面集成。"""

from http.client import RemoteDisconnected
from types import SimpleNamespace
from unittest.mock import Mock, patch
import unittest

from shiboken6 import delete
import requests

from stockwidget.ui.floating.presenter import QuotePresenter
from stockwidget.ui.floating.taskbar import taskbar_message, render_taskbar
from stockwidget.ui.floating.widget import FloatLabel
from stockwidget.ui.settings.dialog import SettingsDialog

from tests.support import QtTestCase


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
                    with patch("stockwidget.data.quotes.request_quote", side_effect=error):
                        QuotePresenter._fetch_worker(owner, CODES, source, 7)
                    signal.emit.assert_called_once_with((False, None, f"{label}：{text}", 7))

    def test_successful_quote_payload_is_unchanged(self):
        signal = Mock()
        data = {"sh600000": {"current_price": 9.48}}
        with patch("stockwidget.data.quotes.request_quote", return_value=data) as fetch:
            QuotePresenter._fetch_worker(SimpleNamespace(data_ready=signal), CODES, "eastmoney", 3)
        fetch.assert_called_once_with(CODES, source="eastmoney")
        signal.emit.assert_called_once_with((True, data, None, 3))


class RequestMessageIntegrationTests(QtTestCase):

    def test_quote_errors_show_on_both_surfaces_and_ignore_stale_results(self):
        with patch.object(QuotePresenter, "refresh"), patch(
            "stockwidget.ui.floating.widget.GlobalHotkeyManager"
        ), patch("stockwidget.ui.floating.widget.apply_click_through"):
            window = FloatLabel({"watchlist": CODES, "taskbar_enabled": True}, CODES)
        try:
            window.quotes.project_rows([{"现价": "9.48"}], [{"现价": "text"}])
            with patch("stockwidget.data.quotes.request_quote", side_effect=requests.ConnectionError(
                RemoteDisconnected("closed")
            )):
                window.quotes._fetch_worker(CODES, "eastmoney", window.quotes._quote_generation)
            self.assertEqual(window.message_label.text(), "东财：远端关闭连接")
            self.assertEqual(taskbar_message(window), "东财：远端关闭连接")
            for dpi in (96, 144, 192):
                image = render_taskbar(window, round(48 * dpi / 96), dpi=dpi)
                self.assertFalse(image.isNull())
                self.assertLessEqual(image.width(), 480)
            window.quotes.accept_result((False, None, "新浪：连接超时", window.quotes._quote_generation - 1))
            self.assertEqual(window.message_label.text(), "东财：远端关闭连接")
        finally:
            delete(window)
            self.app.processEvents()

    def test_manual_update_failure_displays_each_source_reason(self):
        owner = SimpleNamespace(ui=SimpleNamespace(btn_check_update=Mock()), app=None, _setup_about=Mock())
        reasons = "GitHub：连接超时\nGitee：HTTP 403：访问被拒绝"
        with patch("stockwidget.ui.settings.dialog.QMessageBox.warning") as warning:
            SettingsDialog._on_update_check_finished(owner, (False, None, reasons))
        warning.assert_called_once_with(owner, "检查更新", "检查更新失败。\n" + reasons)
        owner.ui.btn_check_update.setEnabled.assert_called_once_with(True)
