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
from stockwidget.data.quotes import _new_entry

from tests.support import QtTestCase, SettingsTestCase


CODES = {"sh600000": {"market": "sh", "code": "600000", "checked": True}}


class QuoteFreshnessTests(SettingsTestCase):
    def setUp(self):
        super().setUp()
        self.codes = {**CODES, "sh600001": {"market": "sh", "code": "600001", "checked": True}}
        _, self.window = self._make_dialog(codes=self.codes, watchlist=self.codes,
                                           visible_metrics=["price"], taskbar_sync_metrics=True)

    @staticmethod
    def quote(price=100, **times):
        quote = _new_entry("测试", 100, 100, price, max(100, price), min(100, price), 1, 100,
                           [0] * 5, [0] * 5, [0] * 5, [0] * 5, "2026-09-30", "10:00:00")
        return {**quote, **times}

    def accept(self, data):
        self.window.quotes.accept_result((True, data, None, self.window.quotes._quote_generation))

    def test_older_quotes_are_filtered_per_instrument_before_projection_and_notification(self):
        for first, older, newer in (
            ({"time": "10:00:00"}, {"time": "09:59:59"}, {"time": "10:00:01"}),
            ({"timestamp": 1790733600}, {"timestamp": 1790733599}, {"timestamp": 1790733601}),
            ({"date": "2026-10-01"}, {"date": "2026-09-30", "time": "15:00:00"}, {"date": "2026-10-02"}),
        ):
            with self.subTest(first=first):
                self.window.quotes.invalidate()
                self.accept({code: self.quote(**first) for code in self.codes})
                on_update = Mock()
                self.window.quotes.quotes_updated.connect(on_update)
                with patch.object(self.window.hide_controller, "quotes_refreshed") as hide:
                    self.accept({"sh600000": self.quote(50, **older), "sh600001": self.quote(101, **newer)})
                on_update.assert_called_once()
                self.window.quotes.quotes_updated.disconnect(on_update)
                self.assertEqual(self.window.quotes.quote_for(self.codes["sh600000"])["current_price"], 100)
                self.assertEqual(self.window.quotes.quote_for(self.codes["sh600001"])["current_price"], 101)
                for model in (self.window.model, self.window.taskbar_model):
                    column = model._headers.index("现价")
                    self.assertEqual([model.index(i, column).data() for i in range(2)], ["100.00 ", "101.00↑"])
                self.assertEqual(hide.call_args.args[0]["sh600000"]["current_price"], 100)

    def test_same_timestamp_and_later_quotes_update_but_unknown_cannot_replace_known_time(self):
        self.accept({"sh600000": self.quote()})
        self.accept({"sh600000": self.quote(102)})
        self.assertEqual(self.window.quotes.quote_for(self.codes["sh600000"])["current_price"], 102)
        for invalid in ({"time": ""}, {"timestamp": float("nan")}, {"date": "invalid"}):
            self.accept({"sh600000": self.quote(50, **invalid)})
            self.assertEqual(self.window.quotes.quote_for(self.codes["sh600000"])["current_price"], 102)
        self.accept({"sh600000": self.quote(103, time="10:01:00")})
        self.assertEqual(self.window.quotes.quote_for(self.codes["sh600000"])["current_price"], 103)

    def test_missing_quotes_and_errors_do_not_erase_last_accepted_timestamp(self):
        self.accept({code: self.quote() for code in self.codes})
        for response in ({}, {"sh600001": self.quote(101)}):
            self.accept(response)
            self.window.quotes.accept_result((False, None, "请求失败"))
            self.accept({"sh600000": self.quote(50, time="09:59:00")})
            self.assertEqual(self.window.quotes.quote_for(self.codes["sh600000"])["current_price"], 100)

    def test_invalidation_accepts_fresh_state_and_rejects_late_previous_result(self):
        self.accept({"sh600000": self.quote()})
        generation = self.window.quotes._quote_generation
        self.window.quotes.invalidate()
        self.accept({"sh600000": self.quote(95, time="09:55:00")})
        self.window.quotes.accept_result((True, {"sh600000": self.quote(200, time="11:00:00")}, None, generation))
        self.assertEqual(self.window.quotes.quote_for(self.codes["sh600000"])["current_price"], 95)

    def test_first_quote_with_unknown_time_remains_usable(self):
        self.accept({"sh600000": self.quote(99, time="")})
        self.assertEqual(self.window.quotes.quote_for(self.codes["sh600000"])["current_price"], 99)
        self.accept({"sh600000": self.quote(100)})
        self.assertEqual(self.window.quotes.quote_for(self.codes["sh600000"])["current_price"], 100)


class QuoteWorkerTests(unittest.TestCase):
    def test_errors_identify_source_and_keep_generation(self):
        for error, text in ((requests.ConnectTimeout(), "连接超时"),
                                (requests.ConnectionError(RemoteDisconnected()), "远端关闭连接"),
                                (requests.exceptions.JSONDecodeError("bad", "", 0), "响应不是有效 JSON"),
                                (ValueError("sensitive response"), "行情处理失败（ValueError）")):
            with self.subTest(error=type(error).__name__):
                signal = Mock()
                owner = SimpleNamespace(data_ready=signal)
                with patch("stockwidget.data.quotes.request_quote", side_effect=error):
                    QuotePresenter._fetch_worker(owner, CODES, 7)
                signal.emit.assert_called_once_with((False, None, f"新浪：{text}", 7))

    def test_successful_quote_payload_is_unchanged(self):
        signal = Mock()
        data = {"sh600000": {"current_price": 9.48}}
        with patch("stockwidget.data.quotes.request_quote", return_value=data) as fetch:
            QuotePresenter._fetch_worker(SimpleNamespace(data_ready=signal), CODES, 3)
        fetch.assert_called_once_with(CODES)
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
                window.quotes._fetch_worker(CODES, window.quotes._quote_generation)
            self.assertEqual(window.message_label.text(), "新浪：远端关闭连接")
            self.assertEqual(taskbar_message(window), "新浪：远端关闭连接")
            for dpi in (96, 144, 192):
                image = render_taskbar(window, round(48 * dpi / 96), dpi=dpi)
                self.assertFalse(image.isNull())
                self.assertLessEqual(image.width(), 480)
            window.quotes.accept_result((False, None, "新浪：连接超时", window.quotes._quote_generation - 1))
            self.assertEqual(window.message_label.text(), "新浪：远端关闭连接")
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
