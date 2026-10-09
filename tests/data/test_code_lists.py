"""代码列表加载、缓存、远端同步和服务端列表生成。"""

from datetime import datetime, timedelta, timezone, date
from io import BytesIO
from pathlib import Path
from unittest.mock import MagicMock, Mock, call, patch
from urllib.parse import urlsplit
import base64
import json
import os
import tempfile
import unittest

from openpyxl import Workbook
from scripts import update_codes
import pandas as pd
import requests

from stockwidget.constants import CODES_SOURCE_URLS, CODES_STATUS_FILE, CODE_LIST_FILES
from stockwidget.core.config_store import config_paths, data_cache_dir, load_file, save_file
from stockwidget.data.code_lists import (
    CODES_HTTP_TIMEOUT,
    CodeDownloadError,
    CODES_RETRY_SECONDS,
    CodeListManager,
    fetch_json_from_url,
    next_code_check_delay,
)


UTC8 = timezone(timedelta(hours=8))


def _payload(filename: str, update_date: str) -> dict:
    code = filename.removesuffix(".json")
    return {
        "last_update": update_date,
        "codes": {
            code: {
                "code": code,
                "market": "",
                "type": "",
                "name": filename,
                "name_en": "",
                "py": "",
                "abbr": "",
            }
        },
    }


def _status(run_date: str, started_at: str) -> dict:
    return {
        "run_date": run_date,
        "started_at": started_at,
        "completed_at": started_at,
        "files": {},
    }


def _load_manager(resources: dict, local_status: dict, fetcher) -> CodeListManager:
    def load_local_file(filename, *, directory=None):
        return local_status if filename == CODES_STATUS_FILE else {}

    with (
        patch(
            "stockwidget.data.code_lists.load_json_from_resource",
            side_effect=lambda path: resources.get(path[2:], {}),
        ),
        patch(
            "stockwidget.data.code_lists.load_file",
            side_effect=load_local_file,
        ),
    ):
        manager = CodeListManager(fetcher=fetcher)
        manager.load_local()
    return manager


class NextCodeCheckDelayTests(unittest.TestCase):
    def test_checks_at_nine_and_skips_weekends(self):
        cases = (
            (datetime(2026, 8, 31, 8, 59, tzinfo=UTC8), 60),
            (datetime(2026, 8, 31, 9, 0, tzinfo=UTC8), 24 * 60 * 60),
            (datetime(2026, 8, 28, 9, 0, tzinfo=UTC8), 3 * 24 * 60 * 60),
        )
        for now, delay in cases:
            with self.subTest(now=now):
                self.assertEqual(next_code_check_delay(now), delay)


class CodeListManagerTests(unittest.TestCase):
    def setUp(self):
        self.local_date = "2026-08-28"
        self.local_status = _status(
            self.local_date, "2026-08-28T09:01:00+08:00"
        )
        self.resources = {
            name: _payload(name, self.local_date) for name in CODE_LIST_FILES
        }

    def test_startup_is_cached_and_date_comes_from_local_status(self):
        manager = _load_manager(self.resources, self.local_status, Mock())

        self.assertEqual(manager.state(), ("cached", self.local_date))
        self.assertEqual(len(manager.codes()), len(CODE_LIST_FILES))

    def test_same_remote_run_time_marks_current_without_downloading_lists(self):
        fetched = []

        def fetcher(url):
            fetched.append(url)
            return self.local_status

        manager = _load_manager(self.resources, self.local_status, fetcher)
        with patch("stockwidget.data.code_lists.save_file") as save:
            result = manager.sync_remote(
                now=datetime(2026, 8, 31, 10, 0, tzinfo=UTC8)
            )

        self.assertEqual(result["state"], "current")
        self.assertEqual(result["date"], self.local_date)
        self.assertEqual(result["updated"], ())
        self.assertEqual(len(fetched), 1)
        save.assert_not_called()

    def test_newer_remote_run_downloads_all_nine_and_saves_status_last(self):
        remote_date = "2026-08-31"
        remote_status = _status(
            remote_date, "2026-08-31T09:01:00+08:00"
        )
        payloads = {
            name: _payload(name, remote_date) for name in CODE_LIST_FILES
        }

        def fetcher(url):
            filename = urlsplit(url).path.rsplit("/", 1)[-1]
            return remote_status if filename == CODES_STATUS_FILE else payloads[filename]

        manager = _load_manager(self.resources, self.local_status, fetcher)
        with patch("stockwidget.data.code_lists.save_file") as save:
            result = manager.sync_remote(
                now=datetime(2026, 8, 31, 10, 0, tzinfo=UTC8)
            )

        self.assertEqual(result["state"], "current")
        self.assertEqual(result["date"], remote_date)
        self.assertEqual(result["updated"], CODE_LIST_FILES)
        self.assertEqual(manager.state(), ("current", remote_date))
        self.assertEqual(save.call_count, len(CODE_LIST_FILES) + 1)
        self.assertEqual(save.call_args_list[-1], call(remote_status, CODES_STATUS_FILE, directory=data_cache_dir()))
        for filename in CODE_LIST_FILES:
            self.assertIn(call(payloads[filename], filename, directory=data_cache_dir()), save.call_args_list)

    def test_started_at_distinguishes_two_runs_on_same_date(self):
        remote_status = _status(
            self.local_date, "2026-08-28T10:01:00+08:00"
        )
        payloads = {
            name: _payload(name, self.local_date) for name in CODE_LIST_FILES
        }

        def fetcher(url):
            filename = urlsplit(url).path.rsplit("/", 1)[-1]
            return remote_status if filename == CODES_STATUS_FILE else payloads[filename]

        manager = _load_manager(self.resources, self.local_status, fetcher)
        with patch("stockwidget.data.code_lists.save_file") as save:
            result = manager.sync_remote(
                now=datetime(2026, 8, 28, 11, 0, tzinfo=UTC8)
            )

        self.assertEqual(result["updated"], CODE_LIST_FILES)
        self.assertEqual(save.call_count, len(CODE_LIST_FILES) + 1)

    def test_status_failure_checks_on_weekend_and_retries_in_half_hour(self):
        fetcher = Mock(return_value=None)
        manager = _load_manager(self.resources, self.local_status, fetcher)

        result = manager.sync_remote(
            now=datetime(2026, 8, 29, 12, 0, tzinfo=UTC8)
        )

        self.assertEqual(result["state"], "cached")
        self.assertEqual(result["date"], self.local_date)
        self.assertEqual(result["retry_seconds"], CODES_RETRY_SECONDS)
        self.assertEqual(fetcher.call_count, len(CODES_SOURCE_URLS))
        self.assertIn("raw.githubusercontent.com", fetcher.call_args_list[0].args[0])
        self.assertIn("gitee", fetcher.call_args_list[1].args[0])

    def test_one_list_failure_saves_nothing_and_keeps_cached(self):
        remote_status = _status(
            "2026-08-31", "2026-08-31T09:01:00+08:00"
        )

        def fetcher(url):
            filename = urlsplit(url).path.rsplit("/", 1)[-1]
            if filename == CODES_STATUS_FILE:
                return remote_status
            if filename == "stock_hk.json":
                return None
            return _payload(filename, "2026-08-31")

        manager = _load_manager(self.resources, self.local_status, fetcher)
        with patch("stockwidget.data.code_lists.save_file") as save:
            result = manager.sync_remote(
                now=datetime(2026, 8, 31, 10, 0, tzinfo=UTC8)
            )

        self.assertEqual(result["state"], "cached")
        self.assertEqual(result["date"], self.local_date)
        self.assertEqual(result["retry_seconds"], CODES_RETRY_SECONDS)
        save.assert_not_called()

    def test_incomplete_local_lists_force_full_download(self):
        resources = dict(self.resources)
        resources.pop(CODE_LIST_FILES[-1])
        payloads = {
            name: _payload(name, self.local_date) for name in CODE_LIST_FILES
        }

        def fetcher(url):
            filename = urlsplit(url).path.rsplit("/", 1)[-1]
            return self.local_status if filename == CODES_STATUS_FILE else payloads[filename]

        manager = _load_manager(resources, self.local_status, fetcher)
        with patch("stockwidget.data.code_lists.save_file") as save:
            result = manager.sync_remote(
                now=datetime(2026, 8, 31, 10, 0, tzinfo=UTC8)
            )

        self.assertEqual(result["updated"], CODE_LIST_FILES)
        self.assertEqual(save.call_count, len(CODE_LIST_FILES) + 1)

    def test_remote_file_falls_back_from_github_to_gitee(self):
        fetched = []

        def fetcher(url):
            fetched.append(url)
            return {"ok": True} if "gitee" in url else None

        manager = CodeListManager(fetcher=fetcher)
        self.assertEqual(manager._fetch_remote("test.json"), {"ok": True})
        self.assertIn("raw.githubusercontent.com", fetched[0])
        self.assertIn("gitee", fetched[1])

    def test_failed_github_is_not_retried_for_every_file(self):
        fetched = []
        remote_status = _status("2026-09-09", "2026-09-09T09:00:00+08:00")

        def fetcher(url):
            fetched.append(url)
            if "githubusercontent" in url:
                raise requests.ConnectTimeout("GitHub timeout")
            filename = urlsplit(url).path.rsplit("/", 1)[-1]
            return remote_status if filename == CODES_STATUS_FILE else _payload(filename, "2026-09-09")

        manager = _load_manager(self.resources, self.local_status, fetcher)
        with patch("stockwidget.data.code_lists.save_file"):
            result = manager.sync_remote()

        self.assertEqual(result["state"], "current")
        self.assertEqual(sum("githubusercontent" in url for url in fetched), 1)
        self.assertEqual(len(fetched), len(CODE_LIST_FILES) + 2)
        self.assertEqual(manager.last_error(), "")

    def test_retry_keeps_downloads_from_same_run(self):
        fetched = []
        failed = True
        remote_status = _status("2026-09-09", "2026-09-09T09:00:00+08:00")

        def fetcher(url):
            filename = urlsplit(url).path.rsplit("/", 1)[-1]
            fetched.append(filename)
            if filename == CODES_STATUS_FILE:
                return remote_status
            if filename == "stock_hk.json" and failed:
                return None
            return _payload(filename, "2026-09-09")

        manager = _load_manager(self.resources, self.local_status, fetcher)
        with patch("stockwidget.data.code_lists.save_file") as save:
            self.assertEqual(manager.sync_remote()["state"], "cached")
            self.assertIn("stock_hk.json", manager.last_error())
            save.assert_not_called()
            failed = False
            self.assertEqual(manager.sync_remote()["state"], "current")
        self.assertEqual(fetched.count("stock_sh.json"), 1)
        self.assertEqual(manager.last_error(), "")

    def test_new_remote_run_discards_partial_downloads(self):
        run_date = "2026-09-08"

        def fetcher(url):
            filename = urlsplit(url).path.rsplit("/", 1)[-1]
            if filename == CODES_STATUS_FILE:
                return _status(run_date, f"{run_date}T09:00:00+08:00")
            if filename == "stock_hk.json" and run_date == "2026-09-08":
                return None
            return _payload(filename, run_date)

        manager = _load_manager(self.resources, self.local_status, fetcher)
        with patch("stockwidget.data.code_lists.save_file") as save:
            manager.sync_remote()
            run_date = "2026-09-09"
            manager.sync_remote()
        self.assertEqual(save.call_args_list[0].args[0]["last_update"], "2026-09-09")

    def test_preferred_mirror_can_fall_back_when_it_fails(self):
        manager = CodeListManager(fetcher=Mock(side_effect=[None, {"ok": True}]))
        manager._fetch_remote("test.json")
        fetcher = Mock(side_effect=[None, None, {"ok": True}])
        manager._fetcher = fetcher
        self.assertEqual(manager._fetch_remote("other.json"), {"ok": True})
        self.assertIn("gitee", fetcher.call_args_list[0].args[0])
        self.assertIn("/api/v5/", fetcher.call_args_list[1].args[0])
        self.assertIn("githubusercontent", fetcher.call_args_list[2].args[0])


class FetchJsonTests(unittest.TestCase):
    @patch("stockwidget.data.code_lists.requests.get")
    def test_gitee_api_content_is_decoded(self, get):
        payload = _payload("stock_hk.json", "2026-09-09")
        encoded = base64.b64encode(json.dumps(payload).encode()).decode()
        response = Mock()
        response.iter_content.return_value = [json.dumps({"encoding": "base64", "content": encoded}).encode()]
        get.return_value = MagicMock()
        get.return_value.__enter__.return_value = response
        result = fetch_json_from_url(CODES_SOURCE_URLS[-1].format(name="stock_hk.json"))
        self.assertEqual(result, payload)

    @patch("stockwidget.data.code_lists.requests.get")
    def test_request_failures_preserve_specific_reasons(self, get):
        response = requests.Response()
        response.status_code = 429
        for error, expected in ((requests.ConnectTimeout(), "连接超时"),
                                (requests.ReadTimeout(), "读取超时"),
                                (requests.exceptions.ProxyError(), "代理请求失败"),
                                (requests.exceptions.SSLError(), "安全连接失败（SSL）"),
                                (requests.HTTPError(response=response), "HTTP 429：请求过于频繁"),
                                (requests.HTTPError(), "HTTP 请求失败")):
            with self.subTest(error=type(error).__name__):
                get.side_effect = error
                with self.assertRaises(CodeDownloadError) as caught:
                    fetch_json_from_url("https://example.com/test.json")
                self.assertEqual(str(caught.exception), expected)
                self.assertIs(caught.exception.__cause__, error)

    @patch("stockwidget.data.code_lists.requests.get")
    def test_html_response_is_rejected_and_connection_closed(self, get):
        response = Mock()
        response.iter_content.return_value = [b"<html>Access denied</html>"]
        get.return_value = MagicMock()
        get.return_value.__enter__.return_value = response
        with self.assertRaisesRegex(CodeDownloadError, "JSON"):
            fetch_json_from_url("https://example.com/test.json")
        get.return_value.__exit__.assert_called_once()

    @patch("stockwidget.data.code_lists.monotonic", side_effect=[0, 1, 61])
    @patch("stockwidget.data.code_lists.requests.get")
    def test_slow_continuous_download_is_stopped(self, get, clock):
        response = Mock()
        response.iter_content.return_value = [b'{"ok":', b'true}']
        get.return_value = MagicMock()
        get.return_value.__enter__.return_value = response
        with self.assertRaisesRegex(CodeDownloadError, "耗时过长"):
            fetch_json_from_url("https://example.com/test.json")
        get.return_value.__exit__.assert_called_once()

    @patch("stockwidget.data.code_lists.requests.get")
    def test_separate_connect_and_read_timeouts(self, get):
        response = Mock()
        response.iter_content.return_value = [b'{"ok": true}']
        get.return_value = MagicMock()
        get.return_value.__enter__.return_value = response

        result = fetch_json_from_url("https://example.com/test.json")

        self.assertEqual(result, {"ok": True})
        get.assert_called_once_with(
            "https://example.com/test.json",
            timeout=CODES_HTTP_TIMEOUT,
            stream=True,
            headers={"Cache-Control": "no-cache", "User-Agent": "StockWidget"},
        )


class CodeCacheTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        env = patch.dict(os.environ, {"APPDATA": self.temp.name})
        env.start()
        self.addCleanup(env.stop)
        # 缓存路径和迁移测试不应受内置代码表更新影响。
        resources = patch(
            "stockwidget.data.code_lists.load_json_from_resource", return_value={}
        )
        resources.start()
        self.addCleanup(resources.stop)

    def test_legacy_cache_is_migrated_without_moving_configuration(self):
        payload = _payload("stock_sh.json", "2026-09-09")
        status = _status("2026-09-09", "2026-09-09T09:00:00+08:00")
        save_file(payload, "stock_sh.json")
        save_file(status, CODES_STATUS_FILE)
        save_file({"watchlist": {}}, "stock_widget_config.json")

        manager = CodeListManager()
        manager.load_local()

        self.assertEqual(load_file("stock_sh.json", directory=data_cache_dir()), payload)
        self.assertEqual(load_file(CODES_STATUS_FILE, directory=data_cache_dir()), status)
        self.assertFalse(os.path.exists(os.path.join(config_paths(), "stock_sh.json")))
        self.assertEqual(load_file("stock_widget_config.json"), {"watchlist": {}})
        self.assertEqual(manager.state(), ("cached", "2026-09-09"))

    def test_failed_migration_keeps_legacy_cache_usable(self):
        payload = _payload("stock_sh.json", "2026-09-09")
        save_file(payload, "stock_sh.json")
        with patch("stockwidget.data.code_lists.save_file", side_effect=OSError("read only")):
            manager = CodeListManager()
            manager.load_local()
        self.assertIn("stock_sh", manager.codes())
        self.assertEqual(load_file("stock_sh.json"), payload)

    def test_new_cache_takes_precedence_over_legacy(self):
        save_file(_payload("stock_sh.json", "2026-09-08"), "stock_sh.json")
        payload = _payload("stock_sh.json", "2026-09-09")
        save_file(payload, "stock_sh.json", directory=data_cache_dir())
        manager = CodeListManager()
        manager.load_local()
        self.assertEqual(manager._payloads["stock_sh.json"], payload)


def _frame(code: str, name: str) -> pd.DataFrame:
    return pd.DataFrame(
        [(code, name, "", "美", "us")],
        columns=update_codes._DF_COLUMNS,
    )


def _xlsx_bytes(headers: tuple[str, ...], rows: list[tuple]) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "ListOfSecurities"
    sheet.append(("title",))
    sheet.append(("updated",))
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def _xlsx_response(content: bytes) -> Mock:
    response = Mock()
    response.content = content
    response.headers = {
        "Content-Type": (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    }
    return response


def _text_response(text: str) -> Mock:
    response = Mock()
    response.content = text.encode("utf-8")
    return response


def _us_directory_texts() -> tuple[str, str]:
    nasdaq = "\n".join(
        (
            "Symbol|Security Name|Market Category|Test Issue|Financial Status|Round Lot Size|ETF|NextShares",
            "AAPL|Apple Inc. - Common Stock|Q|N|N|100|N|N",
            "QQQ|Invesco QQQ Trust, Series 1|Q|N|N|100|Y|N",
            "ZXIET|Test Security|Q|Y|N|100|N|N",
            "File Creation Time: 0831202621:31|||||||",
        )
    )
    other = "\n".join(
        (
            "ACT Symbol|Security Name|Exchange|CQS Symbol|ETF|Round Lot Size|Test Issue|NASDAQ Symbol",
            "BRK.B|Berkshire Hathaway Class B|N|BRK.B|N|40|N|BRK.B",
            "AAC.W|Ares Acquisition Warrant|N|AAC.WS|N|100|N|AAC+",
            "ABR$D|Arbor Realty Preferred D|N|ABRpD|N|100|N|ABR-D",
            "SPY|SPDR S&P 500 ETF Trust|P|SPY|Y|100|N|SPY",
            "TEST|Exchange Test Security|N|TEST|N|100|Y|TEST",
            "File Creation Time: 0831202621:31||||||",
        )
    )
    return nasdaq, other


class UpdateCodeFilesTests(unittest.TestCase):
    def test_shanghai_futures_market_preserves_contract_keys(self):
        response = Mock()
        response.json.return_value = [
            {"symbol": "AU0", "name": "黄金连续"},
            {"symbol": "SC2610", "name": "原油2610"},
        ]
        with (
            patch.object(update_codes, "_futures_nodes", return_value=["shfe_au"]),
            patch.object(update_codes.requests, "get", return_value=response),
        ):
            codes = update_codes._df_to_dict(update_codes.futures_info_all())
        self.assertEqual(set(codes), {"au0", "sc2610"})
        for entry in codes.values():
            self.assertEqual(entry["market"], "sh")
            self.assertEqual(entry["type"], "期")

    def test_long_running_tasks_are_scheduled_first(self):
        filenames = [task[0] for task in update_codes._tasks()]
        self.assertEqual(
            filenames[:3],
            ["stock_us.json", "stock_hk.json", "futures_sh.json"],
        )

    def test_szse_download_relies_on_task_level_retry(self):
        with (
            patch.object(
                update_codes.requests,
                "get",
                side_effect=update_codes.requests.ConnectionError("temporary"),
            ) as request,
            patch.object(update_codes.time, "sleep") as sleep,
        ):
            with self.assertRaises(update_codes.requests.ConnectionError):
                update_codes._szse_xlsx("1110", "tab1", "https://www.szse.cn/")

        request.assert_called_once()
        sleep.assert_not_called()

    def test_hkex_list_includes_equities_funds_and_reits(self):
        english = _xlsx_bytes(
            ("Stock Code", "Name of Securities", "Category"),
            [
                ("00001", "CKH HOLDINGS", "Equity"),
                ("02800", "TRACKER FUND", "Exchange Traded Products"),
                ("00405", "YUEXIU REIT", "Real Estate Investment Trusts"),
                ("04000", "TEST BOND", "Debt Securities"),
            ],
        )
        chinese = _xlsx_bytes(
            ("股份代號", "股份名稱"),
            [
                ("00001", "長和"),
                ("02800", "盈富基金"),
                ("00405", "越秀房產信託基金"),
                ("04000", "測試債券"),
            ],
        )

        with patch.object(
            update_codes.requests,
            "get",
            side_effect=[_xlsx_response(english), _xlsx_response(chinese)],
        ) as request:
            frame = update_codes._stock_hk_name_code()

        self.assertEqual(frame["code"].tolist(), ["00001", "02800", "00405"])
        by_code = frame.set_index("code")
        self.assertEqual(by_code.at["00001", "name"], "長和")
        self.assertEqual(by_code.at["00001", "name_en"], "CKH HOLDINGS")
        self.assertEqual(by_code.at["00001", "type"], "港")
        self.assertEqual(by_code.at["02800", "type"], "基")
        self.assertEqual(by_code.at["00405", "type"], "基")
        self.assertNotIn("04000", by_code.index)
        self.assertEqual(
            [call.args[0] for call in request.call_args_list],
            [
                update_codes._HKEX_SECURITIES_EN_URL,
                update_codes._HKEX_SECURITIES_ZH_URL,
            ],
        )

    def test_nasdaq_official_lists_keep_stocks_and_etfs_as_us(self):
        nasdaq, other = _us_directory_texts()
        with tempfile.TemporaryDirectory() as directory:
            cache_path = str(Path(directory) / "aliases.json")
            Path(cache_path).write_text(
                json.dumps(
                    {"last_update": "2026-08-03", "aliases": {"old": "旧证券"}},
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            with (
                patch.object(
                    update_codes.requests,
                    "get",
                    side_effect=[_text_response(nasdaq), _text_response(other)],
                ) as request,
                patch.object(update_codes, "_fetch_us_cn_aliases") as fetch_aliases,
            ):
                frame = update_codes._stock_us_name_code(
                    today=date(2026, 8, 31),
                    output_dir=directory,
                    cache_path=cache_path,
                )

        self.assertEqual(
            frame["code"].tolist(),
            ["aapl", "qqq", "brk_b", "aac_ws", "abr_d", "spy"],
        )
        self.assertEqual(set(frame["type"]), {"美"})
        self.assertEqual(set(frame["market"]), {"us"})
        self.assertNotIn("zxiet", set(frame["code"]))
        self.assertNotIn("test", set(frame["code"]))
        fetch_aliases.assert_not_called()
        self.assertEqual(
            [call.args[0] for call in request.call_args_list],
            [
                update_codes._NASDAQ_LISTED_URL,
                update_codes._NASDAQ_OTHER_LISTED_URL,
            ],
        )

    def test_us_names_use_git_cached_aliases_without_daily_eastmoney_request(self):
        nasdaq, other = _us_directory_texts()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            cache_path = root / "cache_us_cn_aliases.json"
            cache_path.write_text(
                json.dumps(
                    {
                        "last_update": "2026-08-03",
                        "aliases": {"aapl": "苹果"},
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            with (
                patch.object(
                    update_codes.requests,
                    "get",
                    side_effect=[_text_response(nasdaq), _text_response(other)],
                ),
                patch.object(update_codes, "_fetch_us_cn_aliases") as fetch_aliases,
            ):
                frame = update_codes._stock_us_name_code(
                    today=date(2026, 8, 31),
                    output_dir=directory,
                    cache_path=str(cache_path),
                )

            by_code = frame.set_index("code")
            self.assertEqual(by_code.at["aapl", "name"], "苹果")
            self.assertEqual(
                by_code.at["aapl", "name_en"], "Apple Inc. - Common Stock"
            )
            fetch_aliases.assert_not_called()
            self.assertTrue(cache_path.exists())

    def test_us_aliases_refresh_when_month_changes(self):
        nasdaq, other = _us_directory_texts()
        with tempfile.TemporaryDirectory() as directory:
            cache_path = str(Path(directory) / "aliases.json")
            Path(cache_path).write_text(
                json.dumps(
                    {"last_update": "2026-08-03", "aliases": {"aapl": "旧苹果"}},
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            with (
                patch.object(
                    update_codes.requests,
                    "get",
                    side_effect=[_text_response(nasdaq), _text_response(other)],
                ),
                patch.object(
                    update_codes,
                    "_fetch_us_cn_aliases",
                    return_value={
                        "aapl": "苹果",
                        "spy": "标普500ETF",
                        "not_listed": "已退市测试",
                    },
                ) as fetch_aliases,
            ):
                frame = update_codes._stock_us_name_code(
                    today=date(2026, 9, 1),
                    output_dir=directory,
                    cache_path=cache_path,
                )

            by_code = frame.set_index("code")
            self.assertEqual(by_code.at["aapl", "name"], "苹果")
            self.assertEqual(by_code.at["spy", "name"], "标普500ETF")
            self.assertEqual(by_code.at["qqq", "name"], "Invesco QQQ Trust, Series 1")
            self.assertEqual(set(frame["type"]), {"美"})
            fetch_aliases.assert_called_once_with()
            cache = json.loads(Path(cache_path).read_text(encoding="utf-8"))
            self.assertEqual(cache["last_update"], "2026-09-01")
            self.assertNotIn("not_listed", cache["aliases"])

    def test_us_aliases_are_not_fetched_again_in_same_month(self):
        nasdaq, other = _us_directory_texts()
        with tempfile.TemporaryDirectory() as directory:
            cache_path = Path(directory) / "aliases.json"
            cache_path.write_text(
                json.dumps(
                    {"last_update": "2026-09-01", "aliases": {"aapl": "苹果"}},
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            with (
                patch.object(
                    update_codes.requests,
                    "get",
                    side_effect=[_text_response(nasdaq), _text_response(other)],
                ),
                patch.object(update_codes, "_fetch_us_cn_aliases") as fetch_aliases,
            ):
                frame = update_codes._stock_us_name_code(
                    today=date(2026, 9, 1),
                    output_dir=directory,
                    cache_path=str(cache_path),
                )

            self.assertEqual(frame.set_index("code").at["aapl", "name"], "苹果")
            fetch_aliases.assert_not_called()

    def test_us_alias_refresh_failure_does_not_block_official_list(self):
        nasdaq, other = _us_directory_texts()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            cache_path = root / "aliases.json"
            cache_path.write_text(
                json.dumps(
                    {"last_update": "2026-08-03", "aliases": {"aapl": "苹果"}},
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            with (
                patch.object(
                    update_codes.requests,
                    "get",
                    side_effect=[_text_response(nasdaq), _text_response(other)],
                ),
                patch.object(
                    update_codes,
                    "_fetch_us_cn_aliases",
                    side_effect=RuntimeError("eastmoney unavailable"),
                ),
            ):
                frame = update_codes._stock_us_name_code(
                    today=date(2026, 9, 1),
                    output_dir=directory,
                    cache_path=str(cache_path),
                )

            self.assertEqual(frame.set_index("code").at["aapl", "name"], "苹果")
            self.assertEqual(len(frame), 6)

    def test_successful_unchanged_and_failed_files_are_recorded_independently(self):
        tasks = [
            ("same.json", "相同", None, (), {}),
            ("changed.json", "变化", None, (), {}),
            ("failed.json", "失败", None, (), {}),
        ]
        same_codes = update_codes._df_to_dict(_frame("same", "Same"))
        failed_codes = update_codes._df_to_dict(_frame("failed", "Failed"))

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "same.json").write_text(
                json.dumps({"last_update": "2026-08-27", "codes": same_codes}),
                encoding="utf-8",
            )
            failed_path = root / "failed.json"
            failed_path.write_text(
                json.dumps({"last_update": "2026-08-25", "codes": failed_codes}),
                encoding="utf-8",
            )
            (root / update_codes.STATUS_FILE).write_text(
                json.dumps(
                    {
                        "files": {
                            "failed.json": {
                                "last_checked": "2026-08-26",
                                "last_update": "2026-08-25",
                            }
                        }
                    }
                ),
                encoding="utf-8",
            )
            failed_before = failed_path.read_bytes()

            with (
                patch.object(update_codes, "_tasks", return_value=tasks),
                patch.object(
                    update_codes,
                    "_run_tasks",
                    return_value={
                        "same.json": (_frame("same", "Same"), None),
                        "changed.json": (_frame("changed", "Changed"), None),
                        "failed.json": (None, "network error"),
                    },
                ),
                patch.object(
                    update_codes,
                    "_iso_now",
                    side_effect=(
                        "2026-08-28T09:00:00+08:00",
                        "2026-08-28T09:01:00+08:00",
                    ),
                ),
            ):
                status = update_codes.update_code_files(directory)

            self.assertFalse(status["files"]["same.json"]["updated"])
            self.assertEqual(status["files"]["same.json"]["last_update"], "2026-08-27")
            self.assertTrue(status["files"]["changed.json"]["updated"])
            self.assertEqual(status["files"]["changed.json"]["last_update"], "2026-08-28")
            self.assertTrue(status["files"]["failed.json"]["error"])
            self.assertEqual(status["files"]["failed.json"]["last_checked"], "2026-08-26")
            self.assertEqual(failed_path.read_bytes(), failed_before)

    def test_task_gets_two_retries(self):
        calls = 0

        def flaky():
            nonlocal calls
            calls += 1
            if calls < 3:
                raise RuntimeError("temporary")
            return _frame("ok", "OK")

        task = ("test.json", "测试", flaky, (), {})
        with patch.object(update_codes.time, "sleep"):
            frame, error = update_codes._run_task(task)
        self.assertIsNone(error)
        self.assertEqual(len(frame), 1)
        self.assertEqual(calls, 3)
