# -*- coding: utf-8 -*-
"""代码列表启动加载、远端同步和重试调度测试。"""

import base64
import json
import os
import tempfile
import unittest

import requests
from datetime import datetime, timedelta, timezone
from urllib.parse import urlsplit
from unittest.mock import MagicMock, Mock, call, patch

from stockwidget.constants import CODES_SOURCE_URLS, CODES_STATUS_FILE, CODE_LIST_FILES
from stockwidget.data.code_lists import (
    CODES_HTTP_TIMEOUT,
    CodeDownloadError,
    CODES_RETRY_SECONDS,
    CodeListManager,
    fetch_json_from_url,
    next_code_check_delay,
)


from stockwidget.core.config_store import config_paths, data_cache_dir, load_file, save_file


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
    def test_weekday_before_nine_schedules_same_day(self):
        monday = datetime(2026, 8, 31, 8, 59, tzinfo=UTC8)
        self.assertEqual(next_code_check_delay(monday), 60)

    def test_weekday_at_nine_schedules_next_workday(self):
        monday = datetime(2026, 8, 31, 9, 0, tzinfo=UTC8)
        self.assertEqual(next_code_check_delay(monday), 24 * 60 * 60)

    def test_friday_at_nine_skips_weekend(self):
        friday = datetime(2026, 8, 28, 9, 0, tzinfo=UTC8)
        self.assertEqual(next_code_check_delay(friday), 3 * 24 * 60 * 60)


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
    def test_timeout_reports_reason(self, get):
        get.side_effect = requests.ReadTimeout()
        with self.assertRaisesRegex(CodeDownloadError, "超时"):
            fetch_json_from_url("https://example.com/test.json")

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


if __name__ == "__main__":
    unittest.main()
