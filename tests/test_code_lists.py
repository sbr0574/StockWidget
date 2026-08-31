# -*- coding: utf-8 -*-
"""代码列表启动加载、远端同步和重试调度测试。"""

import unittest
from datetime import datetime, timedelta, timezone
from unittest.mock import Mock, call, patch

from stockwidget.constants import CODES_RAW_URLS, CODES_STATUS_FILE, CODE_LIST_FILES
from stockwidget.data.code_lists import (
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
    def load_local_file(filename):
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
            filename = url.rsplit("/", 1)[-1]
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
        self.assertEqual(save.call_args_list[-1], call(remote_status, CODES_STATUS_FILE))
        for filename in CODE_LIST_FILES:
            self.assertIn(call(payloads[filename], filename), save.call_args_list)

    def test_started_at_distinguishes_two_runs_on_same_date(self):
        remote_status = _status(
            self.local_date, "2026-08-28T10:01:00+08:00"
        )
        payloads = {
            name: _payload(name, self.local_date) for name in CODE_LIST_FILES
        }

        def fetcher(url):
            filename = url.rsplit("/", 1)[-1]
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
        self.assertEqual(fetcher.call_count, len(CODES_RAW_URLS))
        self.assertIn("raw.githubusercontent.com", fetcher.call_args_list[0].args[0])
        self.assertIn("gitee", fetcher.call_args_list[1].args[0])

    def test_one_list_failure_saves_nothing_and_keeps_cached(self):
        remote_status = _status(
            "2026-08-31", "2026-08-31T09:01:00+08:00"
        )

        def fetcher(url):
            filename = url.rsplit("/", 1)[-1]
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
            filename = url.rsplit("/", 1)[-1]
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


class FetchJsonTests(unittest.TestCase):
    @patch("stockwidget.data.code_lists.requests.get")
    def test_default_timeout_is_five_seconds(self, get):
        response = Mock()
        response.json.return_value = {"ok": True}
        get.return_value = response

        result = fetch_json_from_url("https://example.com/test.json")

        self.assertEqual(result, {"ok": True})
        get.assert_called_once_with(
            "https://example.com/test.json",
            timeout=5,
            headers={"Cache-Control": "no-cache", "User-Agent": "StockWidget"},
        )


if __name__ == "__main__":
    unittest.main()
