# -*- coding: utf-8 -*-
"""市场代码缓存的逐文件更新与北京时间策略测试。"""

import os
import tempfile
import unittest
from datetime import datetime, timezone
from unittest.mock import patch

from stockwidget.constants import LIST_FILES
from stockwidget.core.config_store import load_file, save_file
from stockwidget.data import code_lists


class TestCodeLists(unittest.TestCase):
    def setUp(self):
        self._old_appdata = os.environ.get("APPDATA")
        self._tmp = tempfile.TemporaryDirectory()
        os.environ["APPDATA"] = self._tmp.name
        self.app_name = "StockWidgetTest"

    def tearDown(self):
        self._tmp.cleanup()
        if self._old_appdata is None:
            os.environ.pop("APPDATA", None)
        else:
            os.environ["APPDATA"] = self._old_appdata

    def test_stale_files_are_checked_individually(self):
        stock_file, futures_file = LIST_FILES
        save_file(
            {"last_update": "2026-08-25", "codes": {"sh1": {}}},
            self.app_name,
            stock_file,
        )
        save_file(
            {"last_update": "2026-08-24", "codes": {"au": {}}},
            self.app_name,
            futures_file,
        )

        self.assertEqual(
            code_lists.stale_code_files(self.app_name, "2026-08-25"),
            (futures_file,),
        )

    def test_download_saves_only_requested_file_with_expected_date(self):
        stock_file, futures_file = LIST_FILES
        payload = {"last_update": "2026-08-25", "codes": {"sh1": {}}}
        with patch.object(code_lists, "fetch_json_from_url", return_value=payload) as fetch:
            updated = code_lists.download_code_files(
                self.app_name,
                (stock_file,),
                expected_date="2026-08-25",
            )

        self.assertEqual(updated, (stock_file,))
        self.assertEqual(load_file(self.app_name, stock_file), payload)
        self.assertEqual(load_file(self.app_name, futures_file), {})
        self.assertEqual(fetch.call_count, 1)

    def test_download_does_not_replace_cache_with_old_remote_file(self):
        stock_file = LIST_FILES[0]
        cached = {"last_update": "2026-08-24", "codes": {"old": {}}}
        remote = {"last_update": "2026-08-24", "codes": {"remote": {}}}
        save_file(cached, self.app_name, stock_file)
        with patch.object(code_lists, "fetch_json_from_url", return_value=remote):
            updated = code_lists.download_code_files(
                self.app_name,
                (stock_file,),
                expected_date="2026-08-25",
            )

        self.assertEqual(updated, ())
        self.assertEqual(load_file(self.app_name, stock_file), cached)

    def test_best_codes_are_selected_per_file(self):
        stock_file, futures_file = LIST_FILES
        save_file(
            {"last_update": "2026-08-25", "codes": {"local-stock": {}}},
            self.app_name,
            stock_file,
        )
        resources = {
            f":/{stock_file}": {
                "last_update": "2026-08-24",
                "codes": {"resource-stock": {}},
            },
            f":/{futures_file}": {
                "last_update": "2026-08-24",
                "codes": {"resource-future": {}},
            },
        }
        with patch.object(
            code_lists,
            "load_json_from_resource",
            side_effect=lambda path: resources[path],
        ):
            codes = code_lists.load_best_codes(self.app_name)

        self.assertEqual(set(codes), {"local-stock", "resource-future"})

    def test_missing_one_local_file_is_not_reported_online(self):
        stock_file = LIST_FILES[0]
        save_file(
            {"last_update": "2026-08-25", "codes": {"sh1": {}}},
            self.app_name,
            stock_file,
        )
        with (
            patch.object(code_lists, "beijing_today", return_value="2026-08-25"),
            patch.object(code_lists, "load_json_from_resource", return_value={}),
        ):
            state, _ = code_lists.code_data_state(self.app_name)

        self.assertEqual(state, "cached")

    def test_refresh_waits_until_nine_in_beijing(self):
        self.assertEqual(
            code_lists.code_refresh_delay_seconds(
                datetime(2026, 8, 25, 0, 30, tzinfo=timezone.utc)
            ),
            30 * 60,
        )
        self.assertEqual(
            code_lists.code_refresh_delay_seconds(
                datetime(2026, 8, 25, 1, 0, tzinfo=timezone.utc)
            ),
            0,
        )

    def test_weekend_does_not_mark_any_code_file_stale(self):
        saturday = datetime(2026, 8, 29, 2, 0, tzinfo=timezone.utc)
        self.assertFalse(code_lists.is_code_update_day(saturday))
        self.assertEqual(
            code_lists.stale_code_files(self.app_name, now=saturday),
            (),
        )

    def test_weekend_data_state_reports_current(self):
        saturday = datetime(2026, 8, 29, 2, 0, tzinfo=timezone.utc)
        resource = {"last_update": "2026-08-28", "codes": {"sh1": {}}}
        with (
            patch.object(code_lists, "_beijing_time", return_value=saturday),
            patch.object(code_lists, "load_json_from_resource", return_value=resource),
        ):
            state, date = code_lists.code_data_state(self.app_name)

        self.assertEqual((state, date), ("current", "2026-08-28"))


if __name__ == "__main__":
    unittest.main()
