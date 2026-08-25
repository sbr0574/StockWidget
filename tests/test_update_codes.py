# -*- coding: utf-8 -*-
"""Tests for the standalone market-code update script."""

import importlib.util
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import call, patch

import pandas as pd


SCRIPT_PATH = Path(__file__).parents[1] / ".github" / "scripts" / "update_codes.py"
SPEC = importlib.util.spec_from_file_location("update_codes_script", SCRIPT_PATH)
update_codes = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(update_codes)


class TestUpdateCodes(unittest.TestCase):
    def test_cn_index_queries_have_explicit_markets(self):
        empty = pd.DataFrame(columns=update_codes._DF_COLUMNS)
        with patch.object(update_codes, "_em_stock_df", return_value=empty) as fetch:
            update_codes._index_cn_em()

        self.assertEqual(
            fetch.call_args_list,
            [
                call("m:1+t:1", "指", market="sh"),
                call("m:0+t:5", "指", market="sz"),
            ],
        )

    def test_etf_and_lof_queries_have_explicit_markets(self):
        empty = pd.DataFrame(columns=update_codes._DF_COLUMNS)
        for fetch_funds in (update_codes._fund_etf_em, update_codes._fund_lof_em):
            with self.subTest(fetch_funds=fetch_funds.__name__):
                with patch.object(update_codes, "_em_stock_df", return_value=empty) as fetch:
                    fetch_funds()

                self.assertEqual(fetch.call_count, 2)
                sh_call, sz_call = fetch.call_args_list
                self.assertTrue(
                    all(part.startswith("m:1+b:") for part in sh_call.args[0].split(","))
                )
                self.assertTrue(
                    all(part.startswith("m:0+b:") for part in sz_call.args[0].split(","))
                )
                self.assertEqual(sh_call.kwargs, {"market": "sh", "active_only": True})
                self.assertEqual(sz_call.kwargs, {"market": "sz", "active_only": True})

    def test_fund_close_fetches_shanghai_and_shenzhen_separately(self):
        empty = pd.DataFrame(columns=update_codes._DF_COLUMNS)
        with patch.object(update_codes, "_em_stock_df", return_value=empty) as fetch:
            update_codes._fund_close_em()

        self.assertEqual(
            fetch.call_args_list,
            [
                call("m:1+t:9+e:97", "基", market="sh", active_only=True),
                call("m:0+t:10+e:97", "基", market="sz", active_only=True),
            ],
        )

    def test_active_filter_removes_retired_and_unlisted_securities(self):
        today = int(datetime.now().strftime("%Y%m%d"))
        source = [
            {"f12": "000001", "f14": "平安银行", "f2": 10, "f18": 9, "f26": "19910403"},
            {"f12": "000002", "f14": "正常停牌", "f2": "-", "f18": 8, "f26": "19910129"},
            {"f12": "000003", "f14": "待上市", "f2": "-", "f18": "-", "f26": "-"},
            {"f12": "000004", "f14": "未来上市", "f2": "-", "f18": "-", "f26": str(today + 1)},
            {"f12": "000005", "f14": "今日上市", "f2": "-", "f18": "-", "f26": str(today)},
            {"f12": "000006", "f14": "示例退", "f2": "-", "f18": 1, "f26": "19910101"},
        ]

        with patch.object(update_codes, "_em_clist_all", return_value=source):
            frame = update_codes._em_stock_df(
                "m:0+t:6", "深", market="sz", active_only=True
            )

        self.assertEqual(frame["code"].tolist(), ["000001", "000002", "000005"])
        self.assertEqual(frame["market"].tolist(), ["sz", "sz", "sz"])

    def test_cjk_detection_replaces_ascii_special_case(self):
        self.assertFalse(update_codes._has_cjk("Berkshire Hathaway"))
        self.assertTrue(update_codes._has_cjk("苹果公司 Apple"))


if __name__ == "__main__":
    unittest.main()
