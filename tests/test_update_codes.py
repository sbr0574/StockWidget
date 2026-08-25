# -*- coding: utf-8 -*-
"""Tests for the standalone market-code update script."""

import importlib.util
import unittest
from pathlib import Path
from unittest.mock import Mock, call, patch

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

    def test_em_lists_no_longer_request_active_status_fields(self):
        source = [
            {"f12": "920001", "f14": "正常股票"},
            {"f12": "920002", "f14": "待上市股票"},
        ]
        with patch.object(update_codes, "_em_clist_all", return_value=source) as fetch:
            frame = update_codes._em_stock_df("m:0+t:81+s:2048", "京", market="bj")

        fetch.assert_called_once_with("m:0+t:81+s:2048", fields="f12,f14")
        self.assertEqual(frame["code"].tolist(), ["920001", "920002"])

    def test_shanghai_stock_request_uses_exchange_category(self):
        response = Mock()
        response.json.return_value = {
            "result": [{"A_STOCK_CODE": "600001", "SEC_NAME_CN": "示例股份"}]
        }
        with patch.object(update_codes.requests, "get", return_value=response) as get:
            frame = update_codes._stock_sh_name_code("主板A股")

        response.raise_for_status.assert_called_once_with()
        self.assertEqual(get.call_args.kwargs["params"]["STOCK_TYPE"], "1")
        self.assertEqual(frame.iloc[0].to_dict(), {
            "code": "600001",
            "name": "示例股份",
            "name_en": "",
            "type": "沪",
            "market": "sh",
        })

    def test_shenzhen_stock_table_splits_board_column(self):
        table = pd.DataFrame([
            {"板块": "主板", "A股代码": "1", "A股简称": "主板示例"},
            {"板块": "创业板", "A股代码": "300001", "A股简称": "创业板示例"},
        ])
        with patch.object(update_codes, "_szse_xlsx", return_value=table):
            frame = update_codes._stock_sz_a_name_code()

        self.assertEqual(frame["code"].tolist(), ["000001", "300001"])
        self.assertTrue(frame["market"].eq("sz").all())

    def test_sse_funds_are_classified_by_official_fund_type(self):
        response = Mock()
        response.json.return_value = {
            "result": [
                {"fundType": "00", "fundCode": "510001", "secNameFull": "ETF示例"},
                {"fundType": "10", "fundCode": "501001", "secNameFull": "LOF示例"},
                {"fundType": "50", "fundCode": "508001", "secNameFull": "REIT示例"},
                {"fundType": "20", "fundCode": "519001", "fundAbbr": "场外基金"},
            ],
            "pageHelp": {"pageCount": 1},
        }
        with patch.object(update_codes.requests, "get", return_value=response):
            rows = update_codes._fund_sse_rows()

        self.assertEqual([row[0] for row in rows["ETF基金"]], ["510001"])
        self.assertEqual([row[0] for row in rows["LOF基金"]], ["501001"])
        self.assertEqual([row[0] for row in rows["封闭式基金"]], ["508001"])

    def test_szse_funds_use_workbook_category(self):
        table = pd.DataFrame([
            {"基金代码": "159001", "基金简称": "ETF示例", "基金类别": "ETF"},
            {"基金代码": "160001", "基金简称": "LOF示例", "基金类别": "LOF"},
            {"基金代码": "180001", "基金简称": "REIT示例", "基金类别": "不动产基金"},
        ])
        with patch.object(update_codes, "_szse_xlsx", return_value=table):
            rows = update_codes._fund_szse_rows()

        self.assertEqual([row[0] for row in rows["ETF基金"]], ["159001"])
        self.assertEqual([row[0] for row in rows["LOF基金"]], ["160001"])
        self.assertEqual([row[0] for row in rows["封闭式基金"]], ["180001"])

    def test_cjk_detection_replaces_ascii_special_case(self):
        self.assertFalse(update_codes._has_cjk("Berkshire Hathaway"))
        self.assertTrue(update_codes._has_cjk("苹果公司 Apple"))


if __name__ == "__main__":
    unittest.main()
