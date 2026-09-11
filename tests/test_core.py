# -*- coding: utf-8 -*-
"""格式化和自选列表测试。"""

import unittest

from stockwidget.core.formatters import (
    format_value,
    should_use_english_units,
)
from stockwidget.core.watchlist import normalize_watchlist


class FormatterTests(unittest.TestCase):
    def test_format_value_chinese_units(self):
        self.assertEqual(format_value(100000, unit_cn=True), "10.00万")
        self.assertEqual(format_value(123456, unit_cn=True), "12.35万")
        self.assertEqual(format_value(100000000, unit_cn=True), "1.00亿")
        self.assertEqual(format_value(1000000000000, unit_cn=True), "1.00万亿")

    def test_format_value_english_units(self):
        self.assertEqual(format_value(999, unit_cn=False), "999")
        self.assertEqual(format_value(1500, unit_cn=False), "1.50k")
        self.assertEqual(format_value(2500000, unit_cn=False), "2.50M")
        self.assertEqual(format_value(3000000000, unit_cn=False), "3.00B")
        self.assertEqual(format_value(4000000000000, unit_cn=False), "4.00T")

    def test_format_value_applies_lot_size(self):
        self.assertEqual(format_value(100000, lot_size=100, unit_cn=True), "1000")
        self.assertEqual(format_value(100000, lot_size=1, unit_cn=True), "10.00万")

    def test_should_use_english_units_by_mode_and_market(self):
        # 自动：美股与国际指数用英文，国内/港股等用中文
        self.assertTrue(should_use_english_units("auto", "us"))
        self.assertTrue(should_use_english_units("auto", "gb"))
        self.assertFalse(should_use_english_units("auto", "sh"))
        self.assertFalse(should_use_english_units("auto", "hk"))
        self.assertFalse(should_use_english_units("auto", ""))
        # 显式模式不受市场影响
        self.assertTrue(should_use_english_units("en", "sh"))
        self.assertFalse(should_use_english_units("cn", "us"))
        # 中文别名与非法值回退到自动
        self.assertFalse(should_use_english_units("中文", "us"))
        self.assertTrue(should_use_english_units("bad", "us"))


class WatchlistTests(unittest.TestCase):
    def test_invalid_costs_are_rejected_consistently(self):
        for cost in (None, "", "invalid", 0, -1, "nan", "inf", float("-inf")):
            with self.subTest(cost=cost):
                entry = normalize_watchlist({"sh600519": {"cost": cost}})["sh600519"]
                self.assertIsNone(entry["cost"])

    def test_normalize(self):
        watchlist = normalize_watchlist(
            {
                "SH600519": {
                    "checked": "1",
                    "cost": "1500.5",
                    "name": " 贵州茅台 ",
                    "type": "沪",
                },
                "": {"checked": True},
                "sz000001": {"cost": None},
            }
        )
        self.assertTrue(watchlist["sh600519"]["checked"])
        self.assertEqual(watchlist["sh600519"]["cost"], 1500.5)
        self.assertEqual(watchlist["sh600519"]["name"], "贵州茅台")
        self.assertEqual(watchlist["sh600519"]["type"], "沪")
        self.assertNotIn("", watchlist)
        self.assertIsNone(watchlist["sz000001"]["cost"])

    def test_integer_cost_stays_int(self):
        watchlist = normalize_watchlist({"sh600519": {"cost": "1500"}})
        self.assertEqual(watchlist["sh600519"]["cost"], 1500)
        self.assertIsInstance(watchlist["sh600519"]["cost"], int)

    def test_invalid_cost_becomes_none(self):
        watchlist = normalize_watchlist({"sh600519": {"cost": "abc"}})
        self.assertIsNone(watchlist["sh600519"]["cost"])

    def test_none_watchlist(self):
        self.assertEqual(normalize_watchlist(None), {})

    def test_old_watchlist_is_hydrated_from_codes(self):
        codes = {
            "sz000001": {
                "code": "000001",
                "market": "sz",
                "name": "平安银行",
                "type": "深",
            }
        }
        watchlist = normalize_watchlist({"sz000001": {"checked": True}}, codes)
        self.assertEqual(watchlist["sz000001"]["code"], "000001")
        self.assertEqual(watchlist["sz000001"]["market"], "sz")


if __name__ == "__main__":
    unittest.main()
