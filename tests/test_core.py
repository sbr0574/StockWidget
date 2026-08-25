# -*- coding: utf-8 -*-
"""核心功能函数（格式化、自选列表）的单元测试。"""

import unittest

from stockwidget.core.formatters import format_amount, format_volume
from stockwidget.core.watchlist import normalize_watchlist


class TestFormatters(unittest.TestCase):
    def test_format_volume(self):
        self.assertEqual(format_volume(100000), "1000")       # 1000 手
        self.assertEqual(format_volume(1000000), "1.00万")    # 10000 手
        self.assertEqual(format_volume(10000000000), "1.00亿")  # 1e8 手

    def test_format_amount(self):
        self.assertEqual(format_amount(1000000), "100.00万")
        self.assertEqual(format_amount(100000000), "1.00亿")
        self.assertEqual(format_amount(1000000000000), "1.00万亿")


class TestWatchlist(unittest.TestCase):
    def test_normalize(self):
        wl = normalize_watchlist({
            "SH600519": {"checked": "1", "cost": "1500.5", "name": " 贵州茅台 ", "type": "沪"},
            "": {"checked": True},
            "sz000001": {"cost": None},
        })
        self.assertEqual(wl["sh600519"]["checked"], True)
        self.assertEqual(wl["sh600519"]["cost"], 1500.5)
        self.assertEqual(wl["sh600519"]["name"], "贵州茅台")
        self.assertEqual(wl["sh600519"]["type"], "沪")
        self.assertNotIn("", wl)  # 空 key 被丢弃
        self.assertIsNone(wl["sz000001"]["cost"])

    def test_integer_cost_stays_int(self):
        wl = normalize_watchlist({"sh600519": {"cost": "1500"}})
        self.assertEqual(wl["sh600519"]["cost"], 1500)
        self.assertIsInstance(wl["sh600519"]["cost"], int)

    def test_invalid_cost_becomes_none(self):
        wl = normalize_watchlist({"sh600519": {"cost": "abc"}})
        self.assertIsNone(wl["sh600519"]["cost"])

    def test_none_watchlist(self):
        self.assertEqual(normalize_watchlist(None), {})

    def test_old_watchlist_is_hydrated_from_codes(self):
        codes = {
            "sz000001": {"code": "000001", "market": "sz", "name": "平安银行", "type": "深"},
        }
        wl = normalize_watchlist({"sz000001": {"checked": True}}, codes)
        self.assertEqual(wl["sz000001"]["code"], "000001")
        self.assertEqual(wl["sz000001"]["market"], "sz")


if __name__ == "__main__":
    unittest.main()
