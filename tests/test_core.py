# -*- coding: utf-8 -*-
"""核心功能函数（市场代码、格式化、自选列表）的单元测试。"""

import unittest

from stockwidget.core.formatters import format_amount, format_volume
from stockwidget.core.markets import market_of, strip_market
from stockwidget.core.watchlist import normalize_watchlist


class TestMarkets(unittest.TestCase):
    def test_market_of(self):
        self.assertEqual(market_of("sh600519"), "sh")
        self.assertEqual(market_of("sz000001"), "sz")
        self.assertEqual(market_of("bj430047"), "bj")
        self.assertEqual(market_of("hk00700"), "hk")
        self.assertEqual(market_of("usaapl"), "us")
        self.assertEqual(market_of("gbnky"), "gb")     # 全球指数
        self.assertEqual(market_of("gnky"), "g")       # 兼容旧全球指数
        self.assertEqual(market_of("au0"), "")         # 期货裸码
        self.assertEqual(market_of(""), "")

    def test_strip_market(self):
        self.assertEqual(strip_market("sh600519"), "600519")
        self.assertEqual(strip_market("hk00700"), "00700")
        self.assertEqual(strip_market("usaapl"), "aapl")
        self.assertEqual(strip_market("gbnky"), "nky")
        self.assertEqual(strip_market("gnky"), "nky")
        self.assertEqual(strip_market("au0"), "au0")


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


if __name__ == "__main__":
    unittest.main()
