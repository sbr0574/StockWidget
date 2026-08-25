# -*- coding: utf-8 -*-
"""代码搜索/建议逻辑的单元测试。"""

import unittest

from stockwidget.core.code_search import find_suggestions, normalize_stock_entry

CODES = {
    "sh600519": {"code": "600519", "market": "sh", "name": "贵州茅台", "type": "沪",
                 "py": "guizhoumaotai", "abbr": "gzmt"},
    "sh600036": {"code": "600036", "market": "sh", "name": "招商银行", "type": "沪",
                 "py": "zhaoshangyinhang", "abbr": "zsyh"},
    "usaapl": {"code": "aapl", "market": "us", "name": "苹果", "type": "美",
               "name_en": "Apple Inc.", "abbr": ""},
    "usmsft": {"code": "msft", "market": "us", "name": "微软", "type": "美",
               "engname": "Microsoft Corporation", "abbr": "wr"},
    "gbnky": {"code": "nky", "market": "gb", "name": "日经225指数", "type": "指",
              "py": "rijing225zhishu", "abbr": "rj225zs", "name_en": ""},
}


class TestNormalizeEntry(unittest.TestCase):
    def test_key_built_from_market_and_code(self):
        e = normalize_stock_entry({"market": "sh", "code": "600519", "name": "茅台"})
        self.assertEqual(e["key"], "sh600519")
        self.assertEqual(e["code"], "600519")

    def test_code_zfill_for_a_share(self):
        e = normalize_stock_entry({"market": "sz", "code": "1"})
        self.assertEqual(e["code"], "000001")

    def test_name_en_accepts_legacy_engname(self):
        e = normalize_stock_entry({"engname": "Microsoft Corporation"})
        self.assertEqual(e["name_en"], "microsoft corporation")


class TestFindSuggestions(unittest.TestCase):
    def test_by_code(self):
        self.assertEqual(find_suggestions(CODES, "600519")[0]["key"], "sh600519")

    def test_by_pinyin_abbr(self):
        self.assertEqual(find_suggestions(CODES, "gzmt")[0]["name"], "贵州茅台")

    def test_by_pinyin_full(self):
        self.assertEqual(find_suggestions(CODES, "zhaoshang")[0]["name"], "招商银行")

    def test_by_english_name(self):
        self.assertEqual(find_suggestions(CODES, "apple")[0]["key"], "usaapl")

    def test_by_legacy_english_name(self):
        self.assertEqual(find_suggestions(CODES, "microsoft")[0]["key"], "usmsft")

    def test_by_global_index_prefix(self):
        self.assertEqual(find_suggestions(CODES, "gbnky")[0]["key"], "gbnky")
        self.assertEqual(find_suggestions(CODES, "nky")[0]["key"], "gbnky")

    def test_empty_query(self):
        self.assertEqual(find_suggestions(CODES, ""), [])

    def test_no_match(self):
        self.assertEqual(find_suggestions(CODES, "zzzzzznothing"), [])

if __name__ == "__main__":
    unittest.main()
