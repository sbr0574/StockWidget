# -*- coding: utf-8 -*-
"""代码搜索和建议逻辑测试。"""

import unittest

from stockwidget.core.code_search import (
    build_search_index,
    find_suggestions,
    normalize_stock_entry,
    search_suggestions,
)


CODES = {
    "sh600519": {
        "code": "600519",
        "market": "sh",
        "name": "贵州茅台",
        "type": "沪",
        "py": "guizhoumaotai",
        "abbr": "gzmt",
    },
    "sh600036": {
        "code": "600036",
        "market": "sh",
        "name": "招商银行",
        "type": "沪",
        "py": "zhaoshangyinhang",
        "abbr": "zsyh",
    },
    "usaapl": {
        "code": "aapl",
        "market": "us",
        "name": "苹果",
        "type": "美",
        "name_en": "Apple Inc.",
        "abbr": "",
    },
    "usmsft": {
        "code": "msft",
        "market": "us",
        "name": "微软",
        "type": "美",
        "engname": "Microsoft Corporation",
        "abbr": "wr",
    },
    "gbnky": {
        "code": "nky",
        "market": "gb",
        "name": "日经225指数",
        "type": "指",
        "py": "rijing225zhishu",
        "abbr": "rj225zs",
        "name_en": "",
    },
    "sh501001": {
        "code": "501001",
        "market": "sh",
        "name": "财通精选混合LOF",
        "type": "基",
        "py": "caitongjingxuanhunhelof",
        "abbr": "ctjxhhlof",
    },
    "ad0": {
        "code": "ad0",
        "market": "",
        "name": "铸造铝合金连续",
        "type": "期",
        "py": "zhuzaolvhejinlianxu",
        "abbr": "zzlhjlx",
    },
    "sz000002": {
        "code": "000002",
        "market": "sz",
        "name": "万  科A",
        "type": "深",
        "py": "wankea",
        "abbr": "wka",
    },
}


class NormalizeEntryTests(unittest.TestCase):
    def test_key_built_from_market_and_code(self):
        entry = normalize_stock_entry(
            {"market": "sh", "code": "600519", "name": "茅台"}
        )
        self.assertEqual(entry["key"], "sh600519")
        self.assertEqual(entry["code"], "600519")

    def test_code_zfill_for_a_share(self):
        entry = normalize_stock_entry({"market": "sz", "code": "1"})
        self.assertEqual(entry["code"], "000001")

    def test_name_en_accepts_legacy_engname(self):
        entry = normalize_stock_entry({"engname": "Microsoft Corporation"})
        self.assertEqual(entry["name_en"], "microsoft corporation")


class FindSuggestionsTests(unittest.TestCase):
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

    def test_multiple_keywords_can_match_different_fields(self):
        result = find_suggestions(CODES, "600519 茅台")
        self.assertEqual([item["key"] for item in result], ["sh600519"])
        result = find_suggestions(CODES, "茅台 gzmt")
        self.assertEqual([item["key"] for item in result], ["sh600519"])

    def test_multiple_keywords_use_and_semantics(self):
        self.assertEqual(find_suggestions(CODES, "茅台 zsyh"), [])

    def test_arbitrary_whitespace_splits_keywords(self):
        result = find_suggestions(CODES, "  600519\t  茅台　")
        self.assertEqual([item["key"] for item in result], ["sh600519"])

    def test_compact_query_matches_name_with_layout_spaces(self):
        self.assertEqual(find_suggestions(CODES, "万科")[0]["key"], "sz000002")

    def test_category_filter_is_mutually_exclusive(self):
        self.assertEqual(
            find_suggestions(CODES, "茅台", category="stock")[0]["key"],
            "sh600519",
        )
        self.assertEqual(
            find_suggestions(CODES, "财通", category="fund")[0]["key"],
            "sh501001",
        )
        self.assertEqual(
            find_suggestions(CODES, "日经", category="index")[0]["key"],
            "gbnky",
        )
        self.assertEqual(
            find_suggestions(CODES, "铝合金", category="futures")[0]["key"],
            "ad0",
        )
        self.assertEqual(find_suggestions(CODES, "日经", category="stock"), [])

    def test_unknown_type_stays_in_stock_category(self):
        codes = {"custom": {"code": "custom", "name": "自定义", "type": ""}}
        self.assertEqual(
            find_suggestions(codes, "自定义", category="stock")[0]["key"],
            "custom",
        )

    def test_prebuilt_index_matches_direct_search(self):
        index = build_search_index(CODES)
        direct = find_suggestions(CODES, "apple inc", category="stock")
        indexed = search_suggestions(index, "apple inc", category="stock")
        self.assertEqual(indexed, direct)

    def test_empty_query(self):
        self.assertEqual(find_suggestions(CODES, ""), [])

    def test_no_match(self):
        self.assertEqual(find_suggestions(CODES, "zzzzzznothing"), [])


if __name__ == "__main__":
    unittest.main()
