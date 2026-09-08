# -*- coding: utf-8 -*-
"""代码搜索和建议逻辑测试。"""

import unittest

from stockwidget.core.code_search import (
    build_search_index,
    normalize_stock_entry,
    query_search_index,
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


class SearchSuggestionsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.index = build_search_index(CODES)

    def search(self, text):
        return search_suggestions(self.index, text)

    def test_by_code(self):
        self.assertEqual(self.search("600519")[0]["key"], "sh600519")

    def test_by_pinyin_abbr(self):
        self.assertEqual(self.search("gzmt")[0]["name"], "贵州茅台")

    def test_by_pinyin_full(self):
        self.assertEqual(self.search("zhaoshang")[0]["name"], "招商银行")

    def test_by_english_name(self):
        self.assertEqual(self.search("apple")[0]["key"], "usaapl")

    def test_by_legacy_english_name(self):
        self.assertEqual(self.search("microsoft")[0]["key"], "usmsft")

    def test_by_global_index_prefix(self):
        self.assertEqual(self.search("gbnky")[0]["key"], "gbnky")
        self.assertEqual(self.search("nky")[0]["key"], "gbnky")

    def test_multiple_keywords_can_match_different_fields(self):
        result = self.search("600519 茅台")
        self.assertEqual([item["key"] for item in result], ["sh600519"])
        result = self.search("茅台 gzmt")
        self.assertEqual([item["key"] for item in result], ["sh600519"])

    def test_multiple_keywords_use_and_semantics(self):
        self.assertEqual(self.search("茅台 zsyh"), [])

    def test_arbitrary_whitespace_splits_keywords(self):
        result = self.search("  600519\t  茅台　")
        self.assertEqual([item["key"] for item in result], ["sh600519"])

    def test_compact_query_matches_name_with_layout_spaces(self):
        self.assertEqual(self.search("万科")[0]["key"], "sz000002")

    def test_unknown_type_stays_in_stock_category(self):
        codes = {"custom": {"code": "custom", "name": "自定义", "type": ""}}
        result = query_search_index(
            build_search_index(codes), "自定义", categories={"stock"}
        )
        self.assertEqual(result.items[0]["key"], "custom")

    def test_prebuilt_index_can_be_reused(self):
        first = search_suggestions(self.index, "apple inc")
        second = search_suggestions(self.index, "apple inc")
        self.assertEqual(second, first)

    def test_empty_query(self):
        self.assertEqual(self.search(""), [])

    def test_no_match(self):
        self.assertEqual(self.search("zzzzzznothing"), [])


class PagedSearchTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.index = build_search_index(CODES)

    def test_empty_query_browses_all_records_and_paginates(self):
        first = query_search_index(self.index, page=1, page_size=3)
        second = query_search_index(self.index, page=2, page_size=3)

        self.assertEqual(first.total, len(CODES))
        self.assertEqual(first.page_count, 3)
        self.assertEqual(len(first.items), 3)
        self.assertEqual(len(second.items), 3)
        self.assertTrue(
            set(item["key"] for item in first.items).isdisjoint(
                item["key"] for item in second.items
            )
        )

    def test_category_and_region_filters_intersect(self):
        result = query_search_index(
            self.index,
            categories={"fund", "index"},
            regions={"sh"},
        )
        self.assertEqual(
            {item["key"] for item in result.items},
            {"sh501001"},
        )

    def test_region_uses_market_and_other_is_the_complement(self):
        sz = query_search_index(
            self.index, categories={"stock"}, regions={"sz"}
        )
        other = query_search_index(
            self.index, categories={"index", "futures"}, regions={"other"}
        )
        self.assertEqual({item["key"] for item in sz.items}, {"sz000002"})
        self.assertEqual(
            {item["key"] for item in other.items}, {"gbnky", "ad0"}
        )

    def test_us_entries_follow_existing_stock_label(self):
        stocks = query_search_index(
            self.index, categories={"stock"}, regions={"us"}
        )
        funds = query_search_index(
            self.index, categories={"fund"}, regions={"us"}
        )
        self.assertEqual(
            {item["key"] for item in stocks.items}, {"usaapl", "usmsft"}
        )
        self.assertEqual(funds.total, 0)

    def test_empty_filter_selection_returns_no_results(self):
        no_categories = query_search_index(self.index, categories=set())
        no_regions = query_search_index(self.index, regions=set())
        self.assertEqual(no_categories.total, 0)
        self.assertEqual(no_categories.page, 0)
        self.assertEqual(no_regions.total, 0)

    def test_page_is_clamped_after_result_count_changes(self):
        result = query_search_index(
            self.index, "茅台", page=99, page_size=3
        )
        self.assertEqual(result.page, 1)
        self.assertEqual(result.page_count, 1)
        self.assertEqual([item["key"] for item in result.items], ["sh600519"])


if __name__ == "__main__":
    unittest.main()
