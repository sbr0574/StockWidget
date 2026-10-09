"""自选规范化、代码搜索和编辑交互。"""

import unittest

from stockwidget.core.watchlist import (
    normalize_watchlist,
    build_search_index,
    normalize_stock_entry,
    query_search_index,
    search_suggestions,
)


class WatchlistTests(unittest.TestCase):
    def test_invalid_costs_are_rejected_consistently(self):
        for cost in (None, "", "invalid", 0, -1, "nan", "inf", float("-inf")):
            with self.subTest(cost=cost):
                entry = normalize_watchlist({"sh600519": {"cost": cost}})["sh600519"]
                self.assertIsNone(entry["cost"])

    def test_normalize(self):
        self.assertEqual(normalize_watchlist(None), {})
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
                "sh600036": {"cost": "1500"},
            }
        )
        self.assertTrue(watchlist["sh600519"]["checked"])
        self.assertEqual(watchlist["sh600519"]["cost"], 1500.5)
        self.assertEqual(watchlist["sh600519"]["name"], "贵州茅台")
        self.assertEqual(watchlist["sh600519"]["type"], "沪")
        self.assertNotIn("", watchlist)
        self.assertIsNone(watchlist["sz000001"]["cost"])
        self.assertEqual(watchlist["sh600036"]["cost"], 1500)
        self.assertIsInstance(watchlist["sh600036"]["cost"], int)

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
    def test_entry_identity_and_legacy_names_are_normalized(self):
        cases = (
            ({"market": "sh", "code": "600519", "name": "茅台"},
             {"key": "sh600519", "code": "600519"}),
            ({"market": "sz", "code": "1"}, {"code": "000001"}),
            ({"engname": "Microsoft Corporation"}, {"name_en": "microsoft corporation"}),
        )
        for raw, expected in cases:
            with self.subTest(raw=raw):
                entry = normalize_stock_entry(raw)
                self.assertEqual({key: entry[key] for key in expected}, expected)


class SearchSuggestionsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.index = build_search_index(CODES)

    def test_search_fields_and_keyword_intersections(self):
        cases = (
            ("600519", "sh600519"), ("gzmt", "sh600519"),
            ("zhaoshang", "sh600036"), ("apple", "usaapl"),
            ("microsoft", "usmsft"), ("gbnky", "gbnky"), ("nky", "gbnky"),
            ("万科", "sz000002"), ("600519 茅台", "sh600519"),
            ("茅台 gzmt", "sh600519"), ("  600519\t  茅台　", "sh600519"),
        )
        for query, key in cases:
            with self.subTest(query=query):
                self.assertEqual([item["key"] for item in search_suggestions(self.index, query)], [key])

    def test_empty_unmatched_and_conflicting_keywords_return_no_suggestions(self):
        for query in ("", "zzzzzznothing", "茅台 zsyh"):
            with self.subTest(query=query):
                self.assertEqual(search_suggestions(self.index, query), [])

    def test_unknown_type_stays_in_stock_category(self):
        codes = {"custom": {"code": "custom", "name": "自定义", "type": ""}}
        result = query_search_index(
            build_search_index(codes), "自定义", categories={"stock"}
        )
        self.assertEqual(result.items[0]["key"], "custom")

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
