"""行情格式、指标迁移、买卖盘和排序原始值。"""

from copy import deepcopy
import unittest

from stockwidget.core.quote_presentation import (
    format_value,
    should_use_english_units,
    BidAskCell,
    QuoteDisplayOptions,
    format_quote,
    COLOR_ROLE_TEXT,
    COLOR_ROLE_UP,
    COLOR_ROLE_DOWN,
    COLOR_ROLE_NEUTRAL,
    DEFAULT_VISIBLE_METRICS,
    METRIC_IDS,
    metric_headers,
    legacy_visibility,
    normalize_visible_metrics,
    visible_metrics_from_config,
)


class FormatterTests(unittest.TestCase):
    def test_value_units_rounding_and_lot_conversion(self):
        cases = (
            (100000, True, 1, "10.00万"), (123456, True, 1, "12.35万"),
            (100000000, True, 1, "1.00亿"), (1000000000000, True, 1, "1.00万亿"),
            (999, False, 1, "999"), (1500, False, 1, "1.50k"),
            (2500000, False, 1, "2.50M"), (3000000000, False, 1, "3.00B"),
            (4000000000000, False, 1, "4.00T"), (100000, True, 100, "1000"),
        )
        for value, unit_cn, lot_size, expected in cases:
            with self.subTest(value=value, unit_cn=unit_cn, lot_size=lot_size):
                self.assertEqual(format_value(value, lot_size=lot_size, unit_cn=unit_cn), expected)

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


def _quote(volume: int = 123456, amount: float = 0) -> dict:
    return {
        "name": "Test",
        "opening_price": 11.0,
        "prev_close": 11.111,
        "current_price": 12.3456,
        "high_price": 12.5,
        "low_price": 10.5,
        "deals_vol": volume,
        "deals_amt": amount,
        "purchaser_vol": [0, 0, 0, 0, 0],
        "purchaser_price": [0, 0, 0, 0, 0],
        "seller_vol": [0, 0, 0, 0, 0],
        "seller_price": [0, 0, 0, 0, 0],
    }


class WidgetFormattingTests(unittest.TestCase):
    def test_single_bid_ask_cell_preserves_continuous_trade_markers_and_both_colors(self):
        quote = _quote()
        quote["purchaser_price"][0], quote["seller_price"][0] = 11, 13
        quote["purchaser_vol"][0], quote["seller_vol"][0] = 1234, 8900
        for price, marker in ((11, "◀"), (13, "▶"), (12, ""), (0, "")):
            quote["current_price"] = price
            row, _, _sort_values = format_quote(quote, "沪", "600000", market="sh")
            self.assertEqual(row["买一/卖一"], BidAskCell("12", "89", COLOR_ROLE_UP, COLOR_ROLE_DOWN, marker))
            self.assertNotIn("买一", row)
            self.assertNotIn("卖一", row)

    def test_book_volume_uses_lots_then_wan_with_two_decimals_in_all_markets(self):
        for market, security_type, lot_size in (("sh", "沪", 100), ("sz", "深", 100),
                                                ("bj", "北", 100), ("hk", "港", 1),
                                                ("us", "美", 1), ("sh", "期", 1)):
            for volume, label in ((9999, "9999"), (10000, "10000"),
                                  (10001, "1.00万"), (123456, "12.35万")):
                with self.subTest(market=market, security_type=security_type, volume=volume):
                    quote = _quote()
                    quote["purchaser_price"][0], quote["seller_price"][0] = 11, 13
                    quote["purchaser_vol"][0] = quote["seller_vol"][0] = volume * lot_size
                    row, _, _ = format_quote(quote, security_type, "test", market=market)
                    self.assertEqual(row["买一/卖一"], BidAskCell(label, label, "up", "down"))

    def test_auction_keeps_direction_slot_empty_and_shortens_signed_unpaired_volume(self):
        for unpaired, label, role in ((123456, "+12.35万", "up"),
                                      (-123456, "-12.35万", "down"), (0, "+0", "neutral")):
            with self.subTest(unpaired=unpaired):
                quote = _quote()
                quote["purchaser_price"][0] = quote["seller_price"][0] = 12
                quote["seller_vol"][0] = 1000100
                quote["purchaser_vol"][1] = max(0, unpaired) * 100
                quote["seller_vol"][1] = max(0, -unpaired) * 100
                row, _, _ = format_quote(quote, "沪", "600000", market="sh")
                self.assertEqual(row["买一/卖一"], BidAskCell("1.00万", label, role, role))

    def test_auction_formatting_preserves_source_and_uses_auction_price(self):
        quote = _quote()
        quote["purchaser_price"][0] = quote["seller_price"][0] = 12.0
        quote["seller_vol"][0] = 1000
        quote["purchaser_vol"][1] = 200
        original = deepcopy(quote)

        row, roles, _sort_values = format_quote(quote, "沪", "600000", market="sh", cost=10)

        self.assertEqual(quote, original)
        self.assertEqual(row["现价"], "12.00 ")
        self.assertEqual(row["买一/卖一"], BidAskCell("10", "+2", COLOR_ROLE_UP, COLOR_ROLE_UP))
        self.assertEqual(row["浮盈"], "+20.00%")
        self.assertEqual(roles["买一/卖一"], COLOR_ROLE_TEXT)

    def test_preopen_defaults_preserve_source_and_produce_flat_candle(self):
        quote = _quote()
        quote.update(current_price=0, opening_price=0, high_price=0, low_price=0)
        original = deepcopy(quote)

        row, roles, _sort_values = format_quote(quote, "沪", "600000", market="sh")

        self.assertEqual(quote, original)
        self.assertEqual(row["K线"]["k"], (quote["prev_close"],) * 5)
        self.assertEqual(row["涨幅"], "+0.00%")
        self.assertEqual(roles["现价"], COLOR_ROLE_NEUTRAL)

    def test_index_hides_unavailable_metrics_even_with_cost(self):
        row, roles, _sort_values = format_quote(_quote(0), "指", "000001", cost=10)
        for header in ("浮盈", "委比", "均价", "成交量", "成交额"):
            self.assertEqual(row[header], "-")
        self.assertEqual(roles["浮盈"], COLOR_ROLE_NEUTRAL)
        self.assertEqual(row["买一/卖一"], BidAskCell("-", "-"))

    def test_name_options_and_fund_precision(self):
        row, _, _sort_values = format_quote(
            _quote(), "基", "501001", market="sh",
            options=QuoteDisplayOptions(name_length=2, code_visible=True, type_visible=True),
        )
        self.assertEqual(row["名称"], "(基)501001 Te")
        self.assertEqual(row["现价"], "12.346 ")

    def test_market_units_and_explicit_overrides_apply_to_volume_and_amount(self):
        cases = (
            ("us", "美", "aapl", 123456, "auto",
             {"现价": "12.346 ", "涨跌": "+1.235", "成交量": "123.46k", "成交额": "123.46M"}),
            ("us", "指", "dji", 2500000, "auto", {"成交量": "2.50M"}),
            ("sh", "沪", "600000", 123400, "auto", {"成交量": "1234", "成交额": "1.23亿"}),
            ("", "期", "au0", 604495, "auto", {"成交量": "60.45万"}),
            ("sh", "期", "au0", 604495, "auto", {"成交量": "60.45万"}),
            ("sh", "沪", "600000", 123400, "en", {"成交量": "1.23k"}),
            ("us", "美", "aapl", 123456, "cn", {"成交量": "12.35万"}),
        )
        for market, security_type, code, volume, mode, expected in cases:
            with self.subTest(market=market, security_type=security_type, mode=mode):
                row, _, _ = format_quote(_quote(volume, amount=123456789), security_type, code,
                                         market=market, options=QuoteDisplayOptions(unit_mode=mode))
                self.assertEqual({header: row[header] for header in expected}, expected)

    def test_ordinary_text_and_directional_values_have_separate_color_roles(self):
        _row, roles, _sort_values = format_quote(
            _quote(), "沪", "600000", market="sh"
        )

        self.assertEqual(roles["名称"], COLOR_ROLE_TEXT)
        self.assertEqual(roles["成交量"], COLOR_ROLE_TEXT)
        self.assertEqual(roles["成交额"], COLOR_ROLE_TEXT)
        self.assertEqual(roles["现价"], COLOR_ROLE_UP)
        self.assertEqual(roles["涨跌"], COLOR_ROLE_UP)
        self.assertEqual(roles["涨幅"], COLOR_ROLE_UP)


class MetricLayoutTests(unittest.TestCase):
    def test_defaults_match_existing_visibility_and_include_name(self):
        self.assertEqual(DEFAULT_VISIBLE_METRICS, ("name", "price", "change_pct"))
        self.assertEqual(
            visible_metrics_from_config({}),
            ["name", "price", "change_pct"],
        )
        self.assertIn("name", METRIC_IDS)

    def test_legacy_flags_migrate_in_canonical_order(self):
        config = {
            "price_visible": False,
            "change_visible": True,
            "change_pct_visible": False,
            "b1s1_visible": True,
            "kline_visible": True,
        }
        self.assertEqual(
            visible_metrics_from_config(config),
            ["name", "change", "b1s1", "kline"],
        )

    def test_explicit_order_keeps_legacy_name_switch_behavior(self):
        # 旧配置的 visible_metrics 不含名称：未关闭 name_visible 时名称补在首位。
        self.assertEqual(
            visible_metrics_from_config({"visible_metrics": []}),
            ["name"],
        )
        self.assertEqual(
            visible_metrics_from_config(
                {"visible_metrics": ["kline", "price", "kline", "bad"]}
            ),
            ["name", "kline", "price"],
        )
        self.assertEqual(
            visible_metrics_from_config(
                {"name_visible": False, "visible_metrics": ["kline", "price"]}
            ),
            ["kline", "price"],
        )
        self.assertEqual(
            visible_metrics_from_config(
                {"name_visible": True, "visible_metrics": ["kline", "name"]}
            ),
            ["kline", "name"],
        )

    def test_normalize_and_project_combined_level_one_metric(self):
        normalized = normalize_visible_metrics(
            ["amount", "b1s1", "change_pct"]
        )
        self.assertEqual(normalized, ["amount", "b1s1", "change_pct"])
        self.assertEqual(
            metric_headers(normalized),
            ["成交额", "买一/卖一", "涨幅"],
        )

    def test_legacy_visibility_is_kept_in_sync(self):
        flags = legacy_visibility(["volume", "price"])
        self.assertTrue(flags["vol_visible"])
        self.assertTrue(flags["price_visible"])
        self.assertFalse(flags["amount_visible"])
