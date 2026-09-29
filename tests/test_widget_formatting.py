# -*- coding: utf-8 -*-
"""按市场格式化行情显示的测试。"""

import unittest
from copy import deepcopy

from stockwidget.core.quote_presentation import (
    QuoteDisplayOptions, format_quote,
    COLOR_ROLE_TEXT,
    COLOR_ROLE_UP,
    COLOR_ROLE_DOWN,
    COLOR_ROLE_NEUTRAL,
)


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
    def test_auction_formatting_preserves_source_and_uses_auction_price(self):
        quote = _quote()
        quote["purchaser_price"][0] = quote["seller_price"][0] = 12.0
        quote["seller_vol"][0] = 1000
        quote["purchaser_vol"][1] = 200
        original = deepcopy(quote)

        row, roles = format_quote(quote, "沪", "600000", market="sh", cost=10)

        self.assertEqual(quote, original)
        self.assertEqual(row["现价"], "12.00 ")
        self.assertEqual(row["买一"], "10")
        self.assertEqual(row["卖一"], "+2")
        self.assertEqual(row["浮盈"], "+20.00%")
        self.assertEqual(roles["买一"], COLOR_ROLE_UP)

    def test_preopen_defaults_preserve_source_and_produce_flat_candle(self):
        quote = _quote()
        quote.update(current_price=0, opening_price=0, high_price=0, low_price=0)
        original = deepcopy(quote)

        row, roles = format_quote(quote, "沪", "600000", market="sh")

        self.assertEqual(quote, original)
        self.assertEqual(row["K线"]["k"], (quote["prev_close"],) * 5)
        self.assertEqual(row["涨幅"], "+0.00%")
        self.assertEqual(roles["现价"], COLOR_ROLE_NEUTRAL)

    def test_index_hides_unavailable_metrics_even_with_cost(self):
        row, roles, sort_values = format_quote(_quote(0), "指", "000001", cost=10, include_sort=True)
        for header in ("成本", "持仓数量", "浮盈", "持仓盈亏", "买一", "卖一", "委比", "均价", "成交量", "成交额"):
            self.assertEqual(row[header], "-")
        self.assertEqual(roles["浮盈"], COLOR_ROLE_NEUTRAL)
        self.assertIsNone(sort_values["成本"])
        self.assertIsNone(sort_values["持仓数量"])

    def test_cost_and_quantity_are_displayed_as_configured(self):
        quote = _quote()
        quote["current_price"] = 12.0

        row, roles, sort_values = format_quote(
            quote, "沪", "600000", market="sh", cost=10.0, quantity=1500, include_sort=True
        )
        self.assertEqual(row["成本"], "10.00")
        self.assertEqual(row["持仓数量"], "1500")
        self.assertEqual(roles["成本"], COLOR_ROLE_TEXT)
        self.assertEqual(roles["持仓数量"], COLOR_ROLE_TEXT)
        self.assertEqual(sort_values["成本"], 10.0)
        self.assertEqual(sort_values["持仓数量"], 1500)

        # 小数成本/数量按原值显示
        row, _ = format_quote(
            deepcopy(quote), "沪", "600000", market="sh", cost=10.5, quantity=100.5
        )
        self.assertEqual(row["成本"], "10.50")
        self.assertEqual(row["持仓数量"], "100.5")

    def test_cost_and_quantity_are_hidden_without_position(self):
        for kwargs in (
            {},
            {"cost": None, "quantity": None},
            {"cost": 0, "quantity": 0},
        ):
            with self.subTest(**kwargs):
                row, roles, sort_values = format_quote(
                    deepcopy(_quote()), "沪", "600000", market="sh",
                    include_sort=True, **kwargs,
                )
                self.assertEqual(row["成本"], "-")
                self.assertEqual(row["持仓数量"], "-")
                self.assertEqual(roles["成本"], COLOR_ROLE_TEXT)
                self.assertEqual(roles["持仓数量"], COLOR_ROLE_TEXT)
                self.assertIsNone(sort_values["成本"])
                self.assertIsNone(sort_values["持仓数量"])

    def test_position_pl_is_cost_times_quantity(self):
        quote = _quote()
        quote["current_price"] = 12.0

        row, roles = format_quote(quote, "沪", "600000", market="sh", cost=10.0, quantity=100)
        self.assertEqual(row["持仓盈亏"], "+200.00")
        self.assertEqual(roles["持仓盈亏"], COLOR_ROLE_UP)

        row, roles = format_quote(quote, "沪", "600000", market="sh", cost=12.5, quantity=200)
        self.assertEqual(row["持仓盈亏"], "-100.00")
        self.assertEqual(roles["持仓盈亏"], COLOR_ROLE_DOWN)

        row, roles = format_quote(quote, "沪", "600000", market="sh", cost=12.0, quantity=100)
        self.assertEqual(row["持仓盈亏"], "+0.00")
        self.assertEqual(roles["持仓盈亏"], COLOR_ROLE_NEUTRAL)

    def test_position_pl_requires_cost_and_quantity(self):
        quote = _quote()
        quote["current_price"] = 12.0
        for kwargs in (
            {"cost": 10.0},
            {"quantity": 100},
            {"cost": None, "quantity": 100},
            {"cost": 0, "quantity": 100},
            {"cost": 10.0, "quantity": 0},
        ):
            with self.subTest(**kwargs):
                row, roles = format_quote(
                    deepcopy(quote), "沪", "600000", market="sh", **kwargs
                )
                self.assertEqual(row["持仓盈亏"], "-")
                self.assertEqual(roles["持仓盈亏"], COLOR_ROLE_NEUTRAL)

    def test_name_options_and_fund_precision(self):
        row, _ = format_quote(
            _quote(), "基", "501001", market="sh",
            options=QuoteDisplayOptions(name_length=2, code_visible=True, type_visible=True),
        )
        self.assertEqual(row["名称"], "(基)501001 Te")
        self.assertEqual(row["现价"], "12.346 ")

    def test_auto_mode_uses_english_units_for_us_securities(self):
        row, _ = format_quote(
            _quote(123456), "美", "aapl", market="us"
        )
        self.assertEqual(row["现价"], "12.346 ")
        self.assertEqual(row["涨跌"], "+1.235")
        self.assertEqual(row["成交量"], "123.46k")

    def test_auto_mode_uses_english_units_for_international_indices(self):
        row, _ = format_quote(
            _quote(2500000), "指", "dji", market="us"
        )
        self.assertEqual(row["成交量"], "2.50M")

    def test_domestic_equities_display_lots_in_chinese(self):
        row, _ = format_quote(
            _quote(123400), "沪", "600000", market="sh"
        )
        self.assertEqual(row["成交量"], "1234")

    def test_futures_volume_is_not_divided_by_one_hundred(self):
        for market in ("", "sh"):
            with self.subTest(market=market):
                row, _ = format_quote(_quote(604495), "期", "au0", market=market)
                self.assertEqual(row["成交量"], "60.45万")

    def test_explicit_unit_modes_override_market(self):
        row, _ = format_quote(
            _quote(123400), "沪", "600000", market="sh", options=QuoteDisplayOptions(unit_mode="en")
        )
        self.assertEqual(row["成交量"], "1.23k")
        row, _ = format_quote(
            _quote(123456), "美", "aapl", market="us", options=QuoteDisplayOptions(unit_mode="cn")
        )
        self.assertEqual(row["成交量"], "12.35万")

    def test_amount_uses_the_same_unit_mode_as_volume(self):
        row, _ = format_quote(
            _quote(0, amount=123456789), "美", "aapl", market="us"
        )
        self.assertEqual(row["成交额"], "123.46M")
        row, _ = format_quote(
            _quote(0, amount=123456789), "沪", "600000", market="sh"
        )
        self.assertEqual(row["成交额"], "1.23亿")

    def test_ordinary_text_and_directional_values_have_separate_color_roles(self):
        _row, roles = format_quote(
            _quote(), "沪", "600000", market="sh"
        )

        self.assertEqual(roles["名称"], COLOR_ROLE_TEXT)
        self.assertEqual(roles["成交量"], COLOR_ROLE_TEXT)
        self.assertEqual(roles["成交额"], COLOR_ROLE_TEXT)
        self.assertEqual(roles["现价"], COLOR_ROLE_UP)
        self.assertEqual(roles["涨跌"], COLOR_ROLE_UP)
        self.assertEqual(roles["涨幅"], COLOR_ROLE_UP)


if __name__ == "__main__":
    unittest.main()
