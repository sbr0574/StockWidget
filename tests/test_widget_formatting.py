# -*- coding: utf-8 -*-
"""按市场格式化行情显示的测试。"""

import unittest
from types import SimpleNamespace

from stockwidget.ui.table_model import (
    COLOR_ROLE_TEXT,
    COLOR_ROLE_UP,
)
from stockwidget.ui.widget import FloatLabel


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


def _widget(unit_mode: str = "auto"):
    return SimpleNamespace(
        type_visible=False,
        code_visible=False,
        name_length=-1,
        unit_mode=unit_mode,
        costs={},
    )


class WidgetFormattingTests(unittest.TestCase):
    def test_auto_mode_uses_english_units_for_us_securities(self):
        row, _ = FloatLabel._format_data(
            _widget(), "usaapl", _quote(123456), "美", "aapl", market="us"
        )
        self.assertEqual(row["现价"], "12.346 ")
        self.assertEqual(row["涨跌"], "+1.235")
        self.assertEqual(row["成交量"], "123.46k")

    def test_auto_mode_uses_english_units_for_international_indices(self):
        row, _ = FloatLabel._format_data(
            _widget(), "usdji", _quote(2500000), "指", "dji", market="us"
        )
        self.assertEqual(row["成交量"], "2.50M")

    def test_domestic_equities_display_lots_in_chinese(self):
        row, _ = FloatLabel._format_data(
            _widget(), "sh600000", _quote(123400), "沪", "600000", market="sh"
        )
        self.assertEqual(row["成交量"], "1234")

    def test_futures_volume_is_not_divided_by_one_hundred(self):
        row, _ = FloatLabel._format_data(
            _widget(), "au0", _quote(604495), "期", "au0", market=""
        )
        self.assertEqual(row["成交量"], "60.45万")

    def test_explicit_unit_modes_override_market(self):
        row, _ = FloatLabel._format_data(
            _widget("en"), "sh600000", _quote(123400), "沪", "600000", market="sh"
        )
        self.assertEqual(row["成交量"], "1.23k")
        row, _ = FloatLabel._format_data(
            _widget("cn"), "usaapl", _quote(123456), "美", "aapl", market="us"
        )
        self.assertEqual(row["成交量"], "12.35万")

    def test_amount_uses_the_same_unit_mode_as_volume(self):
        row, _ = FloatLabel._format_data(
            _widget(), "usaapl", _quote(0, amount=123456789), "美", "aapl", market="us"
        )
        self.assertEqual(row["成交额"], "123.46M")
        row, _ = FloatLabel._format_data(
            _widget(), "sh600000", _quote(0, amount=123456789), "沪", "600000", market="sh"
        )
        self.assertEqual(row["成交额"], "1.23亿")

    def test_ordinary_text_and_directional_values_have_separate_color_roles(self):
        _row, roles = FloatLabel._format_data(
            _widget(), "sh600000", _quote(), "沪", "600000", market="sh"
        )

        self.assertEqual(roles["名称"], COLOR_ROLE_TEXT)
        self.assertEqual(roles["成交量"], COLOR_ROLE_TEXT)
        self.assertEqual(roles["成交额"], COLOR_ROLE_TEXT)
        self.assertEqual(roles["现价"], COLOR_ROLE_UP)
        self.assertEqual(roles["涨跌"], COLOR_ROLE_UP)
        self.assertEqual(roles["涨幅"], COLOR_ROLE_UP)


if __name__ == "__main__":
    unittest.main()
