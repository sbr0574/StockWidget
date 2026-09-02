# -*- coding: utf-8 -*-
"""按市场格式化行情显示的测试。"""

import unittest
from types import SimpleNamespace

from stockwidget.ui.widget import FloatLabel


def _quote(volume: int = 123456) -> dict:
    return {
        "name": "Test",
        "opening_price": 11.0,
        "prev_close": 11.111,
        "current_price": 12.3456,
        "high_price": 12.5,
        "low_price": 10.5,
        "deals_vol": volume,
        "deals_amt": 0,
        "purchaser_vol": [0, 0, 0, 0, 0],
        "purchaser_price": [0, 0, 0, 0, 0],
        "seller_vol": [0, 0, 0, 0, 0],
        "seller_price": [0, 0, 0, 0, 0],
    }


class WidgetFormattingTests(unittest.TestCase):
    def setUp(self):
        self.widget = SimpleNamespace(
            type_visible=False,
            code_visible=False,
            name_length=-1,
            costs={},
        )

    def test_all_us_securities_use_three_decimal_prices_and_share_volume(self):
        row, _ = FloatLabel._format_data(
            self.widget, "usaapl", _quote(), "美", "aapl", market="us"
        )
        self.assertEqual(row["现价"], "12.346 ")
        self.assertEqual(row["涨跌"], "+1.235")
        self.assertEqual(row["成交量"], "12.35万")

    def test_domestic_equities_display_lots(self):
        row, _ = FloatLabel._format_data(
            self.widget, "sh600000", _quote(123400), "沪", "600000", market="sh"
        )
        self.assertEqual(row["成交量"], "1234")

    def test_futures_volume_is_not_divided_by_one_hundred(self):
        row, _ = FloatLabel._format_data(
            self.widget, "au0", _quote(604495), "期", "au0", market=""
        )
        self.assertEqual(row["成交量"], "60.45万")


if __name__ == "__main__":
    unittest.main()
