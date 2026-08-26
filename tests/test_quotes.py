# -*- coding: utf-8 -*-
"""行情请求代码只使用显式证券元数据。"""

import unittest
from unittest.mock import Mock, patch

from stockwidget.data import quotes
from stockwidget.data.quotes import _em_secid, _sina_code


class TestExplicitMarketMetadata(unittest.TestCase):
    def test_same_raw_code_keeps_shanghai_and_shenzhen_distinct(self):
        sh = {"market": "sh", "code": "000001"}
        sz = {"market": "sz", "code": "000001"}

        self.assertEqual(_sina_code(sh), "sh000001")
        self.assertEqual(_sina_code(sz), "sz000001")
        self.assertEqual(_em_secid(sh), "1.000001")
        self.assertEqual(_em_secid(sz), "0.000001")

    def test_futures_uses_empty_market_and_raw_code(self):
        future = {"market": "", "code": "au0"}
        self.assertEqual(_sina_code(future), "nf_AU0")
        self.assertEqual(_em_secid(future), "113.aum")


class TestEastmoneyQuotes(unittest.TestCase):
    @staticmethod
    def _response(payload):
        response = Mock()
        response.json.return_value = payload
        return response

    def test_stable_host_and_preopen_placeholders(self):
        response = self._response({
            "data": {
                "diff": [{
                    "f12": "600000",
                    "f13": 1,
                    "f14": "浦发银行",
                    "f2": "-",
                    "f5": "-",
                    "f6": "-",
                    "f15": "-",
                    "f16": "-",
                    "f17": "-",
                    "f18": 9.08,
                    "f31": "-",
                    "f32": "-",
                }]
            }
        })
        instruments = {
            "sh600000": {
                "market": "sh",
                "code": "600000",
                "type": "沪",
            }
        }
        with patch.object(quotes.requests, "get", return_value=response) as get:
            _, data = quotes.request_eastmoney(instruments)

        self.assertEqual(
            quotes._EM_QUOTE_URL,
            "https://push2delay.eastmoney.com/api/qt/ulist.np/get",
        )
        self.assertEqual(data["sh600000"]["current_price"], 0)
        self.assertEqual(data["sh600000"]["prev_close"], 9.08)
        self.assertEqual(data["sh600000"]["deals_vol"], 0)
        self.assertEqual(data["sh600000"]["purchaser_price"], [0, 0, 0, 0, 0])
        self.assertEqual(data["sh600000"]["seller_price"], [0, 0, 0, 0, 0])
        self.assertNotIn("f31", get.call_args.kwargs["params"]["fields"])
        self.assertEqual(get.call_count, 1)

    def test_eastmoney_does_not_call_sina(self):
        bulk = self._response({
            "data": {
                "diff": [{
                    "f12": "600000",
                    "f13": 1,
                    "f14": "浦发银行",
                    "f2": 9.05,
                    "f5": 1724,
                    "f6": 1560220,
                    "f15": 9.05,
                    "f16": 9.05,
                    "f17": 9.05,
                    "f18": 9.08,
                    "f31": 9.05,
                    "f32": 9.06,
                }]
            }
        })
        instruments = {
            "sh600000": {
                "market": "sh",
                "code": "600000",
                "type": "沪",
            }
        }
        with (
            patch.object(quotes.requests, "get", return_value=bulk) as get,
            patch.object(quotes, "request_sina") as request_sina,
        ):
            _, data = quotes.request_eastmoney(instruments)

        entry = data["sh600000"]
        self.assertEqual(entry["name"], "浦发银行")
        self.assertEqual(entry["current_price"], 9.05)
        self.assertEqual(entry["deals_vol"], 172400)
        self.assertEqual(entry["purchaser_price"], [0, 0, 0, 0, 0])
        self.assertEqual(entry["purchaser_vol"], [0, 0, 0, 0, 0])
        self.assertEqual(entry["seller_price"], [0, 0, 0, 0, 0])
        self.assertEqual(entry["seller_vol"], [0, 0, 0, 0, 0])
        self.assertEqual(get.call_count, 1)
        request_sina.assert_not_called()


if __name__ == "__main__":
    unittest.main()
