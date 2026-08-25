# -*- coding: utf-8 -*-
"""行情请求代码只使用显式证券元数据。"""

import unittest

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


if __name__ == "__main__":
    unittest.main()
