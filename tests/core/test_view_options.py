"""迷你图类型兼容旧配置并沿用显示配置保存流程。"""

import unittest

from stockwidget.core.view_options import ViewOptions


class MiniChartOptionsTests(unittest.TestCase):
    def test_average_period_roundtrip_and_invalid_values_use_intraday_mean(self):
        for period in (0, 5, 10, 20, 30, 60):
            with self.subTest(period=period):
                options = ViewOptions.from_config({"average_period": period})
                self.assertEqual(options.average_period, period)
                self.assertEqual(ViewOptions.from_config(options.to_config()), options)
        for period in (None, "invalid", 7, [], -5):
            self.assertEqual(ViewOptions.from_config({"average_period": period}).average_period, 0)

    def test_mini_chart_types_round_trip_and_legacy_defaults_to_daily_candle(self):
        for kind in ("kline", "trend", "amplitude"):
            with self.subTest(kind=kind):
                options = ViewOptions.from_config({"mini_chart_type": kind, "chart_enabled": False})
                self.assertEqual(options.mini_chart_type, kind)
                self.assertEqual(ViewOptions.from_config(options.to_config()), options)
        for config in ({}, {"mini_chart_type": "unknown"}, {"mini_chart_type": None},
                       {"mini_chart_type": ["trend"]}):
            with self.subTest(config=config):
                self.assertEqual(ViewOptions.from_config(config).mini_chart_type, "kline")
