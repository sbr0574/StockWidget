"""The chart's volume ratio uses real completed daily volumes and quote time."""

from datetime import datetime, timedelta
import unittest

from stockwidget.data.bars import Bar, quote_moving_average, quote_volume_ratio


class VolumeRatioTests(unittest.TestCase):
    instrument = {"market": "sh", "code": "600000", "type": "沪"}
    quote = {"date": "2026-10-28", "time": "10:30:00", "opening_price": 10, "deals_vol": 1200}
    bars = tuple(Bar((datetime(2026, 10, 20) + timedelta(days=i)).date().isoformat(),
                     10, 11, 9, 10, 1000) for i in range(5))

    def test_five_completed_days_exclude_today_and_pause_for_lunch(self):
        bars = (*self.bars, Bar("2026-10-28", 10, 11, 9, 10, 100000))
        for clock, expected in (("10:30:00", 4.8), ("11:30:00", 2.4), ("12:20:00", 2.4),
                                ("13:00:00", 2.4), ("14:00:00", 1.6), ("15:00:00", 1.2)):
            with self.subTest(clock=clock):
                self.assertAlmostEqual(quote_volume_ratio(bars, {**self.quote, "time": clock}, self.instrument), expected)

    def test_insufficient_unsupported_and_preopen_quotes_have_no_ratio(self):
        for bars, quote, instrument in ((self.bars[:4], self.quote, self.instrument),
                                        (self.bars, {**self.quote, "time": "09:29:00"}, self.instrument),
                                        (self.bars, {**self.quote, "time": "09:30:00"}, self.instrument),
                                        (self.bars, {**self.quote, "opening_price": 0}, self.instrument),
                                        (self.bars, {**self.quote, "date": "invalid"}, self.instrument),
                                        (self.bars, self.quote, {**self.instrument, "market": "hk"}),
                                        (self.bars, self.quote, {**self.instrument, "type": "指"}),
                                        (self.bars, self.quote, {**self.instrument, "type": "期"})):
            with self.subTest(quote=quote, instrument=instrument, count=len(bars)):
                self.assertIsNone(quote_volume_ratio(bars, quote, instrument))

    def test_daily_average_includes_live_price_once_and_requires_full_period(self):
        bars = tuple(Bar(f"2026-10-{day:02d}", close, close, close, close, 1000)
                     for close, day in enumerate((14, 15, 16, 19, 20, 21, 22, 23, 26, 27), 1))
        quote = {**self.quote, "current_price": 20, "opening_price": 20, "high_price": 20,
                 "low_price": 20, "prev_close": 10}
        self.assertAlmostEqual(quote_moving_average(bars, quote, self.instrument, 5), 10.8)
        self.assertAlmostEqual(quote_moving_average(bars, {**quote, "current_price": 25}, self.instrument, 5), 11.8)
        self.assertIsNone(quote_moving_average(bars, quote, self.instrument, 20))
        self.assertIsNone(quote_moving_average(bars, quote, self.instrument, 0))
