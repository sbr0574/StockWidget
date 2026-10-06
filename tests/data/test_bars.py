"""History parsing, source preference, full-window MAs and cache boundaries."""

from datetime import datetime, timedelta, timezone
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch

import requests

from stockwidget.core.config_store import history_cache_dir
from stockwidget.data.bar_cache import BarCache, cache_state
from stockwidget.data.bars import (
    Bar, BarResult, DAILY_HISTORY, fetch_bars, market_now, moving_average,
    parse_eastmoney, parse_sina, request_eastmoney_bars, request_sina_bars, select_bars,
)


A_SHARE = {"market": "sh", "code": "600519", "type": "沪"}


def sample_bars(count=100):
    day = datetime(2026, 1, 1)
    dates = []
    while len(dates) < count:
        if day.weekday() < 5:
            dates.append(day.strftime("%Y-%m-%d"))
        day += timedelta(days=1)
    return tuple(Bar(date, i + 1, i + 2, i + .5, i + 1, 100) for i, date in enumerate(dates))


class HistoryTests(unittest.TestCase):
    def test_eastmoney_fields_sort_deduplicate_and_ignore_bad_rows(self):
        payload = {"data": {"trends": ["bad", "2026-09-30 09:31,10,11,12,9,123,456,10.5",
                                        "2026-09-30 09:30,10,10,11,9,1,10,10",
                                        "2026-09-30 09:31,11,11,12,10,123,456,10.5",
                                        "2026-09-30 09:32,10,--,12,9,123,456,10.5",
                                        "2026-09-30 09:33,10,11,NaN,9,123,456,10.5"]}}
        bars = parse_eastmoney(payload)
        self.assertEqual(len(bars), 2)
        self.assertEqual(bars[1], Bar("2026-09-30 09:31:00", 11, 12, 10, 11, 123, 456, 10.5))
        self.assertEqual(parse_eastmoney({"data": None}), ())
        self.assertIsNone(parse_eastmoney({"data": {"klines": ["2026-09-30,10,11,12,9,1,2,99"]}}, daily=True)[0].average)

    def test_sina_jsonp_dictionary_and_futures_array(self):
        text = 'var foo=([{day:"2026-09-30 09:31:00",open:"10",high:"12",low:"9",close:"11",volume:"5"}]);'
        result = parse_sina(text)
        self.assertEqual(result[0].close, 11)
        self.assertEqual(result[0].volume, 5)
        self.assertIsNone(result[0].amount)
        futures = parse_sina('=([["2026-09-30",10,12,9,11,123,777,888]]);')
        self.assertEqual(futures[0].volume, 123)
        short_keys = parse_sina('[{"d":"2026-09-30","o":10,"h":12,"l":9,"c":11,"v":7}]')
        self.assertEqual(short_keys[0].close, 11)
        self.assertEqual(parse_sina("null"), ())
        with self.assertRaises(ValueError):
            parse_sina('[__import__("os")]')

    def test_us_clock_dst_and_beijing_minute_conversion(self):
        instrument = {"market": "us", "code": "aapl"}
        for utc, hour in ((datetime(2026, 7, 1, 14, tzinfo=timezone.utc), 10),
                          (datetime(2026, 1, 2, 15, tzinfo=timezone.utc), 10)):
            self.assertEqual(market_now(instrument, utc).hour, hour)
        payload = {"data": {"trends": ["2026-07-02 03:59,0,10,0,0,123,1230,10"]}}
        bar = parse_eastmoney(payload, instrument=instrument)[0]
        self.assertEqual(bar.time, "2026-07-01 15:59:00")
        self.assertEqual((bar.open, bar.high, bar.low), (10, 10, 10))

    def test_minute_selection_and_futures_night_session(self):
        bars = tuple(Bar(f"2026-09-{day:02d} 10:00:00", 10, 10, 10, 10) for day in range(20, 30))
        self.assertEqual(len(select_bars(bars, "five_day", A_SHARE)), 5)
        self.assertEqual(len(select_bars(bars, "intraday", A_SHARE)), 1)
        future = {"market": "", "code": "au0", "type": "期"}
        night = (Bar("2026-09-28 21:00:00", 10, 10, 10, 10),
                 Bar("2026-09-29 02:00:00", 10, 10, 10, 10),
                 Bar("2026-09-29 10:00:00", 10, 10, 10, 10))
        self.assertEqual(select_bars(night, "intraday", future), night)

    def test_ma60_uses_history_before_visible_30_candles(self):
        bars = sample_bars()
        averages = moving_average(bars, 60)
        self.assertTrue(all(value is None for value in averages[:59]))
        self.assertEqual(averages[59], 30.5)
        self.assertEqual(averages[-30], 41.5)
        self.assertEqual(averages[-1], 70.5)
        self.assertEqual(moving_average(bars[:4], 5), [None] * 4)
        with self.assertRaises(ValueError):
            moving_average(bars, 0)

    def test_source_preference_fallback_and_safe_empty_state(self):
        bars = sample_bars()
        for source in ("sina", "eastmoney"):
            with self.subTest(source=source), patch("stockwidget.data.bars.request_sina_bars", return_value=bars) as sina, \
                    patch("stockwidget.data.bars.request_eastmoney_bars", return_value=bars) as em:
                result = fetch_bars(A_SHARE, "daily", source)
                self.assertEqual(result.source, source)
                (em if source == "sina" else sina).assert_not_called()
        with patch("stockwidget.data.bars.request_sina_bars", side_effect=requests.Timeout("secret URL")), \
                patch("stockwidget.data.bars.request_eastmoney_bars", return_value=bars):
            self.assertEqual(fetch_bars(A_SHARE, "daily", "sina").source, "eastmoney")
        with patch("stockwidget.data.bars.request_sina_bars", return_value=()), \
                patch("stockwidget.data.bars.request_eastmoney_bars", side_effect=ValueError("secret URL")):
            result = fetch_bars(A_SHARE, "daily", "sina")
            self.assertIn("暂无可用历史数据", result.message)
            self.assertNotIn("secret", result.message)

    def test_short_five_day_window_tries_backup_then_reports_available_days(self):
        short = tuple(Bar(f"2026-09-{day} 10:00:00", 10, 10, 10, 10) for day in range(27, 30))
        full = tuple(Bar(f"2026-09-{day} 10:00:00", 10, 10, 10, 10) for day in range(25, 30))
        with patch("stockwidget.data.bars.request_sina_bars", return_value=short), \
                patch("stockwidget.data.bars.request_eastmoney_bars", return_value=full):
            result = fetch_bars(A_SHARE, "five_day", "sina")
            self.assertEqual(result.source, "eastmoney")
            self.assertEqual(len(result.bars), 5)
        with patch("stockwidget.data.bars.request_sina_bars", return_value=short), \
                patch("stockwidget.data.bars.request_eastmoney_bars", side_effect=requests.Timeout):
            result = fetch_bars(A_SHARE, "five_day", "sina")
            self.assertEqual(len(result.bars), 3)
            self.assertIn("3 个交易日", result.message)

    @patch("stockwidget.data.bars.requests.get")
    def test_request_parameters_and_asset_compatibility(self, get):
        get.return_value.text = '=[{"day":"2026-09-30","open":10,"high":12,"low":9,"close":11}]'
        get.return_value.json.return_value = {"data": {"klines": ["2026-09-30,10,11,12,9,100,1100,1"]}}
        for instrument in (A_SHARE, {"market": "sz", "code": "399001", "type": "指"},
                           {"market": "", "code": "au0", "type": "期"}):
            self.assertTrue(request_sina_bars(instrument, "daily"))
        self.assertEqual(get.call_args.kwargs["params"]["symbol"], "AU0")
        request_sina_bars(A_SHARE, "daily")
        self.assertEqual(get.call_args.kwargs["params"]["datalen"], DAILY_HISTORY)
        get.reset_mock()
        self.assertEqual(request_sina_bars({"market": "us", "code": "aapl"}, "daily"), ())
        get.assert_not_called()
        for instrument, secid in ((A_SHARE, "1.600519"), ({"market": "hk", "code": "00700"}, "116.00700"),
                                  ({"market": "us", "code": "aapl"}, "105.AAPL"),
                                  ({"market": "", "code": "au0", "type": "期"}, "113.aum")):
            self.assertTrue(request_eastmoney_bars(instrument, "daily"))
            self.assertEqual(get.call_args.kwargs["params"]["secid"], secid)
            self.assertEqual(get.call_args.kwargs["params"]["lmt"], 100)
            self.assertEqual(get.call_args.kwargs["params"]["fqt"], 0)
        get.return_value.json.side_effect = [{"data": None}, {"data": {"klines": ["2026-09-30,10,11,12,9,100,1100,1"]}}]
        self.assertTrue(request_eastmoney_bars({"market": "us", "code": "ibm"}, "daily"))
        self.assertEqual(get.call_args.kwargs["params"]["secid"], "106.IBM")


class CacheTests(unittest.TestCase):
    def setUp(self):
        self.tmp = self.enterContext(tempfile.TemporaryDirectory())
        self.now = datetime(2026, 9, 30, 2, tzinfo=timezone.utc)  # 10:00 Beijing.
        self.fetch = Mock(return_value=BarResult(sample_bars(), "sina"))
        self.cache = BarCache(self.tmp, clock=lambda: self.now, fetcher=self.fetch)

    def test_cache_path_and_keys_separate_markets_sources_views(self):
        with patch("stockwidget.core.config_store.config_paths", return_value=self.tmp):
            self.assertEqual(history_cache_dir(), str(Path(self.tmp) / "cache"))
        key = self.cache.key(A_SHARE, "intraday", "sina")
        self.assertEqual(key, self.cache.key(A_SHARE, "five_day", "sina"))
        for inst, view, source in (({**A_SHARE, "market": "sz"}, "intraday", "sina"),
                                   (A_SHARE, "daily", "sina"), (A_SHARE, "intraday", "eastmoney")):
            self.assertNotEqual(key, self.cache.key(inst, view, source))

    def test_shared_minutes_hit_then_expire_during_trading(self):
        bars = tuple(Bar(f"2026-09-{day} 10:00:00", 10, 10, 10, 10) for day in range(24, 30))
        self.fetch.return_value = BarResult(bars, "sina")
        self.cache.get(A_SHARE, "intraday", "sina")
        self.assertEqual(self.fetch.call_args.args[1], "five_day")
        result = self.cache.get(A_SHARE, "five_day", "sina")
        self.assertTrue(result.cached)
        self.assertEqual(len(result.bars), 5)
        self.assertEqual(self.fetch.call_count, 1)
        self.now += timedelta(seconds=61)
        self.cache.get(A_SHARE, "intraday", "sina")
        self.assertEqual(self.fetch.call_count, 2)

    def test_daily_ttl_close_and_next_day_invalidate(self):
        self.cache.get(A_SHARE, "daily", "sina")
        self.now += timedelta(seconds=61)
        self.assertTrue(self.cache.get(A_SHARE, "daily", "sina").cached)
        self.now += timedelta(seconds=240)
        self.cache.get(A_SHARE, "daily", "sina")
        self.assertEqual(self.fetch.call_count, 2)
        self.now = datetime(2026, 9, 30, 8, tzinfo=timezone.utc)
        self.cache.get(A_SHARE, "daily", "sina")
        self.now += timedelta(hours=2)
        self.assertTrue(self.cache.get(A_SHARE, "daily", "sina").cached)
        self.now += timedelta(days=1)
        self.assertFalse(self.cache.get(A_SHARE, "daily", "sina").cached)

    def test_failed_refresh_preserves_stale_data_and_corruption_recovers(self):
        self.cache.get(A_SHARE, "daily", "sina")
        self.now += timedelta(minutes=6)
        self.fetch.return_value = BarResult(message="暂无可用历史数据")
        result = self.cache.get(A_SHARE, "daily", "sina")
        self.assertTrue(result.stale)
        self.assertEqual(len(result.bars), 100)
        path = next(Path(self.tmp).glob("*.json"))
        self.assertEqual(len(json.loads(path.read_text())["bars"]), 100)
        self.cache.get(A_SHARE, "daily", "sina")
        self.assertEqual(self.fetch.call_count, 2)  # Five-minute outage backoff.
        path.write_text("{broken", encoding="utf-8")
        self.fetch.return_value = BarResult(sample_bars(), "sina")
        self.assertTrue(self.cache.get(A_SHARE, "daily", "sina").bars)

    def test_negative_cache_recovers_and_disk_failure_is_nonfatal(self):
        self.fetch.return_value = BarResult(message="暂无可用历史数据")
        self.cache.get(A_SHARE, "daily", "sina")
        self.assertTrue(self.cache.get(A_SHARE, "daily", "sina").cached)
        self.assertEqual(self.fetch.call_count, 1)
        self.now += timedelta(minutes=6)
        self.fetch.return_value = BarResult(sample_bars(), "sina")
        with patch("stockwidget.data.bar_cache.os.replace", side_effect=OSError):
            self.assertTrue(self.cache.get(A_SHARE, "daily", "sina").bars)

    def test_us_and_futures_trading_windows(self):
        us = {"market": "us", "code": "aapl"}
        self.assertTrue(cache_state(us, datetime(2026, 7, 1, 14, tzinfo=timezone.utc))[1])
        self.assertFalse(cache_state(us, datetime(2026, 7, 1, 22, tzinfo=timezone.utc))[1])
        futures = {"market": "", "code": "au0", "type": "期"}
        self.assertTrue(cache_state(futures, datetime(2026, 9, 28, 14, tzinfo=timezone.utc))[1])
        self.assertTrue(cache_state(futures, datetime(2026, 9, 28, 17, tzinfo=timezone.utc))[1])

    def test_minute_cache_accumulates_short_windows_without_mixing_sources(self):
        short = lambda days: tuple(Bar(f"2026-09-{day} 10:00:00", 10, 10, 10, 10) for day in days)
        self.fetch.return_value = BarResult(short(range(24, 27)), "sina", "接口仅返回 3 个交易日")
        self.cache.get(A_SHARE, "five_day", "sina")
        self.now += timedelta(minutes=2)
        self.fetch.return_value = BarResult(short(range(27, 30)), "sina", "接口仅返回 3 个交易日")
        result = self.cache.get(A_SHARE, "five_day", "sina")
        self.assertEqual(len(result.bars), 5)
        self.assertEqual(result.message, "")
        self.now += timedelta(minutes=2)
        self.fetch.return_value = BarResult(short(range(27, 30)), "eastmoney", "接口仅返回 3 个交易日")
        self.assertEqual(len(self.cache.get(A_SHARE, "five_day", "sina").bars), 3)
