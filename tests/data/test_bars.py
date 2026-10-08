"""History parsing, source preference, full-window MAs and cache boundaries."""

from dataclasses import replace
from datetime import datetime, timedelta, timezone
import sqlite3
from pathlib import Path
import tempfile
import threading
import unittest
from unittest.mock import Mock, patch

import requests

from stockwidget.core.config_store import history_cache_dir
from stockwidget.data.bar_cache import BarCache, cache_state
from stockwidget.data.bars import (
    Bar, BarResult, DAILY_HISTORY, HistorySeries, fetch_bars, history_day, intraday_average,
    intraday_timeline, market_now, merge_quote, minute_gaps, moving_average,
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

    def test_sina_daily_time_labels_normalize_before_live_candle_update(self):
        text = '=[{"day":"2026-09-30 15:00:00","open":10,"high":12,"low":9,"close":11,"volume":100}]'
        bars = parse_sina(text, daily=True)
        self.assertEqual(bars[0].time, "2026-09-30")
        quote = {"date": "2026-09-30", "time": "10:00:00", "opening_price": 10,
                 "current_price": 12, "high_price": 12, "low_price": 9, "deals_vol": 120}
        live = merge_quote(bars, A_SHARE, "daily", quote)
        self.assertEqual(len(live), 1)
        self.assertEqual(live[0].close, 12)
        self.assertEqual(parse_sina(text)[0].time, "2026-09-30 15:00:00")

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

    def test_sina_intraday_average_uses_cumulative_turnover_and_resets_each_day(self):
        bars = parse_sina('=[{"day":"2026-09-29 09:31:00","o":10,"h":11,"l":9,"c":10,"v":100,"a":1000},'
                          '{"day":"2026-09-29 09:32:00","o":20,"h":21,"l":19,"c":20,"v":300,"a":6000},'
                          '{"day":"2026-09-30 09:31:00","o":30,"h":31,"l":29,"c":30,"v":200,"a":6000}]')
        self.assertEqual(intraday_average(bars, A_SHARE), [10, 17.5, 30])
        self.assertTrue(all(bar.average is None for bar in bars))

    def test_intraday_average_preserves_supplied_values_and_handles_missing_turnover(self):
        bars = (Bar("2026-09-29 23:59:00", 10, 10, 10, 10, 100),
                Bar("2026-09-30 00:00:00", 20, 20, 20, 20, 300),
                Bar("2026-09-30 09:00:00", 30, 30, 30, 30, 0))
        futures = {"market": "", "code": "au0", "type": "期"}
        # Night and morning are one futures session; turnover includes a multiplier.
        multiplied = tuple(replace(bar, amount=bar.volume * bar.close * 100) for bar in bars)
        self.assertEqual(intraday_average(multiplied, futures), [10, 17.5, 17.5])
        supplied = (replace(bars[0], time="2026-09-30 09:31:00", average=12),
                    replace(bars[1], time="2026-09-30 09:32:00", average=18))
        self.assertEqual(intraday_average(supplied, A_SHARE), [12, 18])
        missing = tuple(replace(bar, average=None) for bar in supplied)
        self.assertEqual(intraday_average(missing, A_SHARE), [10, 17.5])
        self.assertEqual(intraday_average((replace(missing[0], volume=0),), A_SHARE), [10])

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
        get.return_value.text = '=[{"day":"2026-09-30 15:00:00","open":10,"high":12,"low":9,"close":11}]'
        get.return_value.json.return_value = {"data": {"klines": ["2026-09-30,10,11,12,9,100,1100,1"]}}
        for instrument in (A_SHARE, {"market": "sz", "code": "399001", "type": "指"},
                           {"market": "", "code": "au0", "type": "期"}):
            self.assertEqual(request_sina_bars(instrument, "daily")[-1].time, "2026-09-30")
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

    def test_daily_download_once_per_exchange_day_even_after_close(self):
        self.cache.get(A_SHARE, "daily", "sina")
        self.now += timedelta(seconds=61)
        self.assertTrue(self.cache.get(A_SHARE, "daily", "sina").cached)
        self.now += timedelta(seconds=240)
        self.cache.get(A_SHARE, "daily", "sina")
        self.assertEqual(self.fetch.call_count, 1)
        self.now = datetime(2026, 9, 30, 8, tzinfo=timezone.utc)
        self.assertTrue(self.cache.get(A_SHARE, "daily", "sina").cached)
        self.now += timedelta(hours=2)
        self.assertTrue(self.cache.get(A_SHARE, "daily", "sina").cached)
        self.now += timedelta(days=1)
        self.assertFalse(self.cache.get(A_SHARE, "daily", "sina").cached)
        self.assertEqual(self.fetch.call_count, 2)

    def test_failed_refresh_preserves_stale_data_and_corruption_recovers(self):
        self.cache.get(A_SHARE, "daily", "sina")
        self.now += timedelta(days=1)
        self.fetch.return_value = BarResult(message="暂无可用历史数据")
        result = self.cache.get(A_SHARE, "daily", "sina")
        self.assertTrue(result.stale)
        self.assertEqual(len(result.bars), 100)
        db = sqlite3.connect(self.cache.database)
        try:
            self.assertEqual(db.execute("SELECT COUNT(*) FROM bars").fetchone()[0], 100)
        finally:
            db.close()
        self.cache.get(A_SHARE, "daily", "sina")
        self.assertEqual(self.fetch.call_count, 2)  # Five-minute outage backoff.
        self.cache.database.write_bytes(b"broken SQLite cache")
        self.fetch.return_value = BarResult(sample_bars(), "sina")
        self.assertTrue(self.cache.get(A_SHARE, "daily", "sina").bars)

    def test_negative_cache_recovers_and_disk_failure_is_nonfatal(self):
        self.fetch.return_value = BarResult(message="暂无可用历史数据")
        self.cache.get(A_SHARE, "daily", "sina")
        self.assertTrue(self.cache.get(A_SHARE, "daily", "sina").cached)
        self.assertEqual(self.fetch.call_count, 1)
        self.now += timedelta(minutes=6)
        self.fetch.return_value = BarResult(sample_bars(), "sina")
        with patch("stockwidget.data.bar_cache.sqlite3.connect", side_effect=sqlite3.OperationalError("Disk unavailable")):
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

    def test_missing_minutes_repair_bypasses_ttl_but_preserves_failure_backoff(self):
        self.cache.get(A_SHARE, "five_day", "sina")
        self.cache.get(A_SHARE, "five_day", "sina", repair=True)
        self.assertEqual(self.fetch.call_count, 2)
        self.fetch.return_value = BarResult(message="暂无可用历史数据")
        self.assertTrue(self.cache.get(A_SHARE, "five_day", "sina", repair=True).stale)
        self.cache.get(A_SHARE, "five_day", "sina", repair=True)
        self.assertEqual(self.fetch.call_count, 3)

    def test_sqlite_roundtrip_preserves_precision_nulls_and_separate_series(self):
        precise = Bar("2026-09-30", 1234.567890123456, 1235.678901234567,
                      1233.456789012345, 1234.678901234567, 123456789.5, None, None)
        for instrument, source, actual in ((A_SHARE, "sina", "eastmoney"),
                                           (A_SHARE, "eastmoney", "eastmoney"),
                                           ({**A_SHARE, "market": "sz"}, "sina", "sina")):
            self.fetch.return_value = BarResult((precise,), actual)
            self.cache.get(instrument, "daily", source)
        restarted = BarCache(self.tmp, clock=lambda: self.now, fetcher=Mock(side_effect=AssertionError("Cache miss")))
        for instrument, source, actual in ((A_SHARE, "sina", "eastmoney"),
                                           (A_SHARE, "eastmoney", "eastmoney"),
                                           ({**A_SHARE, "market": "sz"}, "sina", "sina")):
            with self.subTest(market=instrument["market"], source=source):
                result = restarted.get(instrument, "daily", source)
                self.assertTrue(result.cached)
                self.assertEqual(result.bars, (precise,))
                self.assertEqual(result.source, actual)
        self.assertEqual(list(Path(self.tmp).glob("*.json")), [])

    def test_failed_sqlite_transaction_keeps_previous_complete_entry(self):
        self.cache.get(A_SHARE, "daily", "sina")
        db = sqlite3.connect(self.cache.database)
        try:
            db.execute("CREATE TRIGGER fail_write BEFORE INSERT ON bars BEGIN SELECT RAISE(ABORT, 'disk failure'); END")
        finally:
            db.close()
        original_day = self.now
        self.now += timedelta(days=1)
        new = (Bar("2026-10-01", 10, 11, 9, 10),)
        self.fetch.return_value = BarResult(new, "sina")
        self.assertEqual(self.cache.get(A_SHARE, "daily", "sina").bars, new)
        restarted = BarCache(self.tmp, clock=lambda: original_day, fetcher=Mock(side_effect=AssertionError("Cache miss")))
        self.assertEqual(restarted.get(A_SHARE, "daily", "sina").bars, sample_bars())

    def test_clear_removes_history_only_and_late_download_cannot_restore_it(self):
        self.cache.get(A_SHARE, "daily", "sina")
        legacy = Path(self.tmp) / ("a" * 64 + ".json")
        legacy.write_text("{}", encoding="utf-8")
        unrelated = Path(self.tmp) / "config.json"
        unrelated.write_text("{}", encoding="utf-8")
        started, release = threading.Event(), threading.Event()
        def download(*_args):
            started.set()
            release.wait(3)
            return BarResult(sample_bars(), "sina")
        self.cache.fetcher = download
        old_generation = self.cache.generation
        worker = threading.Thread(target=lambda: self.cache.get(A_SHARE, "five_day", "sina"))
        worker.start()
        try:
            self.assertTrue(started.wait(2))
            self.assertTrue(self.cache.clear())
            self.assertTrue(worker.is_alive())  # Clear doesn't wait for the network.
            release.set()
            worker.join(2)
            self.assertFalse(worker.is_alive())
            self.assertFalse(self.cache.database.exists())
            self.assertFalse(legacy.exists())
            self.assertTrue(unrelated.exists())
            self.cache.fetcher = Mock(return_value=BarResult(sample_bars(), "sina"))
            self.cache.get(A_SHARE, "daily", "sina", generation=old_generation)
            self.cache.fetcher.assert_not_called()  # A worker queued before clear.
            self.assertTrue(self.cache.get(A_SHARE, "daily", "sina").bars)
            self.assertTrue(self.cache.database.exists())
        finally:
            release.set()
            worker.join(3)

    def test_locked_database_is_preserved_and_clear_failure_is_reported(self):
        self.cache.get(A_SHARE, "daily", "sina")
        before = self.cache.database.read_bytes()
        with patch("stockwidget.data.bar_cache.sqlite3.connect", side_effect=sqlite3.OperationalError("database is locked")):
            self.assertTrue(self.cache.get(A_SHARE, "daily", "sina").bars)
        self.assertEqual(self.cache.database.read_bytes(), before)
        with patch.object(self.cache, "_remove_database", side_effect=PermissionError("in use")):
            self.assertFalse(self.cache.clear())
        self.assertTrue(self.cache.database.exists())


class IntradayTimelineTests(unittest.TestCase):
    def setUp(self):
        self.now = datetime(2026, 9, 30, 2, tzinfo=timezone.utc)

    def test_intraday_timeline_uses_exchange_close_and_skips_breaks_and_past_days(self):
        for instrument, stamp, now, opening, closing in (
                (A_SHARE, "2026-09-30 10:00:00", self.now, "09:30", "15:00"),
                ({"market": "hk"}, "2026-09-30 10:00:00", self.now, "09:30", "16:00"),
                ({"market": "us"}, "2026-07-01 10:00:00", datetime(2026, 7, 1, 14, tzinfo=timezone.utc), "09:30", "16:00"),
                ({"market": "us"}, "2026-01-05 10:00:00", datetime(2026, 1, 5, 15, tzinfo=timezone.utc), "09:30", "16:00")):
            with self.subTest(instrument=instrument, stamp=stamp):
                bars = (Bar(stamp, 10, 10, 10, 10),)
                timeline = intraday_timeline(bars, instrument, now)
                self.assertEqual((timeline[0][11:16], timeline[-1][11:16]), (opening, closing))
                self.assertIn(stamp, timeline)
                self.assertEqual(intraday_timeline(bars, instrument, now + timedelta(days=1)), ())
                if instrument.get("market") in {"sh", "hk"}:
                    self.assertFalse(any("12:30" in time for time in timeline))
        self.assertEqual(intraday_timeline((Bar("2026-09-30 10:00:00", 10, 10, 10, 10),), {"market": "gb"}, self.now), ())

    def test_futures_timeline_keeps_observed_night_and_extends_to_day_close(self):
        night = Bar("2026-09-25 21:30:00", 10, 10, 10, 10)  # Friday night belongs to Monday.
        for product, night_close in (("au", "02:30"), ("cu", "01:00"), ("rb", "23:00")):
            with self.subTest(product=product):
                instrument = {"market": "sh", "type": "期", "code": product + "0"}
                timeline = intraday_timeline((night,), instrument, datetime(2026, 9, 25, 14, tzinfo=timezone.utc))
                self.assertTrue(timeline[-1].endswith("2026-09-28 15:00:00"))
                self.assertTrue(any(stamp[11:16] == night_close for stamp in timeline))
                self.assertFalse(any(stamp.startswith("2026-09-27") for stamp in timeline))
                day_only = (replace(night, time="2026-09-28 10:00:00"),)
                self.assertTrue(intraday_timeline(day_only, instrument, datetime(2026, 9, 28, 2, tzinfo=timezone.utc))[0].endswith("09:00:00"))


class LiveHistoryTests(unittest.TestCase):
    def setUp(self):
        self.now = datetime(2026, 9, 30, 1, 31, tzinfo=timezone.utc)
        self.quote = {"date": "2026-09-30", "time": "09:31:10", "current_price": 11,
                      "opening_price": 10, "high_price": 12, "low_price": 9,
                      "deals_vol": 120, "deals_amt": 1320}
        self.minute = Bar("2026-09-30 09:31:00", 10, 11, 9, 10, 100, 1000, 10)

    def test_index_turnover_is_not_an_average_in_index_points(self):
        for market, code, price in (("sh", "000001", 3842), ("sz", "399001", 12887), ("sz", "399006", 3135)):
            with self.subTest(code=code):
                instrument = {"market": market, "code": code, "type": "指"}
                # Old caches can contain a component-price average near 15.
                minute = replace(self.minute, open=price, high=price, low=price, close=price, average=15)
                quote = {**self.quote, "opening_price": price, "current_price": price,
                         "deals_vol": 100, "deals_amt": 1500}
                self.assertIsNone(merge_quote((minute,), instrument, "intraday", quote)[-1].average)
                payload = {"data": {"trends": [f"2026-09-30 09:31,{price},{price},{price},{price},1,1500,15"]}}
                self.assertIsNone(parse_eastmoney(payload, instrument=instrument)[0].average)
                later = replace(minute, time="2026-09-30 09:32:00", close=price + 2, volume=300)
                self.assertEqual(intraday_average((minute, later), instrument), [price, price + 1])

    def test_zero_change_uses_latest_session_previous_close_without_guessing(self):
        cases = ((A_SHARE, "2026-09-29 15:00:00"),
                 ({"market": "hk", "code": "00700", "type": "港"}, "2026-09-29 16:00:00"),
                 ({"market": "us", "code": "aapl", "type": "美"}, "2026-09-29 15:59:00"))
        for instrument, previous_time in cases:
            with self.subTest(market=instrument["market"]):
                previous = replace(self.minute, time=previous_time, close=9)
                # Five-day reference belongs to its newest day, not the first day's open.
                first = replace(self.minute, time="2026-09-28 09:31:00", open=7, close=7)
                series = HistorySeries(instrument, "five_day")
                series.set_history(BarResult((first, previous, self.minute), "sina"), "2026-09-30", self.now)
                self.assertEqual(series.reference_price, 9)
        series = HistorySeries(A_SHARE, "intraday")
        previous = replace(self.minute, time="2026-09-29 15:00:00", close=9)
        series.set_history(BarResult((previous, self.minute), "sina"), "2026-09-30", self.now)
        series.update_quote({**self.quote, "prev_close": 8.5})
        self.assertEqual(series.reference_price, 8.5)
        series.update_quote({**self.quote, "time": "09:30:00", "prev_close": 5})
        self.assertEqual(series.reference_price, 8.5)  # Rejected old quote cannot change zero.
        for previous_close in (0, float("nan"), float("inf")):
            series.update_quote({**self.quote, "prev_close": previous_close})
            self.assertEqual(series.reference_price, 9)
        for bars in ((self.minute,), (replace(previous, time="2026-09-29 09:31:00"), self.minute), ()):
            with self.subTest(bars=bars):
                series = HistorySeries(A_SHARE, "five_day")
                series.set_history(BarResult(bars, "sina"), "2026-09-30", self.now)
                self.assertIsNone(series.reference_price)

    def test_current_minute_updates_then_appends_without_downloading(self):
        series = HistorySeries(A_SHARE, "intraday")
        series.update_quote(self.quote)
        self.assertTrue(series.needs_download(self.now))
        series.set_history(BarResult((self.minute,), "sina"), "2026-09-30", self.now)
        self.assertEqual(series.result.bars[-1], Bar(self.minute.time, 10, 11, 9, 11, 120, 1320, 11))
        for second, price, volume in ((20, 12, 130), (30, 10.5, 140)):
            series.update_quote({**self.quote, "time": f"09:31:{second}", "current_price": price,
                                 "deals_vol": volume, "deals_amt": volume * 11})
        self.assertEqual(len(series.result.bars), 1)
        self.assertEqual((series.result.bars[-1].close, series.result.bars[-1].high), (10.5, 12))
        series.update_quote({**self.quote, "time": "09:32:01", "current_price": 12, "deals_vol": 150, "deals_amt": 1650})
        self.assertEqual(len(series.result.bars), 2)
        self.assertEqual(series.result.bars[-1].volume, 10)
        self.assertEqual(series.result.bars[-1].average, 11)
        self.assertFalse(series.needs_download(self.now + timedelta(minutes=10)))

    def test_missing_minute_requests_repair_and_does_not_invent_its_volume(self):
        series = HistorySeries(A_SHARE, "five_day")
        series.set_history(BarResult((self.minute,), "sina"), "2026-09-30", self.now)
        series.update_quote({**self.quote, "time": "09:33:01", "deals_vol": 180})
        self.assertTrue(series.needs_download(self.now + timedelta(minutes=2)))
        self.assertEqual(series.result.bars[-1].volume, 0)
        series.update_quote({**self.quote, "time": "09:33:20", "deals_vol": 190})
        self.assertEqual(series.result.bars[-1].volume, 10)
        series.set_history(BarResult((self.minute,), "sina"), "2026-09-30", self.now + timedelta(minutes=2))
        self.assertFalse(series.needs_download(self.now + timedelta(minutes=3)))
        self.assertTrue(series.needs_download(self.now + timedelta(minutes=8)))
        missing = Bar("2026-09-30 09:32:00", 11, 11, 11, 11, 60, 660, 11)
        series.set_history(BarResult((self.minute, missing), "sina"), "2026-09-30", self.now + timedelta(minutes=8))
        self.assertFalse(series.repair_needed)
        self.assertEqual(series.result.bars[-1].volume, 30)

    def test_lunch_break_is_not_missing_data_but_an_actual_hole_is(self):
        start = datetime(2026, 9, 30, 9, 31)
        bars = tuple(Bar((start + timedelta(minutes=i)).isoformat(" "), 10, 10, 10, 10) for i in range(120))
        current = datetime(2026, 9, 30, 13)
        self.assertFalse(minute_gaps(bars, A_SHARE, current))
        self.assertTrue(minute_gaps(bars[:-2], A_SHARE, current))

    def test_daily_candle_updates_cumulative_ohlc_and_refreshes_next_day(self):
        series = HistorySeries(A_SHARE, "daily")
        bars = (Bar("2026-09-29", 9, 10, 8, 9, 100), Bar("2026-09-30", 10, 11, 9, 10, 100))
        series.set_history(BarResult(bars, "sina"), "2026-09-30", self.now)
        series.update_quote(self.quote)
        self.assertEqual(series.result.bars[-1], Bar("2026-09-30", 10, 12, 9, 11, 120, 1320))
        self.assertFalse(series.needs_download(self.now + timedelta(hours=6)))
        series.update_quote({**self.quote, "date": "2026-10-01", "current_price": 12, "opening_price": 11})
        self.assertEqual(len(series.result.bars), 3)
        self.assertEqual(series.result.bars[-1].time, "2026-10-01")
        self.assertTrue(series.needs_download(self.now + timedelta(days=1)))

    def test_unknown_invalid_and_out_of_order_quotes_preserve_data(self):
        series = HistorySeries(A_SHARE, "intraday")
        series.set_history(BarResult((self.minute,), "sina"), "2026-09-30", self.now)
        series.update_quote(self.quote)
        bars = series.result.bars
        for change in ({"date": "", "time": ""}, {"current_price": 0}, {"opening_price": 0},
                       {"current_price": float("nan")}, {"time": "09:30:59"}):
            series.update_quote({**self.quote, **change})
            self.assertEqual(series.result.bars, bars)

    def test_us_clock_and_futures_night_use_the_exchange_trading_day(self):
        us = {"market": "us", "code": "aapl", "type": "美"}
        quote = {**self.quote, "date": "2026-07-02", "time": "03:59:30"}
        self.assertEqual(merge_quote((), us, "daily", quote)[-1].time, "2026-07-01")
        self.assertEqual(merge_quote((), us, "five_day", quote)[-1].time, "2026-07-01 15:59:00")
        future = {"market": "", "code": "au0", "type": "期"}
        quote = {**self.quote, "date": "2026-09-30", "time": "21:01:00"}
        self.assertEqual(merge_quote((), future, "daily", quote)[-1].time, "2026-10-01")
        self.assertFalse(minute_gaps((), future, datetime(2026, 9, 30, 21, 1)))
        self.assertEqual(history_day(future, datetime(2026, 9, 30, 13, 1, tzinfo=timezone.utc)), "2026-10-01")
        midnight = (Bar("2026-09-30 23:59:00", 10, 10, 10, 10),
                    Bar("2026-10-01 00:00:00", 10, 10, 10, 10))
        self.assertFalse(minute_gaps(midnight, future, datetime(2026, 10, 1, 0, 1)))
        self.assertTrue(minute_gaps(midnight[:1], future, datetime(2026, 10, 1, 0, 1)))

    def test_live_and_eastmoney_history_volumes_use_the_same_unit(self):
        payload = {"data": {"trends": ["2026-09-30 09:31,10,10,11,9,1,1000,10"]}}
        bars = parse_eastmoney(payload, instrument=A_SHARE)
        self.assertEqual(bars[0].volume, 100)
        self.assertEqual(merge_quote(bars, A_SHARE, "intraday", self.quote)[-1].volume, 120)

    def test_failed_repairs_keep_minutes_already_observed_live(self):
        series = HistorySeries(A_SHARE, "five_day")
        series.set_history(BarResult((self.minute,), "sina"), "2026-09-30", self.now)
        series.update_quote({**self.quote, "time": "09:32:10"})
        series.update_quote({**self.quote, "time": "09:34:10", "current_price": 12})
        series.set_history(BarResult((self.minute,), "sina", "更新失败", cached=True, stale=True),
                           "2026-09-30", self.now)
        self.assertEqual([bar.time[11:16] for bar in series.result.bars], ["09:31", "09:32", "09:34"])
        self.assertEqual(series.result.bars[-1].close, 12)
        series.set_history(BarResult(message="处理失败"), "2026-09-30", self.now)
        self.assertEqual(len(series.result.bars), 3)
        self.assertTrue(series.result.stale)
        # A real provider change resets history rather than mixing sources.
        series.set_history(BarResult((self.minute,), "eastmoney"), "2026-09-30", self.now)
        self.assertNotIn("09:32", [bar.time[11:16] for bar in series.result.bars])

    def test_repair_keeps_the_current_minute_extrema_observed_during_download(self):
        series = HistorySeries(A_SHARE, "intraday")
        series.set_history(BarResult((self.minute,), "sina"), "2026-09-30", self.now)
        series.update_quote({**self.quote, "current_price": 12})
        series.update_quote({**self.quote, "time": "09:31:20", "current_price": 10})
        series.set_history(BarResult((self.minute,), "sina"), "2026-09-30", self.now)
        self.assertEqual((series.result.bars[-1].high, series.result.bars[-1].close), (12, 10))

    def test_empty_history_retries_after_backoff_even_if_a_live_candle_exists(self):
        series = HistorySeries(A_SHARE, "daily")
        series.update_quote(self.quote)
        series.set_history(BarResult(message="暂无可用历史数据"), "2026-09-30", self.now)
        self.assertTrue(series.result.bars)
        self.assertFalse(series.needs_download(self.now + timedelta(minutes=4)))
        self.assertTrue(series.needs_download(self.now + timedelta(minutes=6)))
