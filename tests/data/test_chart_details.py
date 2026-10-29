"""The chart's volume ratio uses real completed daily volumes and quote time."""

from datetime import datetime, timedelta, timezone
from contextlib import closing
import tempfile
import unittest
from unittest.mock import Mock, patch

import requests

from stockwidget.data.bars import Bar, quote_moving_average, quote_volume_ratio
from stockwidget.data.bar_cache import BarCache
from stockwidget.data.quotes import parse_sina_fundamentals, parse_sina_trade_volumes, request_quote_details, _parse_sina_us
from stockwidget.core.quote_presentation import quote_detail_values


class QuoteDetailsDataTests(unittest.TestCase):
    instrument = {"market": "sh", "code": "600000", "type": "沪"}

    @staticmethod
    def payload(label="sh600000_i", **fields):
        parts = [""] * 42
        for index, value in {0: "A", 2: "1.52", 3: "1.5552", 4: ".9293", 5: "22.63",
                             7: "100", 8: "90", 9: "80", 12: ".02", 13: ".03", 29: "1/1",
                             39: "20260930|3000000", **{int(k): v for k, v in fields.items()}}.items():
            parts[index] = value
        return 'var hq_str_' + label + '="' + ','.join(parts) + '";'

    def test_domestic_share_counts_earnings_and_currency_scope(self):
        for market in ("sh", "sz", "bj"):
            with self.subTest(market=market):
                result = parse_sina_fundamentals(self.payload(market + "600000_i"), {**self.instrument, "market": market})
                self.assertEqual(result["total_shares"], 1_000_000)
                self.assertEqual(result["float_shares"], 800_000)
                self.assertEqual(result["annual_profit"], 2_000_000)
                self.assertEqual(result["ttm_profit"], 3_000_000)
                self.assertEqual(result["annualized_profit"], 4_000_000)
                self.assertNotIn("eps", result)
        result = parse_sina_fundamentals(self.payload(**{"29": "2/1", "5": "NaN"}), self.instrument)
        self.assertIsNone(result["book_value_per_share"])
        self.assertEqual(parse_sina_fundamentals(self.payload(**{"0": "B"}), self.instrument), {})
        self.assertEqual(parse_sina_fundamentals(self.payload(), {**self.instrument, "type": "基"}), {})

    def test_optional_endpoint_failures_do_not_mask_other_details(self):
        response = Mock(text="var trade_INVOL_OUTVOL=[200,300];")
        with patch("stockwidget.data.quotes.requests.get", side_effect=[requests.Timeout(), response]):
            fundamental, activity = request_quote_details(self.instrument)
        self.assertEqual(fundamental, {})
        self.assertEqual(activity, {"outside_volume": 200, "inside_volume": 300})
        for cached in ({"annual_profit": 50017000000, "total_shares": 10000}, {}):
            with self.subTest(cached=cached):
                with patch("stockwidget.data.quotes.requests.get", side_effect=requests.Timeout()) as request:
                    fundamental, activity = request_quote_details(self.instrument, cached)
                self.assertEqual(fundamental, cached)
                self.assertFalse(activity)
                self.assertEqual(request.call_count, 1)  # Includes negative-cache backoff.

    def test_trade_payload_never_executes_javascript_or_converts_absence_to_zero(self):
        self.assertEqual(parse_sina_trade_volumes("Access denied"), {})
        self.assertEqual(parse_sina_trade_volumes("var trade_INVOL_OUTVOL=[NaN,0];"), {})
        self.assertEqual(parse_sina_trade_volumes("var trade_INVOL_OUTVOL=[0,0];"),
                         {"outside_volume": 0, "inside_volume": 0})

    def test_live_price_recomputes_valuations_and_turnover_with_explicit_missingness(self):
        quote = {"current_price": 12, "prev_close": 10, "opening_price": 11, "high_price": 13,
                 "low_price": 9, "deals_vol": 40000}
        details = {"annual_profit": 2_000_000, "ttm_profit": 3_000_000, "annualized_profit": 4_000_000, "book_value_per_share": 4,
                   "total_shares": 1_000_000, "float_shares": 800_000, "outside_volume": 0, "inside_volume": 0}
        result = quote_detail_values(quote, self.instrument, details)
        for key, expected in {"pe_static": 6, "pe_ttm": 4, "pb": 3, "market_cap": 12_000_000,
                              "float_market_cap": 9_600_000, "turnover_rate": 5, "amplitude": 40, "pe_dynamic": 3}.items():
            self.assertEqual(result[key], expected)
        self.assertIsNone(result["inside_volume"])
        self.assertIsNone(result["outside_volume"])
        for key in ("annual_profit", "ttm_profit", "annualized_profit", "book_value_per_share"):
            details[key] = -1
        result = quote_detail_values(quote, self.instrument, details)
        self.assertTrue(all(result[key] is None for key in ("pe_static", "pe_dynamic", "pe_ttm", "pb")))
        result = quote_detail_values(quote, {**self.instrument, "type": "指"}, details)
        self.assertIsNone(result["total_shares"])
        self.assertEqual(result["amplitude"], 40)

    def test_us_ttm_earnings_shares_and_market_value_reuse_basic_quote(self):
        parts = ["0"] * 31
        for index, value in {0: "APPLE", 1: "30", 3: "2026-10-29 05:00:00", 5: "29", 6: "31", 7: "28",
                             12: "3000000", 13: "3", 19: "100000", 26: "29"}.items():
            parts[index] = value
        quote = _parse_sina_us(parts)
        result = quote_detail_values(quote, {"market": "us", "type": "美"})
        self.assertEqual(result["pe_ttm"], 10)
        self.assertEqual(result["market_cap"], 3_000_000)
        self.assertEqual(result["total_shares"], 100000)
        self.assertIsNone(result["pe_static"])

    def test_capital_changes_do_not_reuse_weighted_eps_as_current_earnings(self):
        # User-reported cases, independently checked against Tencent snapshots.
        for market, code, price, total, annual, ttm, report, expected in (
                ("sh", "600089", 17.79, "502024.8734", "59.5429", "53.2274", "20260630|2552310324.8", (15.00, 17.50, 16.78)),
                ("sz", "300750", 229.07, "462724.9686", "722.0128", "876.9724", "20260930|64530063000", (14.68, 12.32, 12.09))):
            with self.subTest(code=code):
                instrument = {"market": market, "code": code, "type": "沪"}
                payload = self.payload(market + code + "_i", **{"7": total, "12": annual, "13": ttm, "39": report})
                details = parse_sina_fundamentals(payload, instrument)
                values = quote_detail_values({"current_price": price}, instrument, details)
                self.assertEqual(tuple(round(values[key], 2) for key in ("pe_static", "pe_dynamic", "pe_ttm")), expected)

    def test_dynamic_profit_fallback_annualizes_average_quarter_instead_of_ytd_twice(self):
        payload = self.payload(**{"39": "", "25": "20260930|1000000"})
        self.assertEqual(parse_sina_fundamentals(payload, self.instrument)["annualized_profit"], 4_000_000)


class FundamentalsCacheTests(unittest.TestCase):
    instrument = {"market": "sh", "code": "600089", "type": "沪"}
    values = {"annual_profit": 5954290000, "ttm_profit": 5322740000, "report_date": "20260630"}

    def setUp(self):
        self.root = self.enterContext(tempfile.TemporaryDirectory())
        self.now = datetime(2026, 10, 29, 2, tzinfo=timezone.utc)
        self.cache = BarCache(self.root, clock=lambda: self.now)

    def test_daily_cache_survives_restart_then_expires_at_market_day_boundary(self):
        self.assertTrue(self.cache.write_fundamentals(self.instrument, self.values, self.now.timestamp(), generation=0))
        restarted = BarCache(self.root, clock=lambda: self.now)
        self.assertEqual(restarted.read_fundamentals(self.instrument), (self.values, self.now.timestamp()))
        self.now += timedelta(hours=5)
        self.assertEqual(restarted.read_fundamentals(self.instrument)[0], self.values)
        self.now += timedelta(days=1)
        self.assertIsNone(restarted.read_fundamentals(self.instrument))

    def test_failed_fetch_has_short_backoff_and_clear_rejects_old_writes(self):
        old_generation = self.cache.generation
        self.cache.write_fundamentals(self.instrument, {}, self.now.timestamp(), generation=old_generation)
        self.now += timedelta(seconds=599)
        self.assertEqual(self.cache.read_fundamentals(self.instrument)[0], {})
        self.now += timedelta(seconds=1)
        self.assertIsNone(self.cache.read_fundamentals(self.instrument))
        self.assertTrue(self.cache.clear())
        self.assertFalse(self.cache.write_fundamentals(self.instrument, self.values, self.now.timestamp(), generation=old_generation))
        self.assertFalse(self.cache.database.exists())

    def test_invalid_json_does_not_rebuild_or_delete_other_cache_data(self):
        self.cache.write_fundamentals(self.instrument, self.values, self.now.timestamp(), generation=0)
        with closing(self.cache._connect()) as db, db:
            db.execute("UPDATE fundamentals SET payload = '{invalid'")
            db.execute("INSERT INTO entries(cache_key,saved_at,stamp,active,day,source,message) VALUES ('sentinel',0,'',0,'','sina','')")
        self.assertIsNone(self.cache.read_fundamentals(self.instrument))
        with closing(self.cache._connect()) as db:
            self.assertEqual(db.execute("SELECT cache_key FROM entries").fetchone(), ("sentinel",))

    def test_us_daily_expiry_follows_exchange_date_instead_of_utc_midnight(self):
        instrument = {"market": "us", "code": "aapl"}
        self.now = datetime(2026, 10, 28, 23, 30, tzinfo=timezone.utc)
        self.cache.write_fundamentals(instrument, self.values, self.now.timestamp(), generation=0)
        self.now += timedelta(hours=1)
        self.assertIsNotNone(self.cache.read_fundamentals(instrument))
        self.now += timedelta(hours=4)
        self.assertIsNone(self.cache.read_fundamentals(instrument))


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
