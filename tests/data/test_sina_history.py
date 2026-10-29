"""Sina-only quote/history adapters: real formats, safe failures and partial windows."""

from datetime import datetime
import json
import unittest
from unittest.mock import Mock, patch

import requests

from stockwidget.data.bars import (
    Bar, decode_sina_daily, fetch_bars, _parse_us_minutes, quote_clock, request_sina_bars,
)
from stockwidget.data.quotes import fetch_quote_result, request_quote, _parse_sina_global
from stockwidget.core.watchlist import normalize_watchlist


# Two candles checked against Sina's xh5_S_KLC_D decoder (format 3466).
DAILY = 'var KLC_K2_TEST="K2/hQfkBFUAZAAAAFAKog9LDAAAMA";'
SHARE = {"market": "sh", "code": "600000", "type": "沪"}


class SinaDailyCodecTests(unittest.TestCase):
    def test_daily_delta_prices_volume_and_dates_match_reference(self):
        bars = decode_sina_daily(DAILY + " /* optional transport signature */")
        self.assertEqual([bar.time for bar in bars], ["1990-12-20", "1990-12-21"])
        for bar, expected in zip(bars, ((10, 10.5, 9.8, 10.1, 100), (10.2, 10.4, 10.1, 10.15, 150))):
            for actual, value in zip((bar.open, bar.high, bar.low, bar.close, bar.volume), expected):
                self.assertAlmostEqual(actual, value)

    def test_malformed_truncated_or_executable_responses_fail_closed(self):
        for text in ("<html>denied</html>", 'var x="AAAA";', DAILY[:-8],
                     DAILY + "alert(1)", 'var x="K2/!!";', 'var x="K2/";'):
            with self.subTest(text=text):
                with self.assertRaises((ValueError, IndexError)):
                    decode_sina_daily(text)


class SinaHistoryAdapterTests(unittest.TestCase):
    def test_hk_us_stock_and_index_daily_use_sina_and_local_decode(self):
        for instrument, suffix in (
            ({"market":"hk", "code":"00700"}, "/00700/klc2_kl.js"),
            ({"market":"hk", "code":"hsi", "type":"指"}, "/HSI/klc2_kl.js"),
            ({"market":"us", "code":"aapl"}, "/staticdata/us/AAPL"),
            ({"market":"us", "code":"brk_b"}, "/staticdata/us/BRK.B"),
            ({"market":"us", "code":"ixic", "type":"指"}, "/staticdata/us/.IXIC"),
        ):
            with self.subTest(instrument=instrument), patch(
                "stockwidget.data.bars.requests.get", return_value=Mock(text=DAILY)
            ) as get:
                self.assertEqual(len(request_sina_bars(instrument, "daily")), 2)
                self.assertTrue(get.call_args.args[0].endswith(suffix))
                self.assertNotIn("verify", get.call_args.kwargs)
                self.assertNotIn("proxies", get.call_args.kwargs)

    def test_global_daily_json_uses_original_index_identity(self):
        payload = {"code":0,"result":{"data":[{"d":"2026-10-26","o":"10","h":"12","l":"9","c":"11","v":"0"}]}}
        with patch("stockwidget.data.bars.requests.get", return_value=Mock(text=json.dumps(payload))) as get:
            result = fetch_bars({"market":"gb","code":"nky","type":"指"}, "daily")
        self.assertEqual(result.bars[0].time, "2026-10-26")
        self.assertEqual(get.call_args.kwargs["params"]["symbol"], "NKY")
        self.assertEqual(get.call_args.args[0], "https://gi.finance.sina.com.cn/hq/daily")

    def test_us_minute_dates_are_exchange_dates_and_volume_is_not_cumulative(self):
        text = '/* transport */=(" 2026-10-23,09:30:00,100,10.5,11,10;09:31:00,50,10.8,12 2026-10-26,09:30:00,200,13,13,12");'
        bars = _parse_us_minutes(text)
        self.assertEqual([bar.time for bar in bars], ["2026-10-23 09:30:00","2026-10-23 09:31:00","2026-10-26 09:30:00"])
        self.assertEqual([bar.volume for bar in bars], [100,50,200])
        self.assertEqual([bar.close for bar in bars], [11,12,13])
        self.assertEqual([bar.average for bar in bars], [10.5,10.8,13])
        with patch("stockwidget.data.bars.requests.get", return_value=Mock(text=text)) as get:
            result = fetch_bars({"market":"us","code":"ixic","type":"指"}, "five_day")
        self.assertIn("2 个交易日", result.message)
        self.assertEqual(get.call_args.kwargs["params"], {"symbol":".IXIC","day":5})

    def test_hk_minutes_use_quote_trading_date_and_report_one_available_session(self):
        payload = {"result":{"status":{"code":0},"data":[
            {"m":"09:30:00","p":"390.2","v":"385012","a":"149174533","avg_p":"387.454"},
            {"m":"09:31:00","p":"389.2","v":"127660","a":"49789444","avg_p":"388.1"},
        ]}}
        with patch("stockwidget.data.bars.requests.get", return_value=Mock(text="sw("+json.dumps(payload)+")")), \
             patch("stockwidget.data.bars.request_quote", return_value={"hk00700":{"date":"2026/10/23"}}):
            result = fetch_bars({"market":"hk","code":"00700"}, "five_day")
        self.assertEqual(result.bars[0].time, "2026-10-23 09:30:00")
        self.assertEqual(result.bars[0].volume, 385012)
        self.assertEqual(result.bars[0].amount, 149174533)
        self.assertEqual(result.bars[0].average, 387.454)
        self.assertIn("1 个交易日", result.message)

    def test_partial_windows_are_preserved_and_errors_are_safe(self):
        bars = (Bar("2026-10-26 10:00:00",10,10,10,10),)
        with patch("stockwidget.data.bars.request_sina_bars", return_value=bars) as fetch:
            result = fetch_bars(SHARE, "five_day")
        self.assertEqual(result.bars, bars)
        self.assertIn("1 个交易日", result.message)
        fetch.assert_called_once_with(SHARE,"five_day")
        for error, message in ((requests.Timeout("https://secret.invalid"),"超时"),
                               (ValueError("private payload"),"格式不可用")):
            with self.subTest(error=type(error).__name__), patch("stockwidget.data.bars.request_sina_bars",side_effect=error):
                result=fetch_bars(SHARE,"daily")
            self.assertFalse(result.bars)
            self.assertIn(message,result.message)
            self.assertNotIn("secret",result.message)
            self.assertNotIn("private",result.message)

    def test_global_minutes_are_explicitly_unavailable_without_guessing(self):
        with patch("stockwidget.data.bars.requests.get") as get:
            result=fetch_bars({"market":"gb","code":"nky","type":"指"},"five_day")
        get.assert_not_called()
        self.assertFalse(result.bars)
        self.assertIn("分钟历史",result.message)


class SinaQuoteTransportTests(unittest.TestCase):
    def test_global_date_placeholder_uses_actual_time_and_exchange_day(self):
        raw = ["富时100指数","10666.76","-100.57","-.93","9/26/2025","2025-09-26",
               "2026-10-28","00:15:00","10767.81","10767.33","10767.81","10649.65","0"]
        quote = _parse_sina_global(raw)
        self.assertEqual(quote["time"],"00:15:00")
        info={"market":"gb","code":"ukx","type":"指","timezone":"Europe/London"}
        enriched=normalize_watchlist({"gbukx":{"checked":True}},{"gbukx":info})["gbukx"]
        local = quote_clock(quote,enriched)
        self.assertEqual(local,datetime(2026,10,27,16,15))
        self.assertIsNone(quote_clock(quote,{"market":"gb","code":"ukx"}))
        raw[5] = "14:12:00"
        self.assertEqual(_parse_sina_global(raw)["time"],"14:12:00")

    @staticmethod
    def response(code="sh600000", name="浦发银行"):
        parts=[name,"10","9","11","12","8"]+["0"]*26
        parts[30:32]=["2026-10-27","15:00:00"]
        return Mock(text=f'var hq_str_{code}="'+','.join(parts)+'";')

    def test_quotes_preserve_http_and_decode_errors_without_raw_details(self):
        for error in (requests.ReadTimeout("secret url"), requests.ConnectionError("secret"), ValueError("private")):
            with self.subTest(error=type(error).__name__),patch("stockwidget.data.quotes.requests.get",side_effect=error):
                ok,data,message=fetch_quote_result({"sh600000":SHARE})
            self.assertFalse(ok)
            self.assertIsNone(data)
            self.assertTrue(message.startswith("新浪："))
            self.assertNotIn("secret",message)
            self.assertNotIn("private",message)

    def test_same_raw_codes_are_mapped_by_explicit_markets_and_http_errors_raise(self):
        response=self.response("sh000001","上证指数")
        response.text+='\n'+self.response("sz000001","平安银行").text
        with patch("stockwidget.data.quotes.requests.get",return_value=response):
            data=request_quote({"sh000001":{"market":"sh","code":"000001","type":"指"},
                                "sz000001":{"market":"sz","code":"000001","type":"深"}})
        self.assertEqual(data["sh000001"]["name"],"上证指数")
        self.assertEqual(data["sz000001"]["name"],"平安银行")
        self.assertEqual(data["sh000001"]["time"],"15:00:00")
        response.close.assert_called_once()
        denied=Mock()
        denied.raise_for_status.side_effect=requests.HTTPError("secret")
        with patch("stockwidget.data.quotes.requests.get",return_value=denied):
            with self.assertRaises(requests.HTTPError): request_quote({"sh600000":SHARE})

    def test_large_watchlists_are_batched_and_html_is_not_successful_empty_data(self):
        instruments={f"sh{n:06}":{"market":"sh","code":f"{n:06}"} for n in range(121)}
        def get(url,**kwargs):
            return Mock(text='\n'.join(self.response(label).text for label in url.split('list=')[1].split(',')))
        with patch("stockwidget.data.quotes.requests.get",side_effect=get) as fetch:
            self.assertEqual(len(request_quote(instruments)),121)
        self.assertEqual(fetch.call_count,2)
        with patch("stockwidget.data.quotes.requests.get",return_value=Mock(text="<html>denied</html>")):
            self.assertFalse(fetch_quote_result({"sh600000":SHARE})[0])
        with patch("stockwidget.data.quotes.requests.get",return_value=Mock(text='var hq_str_sh600000="";')):
            self.assertEqual(fetch_quote_result({"sh600000":SHARE}), (False,None,"新浪：未返回有效行情"))
