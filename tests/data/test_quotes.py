"""新浪 / 东财行情解析和安全的网络错误分类。"""

from http.client import RemoteDisconnected
from unittest.mock import Mock, patch
import socket
import ssl
import unittest

from urllib3.exceptions import MaxRetryError, NameResolutionError, ProtocolError, ProxyError
import requests

from stockwidget.data import quotes
from stockwidget.data.network_errors import request_error_message
from stockwidget.data.quotes import _em_secid, _sina_code


class ExplicitMarketMetadataTests(unittest.TestCase):
    def test_shanghai_futures_use_futures_provider_symbols(self):
        for code, secid in (("au0", "113.aum"), ("au2610", "113.au2610"),
                            ("sc0", "142.scm"), ("sc2610", "142.sc2610")):
            with self.subTest(code=code):
                future = {"market": "sh", "code": code, "type": "期"}
                self.assertEqual(_sina_code(future), "nf_" + code.upper())
                self.assertEqual(_em_secid(future), secid)

    def test_shanghai_futures_sina_parser_and_volume(self):
        parts = ["黄金连续", "150000"] + ["0"] * 26
        parts[7], parts[10], parts[14] = "800", "790", "1234"
        response = Mock(text='var hq_str_nf_AU0="' + ','.join(parts) + '";')
        with patch.object(quotes.requests, "get", return_value=response):
            data = quotes.request_sina({"au0": {"market": "sh", "code": "au0", "type": "期"}})
        self.assertEqual(data["au0"]["current_price"], 800)
        self.assertEqual(data["au0"]["prev_close"], 790)
        self.assertEqual(data["au0"]["deals_vol"], 1234)

    def test_shanghai_futures_eastmoney_volume_stays_in_contracts(self):
        response = Mock()
        response.json.return_value = {"data": {"diff": [
            {"f12": "aum", "f13": 113, "f14": "黄金连续", "f5": 1234},
        ]}}
        with patch.object(quotes.requests, "get", return_value=response):
            _, data = quotes.request_eastmoney({"au0": {"market": "sh", "code": "au0", "type": "期"}})
        self.assertEqual(data["au0"]["deals_vol"], 1234)

    def test_same_raw_code_keeps_markets_distinct(self):
        shanghai = {"market": "sh", "code": "000001"}
        shenzhen = {"market": "sz", "code": "000001"}

        self.assertEqual(_sina_code(shanghai), "sh000001")
        self.assertEqual(_sina_code(shenzhen), "sz000001")
        self.assertEqual(_em_secid(shanghai), "1.000001")
        self.assertEqual(_em_secid(shenzhen), "0.000001")

    def test_futures_uses_empty_market_and_raw_code(self):
        future = {"market": "", "code": "au0"}
        self.assertEqual(_sina_code(future), "nf_AU0")
        self.assertEqual(_em_secid(future), "113.aum")


class SinaVolumeUnitTests(unittest.TestCase):
    @staticmethod
    def _index_parts(volume: str) -> list[str]:
        parts = ["0"] * 32
        parts[0] = "指数"
        parts[1:6] = ["10", "9", "11", "12", "8"]
        parts[8] = volume
        parts[9] = "5000"
        return parts

    def test_shanghai_index_lots_are_normalized_to_shares(self):
        entry = quotes._parse_sina_a(
            self._index_parts("1234"), is_index=True, market="sh"
        )
        self.assertEqual(entry["deals_vol"], 123400)

    def test_shenzhen_index_is_already_in_shares(self):
        entry = quotes._parse_sina_a(
            self._index_parts("1234"), is_index=True, market="sz"
        )
        self.assertEqual(entry["deals_vol"], 1234)

    def test_hong_kong_index_is_already_in_shares(self):
        parts = ["0"] * 19
        parts[0:7] = ["HSI", "恒生指数", "10", "9", "12", "8", "11"]
        parts[11] = "5000"
        parts[12] = "1234"
        entry = quotes._parse_sina_hk(parts, is_index=True)
        self.assertEqual(entry["deals_vol"], 1234)

    def test_us_index_is_already_in_shares(self):
        parts = ["0"] * 31
        parts[0] = "纳斯达克"
        parts[1] = "11"
        parts[3] = "2026-08-29 05:30:00"
        parts[5:8] = ["10", "12", "8"]
        parts[10] = "1234"
        parts[26] = "9"
        entry = quotes._parse_sina_us(parts, is_index=True)
        self.assertEqual(entry["deals_vol"], 1234)

    def test_futures_volume_is_contract_count(self):
        parts = ["0"] * 18
        parts[0] = "黄金连续"
        parts[1] = "150000"
        parts[2:8] = ["10", "12", "8", "10", "11", "11"]
        parts[10] = "9"
        parts[14] = "1234"
        entry = quotes._parse_sina_futures(parts)
        self.assertEqual(entry["deals_vol"], 1234)


class SinaRequestFailureTests(unittest.TestCase):
    def test_http_error_is_not_parsed_as_an_empty_watchlist(self):
        response = requests.Response()
        response.status_code = 403
        response.url = "https://hq.sinajs.cn/list=sh600000"
        response._content = b"Access denied"
        with patch.object(quotes.requests, "get", return_value=response):
            with self.assertRaises(requests.HTTPError) as caught:
                quotes.request_sina({"sh600000": {"market": "sh", "code": "600000"}})
        self.assertIs(caught.exception.response, response)


class EastmoneyQuoteTests(unittest.TestCase):
    @staticmethod
    def _response(payload):
        response = Mock()
        response.json.return_value = payload
        return response

    def test_quote_timestamp_is_requested_and_normalized_for_each_market(self):
        instruments = {"sh600000": {"market": "sh", "code": "600000"},
                       "hk00700": {"market": "hk", "code": "00700"},
                       "usaapl": {"market": "us", "code": "aapl"},
                       "au0": {"market": "sh", "code": "au0", "type": "期"}}
        response = self._response({"data": {"diff": [
            {"f12": code, "f13": market, "f124": 1790753121}
            for code, market in (("600000", 1), ("00700", 116), ("AAPL", 105), ("aum", 113))
        ]}})
        with patch.object(quotes.requests, "get", return_value=response) as request:
            _, data = quotes.request_eastmoney(instruments)
        self.assertIn("f124", request.call_args.kwargs["params"]["fields"].split(","))
        for key in instruments:
            self.assertEqual(data[key]["timestamp"], 1790753121)
            self.assertEqual(data[key]["date"], "2026-09-30")
            self.assertEqual(data[key]["time"], "15:25:21")

    def test_missing_or_invalid_quote_time_is_not_replaced_by_request_time(self):
        for invalid in (None, "-", 0, "bad", 1e99):
            with self.subTest(invalid=invalid):
                response = self._response({"data": {"diff": [
                    {"f12": "600000", "f13": 1, "f124": invalid}
                ]}})
                with patch.object(quotes.requests, "get", return_value=response):
                    _, data = quotes.request_eastmoney({"sh600000": {"market": "sh", "code": "600000"}})
                self.assertIsNone(data["sh600000"]["timestamp"])
                self.assertEqual((data["sh600000"]["date"], data["sh600000"]["time"]), ("", ""))

    def test_stable_host_and_preopen_placeholders(self):
        response = self._response(
            {
                "data": {
                    "diff": [
                        {
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
                        }
                    ]
                }
            }
        )
        instruments = {
            "sh600000": {"market": "sh", "code": "600000", "type": "沪"}
        }
        with patch.object(quotes.requests, "get", return_value=response) as request:
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
        self.assertNotIn("f31", request.call_args.kwargs["params"]["fields"])
        self.assertEqual(request.call_count, 1)

    def test_eastmoney_does_not_call_sina(self):
        response = self._response(
            {
                "data": {
                    "diff": [
                        {
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
                        }
                    ]
                }
            }
        )
        instruments = {
            "sh600000": {"market": "sh", "code": "600000", "type": "沪"}
        }
        with (
            patch.object(quotes.requests, "get", return_value=response) as request,
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
        self.assertEqual(request.call_count, 1)
        request_sina.assert_not_called()

    def test_transient_failures_retry_then_recover_on_an_eastmoney_node(self):
        disconnected = RemoteDisconnected("closed without response")
        response = self._response({"data": {"diff": [
            {"f12": "600000", "f13": 1, "f14": "浦发银行", "f2": 9.05, "f5": 1724},
        ]}})
        instruments = {"sh600000": {"market": "sh", "code": "600000", "type": "沪"}}
        failures = (requests.ConnectionError(ProtocolError("aborted", disconnected)),
                    requests.exceptions.ProxyError(ProtocolError("aborted", disconnected)),
                    requests.ReadTimeout(), requests.exceptions.ChunkedEncodingError())
        for error in failures:
            with (
                self.subTest(error=type(error).__name__),
                patch.object(quotes.requests, "get", side_effect=[error, error, response]) as get,
                patch("stockwidget.data.quotes.time.sleep") as sleep,
                patch.object(quotes, "request_sina") as sina,
            ):
                ok, data, reason = quotes.fetch_quote_result(instruments, "eastmoney")
                self.assertTrue(ok)
                self.assertIsNone(reason)
                self.assertEqual(data["sh600000"]["deals_vol"], 172400)
                self.assertEqual([call.args[0] for call in get.call_args_list], [
                    quotes._EM_QUOTE_URL, quotes._EM_QUOTE_URL,
                    "https://push2.eastmoney.com/api/qt/ulist.np/get",
                ])
                self.assertEqual([call.args[0] for call in sleep.call_args_list], [.25, .5])
                for call in get.call_args_list:
                    self.assertEqual(call.kwargs["params"]["secids"], "1.600000")
                    self.assertNotIn("proxies", call.kwargs)
                    self.assertNotIn("verify", call.kwargs)
                sina.assert_not_called()

    def test_retry_exhaustion_is_bounded_and_preserves_safe_error(self):
        with (
            patch.object(quotes.requests, "get", side_effect=requests.ConnectionError(
                RemoteDisconnected("https://user:secret@example.com"))) as get,
            patch("stockwidget.data.quotes.time.sleep") as sleep,
            patch.object(quotes, "request_sina") as sina,
        ):
            result = quotes.fetch_quote_result({"sh600000": {"market": "sh", "code": "600000"}}, "eastmoney")
        self.assertEqual(result, (False, None, "东财：远端关闭连接"))
        self.assertEqual(get.call_count, 3)
        self.assertEqual(sleep.call_count, 2)
        sina.assert_not_called()

    def test_gateway_errors_retry_but_denial_rate_limit_and_ssl_errors_stop(self):
        instruments = {"sh600000": {"market": "sh", "code": "600000"}}
        for status in (500, 502, 503, 504, 403, 429):
            with self.subTest(status=status):
                response = requests.Response()
                response.status_code = status
                response._content = b"unavailable"
                response._content_consumed = True
                good = self._response({"data": {"diff": [{"f12": "600000", "f13": 1, "f2": 9.05}]}})
                with (patch.object(quotes.requests, "get", side_effect=[response, good]) as get,
                      patch("stockwidget.data.quotes.time.sleep")):
                    ok, data, reason = quotes.fetch_quote_result(instruments, "eastmoney")
                self.assertEqual(ok, status in (500, 502, 503, 504))
                self.assertEqual(get.call_count, 2 if ok else 1)
        for error in (requests.exceptions.SSLError(), requests.exceptions.InvalidURL(),
                      requests.exceptions.JSONDecodeError("bad", "", 0)):
            with (
                self.subTest(error=type(error).__name__),
                patch.object(quotes.requests, "get", side_effect=error) as get,
                patch("stockwidget.data.quotes.time.sleep") as sleep,
            ):
                self.assertFalse(quotes.fetch_quote_result(instruments, "eastmoney")[0])
                get.assert_called_once()
                sleep.assert_not_called()


class NetworkErrorTests(unittest.TestCase):
    def test_specific_request_subclasses_are_not_hidden_by_their_base_classes(self):
        cases = (
            (requests.exceptions.ProxyError(), "代理请求失败"),
            (requests.exceptions.SSLError(), "安全连接失败（SSL）"),
            (requests.exceptions.SSLError(ssl.SSLCertVerificationError()), "证书验证失败"),
            (requests.ConnectTimeout(), "连接超时"),
            (requests.ReadTimeout(), "读取超时"),
            (requests.Timeout(), "请求超时"),
            (requests.ConnectionError(), "连接失败或中断"),
            (requests.exceptions.JSONDecodeError("bad", "", 0), "响应不是有效 JSON"),
            (requests.exceptions.ChunkedEncodingError(), "响应传输中断"),
            (requests.exceptions.ContentDecodingError(), "响应解码失败"),
            (requests.TooManyRedirects(), "重定向次数过多"),
            (requests.exceptions.InvalidURL(), "请求地址无效"),
            (requests.exceptions.MissingSchema(), "请求地址无效"),
            (requests.exceptions.InvalidSchema(), "请求地址无效"),
            (requests.exceptions.URLRequired(), "请求地址无效"),
            (requests.exceptions.InvalidHeader(), "请求头无效"),
            (requests.RequestException(), "网络请求失败"),
        )
        for error, expected in cases:
            with self.subTest(error=type(error).__name__):
                self.assertEqual(request_error_message(error), expected)

    def test_http_errors_keep_status_even_when_response_is_false(self):
        for status, reason in ((401, "访问未授权"), (403, "访问被拒绝"), (404, "接口不存在"),
                               (429, "请求过于频繁"), (500, "服务端内部错误"), (502, "网关错误"),
                               (503, "服务暂不可用"), (504, "网关超时"), (418, "请求失败")):
            with self.subTest(status=status):
                response = requests.Response()
                response.status_code = status
                self.assertFalse(response)
                error = requests.HTTPError(response=response)
                self.assertEqual(request_error_message(error), f"HTTP {status}：{reason}")
        self.assertEqual(request_error_message(requests.HTTPError()), "HTTP 请求失败")

    def test_real_urllib3_wrappers_preserve_disconnect_and_dns_reasons(self):
        disconnected = RemoteDisconnected("Remote end closed connection without response")
        connection = requests.ConnectionError(ProtocolError("Connection aborted.", disconnected))
        self.assertEqual(request_error_message(connection), "远端关闭连接")
        proxy = requests.exceptions.ProxyError(MaxRetryError(
            None, "/", ProxyError("Unable to connect to proxy", disconnected)
        ))
        self.assertEqual(request_error_message(proxy), "代理请求失败（远端关闭连接）")
        dns = requests.ConnectionError(MaxRetryError(
            None, "/", NameResolutionError("example.com", None, socket.gaierror())
        ))
        self.assertEqual(request_error_message(dns), "域名解析失败")
        for cause, expected in ((ConnectionRefusedError(), "连接被拒绝"),
                                (ConnectionResetError(), "连接中断"),
                                (ConnectionAbortedError(), "连接中断")):
            self.assertEqual(request_error_message(requests.ConnectionError(cause)), expected)

    def test_raw_error_urls_credentials_and_markup_are_never_shown(self):
        raw = "https://user:secret@example.com/?token=private <script>bad</script>"
        for error in (requests.exceptions.ProxyError(raw), requests.ConnectionError(raw),
                      requests.RequestException(raw), requests.exceptions.InvalidURL(raw)):
            with self.subTest(error=type(error).__name__):
                message = request_error_message(error)
                for sensitive in ("user", "secret", "example.com", "private", "script"):
                    self.assertNotIn(sensitive, message)

    def test_cyclic_exception_chain_does_not_loop(self):
        error = requests.ConnectionError()
        error.__cause__ = error
        self.assertEqual(request_error_message(error), "连接失败或中断")
