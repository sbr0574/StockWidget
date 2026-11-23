"""新浪行情解析和安全的网络错误分类。"""

from http.client import RemoteDisconnected
import json
from pathlib import Path
from unittest.mock import Mock, patch
import socket
import ssl
import unittest

from urllib3.exceptions import MaxRetryError, NameResolutionError, ProtocolError, ProxyError
import requests

from stockwidget.core.quote_presentation import MARKET_BREADTH_BUCKETS, is_cn_index
from stockwidget.data import quotes
from stockwidget.data.network_errors import request_error_message
from stockwidget.data.quotes import MARKET_BREADTH_KEY, _sina_code, parse_sina_market_breadth


class ExplicitMarketMetadataTests(unittest.TestCase):
    def test_shanghai_futures_use_futures_provider_symbols(self):
        for code in ("au0", "au2610", "sc0", "sc2610"):
            with self.subTest(code=code):
                future = {"market": "sh", "code": code, "type": "期"}
                self.assertEqual(_sina_code(future), "nf_" + code.upper())

    def test_shanghai_futures_sina_parser_and_volume(self):
        parts = ["黄金连续", "150000"] + ["0"] * 26
        parts[7], parts[10], parts[14] = "800", "790", "1234"
        response = Mock(text='var hq_str_nf_AU0="' + ','.join(parts) + '";')
        with patch.object(quotes.requests, "get", return_value=response):
            data = quotes.request_quote({"au0": {"market": "sh", "code": "au0", "type": "期"}})
        self.assertEqual(data["au0"]["current_price"], 800)
        self.assertEqual(data["au0"]["prev_close"], 790)
        self.assertEqual(data["au0"]["deals_vol"], 1234)

    def test_same_raw_code_keeps_markets_distinct(self):
        shanghai = {"market": "sh", "code": "000001"}
        shenzhen = {"market": "sz", "code": "000001"}

        self.assertEqual(_sina_code(shanghai), "sh000001")
        self.assertEqual(_sina_code(shenzhen), "sz000001")

    def test_futures_uses_empty_market_and_raw_code(self):
        future = {"market": "", "code": "au0"}
        self.assertEqual(_sina_code(future), "nf_AU0")


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
        response._content_consumed = True
        with patch.object(quotes.requests, "get", return_value=response):
            with self.assertRaises(requests.HTTPError) as caught:
                quotes.request_quote({"sh600000": {"market": "sh", "code": "600000"}})
        self.assertIs(caught.exception.response, response)


class SinaMarketBreadthTests(unittest.TestCase):
    SAMPLE_ZDP = "4632536,2026-11-23,179,306,508,1931,1573,165,756,187,27,10,5,14:45:04"

    def test_is_cn_index_matches_index_cn_json_and_excludes_other_categories(self):
        root = Path(__file__).resolve().parents[2] / "resources" / "data"
        cn_codes = json.loads((root / "index_cn.json").read_text(encoding="utf-8"))["codes"]
        global_codes = json.loads((root / "index_global.json").read_text(encoding="utf-8"))["codes"]
        self.assertTrue(cn_codes)
        self.assertTrue(all(is_cn_index(item) for item in cn_codes.values()))
        self.assertFalse(any(is_cn_index(item) for item in global_codes.values()))
        for non_cn_index in (
            {"market": "sh", "code": "600000", "type": "沪"},
            {"market": "sz", "code": "000001", "type": "深"},
            {"market": "bj", "code": "920000", "type": "京"},
            {"market": "sh", "code": "510300", "type": "基"},
            {"market": "sh", "code": "au0", "type": "期"},
            None,
            {},
        ):
            self.assertFalse(is_cn_index(non_cn_index))

    def test_parse_sina_market_breadth_valid_and_invalid_payloads(self):
        parsed = parse_sina_market_breadth(self.SAMPLE_ZDP.split(","))
        self.assertIsNotNone(parsed)
        self.assertEqual(parsed["seq"], 4632536)
        self.assertEqual(parsed["date"], "2026-11-23")
        self.assertEqual(parsed["time"], "14:45:04")
        self.assertIsNotNone(parsed["timestamp"])
        self.assertEqual(parsed["buckets"], (179, 306, 508, 1931, 1573, 165, 756, 187, 27, 10, 5))
        self.assertEqual(len(parsed["buckets"]), len(MARKET_BREADTH_BUCKETS))
        self.assertEqual((parsed["up"], parsed["flat"], parsed["down"], parsed["total"]), (4497, 165, 985, 5647))

        zeros = parse_sina_market_breadth(["0", "2026-11-23", *(["0"] * 11), "09:15:00"])
        self.assertEqual((zeros["up"], zeros["flat"], zeros["down"], zeros["total"]), (0, 0, 0, 0))

        for bad in ([], ["4632536", "2026-11-23", "1"], ["1", "2026-11-23", "-1", *(["0"] * 10), "10:00:00"],
                    ["1", "2026-11-23", "abc", *(["0"] * 10), "10:00:00"]):
            self.assertIsNone(parse_sina_market_breadth(bad))

    def test_request_quote_merges_zdp_updown_only_when_requested(self):
        index_parts = ["上证指数", "3500", "3510", "3520", "3530", "3490", "0", "0", "100", "200"] + ["0"] * 20 + ["2026-11-23", "14:45:04"]
        index_line = 'var hq_str_sh000001="' + ",".join(index_parts) + '";'
        zdp_line = f'var hq_str_zdp_updown="{self.SAMPLE_ZDP}";'
        instruments = {"sh000001": {"market": "sh", "code": "000001", "type": "指"}}

        plain_resp = Mock(text=index_line)
        with patch.object(quotes.requests, "get", return_value=plain_resp) as get_plain:
            plain = quotes.request_quote(instruments)
        self.assertNotIn("zdp_updown", get_plain.call_args.args[0])
        self.assertNotIn(MARKET_BREADTH_KEY, plain)

        merged_resp = Mock(text=index_line + "\n" + zdp_line)
        with patch.object(quotes.requests, "get", return_value=merged_resp) as get_merged:
            merged = quotes.request_quote(instruments, include_breadth=True)
        self.assertTrue(get_merged.call_args.args[0].endswith("list=sh000001,zdp_updown"))
        self.assertEqual(merged["sh000001"]["current_price"], 3520.0)
        self.assertEqual(merged[MARKET_BREADTH_KEY]["buckets"], (179, 306, 508, 1931, 1573, 165, 756, 187, 27, 10, 5))

        bad_zdp_resp = Mock(text=index_line + '\nvar hq_str_zdp_updown="";')
        with patch.object(quotes.requests, "get", return_value=bad_zdp_resp):
            partial = quotes.request_quote(instruments, include_breadth=True)
        self.assertEqual(partial["sh000001"]["current_price"], 3520.0)
        self.assertNotIn(MARKET_BREADTH_KEY, partial)

        only_zdp_resp = Mock(text=zdp_line)
        with patch.object(quotes.requests, "get", return_value=only_zdp_resp):
            ok, payload, error = quotes.fetch_quote_result(instruments, include_breadth=True)
        self.assertFalse(ok)
        self.assertIsNone(payload)
        self.assertIn("新浪", error)

    def test_chunked_request_appends_zdp_updown_only_once(self):
        instruments = {f"sh{600000 + i}": {"market": "sh", "code": str(600000 + i), "type": "沪"} for i in range(125)}
        calls = []

        def fake_get(url, **kwargs):
            calls.append(url)
            symbols = url.split("list=", 1)[1].split(",")
            lines = []
            for sym in symbols:
                if sym == "zdp_updown":
                    lines.append(f'var hq_str_zdp_updown="{self.SAMPLE_ZDP}";')
                else:
                    parts = ["股", "10", "10", "10", "10", "10", "0", "0", "100", "1000"] + ["0"] * 20 + ["2026-11-23", "14:45:04"]
                    lines.append(f'var hq_str_{sym}="{",".join(parts)}";')
            return Mock(text="\n".join(lines))

        with patch.object(quotes.requests, "get", side_effect=fake_get):
            data = quotes.request_quote(instruments, include_breadth=True)
        self.assertEqual(len(calls), 2)
        self.assertIn("zdp_updown", calls[0])
        self.assertNotIn("zdp_updown", calls[1])
        self.assertIn(MARKET_BREADTH_KEY, data)




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
