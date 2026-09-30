"""请求错误分类、底层连接原因与凭据保护。"""

from http.client import RemoteDisconnected
import socket
import ssl
import unittest

import requests
from urllib3.exceptions import MaxRetryError, NameResolutionError, ProtocolError, ProxyError

from stockwidget.data.network_errors import request_error_message


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
