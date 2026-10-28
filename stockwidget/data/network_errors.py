"""把请求异常转换为简短、可展示且不含 URL 或代理凭据的原因。"""

from http.client import RemoteDisconnected
import socket
import ssl

import requests


def _exception_chain(error):
    """Requests / urllib3 常将底层异常放在 reason 或 args 中。"""
    pending = [error]
    seen = set()
    while pending:
        current = pending.pop()
        if not isinstance(current, BaseException) or id(current) in seen:
            continue
        seen.add(id(current))
        yield current
        pending.extend(current.args)
        for attr in ("reason", "_reason", "original_error", "__cause__", "__context__"):
            pending.append(getattr(current, attr, None))


def _connection_reason(error):
    for cause in _exception_chain(error):
        if isinstance(cause, socket.gaierror):
            return "域名解析失败"
        if isinstance(cause, ConnectionRefusedError):
            return "连接被拒绝"
        if isinstance(cause, RemoteDisconnected):
            return "远端关闭连接"
        if isinstance(cause, (ConnectionResetError, ConnectionAbortedError)):
            return "连接中断"
    return None


def request_error_message(error: requests.RequestException) -> str:
    """先判断具体子类，避免代理、证书、JSON 等错误被归为普通连接失败。"""
    if isinstance(error, requests.exceptions.ProxyError):
        reason = _connection_reason(error)
        return f"代理请求失败（{reason}）" if reason else "代理请求失败"
    if isinstance(error, requests.exceptions.SSLError):
        if any(isinstance(cause, ssl.SSLCertVerificationError)
               for cause in _exception_chain(error)):
            return "证书验证失败"
        return "安全连接失败（SSL）"
    if isinstance(error, requests.ConnectTimeout):
        return "连接超时"
    if isinstance(error, requests.ReadTimeout):
        return "读取超时"
    if isinstance(error, requests.Timeout):
        return "请求超时"
    if isinstance(error, requests.HTTPError):
        response = error.response
        if response is None or response.status_code is None:
            return "HTTP 请求失败"
        status = response.status_code
        reason = {
            401: "访问未授权", 403: "访问被拒绝", 404: "接口不存在",
            429: "请求过于频繁", 500: "服务端内部错误", 502: "网关错误",
            503: "服务暂不可用", 504: "网关超时",
        }.get(status, "请求失败")
        return f"HTTP {status}：{reason}"
    if isinstance(error, requests.exceptions.JSONDecodeError):
        return "响应不是有效 JSON"
    if isinstance(error, requests.exceptions.ChunkedEncodingError):
        return "响应传输中断"
    if isinstance(error, requests.exceptions.ContentDecodingError):
        return "响应解码失败"
    if isinstance(error, requests.TooManyRedirects):
        return "重定向次数过多"
    if isinstance(error, (requests.exceptions.InvalidURL, requests.exceptions.MissingSchema,
                          requests.exceptions.InvalidSchema, requests.exceptions.URLRequired)):
        return "请求地址无效"
    if isinstance(error, requests.exceptions.InvalidHeader):
        return "请求头无效"
    if isinstance(error, requests.ConnectionError):
        return _connection_reason(error) or "连接失败或中断"
    return "网络请求失败"
