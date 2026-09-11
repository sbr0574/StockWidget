# -*- coding: utf-8 -*-
"""分类代码列表的本地加载、状态检查和远端同步。"""

import base64
from datetime import datetime, time as date_time, timedelta, timezone
import json
import os
from threading import RLock
from time import monotonic

import requests
from PySide6.QtCore import QFile, QIODevice

from stockwidget.constants import (
    CODES_SOURCE_URLS,
    CODES_STATUS_FILE,
    CODE_LIST_FILES,
)
from stockwidget.core.config_store import config_paths, data_cache_dir, load_file, save_file


CODES_CHECK_HOUR = 9
CODES_RETRY_SECONDS = 30 * 60
CODES_HTTP_TIMEOUT = (3, 15)
CODES_DOWNLOAD_SECONDS = 60
_UTC8 = timezone(timedelta(hours=8))


def load_json_from_resource(path: str) -> dict:
    """读取内置代码表；缺失或损坏时让调用方回退到本地缓存。"""
    file = QFile(path)
    if not file.open(QIODevice.ReadOnly | QIODevice.Text):
        return {}
    try:
        data = json.loads(bytes(file.readAll()).decode("utf-8"))
        return data if isinstance(data, dict) else {}
    except ValueError:
        return {}
    finally:
        file.close()


def _time_utc8(now: datetime | None = None) -> datetime:
    if now is None:
        return datetime.now(_UTC8)
    if now.tzinfo is None:
        return now.replace(tzinfo=_UTC8)
    return now.astimezone(_UTC8)


def _next_workday(value):
    value += timedelta(days=1)
    while value.weekday() >= 5:
        value += timedelta(days=1)
    return value


def next_code_check_delay(now: datetime | None = None) -> int:
    """返回下一次工作日 9:00 的等待秒数。"""
    current = _time_utc8(now)
    today_target = datetime.combine(
        current.date(), date_time(CODES_CHECK_HOUR), _UTC8
    )
    if current.weekday() < 5 and current < today_target:
        target_date = current.date()
    else:
        target_date = _next_workday(current.date())
    target = datetime.combine(target_date, date_time(CODES_CHECK_HOUR), _UTC8)
    return max(1, int((target - current).total_seconds()))


class CodeDownloadError(RuntimeError):
    """可向用户展示的下载失败原因。"""


def fetch_json_from_url(url: str, timeout=CODES_HTTP_TIMEOUT):
    """分别限制连接与读取等待，并在分块下载时检查总耗时。"""
    started = monotonic()
    try:
        with requests.get(
            url,
            timeout=timeout,
            stream=True,
            headers={"Cache-Control": "no-cache", "User-Agent": "StockWidget"},
        ) as response:
            response.raise_for_status()
            chunks = []
            for chunk in response.iter_content(chunk_size=64 * 1024):
                if monotonic() - started > CODES_DOWNLOAD_SECONDS:
                    raise CodeDownloadError("下载耗时过长")
                chunks.append(chunk)
            data = json.loads(b"".join(chunks))
            # Gitee 的公开文件 API 使用 Base64 包装，原始文件接口受限时仍可用。
            if "/api/v5/repos/" in url:
                if not isinstance(data, dict) or data.get("encoding") != "base64":
                    raise CodeDownloadError("文件 API 未返回有效内容")
                content = data.get("content")
                if not isinstance(content, str):
                    raise CodeDownloadError("文件 API 缺少内容")
                data = json.loads(base64.b64decode(content))
            if not isinstance(data, dict):
                raise CodeDownloadError("返回内容不是代码数据")
            return data
    except requests.Timeout as exc:
        raise CodeDownloadError("连接或读取超时") from exc
    except requests.HTTPError as exc:
        raise CodeDownloadError(f"HTTP {exc.response.status_code}") from exc
    except requests.RequestException as exc:
        raise CodeDownloadError("网络连接失败或传输中断") from exc
    except ValueError as exc:
        raise CodeDownloadError("返回内容不是有效 JSON") from exc


def _valid_codes(data) -> bool:
    return (
        isinstance(data, dict)
        and isinstance(data.get("codes"), dict)
        and bool(data["codes"])
        and bool(str(data.get("last_update") or ""))
    )


def _status_run_time(status: dict) -> str:
    """返回可排序的状态运行时间，兼容只有 run_date 的旧状态文件。"""
    for key in ("started_at", "completed_at", "run_date"):
        value = str((status or {}).get(key) or "").strip()
        if value:
            return value
    return ""


def _valid_status(status) -> bool:
    return (
        isinstance(status, dict)
        and bool(str(status.get("run_date") or ""))
        and bool(_status_run_time(status))
        and isinstance(status.get("files"), dict)
    )


class CodeListManager:
    """一次性加载各分类，并在内存中维护合并后的代码索引。"""

    def __init__(self, fetcher=fetch_json_from_url):
        self._fetcher = fetcher
        self._lock = RLock()
        self._payloads: dict[str, dict] = {}
        self._codes: dict[str, dict] = {}
        self._state = "cached"
        self._state_date = ""
        self._local_status: dict = {}
        self._preferred_source = CODES_SOURCE_URLS[0]
        self._failed_sources = set()
        self._last_error = ""
        self._pending_run = ""
        self._pending_downloads: dict[str, dict] = {}

    def _load_cached(self, filename, validator):
        cached = load_file(filename, directory=data_cache_dir())
        if validator(cached):
            return cached
        legacy = load_file(filename)
        if not validator(legacy):
            return {}
        try:
            save_file(legacy, filename, directory=data_cache_dir())
            os.remove(os.path.join(config_paths(), filename))
        except OSError:
            pass
        return legacy

    def load_local(self) -> dict:
        """逐份选择 qrc 与本地缓存中内容日期较新的有效版本。"""
        selected = {}
        for filename in CODE_LIST_FILES:
            resource = load_json_from_resource(f":/{filename}")
            local = self._load_cached(filename, _valid_codes)
            candidates = [data for data in (resource, local) if _valid_codes(data)]
            if candidates:
                # 同日时本地缓存排在后面并胜出。
                selected[filename] = max(
                    enumerate(candidates),
                    key=lambda item: (str(item[1].get("last_update") or ""), item[0]),
                )[1]
        local_status = self._load_cached(CODES_STATUS_FILE, _valid_status)
        with self._lock:
            self._payloads = selected
            self._local_status = local_status if _valid_status(local_status) else {}
            self._rebuild_codes_locked()
            self._state = "cached"
            self._state_date = str(self._local_status.get("run_date") or "")
            return self._codes

    def codes(self) -> dict:
        with self._lock:
            return self._codes

    def state(self) -> tuple[str, str]:
        with self._lock:
            return self._state, self._state_date

    def last_error(self) -> str:
        with self._lock:
            return self._last_error

    def begin_remote_check(self) -> None:
        """开始启动/定时检查时立即将界面状态置为缓存。"""
        with self._lock:
            self._state = "cached"
            self._state_date = str(self._local_status.get("run_date") or "")
            self._last_error = ""
        self._failed_sources.clear()

    def sync_remote(self, now: datetime | None = None) -> dict:
        """比较本地与远端运行时间；远端较新时完整下载九个代码列表。"""
        current = _time_utc8(now)
        self.begin_remote_check()
        status = self._fetch_remote(CODES_STATUS_FILE, validator=_valid_status)
        if status is None:
            return self._sync_result(CODES_RETRY_SECONDS)

        with self._lock:
            local_run_time = _status_run_time(self._local_status)
            local_complete = all(
                _valid_codes(self._payloads.get(filename))
                for filename in CODE_LIST_FILES
            )

        remote_run_time = _status_run_time(status)
        needs_download = remote_run_time > local_run_time or not local_complete
        updated = ()
        if needs_download:
            if self._pending_run != remote_run_time:
                self._pending_run = remote_run_time
                self._pending_downloads = {}
            downloaded = self._pending_downloads
            for filename in CODE_LIST_FILES:
                if filename in downloaded:
                    continue
                payload = self._fetch_remote(filename, validator=_valid_codes)
                if payload is None:
                    return self._sync_result(CODES_RETRY_SECONDS)
                downloaded[filename] = payload

            try:
                for filename in CODE_LIST_FILES:
                    save_file(downloaded[filename], filename, directory=data_cache_dir())
                # 状态文件最后保存：它代表九个列表均已完整落盘。
                save_file(status, CODES_STATUS_FILE, directory=data_cache_dir())
            except OSError as exc:
                with self._lock:
                    self._last_error = f"代码缓存保存失败：{exc}"
                return self._sync_result(CODES_RETRY_SECONDS)

            with self._lock:
                self._payloads = downloaded
                self._local_status = status
                self._rebuild_codes_locked()
            updated = tuple(CODE_LIST_FILES)
            self._pending_downloads = {}

        with self._lock:
            self._state = "current"
            self._state_date = str(self._local_status.get("run_date") or "")
            state_date = self._state_date
        return {
            "state": "current",
            "date": state_date,
            "retry_seconds": next_code_check_delay(current),
            "updated": updated,
        }

    def _sync_result(self, retry_seconds: int) -> dict:
        """保持缓存状态，并返回统一的失败重试结果。"""
        with self._lock:
            self._state = "cached"
            self._state_date = str(self._local_status.get("run_date") or "")
            state_date = self._state_date
        return {
            "state": "cached",
            "date": state_date,
            "retry_seconds": retry_seconds,
            "updated": (),
        }

    def _fetch_remote(self, filename: str, validator=None):
        errors = []
        sources = sorted(
            CODES_SOURCE_URLS,
            key=lambda source: (source in self._failed_sources, source != self._preferred_source),
        )
        for template in sources:
            url = template.format(name=filename)
            try:
                data = self._fetcher(url)
                error = "内容缺失或格式不正确"
            except (CodeDownloadError, requests.RequestException, ValueError) as exc:
                data, error = None, str(exc)
            if isinstance(data, dict) and (validator is None or validator(data)):
                self._preferred_source = template
                return data
            source_name = (
                "GitHub" if "githubusercontent" in template
                else "Gitee API" if "/api/" in template else "Gitee"
            )
            self._failed_sources.add(template)
            errors.append(f"{source_name}：{error}")
        with self._lock:
            self._last_error = f"{filename} 更新失败；" + "；".join(errors)
        return None

    def _rebuild_codes_locked(self) -> None:
        merged = {}
        for filename in CODE_LIST_FILES:
            payload = self._payloads.get(filename)
            if _valid_codes(payload):
                merged.update(payload["codes"])
        self._codes = merged
