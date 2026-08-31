# -*- coding: utf-8 -*-
"""分类代码列表的本地加载、状态检查和远端同步。"""

from datetime import datetime, time as date_time, timedelta, timezone
from threading import RLock

import requests

from stockwidget.constants import (
    CODES_RAW_URLS,
    CODES_STATUS_FILE,
    CODE_LIST_FILES,
)
from stockwidget.core.config_store import load_file, load_json_from_resource, save_file


CODES_CHECK_HOUR = 9
CODES_RETRY_SECONDS = 30 * 60
_UTC8 = timezone(timedelta(hours=8))


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


def fetch_json_from_url(url: str, timeout=5):
    """从 URL 下载 JSON；连接、状态码或解析失败时返回 None。"""
    try:
        response = requests.get(
            url,
            timeout=timeout,
            headers={"Cache-Control": "no-cache", "User-Agent": "StockWidget"},
        )
        response.raise_for_status()
        data = response.json()
        return data if isinstance(data, dict) else None
    except (requests.RequestException, ValueError):
        return None


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

    def load_local(self) -> dict:
        """逐份选择 qrc 与本地缓存中内容日期较新的有效版本。"""
        selected = {}
        for filename in CODE_LIST_FILES:
            resource = load_json_from_resource(f":/{filename}")
            local = load_file(filename)
            candidates = [data for data in (resource, local) if _valid_codes(data)]
            if candidates:
                # 同日时本地缓存排在后面并胜出。
                selected[filename] = max(
                    enumerate(candidates),
                    key=lambda item: (str(item[1].get("last_update") or ""), item[0]),
                )[1]
        local_status = load_file(CODES_STATUS_FILE)
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

    def begin_remote_check(self) -> None:
        """开始启动/定时检查时立即将界面状态置为缓存。"""
        with self._lock:
            self._state = "cached"
            self._state_date = str(self._local_status.get("run_date") or "")

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
            downloaded = {}
            for filename in CODE_LIST_FILES:
                payload = self._fetch_remote(filename, validator=_valid_codes)
                if payload is None:
                    return self._sync_result(CODES_RETRY_SECONDS)
                downloaded[filename] = payload

            try:
                for filename in CODE_LIST_FILES:
                    save_file(downloaded[filename], filename)
                # 状态文件最后保存：它代表九个列表均已完整落盘。
                save_file(status, CODES_STATUS_FILE)
            except OSError:
                return self._sync_result(CODES_RETRY_SECONDS)

            with self._lock:
                self._payloads = downloaded
                self._local_status = status
                self._rebuild_codes_locked()
            updated = tuple(CODE_LIST_FILES)

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
        for template in CODES_RAW_URLS:
            data = self._fetcher(template.format(name=filename))
            if isinstance(data, dict) and (validator is None or validator(data)):
                return data
        return None

    def _rebuild_codes_locked(self) -> None:
        merged = {}
        for filename in CODE_LIST_FILES:
            payload = self._payloads.get(filename)
            if _valid_codes(payload):
                merged.update(payload["codes"])
        self._codes = merged
