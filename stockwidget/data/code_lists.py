# -*- coding: utf-8 -*-
"""分类代码列表的资源选择、内存索引和远端同步。"""

from datetime import date, datetime, time as date_time, timedelta, timezone
from threading import RLock

import requests

from stockwidget.constants import (
    CODES_RAW_URLS,
    CODES_STATUS_FILE,
    CODE_LIST_FILES,
)
from stockwidget.core.config_store import load_file, load_json_from_resource, save_file


CODES_UPDATE_HOUR = 9
CODES_RETRY_SECONDS = 30 * 60
_UTC8 = timezone(timedelta(hours=8))


def _time_utc8(now: datetime | None = None) -> datetime:
    if now is None:
        return datetime.now(_UTC8)
    if now.tzinfo is None:
        return now.replace(tzinfo=_UTC8)
    return now.astimezone(_UTC8)


def _previous_workday(value: date) -> date:
    value -= timedelta(days=1)
    while value.weekday() >= 5:
        value -= timedelta(days=1)
    return value


def _next_workday(value: date) -> date:
    value += timedelta(days=1)
    while value.weekday() >= 5:
        value += timedelta(days=1)
    return value


def expected_status_date(now: datetime | None = None) -> str:
    """返回当前应当使用的服务端运行日期。"""
    current = _time_utc8(now)
    if current.weekday() < 5 and current.hour >= CODES_UPDATE_HOUR:
        expected = current.date()
    else:
        expected = _previous_workday(current.date())
    return expected.isoformat()


def next_status_check_delay(now: datetime | None = None) -> int:
    """返回下一次工作日 9:00 状态日期发生变化前的秒数。"""
    current = _time_utc8(now)
    if current.weekday() < 5 and current.hour < CODES_UPDATE_HOUR:
        target_date = current.date()
    else:
        target_date = _next_workday(current.date())
    target = datetime.combine(target_date, date_time(CODES_UPDATE_HOUR), _UTC8)
    return max(1, int((target - current).total_seconds()))


def fetch_json_from_url(url: str, timeout=(2.5, 15)):
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


class CodeListManager:
    """一次性加载各分类，并在内存中维护合并后的代码索引。"""

    def __init__(self, fetcher=fetch_json_from_url):
        self._fetcher = fetcher
        self._lock = RLock()
        self._payloads: dict[str, dict] = {}
        self._codes: dict[str, dict] = {}
        self._state = "cached"
        self._state_date = ""
        self._remote_status: dict | None = None
        self._remote_expected_date = ""

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
        with self._lock:
            self._payloads = selected
            self._rebuild_codes_locked()
            self._state = (
                "current"
                if _time_utc8().weekday() >= 5 and len(selected) == len(CODE_LIST_FILES)
                else "cached"
            )
            self._state_date = self._latest_local_date_locked()
            return self._codes

    def codes(self) -> dict:
        with self._lock:
            return self._codes

    def state(self) -> tuple[str, str]:
        with self._lock:
            return self._state, self._state_date

    def sync_remote(self, now: datetime | None = None) -> dict:
        """读取远端状态，只下载内容日期与本地不同的分类文件。"""
        current = _time_utc8(now)
        expected = expected_status_date(current)
        if current.weekday() >= 5:
            with self._lock:
                complete = len(self._payloads) == len(CODE_LIST_FILES)
                self._state = "current" if complete else "cached"
                self._state_date = self._latest_local_date_locked()
                state, state_date = self._state, self._state_date
            return {
                "state": state,
                "date": state_date,
                "status_available": False,
                "retry_seconds": next_status_check_delay(current),
                "updated": (),
            }
        with self._lock:
            status = self._remote_status if self._remote_expected_date == expected else None

        if status is None:
            status = self._fetch_remote(
                CODES_STATUS_FILE,
                validator=lambda data: self._valid_status(data, expected),
            )
            if status is None:
                with self._lock:
                    self._state = "cached"
                    self._state_date = self._latest_local_date_locked()
                retry = (
                    CODES_RETRY_SECONDS
                    if current.weekday() < 5 and current.hour >= CODES_UPDATE_HOUR
                    else next_status_check_delay(current)
                )
                return {
                    "state": "cached",
                    "date": self.state()[1],
                    "status_available": False,
                    "retry_seconds": retry,
                    "updated": (),
                }
            save_file(status, CODES_STATUS_FILE)
            with self._lock:
                self._remote_status = status
                self._remote_expected_date = expected

        updated = []
        download_failed = False
        server_complete = True
        files = status.get("files", {})

        for filename in CODE_LIST_FILES:
            metadata = files.get(filename)
            if (
                not isinstance(metadata, dict)
                or metadata.get("error") is not False
                or str(metadata.get("last_checked") or "") != expected
                or not str(metadata.get("last_update") or "")
            ):
                server_complete = False
                continue

            remote_date = str(metadata["last_update"])
            with self._lock:
                local_date = str(self._payloads.get(filename, {}).get("last_update") or "")
            if local_date >= remote_date:
                continue

            payload = self._fetch_remote(
                filename,
                validator=lambda data, expected=remote_date: (
                    _valid_codes(data)
                    and str(data.get("last_update") or "") == expected
                ),
            )
            if payload is None:
                download_failed = True
                continue
            save_file(payload, filename)
            with self._lock:
                self._payloads[filename] = payload
            updated.append(filename)

        with self._lock:
            self._rebuild_codes_locked()
            synchronized = server_complete and not download_failed
            if synchronized:
                for filename in CODE_LIST_FILES:
                    remote_date = str(files[filename].get("last_update") or "")
                    local_date = str(self._payloads.get(filename, {}).get("last_update") or "")
                    if local_date < remote_date:
                        synchronized = False
                        break
            self._state = "current" if synchronized else "cached"
            self._state_date = expected if synchronized else self._latest_local_date_locked()
            state, state_date = self._state, self._state_date

        retry_seconds = CODES_RETRY_SECONDS if download_failed else next_status_check_delay(current)
        return {
            "state": state,
            "date": state_date,
            "status_available": True,
            "retry_seconds": retry_seconds,
            "updated": tuple(updated),
        }

    def _fetch_remote(self, filename: str, validator=None):
        for template in CODES_RAW_URLS:
            data = self._fetcher(template.format(name=filename))
            if isinstance(data, dict) and (validator is None or validator(data)):
                return data
        return None

    @staticmethod
    def _valid_status(status, expected: str) -> bool:
        return (
            isinstance(status, dict)
            and str(status.get("run_date") or "") == expected
            and isinstance(status.get("files"), dict)
        )

    def _rebuild_codes_locked(self) -> None:
        merged = {}
        for filename in CODE_LIST_FILES:
            payload = self._payloads.get(filename)
            if _valid_codes(payload):
                merged.update(payload["codes"])
        self._codes = merged

    def _latest_local_date_locked(self) -> str:
        dates = [
            str(payload.get("last_update") or "")
            for payload in self._payloads.values()
            if str(payload.get("last_update") or "")
        ]
        return max(dates) if dates else ""
