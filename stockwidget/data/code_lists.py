# -*- coding: utf-8 -*-
"""市场代码列表的加载与下载：内置资源兜底 / 本地缓存 / 远程更新。

两份代码 JSON（证券、期货）由 GitHub Action 每日生成；程序逐份使用本地缓存或
内置资源，并在北京时间 9 点后从 GitHub/Gitee 更新尚未取得当日版本的文件。
"""

import json
from datetime import datetime, timedelta, timezone

import requests
from PySide6.QtCore import QFile, QIODevice

from stockwidget.constants import CODES_BRANCHES, CODES_RAW_BACKUP, CODES_RAW_URL, LIST_FILES
from stockwidget.core.config_store import load_file, save_file


BEIJING_TIMEZONE = timezone(timedelta(hours=8))
CODES_UPDATE_HOUR = 9
CODES_RETRY_SECONDS = 30 * 60


def beijing_now() -> datetime:
    return datetime.now(BEIJING_TIMEZONE)


def _beijing_time(now: datetime | None = None) -> datetime:
    current = now or beijing_now()
    if current.tzinfo is None:
        return current.replace(tzinfo=BEIJING_TIMEZONE)
    return current.astimezone(BEIJING_TIMEZONE)


def beijing_today(now: datetime | None = None) -> str:
    return _beijing_time(now).strftime("%Y-%m-%d")


def is_code_update_day(now: datetime | None = None) -> bool:
    """代码列表只在北京时间周一至周五检查更新。"""
    return _beijing_time(now).weekday() < 5


def code_refresh_delay_seconds(now: datetime | None = None) -> int:
    """北京时间 9 点前返回距 9 点的秒数，9 点后返回 0。"""
    current = _beijing_time(now)
    if current.hour >= CODES_UPDATE_HOUR:
        return 0
    target = current.replace(
        hour=CODES_UPDATE_HOUR, minute=0, second=0, microsecond=0
    )
    return max(1, int((target - current).total_seconds()))


def fetch_json_from_url(url: str, timeout: int = 10):
    """从 URL 下载 JSON，失败返回 None。"""
    try:
        r = requests.get(
            url,
            timeout=timeout,
            headers={"Cache-Control": "no-cache", "User-Agent": "StockWidget"},
        )
        r.raise_for_status()
        return r.json()
    except Exception:
        return None


def load_json_from_resource(path: str) -> dict:
    """从 Qt 资源系统读取 JSON 文件（path 形如 ':/stock_codes_list.json'）。"""
    file = QFile(path)
    if not file.open(QIODevice.ReadOnly | QIODevice.Text):
        raise FileNotFoundError(f"无法打开资源文件: {path}")
    content = file.readAll()
    file.close()

    text = bytes(content).decode('utf-8')
    return json.loads(text)


def download_code_files(
    app_name: str,
    filenames,
    source: str = "github",
    expected_date: str | None = None,
) -> tuple[str, ...]:
    """下载指定代码文件；仅当远端日期为今天时保存，返回已更新文件名。"""
    url_template = CODES_RAW_BACKUP if source == "gitee" else CODES_RAW_URL
    expected_date = expected_date or beijing_today()
    requested = set(filenames)
    updated = []
    for fname in (name for name in LIST_FILES if name in requested):
        data = None
        for branch in CODES_BRANCHES:
            url = url_template.format(branch=branch, name=fname)
            candidate = fetch_json_from_url(url, timeout=(3, 30))
            if (
                isinstance(candidate, dict)
                and candidate.get("codes")
                and candidate.get("last_update") == expected_date
            ):
                data = candidate
                break
        if data is not None:
            save_file(data, app_name, fname)
            updated.append(fname)
    return tuple(updated)


def download_codes(app_name: str, source: str = "github") -> dict | None:
    """兼容入口：两份当日文件全部下载成功时返回合并结果。"""
    updated = download_code_files(app_name, LIST_FILES, source)
    if set(updated) != set(LIST_FILES):
        return None
    return load_local_codes(app_name)


def load_local_codes(app_name: str) -> dict:
    """合并本地两份代码 JSON。"""
    merged = {}
    for fname in LIST_FILES:
        f = load_file(app_name, fname)
        merged.update((f or {}).get("codes", {}) or {})
    return merged


def load_resource_codes() -> dict:
    """从 Qt 资源内嵌的代码 JSON 合并。"""
    merged = {}
    for fname in LIST_FILES:
        try:
            res = load_json_from_resource(f":/{fname}")
        except FileNotFoundError:
            res = {}
        merged.update((res or {}).get("codes", {}) or {})
    return merged


def load_best_codes(app_name: str) -> dict:
    """逐文件选择本地缓存与内置资源中日期较新的有效版本并合并。"""
    merged = {}
    for fname in LIST_FILES:
        local = load_file(app_name, fname)
        try:
            resource = load_json_from_resource(f":/{fname}")
        except FileNotFoundError:
            resource = {}
        candidates = [
            data
            for data in (local, resource)
            if isinstance(data, dict) and data.get("codes")
        ]
        if not candidates:
            continue
        selected = max(
            candidates,
            key=lambda data: str(data.get("last_update") or ""),
        )
        merged.update(selected["codes"])
    return merged


def stale_code_files(
    app_name: str,
    today: str | None = None,
    now: datetime | None = None,
) -> tuple[str, ...]:
    """返回本地缓存中缺失、为空或更新日期不是今天的代码文件。"""
    current = _beijing_time(now)
    if not is_code_update_day(current):
        return ()
    expected_date = today or beijing_today(current)
    stale = []
    for fname in LIST_FILES:
        data = load_file(app_name, fname)
        if (
            not isinstance(data, dict)
            or not data.get("codes")
            or data.get("last_update") != expected_date
        ):
            stale.append(fname)
    return tuple(stale)


def all_codes_fresh(app_name: str) -> bool:
    """两份本地代码 JSON 是否都是今天生成。"""
    return not stale_code_files(app_name)


def code_data_state(app_name: str) -> tuple[str, str]:
    """返回市场代码数据的状态与更新日期 (state, date)。

    - ('current', 'YYYY-MM-DD') ：周末不检查远端，现有列表视为最新。
    - ('online', 'YYYY-MM-DD')  ：本地两份 JSON 均为今天生成。
    - ('cached', 'YYYY-MM-DD')  ：本地存在旧 JSON（当日未刷新或刷新失败）。
    - ('offline', 'YYYY-MM-DD') ：无本地缓存，使用内置 qrc 资源。
    """
    current = _beijing_time()
    today = beijing_today(current)
    local_dates = []
    for fname in LIST_FILES:
        f = load_file(app_name, fname)
        if isinstance(f, dict) and f.get("codes"):
            d = str(f.get("last_update") or "").strip()
            if d:
                local_dates.append(d)
    if not is_code_update_day(current):
        dates = list(local_dates)
        for fname in LIST_FILES:
            try:
                res = load_json_from_resource(f":/{fname}")
            except FileNotFoundError:
                continue
            d = str(res.get("last_update") or "").strip() if isinstance(res, dict) else ""
            if d:
                dates.append(d)
        return "current", max(dates) if dates else today
    if local_dates:
        if not stale_code_files(app_name, today, current):
            return "online", today
        return "cached", max(local_dates)

    res_dates = []
    for fname in LIST_FILES:
        try:
            res = load_json_from_resource(f":/{fname}")
        except FileNotFoundError:
            continue
        d = str(res.get("last_update") or "").strip() if isinstance(res, dict) else ""
        if d:
            res_dates.append(d)
    return "offline", max(res_dates) if res_dates else today
