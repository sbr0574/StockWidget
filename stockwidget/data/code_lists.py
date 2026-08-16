# -*- coding: utf-8 -*-
"""市场代码列表的加载与下载：内置资源兜底 / 本地缓存 / GitHub 远程更新。

三份代码 JSON（沪深京、全球、期货）由 GitHub Action 每日生成；程序启动时优先用
当天本地缓存，否则先用内置资源兜底显示，再后台从 GitHub 下载刷新。
"""

import json
import os
from datetime import datetime

import requests
from PySide6.QtCore import QFile, QIODevice

from stockwidget.constants import CODES_BRANCHES, CODES_RAW_URL, LIST_FILES
from stockwidget.core.config_store import load_file, save_file


def fetch_json_from_url(url: str, timeout: int = 10):
    """从 URL 下载 JSON，失败返回 None。"""
    try:
        r = requests.get(url, timeout=timeout)
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


def download_codes(app_name: str) -> dict | None:
    """从 GitHub 下载三份代码 JSON 到本地；全部成功返回合并 codes，否则返回 None。"""
    merged = {}
    for fname in LIST_FILES:
        data = None
        for branch in CODES_BRANCHES:
            url = CODES_RAW_URL.format(branch=branch, name=fname)
            data = fetch_json_from_url(url, timeout=15)
            if data and data.get("codes"):
                break
        if not data or not data.get("codes"):
            return None
        save_file(data, app_name, fname)
        merged.update(data["codes"])
    return merged


def load_local_codes(app_name: str) -> dict:
    """合并本地三份代码 JSON。"""
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


def all_codes_fresh(app_name: str) -> bool:
    """三份本地代码 JSON 是否都是今天生成。"""
    today = datetime.now().strftime("%Y-%m-%d")
    for fname in LIST_FILES:
        f = load_file(app_name, fname)
        if (f or {}).get("last_update") != today or not f.get("codes"):
            return False
    return True
