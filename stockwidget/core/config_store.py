# -*- coding: utf-8 -*-
"""配置文件的读写：路径定位、JSON 加载与原子保存（纯 Python，无 Qt 依赖）。"""

import json
import os


def config_paths(app_name: str) -> str:
    """配置文件所在目录：Windows 用 %APPDATA%，其余平台用用户主目录。"""
    return os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), app_name)


def load_file(app_name: str, file_name: str, except_ret: dict | None = None) -> dict:
    """读取配置文件为 dict；文件不存在或解析失败时返回 fallback（默认 {}）。"""
    fallback = {} if except_ret is None else except_ret
    path = os.path.join(config_paths(app_name), file_name)
    if not os.path.exists(path):
        return fallback
    try:
        with open(path, "r", encoding="utf-8") as file:
            return json.load(file)
    except (OSError, json.JSONDecodeError):
        return fallback


def save_file(data: dict, app_name: str, file_name: str) -> None:
    """原子化保存 dict 到配置文件（先写临时文件再替换，避免写一半损坏）。"""
    os.makedirs(config_paths(app_name), exist_ok=True)
    config_file = os.path.join(config_paths(app_name), file_name)
    tmp_file = config_file + ".tmp"
    with open(tmp_file, "w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)
    os.replace(tmp_file, config_file)
