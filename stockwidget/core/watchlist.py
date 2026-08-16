# -*- coding: utf-8 -*-
"""自选列表（watchlist）的规范化逻辑。"""


def normalize_watchlist(watchlist: dict) -> dict:
    """规范化自选列表：代码小写，cost 转数值（整数值保持 int），name/type 转字符串。"""
    result = {}
    for key, info in (watchlist or {}).items():
        key = str(key).strip().lower()
        if not key:
            continue
        entry = dict(info or {})
        entry["checked"] = bool(entry.get("checked", True))
        try:
            val = float(entry["cost"]) if entry.get("cost") not in (None, "") else None
        except (TypeError, ValueError):
            val = None
        if val is not None and val.is_integer():
            val = int(val)
        entry["cost"] = val
        entry["name"] = str(entry.get("name", "") or "").strip()
        entry["type"] = str(entry.get("type", "") or "").strip()
        result[key] = entry
    return result
