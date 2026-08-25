# -*- coding: utf-8 -*-
"""自选列表（watchlist）的规范化逻辑。"""


def normalize_watchlist(watchlist: dict, codes: dict | None = None) -> dict:
    """规范化自选列表，并从代码表补齐旧配置缺少的证券元数据。"""
    result = {}
    for key, info in (watchlist or {}).items():
        key = str(key).strip().lower()
        if not key:
            continue
        entry = dict((codes or {}).get(key) or {})
        entry.update(info or {})
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
        entry["market"] = str(entry.get("market", "") or "").strip().lower()
        entry["code"] = str(entry.get("code", "") or "").strip().lower()
        result[key] = entry
    return result
