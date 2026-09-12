# -*- coding: utf-8 -*-
"""自选列表（watchlist）的规范化逻辑。"""

from math import isfinite


def parse_positive_cost(value) -> float | None:
    """成本只接受有限正数，供配置读取与界面编辑共用。"""
    try:
        cost = float(value)
    except (TypeError, ValueError):
        return None
    return cost if isfinite(cost) and cost > 0 else None


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
        val = parse_positive_cost(entry.get("cost"))
        if val is not None and val.is_integer():
            val = int(val)
        entry["cost"] = val
        entry["name"] = str(entry.get("name", "") or "").strip()
        entry["type"] = str(entry.get("type", "") or "").strip()
        entry["market"] = str(entry.get("market", "") or "").strip().lower()
        entry["code"] = str(entry.get("code", "") or "").strip().lower()
        result[key] = entry
    return result
