# -*- coding: utf-8 -*-
"""自选股代码搜索与规范化（纯业务逻辑，无 Qt/网络依赖）。

根据数字代码、拼音、首字母、中文名/英文名在代码列表中匹配建议。
"""

import re

from stockwidget.core.markets import MARKET_PREFIXES, strip_market


def code_without_market(code: str) -> str:
    """去掉代码的市场前缀（兼容旧命名，等价于 strip_market）。"""
    return strip_market(code)


def normalize_stock_entry(item: dict) -> dict:
    """把代码列表中的原始条目规范化为统一的搜索字段 dict。"""
    market = str(item.get("market", "") or "").strip().lower()
    code = str(item.get("code", "") or "").strip()
    if market in {"sh", "sz", "bj"} and code.isdigit():
        code = code.zfill(6)
    key = str(item.get("key", "") or "").strip().lower()
    if not key and market and code:
        key = f"{market}{code}"
    return {
        "key": key,
        "market": market,
        "code": code,
        "name": str(item.get("name", "") or "").strip(),
        "type": str(item.get("type", "") or "").strip(),
        "py": str(item.get("py", "") or "").strip().lower(),
        "abbr": str(item.get("abbr", "") or "").strip().lower(),
        "engname": str(item.get("engname", "") or "").strip().lower(),
    }


def _query_variants(text: str) -> set[str]:
    """把用户输入扩展成多个候选查询（去市场前缀、数字补零等）。"""
    q = str(text or "").strip().lower().replace(" ", "")
    variants = {q}
    if len(q) == 8 and q[:2] in MARKET_PREFIXES:
        variants.add(q[2:])
    digits = re.sub(r"\D", "", q)
    if digits:
        variants.add(digits.zfill(6) if len(digits) <= 6 else digits)
    return {v for v in variants if v}


def find_suggestions(codes: dict, text: str, limit: int = 20) -> list[dict]:
    """在代码列表中按相关度返回匹配建议（精确 > 前缀 > 包含）。"""
    queries = _query_variants(text)
    if not queries or not isinstance(codes, dict):
        return []

    scored = []
    for raw_key, raw_info in codes.items():
        info = normalize_stock_entry({"key": raw_key, **(raw_info or {})})
        fields = {
            "key": info["key"],
            "code": info["code"],
            "name": info["name"].lower(),
            "py": info["py"],
            "abbr": info["abbr"],
            "engname": info["engname"],
        }
        best = 0
        for q in queries:
            if (fields["key"] == q or fields["code"] == q or fields["name"] == q
                    or fields["py"] == q or fields["abbr"] == q or fields["engname"] == q):
                best = max(best, 120)
            elif fields["key"].startswith(q):
                best = max(best, 110)
            elif fields["code"].startswith(q):
                best = max(best, 105)
            elif q in fields["key"] or q in fields["code"]:
                best = max(best, 95)
            elif fields["name"].startswith(q) or fields["engname"].startswith(q):
                best = max(best, 90)
            elif q in fields["name"] or q in fields["engname"]:
                best = max(best, 80)
            elif fields["abbr"].startswith(q):
                best = max(best, 75)
            elif q in fields["abbr"]:
                best = max(best, 65)
            elif fields["py"].startswith(q):
                best = max(best, 60)
            elif q in fields["py"]:
                best = max(best, 50)
        if best:
            scored.append((best, info))

    scored.sort(key=lambda pair: (-pair[0], pair[1]["code"], pair[1]["key"]))
    return [item for _, item in scored[:limit]]
