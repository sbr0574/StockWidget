# -*- coding: utf-8 -*-
"""统一市场代码的市场前缀约定（集中定义，避免散落字符串）。

统一代码格式：
- A股    sh600519 / sz000001 / bj430047
- 港股   hk00700
- 美股   usaapl（us + 小写代码）
- 期货   au2512（具体合约）/ au0（主力连续）
- 全球指数 gnky（国际 b_）
"""

# sh/sz/bj 沪深京 | hk 港股/港股指数 | us 美股/美股指数 | g 全球指数（国际 b_）| 其余（au0 等）期货
MARKET_PREFIXES = ("sh", "sz", "bj", "hk", "us")


def market_of(code: str) -> str:
    """返回统一代码的市场标签：sh/sz/bj/hk/us；全球指数返回 "g"；期货裸码返回 ""。"""
    c = str(code or "").strip().lower()
    for p in MARKET_PREFIXES:
        if c.startswith(p) and len(c) > len(p):
            return p
    if len(c) > 1 and c.startswith("g"):
        return "g"
    return ""


def strip_market(code: str) -> str:
    """去掉市场前缀：sh600519->600519, hk00700->00700, usaapl->aapl, gnky->nky, au0->au0"""
    c = str(code or "").strip().lower()
    m = market_of(c)
    return c[len(m):] if m else c
