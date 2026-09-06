# -*- coding: utf-8 -*-
"""行情数值的显示格式化（纯函数，供 UI 层调用）。"""

# 自动单位模式下使用英文单位（k/M/B/T）的市场：美股与国际指数。
ENGLISH_UNIT_MARKETS = frozenset({"us", "gb"})


def format_value(value: float, lot_size: int = 1, unit_cn: bool = True) -> str:
    """统一格式化成交量/成交额等数值。

    unit_cn=True 时按 万/亿/万亿 缩写，否则按 k/M/B/T 缩写；
    lot_size 用于把成交量按每手股数换算后再缩写。
    """
    value = int(value / lot_size)
    if unit_cn:
        if value < 1e4:
            return f"{value}"
        if value < 1e8:
            return f"{value / 1e4:.2f}万"
        if value < 1e12:
            return f"{value / 1e8:.2f}亿"
        return f"{value / 1e12:.2f}万亿"
    if value < 1e3:
        return f"{value}"
    if value < 1e6:
        return f"{value / 1e3:.2f}k"
    if value < 1e9:
        return f"{value / 1e6:.2f}M"
    if value < 1e12:
        return f"{value / 1e9:.2f}B"
    return f"{value / 1e12:.2f}T"


def should_use_english_units(unit_mode: str, market: str = "") -> bool:
    """按单位模式决定是否使用英文单位。

    模式为 中文/英文 时按模式返回；“自动”时美股与国际指数
    （market 为 us/gb）使用英文单位，国内、港股等使用中文单位。
    """
    mode = str(unit_mode or "").strip().lower()
    if mode in ("en", "english", "英文"):
        return True
    if mode in ("cn", "zh", "chinese", "中文"):
        return False
    return str(market or "").strip().lower() in ENGLISH_UNIT_MARKETS
