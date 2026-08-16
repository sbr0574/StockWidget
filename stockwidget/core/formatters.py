# -*- coding: utf-8 -*-
"""行情数值的显示格式化（纯函数，供 UI 层调用）。"""


def format_volume(value: int) -> str:
    """成交量：股 -> 手（÷100），并按 万/亿 缩写。"""
    value = int(value / 100)
    if value < 1e4:
        return f"{value}"
    if value < 1e8:
        return f"{value / 1e4:.2f}万"
    return f"{value / 1e8:.2f}亿"


def format_amount(value: float) -> str:
    """成交额：按 万/亿/万亿 缩写。"""
    if value < 1e8:
        return f"{value / 1e4:.2f}万"
    if value < 1e12:
        return f"{value / 1e8:.2f}亿"
    return f"{value / 1e12:.2f}万亿"
