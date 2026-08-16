# -*- coding: utf-8 -*-
"""窗口位置的跨屏恢复逻辑（纯函数，无 Qt 依赖，可单元测试）。

屏幕矩形统一用 ``(left, top, width, height)`` 元组表示，坐标可为负（副屏在主屏左侧时）。
"""


def screen_containing(x, y, rects):
    """返回包含点 (x, y) 的第一个屏幕矩形；找不到返回 None。"""
    for r in rects:
        left, top, w, h = r
        if left <= x < left + w and top <= y < top + h:
            return r
    return None


def clamp_point(x, y, rect, widget_w=0, widget_h=0):
    """把左上角 (x, y) 夹取到 rect 内，保证 widget 不超出该屏幕。"""
    left, top, w, h = rect
    right = left + w - max(1, widget_w)
    bottom = top + h - max(1, widget_h)
    return max(left, min(x, right)), max(top, min(y, bottom))


def resolve_restore_position(saved, rects, primary, widget_w=0, widget_h=0,
                             margin_x=40, margin_y=80):
    """根据保存位置决定窗口恢复位置。

    - ``saved``: (x, y) 或 None（无保存位置）。
    - ``rects``: 所有屏幕可用区域 [(left, top, w, h), ...]。
    - ``primary``: 主屏可用区域 (left, top, w, h)。

    规则：保存位置落在任一屏幕内则原位恢复（并夹取到该屏幕内）；
    否则（如副屏已断开）回退到主屏右下角默认位置。
    """
    if saved is not None:
        x, y = int(saved[0]), int(saved[1])
        target = screen_containing(x, y, rects) or primary
        return clamp_point(x, y, target, widget_w, widget_h)

    left, top, w, h = primary
    x = left + w - widget_w - margin_x
    y = top + h - widget_h - margin_y
    return max(left, x), max(top, y)
