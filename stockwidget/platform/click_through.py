# -*- coding: utf-8 -*-
"""鼠标穿透的原生实现（Windows: WS_EX_TRANSPARENT；Linux/X11: XShape 输入区域）。"""

import ctypes
import sys

from stockwidget.platform.capabilities import is_x11

# X11 相关库与显示连接的缓存（延迟初始化，复用同一连接）
_x11_xlib = None
_x11_xext = None
_x11_shape_display = None


def apply_click_through(widget, enable: bool) -> None:
    """开关鼠标穿透。Windows 用扩展样式，Linux/X11 用 XShape，其余平台为 no-op。"""
    system = sys.platform
    if system == "win32":
        _click_through_windows(widget, enable)
    elif system == "linux" and is_x11():
        _click_through_x11(widget, enable)


def _click_through_windows(widget, enable: bool) -> None:
    """Windows：通过 WS_EX_TRANSPARENT 扩展样式实现鼠标穿透。"""
    try:
        GWL_EXSTYLE = -20
        WS_EX_LAYERED = 0x00080000
        WS_EX_TRANSPARENT = 0x00000020
        SWP_NOSIZE = 0x0001
        SWP_NOMOVE = 0x0002
        SWP_NOZORDER = 0x0004
        SWP_FRAMECHANGED = 0x0020

        hwnd = int(widget.winId())
        user32 = ctypes.windll.user32
        exstyle = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        if enable:
            exstyle |= WS_EX_LAYERED | WS_EX_TRANSPARENT
        else:
            exstyle &= ~WS_EX_TRANSPARENT
        user32.SetWindowLongW(hwnd, GWL_EXSTYLE, exstyle)
        user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)
    except Exception:
        pass


def _click_through_x11(widget, enable: bool) -> None:
    """Linux/X11：通过 XShape 扩展设置输入区域。
    启用 = 清空输入区域（窗口不接收鼠标事件，实现穿透）；关闭 = 恢复默认输入区域（整个窗口）。
    """
    global _x11_xlib, _x11_xext, _x11_shape_display
    try:
        if _x11_xext is None:
            _x11_xlib = ctypes.CDLL("libX11.so.6")
            _x11_xext = ctypes.CDLL("libXext.so.6")
            xlib = _x11_xlib
            xext = _x11_xext
            xlib.XOpenDisplay.restype = ctypes.c_void_p
            xlib.XOpenDisplay.argtypes = [ctypes.c_char_p]
            _x11_shape_display = xlib.XOpenDisplay(None)  # 使用 $DISPLAY
            if not _x11_shape_display:
                return
            xext.XShapeCombineRectangles.argtypes = [
                ctypes.c_void_p, ctypes.c_ulong, ctypes.c_int,
                ctypes.c_int, ctypes.c_int, ctypes.c_void_p,
                ctypes.c_int, ctypes.c_int]
            xext.XShapeCombineMask.argtypes = [
                ctypes.c_void_p, ctypes.c_ulong, ctypes.c_int,
                ctypes.c_int, ctypes.c_int, ctypes.c_ulong, ctypes.c_int]
            xlib.XSync.argtypes = [ctypes.c_void_p, ctypes.c_int]
        if _x11_shape_display is None:
            return
        dpy = _x11_shape_display
        win = ctypes.c_ulong(int(widget.winId()))
        ShapeInput = 2  # XInputShape
        ShapeSet = 0
        if enable:
            # 0 个矩形 + ShapeSet -> 输入区域为空 -> 鼠标穿透
            _x11_xext.XShapeCombineRectangles(
                dpy, win, ShapeInput, 0, 0, None, 0, ShapeSet)
        else:
            # mask 为 None + ShapeSet -> 恢复默认输入区域（整个窗口）
            _x11_xext.XShapeCombineMask(
                dpy, win, ShapeInput, 0, 0, None, ShapeSet)
        _x11_xlib.XSync(dpy, False)
    except Exception:
        pass
