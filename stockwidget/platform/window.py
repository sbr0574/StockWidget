"""浮窗原生能力：Windows 无激活置顶与 Windows / X11 鼠标穿透。"""

from ctypes import wintypes
from functools import lru_cache
import ctypes
import logging
import sys

from PySide6.QtCore import Qt

from stockwidget.platform.capabilities import is_x11


@lru_cache(maxsize=1)
def get_user32():
    # 延迟加载，其他平台导入此模块时不访问 Windows DLL。
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    user32.GetWindowLongW.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.GetWindowLongW.restype = wintypes.LONG
    user32.SetWindowLongW.argtypes = [wintypes.HWND, ctypes.c_int, wintypes.LONG]
    user32.SetWindowLongW.restype = wintypes.LONG
    user32.SetWindowPos.argtypes = [
        wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int,
        ctypes.c_int, ctypes.c_int, wintypes.UINT,
    ]
    user32.SetWindowPos.restype = wintypes.BOOL
    return user32


def ensure_topmost(widget) -> bool:
    """抬升 Windows 浮窗，保持当前位置、尺寸和前台输入焦点。"""
    if sys.platform != "win32":
        return False
    HWND_TOPMOST = -1
    SWP_NOSIZE = 0x0001
    SWP_NOMOVE = 0x0002
    SWP_NOACTIVATE = 0x0010
    try:
        if not get_user32().SetWindowPos(
            int(widget.winId()), HWND_TOPMOST, 0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        ):
            raise ctypes.WinError(ctypes.get_last_error())
    except OSError:
        logging.getLogger(__name__).debug("原生置顶失败", exc_info=True)
        return False
    return True


def apply_click_through(widget, enable: bool) -> None:
    """Windows 使用扩展样式；X11 让 Qt 管理输入区域及窗口重建。"""
    system = sys.platform
    if system == "win32":
        _click_through_windows(widget, enable)
    elif system == "linux" and is_x11():
        widget.winId()
        widget.windowHandle().setFlag(Qt.WindowTransparentForInput, enable)


def _click_through_windows(widget, enable: bool) -> None:
    """Windows：通过 WS_EX_TRANSPARENT 扩展样式实现鼠标穿透。"""
    try:
        GWL_EXSTYLE = -20
        WS_EX_LAYERED = 0x00080000
        WS_EX_TRANSPARENT = 0x00000020
        SWP_NOSIZE = 0x0001
        SWP_NOMOVE = 0x0002
        SWP_NOZORDER = 0x0004
        SWP_NOACTIVATE = 0x0010
        SWP_FRAMECHANGED = 0x0020

        hwnd = int(widget.winId())
        user32 = get_user32()
        exstyle = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        if enable:
            exstyle |= WS_EX_LAYERED | WS_EX_TRANSPARENT
        else:
            exstyle &= ~WS_EX_TRANSPARENT
        user32.SetWindowLongW(hwnd, GWL_EXSTYLE, exstyle)
        user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER
                            | SWP_NOACTIVATE | SWP_FRAMECHANGED)
    except Exception:
        pass
