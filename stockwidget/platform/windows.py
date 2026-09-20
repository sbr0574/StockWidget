"""Windows 窗口 API：显式声明指针宽度，避免 64 位 HWND 被截断。"""

import ctypes
from ctypes import wintypes
from functools import lru_cache
import logging
import sys


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
