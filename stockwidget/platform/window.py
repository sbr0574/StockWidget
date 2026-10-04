"""浮窗原生能力：Windows 无激活置顶与 Windows / X11 鼠标穿透。"""

from ctypes import wintypes
from functools import lru_cache
import ctypes
import logging
import sys

from PySide6.QtGui import QGuiApplication

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


def _click_through_x11(widget, enable: bool) -> None:
    """Linux/X11：通过 XShape 扩展设置输入区域。
    启用 = 清空输入区域（窗口不接收鼠标事件，实现穿透）；关闭 = 恢复默认输入区域（整个窗口）。
    """
    global _x11_xlib, _x11_xext, _x11_shape_display
    try:
        if not _x11_shape_display:
            xlib = ctypes.CDLL("libX11.so.6")
            xext = ctypes.CDLL("libXext.so.6")
            xlib.XOpenDisplay.restype = ctypes.c_void_p
            xlib.XOpenDisplay.argtypes = [ctypes.c_char_p]
            xext.XShapeCombineRectangles.argtypes = [
                ctypes.c_void_p, ctypes.c_ulong, ctypes.c_int,
                ctypes.c_int, ctypes.c_int, ctypes.c_void_p,
                ctypes.c_int, ctypes.c_int, ctypes.c_int]
            xext.XShapeCombineRectangles.restype = None
            xext.XShapeCombineMask.argtypes = [
                ctypes.c_void_p, ctypes.c_ulong, ctypes.c_int,
                ctypes.c_int, ctypes.c_int, ctypes.c_ulong, ctypes.c_int]
            xext.XShapeCombineMask.restype = None
            xlib.XSync.argtypes = [ctypes.c_void_p, ctypes.c_int]
            dpy = xlib.XOpenDisplay(None)  # 使用 $DISPLAY；失败时允许下次重试
            if not dpy:
                return
            _x11_xlib, _x11_xext, _x11_shape_display = xlib, xext, dpy
        dpy = _x11_shape_display
        win = ctypes.c_ulong(int(widget.winId()))
        # 窗口由 Qt 的另一条连接创建，先确保服务器已经收到创建/映射请求。
        QGuiApplication.sync()
        ShapeInput = 2  # XInputShape
        ShapeSet = 0
        if enable:
            # 0 个矩形 + ShapeSet -> 输入区域为空 -> 鼠标穿透
            _x11_xext.XShapeCombineRectangles(
                dpy, win, ShapeInput, 0, 0, None, 0, ShapeSet, 0)  # Unsorted
        else:
            # mask 为 None + ShapeSet -> 恢复默认输入区域（整个窗口）
            _x11_xext.XShapeCombineMask(
                dpy, win, ShapeInput, 0, 0, 0, ShapeSet)  # X11 None 是整数 0
        _x11_xlib.XSync(dpy, False)
    except Exception:
        pass
