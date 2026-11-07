"""浮窗原生能力：Windows 无激活置顶、Windows / Cocoa / X11 鼠标穿透与 macOS 浮窗光标。"""

from ctypes import wintypes
from functools import lru_cache
import ctypes
import logging
import sys

from PySide6.QtCore import Qt

from stockwidget.platform.capabilities import is_cocoa, is_x11


def sync_macos_dock_icon():
    """Normal visible windows have a Dock entry; the tray-only app does not."""
    if not is_cocoa():
        return
    from PySide6.QtWidgets import QApplication
    visible = any(window.isVisible() and window.windowType() == Qt.Window
                  for window in QApplication.topLevelWidgets())
    _set_macos_activation_policy(0 if visible else 1)


def _set_macos_activation_policy(policy):
    try:
        import ctypes.util
        objc = ctypes.CDLL(ctypes.util.find_library("objc"))
        objc.objc_getClass.argtypes = [ctypes.c_char_p]
        objc.objc_getClass.restype = ctypes.c_void_p
        objc.sel_registerName.argtypes = [ctypes.c_char_p]
        objc.sel_registerName.restype = ctypes.c_void_p
        pointer = ctypes.CFUNCTYPE(ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p)(("objc_msgSend", objc))
        ns_app = pointer(objc.objc_getClass(b"NSApplication"), objc.sel_registerName(b"sharedApplication"))
        setter = ctypes.CFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_long)(("objc_msgSend", objc))
        setter(ns_app, objc.sel_registerName(b"setActivationPolicy:"), policy)
    except Exception:
        logging.getLogger(__name__).debug("程序坞图标更新失败", exc_info=True)


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


# macOS 的 Qt::Popup 是不接受 key window 的 QNSPanel。Qt 的
# QCocoaWindow::setWindowCursor() 只在 key / utility 窗口上同步系统光标，
# 所以这类浮窗调用 QWidget.setCursor() 后原生光标仍保持箭头。这里沿用
# Qt 自身的 NSCursor 映射，在 Cocoa 下直接设置系统光标。
_MACOS_CURSOR_NAMES = {
    Qt.ArrowCursor: (b"arrowCursor",),
    Qt.SizeHorCursor: (b"_windowResizeEastWestCursor", b"resizeLeftRightCursor"),
    Qt.SizeVerCursor: (b"_windowResizeNorthSouthCursor", b"resizeUpDownCursor"),
    Qt.SizeFDiagCursor: (b"_windowResizeNorthWestSouthEastCursor",),
    Qt.SizeBDiagCursor: (b"_windowResizeNorthEastSouthWestCursor",),
}
# macOS 15 起可使用公开的 frameResizeCursorFromPosition:inDirections:；
# 位置值与 NSCursorFrameResizeDirectionsAll 来自 AppKit。
_MACOS_FRAME_RESIZE_POSITIONS = {
    Qt.SizeVerCursor: 1,   # NSCursorFrameResizePositionTop
    Qt.SizeHorCursor: 2,   # NSCursorFrameResizePositionLeft
    Qt.SizeFDiagCursor: 3, # NSCursorFrameResizePositionTopLeft
    Qt.SizeBDiagCursor: 9, # NSCursorFrameResizePositionTopRight
}


@lru_cache(maxsize=1)
def _macos_cursor_bridge():
    """解析 NSCursor 单例；保留 AppKit 引用避免框架被卸载。"""
    try:
        import ctypes.util
        objc = ctypes.CDLL(ctypes.util.find_library("objc"))
        appkit = ctypes.CDLL(ctypes.util.find_library("AppKit"))
        objc.objc_getClass.argtypes = [ctypes.c_char_p]
        objc.objc_getClass.restype = ctypes.c_void_p
        objc.sel_registerName.argtypes = [ctypes.c_char_p]
        objc.sel_registerName.restype = ctypes.c_void_p
        send_pointer = ctypes.CFUNCTYPE(ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p)(
            ("objc_msgSend", objc))
        send_bool = ctypes.CFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p)(
            ("objc_msgSend", objc))
        send_void = ctypes.CFUNCTYPE(None, ctypes.c_void_p, ctypes.c_void_p)(
            ("objc_msgSend", objc))
        send_frame_cursor = ctypes.CFUNCTYPE(
            ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_long, ctypes.c_long,
        )(("objc_msgSend", objc))
        ns_cursor = objc.objc_getClass(b"NSCursor")
        if not ns_cursor:
            return None
        responds = objc.sel_registerName(b"respondsToSelector:")
        frame_selector = objc.sel_registerName(b"frameResizeCursorFromPosition:inDirections:")
        has_frame_cursor = send_bool(ns_cursor, responds, frame_selector)
        set_cursor = objc.sel_registerName(b"set")
        cursors = {}
        for shape, names in _MACOS_CURSOR_NAMES.items():
            for name in names:
                selector = objc.sel_registerName(name)
                if not send_bool(ns_cursor, responds, selector):
                    continue
                cursor = send_pointer(ns_cursor, selector)
                if cursor:
                    cursors[shape] = cursor
                break
            if shape not in cursors and has_frame_cursor and shape in _MACOS_FRAME_RESIZE_POSITIONS:
                cursor = send_frame_cursor(ns_cursor, frame_selector,
                                           _MACOS_FRAME_RESIZE_POSITIONS[shape], 3)
                if cursor:
                    cursors[shape] = cursor
        return cursors, lambda cursor: send_void(cursor, set_cursor), appkit
    except Exception:
        logging.getLogger(__name__).debug("macOS 原生光标初始化失败", exc_info=True)
        return None


def set_macos_cursor(shape) -> bool:
    """Cocoa 浮窗直接应用原生光标，返回是否成功。其他平台不处理。"""
    if not is_cocoa():
        return False
    bridge = _macos_cursor_bridge()
    if not bridge:
        return False
    cursor = bridge[0].get(shape)
    if not cursor:
        return False
    try:
        bridge[1](cursor)
    except Exception:
        logging.getLogger(__name__).debug("macOS 原生光标设置失败", exc_info=True)
        return False
    return True


def apply_click_through(widget, enable: bool) -> None:
    """Windows 使用扩展样式；Cocoa / X11 由 Qt 管理原生输入透明状态。"""
    system = sys.platform
    if system == "win32":
        _click_through_windows(widget, enable)
    elif is_cocoa() or (system == "linux" and is_x11()):
        widget.winId()
        handle = widget.windowHandle()
        # xcb 重建原生窗口后可能保留旧输入透明缓存；先清除才能重新设置输入区域。
        if enable and is_x11():
            handle.setFlag(Qt.WindowTransparentForInput, False)
        handle.setFlag(Qt.WindowTransparentForInput, enable)


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
