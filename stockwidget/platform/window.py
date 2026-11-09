"""浮窗原生能力：无激活置顶、鼠标穿透与光标。"""

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
    user32.LoadCursorW.argtypes = [wintypes.HINSTANCE, ctypes.c_void_p]
    user32.LoadCursorW.restype = wintypes.HANDLE
    user32.SetCursor.argtypes = [wintypes.HANDLE]
    user32.SetCursor.restype = wintypes.HANDLE
    return user32


def set_windows_arrow_cursor():
    """Popup 捕获期间，直接恢复不会收到 WM_SETCURSOR 的窗口光标。"""
    user32 = get_user32()
    user32.SetCursor(user32.LoadCursorW(None, ctypes.c_void_p(32512)))


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


# Cocoa 的非 key、非 utility 窗口需直接同步 NSCursor。
_MACOS_CURSOR_NAMES = {
    Qt.ArrowCursor: (b"arrowCursor",),
    Qt.SizeHorCursor: (b"_windowResizeEastWestCursor", b"resizeLeftRightCursor"),
    Qt.SizeVerCursor: (b"_windowResizeNorthSouthCursor", b"resizeUpDownCursor"),
    Qt.SizeFDiagCursor: (b"_windowResizeNorthWestSouthEastCursor",),
    Qt.SizeBDiagCursor: (b"_windowResizeNorthEastSouthWestCursor",),
}
# AppKit 的 NSCursorFrameResizePosition 值，供 macOS 15 起的公开接口使用。
_MACOS_FRAME_RESIZE_POSITIONS = {
    Qt.SizeVerCursor: 1,  # Top
    Qt.SizeHorCursor: 2,  # Left
    Qt.SizeFDiagCursor: 3,  # TopLeft
    Qt.SizeBDiagCursor: 9,  # TopRight
}


@lru_cache(maxsize=1)
def _macos_cursor_bridge():
    """缓存系统光标与调用入口，保留 AppKit 引用。"""
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
    """设置 Cocoa 浮窗的系统光标。"""
    if not is_cocoa():
        return False
    bridge = _macos_cursor_bridge()
    cursor = bridge[0].get(shape) if bridge else None
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
