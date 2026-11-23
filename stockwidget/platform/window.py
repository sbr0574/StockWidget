"""浮窗原生能力：无激活置顶、鼠标穿透、背景模糊与光标。"""

from ctypes import wintypes
from functools import lru_cache
import ctypes
import ctypes.util
import logging
import sys
import weakref

from PySide6.QtCore import Qt

from stockwidget.platform.capabilities import is_cocoa, is_kde_x11, is_windows_11, is_x11


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


class _ACCENT_POLICY(ctypes.Structure):
    _fields_ = [
        ("AccentState", ctypes.c_int),
        ("AccentFlags", ctypes.c_int),
        ("GradientColor", ctypes.c_uint32),
        ("AnimationId", ctypes.c_int),
    ]


class _WINDOWCOMPOSITIONATTRIBDATA(ctypes.Structure):
    _fields_ = [
        ("Attribute", ctypes.c_int),
        ("Data", ctypes.c_void_p),
        ("SizeOfData", ctypes.c_size_t),
    ]


@lru_cache(maxsize=1)
def _get_windows_blur_apis():
    user32 = get_user32()
    set_wca = user32.SetWindowCompositionAttribute
    set_wca.argtypes = [wintypes.HWND, ctypes.POINTER(_WINDOWCOMPOSITIONATTRIBDATA)]
    set_wca.restype = wintypes.BOOL
    dwmapi = ctypes.WinDLL("dwmapi", use_last_error=True)
    dwm_set_attr = dwmapi.DwmSetWindowAttribute
    dwm_set_attr.argtypes = [wintypes.HWND, wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD]
    dwm_set_attr.restype = ctypes.c_long
    return set_wca, dwm_set_attr


def _apply_blur_windows(widget, enable: bool, corner_radius: int) -> bool:
    from PySide6.QtWidgets import QApplication
    app = QApplication.instance()
    if app is None or app.platformName() != "windows":
        return False
    WCA_ACCENT_POLICY = 19
    ACCENT_DISABLED = 0
    ACCENT_ENABLE_BLURBEHIND = 3
    DWMWA_WINDOW_CORNER_PREFERENCE = 33
    DWMWCP_DONOTROUND = 1
    DWMWCP_ROUND = 2
    DWMWCP_ROUNDSMALL = 3
    try:
        hwnd = int(widget.winId())
        set_wca, dwm_set_attr = _get_windows_blur_apis()
        if is_windows_11():
            pref = (DWMWCP_ROUNDSMALL if corner_radius <= 6 else DWMWCP_ROUND) if enable else DWMWCP_DONOTROUND
            value = ctypes.c_int(pref)
            dwm_set_attr(hwnd, DWMWA_WINDOW_CORNER_PREFERENCE, ctypes.byref(value), ctypes.sizeof(value))
        policy = _ACCENT_POLICY(ACCENT_ENABLE_BLURBEHIND if enable else ACCENT_DISABLED, 0, 0, 0)
        data = _WINDOWCOMPOSITIONATTRIBDATA(
            WCA_ACCENT_POLICY,
            ctypes.cast(ctypes.pointer(policy), ctypes.c_void_p),
            ctypes.sizeof(policy),
        )
        return bool(set_wca(hwnd, ctypes.byref(data)))
    except Exception:
        logging.getLogger(__name__).debug("Windows 背景模糊设置失败", exc_info=True)
        return False


class _NSPoint(ctypes.Structure):
    _fields_ = [("x", ctypes.c_double), ("y", ctypes.c_double)]


class _NSSize(ctypes.Structure):
    _fields_ = [("width", ctypes.c_double), ("height", ctypes.c_double)]


class _NSRect(ctypes.Structure):
    _fields_ = [("origin", _NSPoint), ("size", _NSSize)]


_MACOS_BLUR_VIEWS: weakref.WeakKeyDictionary = weakref.WeakKeyDictionary()


@lru_cache(maxsize=1)
def _macos_blur_bridge():
    try:
        objc = ctypes.CDLL(ctypes.util.find_library("objc"))
        appkit = ctypes.CDLL(ctypes.util.find_library("AppKit"))
        objc.objc_getClass.argtypes = [ctypes.c_char_p]
        objc.objc_getClass.restype = ctypes.c_void_p
        objc.sel_registerName.argtypes = [ctypes.c_char_p]
        objc.sel_registerName.restype = ctypes.c_void_p
        return {
            "objc": objc,
            "appkit": appkit,
            "cls": objc.objc_getClass,
            "sel": objc.sel_registerName,
            "ptr": ctypes.CFUNCTYPE(ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p)(("objc_msgSend", objc)),
            "void": ctypes.CFUNCTYPE(None, ctypes.c_void_p, ctypes.c_void_p)(("objc_msgSend", objc)),
            "void_long": ctypes.CFUNCTYPE(None, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_long)(("objc_msgSend", objc)),
            "void_bool": ctypes.CFUNCTYPE(None, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_bool)(("objc_msgSend", objc)),
            "void_double": ctypes.CFUNCTYPE(None, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_double)(("objc_msgSend", objc)),
            "void_ptr": ctypes.CFUNCTYPE(None, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p)(("objc_msgSend", objc)),
            "init_rect": ctypes.CFUNCTYPE(ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, _NSRect)(("objc_msgSend", objc)),
            "void_rect": ctypes.CFUNCTYPE(None, ctypes.c_void_p, ctypes.c_void_p, _NSRect)(("objc_msgSend", objc)),
            "add_subview": ctypes.CFUNCTYPE(
                None, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_long, ctypes.c_void_p,
            )(("objc_msgSend", objc)),
        }
    except Exception:
        logging.getLogger(__name__).debug("macOS 背景模糊桥接初始化失败", exc_info=True)
        return None


def _apply_blur_macos(widget, enable: bool, corner_radius: int, dark: bool) -> bool:
    bridge = _macos_blur_bridge()
    if bridge is None:
        return False
    NSVisualEffectBlendingModeBehindWindow = 0
    NSVisualEffectStateActive = 1
    NSVisualEffectMaterialPopover = 6
    NSVisualEffectMaterialHUDWindow = 13
    NSViewWidthAndHeightSizable = 18
    NSWindowBelow = -1
    try:
        ns_view = ctypes.c_void_p(int(widget.winId()))
        if not ns_view.value:
            return False
        sel = bridge["sel"]
        ns_window = bridge["ptr"](ns_view, sel(b"window"))
        superview = bridge["ptr"](ns_view, sel(b"superview"))
        existing = _MACOS_BLUR_VIEWS.get(widget)
        if not enable:
            if existing:
                bridge["void"](existing, sel(b"removeFromSuperview"))
                bridge["void"](existing, sel(b"release"))
                _MACOS_BLUR_VIEWS.pop(widget, None)
            if ns_window:
                bridge["void"](ns_window, sel(b"invalidateShadow"))
            return True
        if not ns_window or not superview:
            return False
        if existing and bridge["ptr"](existing, sel(b"superview")) != superview:
            bridge["void"](existing, sel(b"removeFromSuperview"))
            bridge["void"](existing, sel(b"release"))
            _MACOS_BLUR_VIEWS.pop(widget, None)
            existing = None
        rect = _NSRect(_NSPoint(0.0, 0.0), _NSSize(float(max(1, widget.width())), float(max(1, widget.height()))))
        if not existing:
            effect_cls = bridge["cls"](b"NSVisualEffectView")
            if not effect_cls:
                return False
            effect = bridge["ptr"](effect_cls, sel(b"alloc"))
            effect = bridge["init_rect"](effect, sel(b"initWithFrame:"), rect)
            if not effect:
                return False
            bridge["void_long"](effect, sel(b"setBlendingMode:"), NSVisualEffectBlendingModeBehindWindow)
            bridge["void_long"](effect, sel(b"setState:"), NSVisualEffectStateActive)
            bridge["void_long"](effect, sel(b"setAutoresizingMask:"), NSViewWidthAndHeightSizable)
            bridge["void_bool"](effect, sel(b"setWantsLayer:"), True)
            bridge["add_subview"](superview, sel(b"addSubview:positioned:relativeTo:"), effect, NSWindowBelow, ns_view)
            _MACOS_BLUR_VIEWS[widget] = effect
            existing = effect
        else:
            bridge["void_rect"](existing, sel(b"setFrame:"), rect)
        material = NSVisualEffectMaterialHUDWindow if dark else NSVisualEffectMaterialPopover
        bridge["void_long"](existing, sel(b"setMaterial:"), material)
        layer = bridge["ptr"](existing, sel(b"layer"))
        if layer:
            bridge["void_double"](layer, sel(b"setCornerRadius:"), float(max(0, corner_radius)))
            bridge["void_bool"](layer, sel(b"setMasksToBounds:"), True)
        bridge["void_bool"](ns_window, sel(b"setOpaque:"), False)
        clear = bridge["ptr"](bridge["cls"](b"NSColor"), sel(b"clearColor"))
        if clear:
            bridge["void_ptr"](ns_window, sel(b"setBackgroundColor:"), clear)
        bridge["void"](ns_window, sel(b"invalidateShadow"))
        return True
    except Exception:
        logging.getLogger(__name__).debug("macOS 背景模糊设置失败", exc_info=True)
        return False


def _apply_blur_kde_x11(widget, enable: bool) -> bool:
    try:
        lib_name = ctypes.util.find_library("X11") or "libX11.so.6"
        x11 = ctypes.CDLL(lib_name)
        x11.XOpenDisplay.argtypes = [ctypes.c_char_p]
        x11.XOpenDisplay.restype = ctypes.c_void_p
        x11.XCloseDisplay.argtypes = [ctypes.c_void_p]
        x11.XCloseDisplay.restype = ctypes.c_int
        x11.XInternAtom.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_int]
        x11.XInternAtom.restype = ctypes.c_ulong
        x11.XChangeProperty.argtypes = [
            ctypes.c_void_p, ctypes.c_ulong, ctypes.c_ulong, ctypes.c_ulong,
            ctypes.c_int, ctypes.c_int, ctypes.c_void_p, ctypes.c_int,
        ]
        x11.XChangeProperty.restype = ctypes.c_int
        x11.XDeleteProperty.argtypes = [ctypes.c_void_p, ctypes.c_ulong, ctypes.c_ulong]
        x11.XDeleteProperty.restype = ctypes.c_int
        x11.XFlush.argtypes = [ctypes.c_void_p]
        x11.XFlush.restype = ctypes.c_int
        dpy = x11.XOpenDisplay(None)
        if not dpy:
            return False
        try:
            atom = x11.XInternAtom(dpy, b"_KDE_NET_WM_BLUR_BEHIND_REGION", 0)
            if not atom:
                return False
            win_id = int(widget.winId())
            if enable:
                dummy = ctypes.c_ulong(0)
                # 0-length CARDINAL property instructs KWin to blur the full window shape.
                x11.XChangeProperty(dpy, win_id, atom, 6, 32, 0, ctypes.byref(dummy), 0)
            else:
                x11.XDeleteProperty(dpy, win_id, atom)
            x11.XFlush(dpy)
            return True
        finally:
            x11.XCloseDisplay(dpy)
    except Exception:
        logging.getLogger(__name__).debug("Linux KDE X11 背景模糊设置失败", exc_info=True)
        return False


def apply_background_blur(widget, enable: bool, *, corner_radius: int = 8, dark: bool = True) -> bool:
    """为无边框浮窗应用或关闭系统原生背景模糊。"""
    enable = bool(enable)
    if sys.platform == "win32":
        return _apply_blur_windows(widget, enable, corner_radius)
    if is_cocoa():
        return _apply_blur_macos(widget, enable, corner_radius, dark)
    if sys.platform == "linux" and is_kde_x11():
        return _apply_blur_kde_x11(widget, enable)
    return False
