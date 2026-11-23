# -*- coding: utf-8 -*-
"""
平台能力探测层：统一封装平台相关的能力探测与平台默认值，供 UI 层按需调用，避免平台分支散落各处。

职责：
- 能力探测：会话类型(X11/Wayland)、全局快捷键、鼠标穿透、窗口透明度、强制置顶、开机自启。
- 平台默认值：默认字体、自定义图标支持等。

Linux 的关键差异在于 X11 与 Wayland：
- 全局快捷键：仅 X11 可用（基于 XGrabKey）；Wayland 没有客户端级的全局按键抓取协议。
- 鼠标穿透：仅 X11 可用（基于 XShape 输入区域）；Wayland 无法让客户端窗口穿透鼠标。
- 窗口整体透明度：仅 X11 可用；Wayland 平台插件不支持设置窗口透明度。
- 开机自启：两者皆可用（XDG autostart .desktop，桌面环境层面实现）。
- 窗口拖动：Wayland 需用 QWindow.startSystemMove() 由合成器接管；X11 可直接 move()。
- 强制置顶：仅 Windows 支持（轮询 SetWindowPos，不激活窗口）；Linux 上 raise_() 受窗口管理器限制不可靠，
  macOS 上轮询 raise_() 会不断抢焦点，故这两个平台禁用该选项。
  （macOS 浮窗使用 Qt.WindowStaysOnTopHint，并保留 Qt.WA_MacAlwaysShowToolWindow。）

本模块只做“探测/判断”，不包含原生实现：原生实现见 window.py / autostart.py。
"""

import ctypes
import ctypes.util
import os
import sys


def session_type() -> str | None:
    """返回当前 Linux 图形会话类型：'wayland' / 'x11' / None（非 Linux 或无法判断）。

    优先按实际 Qt 平台插件判断（wayland / xcb），比环境变量更准确——
    例如在 Wayland 会话中强制以 xcb（XWayland）运行时，应视为 X11。
    """
    if sys.platform != "linux":
        return None
    try:
        from PySide6.QtWidgets import QApplication
        app = QApplication.instance()
        if app is not None:
            pn = app.platformName()
            if pn.startswith("wayland"):
                return "wayland"
            if pn == "xcb":
                return "x11"
            # Headless plugins and other platforms expose no X11 capability.
            return None
    except Exception:
        pass
    # 兜底：按环境变量判断
    if os.environ.get("WAYLAND_DISPLAY"):
        return "wayland"
    if os.environ.get("DISPLAY"):
        return "x11"
    return None


def is_wayland() -> bool:
    """是否运行在 Wayland 会话（合成器接管窗口位置，必须用 startSystemMove 拖动）。"""
    return session_type() == "wayland"


def is_x11() -> bool:
    """是否运行在 X11/XWayland 会话。"""
    return session_type() == "x11"


def is_cocoa() -> bool:
    """macOS 原生 Qt 会话；无头插件不能提供 AppKit 窗口能力。"""
    if sys.platform != "darwin":
        return False
    from PySide6.QtWidgets import QApplication
    app = QApplication.instance()
    return app is not None and app.platformName() == "cocoa"


def is_windows_11() -> bool:
    """Windows 11（Build 22000 及以上）支持 DWM 原生圆角裁剪。"""
    if sys.platform != "win32":
        return False
    try:
        return sys.getwindowsversion().build >= 22000
    except AttributeError:
        return False


def _x11_kwin_blur_supported() -> bool | None:
    """探测 X11 根窗口的 _NET_SUPPORTED 是否声明 _KDE_NET_WM_BLUR_BEHIND_REGION。"""
    try:
        lib_name = ctypes.util.find_library("X11") or "libX11.so.6"
        x11 = ctypes.CDLL(lib_name)
        x11.XOpenDisplay.argtypes = [ctypes.c_char_p]
        x11.XOpenDisplay.restype = ctypes.c_void_p
        x11.XCloseDisplay.argtypes = [ctypes.c_void_p]
        x11.XCloseDisplay.restype = ctypes.c_int
        x11.XDefaultRootWindow.argtypes = [ctypes.c_void_p]
        x11.XDefaultRootWindow.restype = ctypes.c_ulong
        x11.XInternAtom.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_int]
        x11.XInternAtom.restype = ctypes.c_ulong
        x11.XGetWindowProperty.argtypes = [
            ctypes.c_void_p, ctypes.c_ulong, ctypes.c_ulong,
            ctypes.c_long, ctypes.c_long, ctypes.c_int, ctypes.c_ulong,
            ctypes.POINTER(ctypes.c_ulong), ctypes.POINTER(ctypes.c_int),
            ctypes.POINTER(ctypes.c_ulong), ctypes.POINTER(ctypes.c_ulong),
            ctypes.POINTER(ctypes.c_void_p),
        ]
        x11.XGetWindowProperty.restype = ctypes.c_int
        x11.XFree.argtypes = [ctypes.c_void_p]
        x11.XFree.restype = ctypes.c_int
        dpy = x11.XOpenDisplay(None)
        if not dpy:
            return None
        try:
            blur_atom = x11.XInternAtom(dpy, b"_KDE_NET_WM_BLUR_BEHIND_REGION", 1)
            if not blur_atom:
                return False
            supported_atom = x11.XInternAtom(dpy, b"_NET_SUPPORTED", 1)
            if not supported_atom:
                return None
            root = x11.XDefaultRootWindow(dpy)
            actual_type = ctypes.c_ulong()
            actual_format = ctypes.c_int()
            nitems = ctypes.c_ulong()
            bytes_after = ctypes.c_ulong()
            prop = ctypes.c_void_p()
            status = x11.XGetWindowProperty(
                dpy, root, supported_atom, 0, 4096, 0, 4,  # XA_ATOM = 4
                ctypes.byref(actual_type), ctypes.byref(actual_format),
                ctypes.byref(nitems), ctypes.byref(bytes_after), ctypes.byref(prop),
            )
            if status != 0 or not prop.value:
                return None
            try:
                atoms = (ctypes.c_ulong * nitems.value).from_address(prop.value)
                return any(atom == blur_atom for atom in atoms)
            finally:
                x11.XFree(prop)
        finally:
            x11.XCloseDisplay(dpy)
    except Exception:
        return None


def is_kde_x11() -> bool:
    """Linux X11 会话下是否运行在支持背景模糊的 KDE Plasma（KWin）桌面。"""
    if sys.platform != "linux" or not is_x11():
        return False
    supported = _x11_kwin_blur_supported()
    if supported is not None:
        return supported
    desktop = os.environ.get("XDG_CURRENT_DESKTOP", "").lower()
    return "kde" in desktop or bool(os.environ.get("KDE_FULL_SESSION"))


def background_blur_supported() -> bool:
    """窗口背景模糊是否可用：Windows、macOS Cocoa 及 Linux KDE X11 支持。"""
    system = sys.platform
    if system == "win32":
        return True
    if system == "darwin":
        return is_cocoa()
    if system == "linux":
        return is_kde_x11()
    return False


def hotkeys_supported() -> bool:
    """全局快捷键是否可用：Windows / macOS 支持；Linux 仅 X11 支持，Wayland 不支持。"""
    system = sys.platform
    if system in ("win32", "darwin"):
        return True
    if system == "linux":
        return is_x11()
    return False


def click_through_supported() -> bool:
    """Windows、macOS Cocoa 及 Linux X11 支持；Wayland 不支持。"""
    system = sys.platform
    if system == "win32":
        return True
    if system == "darwin":
        return is_cocoa()
    if system == "linux":
        return is_x11()
    return False


def opacity_supported() -> bool:
    """窗口整体透明度是否可用：Wayland 平台插件不支持设置窗口透明度，其余平台可用。"""
    if sys.platform == "linux":
        return not is_wayland()
    return True


def force_top_supported() -> bool:
    """强制置顶是否可用：仅 Windows 支持；Linux / macOS 上 raise_() 不可靠，禁用该选项。"""
    return sys.platform == "win32"


def boundary_check_supported() -> bool:
    """手动边界检测设置仅在 Windows 显示；其他桌面交由窗口管理器约束。"""
    return sys.platform == "win32"


def edge_hide_supported() -> bool:
    """贴边收起需要客户端能定位与调整窗口；Wayland 不提供该能力。"""
    return sys.platform in ("win32", "darwin") or (sys.platform == "linux" and is_x11())


def start_on_boot_supported() -> bool:
    """开机自启是否可用：Windows / Linux / macOS 均支持。
    macOS 通过 ~/Library/LaunchAgents 下的 LaunchAgent plist 实现。
    """
    return sys.platform in ("win32", "linux", "darwin")


def default_font_family() -> str:
    """平台默认中文字体名：macOS 用苹方，其余用微软雅黑。"""
    return "PingFang SC" if sys.platform == "darwin" else "Microsoft YaHei"


def tray_click_toggles() -> bool:
    """托盘单击是否切换显示/隐藏：仅 Windows 支持；macOS/Linux 单击直接弹出菜单。"""
    return sys.platform == "win32"


def unsupported_tooltip(feature: str, suggest_x11: bool = True) -> str:
    """为不支持的功能生成中文提示，用于控件 tooltip。feature: 中文功能名。

    suggest_x11=False 时，在 Wayland 下也不提示“切换到 Xorg”——
    例如强制置顶在 X11 下同样不支持，提示切换反而误导。
    """
    st = session_type()
    if st == "wayland" and suggest_x11:
        return f"{feature}:当前 Wayland 会话不支持，请切换到 Xorg（X11）会话后使用。"
    return f"{feature}:当前平台不支持该功能。"


def background_blur_unsupported_tooltip() -> str:
    """为不支持窗口背景模糊的桌面或会话生成中文提示。"""
    if sys.platform == "linux":
        if session_type() == "wayland":
            return "背景模糊：当前 Wayland 会话不支持，请在 KDE Plasma 的 Xorg（X11）会话下使用。"
        return "背景模糊：当前桌面环境不支持，仅 KDE Plasma（X11）支持背景模糊。"
    return unsupported_tooltip("背景模糊", suggest_x11=False)
