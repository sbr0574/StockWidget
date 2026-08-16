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
- 强制置顶：仅 Windows 支持（轮询 raise_）；Linux 上 raise_() 受窗口管理器限制不可靠，
  macOS 上轮询 raise_() 会不断抢焦点，故这两个平台禁用该选项。
  （macOS 浮窗置顶由 WidgetPanel 的 Qt.WA_MacAlwaysShowToolWindow 属性保证。）

本模块只做“探测/判断”，不包含原生实现：原生实现见 click_through.py / autostart.py。
"""

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
            if pn == "wayland":
                return "wayland"
            if pn in ("xcb", "offscreen"):
                return "x11"
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


def hotkeys_supported() -> bool:
    """全局快捷键是否可用：Windows / macOS 支持；Linux 仅 X11 支持，Wayland 不支持。"""
    system = sys.platform
    if system in ("win32", "darwin"):
        return True
    if system == "linux":
        return is_x11()
    return False


def click_through_supported() -> bool:
    """鼠标穿透是否可用：Windows 支持；Linux 仅 X11 支持（XShape 输入区域），Wayland 不支持。"""
    system = sys.platform
    if system == "win32":
        return True
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


def start_on_boot_supported() -> bool:
    """开机自启是否可用：Windows / Linux / macOS 均支持。
    macOS 通过 ~/Library/LaunchAgents 下的 LaunchAgent plist 实现。
    """
    return sys.platform in ("win32", "linux", "darwin")


def default_font_family() -> str:
    """平台默认中文字体名：macOS 用苹方，其余用微软雅黑。"""
    return "PingFang SC" if sys.platform == "darwin" else "Microsoft YaHei"


def custom_icon_supported() -> bool:
    """自定义/切换应用图标是否可用（macOS 下图标切换不可用）。"""
    return sys.platform != "darwin"


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
