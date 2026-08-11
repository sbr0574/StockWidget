# -*- coding: utf-8 -*-
"""
跨平台能力探测:判断当前会话下各平台相关功能是否可用。

Linux 下的关键差异在于 X11 与 Wayland:
- 全局快捷键:仅 X11 可用(基于 XGrabKey);Wayland 没有客户端级的全局按键抓取协议。
- 鼠标穿透:  仅 X11 可用(基于 XShape 输入区域);Wayland 无法让客户端窗口穿透鼠标。
- 窗口整体透明度:仅 X11 可用;Wayland 平台插件不支持设置窗口透明度。
- 开机自启:  两者皆可用(XDG autostart .desktop,桌面环境层面实现)。
- 窗口拖动:  Wayland 需用 QWindow.startSystemMove() 由合成器接管;
             X11 可直接 move()。
- 强制置顶:  仅 Windows 支持;Linux / macOS 上 raise_() 受窗口管理器/合成器限制,
             效果不可靠,故禁用该选项。
"""

import os
import sys


def session_type() -> str | None:
    """返回当前 Linux 图形会话类型:'wayland' / 'x11' / None(非 Linux 或无法判断)。

    优先按实际 Qt 平台插件判断(wayland / xcb),比环境变量更准确——
    例如在 Wayland 会话中强制以 xcb(XWayland)运行时,应视为 X11。
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
    # 兜底:按环境变量判断
    if os.environ.get("WAYLAND_DISPLAY"):
        return "wayland"
    if os.environ.get("DISPLAY"):
        return "x11"
    return None


def is_wayland() -> bool:
    """是否运行在 Wayland 会话(合成器接管窗口位置,必须用 startSystemMove 拖动)。"""
    return session_type() == "wayland"


def is_x11() -> bool:
    """是否运行在 X11/XWayland 会话。"""
    return session_type() == "x11"


def hotkeys_supported() -> bool:
    """全局快捷键是否可用:Windows / macOS 支持;Linux 仅 X11 支持,Wayland 不支持。"""
    system = sys.platform
    if system in ("win32", "darwin"):
        return True
    if system == "linux":
        return is_x11()
    return False


def click_through_supported() -> bool:
    """鼠标穿透是否可用:Windows 支持;Linux 仅 X11 支持(XShape 输入区域),Wayland 不支持。"""
    system = sys.platform
    if system == "win32":
        return True
    if system == "linux":
        return is_x11()
    return False


def opacity_supported() -> bool:
    """窗口整体透明度是否可用:Wayland 平台插件不支持设置窗口透明度,其余平台可用。"""
    if sys.platform == "linux":
        return not is_wayland()
    return True


def force_top_supported() -> bool:
    """强制置顶是否可用:仅 Windows 支持;Linux / macOS 上 raise_() 不可靠,禁用该选项。"""
    return sys.platform == "win32"


def start_on_boot_supported() -> bool:
    """开机自启是否可用:Windows / Linux 支持;macOS 暂未实现。"""
    system = sys.platform
    if system in ("win32", "linux"):
        return True
    return False


def unsupported_tooltip(feature: str, suggest_x11: bool = True) -> str:
    """为不支持的功能生成中文提示,用于控件 tooltip。feature: 中文功能名。

    suggest_x11=False 时,在 Wayland 下也不提示"切换到 Xorg"——
    例如强制置顶在 X11 下同样不支持,提示切换反而误导。
    """
    st = session_type()
    if st == "wayland" and suggest_x11:
        return f"{feature}:当前 Wayland 会话不支持,请切换到 Xorg(X11)会话后使用。"
    return f"{feature}:当前平台不支持该功能。"
