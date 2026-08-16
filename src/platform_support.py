# -*- coding: utf-8 -*-
"""
跨平台支持层:统一封装平台相关的能力探测与原生实现，供 UI 层按需调用，避免平台分支散落各处。

职责:
- 能力探测:会话类型(X11/Wayland)、全局快捷键、鼠标穿透、窗口透明度、强制置顶、开机自启。
- 原生实现:开机自启(Win 注册表 / XDG autostart / LaunchAgent)、鼠标穿透(Win WS_EX / X11 XShape)。
- 平台默认值:默认字体、自定义图标支持等。

Linux 的关键差异在于 X11 与 Wayland:
- 全局快捷键:仅 X11 可用(基于 XGrabKey);Wayland 没有客户端级的全局按键抓取协议。
- 鼠标穿透:仅 X11 可用(基于 XShape 输入区域);Wayland 无法让客户端窗口穿透鼠标。
- 窗口整体透明度:仅 X11 可用;Wayland 平台插件不支持设置窗口透明度。
- 开机自启:两者皆可用(XDG autostart .desktop,桌面环境层面实现)。
- 窗口拖动:Wayland 需用 QWindow.startSystemMove() 由合成器接管;X11 可直接 move()。
- 强制置顶:仅 Windows 支持(轮询 raise_);Linux 上 raise_() 受窗口管理器限制不可靠,
  macOS 上轮询 raise_() 会不断抢焦点,故这两个平台禁用该选项。
  (macOS 浮窗置顶由 WidgetPanel 的 Qt.WA_MacAlwaysShowToolWindow 属性保证。)
"""

import ctypes
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
    """开机自启是否可用:Windows / Linux / macOS 均支持。
    macOS 通过 ~/Library/LaunchAgents 下的 LaunchAgent plist 实现。
    """
    return sys.platform in ("win32", "linux", "darwin")


def default_font_family() -> str:
    """平台默认中文字体名:macOS 用苹方，其余用微软雅黑。"""
    return "PingFang SC" if sys.platform == "darwin" else "Microsoft YaHei"


def custom_icon_supported() -> bool:
    """自定义/切换应用图标是否可用（macOS 下图标切换不可用）。"""
    return sys.platform != "darwin"


def unsupported_tooltip(feature: str, suggest_x11: bool = True) -> str:
    """为不支持的功能生成中文提示,用于控件 tooltip。feature: 中文功能名。

    suggest_x11=False 时,在 Wayland 下也不提示"切换到 Xorg"——
    例如强制置顶在 X11 下同样不支持,提示切换反而误导。
    """
    st = session_type()
    if st == "wayland" and suggest_x11:
        return f"{feature}:当前 Wayland 会话不支持,请切换到 Xorg(X11)会话后使用。"
    return f"{feature}:当前平台不支持该功能。"


# ---------------------------------------------------------------------------
# 鼠标穿透（Windows: WS_EX_TRANSPARENT；Linux/X11: XShape 输入区域）
# ---------------------------------------------------------------------------

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
    """Windows:通过 WS_EX_TRANSPARENT 扩展样式实现鼠标穿透。"""
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
    """Linux/X11:通过 XShape 扩展设置输入区域。
    启用 = 清空输入区域(窗口不接收鼠标事件，实现穿透)；关闭 = 恢复默认输入区域(整个窗口)。
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
            # mask 为 None + ShapeSet -> 恢复默认输入区域(整个窗口)
            _x11_xext.XShapeCombineMask(
                dpy, win, ShapeInput, 0, 0, None, ShapeSet)
        _x11_xlib.XSync(dpy, False)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# 开机自启（Windows: 注册表 Run 键；Linux: XDG autostart；macOS: LaunchAgent）
# ---------------------------------------------------------------------------

def _launch_command() -> str:
    """返回开机自启要执行的启动命令（冻结环境为可执行文件，开发环境为解释器 + 入口脚本）。"""
    if getattr(sys, 'frozen', False):
        return f'"{sys.executable}"'
    return f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}"'


def set_start_on_boot(enabled: bool, app_name: str) -> None:
    """启用/禁用开机自启（三平台统一入口，按 sys.platform 分派）。"""
    system = sys.platform
    if system == "win32":
        _start_on_boot_windows(enabled, app_name)
    elif system == "linux":
        _start_on_boot_linux(enabled, app_name)
    elif system == "darwin":
        _start_on_boot_macos(enabled, app_name)


def _start_on_boot_windows(enabled: bool, app_name: str) -> None:
    """Windows:写/删 HKCU Run 注册表键。"""
    try:
        import winreg
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
            if enabled:
                winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, _launch_command())
            else:
                winreg.DeleteValue(key, app_name)
    except Exception:
        pass


def _start_on_boot_linux(enabled: bool, app_name: str) -> None:
    """Linux:写/删 ~/.config/autostart 下的 XDG .desktop 文件。"""
    try:
        autostart_dir = os.path.join(os.path.expanduser("~"), ".config", "autostart")
        desktop_file = os.path.join(autostart_dir, f"{app_name}.desktop")
        if not enabled:
            if os.path.exists(desktop_file):
                os.remove(desktop_file)
            return
        content = (
            "[Desktop Entry]\n"
            "Type=Application\n"
            f"Name={app_name}\n"
            f"Comment={app_name} 透明盯盘浮窗\n"
            f"Exec={_launch_command()}\n"
            "Terminal=false\n"
            "X-GNOME-Autostart-enabled=true\n"
        )
        os.makedirs(autostart_dir, exist_ok=True)
        with open(desktop_file, "w", encoding="utf-8") as f:
            f.write(content)
    except Exception:
        pass


def _start_on_boot_macos(enabled: bool, app_name: str) -> None:
    """macOS:写/删 ~/Library/LaunchAgents 下的 LaunchAgent plist，并 launchctl 加载/卸载。"""
    try:
        import plistlib
        import subprocess
        label = f"com.sbr0574.{app_name}"
        launch_dir = os.path.join(os.path.expanduser("~"), "Library", "LaunchAgents")
        plist_path = os.path.join(launch_dir, f"{label}.plist")
        uid = str(os.getuid())
        if not enabled:
            if os.path.exists(plist_path):
                os.remove(plist_path)
            # 立即卸载(失败忽略,注销后也会自然消失)
            subprocess.run(["launchctl", "bootout", f"gui/{uid}", plist_path],
                           capture_output=True, timeout=10)
            return
        args = [sys.executable] if getattr(sys, 'frozen', False) else \
            [sys.executable, os.path.abspath(sys.argv[0])]
        payload = {
            "Label": label,
            "ProgramArguments": args,
            "WorkingDirectory": os.path.dirname(os.path.abspath(sys.argv[0])),
            "RunAtLoad": True,
            "ProcessType": "Interactive",
        }
        os.makedirs(launch_dir, exist_ok=True)
        with open(plist_path, "wb") as f:
            plistlib.dump(payload, f)
        # 立即加载(若已在运行,launchctl 会报错,忽略即可)
        subprocess.run(["launchctl", "bootstrap", f"gui/{uid}", plist_path],
                       capture_output=True, timeout=10)
    except Exception:
        pass
