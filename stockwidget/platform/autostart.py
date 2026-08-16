# -*- coding: utf-8 -*-
"""开机自启的原生实现（Windows: 注册表 Run 键；Linux: XDG autostart；macOS: LaunchAgent）。"""

import os
import sys


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
    """Windows：写/删 HKCU Run 注册表键。"""
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
    """Linux：写/删 ~/.config/autostart 下的 XDG .desktop 文件。"""
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
    """macOS：写/删 ~/Library/LaunchAgents 下的 LaunchAgent plist，并 launchctl 加载/卸载。"""
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
            # 立即卸载（失败忽略，注销后也会自然消失）
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
        # 立即加载（若已在运行，launchctl 会报错，忽略即可）
        subprocess.run(["launchctl", "bootstrap", f"gui/{uid}", plist_path],
                       capture_output=True, timeout=10)
    except Exception:
        pass
