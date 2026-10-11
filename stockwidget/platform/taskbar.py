"""Windows taskbar child window, following the SetParent approach used by TrafficMonitor.

No code is loaded into Explorer and none of its windows are resized. The small
native window owns its lifetime independently of Qt, so Explorer can destroy it
and the next poll can recreate it. All coordinates here are physical pixels.
"""

import ctypes
from ctypes import wintypes as w
from dataclasses import dataclass
from functools import lru_cache
import logging
import sys


@dataclass(frozen=True)
class TaskbarArea:
    hwnd: int
    width: int
    height: int
    right: int
    dpi: int = 96
    device_name: str = ""
    monitor_rect: tuple[int, int, int, int] | None = None

    def position(self, width, height, offset=0):
        # Keep clear of the notification area. Offset moves towards the left.
        return max(0, self.right - width - offset), max(0, (self.height - height) // 2)


@lru_cache(maxsize=1)
def _api():
    if sys.platform != "win32":
        raise OSError("任务栏显示仅支持 Windows")
    user = ctypes.WinDLL("user32", use_last_error=True)
    gdi = ctypes.WinDLL("gdi32", use_last_error=True)
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    signatures = {
        "FindWindowW": ([w.LPCWSTR, w.LPCWSTR], w.HWND),
        "FindWindowExW": ([w.HWND, w.HWND, w.LPCWSTR, w.LPCWSTR], w.HWND),
        "GetWindowRect": ([w.HWND, ctypes.POINTER(w.RECT)], w.BOOL),
        "GetCursorPos": ([ctypes.POINTER(w.POINT)], w.BOOL),
        "MonitorFromWindow": ([w.HWND, w.DWORD], w.HANDLE),
        "GetMonitorInfoW": ([w.HANDLE, w.LPVOID], w.BOOL),
        "ScreenToClient": ([w.HWND, ctypes.POINTER(w.POINT)], w.BOOL),
        "SetCapture": ([w.HWND], w.HWND),
        "GetCapture": ([], w.HWND),
        "ReleaseCapture": ([], w.BOOL),
        "GetAsyncKeyState": ([ctypes.c_int], ctypes.c_short),
        "GetClientRect": ([w.HWND, ctypes.POINTER(w.RECT)], w.BOOL),
        "MapWindowPoints": ([w.HWND, w.HWND, ctypes.POINTER(w.POINT), w.UINT], ctypes.c_int),
        "GetDpiForWindow": ([w.HWND], w.UINT),
        "GetWindowDpiAwarenessContext": ([w.HWND], w.HANDLE),
        "SetThreadDpiAwarenessContext": ([w.HANDLE], w.HANDLE),
        "IsWindow": ([w.HWND], w.BOOL),
        "GetParent": ([w.HWND], w.HWND),
        "SetParent": ([w.HWND, w.HWND], w.HWND),
        "GetWindowLongW": ([w.HWND, ctypes.c_int], w.LONG),
        "SetWindowLongW": ([w.HWND, ctypes.c_int, w.LONG], w.LONG),
        "CreateWindowExW": ([w.DWORD, w.LPCWSTR, w.LPCWSTR, w.DWORD,
                             ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                             w.HWND, w.HMENU, w.HINSTANCE, w.LPVOID], w.HWND),
        "DefWindowProcW": ([w.HWND, w.UINT, w.WPARAM, w.LPARAM], ctypes.c_ssize_t),
        "DestroyWindow": ([w.HWND], w.BOOL),
        "ShowWindow": ([w.HWND, ctypes.c_int], w.BOOL),
        "SetWindowPos": ([w.HWND, w.HWND, ctypes.c_int, ctypes.c_int,
                          ctypes.c_int, ctypes.c_int, w.UINT], w.BOOL),
        "RegisterClassW": ([w.LPVOID], w.WORD),
        "UnregisterClassW": ([w.LPCWSTR, w.HINSTANCE], w.BOOL),
        "UpdateLayeredWindow": ([w.HWND, w.HDC, w.LPVOID, w.LPVOID, w.HDC,
                                 w.LPVOID, w.DWORD, w.LPVOID, w.DWORD], w.BOOL),
    }
    for name, (args, result) in signatures.items():
        function = getattr(user, name)
        function.argtypes, function.restype = args, result
    for name, args, result in (
        ("CreateCompatibleDC", [w.HDC], w.HDC),
        ("DeleteDC", [w.HDC], w.BOOL),
        ("CreateDIBSection", [w.HDC, w.LPVOID, w.UINT, ctypes.POINTER(w.LPVOID), w.HANDLE, w.DWORD], w.HBITMAP),
        ("SelectObject", [w.HDC, w.HANDLE], w.HANDLE),
        ("DeleteObject", [w.HANDLE], w.BOOL),
    ):
        function = getattr(gdi, name)
        function.argtypes, function.restype = args, result
    kernel.GetModuleHandleW.argtypes = [w.LPCWSTR]
    kernel.GetModuleHandleW.restype = w.HMODULE
    return user, gdi, kernel


def _taskbar_windows(user):
    primary = user.FindWindowW("Shell_TrayWnd", None)
    windows = [primary] if primary else []
    previous = None
    while hwnd := user.FindWindowExW(None, previous, "Shell_SecondaryTrayWnd", None):
        windows.append(hwnd)
        previous = hwnd
    return windows


def _monitor_info(user, hwnd):
    class MonitorInfo(ctypes.Structure):
        _fields_ = [("size", w.DWORD), ("monitor", w.RECT), ("work", w.RECT),
                    ("flags", w.DWORD), ("device", w.WCHAR * 32)]

    info = MonitorInfo()
    info.size = ctypes.sizeof(info)
    monitor = user.MonitorFromWindow(hwnd, 2)  # MONITOR_DEFAULTTONEAREST
    if not monitor or not user.GetMonitorInfoW(monitor, ctypes.byref(info)):
        raise ctypes.WinError(ctypes.get_last_error())
    return info


def _find_child(user, parent, class_name):
    found = user.FindWindowExW(parent, None, class_name, None)
    if found:
        return found
    child = None
    while child := user.FindWindowExW(parent, child, None, None):
        if found := _find_child(user, child, class_name):
            return found
    return None


def find_taskbar(*, window=None, device_name="", hwnd=None):
    """Select a bar by pointer target, saved display, or the floating HWND.

    Monitor selection stays in native coordinates, including mixed-DPI desktops.
    A missing display/bar falls back to the floating window's display or primary.
    """
    user, _, _ = _api()
    windows = _taskbar_windows(user)
    if not windows:
        return None
    if hwnd is not None:
        if hwnd not in windows:
            return None
    else:
        preferred = device_name or (_monitor_info(user, window).device if window else "")
        hwnd = next((bar for bar in windows if _monitor_info(user, bar).device == preferred), None)
        if hwnd is None and window:
            monitor = user.MonitorFromWindow(window, 2)
            hwnd = next((bar for bar in windows if user.MonitorFromWindow(bar, 2) == monitor), None)
        hwnd = hwnd or windows[0]
    rect = w.RECT()
    if not user.GetClientRect(hwnd, ctypes.byref(rect)):
        raise ctypes.WinError(ctypes.get_last_error())
    width, height = rect.right, rect.bottom
    if height > width:
        raise OSError("请将目标屏幕任务栏设置为横向，再启用任务栏显示")
    tray = _find_child(user, hwnd, "TrayNotifyWnd") or _find_child(user, hwnd, "ClockButton")
    if not tray and hwnd == user.FindWindowW("Shell_TrayWnd", None):
        raise OSError("无法定位系统托盘，任务栏显示将在恢复后自动重试")
    dpi = user.GetDpiForWindow(hwnd) or 96
    right = width - 120 * dpi // 96  # Reserve clock space on secondary bars without a legacy tray HWND.
    if tray:
        if not user.GetWindowRect(tray, ctypes.byref(rect)):
            raise ctypes.WinError(ctypes.get_last_error())
        point = w.POINT(rect.left, rect.top)
        user.MapWindowPoints(None, hwnd, ctypes.byref(point), 1)
        right = point.x - 6 * dpi // 96
    info = _monitor_info(user, hwnd)
    monitor = info.monitor
    return TaskbarArea(hwnd, width, height, max(0, right), dpi, info.device,
                       (monitor.left, monitor.top, monitor.right - monitor.left, monitor.bottom - monitor.top))


def cursor_over_taskbar():
    """Return the bar HWND under the native pointer, or None outside all bars."""
    if sys.platform != "win32":
        return None
    user, _, _ = _api()
    point, rect = w.POINT(), w.RECT()
    if user.GetCursorPos(ctypes.byref(point)):
        for hwnd in _taskbar_windows(user):
            if (user.GetWindowRect(hwnd, ctypes.byref(rect))
                    and rect.left <= point.x < rect.right and rect.top <= point.y < rect.bottom):
                return hwnd
    return None


class NativeTaskbarWindow:
    """A per-pixel alpha native child, fed premultiplied BGRA frames by Qt."""

    def __init__(self, on_pointer, on_context_menu):
        self.user, self.gdi, kernel = _api()
        self.hwnd = None
        self.parent = None
        self._last_frame = None
        self._on_pointer = on_pointer
        self._pointer_down = False
        self._last_pointer_position = None
        self._ignore_release = False
        self._on_context_menu = on_context_menu
        self._instance = kernel.GetModuleHandleW(None)
        self._class_name = f"StockWidgetTaskbar_{id(self)}"
        callback_type = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, w.HWND, w.UINT, w.WPARAM, w.LPARAM)
        self._callback = callback_type(self._wndproc)

        class WindowClass(ctypes.Structure):
            _fields_ = [("style", w.UINT), ("proc", callback_type),
                        ("class_extra", ctypes.c_int), ("window_extra", ctypes.c_int),
                        ("instance", w.HINSTANCE), ("icon", w.HICON),
                        ("cursor", w.HANDLE), ("background", w.HBRUSH),
                        ("menu", w.LPCWSTR), ("name", w.LPCWSTR)]

        wc = WindowClass()
        wc.style = 0x0008  # CS_DBLCLKS: receive WM_LBUTTONDBLCLK
        wc.proc, wc.instance, wc.name = self._callback, self._instance, self._class_name
        if not self.user.RegisterClassW(ctypes.byref(wc)):
            raise ctypes.WinError(ctypes.get_last_error())

    def _wndproc(self, hwnd, message, wp, lp):
        try:
            x, y = ctypes.c_short(lp & 0xFFFF).value, ctypes.c_short((lp >> 16) & 0xFFFF).value
            if message == 0x21:  # WM_MOUSEACTIVATE
                return 3  # MA_NOACTIVATE
            if message == 0x201:  # WM_LBUTTONDOWN
                self._pointer_down = True
                self._last_pointer_position = (x, y)
                self._ignore_release = False
                self.user.SetCapture(hwnd)
                self._on_pointer("press", x, y)
                return 0
            if message == 0x200 and self._pointer_down:  # WM_MOUSEMOVE
                if wp & 1:  # MK_LBUTTON; capture continues outside the taskbar
                    self._last_pointer_position = (x, y)
                    self._on_pointer("move", x, y)
                return 0
            if message == 0x202:  # WM_LBUTTONUP
                self.release_pointer()
                if not self._ignore_release:
                    self._on_pointer("release", x, y)
                self._ignore_release = False
                return 0
            if message == 0x203:  # WM_LBUTTONDBLCLK
                self.release_pointer()
                self._ignore_release = True
                self._on_pointer("double_click", x, y)
                return 0
            if message in (0x1F, 0x215) and self._pointer_down:  # CANCELMODE / CAPTURECHANGED
                self.release_pointer()
                self._on_pointer("cancel", x, y)
                return 0
            if message == 0x100 and wp == 0x1B and self._pointer_down:  # Escape
                self.release_pointer()
                self._on_pointer("cancel", x, y)
                return 0
            if message == 0x7B:  # WM_CONTEXTMENU
                self._on_context_menu()
                return 0
            if message == 0x82 and hwnd == self.hwnd:  # WM_NCDESTROY
                self.hwnd = None
                self._last_frame = None
        except Exception:
            logging.getLogger(__name__).exception("任务栏窗口事件处理失败")
        return self.user.DefWindowProcW(hwnd, message, wp, lp)

    def _attach(self, area):
        if self.hwnd and self.user.IsWindow(self.hwnd) and self.user.GetParent(self.hwnd) == area.hwnd:
            return
        # Reparent the same input owner when crossing displays. Destroying it
        # would release pointer capture and interrupt the shared drag gesture.
        existing = self.hwnd if self.hwnd and self.user.IsWindow(self.hwnd) else None
        # Match Explorer at creation; avoid SetParent changing Qt's process DPI mode.
        context = self.user.GetWindowDpiAwarenessContext(area.hwnd)
        previous = self.user.SetThreadDpiAwarenessContext(context)
        try:
            if not existing:
                self.hwnd = self.user.CreateWindowExW(
                    0x08080080, self._class_name, "StockWidget 任务栏行情", 0x80000000,
                    0, 0, 1, 1, None, None, self._instance, None,
                )  # NOACTIVATE | LAYERED | TOOLWINDOW; POPUP, initially hidden
                if not self.hwnd:
                    raise ctypes.WinError(ctypes.get_last_error())
                style = self.user.GetWindowLongW(self.hwnd, -16)
                self.user.SetWindowLongW(self.hwnd, -16, (style & ~0x80000000) | 0x40000000)
            ctypes.set_last_error(0)
            old_parent = self.user.SetParent(self.hwnd, area.hwnd)
            error = ctypes.get_last_error()
            if not old_parent and error:
                raise ctypes.WinError(error)
            if self.user.GetParent(self.hwnd) != area.hwnd:
                raise OSError("无法嵌入任务栏")
            self.parent = area.hwnd
            self._last_frame = None
        except Exception:
            self.hide()
            raise
        finally:
            if previous:
                self.user.SetThreadDpiAwarenessContext(previous)

    def present(self, area, width, height, pixels, offset=0, *, click_through=False):
        self._attach(area)
        self.set_click_through(click_through)
        x, y = area.position(width, height, offset)
        if not self.user.SetWindowPos(self.hwnd, 0, x, y, width, height, 0x50):
            raise ctypes.WinError(ctypes.get_last_error())  # NOACTIVATE | SHOWWINDOW
        frame = (width, height, pixels)
        if frame != self._last_frame:
            self._upload(width, height, pixels)
            self._last_frame = frame

    def set_click_through(self, enabled):
        """Layered windows use WS_EX_TRANSPARENT to pass input across processes."""
        if not self.hwnd:
            return
        style = self.user.GetWindowLongW(self.hwnd, -20)  # GWL_EXSTYLE
        updated = style | 0x20 if enabled else style & ~0x20  # WS_EX_TRANSPARENT
        if updated != style:
            ctypes.set_last_error(0)
            previous = self.user.SetWindowLongW(self.hwnd, -20, updated)
            error = ctypes.get_last_error()
            if not previous and error:
                raise ctypes.WinError(error)
        if enabled and self._pointer_down:
            # Enabling from a hotkey must also end an existing captured drag.
            self.release_pointer()
            self._on_pointer("cancel", 0, 0)

    def _upload(self, width, height, pixels):
        class BitmapInfo(ctypes.Structure):
            _fields_ = [("size", w.DWORD), ("width", w.LONG), ("height", w.LONG),
                        ("planes", w.WORD), ("bits", w.WORD), ("compression", w.DWORD),
                        ("image_size", w.DWORD), ("xppm", w.LONG), ("yppm", w.LONG),
                        ("colors", w.DWORD), ("important", w.DWORD)]

        info = BitmapInfo(40, width, -height, 1, 32)
        bits = w.LPVOID()
        dc = self.gdi.CreateCompatibleDC(None)
        bitmap = None
        old = None
        try:
            if not dc:
                raise ctypes.WinError(ctypes.get_last_error())
            bitmap = self.gdi.CreateDIBSection(dc, ctypes.byref(info), 0, ctypes.byref(bits), None, 0)
            if not bitmap:
                raise ctypes.WinError(ctypes.get_last_error())
            old = self.gdi.SelectObject(dc, bitmap)
            ctypes.memmove(bits, pixels, len(pixels))
            size = w.SIZE(width, height)
            origin = w.POINT(0, 0)
            blend = (ctypes.c_ubyte * 4)(0, 0, 255, 1)  # AC_SRC_OVER, AC_SRC_ALPHA
            if not self.user.UpdateLayeredWindow(self.hwnd, None, None, ctypes.byref(size), dc,
                                                 ctypes.byref(origin), 0, ctypes.byref(blend), 2):
                raise ctypes.WinError(ctypes.get_last_error())
        finally:
            if old:
                self.gdi.SelectObject(dc, old)
            if bitmap:
                self.gdi.DeleteObject(bitmap)
            if dc:
                self.gdi.DeleteDC(dc)

    def hide(self):
        if self._pointer_down and self.hwnd and self.user.IsWindow(self.hwnd):
            # Keep the input owner alive until release; destroying it would
            # interrupt a drag precisely when the pointer leaves the bar.
            # SW_HIDE also releases capture on Windows, so conceal the frame.
            if self._last_frame:
                width, height, _ = self._last_frame
                self._upload(width, height, bytes(width * height * 4))
                self._last_frame = None
            return
        if self.hwnd and self.user.IsWindow(self.hwnd):
            self.user.DestroyWindow(self.hwnd)
        self.hwnd = self.parent = self._last_frame = None
        self._ignore_release = False

    def release_pointer(self):
        self._pointer_down = False
        if self.hwnd and self.user.GetCapture() == self.hwnd:
            self.user.ReleaseCapture()

    def capture_pointer(self):
        """Continue a float-initiated drag while only the taskbar preview is visible."""
        self._pointer_down = True
        self._last_pointer_position = None
        self._ignore_release = False
        self.user.SetCapture(self.hwnd)

    def poll_pointer(self):
        """Continue an initiated drag even when Explorer owns the foreground.

        SetCapture restricts background windows to their visible area:
        https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-setcapture
        Poll only while a press initiated in our float or taskbar window is held.
        """
        if not self._pointer_down or not self.hwnd:
            return None
        point = w.POINT()
        if not self.user.GetCursorPos(ctypes.byref(point)) or not self.user.ScreenToClient(self.hwnd, ctypes.byref(point)):
            return None
        position = (point.x, point.y)
        if self.user.GetAsyncKeyState(0x1B) & 0x8000:
            kind = "cancel"
        elif not self.user.GetAsyncKeyState(1) & 0x8000:
            kind = "release"
        elif position != self._last_pointer_position:
            self._last_pointer_position = position
            return ("move", *position)
        else:
            return None
        self.release_pointer()
        self._ignore_release = True  # do not dispatch a subsequent native up twice
        return (kind, *position)

    def close(self):
        self.release_pointer()
        self.hide()
        self.user.UnregisterClassW(self._class_name, self._instance)
