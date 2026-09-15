# -*- coding: utf-8 -*-
"""
跨平台全局快捷键管理器

- Windows: RegisterHotKey 注册，Qt 原生事件过滤器接收 WM_HOTKEY。
- macOS: Carbon RegisterEventHotKey 注册，事件处理器在主线程分发回调。
- Linux/X11: XGrabKey 注册，QSocketNotifier 从同一 Xlib 连接读取按键事件。
  冲突通过 XSync 后的 BadAccess 异步错误判断；失败时释放本次已抓取的组合，
  注册结束后恢复原错误处理器。支持 CapsLock/NumLock 等锁定状态。
- Wayland 及其他不支持的平台: 返回 unsupported。

所有原生调用使用 ctypes，无额外键盘监听依赖。仅预先拦截 Windows 的
Ctrl+Alt+Del 安全序列；其他组合交给平台注册接口判断，不按常用程度禁止。
Windows 和 Linux/X11 允许单独使用 F1–F12，macOS 仍要求搭配修饰键。
字母等其他主键在所有平台上都需要修饰键。
HotkeyResult 区分冲突、无效、保留、不支持和其他失败。
设置页负责展示结果；失败的配置可保留并修改，不在此处弹窗。

用法示例:
    mgr = GlobalHotkeyManager(parent)
    result = mgr.register("Ctrl+Alt+F", callback)   # 返回 HotkeyResult
    if not result:
        print(result.reason)                         # 'conflict' / 'invalid' / ...
    mgr.unregister_all()                             # 注销全部已注册热键
"""

import ctypes
import sys

# 先判断系统类型,再按需 import 平台相关模块
if sys.platform == "win32":
    from ctypes import wintypes
else:
    wintypes = None

from PySide6.QtCore import QAbstractNativeEventFilter, QCoreApplication, QObject, QSocketNotifier
from stockwidget.platform.capabilities import session_type

# ---------------------------------------------------------------------------
# 通用结果类型
# ---------------------------------------------------------------------------

class HotkeyResult:
    """快捷键注册结果。`ok` 为是否成功,`reason` 为失败原因:
    - 'conflict'    : 已被其他程序占用(热键冲突)
    - 'invalid'     : 快捷键无法解析(缺修饰键 / 键不支持)
    - 'reserved'    : 当前平台明确保留的特殊组合（Windows Ctrl+Alt+Del）
    - 'unsupported' : 当前平台暂未实现
    - 'failed'      : 其他系统错误
    """

    __slots__ = ("ok", "reason")

    def __init__(self, ok: bool, reason: str = ""):
        self.ok = bool(ok)
        self.reason = reason

    def __bool__(self) -> bool:
        return self.ok

    def __repr__(self) -> str:
        return f"HotkeyResult(ok={self.ok}, reason={self.reason!r})"


# ---------------------------------------------------------------------------
# Windows: RegisterHotKey
# ---------------------------------------------------------------------------

WM_HOTKEY = 0x0312
ERROR_HOTKEY_ALREADY_REGISTERED = 1409

MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008
MOD_NOREPEAT = 0x4000  # Vista+ 有效:按住组合键时不因自动重复而反复触发

# 常用虚拟键码(VK)映射
_VK_MAP = {
    "f1": 0x70, "f2": 0x71, "f3": 0x72, "f4": 0x73, "f5": 0x74, "f6": 0x75,
    "f7": 0x76, "f8": 0x77, "f9": 0x78, "f10": 0x79, "f11": 0x7A, "f12": 0x7B,
    "space": 0x20, "return": 0x0D, "enter": 0x0D, "tab": 0x09,
    "esc": 0x1B, "escape": 0x1B, "backspace": 0x08,
    "del": 0x2E, "delete": 0x2E, "insert": 0x2D, "home": 0x24, "end": 0x23,
    "pageup": 0x21, "pagedown": 0x22,
    "left": 0x25, "up": 0x26, "right": 0x27, "down": 0x28,
    "minus": 0xBD, "-": 0xBD, "plus": 0xBB, "=": 0xBB,
    "comma": 0xBC, ",": 0xBC, "period": 0xBE, ".": 0xBE,
    "slash": 0xBF, "/": 0xBF,
}

# ---------------------------------------------------------------------------
# macOS: Carbon RegisterEventHotKey（系统推荐的"注册式"全局快捷键）
# ---------------------------------------------------------------------------

# Carbon 修饰键掩码（Events.h）
CARBON_CMD_KEY = 0x0100          # Command(⌘)
CARBON_SHIFT_KEY = 0x0200        # Shift(⇧)
CARBON_OPTION_KEY = 0x0800       # Option(⌥)
CARBON_CONTROL_KEY = 0x1000      # Control(⌃)


def _fourcc(s: str) -> int:
    """四字符码（OSType）→ 整数（大端）。"""
    return (ord(s[0]) << 24) | (ord(s[1]) << 16) | (ord(s[2]) << 8) | ord(s[3])


# Carbon Events 常量（与系统 Events.h 一致）
_K_EVENT_CLASS_KEYBOARD = _fourcc("keyb")        # kEventClassKeyboard
_K_EVENT_HOTKEY_PRESSED = 5                       # kEventHotKeyPressed
_K_EVENT_PARAM_DIRECT_OBJECT = _fourcc("----")    # kEventParamDirectObject
_TYPE_EVENT_HOTKEY_ID = _fourcc("hkid")           # typeEventHotKeyID
_MAC_HOTKEY_SIGNATURE = _fourcc("SWgt")           # 自定义热键签名
_NO_ERR = 0
_EVENT_HOTKEY_EXISTS_ERR = -9878                  # 组合键已被占用
_EVENT_HOTKEY_INVALID_ERR = -9879                 # 无效组合键


# EventHotKeyID / EventTypeSpec（Carbon 结构体）
class _EventHotKeyID(ctypes.Structure):
    _fields_ = [("signature", ctypes.c_uint32), ("id", ctypes.c_uint32)]


class _EventTypeSpec(ctypes.Structure):
    _fields_ = [("eventClass", ctypes.c_uint32), ("eventKind", ctypes.c_uint32)]


# 事件处理器回调类型：OSStatus (*)(EventHandlerCallRef, EventRef, void*)
_EVENT_HANDLER_CALLBACK = ctypes.CFUNCTYPE(
    ctypes.c_int32, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p,
)

# macOS 硬件键码(keycode,US 布局)
_MAC_KEYCODES = {
    "a": 0, "s": 1, "d": 2, "f": 3, "h": 4, "g": 5, "z": 6, "x": 7, "c": 8, "v": 9,
    "b": 11, "q": 12, "w": 13, "e": 14, "r": 15, "y": 16, "t": 17, "1": 18, "2": 19,
    "3": 20, "4": 21, "6": 22, "5": 23, "=": 24, "9": 25, "7": 26, "-": 27, "8": 28,
    "0": 29, "]": 30, "o": 31, "u": 32, "[": 33, "i": 34, "p": 35, "return": 36,
    "l": 37, "j": 38, "'": 39, "k": 40, ";": 41, "\\": 42, ",": 43, "/": 44, "n": 45,
    "m": 46, ".": 47, "tab": 48, "space": 49, "`": 50,
    "delete": 51, "backspace": 51, "esc": 53, "escape": 53,
    "del": 117, "forwarddelete": 117, "insert": 114,
    "home": 115, "end": 119, "pageup": 116, "pagedown": 121,
    "up": 126, "down": 125, "left": 123, "right": 124,
    "f1": 122, "f2": 120, "f3": 99, "f4": 118, "f5": 96, "f6": 97, "f7": 98,
    "f8": 100, "f9": 101, "f10": 109, "f11": 103, "f12": 111,
}


def _load_carbon():
    """加载 Carbon 框架并设置函数签名（仅 macOS）。"""
    if sys.platform != "darwin":
        return None
    try:
        carbon = ctypes.cdll.LoadLibrary(
            "/System/Library/Frameworks/Carbon.framework/Carbon")
        carbon.GetApplicationEventTarget.restype = ctypes.c_void_p
        carbon.GetApplicationEventTarget.argtypes = []

        carbon.RegisterEventHotKey.restype = ctypes.c_int32
        carbon.RegisterEventHotKey.argtypes = [
            ctypes.c_uint32, ctypes.c_uint32, _EventHotKeyID,
            ctypes.c_void_p, ctypes.c_uint32, ctypes.POINTER(ctypes.c_void_p),
        ]
        carbon.UnregisterEventHotKey.restype = ctypes.c_int32
        carbon.UnregisterEventHotKey.argtypes = [ctypes.c_void_p]

        carbon.InstallEventHandler.restype = ctypes.c_int32
        carbon.InstallEventHandler.argtypes = [
            ctypes.c_void_p, _EVENT_HANDLER_CALLBACK, ctypes.c_uint64,
            ctypes.POINTER(_EventTypeSpec), ctypes.c_void_p, ctypes.POINTER(ctypes.c_void_p),
        ]

        carbon.GetEventParameter.restype = ctypes.c_int32
        carbon.GetEventParameter.argtypes = [
            ctypes.c_void_p, ctypes.c_uint32, ctypes.c_uint32,
            ctypes.POINTER(ctypes.c_uint32), ctypes.c_uint64,
            ctypes.POINTER(ctypes.c_uint64), ctypes.c_void_p,
        ]
        return carbon
    except Exception:
        return None


# 模块级持有：Carbon 框架句柄、当前 manager（事件回调经此分发）
_carbon = _load_carbon()
_mac_manager = None

if sys.platform == "darwin":
    @_EVENT_HANDLER_CALLBACK
    def _mac_hotkey_handler(handler_ref, event, user_data):
        """Carbon 热键事件回调：读取 EventHotKeyID 并按 id 分发（主线程执行）。"""
        try:
            hid = _EventHotKeyID()
            _carbon.GetEventParameter(
                event, _K_EVENT_PARAM_DIRECT_OBJECT, _TYPE_EVENT_HOTKEY_ID,
                None, ctypes.sizeof(_EventHotKeyID), None, ctypes.byref(hid),
            )
            mgr = _mac_manager
            if mgr is not None:
                mgr._dispatch(int(hid.id))
        except Exception:
            pass
        return _NO_ERR
else:
    _mac_hotkey_handler = None


# ---------------------------------------------------------------------------
# 快捷键字符串解析（各平台共享拆分逻辑）
# ---------------------------------------------------------------------------

_UNMODIFIED_FUNCTION_KEYS = frozenset(f"f{i}" for i in range(1, 13))


def _split_hotkey(hotkey: str, *, allow_unmodified_function_keys: bool = False):
    """把 'Ctrl+Alt+F' 解析为 (修饰键集合, 主键名);无法解析返回 None。

    修饰键集合元素为规范化名字: ctrl / alt / shift / meta。
    默认要求修饰键；Windows/X11 解析时显式允许无修饰键的 F1–F12。
    """
    if not hotkey:
        return None
    mods = set()
    key = None
    for part in hotkey.lower().split("+"):
        part = part.strip()
        if not part:
            continue
        if part in ("ctrl", "control"):
            mods.add("ctrl")
        elif part == "alt":
            mods.add("alt")
        elif part == "shift":
            mods.add("shift")
        elif part in ("meta", "win", "super"):
            mods.add("meta")
        else:
            if key is not None:
                return None  # 出现多个主键,视为无效
            key = part
    if key is None:
        return None
    if not mods and not (allow_unmodified_function_keys and key in _UNMODIFIED_FUNCTION_KEYS):
        return None
    return mods, key


def _mods_windows(mods: set) -> int:
    value = 0
    if "ctrl" in mods:
        value |= MOD_CONTROL
    if "alt" in mods:
        value |= MOD_ALT
    if "shift" in mods:
        value |= MOD_SHIFT
    if "meta" in mods:
        value |= MOD_WIN
    return value


def _vk_windows(key: str):
    vk = _VK_MAP.get(key)
    if vk is None and len(key) == 1 and key.isalnum():
        vk = ord(key.upper())
    return vk


def _parse_hotkey(hotkey: str):
    """Windows 用解析:返回 (modifiers, vk);无效返回 None。"""
    parts = _split_hotkey(hotkey, allow_unmodified_function_keys=True)
    if parts is None:
        return None
    mods, key = parts
    vk = _vk_windows(key)
    if vk is None:
        return None
    return _mods_windows(mods), vk


def _mods_macos(mods: set) -> int:
    """Qt 修饰键名 → Carbon 修饰键掩码。

    Qt 在 macOS 的键盘映射:
        "ctrl" → Command 键(⌘)   -> cmdKey
        "alt"  → Option 键(⌥)    -> optionKey
        "shift"→ Shift 键(⇧)     -> shiftKey
        "meta" → Control 键(⌃)   -> controlKey
    """
    value = 0
    if "ctrl" in mods:
        value |= CARBON_CMD_KEY
    if "alt" in mods:
        value |= CARBON_OPTION_KEY
    if "shift" in mods:
        value |= CARBON_SHIFT_KEY
    if "meta" in mods:
        value |= CARBON_CONTROL_KEY
    return value


def _parse_hotkey_macos(hotkey: str):
    """macOS 用解析:返回 (carbon_modmask, keycode);无效返回 None。"""
    parts = _split_hotkey(hotkey)
    if parts is None:
        return None
    mods, key = parts
    keycode = _MAC_KEYCODES.get(key)
    if keycode is None:
        return None
    return _mods_macos(mods), keycode


# ---------------------------------------------------------------------------
# 平台明确保留的特殊组合
# ---------------------------------------------------------------------------

def is_reserved(hotkey: str) -> bool:
    """只拦截 Windows 的安全注意序列，不把常用应用快捷键视为系统保留。

    Linux 的窗口管理器绑定、macOS 和 Windows 的其他系统快捷键是否可用，
    由原生注册接口判断。Windows 的限制不套用到其他平台。
    """
    if sys.platform != "win32":
        return False
    parts = _split_hotkey(hotkey)
    if parts is None:
        return False
    mods, key = parts
    return mods == {"ctrl", "alt"} and key in {"del", "delete"}


# ---------------------------------------------------------------------------
# Windows 事件过滤器
# ---------------------------------------------------------------------------

class _WindowsHotkeyEventFilter(QAbstractNativeEventFilter):
    """监听 WM_HOTKEY 消息,按热键 id 分发回调。"""

    def __init__(self, owner):
        super().__init__()
        self._owner = owner

    def nativeEventFilter(self, eventType, message):
        if bytes(eventType) == b"windows_generic_MSG":
            msg = wintypes.MSG.from_address(int(message))
            if msg.message == WM_HOTKEY:
                handled = self._owner._dispatch(int(msg.wParam))
                return handled, 0
        return False, 0


# ---------------------------------------------------------------------------
# Linux/X11: XGrabKey
# ---------------------------------------------------------------------------

# X11 修饰键掩码
X11_SHIFT_MASK = 1 << 0
X11_LOCK_MASK = 1 << 1
X11_CONTROL_MASK = 1 << 2
X11_MOD1_MASK = 1 << 3   # Alt
X11_MOD2_MASK = 1 << 4   # NumLock 所在位
X11_MOD3_MASK = 1 << 5
X11_MOD4_MASK = 1 << 6   # Super/Meta
X11_MOD5_MASK = 1 << 7   # ScrollLock 所在位

# 大小写锁/数字锁等"锁键"修饰位,匹配事件状态时忽略
X11_IGNORE_MASK = X11_LOCK_MASK | X11_MOD2_MASK | X11_MOD5_MASK

X_KEY_PRESS = 2
GrabModeAsync = 1
BadAccess = 10           # 其他程序已抓取同一组合时 XGrabKey 产生的错误码


def _mods_x11(mods: set) -> int:
    value = 0
    if "ctrl" in mods:
        value |= X11_CONTROL_MASK
    if "alt" in mods:
        value |= X11_MOD1_MASK
    if "shift" in mods:
        value |= X11_SHIFT_MASK
    if "meta" in mods:
        value |= X11_MOD4_MASK
    return value


# 主键 -> XStringToKeysym 的规范名称(区分大小写;F 键需大写)
_X11_KEYSYM_NAMES = {
    "space": "space",
    "return": "Return", "enter": "Return",
    "tab": "Tab",
    "esc": "Escape", "escape": "Escape",
    "backspace": "BackSpace",
    "del": "Delete", "delete": "Delete",
    "insert": "Insert",
    "home": "Home", "end": "End",
    "pageup": "Page_Up", "pagedown": "Page_Down",
    "left": "Left", "up": "Up", "right": "Right", "down": "Down",
    "minus": "minus", "-": "minus",
    "plus": "equal", "=": "equal",
    "comma": "comma", ",": "comma",
    "period": "period", ".": "period",
    "slash": "slash", "/": "slash",
    "f1": "F1", "f2": "F2", "f3": "F3", "f4": "F4",
    "f5": "F5", "f6": "F6", "f7": "F7", "f8": "F8",
    "f9": "F9", "f10": "F10", "f11": "F11", "f12": "F12",
}


def _keysym_name(key: str) -> str:
    return _X11_KEYSYM_NAMES.get(key, key)


class _XKeyEvent(ctypes.Structure):
    _fields_ = [
        ("type", ctypes.c_int), ("serial", ctypes.c_ulong),
        ("send_event", ctypes.c_int), ("display", ctypes.c_void_p),
        ("window", ctypes.c_ulong), ("root", ctypes.c_ulong),
        ("subwindow", ctypes.c_ulong), ("time", ctypes.c_ulong),
        ("x", ctypes.c_int), ("y", ctypes.c_int),
        ("x_root", ctypes.c_int), ("y_root", ctypes.c_int),
        ("state", ctypes.c_uint), ("keycode", ctypes.c_uint),
        ("same_screen", ctypes.c_int),
    ]


class _XEvent(ctypes.Union):
    _fields_ = [("type", ctypes.c_int), ("xkey", _XKeyEvent),
                ("pad", ctypes.c_long * 24)]


class _XErrorEvent(ctypes.Structure):
    _fields_ = [
        ("type", ctypes.c_int), ("display", ctypes.c_void_p),
        ("resourceid", ctypes.c_ulong), ("serial", ctypes.c_ulong),
        ("error_code", ctypes.c_ubyte), ("request_code", ctypes.c_ubyte),
        ("minor_code", ctypes.c_ubyte),
    ]


_XErrorHandler = ctypes.CFUNCTYPE(
    ctypes.c_int, ctypes.c_void_p, ctypes.POINTER(_XErrorEvent))


# ---------------------------------------------------------------------------
# 全局快捷键管理器
# ---------------------------------------------------------------------------

class GlobalHotkeyManager(QObject):
    """跨平台全局快捷键管理器。

    - Windows: 官方 RegisterHotKey,热键回调直接进入 Qt 事件循环(主线程)。
    - macOS:   Carbon RegisterEventHotKey,事件经 InstallEventHandler 安装的处理器
               在应用主运行循环(NSApplication)中派发,回调在主线程执行。
    - Linux/X11: XGrabKey 抓取全局组合键,由 QSocketNotifier 读取同一连接的事件。
    - 其他(如 Wayland): register 返回 'unsupported',不影响程序运行。
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._callbacks = {}      # 热键 id -> callback
        self._next_id = 1
        self._filter = None
        # macOS 专用状态（Carbon RegisterEventHotKey）
        self._mac_refs = {}           # hotkey_id -> EventHotKeyRef
        self._mac_handler_ref = None  # EventHandlerRef（首次注册时安装，进程内复用）
        self._mac_handler_installed = False
        # Linux/X11 专用状态
        self._x11_lib = None
        self._x11_display = None
        self._x11_root = 0
        self._x11_grabs = {}      # keycode -> [(core_modmask, callback, [modmask,...])]
        self._x11_notifier = None

    # ----- 公共接口 -----
    def register(self, hotkey: str, callback) -> HotkeyResult:
        """注册全局快捷键。返回 HotkeyResult,冲突/无效/保留/不支持时 ok=False。"""
        # 仅拦截当前平台明确保留的特殊组合，其余以原生注册结果为准。
        if is_reserved(hotkey):
            return HotkeyResult(False, "reserved")
        system = sys.platform
        if system == "win32":
            return self._register_windows(hotkey, callback)
        if system == "darwin":
            return self._register_macos(hotkey, callback)
        if system == "linux":
            if session_type() == "x11":
                return self._register_x11(hotkey, callback)
            return HotkeyResult(False, "unsupported")
        return HotkeyResult(False, "unsupported")

    def unregister_all(self):
        """注销全部已注册的全局快捷键。"""
        system = sys.platform
        if system == "win32":
            self._unregister_all_windows()
        elif system == "darwin":
            self._unregister_all_macos()
        elif system == "linux" and session_type() == "x11":
            self._unregister_all_x11()
        self._callbacks.clear()
        self._next_id = 1

    # ----- Windows 实现 -----
    def _register_windows(self, hotkey: str, callback) -> HotkeyResult:
        parsed = _parse_hotkey(hotkey)
        if parsed is None:
            return HotkeyResult(False, "invalid")
        mods, vk = parsed
        try:
            user32 = ctypes.WinDLL("user32", use_last_error=True)
            if self._filter is None:
                self._filter = _WindowsHotkeyEventFilter(self)
                app = QCoreApplication.instance()
                if app is not None:
                    app.installNativeEventFilter(self._filter)
            hotkey_id = self._next_id
            # hwnd 传 NULL:注册到当前线程,WM_HOTKEY 进入 Qt 事件循环
            ok = user32.RegisterHotKey(None, hotkey_id, mods | MOD_NOREPEAT, vk)
            if not ok:
                if ctypes.get_last_error() == ERROR_HOTKEY_ALREADY_REGISTERED:
                    return HotkeyResult(False, "conflict")
                return HotkeyResult(False, "failed")
            self._callbacks[hotkey_id] = callback
            self._next_id += 1
            return HotkeyResult(True)
        except Exception:
            return HotkeyResult(False, "failed")

    def _unregister_all_windows(self):
        try:
            user32 = ctypes.WinDLL("user32", use_last_error=True)
            for hotkey_id in self._callbacks:
                user32.UnregisterHotKey(None, hotkey_id)
        except Exception:
            pass

    # ----- Linux/X11 实现 -----
    def _ensure_x11(self):
        """打开 X11 显示并初始化函数签名(只做一次)。失败返回 None。"""
        if self._x11_lib is not None:
            return self._x11_lib
        try:
            lib = ctypes.CDLL("libX11.so.6")
            lib.XOpenDisplay.restype = ctypes.c_void_p
            lib.XOpenDisplay.argtypes = [ctypes.c_char_p]
            dpy = lib.XOpenDisplay(None)  # 使用 $DISPLAY
            if not dpy:
                return None
            lib.XDefaultRootWindow.restype = ctypes.c_ulong
            lib.XDefaultRootWindow.argtypes = [ctypes.c_void_p]
            root = lib.XDefaultRootWindow(dpy)
            lib.XKeysymToKeycode.argtypes = [ctypes.c_void_p, ctypes.c_ulong]
            lib.XKeysymToKeycode.restype = ctypes.c_ubyte
            lib.XStringToKeysym.argtypes = [ctypes.c_char_p]
            lib.XStringToKeysym.restype = ctypes.c_ulong
            lib.XGrabKey.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_uint,
                                     ctypes.c_ulong, ctypes.c_int, ctypes.c_int, ctypes.c_int]
            lib.XGrabKey.restype = ctypes.c_int
            lib.XUngrabKey.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_uint, ctypes.c_ulong]
            lib.XSync.argtypes = [ctypes.c_void_p, ctypes.c_int]
            lib.XSetErrorHandler.argtypes = [ctypes.c_void_p]
            lib.XSetErrorHandler.restype = ctypes.c_void_p
            lib.XConnectionNumber.argtypes = [ctypes.c_void_p]
            lib.XConnectionNumber.restype = ctypes.c_int
            lib.XPending.argtypes = [ctypes.c_void_p]
            lib.XPending.restype = ctypes.c_int
            lib.XNextEvent.argtypes = [ctypes.c_void_p, ctypes.POINTER(_XEvent)]
            lib.XCloseDisplay.argtypes = [ctypes.c_void_p]

            # XOpenDisplay 的连接不属于 Qt，Qt 的 nativeEventFilter 收不到它的事件。
            notifier = QSocketNotifier(lib.XConnectionNumber(dpy), QSocketNotifier.Read, self)
            notifier.setEnabled(False)
            notifier.activated.connect(self._read_x11_events)
            self._x11_notifier = notifier
            self.destroyed.connect(lambda _obj=None: lib.XCloseDisplay(dpy))

            self._x11_lib = lib
            self._x11_display = dpy
            self._x11_root = root
            return lib
        except Exception:
            return None

    def _register_x11(self, hotkey: str, callback) -> HotkeyResult:
        parsed = _split_hotkey(hotkey, allow_unmodified_function_keys=True)
        if parsed is None:
            return HotkeyResult(False, "invalid")
        mods, key = parsed
        modmask = _mods_x11(mods)
        try:
            lib = self._ensure_x11()
            if lib is None:
                return HotkeyResult(False, "unsupported")
            keysym = lib.XStringToKeysym(_keysym_name(key).encode())
            if keysym == 0:
                return HotkeyResult(False, "invalid")
            keycode = lib.XKeysymToKeycode(self._x11_display, keysym)
            if keycode == 0:
                return HotkeyResult(False, "invalid")
            if any(core == modmask for core, _, _ in self._x11_grabs.get(keycode, [])):
                return HotkeyResult(False, "conflict")

            # 一次抓取"核心修饰 + 大小写锁/数字锁/滚动锁"的 8 种组合,
            # 保证在 CapsLock/NumLock 等锁定状态下也能触发。
            combos = []
            lock_bits = (X11_LOCK_MASK, X11_MOD2_MASK, X11_MOD5_MASK)
            for i in range(1 << len(lock_bits)):
                extra = 0
                for j, bit in enumerate(lock_bits):
                    if i & (1 << j):
                        extra |= bit
                m = modmask | extra
                combos.append(m)

            # XGrabKey 没有同步的成功/冲突返回值；错误要等 XSync 后读取。
            # 错误处理器是进程全局的，只在注册期间替换，并转发其他连接的错误。
            lib.XSync(self._x11_display, False)
            errors = []
            previous = None

            @_XErrorHandler
            def on_error(display, event):
                if display == self._x11_display and event.contents.request_code == 33:
                    errors.append(event.contents.error_code)
                    return 0
                if previous:
                    return _XErrorHandler(previous)(display, event)
                return 0

            previous = lib.XSetErrorHandler(ctypes.cast(on_error, ctypes.c_void_p))
            try:
                for m in combos:
                    lib.XGrabKey(self._x11_display, keycode, m, self._x11_root,
                                 False, GrabModeAsync, GrabModeAsync)
                lib.XSync(self._x11_display, False)
                if errors:
                    # 某个锁键组合冲突时也要释放已成功抓取的其他组合。
                    for m in combos:
                        lib.XUngrabKey(self._x11_display, keycode, m, self._x11_root)
                    lib.XSync(self._x11_display, False)
                    return HotkeyResult(False, "conflict" if BadAccess in errors else "failed")
            finally:
                lib.XSetErrorHandler(previous)

            self._x11_grabs.setdefault(keycode, []).append((modmask, callback, combos))
            self._x11_notifier.setEnabled(True)
            # XSync 可能已把事件读入 Xlib 缓冲区，不能只等 fd 再次就绪。
            self._read_x11_events()
            return HotkeyResult(True)
        except Exception:
            return HotkeyResult(False, "failed")

    def _read_x11_events(self, *_args):
        while self._x11_lib.XPending(self._x11_display):
            event = _XEvent()
            self._x11_lib.XNextEvent(self._x11_display, ctypes.byref(event))
            if event.type == X_KEY_PRESS:
                self._dispatch_x11(event.xkey.keycode, event.xkey.state)

    def _unregister_all_x11(self):
        try:
            if self._x11_display is None:
                return
            for keycode, entries in self._x11_grabs.items():
                for _core, _callback, combos in entries:
                    for m in combos:
                        self._x11_lib.XUngrabKey(self._x11_display, keycode, m, self._x11_root)
            self._x11_grabs.clear()
            self._x11_lib.XSync(self._x11_display, False)
            self._x11_notifier.setEnabled(False)
            self._read_x11_events()
        except Exception:
            pass

    def _dispatch_x11(self, keycode: int, state: int) -> bool:
        """X11 按键分发:命中已抓取组合时调用回调。返回是否已处理。"""
        entries = self._x11_grabs.get(keycode)
        if not entries:
            return False
        core = state & ~X11_IGNORE_MASK
        matched = False
        for modmask, callback, _combos in entries:
            if core == modmask:
                callback()
                matched = True
        return matched

    # ----- macOS 实现(Carbon RegisterEventHotKey)-----
    def _register_macos(self, hotkey: str, callback) -> HotkeyResult:
        parsed = _parse_hotkey_macos(hotkey)
        if parsed is None:
            return HotkeyResult(False, "invalid")
        modmask, keycode = parsed
        if _carbon is None:
            return HotkeyResult(False, "failed")
        try:
            if not self._mac_handler_installed and not self._install_mac_handler():
                return HotkeyResult(False, "failed")

            global _mac_manager
            _mac_manager = self

            hotkey_id = self._next_id
            hkid = _EventHotKeyID(_MAC_HOTKEY_SIGNATURE, hotkey_id)
            ref = ctypes.c_void_p()
            status = _carbon.RegisterEventHotKey(
                keycode, modmask, hkid,
                _carbon.GetApplicationEventTarget(), 0, ctypes.byref(ref),
            )
            if status != _NO_ERR:
                if status == _EVENT_HOTKEY_EXISTS_ERR:
                    return HotkeyResult(False, "conflict")
                if status == _EVENT_HOTKEY_INVALID_ERR:
                    return HotkeyResult(False, "invalid")
                return HotkeyResult(False, "failed")
            if not ref.value:
                return HotkeyResult(False, "failed")

            self._mac_refs[hotkey_id] = ref.value
            self._callbacks[hotkey_id] = callback
            self._next_id += 1
            return HotkeyResult(True)
        except Exception:
            return HotkeyResult(False, "failed")

    def _install_mac_handler(self) -> bool:
        """安装应用级 Carbon 键盘事件处理器（首次注册时执行一次）。"""
        global _mac_manager
        _mac_manager = self
        try:
            spec = _EventTypeSpec(_K_EVENT_CLASS_KEYBOARD, _K_EVENT_HOTKEY_PRESSED)
            handler_ref = ctypes.c_void_p()
            status = _carbon.InstallEventHandler(
                _carbon.GetApplicationEventTarget(),
                _mac_hotkey_handler,
                1,
                ctypes.byref(spec),
                None,
                ctypes.byref(handler_ref),
            )
            if status != _NO_ERR or not handler_ref.value:
                return False
            self._mac_handler_ref = handler_ref.value
            self._mac_handler_installed = True
            return True
        except Exception:
            return False

    def _unregister_all_macos(self):
        global _mac_manager
        if _mac_manager is self:
            _mac_manager = None
        for ref in self._mac_refs.values():
            try:
                _carbon.UnregisterEventHotKey(ref)
            except Exception:
                pass
        self._mac_refs.clear()

    # ----- 分发 -----
    def _dispatch(self, hotkey_id: int) -> bool:
        callback = self._callbacks.get(hotkey_id)
        if callback is None:
            return False
        callback()
        return True
