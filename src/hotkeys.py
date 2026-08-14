# -*- coding: utf-8 -*-
"""
跨平台全局快捷键管理器

- Windows: 基于官方 RegisterHotKey API + QAbstractNativeEventFilter 监听 WM_HOTKEY。
           相比第三方 keyboard 库,零额外依赖、更稳定、不易被杀毒软件误报。
           注册失败时可区分"被其他程序占用"(ERROR_HOTKEY_ALREADY_REGISTERED),
           冲突时 RegisterHotKey 不会抢占,不会影响其他应用。
- macOS:   基于 Quartz CGEventTap(通过 ctypes 直接调用,零额外依赖)。
           事件截获经 CFRunLoop 源挂到 Qt 主循环,回调在主线程执行,可靠触发。
           (注:Carbon RegisterEventHotKey 在 Qt/Cocoa 应用中虽能注册成功,
            但事件不会派发,故改用 CGEventTap。)
           首次使用需在"系统设置 → 隐私与安全性 → 辅助功能"中授权本程序,
           未授权时注册返回 reason='permission' 并自动弹出授权对话框。
- 其他:    静默返回不支持(unsupported),不影响程序运行。
- 保护:    `register()` 前先经 `is_reserved()` 黑名单拦截系统/通用快捷键
           (如 Ctrl+C/V/A、Alt+Tab、Ctrl+Alt+Del、Win 组合等),返回
           reason='reserved',避免注册后影响其他应用正常使用。

用法示例:
    mgr = GlobalHotkeyManager(parent)
    result = mgr.register("Ctrl+Alt+F", callback)   # 返回 HotkeyResult
    if not result:
        print(result.reason)                         # 'conflict' / 'invalid' / ...
    mgr.unregister_all()                             # 注销全部已注册热键
"""

import ctypes
import struct
import sys

# 先判断系统类型,再按需 import 平台相关模块
if sys.platform == "win32":
    from ctypes import wintypes
else:
    wintypes = None

from PySide6.QtCore import QAbstractNativeEventFilter, QCoreApplication, QObject
from src.platform_support import session_type

# ---------------------------------------------------------------------------
# 通用结果类型
# ---------------------------------------------------------------------------

class HotkeyResult:
    """快捷键注册结果。`ok` 为是否成功,`reason` 为失败原因:
    - 'conflict'    : 已被其他程序占用(热键冲突)
    - 'invalid'     : 快捷键无法解析(缺修饰键 / 键不支持)
    - 'reserved'    : 系统/通用快捷键,为避免影响其他应用而禁止注册
    - 'unsupported' : 当前平台暂未实现
    - 'permission'  : 缺少系统权限(如 macOS 辅助功能授权),需用户手动授权
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
# macOS: Quartz CGEventTap
# ---------------------------------------------------------------------------

# CGEvent 事件类型
CG_KEY_DOWN = 10                 # kCGEventKeyDown
CG_FLAGS_CHANGED = 12            # kCGEventFlagsChanged
CG_TAP_DISABLED_TIMEOUT = 0xFFFFFFFE  # kCGEventTapDisabledByTimeout
CG_TAP_DISABLED_USER = 0xFFFFFFFF     # kCGEventTapDisabledByUserInput

# CGEvent 修饰键掩码
CG_CMD = 0x00100000              # Command(⌘)
CG_OPT = 0x00080000              # Option(⌥)
CG_CTRL = 0x00040000             # Control(⌃)
CG_SHIFT = 0x00020000            # Shift(⇧)
_CG_MOD_MASK = CG_CMD | CG_OPT | CG_CTRL | CG_SHIFT

# 事件截获相关常量
CG_TAP_PLACE_HID = 0             # kCGHIDEventTap:全局 HID 事件
CG_TAP_POINT_HEAD = 0            # kCGHeadInsertEventTap
CG_TAP_LISTEN_ONLY = 1           # 只监听不拦截,不影响其他应用
CG_FIELD_KEYCODE = 9             # kCGKeyboardEventKeycode
CG_FIELD_AUTOREPEAT = 8          # kCGKeyboardEventAutorepeat

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


def _load_cg():
    """加载 macOS CoreGraphics 并设置函数签名(仅 Darwin 平台)。"""
    if sys.platform != "darwin":
        return None
    try:
        cg = ctypes.cdll.LoadLibrary(
            "/System/Library/Frameworks/CoreGraphics.framework/CoreGraphics")
        cg.CGEventTapCreate.restype = ctypes.c_void_p
        cg.CGEventTapEnable.argtypes = [ctypes.c_void_p, ctypes.c_bool]
        cg.CGEventTapIsEnabled.restype = ctypes.c_bool
        cg.CGEventTapIsEnabled.argtypes = [ctypes.c_void_p]
        cg.CGEventGetFlags.restype = ctypes.c_uint64
        cg.CGEventGetFlags.argtypes = [ctypes.c_void_p]
        cg.CGEventGetIntegerValueField.restype = ctypes.c_int64
        cg.CGEventGetIntegerValueField.argtypes = [ctypes.c_void_p, ctypes.c_int32]
        cg.CGPreflightListenEventAccess.restype = ctypes.c_bool
        cg.CGRequestListenEventAccess.restype = ctypes.c_bool
        return cg
    except Exception:
        return None


def _load_cf():
    """加载 macOS CoreFoundation(用于把事件截获挂到主 RunLoop)。"""
    if sys.platform != "darwin":
        return None
    try:
        cf = ctypes.cdll.LoadLibrary(
            "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation")
        cf.CFMachPortCreateRunLoopSource.restype = ctypes.c_void_p
        cf.CFMachPortCreateRunLoopSource.argtypes = [
            ctypes.c_void_p, ctypes.c_void_p, ctypes.c_long,
        ]
        cf.CFRunLoopGetMain.restype = ctypes.c_void_p
        cf.CFRunLoopAddSource.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p]
        cf.CFRunLoopRemoveSource.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p]
        cf.CFRelease.argtypes = [ctypes.c_void_p]
        return cf
    except Exception:
        return None


def _kCFRunLoopCommonModes(cf):
    """取 kCFRunLoopCommonModes 常量(CFStringRef)。"""
    try:
        return ctypes.c_void_p.in_dll(cf, "kCFRunLoopCommonModes").value
    except Exception:
        return None


def _mac_is_trusted() -> bool:
    """当前进程是否已获 macOS 辅助功能(监听键盘输入)权限。"""
    try:
        ax = ctypes.cdll.LoadLibrary(
            "/System/Library/Frameworks/ApplicationServices.framework/ApplicationServices")
        ax.AXIsProcessTrusted.restype = ctypes.c_bool
        ax.AXIsProcessTrusted.argtypes = []
        return bool(ax.AXIsProcessTrusted())
    except Exception:
        return False


# 模块级持有:事件截获(CFMachPortRef)、RunLoop 源、回调引用(防 GC)、当前 manager
_mac_cg = _load_cg()
_mac_cf = _load_cf()
_mac_tap = None
_mac_source = None
_mac_cb_ref = None
_mac_manager = None


# ---------------------------------------------------------------------------
# 快捷键字符串解析(两种平台共享拆分逻辑)
# ---------------------------------------------------------------------------

def _split_hotkey(hotkey: str):
    """把 'Ctrl+Alt+F' 解析为 (修饰键集合, 主键名);无法解析返回 None。

    修饰键集合元素为规范化名字: ctrl / alt / shift / meta。
    RegisterHotKey / RegisterEventHotKey 都要求至少一个修饰键,否则视为无效。
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
    if key is None or not mods:
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
    parts = _split_hotkey(hotkey)
    if parts is None:
        return None
    mods, key = parts
    vk = _vk_windows(key)
    if vk is None:
        return None
    return _mods_windows(mods), vk


def _mods_macos(mods: set) -> int:
    """Qt 修饰键名 → CGEvent 修饰键掩码。

    Qt 在 macOS 的键盘映射:
        "ctrl" → Command 键(⌘)   -> kCGEventFlagMaskCommand
        "alt"  → Option 键(⌥)    -> kCGEventFlagMaskAlternate
        "shift"→ Shift 键(⇧)     -> kCGEventFlagMaskShift
        "meta" → Control 键(⌃)   -> kCGEventFlagMaskControl
    """
    value = 0
    if "ctrl" in mods:
        value |= CG_CMD
    if "alt" in mods:
        value |= CG_OPT
    if "shift" in mods:
        value |= CG_SHIFT
    if "meta" in mods:
        value |= CG_CTRL
    return value


def _parse_hotkey_macos(hotkey: str):
    """macOS 用解析:返回 (cg_modmask, keycode);无效返回 None。"""
    parts = _split_hotkey(hotkey)
    if parts is None:
        return None
    mods, key = parts
    keycode = _MAC_KEYCODES.get(key)
    if keycode is None:
        return None
    return _mods_macos(mods), keycode


# ---------------------------------------------------------------------------
# 保留组合黑名单(避免影响其他应用)
# ---------------------------------------------------------------------------

# 系统硬保留/特殊组合(无论是否有程序占用都禁止注册)
# 以 (修饰键集合, 主键) 元组匹配,确保 Alt+F4 与 Ctrl+Esc 等精确定位
_RESERVED_EXACT = {
    (frozenset({"ctrl", "alt"}), "del"),    # Ctrl+Alt+Del 安全注意序列
    (frozenset({"alt"}), "tab"),            # 切换窗口
    (frozenset({"alt"}), "esc"),            # 切换窗口
    (frozenset({"alt"}), "space"),          # 窗口系统菜单
    (frozenset({"alt"}), "f4"),             # 关闭窗口
    (frozenset({"ctrl"}), "esc"),           # 开始菜单
    (frozenset({"ctrl", "shift"}), "esc"),  # 任务管理器
}


def _is_generic_key(key: str) -> bool:
    """是否为通用快捷键常用的主键:单字母 / 数字 / 空格。"""
    return (len(key) == 1 and (key.isalpha() or key.isdigit())) or key == "space"


def is_reserved(hotkey: str) -> bool:
    """判断该快捷键是否为"系统保留/通用快捷键",为避免影响其他应用应禁止注册。

    规则(跨平台,在 Windows 与 macOS 上均生效):
    - 含 Win/Meta 键的组合(系统级,且多为系统保留)
    - 系统硬保留组合(Ctrl+Alt+Del、Alt+Tab、Alt+F4、Ctrl+Esc、Ctrl+Shift+Esc 等)
    - 恰好一个修饰键(Ctrl 或 Alt)+ 通用主键(单字母/数字/空格):
      如 Ctrl+C/V/X/A/S、Ctrl+Space(输入法切换)、Alt+F4 等
    - Ctrl+Shift + 通用主键:如 Ctrl+Shift+S/T/Z(另存为/恢复标签/撤销)
    - Ctrl+Alt + 主键:放行(这类组合应用很少占用,是安全的自定义空间)
    """
    parts = _split_hotkey(hotkey)
    if parts is None:
        return False
    mods, key = parts
    if "meta" in mods:
        return True
    if (frozenset(mods), key) in _RESERVED_EXACT:
        return True
    if len(mods) == 1 and ("ctrl" in mods or "alt" in mods):
        return _is_generic_key(key)
    if mods == {"ctrl", "shift"}:
        return _is_generic_key(key)
    return False


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
# macOS: CGEventTap 事件处理器(模块级单例)
# ---------------------------------------------------------------------------

if sys.platform == "darwin":
    _CGEventTapCallback = ctypes.CFUNCTYPE(
        ctypes.c_void_p,   # CGEventRef 返回值
        ctypes.c_void_p,   # CGEventTapProxy
        ctypes.c_uint32,   # CGEventType
        ctypes.c_void_p,   # CGEventRef
        ctypes.c_void_p,   # void* userInfo
    )

    _mac_cg.CGEventTapCreate.argtypes = [
        ctypes.c_uint32, ctypes.c_uint32, ctypes.c_uint32,
        ctypes.c_uint64, _CGEventTapCallback, ctypes.c_void_p,
    ]

    @_CGEventTapCallback
    def _mac_tap_callback(proxy, event_type, event, user_info):
        """CGEventTap 回调:命中已注册的全局快捷键时分发。
        只监听不拦截(透传事件),不影响其他应用;经主 RunLoop 在主线程执行。
        """
        try:
            if event_type == CG_KEY_DOWN:
                # 自动重复的按键不重复触发,避免按住时反复切换
                if _mac_cg.CGEventGetIntegerValueField(event, CG_FIELD_AUTOREPEAT):
                    return event
                keycode = _mac_cg.CGEventGetIntegerValueField(event, CG_FIELD_KEYCODE)
                flags = _mac_cg.CGEventGetFlags(event) & _CG_MOD_MASK
                mgr = _mac_manager
                if mgr is not None:
                    hid = mgr._mac_grabs.get((int(keycode), int(flags)))
                    if hid is not None:
                        mgr._dispatch(hid)
            elif event_type in (CG_TAP_DISABLED_TIMEOUT, CG_TAP_DISABLED_USER):
                # 被系统停用时重新启用
                if _mac_tap is not None:
                    _mac_cg.CGEventTapEnable(_mac_tap, True)
        except Exception:
            pass
        return event

    def _mac_ensure_tap() -> bool:
        """首次注册时创建 CGEventTap 并挂到主 RunLoop。已创建则直接返回 True。"""
        global _mac_tap, _mac_source, _mac_cb_ref
        if _mac_tap is not None:
            return True
        if _mac_cg is None:
            return False
        try:
            events = (1 << CG_KEY_DOWN) | (1 << CG_FLAGS_CHANGED)
            cb = _CGEventTapCallback(_mac_tap_callback)
            _mac_cb_ref = cb  # 持有引用,防止回调被 GC 回收
            tap = _mac_cg.CGEventTapCreate(
                CG_TAP_PLACE_HID, CG_TAP_POINT_HEAD, CG_TAP_LISTEN_ONLY,
                events, cb, None,
            )
            if not tap:
                return False
            _mac_tap = tap
            _mac_cg.CGEventTapEnable(tap, True)
            # 挂到主 RunLoop(Qt 的 NSApplication 驱动主 RunLoop,回调在主线程)
            if _mac_cf is not None:
                source = _mac_cf.CFMachPortCreateRunLoopSource(None, tap, 0)
                if source:
                    _mac_source = source
                    rl = _mac_cf.CFRunLoopGetMain()
                    mode = _kCFRunLoopCommonModes(_mac_cf)
                    if rl and mode:
                        _mac_cf.CFRunLoopAddSource(rl, source, mode)
            return True
        except Exception:
            return False
else:
    _CGEventTapCallback = None
    _mac_tap_callback = None
    _mac_ensure_tap = None


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

XCB_KEY_PRESS = 2        # xcb_key_press_event_t 的 response_type
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


class _X11HotkeyEventFilter(QAbstractNativeEventFilter):
    """监听 X11 键盘事件,命中已抓取的全局快捷键时分发回调。"""

    def __init__(self, owner):
        super().__init__()
        self._owner = owner

    def nativeEventFilter(self, eventType, message):
        try:
            if bytes(eventType) != b"xcb_generic_event_t":
                return False, 0
            ptr = int(message)
            if ptr == 0:
                return False, 0
            data = ctypes.string_at(ptr, 32)
            if data[0] == XCB_KEY_PRESS:
                keycode = data[1]
                state = struct.unpack_from("<H", data, 28)[0]
                if self._owner._dispatch_x11(keycode, state):
                    return True, 0
        except Exception:
            pass
        return False, 0


# ---------------------------------------------------------------------------
# 全局快捷键管理器
# ---------------------------------------------------------------------------

class GlobalHotkeyManager(QObject):
    """跨平台全局快捷键管理器。

    - Windows: 官方 RegisterHotKey,热键回调直接进入 Qt 事件循环(主线程)。
    - macOS:   Carbon RegisterEventHotKey,回调经 Carbon 事件循环进入主线程。
    - Linux/X11: XGrabKey 抓取全局组合键,经 X11 事件过滤器分发。
    - 其他(如 Wayland): register 返回 'unsupported',不影响程序运行。
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._callbacks = {}      # 热键 id -> callback
        self._next_id = 1
        self._filter = None
        # macOS 专用状态
        self._mac_grabs = {}      # (keycode, modmask) -> hotkey_id
        # Linux/X11 专用状态
        self._x11_lib = None
        self._x11_display = None
        self._x11_root = 0
        self._x11_grabs = {}      # keycode -> [(core_modmask, callback, [modmask,...])]
        self._x11_filter = None

    # ----- 公共接口 -----
    def register(self, hotkey: str, callback) -> HotkeyResult:
        """注册全局快捷键。返回 HotkeyResult,冲突/无效/保留/不支持时 ok=False。"""
        # 先拦截系统/通用快捷键,避免注册后影响其他应用正常使用
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

            # 安装空错误处理器,吞掉 BadAccess(组合键已被其他程序占用)等错误,
            # 避免 X 默认错误处理器终止整个进程。
            XErrorHandler = ctypes.CFUNCTYPE(
                ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p)

            @XErrorHandler
            def _err_handler(display, event):
                return 0

            self._x11_error_handler_cb = _err_handler
            lib.XSetErrorHandler.argtypes = [XErrorHandler]
            lib.XSetErrorHandler.restype = XErrorHandler
            self._x11_error_handler = lib.XSetErrorHandler(_err_handler)

            self._x11_lib = lib
            self._x11_display = dpy
            self._x11_root = root
            return lib
        except Exception:
            return None

    def _register_x11(self, hotkey: str, callback) -> HotkeyResult:
        parsed = _split_hotkey(hotkey)
        if parsed is None:
            return HotkeyResult(False, "invalid")
        mods, key = parsed
        modmask = _mods_x11(mods)
        if modmask == 0:
            return HotkeyResult(False, "invalid")
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
                if lib.XGrabKey(self._x11_display, keycode, m, self._x11_root,
                                True, GrabModeAsync, GrabModeAsync) != 0:
                    return HotkeyResult(False, "conflict")
                combos.append(m)
            lib.XSync(self._x11_display, False)

            if self._x11_filter is None:
                self._x11_filter = _X11HotkeyEventFilter(self)
                app = QCoreApplication.instance()
                if app is not None:
                    app.installNativeEventFilter(self._x11_filter)
            self._x11_grabs.setdefault(keycode, []).append((modmask, callback, combos))
            return HotkeyResult(True)
        except Exception:
            return HotkeyResult(False, "failed")

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

    # ----- macOS 实现(CGEventTap)-----
    def _register_macos(self, hotkey: str, callback) -> HotkeyResult:
        parsed = _parse_hotkey_macos(hotkey)
        if parsed is None:
            return HotkeyResult(False, "invalid")
        mods, keycode = parsed
        if _mac_cg is None:
            return HotkeyResult(False, "failed")
        try:
            # 全局键盘事件截获需要"辅助功能/输入监听"权限
            if not _mac_is_trusted():
                try:
                    _mac_cg.CGRequestListenEventAccess()  # 触发系统授权对话框
                except Exception:
                    pass
                return HotkeyResult(False, "permission")
            if not _mac_ensure_tap():
                return HotkeyResult(False, "failed")
            global _mac_manager
            _mac_manager = self
            if (keycode, mods) in self._mac_grabs:
                return HotkeyResult(False, "conflict")
            hotkey_id = self._next_id
            self._mac_grabs[(keycode, mods)] = hotkey_id
            self._callbacks[hotkey_id] = callback
            self._next_id += 1
            return HotkeyResult(True)
        except Exception:
            return HotkeyResult(False, "failed")

    def _unregister_all_macos(self):
        global _mac_manager
        try:
            if _mac_manager is self:
                _mac_manager = None
        except Exception:
            pass
        self._mac_grabs.clear()

    # ----- 分发 -----
    def _dispatch(self, hotkey_id: int) -> bool:
        callback = self._callbacks.get(hotkey_id)
        if callback is None:
            return False
        callback()
        return True
