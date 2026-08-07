# -*- coding: utf-8 -*-
"""
跨平台全局快捷键管理器

- Windows: 基于官方 RegisterHotKey API + QAbstractNativeEventFilter 监听 WM_HOTKEY。
           相比第三方 keyboard 库,零额外依赖、更稳定、不易被杀毒软件误报。
           注册失败时可区分"被其他程序占用"(ERROR_HOTKEY_ALREADY_REGISTERED),
           冲突时 RegisterHotKey 不会抢占,不会影响其他应用。
- macOS:   基于 Carbon RegisterEventHotKey API(通过 ctypes 直接调用,零额外依赖)。
           Carbon 框架虽已废弃,但 RegisterEventHotKey 在 macOS 上持续可用,
           是获取系统级全局热键的标准手段。
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
import platform

if platform.system() == "Windows":
    from ctypes import wintypes
else:
    wintypes = None

from PySide6.QtCore import QAbstractNativeEventFilter, QCoreApplication, QObject

# ---------------------------------------------------------------------------
# 通用结果类型
# ---------------------------------------------------------------------------

class HotkeyResult:
    """快捷键注册结果。`ok` 为是否成功,`reason` 为失败原因:
    - 'conflict'    : 已被其他程序占用(热键冲突)
    - 'invalid'     : 快捷键无法解析(缺修饰键 / 键不支持)
    - 'reserved'    : 系统/通用快捷键,为避免影响其他应用而禁止注册
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
# macOS: Carbon RegisterEventHotKey
# ---------------------------------------------------------------------------

# Carbon 事件常量(4 字符码)
K_EVENT_CLASS_HOT_KEY = 0x686B6869       # 'hkhi'
K_EVENT_HOT_KEY_PRESSED = 1
K_EVENT_PARAM_DIRECT_OBJECT = 0x2D2D2D2D  # '----'
TYPE_EVENT_HOT_KEY_ID = 0x686B6964        # 'hkid'
EVENT_HOT_KEY_EXISTS_ERR = -9878          # 热键已存在(冲突)

# Carbon 修饰键(EventModifiers)
MAC_CMD = 0x0100
MAC_SHIFT = 0x0200
MAC_OPTION = 0x0800
MAC_CONTROL = 0x1000

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
    """加载 macOS Carbon framework(仅 Darwin 平台)。"""
    if platform.system() != "Darwin":
        return None
    try:
        return ctypes.cdll.LoadLibrary("/System/Library/Frameworks/Carbon.framework/Carbon")
    except Exception:
        return None


# 模块级持有:避免回调被 GC 回收;mac 只可能有单个 manager 实例
_mac_carbon = _load_carbon()
_mac_manager = None
_mac_handler_ref = None


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
    """Qt 修饰键名 → Carbon 修饰键位。

    Qt 在 macOS 的键盘映射:
        "ctrl" → Command 键(⌘)   -> cmdKey
        "alt"  → Option 键(⌥)    -> optionKey
        "shift"→ Shift 键(⇧)     -> shiftKey
        "meta" → Control 键(⌃)   -> controlKey
    """
    value = 0
    if "ctrl" in mods:
        value |= MAC_CMD
    if "alt" in mods:
        value |= MAC_OPTION
    if "shift" in mods:
        value |= MAC_SHIFT
    if "meta" in mods:
        value |= MAC_CONTROL
    return value


def _parse_hotkey_macos(hotkey: str):
    """macOS 用解析:返回 (carbon_modifiers, keycode);无效返回 None。"""
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
# macOS Carbon 事件处理器(模块级单例)
# ---------------------------------------------------------------------------

if platform.system() == "Darwin":
    class _EventHotKeyID(ctypes.Structure):
        _fields_ = [
            ("signature", ctypes.c_uint32),
            ("id", ctypes.c_uint32),
        ]

    class _EventTypeSpec(ctypes.Structure):
        _fields_ = [
            ("eventClass", ctypes.c_uint32),
            ("eventKind", ctypes.c_uint32),
        ]

    _mac_carbon.GetApplicationEventTarget.restype = ctypes.c_void_p

    _EventHandlerUPP = ctypes.CFUNCTYPE(
        ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p
    )

    @_EventHandlerUPP
    def _mac_event_handler(call_ref, event_ref, user_data):
        """Carbon 热键事件回调:从 EventRef 提取 EventHotKeyID 并分发。"""
        try:
            if _mac_carbon is None or _mac_manager is None:
                return 0
            hkid = _EventHotKeyID()
            size = ctypes.c_size_t()
            err = _mac_carbon.GetEventParameter(
                event_ref, K_EVENT_PARAM_DIRECT_OBJECT, TYPE_EVENT_HOT_KEY_ID,
                None, ctypes.sizeof(hkid), ctypes.byref(size), ctypes.byref(hkid),
            )
            if err != 0:
                return 0
            _mac_manager._dispatch(int(hkid.id))
            return 0
        except Exception:
            return 0
else:
    _EventHandlerUPP = None
    _mac_event_handler = None


# ---------------------------------------------------------------------------
# 全局快捷键管理器
# ---------------------------------------------------------------------------

class GlobalHotkeyManager(QObject):
    """跨平台全局快捷键管理器。

    - Windows: 官方 RegisterHotKey,热键回调直接进入 Qt 事件循环(主线程)。
    - macOS:   Carbon RegisterEventHotKey,回调经 Carbon 事件循环进入主线程。
    - 其他:    register 返回 'unsupported',不影响程序运行。
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._callbacks = {}      # 热键 id -> callback
        self._next_id = 1
        self._filter = None
        # macOS 专用状态
        self._mac_refs = {}       # 热键 id -> EventHotKeyRef
        self._mac_target = None

    # ----- 公共接口 -----
    def register(self, hotkey: str, callback) -> HotkeyResult:
        """注册全局快捷键。返回 HotkeyResult,冲突/无效/保留/不支持时 ok=False。"""
        # 先拦截系统/通用快捷键,避免注册后影响其他应用正常使用
        if is_reserved(hotkey):
            return HotkeyResult(False, "reserved")
        system = platform.system()
        if system == "Windows":
            return self._register_windows(hotkey, callback)
        if system == "Darwin":
            return self._register_macos(hotkey, callback)
        return HotkeyResult(False, "unsupported")

    def unregister_all(self):
        """注销全部已注册的全局快捷键。"""
        system = platform.system()
        if system == "Windows":
            self._unregister_all_windows()
        elif system == "Darwin":
            self._unregister_all_macos()
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

    # ----- macOS 实现 -----
    def _register_macos(self, hotkey: str, callback) -> HotkeyResult:
        parsed = _parse_hotkey_macos(hotkey)
        if parsed is None:
            return HotkeyResult(False, "invalid")
        mods, keycode = parsed
        if _mac_carbon is None:
            return HotkeyResult(False, "failed")
        try:
            global _mac_manager, _mac_handler_ref
            if self._mac_target is None:
                self._mac_target = _mac_carbon.GetApplicationEventTarget()
            # 首次使用时安装一次事件处理器
            if _mac_handler_ref is None:
                spec = _EventTypeSpec(K_EVENT_CLASS_HOT_KEY, K_EVENT_HOT_KEY_PRESSED)
                ref = ctypes.c_void_p()
                err = _mac_carbon.InstallEventHandler(
                    self._mac_target, _mac_event_handler, 1,
                    ctypes.byref(spec), None, ctypes.byref(ref),
                )
                if err != 0:
                    return HotkeyResult(False, "failed")
                _mac_handler_ref = ref
                _mac_manager = self
            hotkey_id = self._next_id
            hkid = _EventHotKeyID(0x53545747, hotkey_id)  # signature 'STWG'
            ref = ctypes.c_void_p()
            err = _mac_carbon.RegisterEventHotKey(
                keycode, mods, ctypes.byref(hkid),
                self._mac_target, 0, ctypes.byref(ref),
            )
            if err != 0:
                if err == EVENT_HOT_KEY_EXISTS_ERR:
                    return HotkeyResult(False, "conflict")
                return HotkeyResult(False, "failed")
            self._callbacks[hotkey_id] = callback
            self._mac_refs[hotkey_id] = ref  # 保持引用,防止被回收
            self._next_id += 1
            return HotkeyResult(True)
        except Exception:
            return HotkeyResult(False, "failed")

    def _unregister_all_macos(self):
        try:
            for ref in self._mac_refs.values():
                if ref and ref.value:
                    _mac_carbon.UnregisterEventHotKey(ref)
            self._mac_refs.clear()
        except Exception:
            pass

    # ----- 分发 -----
    def _dispatch(self, hotkey_id: int) -> bool:
        callback = self._callbacks.get(hotkey_id)
        if callback is None:
            return False
        callback()
        return True
