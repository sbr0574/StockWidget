from functools import partial
import requests
import threading

from PySide6.QtCore import Qt, QEvent, QTimer, Signal
from PySide6.QtGui import QFont, QAction, QColor
from PySide6.QtWidgets import QApplication, QWidget, QMenu, QVBoxLayout, QLabel, QTableView, QHeaderView, QAbstractItemView, QFrame, QStyledItemDelegate

from src.Display import SimpleTableModel, KLineDelegate
from src.hotkeys import GlobalHotkeyManager, HotkeyResult
from services.stock_data import request_quote, strip_market
from src.platform_support import (
    is_wayland,
    hotkeys_supported, click_through_supported,
    opacity_supported, force_top_supported,
    mac_set_window_level, mac_get_window_level, MAC_LEVEL_STATUS,
    apply_click_through, default_font_family, force_top_uses_native_level,
)

def _format_volume(value: int) -> str:
    value = int(value/100)
    if value < 1e4:
        return f"{value}"
    if value < 1e8:
        return f"{value / 1e4:.2f}万"
    return f"{value / 1e8:.2f}亿"

def _format_amount(value: float) -> str:
    if value < 1e8:
        return f"{value / 1e4:.2f}万"
    if value < 1e12:
        return f"{value / 1e8:.2f}亿"
    return f"{value / 1e12:.2f}万亿"


class FloatLabel(QWidget):
    hotkey_triggered = Signal()
    click_through_hotkey_triggered = Signal()
    click_through_changed = Signal(bool)
    data_ready = Signal(object)  # 后台线程请求完成后发回主线程: (ok, ret, data, error)
    ALL_HEADERS = ["名称", "现价", "涨跌", "涨幅", "浮盈", "买一", "卖一", "委比", "成交量", "成交额", "均价", "K线"]
    HEADER_ATTR_MAP = {
        "名称": "name_visible",
        "现价": "price_visible",
        "涨跌": "change_visible",
        "涨幅": "change_pct_visible",
        "浮盈": "profit_visible",
        "买一": "b1s1_visible",
        "卖一": "b1s1_visible",
        "委比": "commi_visible",
        "成交量": "vol_visible",
        "成交额": "amount_visible",
        "均价": "avg_visible",
        "K线": "kline_visible",
    }

    def __init__(self, cfg: dict, codes_list: dict):
        super().__init__()
        self._on_change = (lambda: None)
        self._open_settings_cb = None

        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setFocusPolicy(Qt.StrongFocus)

        self.codes_list: dict = codes_list
        # 加载自选标的配置（代码 -> {checked, cost, name, type}）
        watchlist_cfg           = cfg.get("watchlist", {})
        self.watchlist: dict    = self._normalize_watchlist(watchlist_cfg)
        # 加载面板配置
        self.name_visible       = bool(cfg.get("name_visible", True))
        self.code_visible       = bool(cfg.get("code_visible", False))
        self.type_visible       = bool(cfg.get("type_visible", False))
        self.name_length        = int(cfg.get("name_length", -1))
        self.price_visible      = bool(cfg.get("price_visible", True))
        self.change_visible     = bool(cfg.get("change_visible", False))
        self.change_pct_visible = bool(cfg.get("change_pct_visible", True))
        self.profit_visible     = bool(cfg.get("profit_visible", False))
        self.b1s1_visible       = bool(cfg.get("b1s1_visible", False))
        self.commi_visible      = bool(cfg.get("commi_visible", False))
        self.vol_visible        = bool(cfg.get("vol_visible", False))
        self.amount_visible     = bool(cfg.get("amount_visible", False))
        self.avg_visible        = bool(cfg.get("avg_visible", False))
        self.kline_visible      = bool(cfg.get("kline_visible", False))
        # 加载外观配置
        self.header_visible     = bool(cfg.get("header_visible", False))
        self.grid_visible       = bool(cfg.get("grid_visible", False))
        font_family             = cfg.get("font_family", default_font_family())
        font_size               = int(cfg.get("font_size", 10))
        self.font               = QFont(font_family, max(5, min(15, font_size)))
        self.line_extra_px      = int(cfg.get("line_extra_px", 1))
        self.fg                 = QColor(cfg.get("fg", "#FFFFFF"))
        bg                      = cfg.get("bg", {"r":0,"g":0,"b":0,"a":191})
        self.bg                 = QColor(bg["r"],bg["g"],bg["b"],bg["a"])
        self.opacity_pct        = int(cfg.get("opacity_pct", 90))
        self.default_color      = bool(cfg.get("default_color", False))
        # 加载其他配置
        self.refresh_seconds    = int(cfg.get("refresh_seconds", 2))
        self.force_top          = bool(cfg.get("force_top", False))
        self.click_through      = bool(cfg.get("click_through", False))
        self.hotkey_enabled     = bool(cfg.get("hotkey_enabled", False))
        self.hotkey             = cfg.get("hotkey", "Ctrl+Alt+F")
        self.hotkey_click_through_enabled = bool(cfg.get("hotkey_click_through_enabled", False))
        self.hotkey_click_through = cfg.get("hotkey_click_through", "Ctrl+Alt+C")
        self.start_on_boot      = bool(cfg.get("start_on_boot", False))

        # 平台能力限制:当前平台不支持时强制关闭对应功能
        # (如 Wayland 下无法实现全局快捷键/鼠标穿透,Linux 下强制置顶不可靠),
        # 并交由设置面板/托盘菜单将相关控件置为不可点按。
        if not hotkeys_supported():
            self.hotkey_enabled = False
            self.hotkey_click_through_enabled = False
        if not click_through_supported():
            self.click_through = False
        if not force_top_supported():
            self.force_top = False

        # Wayland 会话下窗口位置由合成器接管,须用系统级拖动(startSystemMove)
        self._wayland_drag = is_wayland()

        self.hotkey_triggered.connect(self.toggle_win)
        self.click_through_hotkey_triggered.connect(self.toggle_click_through)
        # 全局快捷键管理器:Windows 用官方 RegisterHotKey,Linux/X11 用 XGrabKey,Wayland 不支持
        self._hotkeys = GlobalHotkeyManager(self)
        self._register_hotkey()

        # UI
        self.panel = QWidget(self)
        self.panel.setObjectName("panel")
        self.vbox = QVBoxLayout(self.panel)
        self.vbox.setContentsMargins(10,6,10,6)
        self.vbox.setSpacing(0)

        self.table = QTableView(self.panel)
        self.table.setFrameShape(QFrame.NoFrame)
        self.table.setShowGrid(False)
        self.table.setSelectionMode(QAbstractItemView.NoSelection)
        self.table.setFocusPolicy(Qt.NoFocus)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setMinimumSectionSize(1)
        self.table.verticalHeader().setDefaultSectionSize(1)
        self.table.horizontalHeader().setVisible(self.header_visible)
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.setFont(self.font)
        self.table.horizontalHeader().setFont(self.font)
        self.table.horizontalHeader().setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.table.setTextElideMode(Qt.ElideNone)
        self.message_label = QLabel("", self.panel)
        self.message_label.setStyleSheet("padding: 2px 4px;")
        self.message_label.setVisible(False)
        self.vbox.addWidget(self.message_label)
        self._index_updating = False # 市场代码列表后台更新标志
        self._refresh_thread = None  # 后台刷新线程（避免网络请求阻塞 UI）

        self.model = SimpleTableModel(headers=self.ALL_HEADERS, align_right_cols=[1,2,3,4,5])
        self.model.set_color_scheme(self.default_color, self.fg)
        self.table.setModel(self.model)

        self.k_delegate = KLineDelegate(self.table, base_pt=12)
        self.k_delegate.update_scheme(self.default_color, self.fg)
        self.k_delegate.set_point_size(self.font.pointSize())
        self.k_column_visible_index = None

        self.vbox.addWidget(self.table)

        for w in (self.panel, self.table, self.table.viewport(), self.table.horizontalHeader()):
            w.installEventFilter(self)

        self.apply_style()
        self.set_window_opacity_percent(self.opacity_pct)
        self._fit_to_contents()

        scr = QApplication.primaryScreen().availableGeometry()
        pos = cfg.get("pos")
        if isinstance(pos, dict) and "x" in pos and "y" in pos:
            x, y = int(pos["x"]), int(pos["y"])
            x = max(scr.left(), min(x, scr.right()-self.width()))
            y = max(scr.top(),  min(y, scr.bottom()-self.height()))
            self.move(x, y)
        else:
            self.move(scr.right()-self.width()-40, scr.bottom()-self.height()-80)

        self._drag_pos = None
        self._system_moving = False
        self._mac_orig_level = None  # macOS 强制置顶前的原始窗口层级

        # 定时刷新数据
        self.data_ready.connect(self._process_data)
        self.timer = QTimer(self)
        self.timer.setInterval(max(1, self.refresh_seconds)*1000)
        self.timer.timeout.connect(self._refresh_from_function)
        self.timer.start()
        self._refresh_from_function()
        self._defer_fit()

        # 强制置顶定时
        self._keep_top_timer = QTimer(self)
        self._keep_top_timer.setInterval(1000)  # 每 1000ms 检查一次
        self._keep_top_timer.timeout.connect(self._ensure_on_top)
        if self.force_top:
            self._keep_top_timer.start()

        self.set_click_through(self.click_through)

    # ----- 自选标的派生属性（由 watchlist 生成） -----
    @staticmethod
    def _normalize_watchlist(watchlist: dict) -> dict:
        """规范化自选列表：代码小写，cost 转数值（整数值保持 int），name/type 转字符串"""
        result = {}
        for key, info in (watchlist or {}).items():
            key = str(key).strip().lower()
            if not key:
                continue
            entry = dict(info or {})
            entry["checked"] = bool(entry.get("checked", True))
            try:
                val = float(entry["cost"]) if entry.get("cost") not in (None, "") else None
            except (TypeError, ValueError):
                val = None
            if val is not None and val.is_integer():
                val = int(val)
            entry["cost"] = val
            entry["name"] = str(entry.get("name", "") or "").strip()
            entry["type"] = str(entry.get("type", "") or "").strip()
            result[key] = entry
        return result

    @property
    def codes(self) -> list:
        return list(self.watchlist.keys())

    @property
    def checked_codes(self) -> list:
        return [c for c, e in self.watchlist.items() if e.get("checked")]

    @property
    def costs(self) -> dict:
        return {c: e["cost"] for c, e in self.watchlist.items() if e.get("cost")}

    # 与 App 连接
    def set_open_settings_callback(self, fn): 
        self._open_settings_cb = fn

    def set_on_change(self, fn): 
        self._on_change = fn or (lambda: None)

    def _notify_change(self):
        cb = getattr(self, "_on_change", None)
        if callable(cb): cb()

    def current_config(self):
        return {
            "watchlist":            {c: dict(e) for c, e in self.watchlist.items()},

            "name_visible":         self.name_visible,
            "code_visible":         self.code_visible,
            "type_visible":         self.type_visible,
            "name_length":          self.name_length,
            "price_visible":        self.price_visible,
            "change_visible":       self.change_visible,
            "change_pct_visible":   self.change_pct_visible,
            "profit_visible":       self.profit_visible,
            "b1s1_visible":         self.b1s1_visible,
            "commi_visible":        self.commi_visible,
            "vol_visible":          self.vol_visible,
            "amount_visible":       self.amount_visible,
            "avg_visible":          self.avg_visible,
            "kline_visible":        self.kline_visible,
            
            "header_visible":   self.header_visible,
            "grid_visible":     self.grid_visible,
            "font_family":      self.font.family(),
            "font_size":        self.font.pointSize(),
            "line_extra_px":    self.line_extra_px,
            "fg":               self.fg.name(QColor.HexRgb),
            "bg":               {"r": self.bg.red(), "g": self.bg.green(), "b": self.bg.blue(), "a": self.bg.alpha()},
            "opacity_pct":      int(round(getattr(self, "opacity_pct", 90))),
            "default_color":    self.default_color,

            "refresh_seconds":  self.refresh_seconds,
            "force_top":        self.force_top,
            "click_through":    self.click_through,
            "hotkey_enabled":   self.hotkey_enabled,
            "hotkey":           self.hotkey,
            "hotkey_click_through_enabled": self.hotkey_click_through_enabled,
            "hotkey_click_through": self.hotkey_click_through,
            "start_on_boot":    self.start_on_boot,
            "pos":              {"x": self.x(), "y": self.y()},
        }

    def header_is_visible(self, header: str) -> bool:
        attr = self.HEADER_ATTR_MAP.get(header)
        return bool(getattr(self, attr, False)) if attr else False

    # ----- 外观/尺寸 -----
    def apply_style(self):
        r,g,b,a = self.bg.red(), self.bg.green(), self.bg.blue(), self.bg.alpha()
        fg_r, fg_g, fg_b = self.fg.red(), self.fg.green(), self.fg.blue()
        line_col = f"rgba({fg_r},{fg_g},{fg_b},80)"
        self.panel.setStyleSheet(f"""
            QWidget#panel {{
                background: rgba({r},{g},{b},{a});
                border-radius: 5px;
            }}
            QTableView {{
                background: transparent;
                border: {f"1px solid {line_col}" if self.grid_visible else "none"};
                border-radius: 3px;
                {"" if self.default_color else f"color: {self.fg.name()};"}
                outline: none;
            }}
            QTableView::item {{
                border-right: {f"1px solid {line_col}" if self.grid_visible else "none"};
                border-bottom: {f"1px solid {line_col}" if self.grid_visible else "none"};
            }}
            QHeaderView {{
                background-color: transparent;
            }}
            QHeaderView::section {{
                background: transparent;
                border: none;
                border-bottom: 1px solid {line_col};
                font-weight: 600;
                {"" if self.default_color else f"color: {self.fg.name()};"}
                padding: 2px 4px;
            }}
        """)
        self.table.setFont(self.font)
        self.table.horizontalHeader().setFont(self.font)
        self._defer_fit()

    def _apply_row_heights(self):
        fm = self.table.fontMetrics()
        h = fm.height() + max(0, self.line_extra_px)
        self.table.verticalHeader().setDefaultSectionSize(h)
        for r in range(self.model.rowCount()):
            self.table.setRowHeight(r, h)

    def _fit_to_contents(self):
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.resizeColumnsToContents()
        self._apply_row_heights()

        cols = self.model.columnCount()
        rows = self.model.rowCount()
        self.table.verticalHeader().setFixedWidth(0)
        total_w = 2*self.table.frameWidth()
        for c in range(cols): 
            total_w += self.table.columnWidth(c)
        hh = self.table.horizontalHeader().height() if self.table.horizontalHeader().isVisible() else 0
        total_h = hh + 2*self.table.frameWidth()
        for r in range(rows): 
            total_h += self.table.rowHeight(r)
        self.table.setFixedSize(max(1,total_w), max(1,total_h))
        self.panel.adjustSize()
        self.resize(self.panel.size())

    def _defer_fit(self):
        QTimer.singleShot(0, self._fit_to_contents)

    # ----- 数据 & 投影 -----
    def _show_message(self, msg: str, is_error: bool = False):
        """显示顶部提示；is_error=True 时用红色字体，否则用前景色"""
        text = str(msg) if msg is not None else ""
        color = "#ff6666" if is_error else self.fg.name(QColor.HexRgb)
        self.message_label.setStyleSheet(f"color: {color}; padding: 2px 4px;")
        self.message_label.setText(text)
        self.message_label.setVisible(True)
        self._defer_fit()

    def _clear_message(self):
        """清除顶部提示"""
        self.message_label.setVisible(False)
        self.message_label.setText("")

    def set_index_updating(self, updating: bool):
        """标记市场代码列表是否正在后台更新（期间保持进度提示不被清除）"""
        self._index_updating = bool(updating)

    def _project_columns(self, full_rows: list[dict], sign_data: list[dict]):
        # 名称作为数据列显示；其余按显示顺序筛选已启用的列
        headers = [h for h in self.ALL_HEADERS if self.header_is_visible(h)]

        proj_rows, proj_meta = [], []
        for r, row in enumerate(full_rows):
            proj_rows.append([row[h] for h in headers])
            proj_meta.append([sign_data[r][h] for h in headers])

        # 右对齐：名称、K线、卖一除外
        right_cols = [i for i, h in enumerate(headers) if h not in ("名称", "K线", "卖一")]
        self.model.set_align_right_cols(right_cols)
        self.model.set_rows_headers(proj_rows, headers, proj_meta)
        self.model.set_color_scheme(self.default_color, self.fg)

        if "K线" in headers:
            col = headers.index("K线")
            self.k_column_visible_index = col
            self.k_delegate.update_scheme(self.default_color, self.fg)
            self.k_delegate.set_point_size(self.font.pointSize())
            self.table.setItemDelegateForColumn(col, self.k_delegate)
        else:
            if self.k_column_visible_index is not None:
                self.table.setItemDelegateForColumn(self.k_column_visible_index, QStyledItemDelegate(self.table))
                self.k_column_visible_index = None

        self._fit_to_contents()

    def _format_data(self, code: str, data: dict, type: str):
        # 名称显示
        name = f"({type})" if type is not None and self.type_visible else ""
        name += f"{strip_market(code)} " if self.code_visible else ""
        if self.name_length == -1:
            name += data["name"]
        else:
            name += data["name"][:self.name_length]

        # 一档盘口数据
        b1_label = ""
        s1_label = ""
        b1_color_sign = 0
        s1_color_sign = 0
        pur_1 = data["purchaser_price"][0]
        sell_1 = data["seller_price"][0]
        if pur_1 == sell_1 > 0:
            # 集合竞价阶段
            data["current_price"] = sell_1
            paired = int(data["seller_vol"][0] / 100)
            unpaired = int((data["purchaser_vol"][1] or (-data["seller_vol"][1])) / 100)
            b1_label = f"{paired:d}"
            s1_label = f"{unpaired:+d}"
            b1_color_sign = (unpaired > 0) - (unpaired < 0)
            s1_color_sign = b1_color_sign
        else:
            # 连续交易阶段（有买/卖盘口量时才显示，否则"-"）
            pur_v1 = data["purchaser_vol"][0]
            sell_v1 = data["seller_vol"][0]
            buy_marker = "<" if pur_1 and pur_v1 and data["current_price"] == pur_1 else " "
            sell_marker = ">" if sell_1 and sell_v1 and data["current_price"] == sell_1 else " "
            b1_label = f"{int(pur_v1 / 100)}{buy_marker}" if (pur_1 and pur_v1) else "-"
            s1_label = f"{sell_marker}{int(sell_v1 / 100)}" if (sell_1 and sell_v1) else "-"
            b1_color_sign = 1 if (pur_1 and pur_v1) else 0
            s1_color_sign = -1 if (sell_1 and sell_v1) else 0

        # 盘前数据填充
        if data["current_price"] == 0:
            data["current_price"] = data["prev_close"]
        if data["opening_price"] == 0:
            data["opening_price"] = data["current_price"]
            data["high_price"] = data["current_price"]
            data["low_price"] = data["current_price"]

        # 指标计算
        change = data["current_price"] - data["prev_close"] if data["prev_close"] else 0.0
        change_pct = (data["current_price"] / data["prev_close"] - 1) * 100 if data["prev_close"] else 0.0
        avg = (data["deals_amt"] / data["deals_vol"]) if data["deals_vol"] > 0 else data["prev_close"]
        p_sum, s_sum = sum(data["purchaser_vol"]), sum(data["seller_vol"])
        committee = (100 * (p_sum - s_sum) / (p_sum + s_sum)) if (p_sum + s_sum) > 0 else 0.0
        arrow = " "
        if data["high_price"] > data["low_price"]:
            if data["current_price"] == data["high_price"]: arrow = "↑"
            elif data["current_price"] == data["low_price"]: arrow = "↓"
        k_payload = {"k": (data["opening_price"], data["current_price"], data["high_price"], data["low_price"], data["prev_close"])}

        precision = 3 if type == "基" else 2

        # 浮盈计算（与成本价比较），仅显示百分比
        cost = self.costs.get(code)
        if cost is not None and cost > 0:
            profit_pct = (data["current_price"] / cost - 1) * 100
            profit_label = f"{profit_pct:+.2f}%"
            profit_sign = (profit_pct > 0) - (profit_pct < 0)
        else:
            profit_label = "-"
            profit_sign = 0

        # 数据返回
        is_index = type == "指"
        format_data = {
            "名称": name,
            "现价": f"{data["current_price"]:.{precision}f}{arrow}",
            "涨跌": f"{change:+.{precision}f}",
            "涨幅": f"{change_pct:+.2f}%",
            "浮盈": profit_label,
            "买一": b1_label,
            "卖一": s1_label,
            "委比": f"{committee:+.2f}%" if (p_sum + s_sum) > 0 else "-",
            "成交量": ("-" if is_index and not data["deals_vol"] else _format_volume(data["deals_vol"])),
            "成交额": ("-" if is_index and not data["deals_amt"] else _format_amount(data["deals_amt"])),
            "均价": f"{avg:.{precision}f}",
            "K线": k_payload}
        sign = {
            "名称": 0,
            "现价": (change > 0) - (change < 0),
            "涨跌": (change > 0) - (change < 0),
            "涨幅": (change > 0) - (change < 0),
            "浮盈": profit_sign,
            "买一": b1_color_sign,
            "卖一": s1_color_sign,
            "委比": (committee > 0) - (committee < 0),
            "成交量": 0,
            "成交额": 0,
            "均价": (avg > data["prev_close"]) - (avg < data["prev_close"]),
            "K线": 0}
        # 指数不显示浮盈/买一卖一/委比/均价（均置为"-"）
        if type == "指":
            for key in ("浮盈", "买一", "卖一", "委比", "均价"):
                format_data[key] = "-"
                sign[key] = 0
        return format_data, sign

    def _get_code_info(self, c: str) -> dict:
        return self.codes_list.get(c, {})

    def _refresh_from_function(self):
        """定时入口：将网络请求丢到后台线程执行，避免阻塞 UI。
        若上一轮请求尚未完成则跳过本次刷新，防止请求重叠。"""
        if self._refresh_thread is not None and self._refresh_thread.is_alive():
            return
        self._refresh_thread = threading.Thread(
            target=self._fetch_data_worker,
            args=(self.checked_codes,),
            daemon=True,
        )
        self._refresh_thread.start()

    def _fetch_data_worker(self, codes: list):
        """后台线程：执行网络请求，结果经 data_ready 信号回到主线程。"""
        try:
            data = request_quote(codes)
            payload = (True, data, None)
        except requests.exceptions.RequestException:
            payload = (False, None, "网络请求失败")
        except Exception as e:
            payload = (False, None, str(e))
        self.data_ready.emit(payload)

    def _process_data(self, payload):
        """主线程：处理请求结果并更新表格。payload = (ok, data, error)"""
        ok, data, error = payload
        if not ok:
            self._show_message(error or "请求失败", is_error=True)
            return

        full_rows = []
        full_sign = []
        for c, d in data.items():
            entry = self.watchlist.get(c) or {}
            type_ = entry.get("type") or self._get_code_info(c).get("type")
            row, sign = self._format_data(c, d, type_)
            full_rows.append(row)
            full_sign.append(sign)

        if not self._index_updating:
            if len(data) > 0:
                self._clear_message()
            else:
                self._show_message("请在设置面板中添加自选股", is_error=True)
        self._project_columns(full_rows, full_sign)

    # ----- 应用设置 -----
    def set_watchlist(self, watchlist: dict):
        """整体替换自选列表（代码 -> {checked, cost, name, type}）"""
        self.watchlist = self._normalize_watchlist(watchlist)
        self._notify_change()
        self._refresh_from_function()

    def set_type_visible(self, visible: bool):
        self.type_visible = bool(visible)
        self._notify_change()
        self._refresh_from_function()

    def set_code_visible(self, visible: bool):
        self.code_visible = bool(visible)
        self._notify_change()
        self._refresh_from_function()

    def set_flag(self, header, checked: bool):
        if isinstance(header, int):
            if 0 <= header < len(self.ALL_HEADERS):
                header = self.ALL_HEADERS[header]
            else:
                return
        header = str(header)
        attr = self.HEADER_ATTR_MAP.get(header)
        if not attr:
            return

        checked = bool(checked)
        if bool(getattr(self, attr, False)) == checked:
            return
        setattr(self, attr, checked)
        self._notify_change()
        self._refresh_from_function()

    def set_code_type(self, pure_num: bool):
        self.short_code = bool(pure_num)
        self._notify_change()
        self._refresh_from_function()

    def set_name_length(self, name_len: int):
        # -1 全部显示, 0 不显示, >0 显示前 N 个字
        if name_len == -1 or name_len >= 0:
            self.name_length = name_len
            self._notify_change()
            self._refresh_from_function()

    def set_b1s1_display(self, mode: str):
        """mode: 'qty' | 'price' | 'both'"""
        if mode not in ("qty", "price", "both"):
            return
        self.b1s1_display = mode
        self._notify_change()
        self._refresh_from_function()

    def set_header_visible(self, vis: bool):
        self.header_visible = bool(vis)
        self.table.horizontalHeader().setVisible(self.header_visible)
        self._notify_change()
        self._defer_fit()

    def set_grid_visible(self, vis: bool):
        self.grid_visible = bool(vis)
        self.apply_style()
        self._notify_change()

    def set_refresh_interval(self, seconds: int):
        try:
            seconds = max(1, int(seconds))
        except Exception:
            return
        self.refresh_seconds = seconds
        if self.timer is not None:
            self.timer.setInterval(seconds * 1000)
        self._notify_change()

    def set_fg_color(self, c: QColor):
        if isinstance(c, QColor) and c.isValid():
            self.fg = QColor(c)
            self.apply_style()
            self._notify_change()

    def set_bg_rgb_keep_alpha(self, c: QColor):
        if isinstance(c, QColor) and c.isValid():
            c2 = QColor(c)
            c2.setAlpha(self.bg.alpha())
            self.bg = c2
            self.apply_style()
            self._notify_change()

    def set_bg_alpha_percent(self, percent_0_100: int):
        p = max(0, min(100, int(percent_0_100)))
        self.bg.setAlpha(int(round(p*2.55)))
        self.apply_style()
        self._notify_change()

    def set_window_opacity_percent(self, percent_20_100: int):
        p = max(20, min(100, int(percent_20_100)))
        self.opacity_pct = p
        if opacity_supported():
            # Wayland 平台插件不支持设置窗口透明度,跳过以避免终端告警
            self.setWindowOpacity(p / 100.0)
        self._defer_fit()
        self._notify_change()

    def set_font_size(self, pt: int):
        pt = max(5, min(15, int(pt)))
        self.font.setPointSize(pt)
        self.k_delegate.set_point_size(pt)
        self.apply_style()
        self._notify_change()
        self.table.viewport().update()
        self._defer_fit()

    def set_font_family(self, family: str):
        if family and family != self.font.family():
            self.font.setFamily(family)
            self.apply_style()
            self._notify_change()

    def set_line_extra(self, px: int):
        self.line_extra_px = max(0, int(px))
        self.apply_style()
        self._defer_fit()
        self._notify_change()

    def set_default_color(self, enabled: bool):
        self.default_color = bool(enabled)
        self.model.set_color_scheme(self.default_color, self.fg)
        self.k_delegate.update_scheme(self.default_color, self.fg)
        self.apply_style()
        self._notify_change()
        self._defer_fit()

    # ----- 鼠标穿透 / 强制置顶 / 快捷键开关 -----
    def set_click_through(self, enable: bool):
        enable = bool(enable)
        if self.click_through == enable:
            return
        self.click_through = enable
        apply_click_through(self, self.click_through)
        self.click_through_changed.emit(self.click_through)
        self._notify_change()

    def toggle_click_through(self):
        self.set_click_through(not self.click_through)

    def set_force_top(self, enabled: bool):
        enabled = bool(enabled)
        if not force_top_supported():
            # 仅 Windows / macOS 支持强制置顶,其余平台忽略该开关
            enabled = False
        if self.force_top == enabled:
            return
        self.force_top = enabled
        if self.force_top:
            self._apply_mac_force_top(True)
            # macOS 由原生窗口层级保证置顶,无需轮询 raise_(轮询会抢焦点)
            if (not force_top_uses_native_level() and self.isVisible()
                    and self._keep_top_timer and not self._keep_top_timer.isActive()):
                self._keep_top_timer.start()
            self._ensure_on_top()
        else:
            self._apply_mac_force_top(False)
            if self._keep_top_timer and self._keep_top_timer.isActive():
                self._keep_top_timer.stop()
        self._notify_change()

    def _apply_mac_force_top(self, enabled: bool):
        """macOS:通过原生 NSWindow 层级实现置顶。
        - 启用:抬升到 kCGStatusWindowLevel(25),高于所有普通/浮动窗口,
          并记住开启前的原始层级。
        - 关闭:还原到开启前的原始层级(即 Qt 默认的置顶层级)。
        非 macOS 平台直接忽略。
        """
        if not force_top_uses_native_level():
            return
        try:
            hwnd = int(self.winId())
            if not hwnd:
                return
            if enabled:
                if self._mac_orig_level is None:
                    self._mac_orig_level = mac_get_window_level(hwnd)
                mac_set_window_level(hwnd, MAC_LEVEL_STATUS)
            else:
                if self._mac_orig_level is not None:
                    mac_set_window_level(hwnd, self._mac_orig_level)
                    self._mac_orig_level = None
        except Exception:
            pass

    def set_hotkey_enabled(self, enabled: bool) -> HotkeyResult:
        """启用/停用"显示/隐藏"快捷键。启用失败(如冲突)时回滚状态并返回失败原因。"""
        enabled = bool(enabled)
        if self.hotkey_enabled == enabled:
            return HotkeyResult(True)
        old = self.hotkey_enabled
        self.hotkey_enabled = enabled
        result = self._register_current()
        if not result:
            self.hotkey_enabled = old
            self._register_current()
            return result
        self._notify_change()
        return result

    def set_click_through_hotkey_enabled(self, enabled: bool) -> HotkeyResult:
        """启用/停用"鼠标穿透"快捷键。启用失败(如冲突)时回滚状态并返回失败原因。"""
        enabled = bool(enabled)
        if self.hotkey_click_through_enabled == enabled:
            return HotkeyResult(True)
        old = self.hotkey_click_through_enabled
        self.hotkey_click_through_enabled = enabled
        result = self._register_current()
        if not result:
            self.hotkey_click_through_enabled = old
            self._register_current()
            return result
        self._notify_change()
        return result

    def update_click_through_hotkey(self, new_hotkey: str) -> HotkeyResult:
        """更新"鼠标穿透"快捷键。冲突/无效时不生效并回滚,返回失败原因。"""
        new_hotkey = new_hotkey.strip()
        if new_hotkey == self.hotkey_click_through:
            return HotkeyResult(True)
        if not self.hotkey_click_through_enabled:
            # 未启用时直接保存,启用时再校验
            self.hotkey_click_through = new_hotkey
            self._notify_change()
            return HotkeyResult(True)
        old = self.hotkey_click_through
        self.hotkey_click_through = new_hotkey
        result = self._register_current()
        if not result:
            self.hotkey_click_through = old
            self._register_current()
            return result
        self._notify_change()
        return result

    # ----- 交互 -----
    def contextMenuEvent(self, event):
        menu = QMenu(self)
        sub_cols = QMenu("显示指标", menu)
        for name in self.ALL_HEADERS:
            if name == "卖一":
                continue
            if name == "买一":
                act = QAction("买一/卖一", sub_cols, checkable=True)
                act.setChecked(self.header_is_visible("买一"))
                act.toggled.connect(partial(self.set_flag, "买一"))
                sub_cols.addAction(act)
                continue
            act = QAction(name, sub_cols, checkable=True)
            act.setChecked(self.header_is_visible(name))
            act.toggled.connect(partial(self.set_flag, name))
            sub_cols.addAction(act)
        menu.addMenu(sub_cols)

        act_header = QAction("显示表头", menu, checkable=True)
        act_header.setChecked(self.header_visible)
        act_header.toggled.connect(self.set_header_visible)
        menu.addAction(act_header)

        act_grid = QAction("显示网格",menu, checkable=True)
        act_grid.setChecked(self.grid_visible)
        act_grid.toggled.connect(self.set_grid_visible)
        menu.addAction(act_grid)

        act_color = QAction("默认颜色", menu, checkable=True)
        act_color.setChecked(self.default_color)
        act_color.toggled.connect(self.set_default_color)
        menu.addAction(act_color)

        menu.addSeparator()
        act_open_settings = QAction("设置…", menu)
        if callable(self._open_settings_cb):
            act_open_settings.triggered.connect(self._open_settings_cb)
        else:
            act_open_settings.setEnabled(False)
        menu.addAction(act_open_settings)

        menu.addSeparator()
        menu.addAction(QAction("隐藏浮窗", menu, triggered=self.hide))
        menu.exec(event.globalPos())

    # ----- 窗口拖动 -----
    def _drag_press(self, e):
        """按下左键:记录拖动起点。
        Wayland 下先记全局坐标,待移动超过阈值后再交给合成器(startSystemMove),
        这样普通单击/双击(隐藏浮窗)不受影响。
        """
        if self._wayland_drag:
            self._drag_pos = e.globalPosition().toPoint()
            self._system_moving = False
        else:
            self._drag_pos = e.globalPosition().toPoint() - self.frameGeometry().topLeft()
        self.setFocus(Qt.MouseFocusReason)

    def _drag_move(self, e):
        """按住左键移动:Wayland 用系统级拖动,其余平台(X11/Windows)手动 move。"""
        if getattr(self, "_drag_pos", None) is None or not (e.buttons() & Qt.LeftButton):
            return
        if self._wayland_drag:
            if self._system_moving:
                return
            # 移动超过阈值才触发系统级拖动,避免把单击误判为拖动
            pos = e.globalPosition().toPoint()
            if (pos - self._drag_pos).manhattanLength() <= 4:
                return
            self._system_moving = True
            win = self.windowHandle()
            if win is not None and hasattr(win, "startSystemMove"):
                win.startSystemMove()
            return
        self.move(e.globalPosition().toPoint() - self._drag_pos)
        self._ensure_on_top()

    def _drag_release(self):
        self._drag_pos = None
        self._system_moving = False
        self._ensure_on_top()
        self._notify_change()

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_press(e)

    def mouseMoveEvent(self, e):
        self._drag_move(e)

    def mouseReleaseEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_release()

    def mouseDoubleClickEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_pos = None
            self._system_moving = False
            self.hide()

    def eventFilter(self, obj, ev):
        if ev.type() == QEvent.MouseButtonDblClick and hasattr(ev, "button") and ev.button() == Qt.LeftButton:
            self._drag_pos = None
            self._system_moving = False
            self.hide()
            return True
        if ev.type() == QEvent.MouseButtonPress and hasattr(ev, "button") and ev.button() == Qt.LeftButton:
            self._drag_press(ev)
            return True
        if ev.type() == QEvent.MouseMove and hasattr(ev, "buttons") and (ev.buttons() & Qt.LeftButton) and getattr(self, "_drag_pos", None):
            self._drag_move(ev)
            return True
        if ev.type() == QEvent.MouseButtonRelease and hasattr(ev, "button") and ev.button() == Qt.LeftButton:
            self._drag_release()
            return True
        return QWidget.eventFilter(self, obj, ev)

    def closeEvent(self, event): 
        event.ignore()
        self.hide()

    def showEvent(self, event):
        super().showEvent(event)
        # macOS:每次显示都重新断言原生窗口层级(置顶/强制置顶),
        # 防止被全屏应用或系统(如 hide/show 或睡眠唤醒)重置后掉出置顶。
        if force_top_uses_native_level():
            self._apply_mac_force_top(self.force_top)
        if self.timer and not self.timer.isActive(): 
            self.timer.start()
        # macOS 无需轮询 raise_(原生层级已保证置顶,轮询会抢焦点)
        if (not force_top_uses_native_level() and self.force_top
                and self._keep_top_timer and not self._keep_top_timer.isActive()):
            self._keep_top_timer.start()
        apply_click_through(self, self.click_through)
        self._defer_fit()

    def hideEvent(self, event):
        super().hideEvent(event)
        if self.timer and self.timer.isActive(): 
            self.timer.stop()
        if self._keep_top_timer and self._keep_top_timer.isActive():
            self._keep_top_timer.stop()

    def _ensure_on_top(self):
        if not self.force_top or not self.isVisible():
            return
        if self.click_through:
            return
        if force_top_uses_native_level():
            # macOS:原生 NSWindow 层级(25)已保证置顶,无需轮询 raise_;
            # 周期性 raise_ 会把本程序不断激活,抢夺其他正在使用程序的焦点。
            return
        try:
            aw = QApplication.activeWindow()
            popup = QApplication.activePopupWidget()
            if aw is not None and aw is not self and not self.isAncestorOf(aw):
                return
            if popup is not None and popup is not self and not self.isAncestorOf(popup):
                return
        except Exception:
            pass
        self.raise_()

    def _register_current(self) -> HotkeyResult:
        """按当前 self.* 状态全量注册全局快捷键,返回第一个失败结果。"""
        self._hotkeys.unregister_all()
        if self.hotkey_enabled:
            result = self._hotkeys.register(self.hotkey, lambda: self.hotkey_triggered.emit())
            if not result:
                return result
        if self.hotkey_click_through_enabled:
            result = self._hotkeys.register(
                self.hotkey_click_through,
                lambda: self.click_through_hotkey_triggered.emit(),
            )
            if not result:
                return result
        return HotkeyResult(True)

    def _register_hotkey(self):
        """按当前状态注册全局快捷键(启动/切换时调用)。失败静默,不影响运行。"""
        self._register_current()

    def update_hotkey(self, new_hotkey: str) -> HotkeyResult:
        """更新"显示/隐藏"快捷键。冲突/无效时不生效并回滚,返回失败原因。"""
        new_hotkey = new_hotkey.strip()
        if new_hotkey == self.hotkey:
            return HotkeyResult(True)
        if not self.hotkey_enabled:
            # 未启用时直接保存,启用时再校验
            self.hotkey = new_hotkey
            self._notify_change()
            return HotkeyResult(True)
        old = self.hotkey
        self.hotkey = new_hotkey
        result = self._register_current()
        if not result:
            self.hotkey = old
            self._register_current()
            return result
        self._notify_change()
        return result

    def toggle_win(self):
        if self.isVisible():
            self.hide()
        else:
            self.show()
