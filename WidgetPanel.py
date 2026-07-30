import importlib
from functools import partial
import platform

try:
    if platform.system() == "Windows":
        keyboard = importlib.import_module("keyboard")
    elif platform.system() == "Darwin":
        keyboard = importlib.import_module("keyboardMac")
    else:
        keyboard = None
except Exception:
    keyboard = None

from PySide6.QtCore import Qt, QEvent, QTimer, Signal
from PySide6.QtGui import QFont, QAction, QColor
from PySide6.QtWidgets import QApplication, QWidget, QMenu, QVBoxLayout, QLabel, QTableView, QHeaderView, QAbstractItemView, QFrame, QStyledItemDelegate

from Display import SimpleTableModel, KLineDelegate
from stock_data import request_sina

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
    ALL_HEADERS = ["名称", "现价", "涨跌", "涨幅", "买一", "卖一", "委比", "成交量", "成交额", "均价", "K线"]
    HEADER_ATTR_MAP = {
        "名称": "name_visible",
        "现价": "price_visible",
        "涨跌": "change_visible",
        "涨幅": "change_pct_visible",
        "买一": "b1s1_visible",
        "卖一": "b1s1_visible",
        "委比": "commi_visible",
        "成交量": "vol_visible",
        "成交额": "amount_visible",
        "均价": "avg_visible",
        "K线": "kline_visible",
    }

    def __init__(self, cfg: dict, code_list: dict):
        super().__init__()
        self._on_change = (lambda: None)
        self._open_settings_cb = None

        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setFocusPolicy(Qt.StrongFocus)

        self.code_list_last_update = code_list["last_update"]
        self.code_list: dict = code_list["codes"]
        # 加载自选标的配置
        self.codes              = [str(c).strip() for c in cfg.get("codes",["sh000001"]) if str(c).strip()]
        self.checked_codes      = [str(c).strip() for c in cfg.get("checked_codes", self.codes) if (str(c).strip() and str(c).strip() in self.codes)]
        self.info               = [self.code_list.get(c[2:] if c[:2] in ('sh','sz','bj') else c, {}) for c in self.codes]
        # 加载面板配置
        self.name_visible       = bool(cfg.get("name_visible", False))
        self.market_visible     = bool(cfg.get("market_visible", False))
        self.code_visible       = bool(cfg.get("code_visible", False))
        self.name_length        = int(cfg.get("name_length",0))
        self.price_visible      = bool(cfg.get("price_visible", False))
        self.change_visible     = bool(cfg.get("change_visible", False))
        self.change_pct_visible = bool(cfg.get("change_pct_visible", False))
        self.b1s1_visible       = bool(cfg.get("b1s1_visible", False))
        self.commi_visible      = bool(cfg.get("commi_visible", False))
        self.vol_visible        = bool(cfg.get("vol_visible", False))
        self.amount_visible     = bool(cfg.get("amount_visible", False))
        self.avg_visible        = bool(cfg.get("avg_visible", False))
        self.kline_visible      = bool(cfg.get("kline_visible", False))
        # 加载外观配置
        self.header_visible     = bool(cfg.get("header_visible", False))
        self.grid_visible       = bool(cfg.get("grid_visible", False))
        font_family             = cfg.get("font_family", "PingFang SC" if platform.system() == "Darwin" else "Microsoft YaHei")
        font_size               = int(cfg.get("font_size", 10))
        self.font               = QFont(font_family, max(8, min(15, font_size)))
        self.line_extra_px      = int(cfg.get("line_extra_px", 1))
        self.fg                 = QColor(cfg.get("fg", "#FFFFFF"))
        bg                      = cfg.get("bg", {"r":0,"g":0,"b":0,"a":191})
        self.bg                 = QColor(bg["r"],bg["g"],bg["b"],bg["a"])
        self.opacity_pct        = int(cfg.get("opacity_pct", 90))
        self.default_color      = bool(cfg.get("default_color", False))
        # 加载其他配置
        self.refresh_seconds    = int(cfg.get("refresh_seconds", 2))
        self.hotkey_enabled     = bool(cfg.get("hotkey_enabled", False))
        self.hotkey             = cfg.get("hotkey", "Ctrl+Alt+F")
        self.start_on_boot      = bool(cfg.get("start_on_boot", False))

        self.hotkey_triggered.connect(self.toggle_win)
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
        self.table.horizontalHeader().setVisible(self.header_visible)
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.setFont(self.font)
        self.table.horizontalHeader().setFont(self.font)
        self.table.verticalHeader().setMinimumSectionSize(1)
        self.table.verticalHeader().setDefaultSectionSize(1)
        self.table.horizontalHeader().setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.table.setTextElideMode(Qt.ElideNone)
        self.error_label = QLabel("", self.panel)
        self.error_label.setStyleSheet("color: #ff6666; padding: 2px 4px;")
        self.error_label.setVisible(False)
        self.vbox.addWidget(self.error_label)

        self.model = SimpleTableModel(headers=self.ALL_HEADERS, align_right_cols=[1,2,3,4,5])
        self.model.set_color_scheme(self.default_color, self.fg)
        self.table.setModel(self.model)

        self.k_delegate = KLineDelegate(self.table, base_pt=12)
        self.k_delegate.update_scheme(self.default_color, self.fg)
        self.k_delegate.set_point_size(self.font.pointSize())
        self.k_column_visible_index = None

        self.vbox.addWidget(self.table)

        for w in (self.panel, self.table, self.table.viewport(), self.table.horizontalHeader(), self.table.verticalHeader()):
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

        self.timer = QTimer(self)
        self.timer.setInterval(max(1, self.refresh_seconds)*1000)
        self.timer.timeout.connect(self._refresh_from_function)
        self.timer.start()
        self._refresh_from_function()
        self._defer_fit()

        self._keep_top_timer = QTimer(self)
        self._keep_top_timer.setInterval(1000)  # 每 1000ms 检查一次
        self._keep_top_timer.timeout.connect(self._ensure_on_top)
        self._keep_top_timer.start()

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
            "codes":                self.codes,
            "checked_codes":        self.checked_codes,

            "name_visible":         self.name_visible,
            "market_visible":       self.market_visible,
            "code_visible":         self.code_visible,
            "name_length":          self.name_length,
            "price_visible":        self.price_visible,
            "change_visible":       self.change_visible,
            "change_pct_visible":   self.change_pct_visible,
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
            "opacity_pct":      int(round(self.windowOpacity()*100)),
            "default_color":    self.default_color,

            "refresh_seconds":  self.refresh_seconds,
            "hotkey_enabled":   self.hotkey_enabled,
            "hotkey":           self.hotkey,
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
        total_w = self.table.verticalHeader().width() + 2*self.table.frameWidth()
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
    def _show_error(self, msg: str):
        """显示错误内容"""
        if self.k_column_visible_index is not None:
            self.table.setItemDelegateForColumn(self.k_column_visible_index, QStyledItemDelegate(self.table))
            self.k_column_visible_index = None

        text = str(msg) if msg is not None else ""
        self.error_label.setText(text)
        self.error_label.setVisible(True)
        self._defer_fit()

    def _clear_error(self):
        """清除顶部错误提示"""
        self.error_label.setVisible(False)
        self.error_label.setText("")

    def _project_columns(self, full_rows: list[dict], sign_data: list[dict]):
        # 从 ALL_HEADERS 中按显示顺序筛选已启用的列
        cols = [i for i, h in enumerate(self.ALL_HEADERS) if self.header_is_visible(h)]
        headers = [self.ALL_HEADERS[i] for i in cols]

        proj_rows, proj_meta = [], []
        for r, row in enumerate(full_rows):
            proj_rows.append([row[i] for i in headers])
            proj_meta.append(sign_data[r])

        # 右对齐：除了名称、K线、卖一外的所有列都右对齐
        right_cols = [i for i, h in enumerate(headers) if h not in ("名称", "K线", "卖一")]
        self.model.set_align_right_cols(right_cols)
        self.model.set_rows_headers(proj_rows, headers, meta=proj_meta)
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
        name += f"{code} " if self.code_visible else ""
        name += data["name"] if self.name_length == 0 else data["name"][:self.name_length]

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
            # 连续交易阶段
            buy_marker = "<" if pur_1 and data["current_price"]==pur_1 else " "
            sell_marker = ">" if sell_1 and data["current_price"]==sell_1 else " "
            b1_label = f"{int(data["purchaser_vol"][0] / 100)}{buy_marker}" if pur_1 else "-"
            s1_label = f"{sell_marker}{int(data["seller_vol"][0] / 100)}" if sell_1 else "-"
            b1_color_sign = 1
            s1_color_sign = -1

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

        # 数据返回
        precision = 3 if type == "基" else 2
        format_data = {
            self.ALL_HEADERS[0]: name, 
            self.ALL_HEADERS[1]: f"{data["current_price"]:.{precision}f}{arrow}", 
            self.ALL_HEADERS[2]: f"{change:+.{precision}f}",
            self.ALL_HEADERS[3]: f"{change_pct:+.2f}%",
            self.ALL_HEADERS[4]: b1_label,
            self.ALL_HEADERS[5]: s1_label,
            self.ALL_HEADERS[6]: f"{committee:+.2f}%",
            self.ALL_HEADERS[7]: _format_volume(data["deals_vol"]),
            self.ALL_HEADERS[8]: _format_amount(data["deals_amt"]),
            self.ALL_HEADERS[9]: f"{avg:.{precision}f}",
            self.ALL_HEADERS[10]: k_payload}
        sign = {
            self.ALL_HEADERS[0]: 0, 
            self.ALL_HEADERS[1]: (change > 0) - (change < 0),
            self.ALL_HEADERS[2]: (change > 0) - (change < 0),
            self.ALL_HEADERS[3]: (change > 0) - (change < 0),
            self.ALL_HEADERS[4]: b1_color_sign,
            self.ALL_HEADERS[5]: s1_color_sign,
            self.ALL_HEADERS[6]: (committee > 0) - (committee < 0),
            self.ALL_HEADERS[7]: 0,
            self.ALL_HEADERS[8]: 0,
            self.ALL_HEADERS[9]: (avg > data["prev_close"]) - (avg < data["prev_close"]),
            self.ALL_HEADERS[10]: 0}
        return format_data, sign

    def _refresh_from_function(self):
        try:
            ret, data = request_sina(self.checked_codes)
            full_rows = [self._format_data()]
        except Exception as e:
            self._show_error(str(e))
            return

        self._clear_error()
        self._project_columns(full_rows, sign)

    # ----- 应用设置 -----
    def set_codes(self, codes_list):
        seen = set()
        new = []
        for c in codes_list:
            s = str(c).strip().lower()
            if s and s not in seen:
                seen.add(s)
                new.append(s)
        if not new: 
            new = ["sh000001"]
        self.codes = new
        self._notify_change()
        self._refresh_from_function()

    def set_checked_codes(self, codes_list):
        seen = set()
        new = []
        for c in codes_list:
            s = str(c).strip().lower()
            if s and s not in seen:
                seen.add(s)
                new.append(s)
        if not new: 
            new = ["sh000001"]
        self.checked_codes = new
        self._notify_change()
        self._refresh_from_function()

    def set_code_names(self, code_names: dict):
        self.code_names = {str(k).strip().lower(): str(v).strip() for k, v in (code_names or {}).items() if str(k).strip()}
        self._notify_change()

    def set_type_visible(self, visible: bool):
        self.type_visible = bool(visible)
        self._notify_change()

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
        if name_len >=0:
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
        self.setWindowOpacity(p/100.0)
        self._defer_fit()
        self._notify_change()

    def set_font_size(self, pt: int):
        pt = max(8, min(15, int(pt)))
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

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_pos = e.globalPosition().toPoint() - self.frameGeometry().topLeft()
            self.setFocus(Qt.MouseFocusReason)

    def mouseMoveEvent(self, e):
        if getattr(self, "_drag_pos", None) and (e.buttons() & Qt.LeftButton):
            self.move(e.globalPosition().toPoint() - self._drag_pos)
            self._ensure_on_top()

    def mouseReleaseEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_pos = None
            self._ensure_on_top()
            self._notify_change()

    def mouseDoubleClickEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_pos = None
            self.hide()

    def eventFilter(self, obj, ev):
        if ev.type() == QEvent.MouseButtonDblClick and hasattr(ev, "button") and ev.button() == Qt.LeftButton:
            self._drag_pos = None
            self.hide()
            return True
        if ev.type() == QEvent.MouseButtonPress and hasattr(ev, "button") and ev.button() == Qt.LeftButton:
            self._drag_pos = ev.globalPosition().toPoint() - self.frameGeometry().topLeft()
            self.setFocus(Qt.MouseFocusReason)
            return True
        if ev.type() == QEvent.MouseMove and hasattr(ev, "buttons") and (ev.buttons() & Qt.LeftButton) and getattr(self, "_drag_pos", None):
            self.move(ev.globalPosition().toPoint() - self._drag_pos)
            return True
        if ev.type() == QEvent.MouseButtonRelease and hasattr(ev, "button") and ev.button() == Qt.LeftButton:
            self._drag_pos = None
            self._notify_change()
            return True
        return QWidget.eventFilter(self, obj, ev)

    def closeEvent(self, event): 
        event.ignore()
        self.hide()

    def showEvent(self, event):
        super().showEvent(event)
        if self.timer and not self.timer.isActive(): 
            self.timer.start()
        if self._keep_top_timer and not self._keep_top_timer.isActive():
            self._keep_top_timer.start()
        self._defer_fit()

    def hideEvent(self, event):
        super().hideEvent(event)
        if self.timer and self.timer.isActive(): 
            self.timer.stop()
        if self._keep_top_timer and self._keep_top_timer.isActive():
            self._keep_top_timer.stop()

    def _ensure_on_top(self):
        if not self.isVisible():
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

    def _register_hotkey(self):
        if platform.system() == "Darwin" or keyboard is None:
            return
        try:
            keyboard.remove_all_hotkeys()
        except Exception:
            pass
        try:
            keyboard.add_hotkey(self.hotkey.lower(), lambda: self.hotkey_triggered.emit())
        except Exception:
            pass

    def update_hotkey(self, new_hotkey: str):
        self.hotkey = new_hotkey.strip()
        self._register_hotkey()

    def toggle_win(self):
        if self.isVisible():
            self.hide()
        else:
            self.show()
