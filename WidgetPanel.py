import uuid
import requests, keyboard
from functools import partial

from PySide6.QtCore import Qt, QEvent, QTimer, Signal
from PySide6.QtGui import QFont, QAction, QColor
from PySide6.QtWidgets import QApplication, QWidget, QMenu, QVBoxLayout, QLabel, QTableView, QHeaderView, QAbstractItemView, QFrame

from Display import SimpleTableModel, KLineDelegate, FlashAwareDelegate

class FloatLabel(QWidget):
    hotkey_triggered = Signal()
    def __init__(self, cfg: dict):
        super().__init__()
        self._on_change = (lambda: None)
        self._open_settings_cb = None

        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setFocusPolicy(Qt.StrongFocus)

        # 加载配置
        codes_cfg               = cfg.get("codes",["sh000001"])             # 自选列表
        checked_codes_cfg       = cfg.get("checked_codes", cfg.get("visible_codes", codes_cfg))  # 在浮窗中显示的股票（新名 checked_codes，兼容 visible_codes）
        self.refresh_seconds    = int(cfg.get("refresh_seconds", 2))        # 刷新间隔
        flags_cfg               = cfg.get("flags", {})                      # 指标开关（字典格式）
        self.short_code         = bool(cfg.get("short_code", False))
        self.name_length        = int(cfg.get("name_length",0))
        # b1s1_display: 'qty'|'price'|'both'。兼容旧配置键 b1s1_price (bool)
        b1s1_display_cfg = cfg.get("b1s1_display", None)
        if isinstance(b1s1_display_cfg, str) and b1s1_display_cfg in ("qty", "price", "both"):
            self.b1s1_display = b1s1_display_cfg
        else:
            # 旧配置兼容：若 b1s1_price 为 True 则默认显示价格，否则显示数量
            self.b1s1_display = "price" if bool(cfg.get("b1s1_price", False)) else "qty"
        
        # 防止买一/卖一同步时触发重复处理
        self._syncing_b1s1 = False

        self.header_visible     = bool(cfg.get("header_visible", False))    # 表头可见
        self.grid_visible       = bool(cfg.get("grid_visible", False))      # 网格可见

        font_family             = cfg.get("font_family", "Microsoft YaHei") # 字体类型
        font_size               = int(cfg.get("font_size", 10))             # 字体大小
        self.line_extra_px      = int(cfg.get("line_extra_px", 1))          # 行间距
        self.fg                 = QColor(cfg.get("fg", "#FFFFFF"))        # 前景色
        bg                      = cfg.get("bg", {"r":0,"g":0,"b":0,"a":191})# 背景色
        self.opacity_pct        = int(cfg.get("opacity_pct", 90))           # 透明度
        self.default_color      = bool(cfg.get("default_color", False))     # 默认颜色模式

        self.hotkey             = cfg.get("hotkey", "Ctrl+Alt+F")           # 快捷键
        self.start_on_boot      = bool(cfg.get("start_on_boot", False))
        self.alarms             = self._normalize_alarms(cfg.get("alarms", []))
        self._active_alerts     = {}   # code -> alarm_id，正在闪烁待确认
        self._flash_on          = False
        self._display_row_codes = []   # 当前表格各行对应的完整代码

        # 设置初值
        self.codes = [str(c).strip() for c in codes_cfg if str(c).strip()]
        # 列标题列表（提前定义，供后续旧配置解析使用）
        self.ALL_HEADERS = ["代码", "名称", "现价", "涨跌值", "涨跌幅", "买一", "卖一", "委比", "成交量", "成交额", "均价", "K线"]

        # 列显示标志（独立属性）
        # 解析旧 flags 配置以做回退
        old_flags = {}
        if isinstance(flags_cfg, list):
            for i, h in enumerate(self.ALL_HEADERS):
                old_flags[h] = bool(flags_cfg[i]) if i < len(flags_cfg) else False
        elif isinstance(flags_cfg, dict):
            for h in self.ALL_HEADERS:
                old_flags[h] = bool(flags_cfg.get(h, False))

        # 新：为每一列创建独立的 bool 属性（优先读取新配置，否则回退到 old_flags）
        self.code_visible = bool(cfg.get("code_visible", old_flags.get("代码", False)))
        self.name_visible = bool(cfg.get("name_visible", old_flags.get("名称", False)))
        self.price_visible = bool(cfg.get("price_visible", old_flags.get("现价", False)))
        self.change_visible = bool(cfg.get("change_visible", old_flags.get("涨跌值", False)))
        self.change_pct_visible = bool(cfg.get("change_pct_visible", old_flags.get("涨跌幅", False)))
        # 买一/卖一 使用单一开关 b1s1_visible（用户要求不要拆分控制）
        self.b1s1_visible = bool(cfg.get("b1s1_visible", (old_flags.get("买一", False) or old_flags.get("卖一", False))))
        self.commi_visible = bool(cfg.get("commi_visible", old_flags.get("委比", False)))
        self.vol_visible = bool(cfg.get("vol_visible", old_flags.get("成交量", False)))
        self.amount_visible = bool(cfg.get("amount_visible", old_flags.get("成交额", False)))
        self.avg_visible = bool(cfg.get("avg_visible", old_flags.get("均价", False)))
        self.kline_visible = bool(cfg.get("kline_visible", old_flags.get("K线", False)))

        # 设置自选显示股票（新名 checked_codes）
        self.codes = [str(c).strip() for c in codes_cfg if str(c).strip()]
        self.checked_codes = [str(c).strip() for c in checked_codes_cfg if (str(c).strip() and str(c).strip() in self.codes)]
        self.font = QFont(font_family, max(8, min(15, font_size)))
        self.bg = QColor(bg["r"],bg["g"],bg["b"],bg["a"])
        
        
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
        self.table.setItemDelegate(FlashAwareDelegate(self.table))

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

        self._flash_timer = QTimer(self)
        self._flash_timer.setInterval(450)
        self._flash_timer.timeout.connect(self._on_flash_tick)

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
            "codes": self.codes,
            "checked_codes": self.checked_codes,
            "code_visible": bool(getattr(self, 'code_visible', False)),
            "name_visible": bool(getattr(self, 'name_visible', False)),
            "price_visible": bool(getattr(self, 'price_visible', False)),
            "change_visible": bool(getattr(self, 'change_visible', False)),
            "change_pct_visible": bool(getattr(self, 'change_pct_visible', False)),
            "b1s1_visible": bool(getattr(self, 'b1s1_visible', False)),
            "commi_visible": bool(getattr(self, 'commi_visible', False)),
            "vol_visible": bool(getattr(self, 'vol_visible', False)),
            "amount_visible": bool(getattr(self, 'amount_visible', False)),
            "avg_visible": bool(getattr(self, 'avg_visible', False)),
            "kline_visible": bool(getattr(self, 'kline_visible', False)),
            "short_code": self.short_code,
            "name_length": self.name_length,
            "b1s1_price": (getattr(self, 'b1s1_display', 'qty') == 'price'),
            "b1s1_display": getattr(self, 'b1s1_display', 'qty'),
            "header_visible": self.header_visible,
            "grid_visible": self.grid_visible,
            "refresh_seconds": self.refresh_seconds,
            "fg": self.fg.name(QColor.HexRgb),
            "bg": {"r": self.bg.red(), "g": self.bg.green(), "b": self.bg.blue(), "a": self.bg.alpha()},
            "opacity_pct": int(round(self.windowOpacity()*100)),
            "font_family": self.font.family(),
            "font_size": self.font.pointSize(),
            "line_extra_px": self.line_extra_px,
            "default_color": self.default_color,
            "pos": {"x": self.x(), "y": self.y()},
            "hotkey": self.hotkey,
            "start_on_boot": bool(self.start_on_boot),
            "alarms": [
                {"id": a["id"], "code": a["code"], "price": a["price"], "direction": a["direction"]}
                for a in self.alarms
            ],
        }

    def header_is_visible(self, header: str) -> bool:
        """返回指定列标题对应的独立可见属性值（替代旧的 flags 字典）。"""
        try:
            if header == "代码":
                return bool(getattr(self, 'code_visible', False))
            if header == "名称":
                return bool(getattr(self, 'name_visible', False))
            if header == "现价":
                return bool(getattr(self, 'price_visible', False))
            if header == "涨跌值":
                return bool(getattr(self, 'change_visible', False))
            if header == "涨跌幅":
                return bool(getattr(self, 'change_pct_visible', False))
            if header in ("买一", "卖一"):
                return bool(getattr(self, 'b1s1_visible', False))
            if header == "委比":
                return bool(getattr(self, 'commi_visible', False))
            if header == "成交量":
                return bool(getattr(self, 'vol_visible', False))
            if header == "成交额":
                return bool(getattr(self, 'amount_visible', False))
            if header == "均价":
                return bool(getattr(self, 'avg_visible', False))
            if header == "K线":
                return bool(getattr(self, 'kline_visible', False))
        except Exception:
            pass
        return False

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
        try:
            if self.k_column_visible_index is not None:
                self.table.setItemDelegateForColumn(self.k_column_visible_index, FlashAwareDelegate(self.table))
                self.k_column_visible_index = None
        except Exception:
            pass
        try:
            text = str(msg) if msg is not None else ""
            # 若是 requests 抛出的网络错误，显示更友好的中文提示
            if isinstance(msg, Exception):
                import requests as _req
                if isinstance(msg, _req.exceptions.RequestException):
                    text = "无网络连接"
        except Exception:
            text = str(msg)

        if hasattr(self, 'error_label'):
            self.error_label.setText(text)
            self.error_label.setVisible(True)
        self._defer_fit()

    def _clear_error(self):
        # 清除顶部错误提示
        if hasattr(self, 'error_label'):
            try:
                self.error_label.setVisible(False)
                self.error_label.setText("")
            except Exception:
                pass

    # ----- 数据来源：新浪财经 -----
    def _get_price(self, codes:list):
        label = ",".join([str(c).strip() for c in codes if str(c).strip()])
        if not label:
            raise Exception("暂无数据，请添加自选")

        price_data = []
        sign_data = []
        row_codes = []
        prices_by_code = {}
        url = 'https://hq.sinajs.cn/list=' + label
        headers = {'Referer': 'https://finance.sina.com.cn', 'User-Agent': 'Mozilla/5.0'}
        r = requests.get(url, headers=headers, timeout=3)
        r.encoding = 'gbk'
        for line in r.text.split('\n'):
            if not line or '"' not in line:
                continue
            heads = line.split('="')[0].split('_')
            parts = line.split('="')[1].split(',')
            if len(parts) < 30:
                continue

            code          = heads[2]
            name          = parts[0]
            opening_price = float(parts[1] or 0)   # 开盘
            prev_close    = float(parts[2] or 0)   # 昨收
            current_price = float(parts[3] or 0)   # 现价
            high_price    = float(parts[4] or 0)   # 当日最高
            low_price     = float(parts[5] or 0)   # 当日最低
            first_pur     = float(parts[6] or 0)   # 买一
            first_sell    = float(parts[7] or 0)   # 卖一
            deals_vol     = float(parts[8] or 0)   # 成交量
            deals_amt     = float(parts[9] or 0)   # 成交额
            purchaser     = [int(x or 0) for x in parts[10:19:2]]  # 买盘，股数
            pur_price     = [float(x or 0) for x in parts[11:20:2]]  # 买盘，价格
            seller        = [int(x or 0) for x in parts[20:29:2]]  # 卖盘，股数
            sel_price     = [float(x or 0) for x in parts[21:30:2]]  # 卖盘，价格
            update_date   = [int(x or 0) for x in parts[30].split('-')]  # 日期
            update_time   = [int(x or 0) for x in parts[31].split(':')]  # 时间

            etf = code[2] in ('1','5')

            # 构建买一/卖一数据及其颜色信息，并添加位置箭头
            b1_label = ""
            s1_label = ""
            b1_color_sign = 0  # 买一颜色：1红 0中性 -1绿
            s1_color_sign = 0  # 卖一颜色：1红 0中性 -1绿

            # 决定小数精度用于比较是否相等（避免浮点微小误差）
            dec = 3 if etf else 2
            def almost_eq(a, b):
                try:
                    return round(float(a), dec) == round(float(b), dec)
                except Exception:
                    return False

            # 标记：买一箭头位于右侧 '<'，卖一箭头位于左侧 '>'
            buy_marker = " "
            sell_marker = " "
            if first_pur > 0 and almost_eq(current_price, first_pur):
                buy_marker = "<"
            if first_sell > 0 and almost_eq(current_price, first_sell):
                sell_marker = ">"

            if first_pur == first_sell > 0:
                # 集合竞价：配对量 / 未配对量
                # 此处不显示成交方向箭头（竞价阶段无 <> 指示），且配对量和未配对量使用统一颜色规则
                current_price = first_sell  # 9:15 ~ 9:25; 14:57 ~ 15:00 竞价
                paired = seller[0]
                # unpaired_sign: >0 表示买方优势，<0 表示卖方优势
                unpaired_sign = -seller[1] if seller[1] > 0 else purchaser[1]
                # 显示数量（手）或价格或数量和价格（手数(价格)）
                paired_cnt = int(paired/100)
                unpaired_cnt = int(unpaired_sign/100)
                b_price = f"{first_pur:.3f}" if etf else f"{first_pur:.2f}"
                s_price = f"{first_sell:.3f}" if etf else f"{first_sell:.2f}"
                mode = getattr(self, 'b1s1_display', 'qty')
                if mode == 'price':
                    b1_label = f"{b_price}"
                    s1_label = f"{s_price}"
                elif mode == 'both':
                    b1_label = f"{paired_cnt:d}({b_price})"
                    s1_label = f"{unpaired_cnt:+d}({s_price})"
                else:
                    b1_label = f"{paired_cnt:d}"
                    s1_label = f"{unpaired_cnt:+d}"
                # 竞价颜色：根据未配对量的方向
                if unpaired_sign > 0:
                    b1_color_sign = 1
                    s1_color_sign = 1
                elif unpaired_sign < 0:
                    b1_color_sign = -1
                    s1_color_sign = -1
                else:
                    b1_color_sign = 0
                    s1_color_sign = 0
            else:
                # 连续竞价：买一数量/卖一数量
                if first_pur > 0:
                    cnt = f"{int(purchaser[0]/100)}"
                    b_price = f"{first_pur:.3f}" if etf else f"{first_pur:.2f}"
                    mode = getattr(self, 'b1s1_display', 'qty')
                    if mode == 'price':
                        b1_label = f"{b_price}{buy_marker}"
                    elif mode == 'both':
                        b1_label = f"{cnt}({b_price}){buy_marker}"
                    else:
                        b1_label = f"{cnt}{buy_marker}"
                else:
                    b1_label = f"-{buy_marker}"

                if first_sell > 0:
                    cnt = f"{int(seller[0]/100)}"
                    s_price = f"{first_sell:.3f}" if etf else f"{first_sell:.2f}"
                    mode = getattr(self, 'b1s1_display', 'qty')
                    if mode == 'price':
                        s1_label = f"{sell_marker}{s_price}"
                    elif mode == 'both':
                        s1_label = f"{sell_marker}{cnt}({s_price})"
                    else:
                        s1_label = f"{sell_marker}{cnt}"
                else:
                    s1_label = f"{sell_marker}-"

                # 连续竞价时：买一固定红色，卖一固定绿色
                b1_color_sign = 1
                s1_color_sign = -1
            
            if current_price == 0:
                current_price = prev_close # 9:00 ~ 9:15 无数据
            if opening_price == 0: 
                opening_price = current_price
                high_price = current_price
                low_price = current_price

            change = current_price - prev_close if prev_close else 0.0
            change_pct = (current_price / prev_close - 1) * 100 if prev_close else 0.0
            avg = (deals_amt / deals_vol) if deals_vol > 0 else prev_close # 均价
            p_sum, s_sum = sum(purchaser), sum(seller)
            committee = (100 * (p_sum - s_sum) / (p_sum + s_sum)) if (p_sum + s_sum) > 0 else 0.0 # 委比

            # 触及日高/低显示箭头
            arrow = " "
            if high_price > low_price:
                if current_price == high_price: arrow = "↑"
                elif current_price == low_price: arrow = "↓"

            k_payload = {"k": (opening_price, current_price, high_price, low_price, prev_close)}

            code_key = str(code).strip().lower()
            prices_by_code[code_key] = current_price
            row_codes.append(code_key)

            # "代码", "名称", "现价", "涨跌值", "涨跌幅", "买一", "卖一", "委比", "成交量", "成交额", "均价",  "K线"
            if code[2] not in ('1','5'):
                price_data.append([
                    code[2:] if self.short_code else code,
                    name if self.name_length == 0 else name[:self.name_length],
                    f"{current_price:.2f}{arrow}",
                    f"{change:+.2f}",
                    f"{change_pct:+.2f}%",
                    b1_label,
                    s1_label,
                    f"{committee:+.2f}%",
                    f"{deals_vol}" if deals_vol<1e4 else (f"{deals_vol/1e4:.2f}万" if deals_vol<1e8 else f"{deals_vol/1e8:.2f}亿"),
                    f"{deals_amt/1e4:.2f}万" if deals_amt<1e8 else (f"{deals_amt/1e8:.2f}亿" if deals_amt<1e12 else f"{deals_amt/1e12:.2f}万亿"),
                    f"{avg:.2f}",
                    k_payload
                ])
            else:
                price_data.append([
                    code[2:] if self.short_code else code,
                    name if self.name_length == 0 else name[:self.name_length],
                    f"{current_price:.3f}{arrow}",
                    f"{change:+.3f}",
                    f"{change_pct:+.2f}%",
                    b1_label,
                    s1_label,
                    f"{committee:+.2f}%",
                    f"{deals_vol}" if deals_vol<1e4 else (f"{deals_vol/1e4:.2f}万" if deals_vol<1e8 else f"{deals_vol/1e8:.2f}亿"),
                    f"{deals_amt/1e4:.2f}万" if deals_amt<1e8 else (f"{deals_amt/1e8:.2f}亿" if deals_amt<1e12 else f"{deals_amt/1e12:.2f}万亿"),
                    f"{avg:.3f}",
                    k_payload
                ])
            sign_data.append({
                "delta": (change > 0) - (change < 0), 
                "commi": (committee > 0) - (committee < 0),
                "avg": (avg > prev_close) - (avg < prev_close),
                "b1": b1_color_sign,
                "s1": s1_color_sign,
            })
        
        return price_data, sign_data, row_codes, prices_by_code

    def _project_columns(self, full_rows, sign_data):
        # 从 ALL_HEADERS 中按显示顺序筛选已启用的列
        cols = [i for i, h in enumerate(self.ALL_HEADERS) if self.header_is_visible(h)]
        headers = [self.ALL_HEADERS[i] for i in cols]

        proj_rows, proj_meta = [], []
        for r, row in enumerate(full_rows):
            proj_rows.append([row[i] for i in cols])
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
                self.table.setItemDelegateForColumn(self.k_column_visible_index, FlashAwareDelegate(self.table))
                self.k_column_visible_index = None

        self._fit_to_contents()

    def _codes_for_fetch(self):
        seen = set()
        out = []
        for c in list(self.checked_codes) + [a["code"] for a in self.alarms]:
            s = str(c).strip().lower()
            if s and s not in seen:
                seen.add(s)
                out.append(s)
        return out

    def _normalize_alarms(self, raw):
        out = []
        seen = set()
        for a in raw or []:
            if not isinstance(a, dict):
                continue
            code = str(a.get("code", "")).strip().lower()
            direction = str(a.get("direction", "")).strip().lower()
            if direction not in ("above", "below"):
                continue
            try:
                price = float(a.get("price"))
            except (TypeError, ValueError):
                continue
            if not code or price <= 0:
                continue
            key = (code, direction, round(price, 4))
            if key in seen:
                continue
            seen.add(key)
            aid = str(a.get("id") or uuid.uuid4())
            out.append({"id": aid, "code": code, "price": price, "direction": direction})
        return out

    def list_alarms(self):
        return [dict(a) for a in self.alarms]

    def add_alarm(self, code: str, price: float, direction: str):
        code = str(code).strip().lower()
        direction = str(direction).strip().lower()
        if direction not in ("above", "below"):
            return False
        try:
            price = float(price)
        except (TypeError, ValueError):
            return False
        if not code or price <= 0:
            return False
        for a in self.alarms:
            if a["code"] == code and a["direction"] == direction and abs(a["price"] - price) < 1e-9:
                return False
        self.alarms.append({
            "id": str(uuid.uuid4()),
            "code": code,
            "price": price,
            "direction": direction,
        })
        self._notify_change()
        # 隐藏时若有闹钟需继续刷新
        if not self.isVisible() and self.timer and not self.timer.isActive():
            self.timer.start()
        self._refresh_from_function()
        return True

    def remove_alarm(self, alarm_id: str):
        alarm_id = str(alarm_id)
        before = len(self.alarms)
        self.alarms = [a for a in self.alarms if a["id"] != alarm_id]
        # 若删掉的是正在闪烁的，同步清掉
        drop_codes = [c for c, aid in self._active_alerts.items() if aid == alarm_id]
        for c in drop_codes:
            self._active_alerts.pop(c, None)
        if not self._active_alerts:
            self._stop_flash()
        if len(self.alarms) != before:
            self._notify_change()
            self._maybe_pause_timer_when_hidden()
            return True
        return False

    def _evaluate_alarms(self, prices_by_code: dict):
        newly = False
        for alarm in list(self.alarms):
            code = alarm["code"]
            if code in self._active_alerts:
                continue
            if code not in prices_by_code:
                continue
            cur = float(prices_by_code[code])
            etf = len(code) > 2 and code[2] in ("1", "5")
            dec = 3 if etf else 2
            cur_r = round(cur, dec)
            target = round(float(alarm["price"]), dec)
            hit = False
            if alarm["direction"] == "above" and cur_r >= target:
                hit = True
            elif alarm["direction"] == "below" and cur_r <= target:
                hit = True
            if not hit:
                continue
            self._active_alerts[code] = alarm["id"]
            newly = True
            if code not in self.checked_codes:
                self.checked_codes.append(code)
        if newly:
            if not self.isVisible():
                self.show()
                self.raise_()
                self.activateWindow()
            self._start_flash()
            self._notify_change()

    def _start_flash(self):
        if not self._active_alerts:
            return
        self._flash_on = True
        if not self._flash_timer.isActive():
            self._flash_timer.start()
        self._sync_flash_rows()

    def _stop_flash(self):
        if self._flash_timer.isActive():
            self._flash_timer.stop()
        self._flash_on = False
        self.model.set_flash_state([], False)

    def _on_flash_tick(self):
        if not self._active_alerts:
            self._stop_flash()
            return
        self._flash_on = not self._flash_on
        self._sync_flash_rows()

    def _sync_flash_rows(self):
        rows = [i for i, c in enumerate(self._display_row_codes) if c in self._active_alerts]
        self.model.set_flash_state(rows, self._flash_on and bool(rows))

    def _dismiss_active_alerts(self) -> bool:
        if not self._active_alerts:
            return False
        ids = set(self._active_alerts.values())
        self.alarms = [a for a in self.alarms if a["id"] not in ids]
        self._active_alerts.clear()
        self._stop_flash()
        self._notify_change()
        self._maybe_pause_timer_when_hidden()
        return True

    def _maybe_pause_timer_when_hidden(self):
        if self.isVisible():
            return
        if self.alarms:
            if self.timer and not self.timer.isActive():
                self.timer.start()
            return
        if self.timer and self.timer.isActive():
            self.timer.stop()

    def _refresh_from_function(self):
        try:
            fetch_codes = self._codes_for_fetch()
            full_rows, sign, row_codes, prices = self._get_price(fetch_codes)
        except Exception as e:
            try:
                import requests as _req
                if isinstance(e, _req.exceptions.RequestException):
                    self._show_error(_req.exceptions.RequestException())
                else:
                    self._show_error(str(e))
            except Exception:
                self._show_error(str(e))
            return

        try:
            self._clear_error()
        except Exception:
            pass

        self._evaluate_alarms(prices)

        code_to_idx = {c: i for i, c in enumerate(row_codes)}
        disp_rows, disp_sign, disp_codes = [], [], []
        for c in self.checked_codes:
            s = str(c).strip().lower()
            if s in code_to_idx:
                i = code_to_idx[s]
                disp_rows.append(full_rows[i])
                disp_sign.append(sign[i])
                disp_codes.append(s)
        self._display_row_codes = disp_codes
        self._project_columns(disp_rows, disp_sign)
        self._sync_flash_rows()

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

    def set_flag(self, idx, checked: bool):
        """设置指标显示标志。idx 可以是整数索引（向后兼容）或列标题字符串"""
        # 兼容老版本：若传整数索引，转为列标题
        if isinstance(idx, int):
            if 0 <= idx < len(self.ALL_HEADERS):
                header = self.ALL_HEADERS[idx]
            else:
                return
        else:
            header = str(idx)
            if header not in self.ALL_HEADERS:
                return
        
        checked = bool(checked)
        prev = None
        try:
            if header == "代码":
                prev = bool(getattr(self, 'code_visible', False)); self.code_visible = checked
            elif header == "名称":
                prev = bool(getattr(self, 'name_visible', False)); self.name_visible = checked
            elif header == "现价":
                prev = bool(getattr(self, 'price_visible', False)); self.price_visible = checked
            elif header == "涨跌值":
                prev = bool(getattr(self, 'change_visible', False)); self.change_visible = checked
            elif header == "涨跌幅":
                prev = bool(getattr(self, 'change_pct_visible', False)); self.change_pct_visible = checked
            elif header in ("买一", "卖一"):
                prev = bool(getattr(self, 'b1s1_visible', False)); self.b1s1_visible = checked
            elif header == "委比":
                prev = bool(getattr(self, 'commi_visible', False)); self.commi_visible = checked
            elif header == "成交量":
                prev = bool(getattr(self, 'vol_visible', False)); self.vol_visible = checked
            elif header == "成交额":
                prev = bool(getattr(self, 'amount_visible', False)); self.amount_visible = checked
            elif header == "均价":
                prev = bool(getattr(self, 'avg_visible', False)); self.avg_visible = checked
            elif header == "K线":
                prev = bool(getattr(self, 'kline_visible', False)); self.kline_visible = checked
        except Exception:
            prev = None

        if prev is None or prev == checked:
            # 如果状态没有变化仍然返回（避免额外刷新）
            if prev == checked:
                return
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
        if seconds in {1,2,3,5,10,15,30,60}:
            self.refresh_seconds = seconds
            self.timer.setInterval(seconds*1000)
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

    def set_start_on_boot(self, enabled: bool):
        self.start_on_boot = bool(enabled)
        self._notify_change()
    
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
        act_open_settings.triggered.connect(lambda: self._open_settings_cb() if self._open_settings_cb else None)
        menu.addAction(act_open_settings)
        act_alarms = QAction("闹钟…", menu)
        act_alarms.triggered.connect(lambda: self._open_settings_cb(4) if self._open_settings_cb else None)
        menu.addAction(act_alarms)

        menu.addSeparator()
        menu.addAction(QAction("隐藏浮窗", menu, triggered=self.hide))
        menu.exec(event.globalPos())

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            if self._dismiss_active_alerts():
                return
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
            if self._dismiss_active_alerts():
                return True
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
        self._maybe_pause_timer_when_hidden()
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
        try:
            keyboard.remove_all_hotkeys()
        except Exception:
            pass
        keyboard.add_hotkey(self.hotkey.lower(), lambda: self.hotkey_triggered.emit())

    def update_hotkey(self, new_hotkey: str):
        self.hotkey = new_hotkey.strip()
        self._register_hotkey()

    def toggle_win(self):
        if self.isVisible():
            self.hide()
        else:
            self.show()