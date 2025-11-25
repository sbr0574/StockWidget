# filename: StockWidget.py
# python3 -m PyInstaller -F -w .\StockWidget.py --name StockWidget --icon .\StockWidget.ico --add-data ".\StockWidget.ico;."
import sys, os, json, ctypes, re, requests, keyboard, winreg
from functools import partial

from PySide6.QtCore import (
    Qt, QEvent, QTimer, QRect, QPoint, QAbstractTableModel, QModelIndex, Signal, QSize, QPropertyAnimation
)
from PySide6.QtGui import (
    QFont, QAction, QIcon, QColor, QFontDatabase, QPainter, QPen, QBrush, QKeySequence
)
from PySide6.QtWidgets import (
    QApplication, QWidget, QSystemTrayIcon, QMenu, QStyle, 
    QDialog, QVBoxLayout, QHBoxLayout, QGridLayout, QTabWidget, QPushButton, QSlider,
    QDialogButtonBox, QGroupBox, QLabel as QLabelW, QColorDialog, QComboBox, QTableView, QHeaderView, QAbstractItemView, QFrame,
    QStyledItemDelegate, QCheckBox, QListWidget, QListWidgetItem, QKeySequenceEdit
)

# ----- 程序与资源 -----
APP_NAME = "StockWidget"
APP_ICON_FILE = "StockWidget.ico"

def resource_path(rel_path):
    base = getattr(sys, "_MEIPASS", "")
    return os.path.join(base, rel_path)

def set_windows_app_user_model_id(appid: str):
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(appid)
    except Exception:
        pass

# ----- 配置存档 -----
CONFIG_DIR = os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), APP_NAME)
CONFIG_FILE = os.path.join(CONFIG_DIR, "SW_config.json")

def load_config():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_config(cfg: dict):
    if not os.path.exists(CONFIG_DIR):
        os.makedirs(CONFIG_DIR, exist_ok=True)
    tmp = CONFIG_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    os.replace(tmp, CONFIG_FILE)

# ----- 颜色配置 -----
UP_COLOR = QColor("#dd2100")
DOWN_COLOR = QColor("#019933")
NEUTRAL_COLOR = QColor("#494949")

# ===================== 表格 =====================
class SimpleTableModel(QAbstractTableModel):
    def __init__(self, rows=None, headers=None, align_right_cols=None, parent=None):
        super().__init__(parent)
        self._rows = rows or []
        self._headers = headers or []
        self._align_right = align_right_cols or []
        self.default_color = False
        self.fg_color = QColor("#FFFFFF")
        self._row_meta = []

    def set_color_scheme(self, default: bool, fg: QColor):
        self.default_color = bool(default)
        self.fg_color = QColor(fg)

    def rowCount(self, parent=QModelIndex()):
        return len(self._rows)
    
    def columnCount(self, parent=QModelIndex()):
        return len(self._rows[0]) if self._rows else len(self._headers)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        r, c = index.row(), index.column()
        cell = "" if c >= len(self._rows[r]) else self._rows[r][c]

        if role == Qt.UserRole:
            if isinstance(cell, dict) and "k" in cell:
                return cell["k"]
            return None

        if role == Qt.DisplayRole:
            return "" if isinstance(cell, dict) else str(cell)

        if role == Qt.TextAlignmentRole:
            return (Qt.AlignRight | Qt.AlignVCenter) if c in self._align_right else (Qt.AlignLeft | Qt.AlignVCenter)

        if role == Qt.ForegroundRole:
            if not self.default_color:
                return self.fg_color

            meta = self._row_meta[r] if 0 <= r < len(self._row_meta) else {}
            header = self._headers[c] if 0 <= c < len(self._headers) else ""
            sign = 0
            if header in ("涨跌值", "涨跌幅", "现价"):
                sign = int(meta.get("delta", 0))
            elif header == "委比":
                sign = int(meta.get("commi", 0))
            elif header == "均价":
                sign = int(meta.get("avg", 0))
            else:
                return self.fg_color

            if sign > 0:
                return UP_COLOR
            if sign < 0:
                return DOWN_COLOR
            return NEUTRAL_COLOR

        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and orientation == Qt.Horizontal and 0 <= section < len(self._headers):
            return self._headers[section]
        return None

    def set_rows_headers(self, rows, headers, meta=None):
        self.beginResetModel()
        self._rows = rows or []
        self._headers = headers or []
        self._row_meta = list(meta or [{} for _ in self._rows])
        self.endResetModel()

    def set_align_right_cols(self, cols_idx):
        self._align_right = set(cols_idx or [])

# ===================== 迷你 K 线 =====================
class KLineDelegate(QStyledItemDelegate):
    def __init__(self, parent=None, base_pt=12):
        super().__init__(parent)
        self.default_color = False
        self.fg = QColor("#FFFFFF")
        self.base_pt = max(1, int(base_pt))
        self.scale = 1.0  # 缩放

    def update_scheme(self, default_color: bool, fg: QColor):
        self.default_color = bool(default_color)
        self.fg = QColor(fg)

    def set_point_size(self, pt: int):
        self.scale = max(0.5, min(1.5, float(pt) / float(self.base_pt)))

    def paint(self, painter: QPainter, option, index):
        k = index.data(Qt.UserRole)
        if not k or not isinstance(k, tuple) or len(k) != 5:
            super().paint(painter, option, index)
            return

        o, c, h, l, p = k
        if h < l: h, l = l, h

        cell = option.rect
        rect = cell.adjusted(2, 2, -2, -2)

        sc = max(0.5, min(1.5, self.scale))
        vpad = max(2, int(rect.height() * (0.12 + 0.06 * (sc - 1))))   # ~12%~18%
        h_eff = max(2, rect.height() - 2 * vpad)
        krect = QRect(rect.left(), rect.top() + vpad, rect.width(), h_eff)

        def y_for(v):
            if h == l == p:
                y = 0.5
            else:
                y = (v - min(l,p)) / (max(h,p) - min(l,p))
            return krect.top() + (1 - y) * krect.height()

        y_o, y_c, y_h, y_l, y_p = (y_for(o), y_for(c), y_for(h), y_for(l), y_for(p))

        painter.save()
        painter.setClipRect(cell)
        painter.setRenderHint(QPainter.Antialiasing, True)

        body_w = max(5, min(int(krect.width() * 0.4 * sc), 10))
        x = krect.center().x()

        # 昨收虚线
        dash_col = QColor(NEUTRAL_COLOR if self.default_color else self.fg)
        dash_col.setAlpha(180)
        painter.setPen(QPen(dash_col, 1, Qt.DashLine))
        painter.drawLine(x - body_w, y_p, x + body_w, y_p)

        kcolor = self.fg
        if self.default_color:
            if c>o:
                kcolor = UP_COLOR
            elif c<o:
                kcolor = DOWN_COLOR
            else:
                kcolor = NEUTRAL_COLOR

        top, bot = min(y_o, y_c), max(y_o, y_c)
        body_h = max(2, bot - top)
        body_x = x - body_w // 2

        painter.setPen(QPen(kcolor, 1))
        if c != o:
            # 实体
            painter.drawRect(body_x, top, body_w, body_h)
        else:
            # 一字实体
            painter.drawLine(body_x, y_c, body_x+body_w, y_c)
        if y_h < top:
            # 上影线
            painter.drawLine(x, y_h, x, top)
        if y_l > bot:
            # 下影线
            painter.drawLine(x, bot, x, y_l)
        if c < o: 
            # 填充实体（空阳线）
            painter.fillRect(body_x, top, body_w, body_h, QBrush(kcolor))

        painter.restore()

# ===================== 主浮窗 =====================
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
        visible_codes_cfg       = cfg.get("visible_codes",codes_cfg)               # 在浮窗中显示的股票

        self.refresh_seconds    = int(cfg.get("refresh_seconds", 2))        # 刷新间隔
        flags                   = cfg.get("flags", [False, False, True, False, True, False, False, False, False, False, False]) # 指标开关
        self.short_code         = bool(cfg.get("short_code", False))
        self.name_length        = int(cfg.get("name_length",0))
        self.b1s1_price         = bool(cfg.get("b1s1_price", False))

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

        # 设置初值
        self.codes = [str(c).strip() for c in codes_cfg if str(c).strip()]
        self.visible_codes = [str(c).strip() for c in visible_codes_cfg if (str(c).strip() and str(c).strip() in self.codes)]
        self.font = QFont(font_family, max(8, min(15, font_size)))
        self.bg = QColor(bg["r"],bg["g"],bg["b"],bg["a"])
        self.ALL_HEADERS = ["代码", "名称", "现价", "涨跌值", "涨跌幅", "买一/卖一", "委比", "成交量", "成交额", "均价", "K线"]
        flags = list(flags) + [False] * (len(self.ALL_HEADERS) - len(flags))
        self.flags = [bool(x) for x in flags][:len(self.ALL_HEADERS)]
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
        self.table.setShowGrid(self.grid_visible)
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

        self.model = SimpleTableModel(headers=self.ALL_HEADERS, align_right_cols=[1,2,3,4,5])
        self.model.set_color_scheme(self.default_color, self.fg)
        self.table.setModel(self.model)

        self.k_delegate = KLineDelegate(self.table, base_pt=12)
        self.k_delegate.update_scheme(self.default_color, self.fg)
        self.k_delegate.set_point_size(self.font.pointSize())
        self.k_column_visible_index = None

        self.vbox.addWidget(self.table)

        # 拖动/双击
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
            "codes": self.codes,
            "visible_codes": self.visible_codes,
            "flags": self.flags,
            "short_code": self.short_code,
            "name_length": self.name_length,
            "b1s1_price": self.b1s1_price,
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
        }

    # ----- 外观/尺寸 -----
    def apply_style(self):
        r,g,b,a = self.bg.red(), self.bg.green(), self.bg.blue(), self.bg.alpha()
        self.panel.setStyleSheet(f"""
            QWidget#panel {{
                background: rgba({r},{g},{b},{a});
                border-radius: 5px;
            }}
            QTableView {{
                background: transparent; 
                border: none; 
                {"" if self.default_color else f"color: {self.fg.name()};"}
                gridline-color: rgba(180,180,180,120);
                outline: none;
            }}
            QHeaderView {{
                background-color: transparent;
            }}
            QHeaderView::section {{
                background: transparent; 
                border: none;
                font-weight: 600;
                {"" if self.default_color else f"color: {self.fg.name()};"}
                padding: 2px 4px;
            }}
            # QTableView::item {{
            #     border-right: 1px solid rgba(120,120,120,80);
            #     border-bottom: 1px solid rgba(120,120,120,80);
            # }}
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
        if self.k_column_visible_index is not None:
            self.table.setItemDelegateForColumn(self.k_column_visible_index, QStyledItemDelegate(self.table))
            self.k_column_visible_index = None
        self.model.set_align_right_cols([])
        self.model.set_rows_headers([[msg]], ["错误"])
        self._fit_to_contents()

    # ----- 数据来源：新浪财经 -----
    def _get_price(self, codes:list):
        label = ",".join([str(c).strip() for c in codes if str(c).strip()])
        if not label:
            raise Exception("暂无数据，请添加自选")

        price_data = []
        sign_data = []
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

            b1s1_label = ""
            etf = code[2] in ('1','5')
            if first_pur == first_sell > 0:
                # 集合竞价
                current_price = first_sell # 9:15 ~ 9:25; 14:57 ~ 15:00 竞价
                paired = seller[0]
                unpaired = -seller[1] if seller[1] > 0 else purchaser[1]
                b1s1_label = f"{int(paired/100):d} /{int(unpaired/100):+d}"
            else:
                # 连续竞价
                if first_pur > 0:
                    b1s1_label = f"{int(purchaser[0]/100)}" + ((f"({first_pur:.2f})" if not etf else f"({first_pur:.3f})") if self.b1s1_price else "")
                else:
                    b1s1_label = "-"
                if current_price == first_pur > 0:
                    b1s1_label += "</ "
                elif current_price == first_sell > 0:
                    b1s1_label += " />"
                else:
                    b1s1_label += " / "
                if first_sell > 0:
                    b1s1_label += f"{int(seller[0]/100)}" + ((f"({first_sell:.2f})" if not etf else f"({first_sell:.3f})") if self.b1s1_price else "")
                else:
                    b1s1_label += "-"
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

            # "代码", "名称", "现价", "涨跌值", "涨跌幅", "买一/卖一", "委比", "成交量", "成交额", "均价",  "K线"
            if code[2] not in ('1','5'):
                price_data.append([
                    code[2:] if self.short_code else code,
                    name if self.name_length == 0 else name[:self.name_length],
                    f"{current_price:.2f}{arrow}",
                    f"{change:+.2f}",
                    f"{change_pct:+.2f}%",
                    b1s1_label,
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
                    b1s1_label,
                    f"{deals_vol}" if deals_vol<1e4 else (f"{deals_vol/1e4:.2f}万" if deals_vol<1e8 else f"{deals_vol/1e8:.2f}亿"),
                    f"{deals_amt/1e4:.2f}万" if deals_amt<1e8 else (f"{deals_amt/1e8:.2f}亿" if deals_amt<1e12 else f"{deals_amt/1e12:.2f}万亿"),
                    f"{avg:.3f}",
                    k_payload
                ])
            sign_data.append({
                "delta": (change > 0) - (change < 0), 
                "commi": (committee > 0) - (committee < 0),
                "avg": (avg > prev_close) - (avg < prev_close),
            })
        
        return price_data, sign_data

    def _project_columns(self, full_rows, sign_data):
        cols = [i for i, f in enumerate(self.flags) if f]
        headers = [self.ALL_HEADERS[i] for i in cols]

        proj_rows, proj_meta = [], []
        for r, row in enumerate(full_rows):
            proj_rows.append([row[i] for i in cols])
            proj_meta.append(sign_data[r])

        right_cols = [i for i, h in enumerate(headers) if h not in ("名称","K线")]
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

    def _refresh_from_function(self):
        try:
            full_rows, sign = self._get_price(self.visible_codes)
        except Exception as e:
            self._show_error(str(e))
            return
        
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

    def set_visible_codes(self, codes_list):
        seen = set()
        new = []
        for c in codes_list:
            s = str(c).strip().lower()
            if s and s not in seen:
                seen.add(s)
                new.append(s)
        if not new: 
            new = ["sh000001"]
        self.visible_codes = new
        self._notify_change()
        self._refresh_from_function()

    def set_flag(self, idx: int, checked: bool):
        if 0 <= idx < len(self.ALL_HEADERS):
            self.flags[idx] = bool(checked)
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

    def set_b1s1_price(self, simple_order: bool):
        self.b1s1_price = bool(simple_order)
        self._notify_change()
        self._refresh_from_function()

    def set_header_visible(self, vis: bool):
        self.header_visible = bool(vis)
        self.table.horizontalHeader().setVisible(self.header_visible)
        self._notify_change()
        self._defer_fit()

    def set_grid_visible(self, vis: bool):
        self.grid_visible = bool(vis)
        self.table.setShowGrid(self.grid_visible)
        self._notify_change()
        self._defer_fit()

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
        for i, name in enumerate(self.ALL_HEADERS):
            act = QAction(name, sub_cols, checkable=True)
            act.setChecked(bool(self.flags[i]))
            act.toggled.connect(partial(self.set_flag, i))
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
        if self._open_settings_cb:
            act_open_settings.triggered.connect(self._open_settings_cb)
        else:
            def _fallback_open():
                dlg = SettingsDialog(self, self)
                self.place_dialog_away(dlg, self, margin=16)
                dlg.show()
            act_open_settings.triggered.connect(_fallback_open)
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
        """保持浮窗始终在最前"""
        if self.isVisible():
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
    
    # ----- 放置函数 -----
    def place_dialog_away(self, dlg, anchor_widget, margin=16):
        screen = anchor_widget.screen() or QApplication.primaryScreen()
        sg = screen.availableGeometry()
        ag: QRect = anchor_widget.frameGeometry()
        dlg.adjustSize(); dw, dh = dlg.width(), dlg.height()
        candidates = [
            QPoint(ag.right() + margin, ag.top()),
            QPoint(ag.left() - dw - margin, ag.top()),
            QPoint(max(sg.left(), ag.left()), ag.bottom() + margin),
            QPoint(max(sg.left(), ag.left()), ag.top() - dh - margin),
        ]
        for pt in candidates:
            if (pt.x() >= sg.left() and pt.y() >= sg.top() and
                pt.x() + dw <= sg.right() and pt.y() + dh <= sg.bottom()):
                dlg.move(pt); return
        cx = sg.left() + (sg.width() - dw)//2; cy = sg.top() + (sg.height() - dh)//2
        dlg.move(QPoint(cx, cy))

# ===================== 设置面板 =====================
class SettingsDialog(QDialog):
    def __init__(self, win: FloatLabel, parent: QWidget, app=None):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.win = win
        self.app = app
        self.setModal(False)

        main = QHBoxLayout(self)
        main.setContentsMargins(8, 8, 8, 8)
        main.setSpacing(8)
        self.tabs = QTabWidget()
        main.addWidget(self.tabs)

        self.tab_sizes = {
            0: QSize(300, 300),
            1: QSize(480, 420),
            2: QSize(360, 340),
            3: QSize(300, 160),
        }
        self._apply_tab_size(0)

        # ---- 第一页 ----
        tab_0 = QWidget()
        code_settings = QVBoxLayout(tab_0)

        # 1.自选列表
        g_codes = QGroupBox("自选列表")
        g_codes.setContentsMargins(3,12,3,6)
        lay_codes = QHBoxLayout(g_codes)
        lay_codes.setSpacing(6)
        # 1.1 代码列表
        self.list_codes = QListWidget()
        self.list_codes.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked | QAbstractItemView.EditKeyPressed)
        self.list_codes.setFixedWidth(150)
        for c in self.win.codes:
            it = QListWidgetItem(c)
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsEditable | Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            it.setCheckState(Qt.Checked if c in self.win.visible_codes else Qt.Unchecked)
            it.setData(Qt.UserRole, c)  # 记住上次有效值
            self.list_codes.addItem(it)
        # 1.2 操作按钮
        btn_col = QVBoxLayout()
        btn_col.setSpacing(4)
        self.btn_add = QPushButton("添加")
        self.btn_add.setFixedWidth(60)
        self.btn_del = QPushButton("删除")
        self.btn_del.setFixedWidth(60)
        self.btn_up  = QPushButton("上移")
        self.btn_up.setFixedWidth(60)
        self.btn_dn  = QPushButton("下移")
        self.btn_dn.setFixedWidth(60)
        for b in (self.btn_add, self.btn_del, self.btn_up, self.btn_dn):
            btn_col.addWidget(b)
        btn_col.addStretch(1)

        lay_codes.addWidget(self.list_codes, 1)
        lay_codes.addLayout(btn_col)
        code_settings.addWidget(g_codes)

        self.tabs.addTab(tab_0, "自选列表")

        # ---- 第二页 ----
        tab_1 = QWidget()
        data_settings = QVBoxLayout(tab_1)

        # 2.刷新间隔
        g_interval = QGroupBox("刷新间隔")
        g_interval.setContentsMargins(3,12,3,6)
        self.cmb_interval = QComboBox()
        self.cmb_interval.setFixedWidth(136)
        for s in [1,2,3,5,10,15,30,60]:
            self.cmb_interval.addItem(f"{s} 秒", userData=s)
        idx = self.cmb_interval.findData(self.win.refresh_seconds)
        self.cmb_interval.setCurrentIndex(idx if idx >= 0 else 1)
        v = QVBoxLayout(g_interval)
        v.setContentsMargins(6,6,6,6)
        v.addWidget(self.cmb_interval)
        data_settings.addWidget(g_interval)

        # 3.显示选项
        # 3.1复选框组
        g_flags = QGroupBox("显示指标")
        g_flags.setContentsMargins(3,12,3,6)
        gl_flags = QGridLayout(g_flags)
        self.cbs: list[QCheckBox] = []
        cb_texts = self.win.ALL_HEADERS

        g_flag_name = QGroupBox("名称")
        gl_flag_name = QGridLayout(g_flag_name)
        gl_flag_name.setHorizontalSpacing(6)
        gl_flag_name.setVerticalSpacing(6)
        for i in range(0,2):
            cb = QCheckBox(cb_texts[i])
            cb.setChecked(bool(self.win.flags[i]))
            cb.stateChanged.connect(partial(self._on_cb_changed, i))
            self.cbs.append(cb)
            gl_flag_name.addWidget(cb, i, 0)
        self.cb_short_code = QCheckBox("仅显示数字")
        self.cb_short_code.setChecked(bool(self.win.short_code))
        self.cb_short_code.setEnabled(self.win.flags[0])
        gl_flag_name.addWidget(self.cb_short_code, 0, 1)
        self.cmb_namelength = QComboBox()
        self.cmb_namelength.setFixedWidth(80)
        for l in [0, 1, 2, 3, 4]:
            self.cmb_namelength.addItem(f"前{l}个字" if l>0 else "完整显示", userData=l)
        idx_name = self.cmb_namelength.findData(self.win.name_length)
        self.cmb_namelength.setCurrentIndex(idx_name if idx_name>=0 else 1)
        self.cmb_namelength.setEnabled(self.win.flags[1])
        gl_flag_name.addWidget(self.cmb_namelength, 1, 1)
        gl_flags.addWidget(g_flag_name, 0, 0)

        g_flag_price = QGroupBox("价格")
        gl_flag_price = QGridLayout(g_flag_price)
        gl_flag_price.setHorizontalSpacing(6)
        gl_flag_price.setVerticalSpacing(6)
        for i in range(2,5):
            cb = QCheckBox(cb_texts[i])
            cb.setChecked(bool(self.win.flags[i]))
            cb.stateChanged.connect(partial(self._on_cb_changed, i))
            self.cbs.append(cb)
            gl_flag_price.addWidget(cb, i-2, 0)
        gl_flags.addWidget(g_flag_price, 1, 0)

        g_flag_order = QGroupBox("盘口")
        gl_flag_order = QGridLayout(g_flag_order)
        gl_flag_order.setHorizontalSpacing(6)
        gl_flag_order.setVerticalSpacing(6)
        for i in range(5,7):
            cb = QCheckBox(cb_texts[i])
            cb.setChecked(bool(self.win.flags[i]))
            cb.stateChanged.connect(partial(self._on_cb_changed, i))
            self.cbs.append(cb)
            gl_flag_order.addWidget(cb, i-5, 0)
        self.cb_b1s1_price = QCheckBox("显示价格")
        self.cb_b1s1_price.setChecked(bool(self.win.b1s1_price))
        self.cb_b1s1_price.setEnabled(self.win.flags[5])
        gl_flag_order.addWidget(self.cb_b1s1_price, 0, 1)
        gl_flags.addWidget(g_flag_order, 0, 1)

        g_flag_deal = QGroupBox("成交")
        gl_flag_deal = QGridLayout(g_flag_deal)
        gl_flag_deal.setHorizontalSpacing(6)
        gl_flag_deal.setVerticalSpacing(6)
        for i in range(7,10):
            cb = QCheckBox(cb_texts[i])
            cb.setChecked(bool(self.win.flags[i]))
            cb.stateChanged.connect(partial(self._on_cb_changed, i))
            self.cbs.append(cb)
            gl_flag_deal.addWidget(cb, i-7, 0)
        gl_flags.addWidget(g_flag_deal, 1, 1)

        g_flag_other = QGroupBox("其他")
        gl_flag_other = QGridLayout(g_flag_other)
        gl_flag_other.setHorizontalSpacing(6)
        gl_flag_other.setVerticalSpacing(6)
        for i in range(10,11):
            cb = QCheckBox(cb_texts[i])
            cb.setChecked(bool(self.win.flags[i]))
            cb.stateChanged.connect(partial(self._on_cb_changed, i))
            self.cbs.append(cb)
            gl_flag_other.addWidget(cb, i-10, 0)
        gl_flags.addWidget(g_flag_other, 2, 0)

        data_settings.addWidget(g_flags)

        self.tabs.addTab(tab_1, "显示数据")

        # ---- 第三页 ----
        tab_2 = QWidget()
        appearance_settings = QVBoxLayout(tab_2)

        # 表格外观
        g_table = QGroupBox("表格外观")
        g_table.setContentsMargins(3,12,3,6)
        gl_table = QGridLayout(g_table)
        gl_table.setHorizontalSpacing(6)
        gl_table.setVerticalSpacing(6)
        # 复选框
        self.chk_table_header = QCheckBox("显示表头")
        self.chk_table_header.setChecked(self.win.header_visible)
        self.chk_table_grid = QCheckBox("显示网格")
        self.chk_table_grid.setChecked(self.win.grid_visible)

        gl_table.addWidget(self.chk_table_header,0,0)
        gl_table.addWidget(self.chk_table_grid,0,1)
        appearance_settings.addWidget(g_table)

        # 3.颜色/透明度
        g_color = QGroupBox("颜色与透明度")
        g_color.setContentsMargins(3,12,3,6)
        gl_color = QGridLayout(g_color)
        gl_color.setHorizontalSpacing(6)
        gl_color.setVerticalSpacing(6)
        # 3.1 复选框：默认颜色
        self.chk_default_color = QCheckBox("默认颜色")
        self.chk_default_color.setChecked(self.win.default_color)
        # 3.2 按钮：文字颜色
        self.btn_fg = QPushButton("文字颜色…")
        self.btn_fg.setFixedWidth(90)
        self.btn_fg.setEnabled(not self.win.default_color)
        # 3.3 按钮：背景颜色
        self.btn_bg = QPushButton("背景颜色…")
        self.btn_bg.setFixedWidth(90)
        # 3.4 滑块：背景不透明度
        self.slider_bg_alpha = QSlider(Qt.Horizontal)
        self.slider_bg_alpha.setRange(0, 100)
        self.slider_bg_alpha.setMinimumWidth(150)
        self.slider_bg_alpha.setValue(int(round(self.win.bg.alpha()/2.55)))
        self.lbl_bg_alpha = QLabelW(f"{self.slider_bg_alpha.value()}%")
        # 3.5 滑块：整体不透明度
        self.slider_win_opacity = QSlider(Qt.Horizontal)
        self.slider_win_opacity.setRange(20, 100)
        self.slider_win_opacity.setMinimumWidth(150)
        self.slider_win_opacity.setValue(int(round(self.win.windowOpacity()*100)))
        self.lbl_win_opacity = QLabelW(f"{self.slider_win_opacity.value()}%")

        gl_color.addWidget(self.chk_default_color,0,0,1,2)
        gl_color.addWidget(self.btn_fg,0,2,1,2)
        gl_color.addWidget(self.btn_bg,0,4,1,2)
        gl_color.addWidget(QLabelW("背景不透明度："),1,0,1,2)
        gl_color.addWidget(self.slider_bg_alpha,1,2,1,3)
        gl_color.addWidget(self.lbl_bg_alpha,1,5,1,1)
        gl_color.addWidget(QLabelW("整体不透明度："),2,0,1,2)
        gl_color.addWidget(self.slider_win_opacity,2,2,1,3)
        gl_color.addWidget(self.lbl_win_opacity,2,5,1,1)
        appearance_settings.addWidget(g_color)

        # 4.字体/行距
        g_font = QGroupBox("字体与行距")
        g_font.setContentsMargins(3,12,3,6)
        gl_font = QGridLayout(g_font)
        gl_font.setHorizontalSpacing(6)
        gl_font.setVerticalSpacing(6)
        # 4.1 选项：字体
        self.cmb_family = QComboBox()
        self.cmb_family.setFixedWidth(200)
        for fam in sorted(QFontDatabase.families()):
            self.cmb_family.addItem(fam)
        fi = self.cmb_family.findText(self.win.font.family())
        self.cmb_family.setCurrentIndex(fi if fi >= 0 else 0)
        # 4.2 滑块：字号
        self.slider_font = QSlider(Qt.Horizontal)
        self.slider_font.setRange(8, 15)
        self.slider_font.setMinimumWidth(150)
        self.slider_font.setValue(self.win.font.pointSize())
        self.lbl_font = QLabelW(f"{self.slider_font.value()} pt")
        # 4.3 滑块：行间距
        self.slider_line = QSlider(Qt.Horizontal)
        self.slider_line.setRange(0, 20)
        self.slider_line.setMinimumWidth(150)
        self.slider_line.setValue(getattr(self.win,"line_extra_px",4))
        self.lbl_line = QLabelW(f"+{self.slider_line.value()} px")

        gl_font.addWidget(QLabelW("字体："),0,0,1,2)
        gl_font.addWidget(self.cmb_family,0,2,1,4)
        gl_font.addWidget(QLabelW("字号："),1,0,1,2)
        gl_font.addWidget(self.slider_font,1,2,1,3)
        gl_font.addWidget(self.lbl_font,1,5,1,1)
        gl_font.addWidget(QLabelW("行距："),2,0,1,2)
        gl_font.addWidget(self.slider_line,2,2,1,3)
        gl_font.addWidget(self.lbl_line,2,5,1,1)
        appearance_settings.addWidget(g_font)

        self.tabs.addTab(tab_2, "外观")

        # ---- 第四页 ----
        tab_3 = QWidget()
        other_settings = QVBoxLayout(tab_3)

        # 4.热键
        g_hotkey = QGroupBox("快捷键")
        g_hotkey.setContentsMargins(3,12,3,6)
        gl_hotkey = QGridLayout(g_hotkey)
        gl_hotkey.setHorizontalSpacing(6)
        gl_hotkey.setVerticalSpacing(6)
        gl_hotkey.addWidget(QLabelW("隐藏/显示浮窗："),0,0,1,1)
        self.edit_hotkey = QKeySequenceEdit()
        self.edit_hotkey.setKeySequence(QKeySequence(self.win.hotkey))
        gl_hotkey.addWidget(self.edit_hotkey,0,1)
        # 开机启动复选框（与快捷键同页）
        self.chk_start_on_boot = QCheckBox("开机启动")
        self.chk_start_on_boot.setChecked(bool(self.win.start_on_boot))
        other_settings.addWidget(self.chk_start_on_boot)
        other_settings.addWidget(g_hotkey)

        self.tabs.addTab(tab_3, "常规")

        # ---- 连接 ----
        # 连接：代码列表
        self.list_codes.itemChanged.connect(self._on_codes_changed)
        self.btn_add.clicked.connect(self._add_code)
        self.btn_del.clicked.connect(self._del_code)
        self.btn_up.clicked.connect(self._move_up)
        self.btn_dn.clicked.connect(self._move_down)
        # 连接：其它设置
        self.cmb_interval.currentIndexChanged.connect(self._on_interval_changed)
        self.cmb_namelength.currentIndexChanged.connect(self._on_name_length_changed)
        self.chk_default_color.toggled.connect(self._on_default_color_toggled)
        self.btn_fg.clicked.connect(self.pick_fg)
        self.btn_bg.clicked.connect(self.pick_bg)
        self.slider_bg_alpha.valueChanged.connect(self.apply_bg_alpha)
        self.slider_win_opacity.valueChanged.connect(self.apply_win_opacity)
        self.cmb_family.currentTextChanged.connect(self._on_family_changed)
        self.slider_font.valueChanged.connect(self.apply_font_size)
        self.slider_line.valueChanged.connect(self._on_line_changed)
        self.edit_hotkey.editingFinished.connect(self._on_hotkey_changed)
        self.chk_start_on_boot.toggled.connect(self._on_start_on_boot_toggled)
        self.chk_table_header.toggled.connect(self._on_header_toggled)
        self.chk_table_grid.toggled.connect(self._on_grid_toggled)
        self.tabs.currentChanged.connect(self._apply_tab_size)
        self.cb_b1s1_price.stateChanged.connect(self._on_b1s1_price_toggled)
        self.cb_short_code.stateChanged.connect(self._on_short_code_toggled)

    def _on_start_on_boot_toggled(self, checked: bool):
        try:
            self.win.set_start_on_boot(bool(checked))
            if hasattr(self, 'app') and self.app is not None:
                try:
                    self.app.set_start_on_boot(bool(checked))
                except Exception:
                    pass
        except Exception:
            pass

    # —— 代码规格化 —— #
    _re_full = re.compile(r'^(sh|sz|bj)\d+$')
    _re_6 = re.compile(r'^\d{6}$')

    def _normalize_code_or_none(self, s: str):
        s = (s or "").strip().lower()
        s = re.sub(r'[^a-z0-9]', '', s)
        if not s: return None
        if self._re_full.match(s): return s
        if self._re_6.match(s):
            if s[0] == '6' or s[0:2] == '90' or s[0] == '5':
                return 'sh' + s
            elif s[0] == '0' or s[0] == '3' or s[0] == '2' or s[0] == '1':
                return 'sz' + s
            elif s[0] == '8' or s[0] == '4' or s[0:2] == '92':
                return 'bj' + s
        return None

    def _collect_codes_from_list(self):
        codes = []
        seen = set()
        for i in range(self.list_codes.count()):
            txt = self.list_codes.item(i).text()
            norm = self._normalize_code_or_none(txt)
            if norm:
                if norm not in seen:
                    seen.add(norm); codes.append(norm)
                # 写回规范化文本
                it = self.list_codes.item(i)
                if it.text() != norm:
                    self.list_codes.blockSignals(True)
                    it.setText(norm)
                    it.setData(Qt.UserRole, norm)
                    self.list_codes.blockSignals(False)
            else:
                # 回退到上次有效值
                it = self.list_codes.item(i)
                prev = it.data(Qt.UserRole)
                if prev:
                    self.list_codes.blockSignals(True)
                    it.setText(prev)
                    self.list_codes.blockSignals(False)
                else:
                    # 没有上次有效值则删除
                    self.list_codes.takeItem(i)
                    return self._collect_codes_from_list()
        return codes

    def _on_codes_changed(self, _item):
        codes = self._collect_codes_from_list()
        self.win.set_codes(codes)
        visible_codes = [
            self.list_codes.item(i).text().split()[0]
            for i in range(self.list_codes.count())
            if self.list_codes.item(i).checkState() == Qt.Checked
        ]
        self.win.set_visible_codes(visible_codes)

    def _add_code(self):
        it = QListWidgetItem("sh000001")
        it.setFlags(it.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsEditable | Qt.ItemIsSelectable | Qt.ItemIsEnabled)
        it.setCheckState(Qt.Unchecked)
        it.setData(Qt.UserRole, "sh000001")
        self.list_codes.addItem(it)
        self.list_codes.setCurrentItem(it)
        self.list_codes.editItem(it)
        self._on_codes_changed(it)

    def _del_code(self):
        row = self.list_codes.currentRow()
        if row >= 0:
            self.list_codes.takeItem(row)
            self._on_codes_changed(None)

    def _move_up(self):
        row = self.list_codes.currentRow()
        if row > 0:
            it = self.list_codes.takeItem(row)
            self.list_codes.insertItem(row-1, it)
            self.list_codes.setCurrentRow(row-1)
            self._on_codes_changed(None)

    def _move_down(self):
        row = self.list_codes.currentRow()
        if 0 <= row < self.list_codes.count()-1:
            it = self.list_codes.takeItem(row)
            self.list_codes.insertItem(row+1, it)
            self.list_codes.setCurrentRow(row+1)
            self._on_codes_changed(None)

    # —— 其它槽 —— #
    def _on_interval_changed(self, idx):
        seconds = self.cmb_interval.currentData()
        if isinstance(seconds,int): 
            self.win.set_refresh_interval(seconds)

    def _on_default_color_toggled(self, checked: bool):
        self.btn_fg.setEnabled(not checked)
        self.win.set_default_color(bool(checked))
    
    def _on_grid_toggled(self, checked: bool):
        self.win.set_grid_visible(bool(checked))

    def _on_header_toggled(self, checked: bool):
        self.win.set_header_visible(bool(checked))

    def _on_cb_changed(self, idx: int, state: bool):
        self.win.set_flag(idx, state)
        if idx == 0:
            self.cb_short_code.setEnabled(state)
        elif idx == 1:
            self.cmb_namelength.setEnabled(state)
        elif idx == 5:
            self.cb_b1s1_price.setEnabled(state)
    
    def _on_short_code_toggled(self, checked: bool):
        self.win.set_code_type(checked)

    def _on_name_length_changed(self, length: int):
        self.win.set_name_length(length)

    def _on_b1s1_price_toggled(self, checked:bool):
        self.win.set_b1s1_price(checked)

    def _apply_tab_size(self, index: int):
        size = self.tab_sizes.get(index, QSize(400, 400))
        self.setFixedSize(size)

    def pick_fg(self):
        c = QColorDialog.getColor(self.win.fg, self, "选择文字颜色")
        if c.isValid(): self.win.set_fg_color(c)
    def pick_bg(self):
        base = QColor(self.win.bg)
        base.setAlpha(255)
        c = QColorDialog.getColor(base, self, "选择背景颜色")
        if c.isValid(): self.win.set_bg_rgb_keep_alpha(c)
    def apply_bg_alpha(self, v): 
        self.lbl_bg_alpha.setText(f"{v}%")
        self.win.set_bg_alpha_percent(v)
    def apply_win_opacity(self, v): 
        self.lbl_win_opacity.setText(f"{v}%")
        self.win.set_window_opacity_percent(v)
    def _on_family_changed(self, fam: str): 
        self.win.set_font_family(fam)
    def apply_font_size(self, v):
        self.lbl_font.setText(f"{v} pt")
        self.win.set_font_size(v)  # 同步 K 线缩放
    def _on_line_changed(self, v: int): 
        self.lbl_line.setText(f"+{v} px")
        self.win.set_line_extra(v)
    def _on_hotkey_changed(self):
        new_hotkey = self.edit_hotkey.keySequence().toString()
        try:
            self.win.update_hotkey(new_hotkey)
        except Exception:
            pass

# ===================== 应用 =====================
class App(QApplication):
    def __init__(self, argv):
        super().__init__(argv)
        self.setQuitOnLastWindowClosed(False)
        icon_path = resource_path(APP_ICON_FILE)
        app_icon = QIcon(icon_path) if os.path.exists(icon_path) else self.style().standardIcon(QStyle.SP_ComputerIcon)
        self.setWindowIcon(app_icon)

        cfg = load_config()
        self.win = FloatLabel(cfg)
        # Apply start-on-boot setting from config
        try:
            self.set_start_on_boot(bool(cfg.get("start_on_boot", False)))
        except Exception:
            pass
        self.win.set_on_change(self.save_now)
        self.win.set_open_settings_callback(self.open_settings)

        self.tray = QSystemTrayIcon(app_icon, self)
        self.tray.setToolTip(APP_NAME)
        menu = QMenu()
        menu.addAction(QAction("显示/隐藏 浮窗", self, triggered=self.toggle_win))
        menu.addAction(QAction("设置…", self, triggered=self.open_settings))
        menu.addSeparator()
        menu.addAction(QAction("退出", self, triggered=self.quit_app))
        self.tray.setContextMenu(menu)
        self.tray.activated.connect(self.on_tray_activated)
        self.tray.show()

        self.settings_dlg = None
        self.win.show()
        self.win.raise_()
        self.win.activateWindow()
        self.win.setFocus(Qt.ActiveWindowFocusReason)
        self.save_now()

    def on_tray_activated(self, reason):
        if reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.DoubleClick): self.toggle_win()

    def toggle_win(self):
        if self.win.isVisible():
            self.win.hide()
        else:
            self.win.show()
            self.win.raise_()
            self.win.activateWindow()
            self.win.setFocus(Qt.ActiveWindowFocusReason)
        self.save_now()

    def open_settings(self):
        if self.settings_dlg and self.settings_dlg.isVisible():
            self.settings_dlg.raise_()
            self.settings_dlg.activateWindow()
            return
        self.settings_dlg = SettingsDialog(self.win, self.win, app=self)
        self.win.place_dialog_away(self.settings_dlg, self.win, margin=16)
        self.settings_dlg.show()
        self.settings_dlg.raise_()
        self.settings_dlg.activateWindow()

    def quit_app(self):
        self.tray.hide()
        self.save_now()
        keyboard.unhook_all_hotkeys()
        sys.exit(0)

    def save_now(self):
        cfg = self.win.current_config()
        save_config(cfg)

    def set_start_on_boot(self, enabled: bool):
        """Enable or disable Windows startup by writing/removing Run key in HKCU."""
        try:
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            name = APP_NAME
            if enabled:
                if getattr(sys, 'frozen', False):
                    cmd = f'"{sys.executable}"'
                else:
                    cmd = f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}"'
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                    winreg.SetValueEx(key, name, 0, winreg.REG_SZ, cmd)
            else:
                try:
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                        winreg.DeleteValue(key, name)
                except OSError:
                    # value not present
                    pass
        except Exception:
            pass

if __name__ == "__main__":
    set_windows_app_user_model_id(f"{APP_NAME}.1")
    app = App(sys.argv)
    sys.exit(app.exec())
