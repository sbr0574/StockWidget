# filename: StockWidget.py
# python3 -m PyInstaller -F -w .\StockWidget.py --name StockWidget --icon .\StockWidget.ico --add-data ".\StockWidget.ico;."
import sys, os, json, ctypes, re, requests, keyboard
from functools import partial

from PySide6.QtCore import (
    Qt, QEvent, QTimer, QRect, QPoint, QAbstractTableModel, QModelIndex, Signal
)
from PySide6.QtGui import (
    QFont, QAction, QIcon, QColor, QFontDatabase, QPainter, QPen, QBrush
)
from PySide6.QtWidgets import (
    QApplication, QWidget, QSystemTrayIcon, QMenu, QStyle, 
    QDialog, QVBoxLayout, QHBoxLayout, QGridLayout, QTabWidget, QPushButton, QLineEdit, QSlider,
    QDialogButtonBox, QGroupBox, QLabel as QLabelW, QColorDialog, QComboBox, QTableView, QHeaderView, QAbstractItemView, QFrame,
    QStyledItemDelegate, QCheckBox, QListWidget, QListWidgetItem
)

# ----- 程序与资源 -----
APP_NAME = "StockWidget"
APP_ICON_FILE = "StockWidget.ico"

def resource_path(rel_path: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.abspath("."))
    p = os.path.join(base, rel_path)
    return p if os.path.exists(p) else rel_path

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

# 颜色配置
UP_COLOR = QColor("#dd2100")
DOWN_COLOR = QColor("#019933")
NEUTRAL_COLOR = QColor("#494949")

# ----- 数据来源：新浪财经 -----
def get_price(codes:list) -> list[list]:
    label = ",".join([str(c).strip() for c in codes if str(c).strip()])
    if not label:
        return []

    rows = []
    url = 'https://hq.sinajs.cn/list=' + label
    headers = {'Referer': 'https://finance.sina.com.cn', 'User-Agent': 'Mozilla/5.0'}
    r = requests.get(url, headers=headers, timeout=3)
    r.encoding = 'gbk'
    for line in r.text.split('\n'):
        if not line or '"' not in line:
            continue
        heads = line.split('"')[0].split('_')
        parts = line.split('"')[1].split(',')
        if len(parts) < 30:
            continue
        is_etf = False
        if heads[2][2] in ('1','5'):
            is_etf = True
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
        purchaser     = [float(x or 0) for x in parts[10:19:2]]
        seller        = [float(x or 0) for x in parts[20:29:2]]

        if current_price == 0:
            # 9:15 ~ 9:25 竞价
            current_price = max(first_pur, first_sell)
            # 9:00 ~ 9:15 无数据
            if current_price == 0:
                raise Exception("暂无数据")
            opening_price = current_price
        if high_price == 0 or low_price == 0:
            high_price = current_price
            low_price = current_price

        change = current_price - prev_close if prev_close else 0.0
        change_pct = (current_price / prev_close - 1) * 100 if prev_close else 0.0

        seal = 0
        if first_pur == 0 and seller and seller[0] > 0:
            # 跌停
            seal = int(seller[0] / 100)
        elif first_sell == 0 and purchaser and purchaser[0] > 0:
            # 涨停
            seal = int(purchaser[0] / 100)

        avg = (deals_amt / deals_vol) if deals_vol > 0 else 0.0 # 均价
        p_sum, s_sum = sum(purchaser), sum(seller)
        committee = (100 * (p_sum - s_sum) / (p_sum + s_sum)) if (p_sum + s_sum) > 0 else 0.0 # 委比

        # 触及日高/低显示箭头
        arrow = " "
        if high_price != low_price:
            if current_price == high_price: arrow = "↑"
            elif current_price == low_price: arrow = "↓"

        k_payload = {"k": (opening_price, current_price, high_price, low_price, prev_close)}

        if not is_etf:
            rows.append([
                name,
                f"{current_price:.2f}{arrow}",
                f"{change:+.2f}",
                f"{change_pct:+.2f}%",
                f"{seal}" if 0 < seal < 10000 else (f"{(seal/10000):.1f}w" if seal >=10000 else (f"{avg:.2f}" if avg > 0 else "")),
                f"{committee:+.2f}%",
                k_payload
            ])
        else:
            rows.append([
                name,
                f"{current_price:.3f}{arrow}",
                f"{change:+.3f}",
                f"{change_pct:+.2f}%",
                f"{seal}" if 0 < seal < 10000 else (f"{(seal/10000):.1f}w" if seal >=10000 else (f"{avg:.3f}" if avg > 0 else "")),
                f"{committee:+.2f}%",
                k_payload
            ])
    return rows

# ----- 放置函数 -----
def place_dialog_away(dlg, anchor_widget, margin=16):
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

# ===================== 表格 =====================
class SimpleTableModel(QAbstractTableModel):
    def __init__(self, rows=None, headers=None, align_right_cols=None, parent=None):
        super().__init__(parent)
        self._rows = rows or []
        self._headers = headers or []
        self._align_right = set(align_right_cols or [])
        self.colorful_mode = False
        self.fg_color = QColor("#FFFFFF")
        self.up_color = UP_COLOR
        self.down_color = DOWN_COLOR
        self.neutral_color = NEUTRAL_COLOR
        self._row_meta = []

    def set_color_scheme(self, colorful: bool, fg: QColor):
        self.colorful_mode = bool(colorful)
        self.fg_color = QColor(fg)

    def rowCount(self, parent=QModelIndex()): return len(self._rows)
    def columnCount(self, parent=QModelIndex()): return len(self._rows[0]) if self._rows else len(self._headers)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid(): return None
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
            if not self.colorful_mode:
                return self.fg_color

            meta = self._row_meta[r] if 0 <= r < len(self._row_meta) else {}
            header = self._headers[c] if 0 <= c < len(self._headers) else ""
            sign = 0
            if header in ("涨跌值", "涨跌幅", "现价"): sign = int(meta.get("delta_sign", 0))
            elif header == "委比":                    sign = int(meta.get("weibi_sign", 0))
            else: return self.fg_color

            if sign > 0:  return self.up_color
            if sign < 0:  return self.down_color
            return self.neutral_color

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

    def set_align_right_cols(self, cols_idx): self._align_right = set(cols_idx or [])

# ===================== 迷你 K 线 =====================
class KLineDelegate(QStyledItemDelegate):
    def __init__(self, parent=None, base_pt=12):
        super().__init__(parent)
        self.colorful = False
        self.fg = QColor("#FFFFFF")
        self.up_color = UP_COLOR
        self.down_color = DOWN_COLOR
        self.neutral_color = NEUTRAL_COLOR
        self.base_pt = max(1, int(base_pt))
        self.scale = 1.0  # 缩放

    def update_scheme(self, colorful: bool, fg: QColor):
        self.colorful = bool(colorful)
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
        dash_col = QColor(self.neutral_color if self.colorful else self.fg)
        dash_col.setAlpha(180)
        painter.setPen(QPen(dash_col, 1, Qt.DashLine))
        painter.drawLine(x - body_w, y_p, x + body_w, y_p)

        kcolor = self.fg
        if self.colorful:
            if c>o:
                kcolor = self.up_color
            elif c<o:
                kcolor = self.down_color
            else:
                kcolor = self.neutral_color

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
    def __init__(self, cfg: dict):
        super().__init__()
        self._on_change = (lambda: None)
        self._open_settings_cb = None

        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setFocusPolicy(Qt.StrongFocus)

        # 加载配置/初值
        codes_from_cfg = cfg.get("codes",["sh000001"])
        self.codes = [str(c).strip() for c in codes_from_cfg if str(c).strip()]
        self.header_visible = bool(cfg.get("header_visible", False))
        self.refresh_seconds = int(cfg.get("refresh_seconds", 2))
        self.font_family = cfg.get("font_family", "Microsoft YaHei")
        font_size = int(cfg.get("font_size", 10))
        self.font = QFont(self.font_family, max(8, min(15, font_size)))
        self.line_extra_px = int(cfg.get("line_extra_px", 1))
        self.fg = QColor(cfg.get("fg", "#FFFFFF"))
        bg = cfg.get("bg", {"r":0,"g":0,"b":0,"a":191})
        self.bg = QColor(bg["r"],bg["g"],bg["b"],bg["a"])
        self.opacity_pct = int(cfg.get("opacity_pct", 90))
        self.colorful_mode = bool(cfg.get("colorful_mode", False))
        self.ALL_HEADERS = ["名称", "现价", "涨跌值", "涨跌幅", "均/封", "委比", "K线"]
        default_flags = [False, True, False, True, False, False, False]
        flags = cfg.get("flags", default_flags)
        if not isinstance(flags, list): flags = default_flags
        flags = list(flags) + [False] * (len(self.ALL_HEADERS) - len(flags))
        self.flags = [bool(x) for x in flags][:len(self.ALL_HEADERS)]

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

        self.model = SimpleTableModel(headers=self.ALL_HEADERS, align_right_cols=[1,2,3,4,5])
        self.model.set_color_scheme(self.colorful_mode, self.fg)
        self.table.setModel(self.model)

        self.k_delegate = KLineDelegate(self.table, base_pt=12)
        self.k_delegate.update_scheme(self.colorful_mode, self.fg)
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
            "flags": self.flags,
            "header_visible": self.header_visible,
            "refresh_seconds": self.refresh_seconds,
            "fg": self.fg.name(QColor.HexRgb),
            "bg": {"r": self.bg.red(), "g": self.bg.green(), "b": self.bg.blue(), "a": self.bg.alpha()},
            "opacity_pct": int(round(self.windowOpacity()*100)),
            "font_family": self.font.family(),
            "font_size": self.font.pointSize(),
            "line_extra_px": self.line_extra_px,
            "colorful_mode": self.colorful_mode,
            "pos": {"x": self.x(), "y": self.y()},
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
                {"" if self.colorful_mode else f"color: {self.fg.name()};"}
                gridline-color: transparent; 
                outline: none;
            }}
            QHeaderView {{
                background-color: transparent;
            }}
            QHeaderView::section {{
                background: transparent; 
                border: none;
                font-weight: 600;
                {"" if self.colorful_mode else f"color: {self.fg.name()};"}
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
        if self.k_column_visible_index is not None:
            self.table.setItemDelegateForColumn(self.k_column_visible_index, QStyledItemDelegate(self.table))
            self.k_column_visible_index = None
        self.model.set_align_right_cols([])
        self.model.set_rows_headers([[msg]], ["错误"])
        self._fit_to_contents()

    def _project_columns(self, full_rows):
        meta_list = []
        for row in full_rows:
            delta_sign = 0; weibi_sign = 0
            try:
                o,c,h,l,p = row[6]["k"]
                if c > p: delta_sign = +1
                elif c < p: delta_sign = -1
                wb = str(row[5])
                if wb.startswith("+"): weibi_sign = +1
                elif wb.startswith("-"): weibi_sign = -1
            except Exception: pass
            meta_list.append({"delta_sign": delta_sign, "weibi_sign": weibi_sign})

        cols = [i for i, f in enumerate(self.flags) if f] or [1,3]
        headers = [self.ALL_HEADERS[i] for i in cols]

        proj_rows, proj_meta = [], []
        for r, row in enumerate(full_rows):
            rr = list(row) + [""] * max(0, len(self.ALL_HEADERS) - len(row))
            proj_rows.append([rr[i] for i in cols]); proj_meta.append(meta_list[r])

        right_cols = [i for i, h in enumerate(headers) if h not in ("名称","K线")]
        self.model.set_align_right_cols(right_cols)
        self.model.set_rows_headers(proj_rows, headers, meta=proj_meta)
        self.model.set_color_scheme(self.colorful_mode, self.fg)

        if "K线" in headers:
            col = headers.index("K线")
            self.k_column_visible_index = col
            self.k_delegate.update_scheme(self.colorful_mode, self.fg)
            self.k_delegate.set_point_size(self.font.pointSize())
            self.table.setItemDelegateForColumn(col, self.k_delegate)
        else:
            if self.k_column_visible_index is not None:
                self.table.setItemDelegateForColumn(self.k_column_visible_index, QStyledItemDelegate(self.table))
                self.k_column_visible_index = None

        self._fit_to_contents()

    def _refresh_from_function(self):
        try:
            full_rows = get_price(self.codes)
        except Exception as e:
            self._show_error(str(e))
            return
        
        self._project_columns(full_rows)

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

    def set_flag(self, idx: int, checked: bool):
        if 0 <= idx < len(self.ALL_HEADERS):
            self.flags[idx] = bool(checked)
            self._notify_change()
            self._refresh_from_function()

    def set_header_visible(self, vis: bool):
        self.header_visible = bool(vis)
        self.table.horizontalHeader().setVisible(self.header_visible)
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

    def set_colorful_mode(self, enabled: bool):
        self.colorful_mode = bool(enabled)
        self.model.set_color_scheme(self.colorful_mode, self.fg)
        self.k_delegate.update_scheme(self.colorful_mode, self.fg)
        self.apply_style()
        self._notify_change()
        self._defer_fit()
    
    # ----- 交互 -----
    def contextMenuEvent(self, event):
        menu = QMenu(self)
        sub_cols = QMenu("显示指标", menu)
        headers = ["名称", "现价", "涨跌值", "涨跌幅", "均价/封单", "五档委比", "K线"]
        for i, name in enumerate(headers):
            act = QAction(name, sub_cols, checkable=True)
            act.setChecked(bool(self.flags[i]))
            act.toggled.connect(partial(self.set_flag, i))
            sub_cols.addAction(act)
        menu.addMenu(sub_cols)

        act_header = QAction("显示表头", menu, checkable=True)
        act_header.setChecked(self.header_visible)
        act_header.toggled.connect(self.set_header_visible)
        menu.addAction(act_header)

        act_color = QAction("默认颜色", menu, checkable=True)
        act_color.setChecked(self.colorful_mode)
        act_color.toggled.connect(self.set_colorful_mode)
        menu.addAction(act_color)

        menu.addSeparator()
        act_open_settings = QAction("设置…", menu)
        if self._open_settings_cb:
            act_open_settings.triggered.connect(self._open_settings_cb)
        else:
            def _fallback_open():
                dlg = SettingsDialog(self, self)
                place_dialog_away(dlg, self, margin=16)
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

    def mouseReleaseEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_pos = None
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
        self._defer_fit()

    def hideEvent(self, event):
        super().hideEvent(event)
        if self.timer and self.timer.isActive(): 
            self.timer.stop()

# ===================== 设置面板 =====================
class SettingsDialog(QDialog):
    def __init__(self, win: FloatLabel, parent: QWidget):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.win = win

        main = QHBoxLayout(self)
        main.setContentsMargins(8, 8, 8, 8)
        main.setSpacing(8)
        self.setFixedSize(500,330)
        self.setModal(False)

        # ---- 第一列 ----
        v_line_0 = QVBoxLayout()
        main.addLayout(v_line_0)

        # 1.关注列表
        g_codes = QGroupBox("关注列表")
        g_codes.setContentsMargins(3,12,3,6)
        lay_codes = QHBoxLayout(g_codes)
        lay_codes.setSpacing(6)
        # 1.1 代码列表
        self.list_codes = QListWidget()
        self.list_codes.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked | QAbstractItemView.EditKeyPressed)
        self.list_codes.setFixedWidth(70)
        for c in self.win.codes:
            it = QListWidgetItem(c)
            it.setFlags(it.flags() | Qt.ItemIsEditable | Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            it.setData(Qt.UserRole, c)  # 记住上次有效值
            self.list_codes.addItem(it)
        # 1.2 编辑按钮
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
        v_line_0.addWidget(g_codes)

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
        v_line_0.addWidget(g_interval)

        # ---- 第二列 ----
        v_line_1 = QVBoxLayout()
        main.addLayout(v_line_1)

        # 3.颜色/透明度
        g_color = QGroupBox("颜色与透明度")
        g_color.setContentsMargins(3,12,3,6)
        gl_color = QGridLayout(g_color)
        gl_color.setHorizontalSpacing(6)
        gl_color.setVerticalSpacing(6)
        # 3.1 复选框：默认颜色
        self.chk_colorful = QCheckBox("默认颜色")
        self.chk_colorful.setChecked(self.win.colorful_mode)
        # 3.2 按钮：文字颜色
        self.btn_fg = QPushButton("文字颜色…")
        self.btn_fg.setFixedWidth(90)
        self.btn_fg.setEnabled(not self.win.colorful_mode)
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

        gl_color.addWidget(self.chk_colorful,0,0,1,2)
        gl_color.addWidget(self.btn_fg,0,2,1,2)
        gl_color.addWidget(self.btn_bg,0,4,1,2)
        gl_color.addWidget(QLabelW("背景不透明度："),1,0,1,2)
        gl_color.addWidget(self.slider_bg_alpha,1,2,1,3)
        gl_color.addWidget(self.lbl_bg_alpha,1,5,1,1)
        gl_color.addWidget(QLabelW("整体不透明度："),2,0,1,2)
        gl_color.addWidget(self.slider_win_opacity,2,2,1,3)
        gl_color.addWidget(self.lbl_win_opacity,2,5,1,1)
        v_line_1.addWidget(g_color)

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
        v_line_1.addWidget(g_font)

        # 5.显示选项
        # 5.1复选框组
        g_flags = QGroupBox("显示指标")
        g_flags.setContentsMargins(3,12,3,6)
        gl_flags = QGridLayout(g_flags)
        gl_flags.setHorizontalSpacing(6)
        gl_flags.setVerticalSpacing(6)
        self.cbs: list[QCheckBox] = []
        cb_texts = ["名称", "现价", "涨跌值", "涨跌幅", "均价/封单", "五档委比", "K线"]
        for i in range(len(cb_texts)):
            cb = QCheckBox(cb_texts[i])
            cb.setChecked(bool(self.win.flags[i]))
            cb.stateChanged.connect(partial(self.win.set_flag, i))
            self.cbs.append(cb)
            row, col = divmod(i, 4)
            gl_flags.addWidget(cb, row, col)
        v_line_1.addWidget(g_flags)

        # ---- 连接 ----
        # 连接：代码列表
        self.list_codes.itemChanged.connect(self._on_codes_changed)
        self.btn_add.clicked.connect(self._add_code)
        self.btn_del.clicked.connect(self._del_code)
        self.btn_up.clicked.connect(self._move_up)
        self.btn_dn.clicked.connect(self._move_down)
        # 连接：其它设置
        self.cmb_interval.currentIndexChanged.connect(self._on_interval_changed)
        self.chk_colorful.toggled.connect(self._on_colorful_toggled)
        self.btn_fg.clicked.connect(self.pick_fg)
        self.btn_bg.clicked.connect(self.pick_bg)
        self.slider_bg_alpha.valueChanged.connect(self.apply_bg_alpha)
        self.slider_win_opacity.valueChanged.connect(self.apply_win_opacity)
        self.cmb_family.currentTextChanged.connect(self._on_family_changed)
        self.slider_font.valueChanged.connect(self.apply_font_size)
        self.slider_line.valueChanged.connect(self._on_line_changed)

    # —— 代码规格化 —— #
    _re_full = re.compile(r'^(sh|sz|bj)\d+$')
    _re_6 = re.compile(r'^\d{6}$')
    # _re_5 = re.compile(r'^\d{5}$')  # 港股

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
        # if self._re_5.match(s):
        #     return 'hk' + s
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

    def _add_code(self):
        it = QListWidgetItem("sh000001")
        it.setFlags(it.flags() | Qt.ItemIsEditable | Qt.ItemIsSelectable | Qt.ItemIsEnabled)
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
        if isinstance(seconds,int): self.win.set_refresh_interval(seconds)

    def _on_colorful_toggled(self, checked: bool):
        self.btn_fg.setEnabled(not checked)
        self.win.set_colorful_mode(bool(checked))
    
    def _on_cb_changed(self, idx: int, state: int):
        self.win.set_flag(idx, state == Qt.Checked)

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

# ===================== 应用 =====================
class App(QApplication):
    hotkey_triggered = Signal()

    def __init__(self, argv):
        super().__init__(argv)
        self.setQuitOnLastWindowClosed(False)
        icon_path = resource_path(APP_ICON_FILE)
        app_icon = QIcon(icon_path) if os.path.exists(icon_path) else self.style().standardIcon(QStyle.SP_ComputerIcon)
        self.setWindowIcon(app_icon)

        cfg = load_config()
        self.win = FloatLabel(cfg)
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

        self.hotkey = cfg.get("hotkey", "alt+z")
        self.hotkey_triggered.connect(self.toggle_win)
        self._register_hotkey()

        self.settings_dlg = None
        self.win.show()
        self.win.raise_()
        self.win.activateWindow()
        self.win.setFocus(Qt.ActiveWindowFocusReason)
        self.save_now()

    def _register_hotkey(self):
        try:
            keyboard.remove_all_hotkeys()
        except Exception:
            pass
        keyboard.add_hotkey(self.hotkey, lambda: self.hotkey_triggered.emit())

    def update_hotkey(self, new_hotkey: str):
        self.hotkey = new_hotkey.strip().lower()
        self._register_hotkey()
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
        self.settings_dlg = SettingsDialog(self.win, self.win)
        place_dialog_away(self.settings_dlg, self.win, margin=16)
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
        cfg["hotkey"] = self.hotkey
        save_config(cfg)

if __name__ == "__main__":
    set_windows_app_user_model_id(f"{APP_NAME}.1")
    app = App(sys.argv)
    sys.exit(app.exec())
