from PySide6.QtCore import Qt, QRect, QAbstractTableModel, QModelIndex
from PySide6.QtGui import QColor, QPainter, QPen, QBrush
from PySide6.QtWidgets import QStyledItemDelegate

# ----- 颜色配置 -----
UP_COLOR = QColor("#dd2100")
DOWN_COLOR = QColor("#019933")
NEUTRAL_COLOR = QColor("#494949")

class SimpleTableModel(QAbstractTableModel):
    """
    主浮窗表格数据与格式
    """
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
            elif header == "买一":
                sign = int(meta.get("b1", 0))
            elif header == "卖一":
                sign = int(meta.get("s1", 0))
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


class KLineDelegate(QStyledItemDelegate):
    """
    当日K线图，基于昨收，今开，最高，最低，实时价
    """
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