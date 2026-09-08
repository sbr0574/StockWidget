from PySide6.QtCore import Qt, QRect, QAbstractTableModel, QModelIndex
from PySide6.QtGui import QColor, QPainter, QPen, QBrush
from PySide6.QtWidgets import QStyledItemDelegate

# ----- 颜色配置 -----
DEFAULT_TEXT_COLOR = QColor("#FFFFFF")
DEFAULT_UP_COLOR = QColor("#dd2100")
DEFAULT_DOWN_COLOR = QColor("#019933")
DEFAULT_NEUTRAL_COLOR = QColor("#494949")

COLOR_ROLE_TEXT = "text"
COLOR_ROLE_UP = "up"
COLOR_ROLE_DOWN = "down"
COLOR_ROLE_NEUTRAL = "neutral"


def direction_color_role(value) -> str:
    """把正负方向转换为明确颜色角色，避免 0 同时表示普通文本和平盘。"""
    if value > 0:
        return COLOR_ROLE_UP
    if value < 0:
        return COLOR_ROLE_DOWN
    return COLOR_ROLE_NEUTRAL


class SimpleTableModel(QAbstractTableModel):
    """
    主浮窗表格数据与格式
    """
    def __init__(self, rows=None, headers=None, align_right_cols=None, parent=None):
        super().__init__(parent)
        self.unicolor = True
        self.text_color = QColor(DEFAULT_TEXT_COLOR)
        self.up_color = QColor(DEFAULT_UP_COLOR)
        self.down_color = QColor(DEFAULT_DOWN_COLOR)
        self.neutral_color = QColor(DEFAULT_NEUTRAL_COLOR)
        self._rows = rows or []
        self._headers = headers or []
        self._align_right = align_right_cols or []
        self._color_roles = []

    def set_colors(
        self,
        unicolor: bool,
        text_color: QColor,
        up_color: QColor,
        down_color: QColor,
        neutral_color: QColor,
    ):
        self.unicolor = bool(unicolor)
        self.text_color = QColor(text_color)
        self.up_color = QColor(up_color)
        self.down_color = QColor(down_color)
        self.neutral_color = QColor(neutral_color)
        if self.rowCount() and self.columnCount():
            self.dataChanged.emit(
                self.index(0, 0),
                self.index(self.rowCount() - 1, self.columnCount() - 1),
                [Qt.ItemDataRole.ForegroundRole],
            )
        if self.columnCount():
            self.headerDataChanged.emit(
                Qt.Orientation.Horizontal, 0, self.columnCount() - 1
            )

    def rowCount(self, parent=QModelIndex()):
        return len(self._rows)
    
    def columnCount(self, parent=QModelIndex()):
        return len(self._headers)

    def data(self, index, role=Qt.DisplayRole):
        r, c = index.row(), index.column()
        cell = self._rows[r][c]

        if role == Qt.UserRole:
            if isinstance(cell, dict) and "k" in cell:
                return cell["k"]
            return None

        if role == Qt.DisplayRole:
            return "" if isinstance(cell, dict) else str(cell)

        if role == Qt.TextAlignmentRole:
            return (Qt.AlignRight | Qt.AlignVCenter) if c in self._align_right else (Qt.AlignLeft | Qt.AlignVCenter)

        if role == Qt.ForegroundRole:
            if self.unicolor:
                return self.text_color

            if r >= len(self._color_roles) or c >= len(self._color_roles[r]):
                return self.text_color
            color_role = self._color_roles[r][c]
            if color_role == COLOR_ROLE_UP:
                return self.up_color
            if color_role == COLOR_ROLE_DOWN:
                return self.down_color
            if color_role == COLOR_ROLE_NEUTRAL:
                return self.neutral_color
            return self.text_color

        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and 0 <= section < len(self._headers):
            if role == Qt.DisplayRole:
                return self._headers[section]
            if role == Qt.ForegroundRole:
                return self.text_color
        return None

    def set_rows_headers(self, rows, headers, color_roles):
        self.beginResetModel()
        self._rows = rows
        self._headers = headers
        self._color_roles = color_roles
        self.endResetModel()

    def set_align_right_cols(self, cols_idx):
        self._align_right = set(cols_idx or [])


class KLineDelegate(QStyledItemDelegate):
    """
    当日K线图，基于昨收，今开，最高，最低，实时价
    """
    def __init__(self, parent=None, base_pt=12):
        super().__init__(parent)
        self.unicolor = True
        self.text_color = QColor(DEFAULT_TEXT_COLOR)
        self.up_color = QColor(DEFAULT_UP_COLOR)
        self.down_color = QColor(DEFAULT_DOWN_COLOR)
        self.neutral_color = QColor(DEFAULT_NEUTRAL_COLOR)
        self.base_pt = max(1, int(base_pt))
        self.scale = 1.0  # 缩放

    def set_colors(
        self,
        unicolor: bool,
        text_color: QColor,
        up_color: QColor,
        down_color: QColor,
        neutral_color: QColor,
    ):
        self.unicolor = bool(unicolor)
        self.text_color = QColor(text_color)
        self.up_color = QColor(up_color)
        self.down_color = QColor(down_color)
        self.neutral_color = QColor(neutral_color)

    def candle_color(self, opening, closing) -> QColor:
        if self.unicolor:
            return QColor(self.text_color)
        if closing > opening:
            return QColor(self.up_color)
        if closing < opening:
            return QColor(self.down_color)
        return QColor(self.neutral_color)

    def reference_color(self) -> QColor:
        return QColor(self.text_color if self.unicolor else self.neutral_color)

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
        dash_col = self.reference_color()
        dash_col.setAlpha(180)
        painter.setPen(QPen(dash_col, 1, Qt.DashLine))
        painter.drawLine(x - body_w, y_p, x + body_w, y_p)

        kcolor = self.candle_color(o, c)

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
