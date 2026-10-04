"""行情视图组件：表格模型、指标绘制、表头、统一网格和共用分页控件。"""

from PySide6.QtCore import (
    Qt,
    QRect,
    QRectF,
    QSize,
    QAbstractTableModel,
    QModelIndex,
    QPointF,
    QPoint,
    Signal,
)
from PySide6.QtGui import (
    QColor,
    QFontMetrics,
    QPainter,
    QPen,
    QBrush,
    QPalette,
    QPolygonF,
    QPainterPath,
)
from PySide6.QtWidgets import (
    QStyledItemDelegate,
    QStyleOptionViewItem,
    QApplication,
    QProxyStyle,
    QStyle,
    QStyleOptionHeader,
    QTableView,
    QWidget,
)

from stockwidget.core.quote_presentation import (
    BidAskCell,
    COLOR_ROLE_TEXT,
    COLOR_ROLE_UP,
    COLOR_ROLE_DOWN,
    COLOR_ROLE_NEUTRAL,
)


# ----- 颜色配置 -----
DEFAULT_TEXT_COLOR = QColor("#FFFFFF")
DEFAULT_UP_COLOR = QColor("#dd2100")
DEFAULT_DOWN_COLOR = QColor("#019933")
DEFAULT_NEUTRAL_COLOR = QColor("#494949")


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
        if not index.isValid():
            return None
        r, c = index.row(), index.column()
        cell = self._rows[r][c]

        if role == Qt.UserRole:
            if isinstance(cell, BidAskCell):
                return cell
            if isinstance(cell, dict) and "k" in cell:
                return cell["k"]
            return None

        if role == Qt.DisplayRole:
            return "" if isinstance(cell, dict) else str(cell)

        if role == Qt.TextAlignmentRole:
            if isinstance(cell, BidAskCell):
                return Qt.AlignCenter
            return (Qt.AlignRight | Qt.AlignVCenter) if c in self._align_right else (Qt.AlignLeft | Qt.AlignVCenter)

        if role == Qt.ForegroundRole:
            if r >= len(self._color_roles) or c >= len(self._color_roles[r]):
                return self.text_color
            return self.color_for_role(self._color_roles[r][c])

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

    def color_for_role(self, role):
        if self.unicolor:
            return self.text_color
        return {COLOR_ROLE_UP: self.up_color, COLOR_ROLE_DOWN: self.down_color,
                COLOR_ROLE_NEUTRAL: self.neutral_color}.get(role, self.text_color)


def bid_ask_width(cell, font, padding=4):
    """Symmetric halves keep the axis fixed even when the digit counts differ."""
    fm = QFontMetrics(font)
    half = max(fm.horizontalAdvance(cell.buy), fm.horizontalAdvance(cell.sell))
    slot = max(fm.horizontalAdvance(symbol) for symbol in ("▶", "◀")) + 2
    return half * 2 + slot + padding * 2 + 2


def paint_bid_ask(painter, rect, font, cell, model, padding=4):
    fm = QFontMetrics(font)
    gap = min(max(fm.horizontalAdvance(symbol) for symbol in ("▶", "◀")) + 2,
              max(0, rect.width() - padding * 2))
    center = rect.x() + rect.width() / 2
    left_edge, right_edge = center - gap / 2, center + gap / 2
    left = QRectF(rect.x() + padding, rect.y(), max(0, left_edge - rect.x() - padding), rect.height())
    right = QRectF(right_edge, rect.y(), max(0, rect.right() + 1 - padding - right_edge), rect.height())
    middle = QRectF(left_edge, rect.y(), gap, rect.height())
    painter.save()
    painter.setClipRect(rect)
    painter.setFont(font)
    for area, text, role, alignment in (
        (left, cell.buy, cell.buy_role, Qt.AlignRight),
        (right, cell.sell, cell.sell_role, Qt.AlignLeft),
        (middle, cell.marker, COLOR_ROLE_TEXT, Qt.AlignHCenter),
    ):
        painter.setPen(model.color_for_role(role))
        painter.drawText(area, int(alignment | Qt.AlignVCenter),
                         fm.elidedText(text, Qt.ElideRight, max(0, int(area.width()))))
    painter.restore()


class QuoteItemDelegate(QStyledItemDelegate):
    """Render the paired metric in one cell; ordinary metrics use Qt's delegate."""
    def sizeHint(self, option, index):
        cell = index.data(Qt.UserRole)
        if not isinstance(cell, BidAskCell):
            return super().sizeHint(option, index)
        styled = QStyleOptionViewItem(option)
        self.initStyleOption(styled, index)
        return QSize(bid_ask_width(cell, styled.font), QFontMetrics(styled.font).height())

    def paint(self, painter, option, index):
        cell = index.data(Qt.UserRole)
        if not isinstance(cell, BidAskCell):
            return super().paint(painter, option, index)
        styled = QStyleOptionViewItem(option)
        self.initStyleOption(styled, index)
        paint_bid_ask(painter, option.rect, styled.font, cell, index.model())


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


class SortIndicatorStyle(QProxyStyle):
    def __init__(self, header):
        # 按名称创建独立样式，避免代理接管 QApplication 共享样式的所有权。
        super().__init__(QApplication.style().objectName())
        self.setParent(header)

    def sizeFromContents(self, contents_type, option, size, widget=None):
        if contents_type == QStyle.CT_HeaderSection:
            # QHeaderView 会为所有列附加排序标记；尺寸计算忽略它，避免整表变宽。
            option = QStyleOptionHeader(option)
            option.sortIndicator = QStyleOptionHeader.SortIndicator.None_
        return super().sizeFromContents(contents_type, option, size, widget)

    def drawPrimitive(self, element, option, painter, widget=None):
        # 原生箭头锚定列边界；改在文字绘制后按实际文字位置绘制。
        if element != QStyle.PE_IndicatorHeaderArrow:
            return super().drawPrimitive(element, option, painter, widget)

    def drawItemText(self, painter, rect, flags, palette, enabled, text, text_role=QPalette.NoRole):
        super().drawItemText(painter, rect, flags, palette, enabled, text, text_role)
        header = self.parent()
        if (
            text_role != QPalette.ButtonText
            or not header.isSortIndicatorShown()
            or header.logicalIndexAt(rect.center()) != header.sortIndicatorSection()
        ):
            return

        # 使用样式表已解析的字体与对齐方式；列变宽时，箭头仍紧跟文字。
        text_rect = self.itemTextRect(painter.fontMetrics(), rect, flags, enabled, text)
        arrow_rect = QRectF(text_rect.right() + 3, text_rect.center().y() - 2, 5, 4)
        center = arrow_rect.center()
        half_width = arrow_rect.width() / 2
        half_height = arrow_rect.height() / 2
        direction = 1 if header.sortIndicatorOrder() == Qt.DescendingOrder else -1
        base_y = center.y() - direction * half_height
        arrow = QPolygonF([
            QPointF(center.x() - half_width, base_y),
            QPointF(center.x() + half_width, base_y),
            QPointF(center.x(), center.y() + direction * half_height),
        ])
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(Qt.NoPen)
        painter.setBrush(palette.brush(text_role))
        painter.drawPolygon(arrow)
        painter.restore()


class _TableGrid(QWidget):
    def __init__(self, table):
        super().__init__(table)
        self.table = table
        self.color = QColor()
        # The grid is only paint; hit-testing and dragging reach the usual regions.
        self.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.setFocusPolicy(Qt.NoFocus)
        self.hide()

    def paintEvent(self, event):
        table, model = self.table, self.table.model()
        if model is None or self.width() < 2 or self.height() < 2:
            return
        bounds = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        path = QPainterPath()
        path.addRoundedRect(bounds, 3, 3)
        viewport = table.viewport()
        origin = viewport.mapTo(table, QPoint())
        header = table.horizontalHeader()
        for c in range(model.columnCount()):
            if table.isColumnHidden(c):
                continue
            x = origin.x() + table.columnViewportPosition(c) + table.columnWidth(c) - 0.5
            if bounds.left() < x < bounds.right():
                path.moveTo(x, bounds.top())
                path.lineTo(x, bounds.bottom())
        if header.isVisible() and bounds.top() < origin.y() - 0.5 < bounds.bottom():
            path.moveTo(bounds.left(), origin.y() - 0.5)
            path.lineTo(bounds.right(), origin.y() - 0.5)
        for r in range(model.rowCount()):
            if table.isRowHidden(r):
                continue
            y = origin.y() + table.rowViewportPosition(r) + table.rowHeight(r) - 0.5
            if bounds.top() < y < bounds.bottom():
                path.moveTo(bounds.left(), y)
                path.lineTo(bounds.right(), y)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        pen = QPen(self.color, 1)
        pen.setCapStyle(Qt.FlatCap)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)
        # One stroke prevents alpha accumulating where cell and frame edges meet.
        painter.drawPath(path)
        painter.end()


class QuoteTableView(QTableView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.grid_layer = _TableGrid(self)
        for header in (self.horizontalHeader(), self.verticalHeader()):
            header.sectionResized.connect(self.grid_layer.update)
            header.sectionMoved.connect(self.grid_layer.update)
            header.geometriesChanged.connect(self.grid_layer.update)
        for scrollbar in (self.horizontalScrollBar(), self.verticalScrollBar()):
            scrollbar.valueChanged.connect(self.grid_layer.update)

    def setModel(self, model):
        old = self.model()
        if old is not None:
            old.modelReset.disconnect(self.grid_layer.update)
            old.layoutChanged.disconnect(self.grid_layer.update)
        super().setModel(model)
        if model is not None:
            model.modelReset.connect(self.grid_layer.update)
            model.layoutChanged.connect(self.grid_layer.update)
        self.grid_layer.update()

    def set_grid_appearance(self, enabled, color):
        self.setShowGrid(False)
        self.grid_layer.color = QColor(color)
        self.grid_layer.setGeometry(self.rect())
        self.grid_layer.setVisible(enabled)
        self.grid_layer.raise_()
        self.grid_layer.update()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.grid_layer.setGeometry(self.rect())
        self.grid_layer.raise_()


def pager_regions(rect):
    h = max(1, rect.height() // 3)
    return (QRect(rect.x(), rect.y(), rect.width(), h),
            QRect(rect.x(), rect.y() + h, rect.width(), rect.height() - 2 * h),
            QRect(rect.x(), rect.bottom() - h + 1, rect.width(), h))


def pager_hit(rect, x, y):
    previous, _, following = pager_regions(rect)
    if previous.contains(x, y):
        return -1
    if following.contains(x, y):
        return 1
    return 0


def pager_size(page, font, minimum_width=38):
    fm = QFontMetrics(font)
    digits = len(str(page.count if page else 1))
    # Reserve the widest digits for both numbers, independent of the current page.
    digit_width = max(fm.horizontalAdvance(str(n)) for n in range(10))
    width = max(minimum_width, 2 * digits * digit_width + fm.horizontalAdvance("/") + 8)
    return QSize(width, max(14, fm.height() + 4) * 3)


def paint_pager(painter, rect, page, color, font):
    painter.save()
    painter.setClipRect(rect)
    label = f"{page.index + 1}/{page.count}"
    painter.setFont(font)
    painter.setPen(color)
    previous, middle, following = pager_regions(rect)
    painter.drawText(middle, Qt.AlignCenter, label)
    painter.setRenderHint(QPainter.Antialiasing)
    for area, direction in ((previous, -1), (following, 1)):
        center = QRectF(area).center()
        half_width = min(QFontMetrics(font).height() / 4, area.width() / 4, area.height() / 3)
        path = QPainterPath()
        path.moveTo(center.x() - half_width, center.y() - direction * half_width / 2)
        path.lineTo(center.x(), center.y() + direction * half_width / 2)
        path.lineTo(center.x() + half_width, center.y() - direction * half_width / 2)
        painter.drawPath(path)
    painter.restore()


class PagerWidget(QWidget):
    page_requested = Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.page = None
        self.color = QColor("white")
        self.setCursor(Qt.PointingHandCursor)
        self.setToolTip("单击上方 / 下方翻页；按住任意位置拖动；双击隐藏")
        self.hide()

    def sizeHint(self):
        return pager_size(self.page, self.font())

    def set_page(self, page, color, font):
        self.page, self.color = page, QColor(color)
        self.setFont(font)
        self.setFixedSize(self.sizeHint())
        self.setVisible(page.controls)
        self.update()

    def paintEvent(self, event):
        if self.page:
            painter = QPainter(self)
            paint_pager(painter, self.rect(), self.page, self.color, self.font())

    def activate_at(self, position):
        """由宿主的共享拖拽过滤器确认未拖动后调用。"""
        delta = pager_hit(self.rect(), position.x(), position.y())
        if delta:
            self.page_requested.emit(delta)
