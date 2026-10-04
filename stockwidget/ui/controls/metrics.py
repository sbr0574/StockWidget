"""固定指标目录与已显示指标池；负责选择、拖动排序和编辑交互。"""

from PySide6.QtCore import QByteArray, QMimeData, QPoint, QRectF, QSize, Qt, Signal
from PySide6.QtGui import QColor, QCursor, QDrag, QKeyEvent, QPainter, QPalette, QPixmap
from PySide6.QtWidgets import QListWidget, QListWidgetItem, QWidget

from stockwidget.core.quote_presentation import (
    METRIC_BY_ID,
    METRIC_SPECS,
    normalize_visible_metrics,
)
from stockwidget.ui.controls.style import accent_color, accent_rgba


_METRIC_MIME_TYPE = "application/x-stockwidget-metric"
_POOL_DISPLAYED = "displayed"
_POOL_AVAILABLE = "available"


class MetricListWidget(QListWidget):
    """圆角矩形指标块列表；拖放结果交由父控件统一更新。"""

    drop_requested = Signal(str, str, int)
    metric_activated = Signal(str, str)
    fixed_catalog = False

    def __init__(self, pool_name="", empty_text="", parent=None):
        if isinstance(pool_name, QWidget):
            parent, pool_name = pool_name, ""
        super().__init__(parent)
        self.pool_name = pool_name
        self.empty_text = empty_text
        self.setObjectName(f"metric_{pool_name}_pool")
        self.setAcceptDrops(True)
        self.setDragEnabled(True)
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QListWidget.DragDropMode.DragDrop)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        self.setFlow(QListWidget.Flow.LeftToRight)
        self.setWrapping(True)
        self.setResizeMode(QListWidget.ResizeMode.Adjust)
        # 指标块按文字宽度自适应，禁省略号，避免文字显示成“...”
        self.setTextElideMode(Qt.TextElideMode.ElideNone)
        self.setHorizontalScrollMode(QListWidget.ScrollMode.ScrollPerPixel)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setSpacing(2)
        self.setToolTip("拖动或双击可用指标添加；拖动已显示指标排序，双击或拖回目录移除")
        self.itemClicked.connect(self._on_item_clicked)
        self.itemDoubleClicked.connect(self._activate_item)
        self._drag_chip_bg = QColor(0, 0, 0, 18)
        self._drag_radius = 5.0
        # 拖拽时目标插入位置指示（None 表示未在拖拽中）
        self._drop_indicator_index = None
        self._drop_color = accent_color()
        self.set_theme(False)

    def set_metric_ids(self, metric_ids: list[str]):
        current_metric_id = self.current_metric_id()
        self.clear()
        for metric_id in metric_ids:
            text = METRIC_BY_ID[metric_id].label
            item = QListWidgetItem(text)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item.setData(Qt.ItemDataRole.UserRole, metric_id)
            item.setFlags(
                Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsSelectable
                | Qt.ItemFlag.ItemIsDragEnabled
            )
            width = self.fontMetrics().horizontalAdvance(text) + 14
            item.setSizeHint(QSize(max(38, width), 22))
            self.addItem(item)
            if metric_id == current_metric_id:
                self.setCurrentItem(item)
        self.viewport().update()

    def current_metric_id(self) -> str | None:
        item = self.currentItem()
        if item is None:
            return None
        metric_id = str(item.data(Qt.ItemDataRole.UserRole) or "")
        return metric_id if metric_id in METRIC_BY_ID else None

    def set_theme(self, dark: bool):
        if dark:
            pool_bg = "rgba(255, 255, 255, 0.06)"
            chip = "rgba(255, 255, 255, 0.12)"
            hover = accent_rgba(0.12)
            scroll_handle = "rgba(255, 255, 255, 0.34)"
            self._drag_chip_bg = QColor(255, 255, 255, 31)
        else:
            pool_bg = "rgba(0, 0, 0, 0.05)"
            chip = "rgba(0, 0, 0, 0.07)"
            hover = accent_rgba(0.08)
            scroll_handle = "rgba(0, 0, 0, 0.28)"
            self._drag_chip_bg = QColor(0, 0, 0, 18)
        self._drop_color = accent_color()
        disabled_chip = "rgba(255, 255, 255, 0.04)" if dark else "rgba(0, 0, 0, 0.03)"
        disabled_text = "#A8ABB1" if dark else "#888888"
        if self.fixed_catalog:
            pool_bg = "transparent"
        self.setStyleSheet(f"""
            QListWidget {{
                background-color: {pool_bg};
                border: none;
                border-radius: 8px;
                outline: none;
            }}
            QListWidget::item {{
                background: {chip};
                border: none;
                border-radius: 5px;
                padding: 1px 1px;
                margin: 1px;
            }}
            QListWidget::item:hover {{ background: {hover}; }}
            QListWidget::item:selected {{
                background: {hover};
                color: palette(text);
            }}
            QListWidget::item:disabled {{
                background: {disabled_chip};
                color: {disabled_text};
            }}
            QScrollBar:horizontal {{
                height: 5px;
                background: transparent;
                margin: 0;
            }}
            QScrollBar::handle:horizontal {{
                min-width: 24px;
                background: {scroll_handle};
                border-radius: 2px;
            }}
            QScrollBar::add-line:horizontal,
            QScrollBar::sub-line:horizontal {{ width: 0; }}
            QScrollBar::add-page:horizontal,
            QScrollBar::sub-page:horizontal {{ background: transparent; }}
            QScrollBar:vertical {{
                width: 5px;
                background: transparent;
                margin: 0;
            }}
            QScrollBar::handle:vertical {{
                min-height: 24px;
                background: {scroll_handle};
                border-radius: 2px;
            }}
            QScrollBar::add-line:vertical,
            QScrollBar::sub-line:vertical {{ height: 0; }}
            QScrollBar::add-page:vertical,
            QScrollBar::sub-page:vertical {{ background: transparent; }}
        """)

    def _drag_chip_pixmap(self, text: str, size: QSize) -> QPixmap:
        """渲染跟随鼠标的圆角矩形指标块。"""
        ratio = max(1.0, float(self.devicePixelRatioF()))
        pixmap = QPixmap(
            max(1, round(size.width() * ratio)),
            max(1, round(size.height() * ratio)),
        )
        pixmap.setDevicePixelRatio(ratio)
        pixmap.fill(Qt.GlobalColor.transparent)

        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(self._drag_chip_bg)
        painter.drawRoundedRect(
            QRectF(0, 0, size.width(), size.height()),
            self._drag_radius,
            self._drag_radius,
        )
        painter.setFont(self.font())
        painter.setPen(self.palette().color(QPalette.ColorRole.Text))
        painter.drawText(
            QRectF(0, 0, size.width(), size.height()),
            Qt.AlignmentFlag.AlignCenter,
            text,
        )
        painter.end()
        return pixmap

    def startDrag(self, _supported_actions):
        metric_id = self.current_metric_id()
        if metric_id is None or not self.currentItem().flags() & Qt.ItemFlag.ItemIsDragEnabled:
            return
        rect = self.visualRect(self.currentIndex())
        if not rect.isValid() or rect.isEmpty():
            return
        mime = QMimeData()
        mime.setData(_METRIC_MIME_TYPE, QByteArray(metric_id.encode("utf-8")))
        drag = QDrag(self)
        drag.setMimeData(mime)
        drag.setPixmap(
            self._drag_chip_pixmap(METRIC_BY_ID[metric_id].label, rect.size())
        )
        cursor = self.viewport().mapFromGlobal(QCursor.pos())
        hotspot = QPoint(
            min(max(cursor.x() - rect.x(), 0), rect.width()),
            min(max(cursor.y() - rect.y(), 0), rect.height()),
        )
        drag.setHotSpot(hotspot)
        drag.exec(Qt.DropAction.MoveAction)

    def dragEnterEvent(self, event):
        if self._accepts_drag(event):
            event.setDropAction(Qt.DropAction.MoveAction)
            event.accept()
            return
        event.ignore()

    def dragMoveEvent(self, event):
        if self._accepts_drag(event):
            event.setDropAction(Qt.DropAction.MoveAction)
            event.accept()
            self._drop_indicator_index = None if self.fixed_catalog else self._drop_insert_index(event.position().toPoint())
            self.viewport().update()
            return
        event.ignore()

    def _accepts_drag(self, event):
        source = event.source()
        return event.mimeData().hasFormat(_METRIC_MIME_TYPE) and (not self.fixed_catalog or
            isinstance(source, MetricListWidget) and source.pool_name == _POOL_DISPLAYED)

    def dragLeaveEvent(self, event):
        self._drop_indicator_index = None
        self.viewport().update()
        super().dragLeaveEvent(event)

    def _drop_insert_index(self, pos: QPoint) -> int:
        """按视觉顺序计算投放位置：拖放点之前的指标块数量。

        换行布局下按行带（中心 ± 半高）比较，先比行、再比列，
        行尾空白处也能得到正确插入位置，避免被追加到队尾。
        """
        insert_at = 0
        for row in range(self.count()):
            rect = self.visualRect(self.model().index(row, 0))
            center = rect.center()
            half = max(1, rect.height()) / 2
            if center.y() + half < pos.y():
                # 块所在行整体在拖放点上方
                insert_at += 1
            elif abs(center.y() - pos.y()) <= half and center.x() < pos.x():
                # 同一行内且在拖放点左侧
                insert_at += 1
        return insert_at

    def dropEvent(self, event):
        source = event.source()
        if not isinstance(source, MetricListWidget):
            event.ignore()
            return

        raw_metric_id = bytes(event.mimeData().data(_METRIC_MIME_TYPE))
        metric_id = raw_metric_id.decode("utf-8", errors="ignore")
        if metric_id not in METRIC_BY_ID:
            event.ignore()
            return

        insert_at = self._drop_insert_index(event.position().toPoint())
        self._drop_indicator_index = None
        self.viewport().update()

        self.drop_requested.emit(source.pool_name, metric_id, insert_at)
        event.setDropAction(Qt.DropAction.MoveAction)
        event.accept()

    def keyPressEvent(self, event: QKeyEvent):
        if event.key() in (
            Qt.Key.Key_Return,
            Qt.Key.Key_Enter,
            Qt.Key.Key_Space,
            Qt.Key.Key_Delete,
            Qt.Key.Key_Backspace,
        ):
            metric_id = self.current_metric_id()
            if metric_id is not None:
                self.metric_activated.emit(self.pool_name, metric_id)
                event.accept()
                return
        super().keyPressEvent(event)

    def _on_item_clicked(self, _item: QListWidgetItem):
        self.clearSelection()
        self.setCurrentItem(None)
        self.clearFocus()

    def _activate_item(self, item: QListWidgetItem):
        metric_id = str(item.data(Qt.ItemDataRole.UserRole) or "")
        if metric_id not in METRIC_BY_ID or not item.flags() & Qt.ItemFlag.ItemIsEnabled:
            return
        self.metric_activated.emit(self.pool_name, metric_id)

    def mousePressEvent(self, event):
        # 点击池内空白处时取消选中与焦点，其余交给默认处理。
        if (event.button() == Qt.MouseButton.LeftButton
                and not self.indexAt(event.position().toPoint()).isValid()):
            self.clearSelection()
            self.setCurrentItem(None)
            self.clearFocus()
        super().mousePressEvent(event)

    def paintEvent(self, event):
        super().paintEvent(event)
        self._paint_drop_indicator()
        if self.count() or not self.empty_text:
            return
        painter = QPainter(self.viewport())
        color = self.palette().color(QPalette.ColorRole.PlaceholderText)
        painter.setPen(color)
        painter.drawText(
            self.viewport().rect(),
            Qt.AlignmentFlag.AlignCenter,
            self.empty_text,
        )

    def _paint_drop_indicator(self):
        """拖拽过程中在目标插入位置绘制竖向圆角指示条。"""
        insert_at = self._drop_indicator_index
        if insert_at is None or self.count() == 0:
            return
        if insert_at < self.count():
            rect = self.visualRect(self.model().index(insert_at, 0))
            x = rect.left() - 3
        else:
            rect = self.visualRect(self.model().index(self.count() - 1, 0))
            x = rect.right() + 2
        painter = QPainter(self.viewport())
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(self._drop_color)
        painter.drawRoundedRect(QRectF(x, rect.top(), 3, rect.height()), 1.5, 1.5)
        painter.end()


class MetricCatalogWidget(MetricListWidget):
    """A fixed, frameless catalog whose selected entries stay in place."""

    fixed_catalog = True

    def __init__(self, parent=None):
        super().__init__(_POOL_AVAILABLE, parent=parent)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)


class MetricPoolWidget(QWidget):
    """固定目录在上，可排序的已显示指标池在下，输出显示顺序。"""

    visible_metrics_changed = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._visible_metrics: list[str] = []

        # Import after MetricListWidget is defined: the generated form promotes
        # that class while MetricPoolWidget itself is promoted in settings.ui.
        from stockwidget.ui.generated.ui_metric_pool import Ui_MetricPool
        self.ui = Ui_MetricPool()
        self.ui.setupUi(self)
        self.available_label = self.ui.available_label
        self.displayed_label = self.ui.displayed_label
        self.available_pool = self.ui.metric_available_pool
        self.displayed_pool = self.ui.metric_displayed_pool
        for pool, name in ((self.available_pool, _POOL_AVAILABLE),
                           (self.displayed_pool, _POOL_DISPLAYED)):
            pool.pool_name = name
            pool.empty_text = pool.property("emptyText")

        for pool in (self.available_pool, self.displayed_pool):
            pool.drop_requested.connect(
                lambda source, metric_id, index, target=pool.pool_name:
                self.move_metric(source, target, metric_id, index)
            )
            pool.metric_activated.connect(self._toggle_metric)

        self.available_pool.set_metric_ids([spec.metric_id for spec in METRIC_SPECS])
        self._rebuild_pools()

    @property
    def visible_metrics(self) -> list[str]:
        return list(self._visible_metrics)

    def set_visible_metrics(self, metric_ids):
        normalized = normalize_visible_metrics(metric_ids)
        if normalized == self._visible_metrics:
            return
        self._visible_metrics = normalized
        self._rebuild_pools()

    def move_metric(
        self,
        source_pool: str,
        target_pool: str,
        metric_id: str,
        insert_at: int,
    ):
        """应用一次成功投放；未显示池自身不参与排序。"""
        if metric_id not in METRIC_BY_ID:
            return
        if source_pool == target_pool == _POOL_AVAILABLE:
            return
        if source_pool == _POOL_AVAILABLE and metric_id in self._visible_metrics:
            return

        updated = list(self._visible_metrics)
        source_index = updated.index(metric_id) if metric_id in updated else None
        if source_index is not None:
            updated.pop(source_index)

        if target_pool == _POOL_DISPLAYED:
            if source_pool == _POOL_DISPLAYED and source_index is not None:
                if source_index < insert_at:
                    insert_at -= 1
            insert_at = max(0, min(int(insert_at), len(updated)))
            updated.insert(insert_at, metric_id)
        elif target_pool != _POOL_AVAILABLE:
            return

        if updated == self._visible_metrics:
            return
        self._visible_metrics = updated
        # 移动后一律不保留选中状态
        self._rebuild_pools()
        self.clear_selections()
        self.visible_metrics_changed.emit(list(self._visible_metrics))

    def clear_selections(self):
        for pool in (self.available_pool, self.displayed_pool):
            pool.clearSelection()
            pool.setCurrentItem(None)
            pool.clearFocus()

    def set_theme(self, dark: bool):
        for pool in (self.available_pool, self.displayed_pool):
            pool.set_theme(dark)

    def _toggle_metric(self, source_pool: str, metric_id: str):
        if source_pool == _POOL_DISPLAYED:
            self.move_metric(
                _POOL_DISPLAYED, _POOL_AVAILABLE, metric_id, 0
            )
        else:
            self.move_metric(
                _POOL_AVAILABLE,
                _POOL_DISPLAYED,
                metric_id,
                len(self._visible_metrics),
            )

    def _rebuild_pools(self):
        visible = set(self._visible_metrics)
        self.displayed_pool.set_metric_ids(self._visible_metrics)
        for row in range(self.available_pool.count()):
            item = self.available_pool.item(row)
            selected = item.data(Qt.ItemDataRole.UserRole) in visible
            item.setFlags(Qt.ItemFlag.NoItemFlags if selected else
                          Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsDragEnabled)
            item.setToolTip("已显示；可在下方双击移除" if selected else "拖动或双击添加到已显示指标")
