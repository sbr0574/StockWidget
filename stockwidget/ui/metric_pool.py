"""设置页的指标显示池与拖动排序控件。"""

from PySide6.QtCore import QByteArray, QMimeData, QSize, Qt, Signal
from PySide6.QtGui import QDrag, QKeyEvent, QPainter, QPalette
from PySide6.QtWidgets import QLabel, QListWidget, QListWidgetItem, QVBoxLayout, QWidget

from stockwidget.core.metric_layout import (
    METRIC_BY_ID,
    METRIC_SPECS,
    normalize_visible_metrics,
)


_METRIC_MIME_TYPE = "application/x-stockwidget-metric"
_POOL_DISPLAYED = "displayed"
_POOL_AVAILABLE = "available"


class MetricListWidget(QListWidget):
    """横向指标块列表；拖放结果交由父控件统一更新。"""

    drop_requested = Signal(str, str, int)
    metric_activated = Signal(str, str)

    def __init__(self, pool_name: str, empty_text: str, parent=None):
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
        self.setWrapping(False)
        self.setResizeMode(QListWidget.ResizeMode.Adjust)
        self.setHorizontalScrollMode(QListWidget.ScrollMode.ScrollPerPixel)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setFixedHeight(37)
        self.setSpacing(2)
        self.setToolTip("拖动指标改变显示状态或顺序；双击可快速移入另一侧")
        self.itemDoubleClicked.connect(self._activate_item)

    def set_metric_ids(self, metric_ids: list[str]):
        current_metric_id = self.current_metric_id()
        self.clear()
        for metric_id in metric_ids:
            spec = METRIC_BY_ID[metric_id]
            item = QListWidgetItem(spec.label)
            item.setData(Qt.ItemDataRole.UserRole, metric_id)
            item.setFlags(
                Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsSelectable
                | Qt.ItemFlag.ItemIsDragEnabled
            )
            width = self.fontMetrics().horizontalAdvance(spec.label) + 22
            item.setSizeHint(QSize(max(48, width), 24))
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

    def startDrag(self, _supported_actions):
        metric_id = self.current_metric_id()
        if metric_id is None:
            return
        mime = QMimeData()
        mime.setData(_METRIC_MIME_TYPE, QByteArray(metric_id.encode("utf-8")))
        drag = QDrag(self)
        drag.setMimeData(mime)
        drag.exec(Qt.DropAction.MoveAction)

    def dragEnterEvent(self, event):
        if event.mimeData().hasFormat(_METRIC_MIME_TYPE):
            event.setDropAction(Qt.DropAction.MoveAction)
            event.accept()
            return
        event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasFormat(_METRIC_MIME_TYPE):
            event.setDropAction(Qt.DropAction.MoveAction)
            event.accept()
            return
        event.ignore()

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

        pos = event.position().toPoint()
        index = self.indexAt(pos)
        insert_at = self.count()
        if index.isValid():
            insert_at = index.row()
            if pos.x() > self.visualRect(index).center().x():
                insert_at += 1

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

    def _activate_item(self, item: QListWidgetItem):
        metric_id = str(item.data(Qt.ItemDataRole.UserRole) or "")
        if metric_id in METRIC_BY_ID:
            self.metric_activated.emit(self.pool_name, metric_id)

    def paintEvent(self, event):
        super().paintEvent(event)
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


class MetricPoolWidget(QWidget):
    """维护“已显示/可用指标”双池，并输出有序的已显示指标列表。"""

    visible_metrics_changed = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._visible_metrics: list[str] = []

        self.displayed_label = QLabel("已显示（拖动排序）", self)
        self.available_label = QLabel("可用指标", self)
        for label in (self.displayed_label, self.available_label):
            label.setFixedHeight(14)

        self.displayed_pool = MetricListWidget(
            _POOL_DISPLAYED, "拖入要显示的指标", self
        )
        self.available_pool = MetricListWidget(
            _POOL_AVAILABLE, "已全部显示", self
        )

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)
        layout.addWidget(self.displayed_label)
        layout.addWidget(self.displayed_pool)
        layout.addWidget(self.available_label)
        layout.addWidget(self.available_pool)

        for pool in (self.displayed_pool, self.available_pool):
            pool.drop_requested.connect(
                lambda source, metric_id, index, target=pool.pool_name:
                self.move_metric(source, target, metric_id, index)
            )
            pool.metric_activated.connect(self._toggle_metric)

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
        self._rebuild_pools(select_metric_id=metric_id, selected_pool=target_pool)
        self.visible_metrics_changed.emit(list(self._visible_metrics))

    def set_theme(self, dark: bool):
        border = "rgba(255, 255, 255, 0.26)" if dark else "rgba(0, 0, 0, 0.22)"
        chip = "rgba(255, 255, 255, 0.10)" if dark else "rgba(0, 0, 0, 0.06)"
        hover = "rgba(10, 132, 255, 0.22)" if dark else "rgba(0, 122, 255, 0.14)"
        selected = "rgba(10, 132, 255, 0.34)" if dark else "rgba(0, 122, 255, 0.24)"
        scroll_handle = "rgba(255, 255, 255, 0.34)" if dark else "rgba(0, 0, 0, 0.28)"
        style = f"""
            QListWidget {{
                background: transparent;
                border: 1px solid {border};
                border-radius: 5px;
                outline: none;
            }}
            QListWidget::item {{
                background: {chip};
                border: 1px solid {border};
                border-radius: 8px;
                padding: 2px 7px;
                margin: 1px;
            }}
            QListWidget::item:hover {{ background: {hover}; }}
            QListWidget::item:selected {{
                background: {selected};
                border-color: rgb(10, 132, 255);
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
        """
        self.displayed_pool.setStyleSheet(style)
        self.available_pool.setStyleSheet(style)

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

    def _rebuild_pools(
        self,
        *,
        select_metric_id: str | None = None,
        selected_pool: str | None = None,
    ):
        visible = set(self._visible_metrics)
        available = [
            spec.metric_id for spec in METRIC_SPECS
            if spec.metric_id not in visible
        ]
        self.displayed_pool.set_metric_ids(self._visible_metrics)
        self.available_pool.set_metric_ids(available)

        if select_metric_id is not None:
            pool = (
                self.displayed_pool
                if selected_pool == _POOL_DISPLAYED
                else self.available_pool
            )
            for row in range(pool.count()):
                item = pool.item(row)
                if item.data(Qt.ItemDataRole.UserRole) == select_metric_id:
                    pool.setCurrentItem(item)
                    pool.scrollToItem(item)
                    break
