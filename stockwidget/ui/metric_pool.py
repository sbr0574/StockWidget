"""设置页的指标显示池与拖动排序控件。"""

from PySide6.QtCore import (
    QByteArray, QMimeData, QPoint, QRect, QRectF, QSize, Qt, Signal,
)
from PySide6.QtGui import (
    QColor, QCursor, QDrag, QKeyEvent, QPainter, QPalette, QPixmap,
)
from PySide6.QtWidgets import (
    QLabel, QListWidget, QListWidgetItem, QVBoxLayout, QWidget,
)

from stockwidget.core.metric_layout import (
    METRIC_BY_ID,
    METRIC_SPECS,
    NAME_METRIC_ID,
    normalize_visible_metrics,
)
from stockwidget.ui.theme import accent_color, accent_rgba


_METRIC_MIME_TYPE = "application/x-stockwidget-metric"
_POOL_DISPLAYED = "displayed"
_POOL_AVAILABLE = "available"
# 点击 ⓘ 打开单位设置面板的指标
_UNIT_METRIC_IDS = ("volume", "amount")
# 带设置入口的指标：文本后追加 ⓘ
_CLICKABLE_METRIC_IDS = (NAME_METRIC_ID, "volume", "amount")


def _metric_display_text(metric_id: str) -> str:
    """指标块显示文本：有设置面板的指标追加 ⓘ 入口。"""
    label = METRIC_BY_ID[metric_id].label
    return f"{label} ⓘ" if metric_id in _CLICKABLE_METRIC_IDS else label


class MetricListWidget(QListWidget):
    """圆角矩形指标块列表；拖放结果交由父控件统一更新。"""

    drop_requested = Signal(str, str, int)
    metric_activated = Signal(str, str)
    # 点击“名称”后的 ⓘ 请求打开面板；参数为池自身，用于定位面板。
    name_settings_requested = Signal(object)
    # 点击“成交量/成交额”后的 ⓘ 请求打开面板；参数为池自身。
    unit_settings_requested = Signal(object)

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
        self.setWrapping(True)
        self.setResizeMode(QListWidget.ResizeMode.Adjust)
        # 指标块按文字宽度自适应，禁省略号，避免文字显示成“...”
        self.setTextElideMode(Qt.TextElideMode.ElideNone)
        self.setHorizontalScrollMode(QListWidget.ScrollMode.ScrollPerPixel)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setSpacing(2)
        self.setToolTip("拖动指标改变显示状态或顺序；双击可快速移入另一侧")
        self.itemClicked.connect(self._on_item_clicked)
        self.itemDoubleClicked.connect(self._activate_item)
        self._drag_chip_bg = QColor(0, 0, 0, 18)
        self._drag_radius = 5.0
        # 拖拽时目标插入位置指示（None 表示未在拖拽中）
        self._drop_indicator_index = None
        self._drop_color = accent_color()
        self._pressed_info_id = None
        self.set_theme(False)

    def set_metric_ids(self, metric_ids: list[str]):
        current_metric_id = self.current_metric_id()
        self.clear()
        for metric_id in metric_ids:
            text = _metric_display_text(metric_id)
            item = QListWidgetItem(text)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item.setData(Qt.ItemDataRole.UserRole, metric_id)
            item.setFlags(
                Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsSelectable
                | Qt.ItemFlag.ItemIsDragEnabled
            )
            if metric_id == NAME_METRIC_ID:
                item.setToolTip("单击 ⓘ 设置名称显示（字数/代码/类型）；双击文字移入另一池")
            elif metric_id in _UNIT_METRIC_IDS:
                item.setToolTip("单击 ⓘ 设置数值单位（中文/英文/自动）；双击文字移入另一池")
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
        if metric_id is None:
            return
        rect = self.visualRect(self.currentIndex())
        if not rect.isValid() or rect.isEmpty():
            return
        mime = QMimeData()
        mime.setData(_METRIC_MIME_TYPE, QByteArray(metric_id.encode("utf-8")))
        drag = QDrag(self)
        drag.setMimeData(mime)
        drag.setPixmap(
            self._drag_chip_pixmap(_metric_display_text(metric_id), rect.size())
        )
        cursor = self.viewport().mapFromGlobal(QCursor.pos())
        hotspot = QPoint(
            min(max(cursor.x() - rect.x(), 0), rect.width()),
            min(max(cursor.y() - rect.y(), 0), rect.height()),
        )
        drag.setHotSpot(hotspot)
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
            self._drop_indicator_index = self._drop_insert_index(
                event.position().toPoint()
            )
            self.viewport().update()
            return
        event.ignore()

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
                if metric_id == NAME_METRIC_ID:
                    self.name_settings_requested.emit(self)
                elif metric_id in _UNIT_METRIC_IDS:
                    self.unit_settings_requested.emit(self)
                else:
                    self.metric_activated.emit(self.pool_name, metric_id)
                event.accept()
                return
        super().keyPressEvent(event)

    def _on_item_clicked(self, _item: QListWidgetItem):
        self.clearSelection()
        self.setCurrentItem(None)
        self.clearFocus()

    def info_rect(self, item: QListWidgetItem) -> QRect:
        """居中文本末尾 ⓘ 的独立命中区域（viewport 坐标）。"""
        if item.data(Qt.ItemDataRole.UserRole) not in _CLICKABLE_METRIC_IDS:
            return QRect()
        rect = self.visualItemRect(item)
        metrics = self.fontMetrics()
        text_width = metrics.horizontalAdvance(item.text())
        info_width = metrics.horizontalAdvance("ⓘ")
        right = rect.center().x() + text_width // 2
        return QRect(right - info_width - 2, rect.top(), info_width + 4, rect.height())

    def _request_settings(self, metric_id: str):
        if metric_id == NAME_METRIC_ID:
            self.name_settings_requested.emit(self)
        elif metric_id in _UNIT_METRIC_IDS:
            self.unit_settings_requested.emit(self)

    def _activate_item(self, item: QListWidgetItem):
        metric_id = str(item.data(Qt.ItemDataRole.UserRole) or "")
        if metric_id not in METRIC_BY_ID:
            return
        self.metric_activated.emit(self.pool_name, metric_id)

    def mousePressEvent(self, event):
        self._pressed_info_id = None
        pos = event.position().toPoint()
        item = self.itemAt(pos)
        if (event.button() == Qt.MouseButton.LeftButton and item is not None
                and self.info_rect(item).contains(pos)):
            self._pressed_info_id = item.data(Qt.ItemDataRole.UserRole)
            event.accept()
            return
        # 点击池内空白处时取消选中与焦点，其余交给默认处理。
        if (event.button() == Qt.MouseButton.LeftButton
                and not self.indexAt(event.position().toPoint()).isValid()):
            self.clearSelection()
            self.setCurrentItem(None)
            self.clearFocus()
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        metric_id, self._pressed_info_id = self._pressed_info_id, None
        if metric_id is not None and event.button() == Qt.MouseButton.LeftButton:
            pos = event.position().toPoint()
            item = self.itemAt(pos)
            if (item is not None and item.data(Qt.ItemDataRole.UserRole) == metric_id
                    and self.info_rect(item).contains(pos)):
                self.setCurrentItem(item)
                self._request_settings(metric_id)
            event.accept()
            return
        super().mouseReleaseEvent(event)

    def mouseMoveEvent(self, event):
        if self._pressed_info_id is not None:
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseDoubleClickEvent(self, event):
        pos = event.position().toPoint()
        item = self.itemAt(pos)
        if (event.button() == Qt.MouseButton.LeftButton and item is not None
                and self.info_rect(item).contains(pos)):
            self.mousePressEvent(event)
            return
        super().mouseDoubleClickEvent(event)

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


class MetricPoolWidget(QWidget):
    """上“已显示指标”、下“可用指标”双池，输出有序的已显示指标列表。"""

    visible_metrics_changed = Signal(list)
    name_settings_requested = Signal(object)
    unit_settings_requested = Signal(object)

    # 每行指标块占用的估算高度（22 高圆角块 + 网格间距 2 + 边距 2 + 冗余 2）
    _ROW_PITCH = 28
    # 每个池至少保留两行高度
    _MIN_ROWS = 2

    def __init__(self, parent=None):
        super().__init__(parent)
        self._visible_metrics: list[str] = []

        self.available_label = QLabel("可用指标", self)
        self.displayed_label = QLabel("已显示指标", self)
        for label in (self.available_label, self.displayed_label):
            label.setFixedHeight(14)

        self.available_pool = MetricListWidget(
            _POOL_AVAILABLE, "已全部显示", self
        )
        self.displayed_pool = MetricListWidget(
            _POOL_DISPLAYED, "拖入要显示的指标", self
        )

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)
        layout.addWidget(self.displayed_label)
        layout.addWidget(self.displayed_pool)
        layout.addWidget(self.available_label)
        layout.addWidget(self.available_pool)
        # 两池按内容固定高度后，余量统一留在底部，避免布局把空隙
        # 分散到标题和池之间，造成增删指标时顶部位置上下跳动。
        layout.addStretch(1)

        for pool in (self.available_pool, self.displayed_pool):
            pool.drop_requested.connect(
                lambda source, metric_id, index, target=pool.pool_name:
                self.move_metric(source, target, metric_id, index)
            )
            pool.metric_activated.connect(self._toggle_metric)
            pool.name_settings_requested.connect(self.name_settings_requested)
            pool.unit_settings_requested.connect(self.unit_settings_requested)

        self._rebuild_pools()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._apply_pool_heights()

    def _estimated_rows(self, pool: MetricListWidget) -> int:
        """按当前池宽度估算指标块换行后的行数。"""
        count = pool.count()
        if count <= 0:
            return 0
        width = max(0, pool.viewport().width() - 2)
        if width <= 0:
            # 尚未布局时按每行 3 个粗略估算
            return max(1, (count + 2) // 3)
        step = pool.spacing() + 16  # 网格间距 + QSS padding(7*2) + margin(1*2)
        rows = 0
        used = 0
        for row in range(count):
            item_width = pool.item(row).sizeHint().width() + step
            if rows == 0 or used + item_width > width:
                rows += 1
                used = item_width
            else:
                used += item_width
        return max(1, rows)

    def _apply_pool_heights(self):
        """两个池按内容行数自动分配高度：较空的池让出空间，最少保留两行。"""
        total = max(0, self.height() - 2 * 14 - 3 * 2)  # 减去两个标题与间距
        if total <= 0:
            return
        rows_available = self._estimated_rows(self.available_pool)
        rows_displayed = self._estimated_rows(self.displayed_pool)
        pitch = self._ROW_PITCH
        min_px = self._MIN_ROWS * pitch
        ideal_available = max(rows_available, self._MIN_ROWS) * pitch
        ideal_displayed = max(rows_displayed, self._MIN_ROWS) * pitch

        if ideal_available + ideal_displayed <= total:
            self.available_pool.setFixedHeight(ideal_available)
            self.displayed_pool.setFixedHeight(ideal_displayed)
            return

        # 空间不足：先各保底两行，剩余高度按所需行数比例分配并封顶
        remainder = max(0, total - 2 * min_px)
        weight_available = max(0, rows_available - self._MIN_ROWS)
        weight_displayed = max(0, rows_displayed - self._MIN_ROWS)
        weight_total = weight_available + weight_displayed
        if weight_total > 0:
            extra_available = min(
                weight_available * pitch,
                remainder * weight_available // weight_total,
            )
            extra_displayed = min(
                weight_displayed * pitch,
                remainder - extra_available,
            )
        else:
            extra_available = extra_displayed = 0
        self.available_pool.setFixedHeight(min_px + extra_available)
        self.displayed_pool.setFixedHeight(min_px + extra_displayed)

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
        # 移动后一律不保留选中状态
        self._rebuild_pools()
        self._clear_pool_selections()
        self.visible_metrics_changed.emit(list(self._visible_metrics))

    def _clear_pool_selections(self):
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
        available = [
            spec.metric_id for spec in METRIC_SPECS
            if spec.metric_id not in visible
        ]
        self.displayed_pool.set_metric_ids(self._visible_metrics)
        self.available_pool.set_metric_ids(available)
        self._apply_pool_heights()
