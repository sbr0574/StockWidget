"""指标设置弹窗共用的定位与重复点击处理。"""

from PySide6.QtCore import QPoint, Qt
from PySide6.QtWidgets import QFrame, QWidget


class MetricSettingsPanel(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent, Qt.WindowType.Popup)
        self._anchor = None
        self._metric_id = None

    def show_for(self, anchor: QWidget):
        metric_id = anchor.current_metric_id() if hasattr(anchor, "current_metric_id") else None
        if self.isVisible() and self._anchor is anchor and self._metric_id == metric_id:
            self.hide()
            return
        self._anchor = anchor
        self._metric_id = metric_id
        self.setAttribute(Qt.WidgetAttribute.WA_NoMouseReplay, False)
        self.adjustSize()
        position = anchor.mapToGlobal(QPoint(0, anchor.height() + 2))
        screen = anchor.screen()
        if screen is not None:
            available = screen.availableGeometry()
            max_x = max(available.left(), available.right() - self.width() + 1)
            x = min(max(position.x(), available.left()), max_x)
            above_y = anchor.mapToGlobal(QPoint(0, -self.height() - 2)).y()
            y = position.y() if position.y() + self.height() <= available.bottom() + 1 else above_y
            max_y = max(available.top(), available.bottom() - self.height() + 1)
            position = QPoint(x, min(max(y, available.top()), max_y))
        self.move(position)
        self.show()
        self.raise_()

    def mousePressEvent(self, event):
        # Qt.Popup 会把外部点击重放给下方控件。再次点击同一个 ⓘ 时，
        # 只关闭面板，避免重放把刚关闭的面板又打开；其他点击照常传递。
        if not self.rect().contains(event.position().toPoint()):
            anchor = self._anchor
            same_info = False
            if anchor is not None and hasattr(anchor, "info_rect"):
                pos = anchor.viewport().mapFromGlobal(event.globalPosition().toPoint())
                item = anchor.itemAt(pos)
                same_info = (
                    item is not None
                    and item.data(Qt.ItemDataRole.UserRole) == self._metric_id
                    and anchor.info_rect(item).contains(pos)
                )
            self.setAttribute(Qt.WidgetAttribute.WA_NoMouseReplay, same_info)
        super().mousePressEvent(event)
