"""浮窗屏幕边界与贴边收起；完整位置独立于临时的 5px 窄条。"""

from PySide6.QtCore import QEvent, QObject, QPoint, QRect, QTimer
from PySide6.QtGui import QCursor
from PySide6.QtWidgets import QApplication

from stockwidget.core.geometry import best_screen, clamp_point


class PositionController(QObject):
    STRIP_SIZE = 5

    def __init__(self, source):
        super().__init__(source)
        self.source = source
        self.collapsed = False
        self.edge = None
        self._full = None
        self._applying = False
        self._restored = False
        self._drag_origin = None
        self._panel_origin = QPoint(source.panel.pos())
        self.hover_timer = QTimer(self)
        self.hover_timer.setInterval(100)
        self.hover_timer.timeout.connect(self.check_pointer)
        source.installEventFilter(self)
        app = QApplication.instance()
        app.screenAdded.connect(self._screen_added)
        app.screenRemoved.connect(self.recheck)
        for screen in app.screens():
            self._screen_added(screen)

    def _screen_added(self, screen):
        screen.geometryChanged.connect(self.recheck)
        self.recheck()

    def full_geometry(self):
        return QRect(self._full if self.collapsed else self.source.geometry())

    def _bounded(self, rect):
        if not self.source.boundary_check_enabled:
            return rect, None
        # 使用整个屏幕，允许双开浮窗经过任务栏；任务栏停靠仍由共用拖动处理。
        screens = [s.geometry() for s in QApplication.screens()]
        bounds = best_screen(rect.x(), rect.y(), rect.width(), rect.height(),
                             [(g.x(), g.y(), g.width(), g.height()) for g in screens])
        if bounds is None:
            return rect, None
        rect.moveTopLeft(QPoint(*clamp_point(rect.x(), rect.y(), bounds,
                                            rect.width(), rect.height())))
        return rect, QRect(*bounds)

    def _edge_at(self, rect, bounds):
        if bounds is None or not self.source.edge_hide_enabled:
            return None
        if rect.left() == bounds.left():
            return "left"
        if rect.right() == bounds.right():
            return "right"
        if rect.top() == bounds.top():
            return "top"
        return None

    def _apply(self, rect):
        self._applying = True
        try:
            self.source.setGeometry(rect)
            # 收起方向决定露出的内容边缘：左侧露右边，顶部露下边。
            # 只平移完整面板，保留 Designer 布局与圆角；展开恢复原始偏移。
            panel_pos = QPoint(self._panel_origin)
            if self.collapsed:
                if self.edge == "left":
                    panel_pos.setX(panel_pos.x() + rect.width() - self._full.width())
                elif self.edge == "top":
                    panel_pos.setY(panel_pos.y() + rect.height() - self._full.height())
            self.source.panel.move(panel_pos)
        finally:
            self._applying = False

    def _strip_geometry(self):
        strip = QRect(self._full)
        if self.edge == "top":
            strip.setHeight(min(self.STRIP_SIZE, strip.height()))
        else:
            strip.setWidth(min(self.STRIP_SIZE, strip.width()))
            if self.edge == "right":
                strip.moveRight(self._full.right())
        return strip

    def restore_position(self, point):
        self._restored = True
        rect = self.full_geometry()
        rect.moveTopLeft(point)
        rect, _ = self._bounded(rect)
        self._apply(rect)

    def fit(self, size):
        before = self.full_geometry()
        rect = QRect(before)
        rect.setSize(size)
        if self.edge == "right":
            rect.moveRight(before.right())
        rect, bounds = self._bounded(rect) if self._restored else (rect, None)
        if self.collapsed:
            self._full = rect
            self.edge = self._edge_at(rect, bounds)
            if self.edge:
                self._apply(self._strip_geometry())
            else:
                self.expand()
        else:
            self._apply(rect)
        self.recheck()
        if (before.topLeft() != self.full_geometry().topLeft()
                and self.source._drag_pos is None and hasattr(self.source, "timer")):
            self.source._notify_change()

    def recheck(self, *_):
        if not self._restored or self._applying or self.source._updating_topmost:
            return
        rect, bounds = self._bounded(self.full_geometry())
        edge = self._edge_at(rect, bounds)
        if self.collapsed:
            self._full = rect
            self.edge = edge
            if edge:
                self._apply(self._strip_geometry())
            else:
                self.expand()
        else:
            self._apply(rect)
            self.edge = edge
        active = bool(edge and self.source.isVisible() and self.source.widget_visible)
        if active:
            self.hover_timer.start()
            self.check_pointer()
        else:
            self.hover_timer.stop()

    def check_pointer(self):
        source = self.source
        if (not self.edge or not source.isVisible() or not source.widget_visible
                or source._drag_pos is not None or source.taskbar_preview_active
                or QApplication.activePopupWidget() is not None):
            return
        # 原生鼠标坐标在高 DPI 边界转换时可能向外舍入一个逻辑像素。
        # 使用整个窗口的矩形（含透明圆角），避免最外侧像素触发缩回。
        if source.geometry().adjusted(-1, -1, 1, 1).contains(QCursor.pos()):
            self.expand()
        else:
            self.collapse()

    def collapse(self):
        if self.collapsed or not self.edge:
            return
        self._full = self.source.geometry()
        self.collapsed = True
        self._apply(self._strip_geometry())
        self.source.hide_controller.cancel_countdown()
        self.source.sync_refresh_timer()

    def expand(self):
        if not self.collapsed:
            return
        rect = self._full
        self.collapsed = False
        self._full = None
        self._apply(rect)
        self.source.sync_refresh_timer()

    def prepare_drag(self):
        self._drag_origin = (self.full_geometry(), self.collapsed, self.edge)
        self.expand()

    def finish_drag(self, accepted):
        origin = self._drag_origin
        self._drag_origin = None
        if not accepted and origin is not None:
            rect, collapsed, edge = origin
            self._apply(rect)
            self.edge = edge
            if collapsed:
                self.collapse()
                return
        self.recheck()

    def eventFilter(self, obj, event):
        kind = event.type()
        if kind == QEvent.Enter:
            self.expand()
        elif kind in (QEvent.Leave, QEvent.Show):
            QTimer.singleShot(0, self, self.recheck)
        elif kind == QEvent.Hide:
            self.hover_timer.stop()
        elif kind == QEvent.Move and not self._applying and not self.source._updating_topmost:
            if self.collapsed:
                self._full.translate(self.source.pos() - self._strip_geometry().topLeft())
            self.recheck()
        return False
