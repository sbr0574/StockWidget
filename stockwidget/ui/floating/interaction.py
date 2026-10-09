"""浮窗交互：共用拖动手势、完整位置与贴边状态、定时及旧行情隐藏。"""

from datetime import datetime
import math
import time

from PySide6.QtCore import QPoint, Qt, QEvent, QObject, QRect, QTimer
from PySide6.QtGui import QCursor, QScreen
from PySide6.QtWidgets import QApplication, QWidget

from stockwidget.core.window_rules import best_screen, clamp_point, all_quotes_stale, due_hide_times


class DragBehaviorMixin:
    """begin/move/finish_drag 接受统一坐标，Qt 事件只负责转换输入。"""

    def _init_drag(self, clickable_header=None):
        self._clickable_header = clickable_header
        self._drag_region_clicks = {}
        self._drag_surface = "float"
        self._reset_drag()

    def register_drag_region(self, widget, click=None):
        """子控件复用拖动/双击；交互控件仅在原区域未拖动松开时响应单击。"""
        widget.installEventFilter(self)
        if click is not None:
            self._drag_region_clicks[widget] = click

    def _reset_drag(self):
        self._drag_pos = None
        self._drag_start_pos = None
        self._dragging = False
        self._system_moving = False
        self._pressed_header_section = -1
        self._pressed_region = None

    # ----- 拖拽实现 -----
    def begin_drag(self, global_pos, *, surface="float", offset=None):
        """Shared pointer entry point for Qt and native taskbar adapters."""
        self.position_controller.prepare_drag()
        self._reset_drag()
        self._drag_surface = surface
        self._drag_start_pos = QPoint(global_pos)
        self._drag_start_window_pos = self.pos()
        self._drag_pos = QPoint(offset) if offset is not None else self._drag_start_pos - self.frameGeometry().topLeft()
        self.setFocus(Qt.MouseFocusReason)

    def _is_drag_position(self, global_pos):
        return self._dragging or (
            self._drag_start_pos is not None
            and (global_pos - self._drag_start_pos).manhattanLength()
            >= QApplication.startDragDistance()
        )

    def move_drag(self, global_pos, *, left_pressed=True):
        """超过系统拖动阈值后移动浮窗，避免单击时轻微抖动触发拖动。"""
        if self._drag_pos is None or not left_pressed:
            return
        global_pos = QPoint(global_pos)
        if not self._is_drag_position(global_pos):
            return
        if not self._dragging:
            self._dragging = True
            self.drag_started.emit()
        if self._wayland_drag:
            if self._system_moving:
                return
            self._system_moving = True
            win = self.windowHandle()
            if win is not None and hasattr(win, "startSystemMove"):
                win.startSystemMove()
            return
        self.move(global_pos - self._drag_pos)
        self._ensure_on_top()
        self.drag_moved.emit()

    def finish_drag(self, accepted=True):
        dragged = self._dragging
        if dragged and not accepted:
            self.move(self._drag_start_window_pos)
        self._reset_drag()
        self._ensure_on_top()
        if dragged:
            self.drag_finished.emit(accepted)
        self.position_controller.finish_drag(accepted)
        if dragged:
            self._notify_change()

    def _drag_press(self, e):
        self.begin_drag(e.globalPosition().toPoint())

    def _is_drag(self, e):
        return self._is_drag_position(e.globalPosition().toPoint())

    def _drag_move(self, e):
        self.move_drag(e.globalPosition().toPoint(), left_pressed=bool(e.buttons() & Qt.LeftButton))

    def _drag_release(self):
        self.finish_drag()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape and self._dragging:
            self.finish_drag(False)
            event.accept()
            return
        QWidget.keyPressEvent(self, event)

    def _header_section_at(self, obj, ev):
        header = self._clickable_header
        if header is not None and obj is header.viewport() and header.sectionsClickable():
            pos = ev.position().toPoint()
            if obj.rect().contains(pos):
                return header.logicalIndexAt(pos)
        return -1

    # ----- 鼠标事件 -----
    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_press(e)

    def mouseMoveEvent(self, e):
        self._drag_move(e)

    def mouseReleaseEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_release()

    def mouseDoubleClickEvent(self, e):
        if e.button() == Qt.LeftButton:
            self.hide_widget()

    # ----- 子控件事件过滤（表格区域同样支持拖拽/双击隐藏）-----
    def eventFilter(self, obj, ev):
        if ev.type() == QEvent.MouseButtonDblClick and ev.button() == Qt.LeftButton:
            self.mouseDoubleClickEvent(ev)
            return True
        if ev.type() == QEvent.MouseButtonPress and ev.button() == Qt.LeftButton:
            self._drag_press(ev)
            self._pressed_region = obj
            self._pressed_header_section = self._header_section_at(obj, ev)
            return True
        if ev.type() == QEvent.MouseMove and (ev.buttons() & Qt.LeftButton) and self._drag_pos is not None:
            self._drag_move(ev)
            return True
        if ev.type() == QEvent.MouseButtonRelease and ev.button() == Qt.LeftButton:
            section = self._pressed_header_section
            stationary = self._drag_pos is not None and not self._is_drag(ev)
            clicked = (
                section >= 0 and stationary
                and section == self._header_section_at(obj, ev)
            )
            click = self._drag_region_clicks.get(obj) if (
                stationary and self._pressed_region is obj and obj.rect().contains(ev.position().toPoint())
            ) else None
            self._drag_release()
            if clicked:
                self._clickable_header.sectionClicked.emit(section)
            if click is not None:
                click(ev.position().toPoint())
            return True
        return QWidget.eventFilter(self, obj, ev)


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
            if QApplication.platformName() == "windows":
                # Windows keeps physical screen origins under Qt scaling.
                # A window crossing a logical screen edge can be attributed
                # to its neighbour; retain the screen containing most of it.
                handle = self.source.windowHandle()
                screens = QApplication.screens()
                bounds = best_screen(*rect.getRect(), [s.geometry().getRect() for s in screens])
                screen = next((s for s in screens if s.geometry().getRect() == bounds), None)
                if handle and isinstance(screen, QScreen) and handle.screen() != screen:
                    handle.setScreen(screen)
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
        if QApplication.platformName() == "windows":
            # Create on the initial screen before applying a saved position
            # that may extend into a gap between scaled logical screens.
            self.source.winId()
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


class HideController(QObject):
    def __init__(self, source):
        super().__init__(source)
        self.source = source
        self._last_check = datetime.now().astimezone()
        self._fired = set()
        self._manual_show = False
        self._deadline = None
        self.schedule_timer = QTimer(self)
        self.schedule_timer.setInterval(1000)
        self.schedule_timer.timeout.connect(self.check_schedule)
        self.countdown_timer = QTimer(self)
        self.countdown_timer.setTimerType(Qt.PreciseTimer)
        self.countdown_timer.setInterval(100)
        self.countdown_timer.timeout.connect(self.check_countdown)
        source.widget_visibility_changed.connect(self._visibility_changed)
        self.configure()

    def configure(self):
        if (self.source.hide_enabled and self.source.scheduled_hide_enabled
                and self.source.scheduled_hide_times):
            if not self.schedule_timer.isActive():
                self._last_check = datetime.now().astimezone()
                self.schedule_timer.start()
        else:
            self.schedule_timer.stop()
        if not self.source.hide_enabled or not self.source.auto_hide_enabled:
            self.cancel_countdown()

    def check_schedule(self, now=None):
        now = now or datetime.now().astimezone()
        due = due_hide_times(self.source.scheduled_hide_times, self._last_check, now)
        self._last_check = now
        self._fired = {event for event in self._fired if event[0] == now.date()}
        if not self.source.hide_enabled or not self.source.scheduled_hide_enabled:
            return
        pending = {(now.date(), value) for value in due} - self._fired
        self._fired.update(pending)  # 隐藏时也消耗当天事件，呼出后不会重复隐藏。
        if pending and self.source.widget_visible:
            self.source.hide_widget()

    def _visibility_changed(self):
        self.cancel_countdown()
        self._manual_show = self.source.widget_visible

    def quotes_refreshed(self, data):
        source = self.source
        if (not source.hide_enabled or not source.auto_hide_enabled or not source.widget_visible
                or not all_quotes_stale(data, source.checked_codes, time.time())):
            self.cancel_countdown()
            return
        if self._deadline is not None:
            return  # 连续收到旧行情不会重置倒计时。
        if self._manual_show:
            self._deadline = time.monotonic() + 5
            self.check_countdown()
            self.countdown_timer.start()
        else:
            source.hide_widget()

    def check_countdown(self):
        if self._deadline is None:
            return
        if (not self.source.hide_enabled or not self.source.auto_hide_enabled
                or not self.source.widget_visible):
            self.cancel_countdown()
            return
        remaining = self._deadline - time.monotonic()
        if remaining <= 0:
            self.source.hide_widget()
            return
        label = self.source.hide_notice
        text = f"行情已超过30秒未更新，{math.ceil(remaining)}秒后隐藏"
        if label.text() != text:
            label.setText(text)
            label.show()
            self.source._defer_fit()
            self.source.presentation_changed.emit()

    def cancel_countdown(self):
        self._deadline = None
        self.countdown_timer.stop()
        label = self.source.hide_notice
        if label.text():
            label.clear()
            label.hide()
            self.source._defer_fit()
            self.source.presentation_changed.emit()
