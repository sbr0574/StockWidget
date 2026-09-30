# -*- coding: utf-8 -*-
"""共享拖拽 / 双击隐藏交互；Qt 事件和任务栏原生事件复用同一入口。

依赖宿主 widget 提供以下属性/方法：
- ``_wayland_drag`` (bool)：是否为 Wayland 会话（决定用系统级拖动）。
- ``_ensure_on_top()``：强制置顶逻辑。
- ``_notify_change()``：位置变化后回写配置。
- ``hide_widget()``：隐藏整个行情组件。
- ``drag_started / drag_moved / drag_finished``：显示位置控制信号。
"""

from PySide6.QtCore import QPoint, Qt, QEvent
from PySide6.QtWidgets import QApplication, QWidget


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
