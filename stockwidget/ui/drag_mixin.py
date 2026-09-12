# -*- coding: utf-8 -*-
"""浮窗拖拽 / 双击隐藏交互的混入类（与 UI 构建解耦，供 FloatLabel 复用）。

依赖宿主 widget 提供以下属性/方法：
- ``_wayland_drag`` (bool)：是否为 Wayland 会话（决定用系统级拖动）。
- ``_ensure_on_top()``：强制置顶逻辑。
- ``_notify_change()``：位置变化后回写配置。
"""

from PySide6.QtCore import Qt, QEvent
from PySide6.QtWidgets import QApplication, QWidget


class DragBehaviorMixin:
    """拖拽移动 + 双击隐藏。Wayland 下用 startSystemMove，其余平台手动 move。"""

    def _init_drag(self, clickable_header=None):
        self._clickable_header = clickable_header
        self._reset_drag()

    def _reset_drag(self):
        self._drag_pos = None
        self._drag_start_pos = None
        self._dragging = False
        self._system_moving = False
        self._pressed_header_section = -1

    # ----- 拖拽实现 -----
    def _drag_press(self, e):
        """按下左键：记录拖动起点。
        Wayland 下先记全局坐标，待移动超过阈值后再交给合成器（startSystemMove），
        这样普通单击/双击（隐藏浮窗）不受影响。
        """
        self._reset_drag()
        self._drag_start_pos = e.globalPosition().toPoint()
        self._drag_pos = self._drag_start_pos - self.frameGeometry().topLeft()
        self.setFocus(Qt.MouseFocusReason)

    def _is_drag(self, e):
        return self._dragging or (
            self._drag_start_pos is not None
            and (e.globalPosition().toPoint() - self._drag_start_pos).manhattanLength()
            >= QApplication.startDragDistance()
        )

    def _drag_move(self, e):
        """超过系统拖动阈值后移动浮窗，避免单击时轻微抖动触发拖动。"""
        if self._drag_pos is None or not (e.buttons() & Qt.LeftButton):
            return
        if not self._is_drag(e):
            return
        self._dragging = True
        if self._wayland_drag:
            if self._system_moving:
                return
            self._system_moving = True
            win = self.windowHandle()
            if win is not None and hasattr(win, "startSystemMove"):
                win.startSystemMove()
            return
        self.move(e.globalPosition().toPoint() - self._drag_pos)
        self._ensure_on_top()

    def _drag_release(self):
        dragged = self._dragging
        self._reset_drag()
        self._ensure_on_top()
        if dragged:
            self._notify_change()

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
            self._reset_drag()
            self.hide()

    # ----- 子控件事件过滤（表格区域同样支持拖拽/双击隐藏）-----
    def eventFilter(self, obj, ev):
        if ev.type() == QEvent.MouseButtonDblClick and ev.button() == Qt.LeftButton:
            self.mouseDoubleClickEvent(ev)
            return True
        if ev.type() == QEvent.MouseButtonPress and ev.button() == Qt.LeftButton:
            self._drag_press(ev)
            self._pressed_header_section = self._header_section_at(obj, ev)
            return True
        if ev.type() == QEvent.MouseMove and (ev.buttons() & Qt.LeftButton) and self._drag_pos is not None:
            self._drag_move(ev)
            return True
        if ev.type() == QEvent.MouseButtonRelease and ev.button() == Qt.LeftButton:
            section = self._pressed_header_section
            clicked = (
                section >= 0 and not self._is_drag(ev)
                and section == self._header_section_at(obj, ev)
            )
            self._drag_release()
            if clicked:
                self._clickable_header.sectionClicked.emit(section)
            return True
        return QWidget.eventFilter(self, obj, ev)
