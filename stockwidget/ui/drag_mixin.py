# -*- coding: utf-8 -*-
"""浮窗拖拽 / 双击隐藏交互的混入类（与 UI 构建解耦，供 FloatLabel 复用）。

依赖宿主 widget 提供以下属性/方法：
- ``_wayland_drag`` (bool)：是否为 Wayland 会话（决定用系统级拖动）。
- ``_ensure_on_top()``：强制置顶逻辑。
- ``_notify_change()``：位置变化后回写配置。
"""

from PySide6.QtCore import Qt, QEvent
from PySide6.QtWidgets import QWidget


class DragBehaviorMixin:
    """拖拽移动 + 双击隐藏。Wayland 下用 startSystemMove，其余平台手动 move。"""

    def _init_drag(self):
        self._drag_pos = None
        self._system_moving = False

    # ----- 拖拽实现 -----
    def _drag_press(self, e):
        """按下左键：记录拖动起点。
        Wayland 下先记全局坐标，待移动超过阈值后再交给合成器（startSystemMove），
        这样普通单击/双击（隐藏浮窗）不受影响。
        """
        if self._wayland_drag:
            self._drag_pos = e.globalPosition().toPoint()
            self._system_moving = False
        else:
            self._drag_pos = e.globalPosition().toPoint() - self.frameGeometry().topLeft()
        self.setFocus(Qt.MouseFocusReason)

    def _drag_move(self, e):
        """按住左键移动：Wayland 用系统级拖动，其余平台（X11/Windows）手动 move。"""
        if getattr(self, "_drag_pos", None) is None or not (e.buttons() & Qt.LeftButton):
            return
        if self._wayland_drag:
            if self._system_moving:
                return
            # 移动超过阈值才触发系统级拖动，避免把单击误判为拖动
            pos = e.globalPosition().toPoint()
            if (pos - self._drag_pos).manhattanLength() <= 4:
                return
            self._system_moving = True
            win = self.windowHandle()
            if win is not None and hasattr(win, "startSystemMove"):
                win.startSystemMove()
            return
        self.move(e.globalPosition().toPoint() - self._drag_pos)
        self._ensure_on_top()

    def _drag_release(self):
        self._drag_pos = None
        self._system_moving = False
        self._ensure_on_top()
        self._notify_change()

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
            self._drag_pos = None
            self._system_moving = False
            self.hide()

    # ----- 子控件事件过滤（表格区域同样支持拖拽/双击隐藏）-----
    def eventFilter(self, obj, ev):
        if ev.type() == QEvent.MouseButtonDblClick and hasattr(ev, "button") and ev.button() == Qt.LeftButton:
            self._drag_pos = None
            self._system_moving = False
            self.hide()
            return True
        if ev.type() == QEvent.MouseButtonPress and hasattr(ev, "button") and ev.button() == Qt.LeftButton:
            self._drag_press(ev)
            return True
        if ev.type() == QEvent.MouseMove and hasattr(ev, "buttons") and (ev.buttons() & Qt.LeftButton) and getattr(self, "_drag_pos", None):
            self._drag_move(ev)
            return True
        if ev.type() == QEvent.MouseButtonRelease and hasattr(ev, "button") and ev.button() == Qt.LeftButton:
            self._drag_release()
            return True
        return QWidget.eventFilter(self, obj, ev)
