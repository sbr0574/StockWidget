"""任务栏行情绘制及显示位置协调；复用行情展示和浮窗交互入口。"""

import logging

from PySide6.QtCore import QObject, QPoint, QRect, Qt, QTimer
from PySide6.QtGui import QColor, QCursor, QFont, QFontMetrics, QImage, QPainter
from PySide6.QtWidgets import QApplication, QStyleOptionViewItem

from stockwidget.core.quote_presentation import BidAskCell
from stockwidget.core.view_options import column_ranges, taskbar_content_height, taskbar_font_pixels
from stockwidget.platform.taskbar import NativeTaskbarWindow, cursor_over_taskbar, find_taskbar
from stockwidget.ui.controls.quote_view import bid_ask_width, paint_bid_ask, paint_pager, pager_hit, pager_size


def taskbar_message(source):
    if source.hide_notice.text():
        return source.hide_notice.text()
    if not source.get_surface_metrics("taskbar"):
        return "请选择任务栏显示指标"
    message = source.message_label.text()
    if source._message_kind != "metrics" and message:
        return message
    return "暂无行情" if not source.taskbar_model.rowCount() else ""


def taskbar_pager_rect(source, height, dpi):
    page = source.quotes.get_page("taskbar")
    if page.controls and not taskbar_message(source):
        appearance_font, *_ = source.get_taskbar_appearance()
        font = QFont(appearance_font)
        font.setPixelSize(taskbar_font_pixels(height, dpi, source.view_options.taskbar_rows, font.pointSizeF()))
        width = pager_size(page, font, round(38 * dpi / 96)).width()
        return QRect(0, 0, width, height)
    return QRect()


def render_taskbar(source, height, dpi=96, max_width=480, hit_regions=None):
    """Render independent metrics/style at physical taskbar DPI."""
    scale = dpi / 96
    if hit_regions is not None:
        hit_regions.clear()
    padding = max(2, round(5 * scale))
    options = source.view_options
    row_height = max(1, (height - 4) // options.taskbar_rows)
    appearance_font, color, opacity, _unicolor = source.get_taskbar_appearance()
    font = QFont(appearance_font)
    font.setPixelSize(taskbar_font_pixels(height, dpi, options.taskbar_rows, font.pointSizeF()))
    fm = QFontMetrics(font)
    model = source.taskbar_model
    rows, cols = model.rowCount(), model.columnCount()
    message = taskbar_message(source)
    if message:
        text = message
        width = min(max_width, max(round(120 * scale), fm.horizontalAdvance(text) + padding * 2))
        image = QImage(width, height, QImage.Format_ARGB32_Premultiplied)
        image.fill(QColor(0, 0, 0, 1))
        painter = QPainter(image)
        painter.setOpacity(opacity / 100)
        painter.setFont(font)
        painter.setPen(QColor("#ff6666") if source.message_label.text() and source._message_kind is None else color)
        painter.drawText(image.rect().adjusted(padding, 0, -padding, 0), Qt.AlignCenter,
                         fm.elidedText(text, Qt.ElideRight, width - padding * 2))
        painter.end()
        return image
    pager_rect = taskbar_pager_rect(source, height, dpi)
    split, separator = source.view_options.split_settings("taskbar")
    ranges = column_ranges(rows, split)
    gap = round(9 * scale) if split else 0
    content_width = max(1, max_width - pager_rect.width() - gap)
    widths = []
    for c in range(cols):
        if model.headerData(c, Qt.Horizontal) == "K线":
            widths.append(round(30 * scale))
        else:
            # Integer font metrics can round down; elidedText uses fractional
            # advances internally, so allow two pixels for the final glyph.
            sizes = []
            for r in range(rows):
                index = model.index(r, c)
                cell = index.data(Qt.UserRole)
                sizes.append(bid_ask_width(cell, font, padding) if isinstance(cell, BidAskCell)
                             else fm.horizontalAdvance(str(index.data())) + padding * 2 + 2)
            maximum = 360 if model.headerData(c, Qt.Horizontal) == "买一/卖一" else 180
            widths.append(min(round(maximum * scale), max(sizes)))
    total = sum(widths) * len(ranges)
    if total > content_width:
        widths = [max(1, int(width * content_width / total)) for width in widths]
    image = QImage(max(1, sum(widths) * len(ranges) + gap + pager_rect.width()), height, QImage.Format_ARGB32_Premultiplied)
    # Alpha 1 is visually transparent but retains clicks between text glyphs.
    image.fill(QColor(0, 0, 0, 1))
    painter = QPainter(image)
    painter.setOpacity(opacity / 100)
    painter.setFont(font)
    painter.setRenderHint(QPainter.TextAntialiasing)
    try:
        if not pager_rect.isEmpty():
            paint_pager(painter, pager_rect, source.quotes.get_page("taskbar"), color, font)
        x = pager_rect.width()
        source.taskbar_k_delegate.set_point_size(max(5, round(font.pixelSize() * 72 / dpi)))
        for block, (start, stop) in enumerate(ranges):
            if block:
                if separator:
                    painter.setPen(color)
                    painter.drawLine(x + gap // 2, 2, x + gap // 2, height - 3)
                x += gap
            for c, width in enumerate(widths):
                for r in range(start, stop):
                    rect = QRect(x, 2 + (r - start) * row_height, width, row_height)
                    if hit_regions is not None:
                        hit_regions.append((rect, r - start, block))
                    index = model.index(r, c)
                    cell = model.data(index, Qt.UserRole)
                    painter.save()
                    painter.setClipRect(rect)
                    if isinstance(cell, BidAskCell):
                        paint_bid_ask(painter, rect, font, cell, model, padding)
                    elif model.headerData(c, Qt.Horizontal) == "K线":
                        option = QStyleOptionViewItem()
                        option.rect, option.font = rect, font
                        source.taskbar_k_delegate.paint(painter, option, index)
                    else:
                        painter.setPen(model.data(index, Qt.ForegroundRole))
                        text = fm.elidedText(str(model.data(index, Qt.DisplayRole)), Qt.ElideRight,
                                             max(0, width - padding * 2))
                        painter.drawText(rect.adjusted(padding, 0, -padding, 0),
                                         int(model.data(index, Qt.TextAlignmentRole)), text)
                    painter.restore()
                x += width
    finally:
        painter.end()
    return image


class TaskbarController(QObject):
    def __init__(self, source, open_settings, parent=None):
        super().__init__(parent)
        self.source = source
        source.set_open_settings_callback(open_settings)
        self.native = None
        self._active = False
        self._closed = False
        self._dragging = False
        self._drag_over = False
        self._pointer_over = None
        self._native_press = None
        self._dpi = 96
        self._pager_rect = QRect()
        self._hit_regions = []
        self.timer = QTimer(self)
        self.timer.setInterval(1000)
        self.timer.timeout.connect(self.refresh)
        self.input_timer = QTimer(self)
        self.input_timer.setInterval(25)
        self.input_timer.timeout.connect(self._poll_pointer)
        source.taskbar_options_changed.connect(self.apply_mode)
        source.widget_visibility_changed.connect(self.apply_mode)
        source.presentation_changed.connect(self.schedule_refresh)
        source.drag_started.connect(self.drag_started)
        source.drag_moved.connect(self.drag_moved)
        source.drag_finished.connect(self.drag_finished)
        self._refresh_pending = False

    @property
    def enabled(self):
        return self.source.view_options.taskbar_enabled and self.source.widget_visible and (
            self._drag_over if self._dragging else self.source.display_mode != "float")

    def _status(self, text):
        if self.source.taskbar_status != text:
            self.source.taskbar_status = text
            self.source.taskbar_status_changed.emit(text)

    def apply_mode(self):
        if self._closed:
            return
        if self.enabled:
            if not self.timer.isActive():
                self.timer.start()
            self.refresh()
        else:
            self.timer.stop()
            if self.native:
                self.native.hide()
            self._active = False
            self._status("已隐藏" if not self.source.widget_visible else "任务栏显示已关闭")
        if self._dragging:
            return
        if not self.source.widget_visible or (self.source.display_mode == "taskbar" and self._active):
            self.source.hide()
        else:
            self.source.show()
        self.source.sync_refresh_timer()

    def schedule_refresh(self):
        if not self._closed and self.enabled and not self._refresh_pending:
            self._refresh_pending = True
            QTimer.singleShot(0, self, self.refresh)

    def refresh(self):
        self._refresh_pending = False
        if self._closed or not self.enabled:
            return
        was_active = self._active
        try:
            if QApplication.platformName() != "windows":
                raise OSError("任务栏显示需要 Windows 桌面会话")
            area = find_taskbar()
            if area is None:
                raise OSError("等待 Windows 任务栏恢复")
            height = taskbar_content_height(area.height, area.dpi)
            max_width = min(round(480 * area.dpi / 96), area.right)
            if max_width < 80 or height < 24:
                raise OSError("任务栏空间不足，请使用标准高度的横向任务栏")
            image = render_taskbar(self.source, height, area.dpi, max_width, self._hit_regions)
            self._pager_rect = taskbar_pager_rect(self.source, height, area.dpi)
            self._dpi = area.dpi
            if self.native is None:
                self.native = NativeTaskbarWindow(self._native_pointer, self._context_menu)
            self.native.present(area, image.width(), image.height(), bytes(image.constBits()),
                                round(self.source.taskbar_offset * area.dpi / 96))
            self._active = True
            page = self.source.quotes.get_page("taskbar")
            self._status(f"已显示在主屏任务栏 · 第 {page.index + 1}/{page.count} 页")
            if not was_active and self.source.display_mode == "taskbar" and not self._dragging:
                self.source.hide()
        except (OSError, ValueError) as error:
            if self.native:
                self.native.hide()
            self._active = False
            status = f"{error}；暂用浮窗，将自动重试"
            if self.source.taskbar_status != status:
                logging.getLogger(__name__).warning("%s", status)
            self._status(status)
            if self.source.display_mode == "taskbar" or self._dragging:
                self.source.show()

    def _click(self, x, y):
        delta = pager_hit(self._pager_rect, x, y)
        if delta:
            self.source.quotes.change_page("taskbar", delta)
            return
        for rect, row, block in self._hit_regions:
            if rect.contains(x, y):
                self.source.history.request_row("taskbar", row, block)
                break

    def _double_click(self, x, y):
        self.source.hide_widget()

    def _native_pointer(self, kind, x, y):
        # Snapshot before queuing: the pointer can move again before Qt runs.
        position = QCursor.pos()
        over = False if self.source.display_mode == "both" else cursor_over_taskbar()
        QTimer.singleShot(0, self, lambda: self._dispatch_pointer(kind, x, y, position, over))

    def _dispatch_pointer(self, kind, x, y, position, over):
        if self._closed:
            return
        self._pointer_over = over
        try:
            if kind == "press":
                self.input_timer.start()
                self._native_press = pager_hit(self._pager_rect, x, y)
                full = self.source.position_controller.full_geometry()
                offset = QPoint(max(0, min(round(x * 96 / self._dpi), full.width() - 1)),
                                max(0, min(round(y * 96 / self._dpi), full.height() - 1)))
                self.source.begin_drag(position, surface="taskbar", offset=offset)
            elif kind == "move" and self._native_press is not None:
                self.source.move_drag(position)
            elif kind in ("release", "cancel"):
                self.input_timer.stop()
                clicked = (kind == "release" and self._native_press
                           and self.source._drag_pos is not None and not self.source._is_drag_position(position)
                           and self._native_press == pager_hit(self._pager_rect, x, y))
                self.source.finish_drag(kind == "release")
                if clicked:
                    self._click(x, y)
                self._native_press = None
                if not self._dragging:
                    self.apply_mode()
            elif kind == "double_click":
                self.input_timer.stop()
                self._native_press = None
                self._double_click(x, y)
        finally:
            self._pointer_over = None

    def drag_started(self):
        # 双开时任务栏固定显示，浮窗可在整个桌面自由移动。
        if (not self.source.view_options.taskbar_enabled or self.source.display_mode == "both"
                or QApplication.platformName() != "windows"):
            return
        self._dragging = True
        self._drag_over = None  # force the first hover update, even outside the bar
        self._drag_origin_mode = self.source.display_mode
        self._drag_origin_position = self.source._drag_start_window_pos
        self.drag_moved()

    def _poll_pointer(self):
        if self.native:
            event = self.native.poll_pointer()
            if event:
                self._native_pointer(*event)

    def drag_moved(self):
        if not self._dragging:
            return
        over = cursor_over_taskbar() if self._pointer_over is None else self._pointer_over
        if over == self._drag_over:
            return
        self._drag_over = over
        self.source.taskbar_preview_active = over
        self.apply_mode()
        if over and self._active:
            self.source.hide()
            if self.source._drag_surface == "float" and self.native:
                # Hiding Qt releases its implicit mouse grab. The native input
                # adapter continues the same gesture until release or Escape.
                self._native_press = 0
                self.native.capture_pointer()
                self.input_timer.start()
        elif self.source.widget_visible:
            # Showing a preview must not steal the native pointer capture.
            previous = self.source.testAttribute(Qt.WA_ShowWithoutActivating)
            self.source.setAttribute(Qt.WA_ShowWithoutActivating, True)
            try:
                self.source.show()
                self.source.raise_()
            finally:
                self.source.setAttribute(Qt.WA_ShowWithoutActivating, previous)
        self.source.sync_refresh_timer()

    def drag_finished(self, accepted):
        self.input_timer.stop()
        self._native_press = None
        if self.native:
            self.native.release_pointer()
        if not self._dragging:
            return
        if accepted:
            self.drag_moved()
        docked = accepted and self._drag_over and self._active
        self._dragging = self._drag_over = False
        self.source.taskbar_preview_active = False
        if not accepted:
            mode = self._drag_origin_mode
            self.source.move(self._drag_origin_position)
        elif docked:
            mode = "both" if self.source.view_options.taskbar_dual_open else "taskbar"
            # Remember a usable position for a later drag out of the bar.
            self.source.move(self._drag_origin_position)
        else:
            mode = "float"
        self.source.set_display_mode(mode)
        self.apply_mode()

    def _context_menu(self):
        QTimer.singleShot(0, self, self._show_menu)

    def _show_menu(self):
        self.source.show_context_menu(QCursor.pos(), "taskbar")

    def close(self):
        self._closed = True
        self.timer.stop()
        self.input_timer.stop()
        self.source.taskbar_preview_active = False
        if self.native:
            self.native.close()
            self.native = None
