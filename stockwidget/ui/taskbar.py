"""Two-row taskbar presentation sharing the floating window's projected model."""

import logging

from PySide6.QtCore import QObject, QRect, Qt, QTimer
from PySide6.QtGui import QColor, QCursor, QFont, QFontMetrics, QImage, QPainter
from PySide6.QtWidgets import QApplication, QMenu, QStyleOptionViewItem

from stockwidget.platform.taskbar import NativeTaskbarWindow, find_taskbar


DISPLAY_MODES = (("float", "仅浮窗"), ("taskbar", "仅任务栏"), ("both", "浮窗和任务栏"))


def render_taskbar(source, height, dpi=96, max_width=480):
    """Render at physical taskbar DPI; preserve column order, colors and K lines."""
    scale = dpi / 96
    padding = max(2, round(5 * scale))
    row_height = max(1, (height - 4) // 2)
    font = QFont(source.font)
    font.setPixelSize(max(7, min(round(source.font.pointSizeF() * dpi / 72), row_height - 3)))
    fm = QFontMetrics(font)
    model = source.model
    rows, cols = min(2, model.rowCount()), model.columnCount()
    message = source.message_label.text()
    if not rows or not cols or (message and source._message_kind is None):
        text = message or "暂无行情"
        width = min(max_width, max(round(120 * scale), fm.horizontalAdvance(text) + padding * 2))
        image = QImage(width, height, QImage.Format_ARGB32_Premultiplied)
        image.fill(QColor(0, 0, 0, 1))
        painter = QPainter(image)
        painter.setFont(font)
        painter.setPen(QColor("#ff6666") if message and source._message_kind is None else source.fg)
        painter.drawText(image.rect().adjusted(padding, 0, -padding, 0), Qt.AlignCenter,
                         fm.elidedText(text, Qt.ElideRight, width - padding * 2))
        painter.end()
        return image
    widths = []
    for c in range(cols):
        if model.headerData(c, Qt.Horizontal) == "K线":
            widths.append(round(30 * scale))
        else:
            texts = [str(model.data(model.index(r, c), Qt.DisplayRole)) for r in range(rows)]
            # Integer font metrics can round down; elidedText uses fractional
            # advances internally, so allow two pixels for the final glyph.
            widths.append(min(round(180 * scale), max(fm.horizontalAdvance(t) for t in texts) + padding * 2 + 2))
    total = sum(widths)
    if total > max_width:
        widths = [max(1, int(width * max_width / total)) for width in widths]
    image = QImage(max(1, sum(widths)), height, QImage.Format_ARGB32_Premultiplied)
    # Alpha 1 is visually transparent but retains clicks between text glyphs.
    image.fill(QColor(0, 0, 0, 1))
    painter = QPainter(image)
    painter.setFont(font)
    painter.setRenderHint(QPainter.TextAntialiasing)
    try:
        x = 0
        for c, width in enumerate(widths):
            for r in range(rows):
                rect = QRect(x, 2 + r * row_height, width, row_height)
                index = model.index(r, c)
                painter.save()
                painter.setClipRect(rect)
                if model.headerData(c, Qt.Horizontal) == "K线":
                    option = QStyleOptionViewItem()
                    option.rect, option.font = rect, font
                    source.k_delegate.paint(painter, option, index)
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
        self.open_settings = open_settings
        self.native = None
        self._active = False
        self._closed = False
        self.timer = QTimer(self)
        self.timer.setInterval(1000)
        self.timer.timeout.connect(self.refresh)
        source.taskbar_options_changed.connect(self.apply_mode)
        source.presentation_changed.connect(self.schedule_refresh)
        self._refresh_pending = False

    def _status(self, text):
        if self.source.taskbar_status != text:
            self.source.taskbar_status = text
            self.source.taskbar_status_changed.emit(text)

    def apply_mode(self):
        if self._closed:
            return
        enabled = self.source.display_mode != "float"
        if enabled:
            self.timer.start()
            self.refresh()
        else:
            self.timer.stop()
            if self.native:
                self.native.hide()
            self._active = False
            self._status("任务栏显示已关闭")
        if self.source.display_mode == "taskbar" and self._active:
            self.source.hide()
        else:
            self.source.show()
        self.source.sync_refresh_timer()

    def schedule_refresh(self):
        if not self._closed and self.source.display_mode != "float" and not self._refresh_pending:
            self._refresh_pending = True
            QTimer.singleShot(0, self, self.refresh)

    def refresh(self):
        self._refresh_pending = False
        if self._closed or self.source.display_mode == "float":
            return
        was_active = self._active
        try:
            if QApplication.platformName() != "windows":
                raise OSError("任务栏显示需要 Windows 桌面会话")
            area = find_taskbar()
            if area is None:
                raise OSError("等待 Windows 任务栏恢复")
            height = min(area.height, round(44 * area.dpi / 96))
            max_width = min(round(480 * area.dpi / 96), area.right)
            if max_width < 80 or height < 24:
                raise OSError("任务栏空间不足，请使用标准高度的横向任务栏")
            image = render_taskbar(self.source, height, area.dpi, max_width)
            if self.native is None:
                self.native = NativeTaskbarWindow(self._click, self._context_menu)
            self.native.present(area, image.width(), image.height(), bytes(image.constBits()),
                                round(self.source.taskbar_offset * area.dpi / 96))
            self._active = True
            self._status("已显示在主屏任务栏 · 当前排序前两行")
            if not was_active and self.source.display_mode == "taskbar":
                self.source.hide()
        except (OSError, ValueError) as error:
            if self.native:
                self.native.hide()
            self._active = False
            status = f"{error}；暂用浮窗，将自动重试"
            if self.source.taskbar_status != status:
                logging.getLogger(__name__).warning("%s", status)
            self._status(status)
            if self.source.display_mode == "taskbar":
                self.source.show()

    def _click(self):
        QTimer.singleShot(0, self, self.source.toggle_win)

    def _context_menu(self):
        QTimer.singleShot(0, self, self._show_menu)

    def _show_menu(self):
        menu = QMenu()
        menu.addAction("显示/隐藏 浮窗", self.source.toggle_win)
        menu.addAction("设置…", self.open_settings)
        menu.addSeparator()
        for mode, label in DISPLAY_MODES:
            action = menu.addAction(label)
            action.setCheckable(True)
            action.setChecked(self.source.display_mode == mode)
            action.triggered.connect(lambda checked=False, value=mode: self.source.set_display_mode(value))
        menu.exec(QCursor.pos())

    def close(self):
        self._closed = True
        self.timer.stop()
        if self.native:
            self.native.close()
            self.native = None
