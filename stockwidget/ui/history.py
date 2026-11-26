"""Shared history window/popup and bounded asynchronous request lifecycle."""

import sys
import threading

from PySide6.QtCore import QEvent, QObject, QPoint, QRect, QSize, Qt, QTimer, QSignalBlocker, Signal
from PySide6.QtGui import QColor, QContextMenuEvent, QCursor, QFont, QFontInfo, QFontMetrics, QPainter, QPalette
from PySide6.QtWidgets import (
    QApplication, QAbstractButton, QButtonGroup, QDialog, QMenu, QSizePolicy, QToolTip, QVBoxLayout, QWidget, QWidgetItem,
)

from stockwidget.core.quote_presentation import MARKET_OVERVIEW_QUOTES, MARKET_TURNOVER_INDICES, is_cn_index
from stockwidget.core.window_rules import adjacent_popup_position, best_screen
from stockwidget.core.config_store import history_cache_dir
from stockwidget.data.bar_cache import BarCache
from stockwidget.data.bars import MA_PERIODS, BarResult, HistorySeries, history_day, merge_quote, mini_trend_points, quote_moving_average, quote_volume_ratio, select_bars, trading_date
from stockwidget.platform.capabilities import is_cocoa, is_wayland
from stockwidget.platform.window import apply_background_blur, set_macos_cursor, set_windows_arrow_cursor, sync_macos_dock_icon
from stockwidget.ui.controls.history_chart import MA_COLORS
from stockwidget.ui.generated.ui_history import Ui_HistoryDialog
from stockwidget.ui.controls.style import (
    apply_history_theme, is_dark_theme, market_overview_icon, navigation_arrow_icon,
    order_book_icon, pin_icon, theme_palette, volume_icon,
)
from stockwidget.ui.menus import build_history_menu


class HistoryDialog(QDialog):
    view_changed = Signal()
    dismissed = Signal()
    hide_quotes_requested = Signal()
    pop_out = Signal()
    size_changed = Signal(str, object)
    options_changed = Signal(object)
    instrument_changed = Signal(int)
    overview_toggled = Signal(bool)

    def __init__(self, parent, display_mode="popup"):
        floating = display_mode == "popup"
        flags = (Qt.Tool | Qt.FramelessWindowHint | Qt.NoDropShadowWindowHint if floating else
                 Qt.Window | Qt.WindowTitleHint | Qt.WindowSystemMenuHint | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)
        super().__init__(parent if floating else None, flags)
        self.display_mode = display_mode
        self._restoring = False
        self._saved_size = None
        self._preferred_size = None
        self._gesture = None
        self._drag_started = False
        self._last_press_on_button = False
        self._outside_grab = False
        self._context_menu = None
        self._native_resize_cursor = False
        self._book_enabled = False
        self._breadth_mode = False
        self._overview_mode = False
        self._average_enabled = True
        self._bg_blur = False
        self._bg_dark = True
        self._theme_palette = None
        self.setMouseTracking(True)
        self.setAttribute(Qt.WA_TranslucentBackground, floating)
        self.setModal(False)
        self.ui = Ui_HistoryDialog()
        self.ui.setupUi(self)
        for name in ("title_row", "title", "volume_button", "order_book_button", "overview_button",
                     "pin_button", "pop_out_button", "close_button",
                     "quote_header", "chart", "market_overview", "order_book", "status"):
            setattr(self, name, getattr(self.ui, name))
        self.quote_header.bind_fields()
        for name, direction in (("previous", -1), ("next", 1)):
            button = getattr(self.ui, f"btn_{name}")
            button.setIconSize(QSize(14, 14))
            button.clicked.connect(lambda checked=False, direction=direction: self.instrument_changed.emit(direction))
        self._set_navigation_icons()
        for button in (self.volume_button, self.order_book_button, self.overview_button, self.pin_button):
            button.setIconSize(QSize(14, 14))
        self.overview_button.toggled.connect(self._overview_changed)
        self.pin_button.toggled.connect(self._pin_changed)
        self.setAttribute(Qt.WA_MacAlwaysShowToolWindow, floating)
        self.chart.flat = self.order_book.flat = self.market_overview.flat = floating
        self.chart.readout_changed.connect(self._update_readout)
        self.chart.readout_layout_changed.connect(self._sync_readout_height)
        self._title_text = self._status_text = ""
        for period in MA_PERIODS:
            getattr(self.ui, f"cb_ma{period}").set_subtext("—")
        self._left_layout = None
        if floating:
            # Popup uses top-right buttons / context menu for volume, depth/breadth, and average;
            # detached window keeps all toolbar checkboxes and its full-width layout unchanged.
            self.ui.cb_volume.hide()
            self.ui.cb_average.hide()
            self.ui.cb_order_book.hide()
            self.ui.chart_layout.removeWidget(self.ui.toolbar)
            self.ui.chart_layout.removeWidget(self.ui.ma_row)
            self.ui.chart_layout.removeWidget(self.ui.chart_readout)
            self.ui.content_layout.removeWidget(self.chart)
            self.ui.content_layout.removeWidget(self.market_overview)
            self._left_layout = QVBoxLayout()
            self._left_layout.setContentsMargins(0, 0, 0, 0)
            self._left_layout.setSpacing(4)
            self._left_layout.addWidget(self.ui.toolbar, 0)
            self._left_layout.addWidget(self.ui.ma_row, 0)
            self._left_layout.addWidget(self.chart, 1)
            self._left_layout.addWidget(self.market_overview, 1)
            self._left_layout.addWidget(self.ui.chart_readout, 0)
            self.ui.content_layout.insertLayout(0, self._left_layout, 1)
            self.ui.chart_layout.setStretch(self.ui.chart_layout.indexOf(self.ui.content_layout), 1)
        self._sync_fixed_row_heights()
        if not floating:
            self.ui.chart_layout.setContentsMargins(12, 10, 12, 10)
            self.ui.chart_layout.setSpacing(8)
            self.ui.content_layout.setSpacing(8)
            self.ui.ma_layout.setSpacing(8)
            self.ui.quote_header_layout.setContentsMargins(12, 8, 12, 8)
            self.quote_header.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.title_row.setVisible(floating)
        self.quote_header.setVisible(not floating)
        self.market_overview.hide()
        self.order_book.hide()
        self.ui.ma_row.hide()
        self.status.hide()
        self._sync_readout_height()
        self.view_group = QButtonGroup(self)
        self._view = "intraday"
        self.view_buttons = {view: getattr(self.ui, f"btn_{view}") for view in ("intraday", "five_day", "daily")}
        for button in self.view_buttons.values():
            self.view_group.addButton(button)
            button.setFixedWidth(button.fontMetrics().horizontalAdvance(button.text()) + 12)
        self.view_group.buttonToggled.connect(self._view_changed)
        self.pop_out_button.clicked.connect(self.pop_out)
        self.close_button.clicked.connect(self.close)
        for control, key in ((self.ui.cb_order_book, "chart_order_book_enabled"),
                             (self.order_book_button, "chart_order_book_enabled"),
                             (self.ui.cb_volume, "chart_volume_enabled"),
                             (self.volume_button, "chart_volume_enabled"),
                             (self.ui.cb_average, "chart_average_enabled")):
            control.toggled.connect(lambda checked, key=key: self.options_changed.emit({key: checked}))
        self._sync_top_buttons()
        for period in MA_PERIODS:
            getattr(self.ui, f"cb_ma{period}").toggled.connect(self._ma_changed)
        self.setWindowOpacity(.92 if floating else 1.)
        self._save_timer = QTimer(self)
        self._save_timer.setSingleShot(True)
        self._save_timer.setInterval(200)
        self._save_timer.timeout.connect(self.save_size)

    def _sync_top_buttons(self):
        book_label = "涨跌分布" if (self._overview_mode or self._breadth_mode) else "五档"
        self.order_book_button.setToolTip(book_label)
        self.order_book_button.setAccessibleName(book_label)
        self.order_book_button.setEnabled(self._overview_mode or self.view == "intraday")
        self.ui.cb_average.setEnabled(not self._overview_mode)

    def set_chart_options(self, options, unit_mode):
        restoring, self._restoring = self._restoring, True
        if self.display_mode == "popup" and self.pin_button.isChecked() != options.chart_popup_pinned:
            with QSignalBlocker(self.pin_button):
                self.pin_button.setChecked(options.chart_popup_pinned)
            self._sync_outside_grab()
        self.chart.set_options(options.chart_ma_periods, options.chart_average_enabled,
                               options.chart_volume_enabled, unit_mode)
        turnover_changed = self.market_overview.set_show_turnover(options.chart_volume_enabled)
        for control, checked in ((self.ui.cb_order_book, options.chart_order_book_enabled),
                                 (self.order_book_button, options.chart_order_book_enabled),
                                 (self.ui.cb_volume, options.chart_volume_enabled),
                                 (self.volume_button, options.chart_volume_enabled),
                                 (self.ui.cb_average, options.chart_average_enabled)):
            with QSignalBlocker(control):
                control.setChecked(checked)
        for period, color in zip(MA_PERIODS, MA_COLORS):
            control = getattr(self.ui, f"cb_ma{period}")
            with QSignalBlocker(control):
                control.setChecked(period in options.chart_ma_periods)
            palette = QPalette(self.palette())
            palette.setColor(QPalette.WindowText, QColor(color))
            control.setPalette(palette)
        self._average_enabled = bool(options.chart_average_enabled)
        self.ui.ma_row.setVisible(not self._overview_mode and self.view == "daily" and options.chart_average_enabled)
        self.ui.cb_order_book.setEnabled(self.view == "intraday")
        self._sync_top_buttons()
        self.set_order_book(options.chart_order_book_enabled)
        self.update_size_limits()
        if turnover_changed and self._overview_mode and self._preferred_size is not None:
            self.resize(QSize(*self._preferred_size).expandedTo(self.minimumSize()).boundedTo(self.maximumSize()))
        self.layout().activate()
        if self.display_mode == "popup":
            self._keep_within_screen()
        self._restoring = restoring

    def _ma_changed(self):
        self.options_changed.emit({"chart_ma_periods": [period for period in MA_PERIODS
                                                       if getattr(self.ui, f"cb_ma{period}").isChecked()]})

    @property
    def overview_mode(self) -> bool:
        return self._overview_mode

    def set_overview_mode(self, enabled: bool):
        overview = bool(enabled) and self.display_mode == "popup"
        if self.overview_button.isChecked() != overview:
            with QSignalBlocker(self.overview_button):
                self.overview_button.setChecked(overview)
        if self._overview_mode == overview:
            return
        restoring, self._restoring = self._restoring, True
        self._overview_mode = overview
        if overview:
            self.ui.toolbar.hide()
            self.ui.ma_row.hide()
            self.chart.hide()
            self.ui.chart_readout.hide()
            self.status.hide()
            self.market_overview.show()
            self.order_book.set_breadth_mode(True)
            self.order_book.setVisible(self._book_enabled)
        else:
            self.market_overview.hide()
            self.ui.toolbar.show()
            self.ui.ma_row.setVisible(self.view == "daily" and self._average_enabled)
            self.chart.show()
            self.ui.chart_readout.show()
            self.status.setVisible(bool(self._status_text))
            self.order_book.set_breadth_mode(self._breadth_mode)
            self.order_book.setVisible(self._book_enabled and self.view == "intraday")
        self._sync_top_buttons()
        self.update_size_limits()
        if self._preferred_size is not None:
            self.resize(QSize(*self._preferred_size).expandedTo(self.minimumSize()).boundedTo(self.maximumSize()))
        self.layout().activate()
        self._sync_readout_height()
        if self.display_mode == "popup":
            self._keep_within_screen()
        self._restoring = restoring

    def _overview_changed(self, checked):
        if self.display_mode != "popup":
            return
        self.set_overview_mode(checked)
        self.overview_toggled.emit(self._overview_mode)

    def set_order_book(self, enabled):
        restoring, self._restoring = self._restoring, True
        changed = self._book_enabled != bool(enabled)
        self._book_enabled = bool(enabled)
        self.order_book.setVisible(self._book_enabled if self._overview_mode else (self._book_enabled and self.view == "intraday"))
        self.update_size_limits()
        if changed and self._preferred_size is not None:
            self.resize(QSize(*self._preferred_size).expandedTo(self.minimumSize()).boundedTo(self.maximumSize()))
        self.layout().activate()
        self._sync_readout_height()
        if self.display_mode == "popup":
            self._keep_within_screen()
        self._restoring = restoring

    def set_breadth_mode(self, enabled):
        restoring, self._restoring = self._restoring, True
        breadth = bool(enabled)
        changed = self._breadth_mode != breadth
        self._breadth_mode = breadth
        self.ui.cb_order_book.setText("涨跌分布" if breadth else "五档")
        self._sync_top_buttons()
        self.order_book.set_breadth_mode(True if self._overview_mode else breadth)
        if changed:
            self.ui.toolbar.updateGeometry()
            self.update_size_limits()
            if not self._overview_mode and self._preferred_size is not None:
                self.resize(QSize(*self._preferred_size).expandedTo(self.minimumSize()).boundedTo(self.maximumSize()))
            self.layout().activate()
            self._sync_readout_height()
            if self.display_mode == "popup":
                self._keep_within_screen()
        self._restoring = restoring

    def _sync_fixed_row_heights(self):
        for row, floor in ((self.title_row, 26), (self.ui.toolbar, 24), (self.ui.ma_row, 22)):
            row.ensurePolished()
            height = max(floor, row.layout().minimumSize().height())
            if row.minimumHeight() != height or row.maximumHeight() != height:
                row.setFixedHeight(height)

    def update_size_limits(self, bounds=None):
        self.ensurePolished()
        self.status.ensurePolished()
        self._sync_fixed_row_heights()
        screen = QApplication.screenAt(self.frameGeometry().center()) or QApplication.primaryScreen()
        bounds = bounds or screen.availableGeometry()
        # Reserve native title-bar/frame space for the detached window.
        frame = self.frameGeometry().size() - self.size() if self.isVisible() else QSize(16, 40)
        screen_max = bounds.size() if self.display_mode == "popup" else bounds.size() - frame
        base_min = QSize(640, 400) if self.display_mode == "window" else QSize(300, 210)
        minimum = base_min.expandedTo(QSize(460, base_min.height())) if self._book_enabled else QSize(base_min)
        if self.layout():
            layout = self.layout()
            if self._left_layout is not None:
                self._left_layout.invalidate()
            self.ui.content_layout.invalidate()
            self.ui.ma_layout.invalidate()
            layout.invalidate()
            content = layout.minimumSize()
            margins = layout.contentsMargins()
            horizontal = margins.left() + margins.right()
            if self.display_mode == "popup":
                # Popup uses one unified minimum/maximum geometry across stock/index,
                # intraday/five-day/daily, volume/average toggles, and market overview;
                # only enabling/disabling the right depth/breadth panel changes minimum size.
                self.market_overview._sync_size()
                self.order_book._sync_size()
                self._left_layout.invalidate()
                self.ui.content_layout.invalidate()
                self.ui.toolbar.layout().invalidate()
                self.title_row.layout().invalidate()
                layout.invalidate()
                content = layout.minimumSize()
                vertical = margins.top() + margins.bottom()
                v_gap = layout.spacing()
                left_gap = self._left_layout.spacing()
                title_min = self.title_row.minimumSize().expandedTo(self.title_row.layout().minimumSize())
                toolbar_min = self.ui.toolbar.minimumSize().expandedTo(self.ui.toolbar.layout().minimumSize())
                ma_min = self.ui.ma_row.minimumSize().expandedTo(self.ui.ma_row.layout().minimumSize())
                chart_min = self.chart.minimumSize()
                overview_min = self.market_overview.minimumSize()
                readout_h = max(self.ui.chart_readout.minimumHeight(), self.ui.chart_readout.sizeHint().height())
                status_h = max(self.status.minimumHeight(), self.status.sizeHint().height(), self.status.fontMetrics().height())
                book_min_w, book_min_h = self.order_book.unified_minimum_size()
                left_intraday_w = max(toolbar_min.width(), chart_min.width(), overview_min.width())
                left_daily_w = max(left_intraday_w, ma_min.width())
                left_intraday_h = toolbar_min.height() + left_gap + chart_min.height() + left_gap + readout_h
                left_daily_h = max(left_intraday_h + left_gap + ma_min.height(), overview_min.height())
                overhead_h = vertical + title_min.height() + v_gap + v_gap + status_h
                no_book_w = horizontal + max(title_min.width(), left_daily_w)
                no_book_h = overhead_h + left_daily_h
                content_without_book = QSize(no_book_w, no_book_h)
                if self._book_enabled:
                    book_w = book_min_w + self.ui.content_layout.spacing()
                    with_book_w = max(no_book_w, horizontal + left_intraday_w + book_w)
                    with_book_h = overhead_h + max(left_daily_h, book_min_h)
                    content_with_book = QSize(with_book_w, with_book_h).expandedTo(content)
                else:
                    content_with_book = QSize(content_without_book).expandedTo(content)
                base_min = base_min.expandedTo(content_without_book)
                minimum = minimum.expandedTo(content_with_book)
            else:
                # Detached window keeps toolbar/ma_row above and readout below content_layout.
                self.order_book._sync_size()
                self.ui.content_layout.invalidate()
                layout.invalidate()
                content = layout.minimumSize()
                extra_w = 0
                for row in (self.ui.ma_row, self.status):
                    hint = QWidgetItem(row).minimumSize()
                    if row.isHidden():
                        if row.layout():
                            hint = row.minimumSize().expandedTo(row.layout().minimumSize())
                        else:
                            hint = QSize(0, max(row.minimumHeight(), row.fontMetrics().height()))
                    extra_w = max(extra_w, hint.width() + horizontal)
                    content.setWidth(max(content.width(), hint.width() + horizontal))
                    if row.isHidden():
                        content.setHeight(content.height() + hint.height() + layout.spacing())
                chart = self.ui.content_layout.minimumSize()
                book_w = self.order_book.minimumWidth() + self.ui.content_layout.spacing()
                if self._book_enabled and self.order_book.isHidden():
                    content_with_book = QSize(
                        max(content.width(), chart.width() + book_w + horizontal),
                        content.height() + max(0, self.order_book.minimumHeight() - chart.height()),
                    )
                    content_without_book = QSize(content)
                elif not self.order_book.isHidden():
                    content_with_book = QSize(content)
                    non_book_w = max(
                        extra_w,
                        self.title_row.minimumSizeHint().width() + horizontal if not self.title_row.isHidden() else 0,
                        self.ui.toolbar.minimumSizeHint().width() + horizontal,
                        self.chart.minimumWidth() + horizontal,
                    )
                    content_without_book = QSize(
                        non_book_w,
                        content.height() - max(0, self.order_book.minimumHeight() - self.chart.minimumHeight()),
                    )
                else:
                    content_with_book = QSize(content)
                    content_without_book = QSize(content)
                base_min = base_min.expandedTo(content_without_book)
                minimum = minimum.expandedTo(content_with_book)
        if self.display_mode == "popup":
            # Cap popup size at ~2x its no-order-book minimum dimensions.
            popup_cap = QSize(base_min.width() * 2, base_min.height() * 2)
            maximum = popup_cap.boundedTo(screen_max).expandedTo(minimum).boundedTo(screen_max)
        else:
            maximum = screen_max
        maximum = maximum.expandedTo(QSize(200, 200))
        self.setMinimumSize(minimum.boundedTo(maximum))
        self.setMaximumSize(maximum)

    def _keep_within_screen(self):
        if not self.isVisible() or self.isMaximized() or self.isMinimized() or self._gesture:
            return
        frame = self.frameGeometry()
        bounds = best_screen(
            *frame.getRect(),
            [screen.availableGeometry().getRect() for screen in QApplication.screens()],
        )
        if not bounds:
            return
        area = QRect(*bounds)
        x = max(area.left(), min(area.right() - frame.width() + 1, frame.left()))
        y = max(area.top(), min(area.bottom() - frame.height() + 1, frame.top()))
        if (x, y) != (frame.x(), frame.y()):
            self.move(self.pos() + QPoint(x - frame.x(), y - frame.y()))

    def restore_size(self, size, bounds):
        self._restoring = True
        self._preferred_size = list(size)
        self.update_size_limits(bounds)
        # Settle the footer's fitted font before saving the restored geometry.
        for _ in range(3):
            self.resize(QSize(*size).expandedTo(self.minimumSize()).boundedTo(self.maximumSize()))
            self.layout().activate()
            self._sync_readout_height()
            if self.size() == QSize(*size).expandedTo(self.minimumSize()).boundedTo(self.maximumSize()):
                break
        self._restoring = False

    def save_size(self):
        self._save_timer.stop()
        if self._saved_size is not None:
            size, self._saved_size = self._saved_size, None
            self.size_changed.emit(self.display_mode, size)

    def _pin_changed(self, pinned):
        if self.display_mode != "popup":
            return
        self._sync_outside_grab()
        self.options_changed.emit({"chart_popup_pinned": pinned})

    def _sync_outside_grab(self, *, closing=False):
        # The Windows taskbar and X11 panels can receive clicks without taking focus.
        handle = self.windowHandle()
        grab = (not closing and self.isVisible() and self._context_menu is None
                and not self.pin_button.isChecked() and handle is not None
                and handle.isVisible() and QApplication.platformName() in ("windows", "xcb"))
        if self._outside_grab != grab:
            grabbed = handle.setMouseGrabEnabled(grab) if handle is not None else False
            self._outside_grab = grab and grabbed

    def hideEvent(self, event):
        if self.display_mode == "popup":
            QApplication.instance().removeEventFilter(self)
            self._sync_outside_grab(closing=True)
            if menu := self._context_menu:
                menu.setAttribute(Qt.WA_NoMouseReplay)
                menu.close()
        self._gesture = None
        self._drag_started = False
        self._apply_cursor(Qt.ArrowCursor)
        self.save_size()
        super().hideEvent(event)
        QToolTip.hideText()
        if not (self.display_mode == "window" and event.spontaneous() and self.isMinimized()):
            self.dismissed.emit()
        if self.display_mode == "window":
            sync_macos_dock_icon()

    def set_background_blur(self, enabled: bool, *, dark: bool = True):
        if self.display_mode != "popup":
            return
        self._bg_blur = bool(enabled)
        self._bg_dark = bool(dark)
        apply_background_blur(self, self._bg_blur, corner_radius=10, dark=self._bg_dark)
        self.update()

    def showEvent(self, event):
        super().showEvent(event)
        self.layout().activate()
        self._elide_labels()
        if self.display_mode == "popup":
            apply_background_blur(self, self._bg_blur, corner_radius=10, dark=self._bg_dark)
            QApplication.instance().installEventFilter(self)
            # Wait until the platform window is visible before grabbing the mouse.
            QTimer.singleShot(0, self, self._sync_outside_grab)
        else:
            sync_macos_dock_icon()

    @property
    def view(self):
        return self._view

    def set_view(self, view):
        self.view_buttons[view].setChecked(True)

    def _view_changed(self, button, checked):
        if not checked:
            return
        self._view = next(view for view, candidate in self.view_buttons.items() if candidate is button)
        self._sync_top_buttons()
        self.view_changed.emit()

    def set_title(self, text):
        self._title_text = text
        self.setWindowTitle(text)
        self._elide_labels()

    def set_status(self, text):
        self._status_text = text
        self.status.setVisible(bool(text) and not self._overview_mode)
        self._elide_labels()

    def _elide_labels(self):
        for label, text in ((self.title, self._title_text), (self.status, self._status_text)):
            label.setText(label.fontMetrics().elidedText(text, Qt.ElideRight, max(1, label.width())))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.layout().activate()
        self._sync_readout_height()
        self._elide_labels()
        if self.display_mode == "popup" and self._bg_blur and is_cocoa():
            apply_background_blur(self, True, corner_radius=10, dark=self._bg_dark)
        if self.isVisible() and not self.isMaximized() and not self.isMinimized():
            if self.display_mode == "popup" or not self._restoring:
                self._keep_within_screen()
            if not self._restoring:
                self._saved_size = [self.width(), self.height()]
                self._preferred_size = list(self._saved_size)
                if not self._gesture:
                    self._save_timer.start()

    def _sync_readout_height(self):
        if not hasattr(self, "chart") or not hasattr(self.ui, "chart_readout"):
            return
        label = self.ui.chart_readout
        lines = self.chart.readout_size_sample().splitlines()
        base_font = self.font()
        metrics = QFontMetrics(base_font)
        width = max(metrics.horizontalAdvance(line) for line in lines) + 2
        # Fit whole rows in compact panels without changing font on hover.
        available = label.contentsRect().width() - 2
        scale = min(1, available / width) if available > 20 else 1
        font = QFont(base_font)
        if scale < 1:
            base_px = QFontInfo(base_font).pixelSize()
            if base_px > 0:
                font.setPixelSize(max(1, int(base_px * scale)))
            elif base_font.pointSizeF() > 0:
                font.setPointSizeF(max(1.0, base_font.pointSizeF() * scale))
        if label.font() != font:
            label.setFont(font)
        height = metrics.height() + 2
        if label.minimumHeight() != height or label.maximumHeight() != height:
            label.setFixedHeight(height)
            self.update_size_limits()

    def _update_readout(self, text):
        self.ui.chart_readout.setText(self.chart.readout_text(rich=True) if self.display_mode == "popup" else text)
        ma_values = self.chart.ma_values()
        for period in MA_PERIODS:
            getattr(self.ui, f"cb_ma{period}").set_subtext(ma_values[period])

    def paintEvent(self, event):
        if self.display_mode != "popup":
            return super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(Qt.NoPen)
        painter.setBrush(self.palette().window())
        if sys.platform == "win32" and self._bg_blur:
            painter.fillRect(self.rect(), self.palette().window())
        else:
            painter.drawRoundedRect(self.rect(), 10, 10)

    def _set_navigation_icons(self):
        ratio, palette = self.devicePixelRatioF(), self.palette()
        self.volume_button.setIcon(volume_icon(ratio, palette))
        self.order_book_button.setIcon(order_book_icon(ratio, palette))
        self.overview_button.setIcon(market_overview_icon(ratio, palette))
        self.pin_button.setIcon(pin_icon(ratio, palette))
        for name, direction in (("previous", -1), ("next", 1)):
            getattr(self.ui, f"btn_{name}").setIcon(navigation_arrow_icon(direction, ratio, palette))

    def changeEvent(self, event):
        super().changeEvent(event)
        if event.type() in {QEvent.PaletteChange, QEvent.DevicePixelRatioChange} and hasattr(self, "chart"):
            self._set_navigation_icons()
        if event.type() in {QEvent.FontChange, QEvent.DevicePixelRatioChange} and hasattr(self, "chart"):
            self._sync_fixed_row_heights()
            self._sync_readout_height()
            if self._overview_mode:
                self.update_size_limits()

    def moveEvent(self, event):
        super().moveEvent(event)
        if self.isVisible() and not self._restoring:
            self._restoring = True
            self.update_size_limits()
            self._restoring = False

    def _edges(self, position):
        if self.display_mode != "popup" or not self.rect().contains(position):
            return Qt.Edges()
        edges = Qt.Edges()
        for hit, edge in ((position.x() < 7, Qt.LeftEdge), (position.x() >= self.width() - 7, Qt.RightEdge),
                          (position.y() < 7, Qt.TopEdge), (position.y() >= self.height() - 7, Qt.BottomEdge)):
            if hit:
                edges |= edge
        return edges

    @staticmethod
    def _edge_cursor(edges):
        horizontal = bool(edges & (Qt.LeftEdge | Qt.RightEdge))
        vertical = bool(edges & (Qt.TopEdge | Qt.BottomEdge))
        if horizontal and vertical:
            return (Qt.SizeFDiagCursor if bool(edges & Qt.LeftEdge) == bool(edges & Qt.TopEdge)
                    else Qt.SizeBDiagCursor)
        if horizontal:
            return Qt.SizeHorCursor
        return Qt.SizeVerCursor if vertical else Qt.ArrowCursor

    def _apply_cursor(self, cursor):
        previous = self.cursor().shape()
        self.setCursor(cursor)
        if self.display_mode == "popup" and QApplication.platformName() == "windows":
            if cursor == Qt.ArrowCursor and previous != cursor:
                set_windows_arrow_cursor()
        if self.display_mode != "popup" or (cursor == Qt.ArrowCursor and not self._native_resize_cursor):
            return
        if set_macos_cursor(cursor):
            self._native_resize_cursor = cursor != Qt.ArrowCursor

    def _sync_cursor(self):
        self._apply_cursor(self._edge_cursor(self._edges(self.mapFromGlobal(QCursor.pos()))))

    def _start_gesture(self, position):
        edges = self._edges(self.mapFromGlobal(position))
        if edges:
            self._apply_cursor(self._edge_cursor(edges))
        self.save_size()
        self._gesture = edges, position, self.geometry()
        self._drag_started = False

    def mouseMoveEvent(self, event):
        if self._gesture:
            edges, start, original = self._gesture
            delta = event.globalPosition().toPoint() - start
            if not edges:
                if not self._drag_started:
                    if delta.manhattanLength() < QApplication.startDragDistance():
                        event.accept()
                        return
                    self._drag_started = True
                    if is_wayland() and self.windowHandle().startSystemMove():
                        self._gesture = None
                        event.accept()
                        return
                self.move(original.topLeft() + delta)
                event.accept()
                return
            width = original.width() + (delta.x() if edges & Qt.RightEdge else -delta.x() if edges & Qt.LeftEdge else 0)
            height = original.height() + (delta.y() if edges & Qt.BottomEdge else -delta.y() if edges & Qt.TopEdge else 0)
            width = max(self.minimumWidth(), min(self.maximumWidth(), width))
            height = max(self.minimumHeight(), min(self.maximumHeight(), height))
            x = original.right() - width + 1 if edges & Qt.LeftEdge else original.x()
            y = original.bottom() - height + 1 if edges & Qt.TopEdge else original.y()
            screen = QApplication.screenAt(original.center()) or QApplication.primaryScreen()
            bounds = screen.availableGeometry()
            self.setGeometry(max(bounds.left(), min(bounds.right() - width + 1, x)),
                             max(bounds.top(), min(bounds.bottom() - height + 1, y)), width, height)
            event.accept()
            return
        self._apply_cursor(self._edge_cursor(self._edges(event.position().toPoint())))
        super().mouseMoveEvent(event)

    def enterEvent(self, event):
        super().enterEvent(event)
        if self.display_mode == "popup" and not self._gesture:
            self._apply_cursor(self._edge_cursor(self._edges(event.position().toPoint())))

    def leaveEvent(self, event):
        if self.display_mode == "popup" and not self._gesture:
            self._apply_cursor(Qt.ArrowCursor)
        super().leaveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._gesture and event.button() == Qt.LeftButton:
            moved = not self._gesture[0] and self._drag_started
            self._gesture = None
            self._drag_started = False
            if moved:
                bounds = best_screen(*self.geometry().getRect(),
                                     [screen.availableGeometry().getRect() for screen in QApplication.screens()])
                if bounds:
                    area = QRect(*bounds)
                    self.move(max(area.left(), min(area.right() - self.width() + 1, self.x())),
                              max(area.top(), min(area.bottom() - self.height() + 1, self.y())))
            self._sync_cursor()
            self.save_size()
            event.accept()
        else:
            super().mouseReleaseEvent(event)

    def keyPressEvent(self, event):
        if event.key() != Qt.Key_Escape:
            return super().keyPressEvent(event)
        if self._gesture:
            original = self._gesture[2]
            self._gesture = None
            self._drag_started = False
            self.setGeometry(original)
            self._sync_cursor()
            self.save_size()
        elif self.display_mode != "popup" or not self.pin_button.isChecked():
            return super().keyPressEvent(event)
        event.accept()

    def contextMenuEvent(self, event):
        if self.display_mode != "popup":
            return super().contextMenuEvent(event)
        if not self.rect().contains(event.pos()):
            event.ignore()
            return
        if self._context_menu is None:
            menu = self._context_menu = build_history_menu(self)
            menu.aboutToHide.connect(self._context_menu_hidden)
            self._sync_outside_grab()
            menu.popup(event.globalPos())
        event.accept()

    def _context_menu_hidden(self):
        menu, self._context_menu = self._context_menu, None
        menu.deleteLater()
        # Qt releases the menu's mouse grab after emitting aboutToHide.
        QTimer.singleShot(0, self, self._sync_outside_grab)

    def eventFilter(self, obj, event):
        kind = event.type()
        if kind == QEvent.ApplicationDeactivate and not self.pin_button.isChecked():
            self.close()
        if (kind == QEvent.ContextMenu and obj is self.windowHandle()
                and event.reason() == QContextMenuEvent.Mouse):
            # After a menu closes, Qt can still have another window as its last mouse receiver.
            self.contextMenuEvent(event)
            return True
        if kind not in (QEvent.MouseButtonPress, QEvent.MouseButtonDblClick, QEvent.MouseButtonRelease,
                        QEvent.MouseMove, QEvent.Enter, QEvent.KeyPress) or not isinstance(obj, QWidget):
            return False
        inside = obj.window() is self
        if kind in (QEvent.MouseButtonPress, QEvent.MouseButtonDblClick):
            position = event.globalPosition().toPoint()
            point = self.mapFromGlobal(position)
            over_chart = self.rect().contains(point)
            if menu := self._context_menu:
                if over_chart or any(
                        popup.isVisible() and popup.rect().contains(popup.mapFromGlobal(position))
                        for popup in (menu, *menu.findChildren(QMenu))):
                    return False
            if not inside or not over_chart:
                if not self.pin_button.isChecked():
                    self.close()
                    return True
            elif event.button() == Qt.LeftButton:
                over_button = isinstance(self.childAt(point), QAbstractButton)
                if kind == QEvent.MouseButtonPress:
                    self._last_press_on_button = over_button
                    if not over_button:
                        self._start_gesture(position)
                        return True
                elif not over_button and not self._last_press_on_button:
                    self.hide_quotes_requested.emit()
                    self.close()
                    return True
                else:
                    return False
            return False
        if not inside or obj is self:
            return False
        # Keep move and release with the same gesture across chart children.
        if self._gesture:
            if kind == QEvent.MouseMove:
                self.mouseMoveEvent(event)
                return True
            if kind == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
                self.mouseReleaseEvent(event)
                return True
            if kind == QEvent.KeyPress and event.key() == Qt.Key_Escape:
                self.keyPressEvent(event)
                return True
        elif (kind in (QEvent.Enter, QEvent.MouseMove)
              and (self._native_resize_cursor or self.cursor().shape() != Qt.ArrowCursor)):
            # Child entry can bypass the dialog's border mouse move.
            self._apply_cursor(Qt.ArrowCursor)
        return False


class HistoryController(QObject):
    result_ready = Signal(object)

    def __init__(self, window):
        super().__init__(window)
        self.window = window
        for block, table in enumerate(window.float_tables):
            window.register_drag_region(table.viewport(),
                                        lambda pos, table=table, block=block: self.request_row(
                                            "float", table.indexAt(pos).row(), block))
        self.dialog = None
        self.instrument = None
        self.cache = BarCache(history_cache_dir())
        self._generation = 0
        self._discard_before = 0
        self._busy = False
        self._pending = None
        self._displayed_key = None
        self._series = {}
        self._history_pending = {}
        self._history_scheduled = False
        self.result_ready.connect(self._accept)
        self.window.view_options_changed.connect(self._options_changed)
        self.window.configuration_changed.connect(self._apply_theme)
        QApplication.instance().paletteChanged.connect(self._apply_theme)
        self.window.display_flags_changed.connect(self._apply_chart_options)
        self.window.quotes.quotes_updated.connect(self._quotes_updated)
        self.window.quotes.details_updated.connect(self._update_details)
        self.window.presentation_changed.connect(self.schedule_history_refresh)
        self.click_timer = QTimer(self)
        self.click_timer.setSingleShot(True)
        self.click_timer.setTimerType(Qt.PreciseTimer)
        self.click_timer.timeout.connect(self._open_pending)
        self._clicked_instrument = None
        self._surface = "float"
        self._taskbar_position = QPoint()
        self.window.widget_visibility_changed.connect(self._visibility_changed)

    def request_row(self, surface, row, block=0, *, global_pos=None):
        if not self.window.view_options.chart_enabled:
            return
        self._clicked_instrument = self.window.quotes.instrument_at(surface, row, block)
        if self._clicked_instrument:
            self._set_anchor(surface, global_pos)
            # Preserve the existing double-click-to-hide gesture without a popup.
            self.click_timer.start(QApplication.doubleClickInterval())

    def _open_pending(self):
        if self._clicked_instrument and self.window.view_options.chart_enabled and self.window.widget_visible:
            self._open(self._clicked_instrument)

    def _visibility_changed(self):
        if not self.window.widget_visible:
            self.click_timer.stop()
            self._closed()
            self._discard_before = self._generation
            self._history_pending.clear()
            if self.dialog:
                self.dialog.close()
        else:
            self.schedule_history_refresh()

    def open_row(self, surface, row, block=0, *, global_pos=None):
        if not self.window.view_options.chart_enabled:
            return
        instrument = self.window.quotes.instrument_at(surface, row, block)
        if instrument is None:
            return
        self._set_anchor(surface, global_pos)
        self._open(instrument)

    def _set_anchor(self, surface, global_pos):
        self._surface = surface
        if surface == "taskbar":
            self._taskbar_position = QPoint(global_pos if global_pos is not None else QCursor.pos())

    def _sync_dialog_title(self):
        if not self.dialog:
            return
        if self.dialog.overview_mode:
            self.dialog.set_title("大盘总览")
        elif self.instrument:
            self.dialog.set_title(str(self.instrument.get("name") or self.instrument.get("code") or ""))

    def _open(self, instrument):
        if self.dialog is None or (self.dialog.display_mode == "window" and not self.dialog.isVisible()):
            previous = self.dialog
            self.dialog = self._create_dialog("popup")
            if previous:
                previous.deleteLater()
        elif self.dialog.overview_mode:
            self.dialog.set_overview_mode(False)
            self.window.quotes.sync_market_overview(False)
        self.instrument = dict(instrument)
        self._sync_dialog_title()
        self._apply_theme()
        self._apply_chart_options()
        self._show()
        self.reload()

    def _create_dialog(self, mode):
        dialog = HistoryDialog(self.window, mode)
        self.window.destroyed.connect(dialog.deleteLater)
        dialog.view_changed.connect(self._view_changed)
        dialog.dismissed.connect(self._closed)
        dialog.hide_quotes_requested.connect(self.window.hide_widget)
        dialog.pop_out.connect(self._pop_out)
        dialog.size_changed.connect(self._save_size)
        dialog.options_changed.connect(lambda changes: self.window.set_view_options(**changes))
        dialog.instrument_changed.connect(self._navigate)
        dialog.overview_toggled.connect(self._overview_toggled)
        return dialog

    def _navigate(self, direction):
        if not self.dialog or self.dialog.display_mode != "window" or not self.dialog.isVisible():
            return
        instruments = list(self.window.checked_codes.values())
        if not instruments:
            self.dialog.close()
            return
        identity = lambda entry: (entry.get("market"), entry.get("code"))
        current = next((i for i, entry in enumerate(instruments) if identity(entry) == identity(self.instrument or {})), -1)
        if current == -1 and direction < 0:
            current = 0
        target = instruments[(current + direction) % len(instruments)]
        if identity(target) != identity(self.instrument or {}):
            self._open(target)

    def _save_size(self, mode, size):
        self.window.set_view_options(**{f"chart_{mode}_size": size})

    def _view_changed(self):
        self.window.quotes.sync_market_breadth(self.needs_market_breadth())
        self._apply_chart_options()
        self.reload()

    def needs_market_overview(self) -> bool:
        return bool(
            self.window.widget_visible
            and self.dialog is not None
            and self.dialog.isVisible()
            and not self.dialog.isMinimized()
            and self.dialog.overview_mode
        )

    def needs_market_breadth(self) -> bool:
        if not (
            self.window.widget_visible
            and self.dialog is not None
            and self.dialog.isVisible()
            and not self.dialog.isMinimized()
            and self.window.view_options.chart_order_book_enabled
        ):
            return False
        if self.dialog.overview_mode:
            return True
        return self.dialog.view == "intraday" and is_cn_index(self.instrument)

    def _overview_toggled(self, enabled):
        if not self.dialog:
            return
        self._sync_dialog_title()
        self.window.quotes.sync_market_overview(self.needs_market_overview())
        self.window.quotes.sync_market_breadth(self.needs_market_breadth())
        if enabled:
            self._update_details()
            self._start_next_background()
            overview_quotes = self.window.quotes.market_overview_quotes()
            if (len(overview_quotes) < len(MARKET_OVERVIEW_QUOTES)
                    or (self.needs_market_breadth() and self.window.quotes.market_breadth() is None)):
                self.window.quotes.refresh()
        else:
            self._apply_chart_options()
            self.reload()

    def _pop_out(self):
        if not self.dialog or self.dialog.display_mode != "popup":
            return
        previous = self.dialog
        was_overview = previous.overview_mode
        previous.save_size()
        # Separate native handles prevent queued Popup close/resize events from
        # affecting the detached window. Keep request generation and series.
        previous.dismissed.disconnect(self._closed)
        self.dialog = self._create_dialog("window")
        self.dialog.set_view(previous.view)
        self._sync_dialog_title()
        self.dialog.set_status(previous._status_text)
        previous.close()
        previous.deleteLater()
        self.window.quotes.sync_market_overview(False)
        self._apply_theme()
        self._apply_chart_options()
        self._show()
        series = self._series_for(self.instrument, self.dialog.view)
        if series.loaded_day is not None:
            self._render(series, self.dialog.view)
        elif was_overview:
            self.reload()
        self._start_next_background()

    def _show(self):
        self._place_floating(restore=not self.dialog.isVisible())
        self.dialog.show()
        self.window.quotes.sync_market_overview(self.needs_market_overview())
        self.window.quotes.sync_market_breadth(self.needs_market_breadth())
        self.window.quotes.set_detail_instrument(self.instrument if self.dialog.display_mode == "window" else None)
        self._update_details()
        self.dialog.raise_()
        self.dialog.activateWindow()

    def _place_floating(self, *, restore=False):
        self.dialog.layout().activate()
        anchor = (self.window.position_controller.full_geometry() if self._surface == "float"
                  else QRect(self._taskbar_position, QSize(1, 1)))
        bounds = best_screen(*anchor.getRect(), [screen.availableGeometry().getRect()
                                                for screen in QApplication.screens()])
        if bounds:
            if restore:
                size = getattr(self.window.view_options, f"chart_{self.dialog.display_mode}_size")
                self.dialog.restore_size(size, QRect(*bounds))
            if self.dialog.display_mode == "popup" and restore:
                self.dialog.move(*adjacent_popup_position(anchor.getRect(), bounds,
                                                         self.dialog.width(), self.dialog.height(),
                                                         prefer_above=self._surface == "taskbar"))
            elif restore:
                rect = QRect(QPoint(), self.dialog.size())
                rect.moveCenter(QRect(*bounds).center())
                self.dialog.move(rect.topLeft())

    def _apply_theme(self):
        if self.dialog is None:
            return
        options = self.window.view_options
        mode = options.color_mode
        dark = mode == "dark" or (mode == "system" and is_dark_theme())
        palette = theme_palette(dark, QApplication.palette())
        floating = self.dialog.display_mode == "popup"
        colors = (QColor(self.window.fg), QColor(self.window.bg)) if floating else (None, None)
        blur = floating and self.window.chart_blur_active
        opacity = 1.
        if floating:
            colors[1].setAlpha(round(options.chart_bg_alpha_pct * 2.55))
            opacity = 1. if blur else options.chart_opacity_pct / 100
        theme = (palette, *colors, QColor(self.window.up_color), QColor(self.window.down_color),
                 QColor(self.window.neutral_color), opacity, blur)
        if self.dialog._theme_palette == theme:
            return
        if self.dialog._theme_palette is None or self.dialog._theme_palette[:3] != theme[:3]:
            apply_history_theme(self.dialog, dark, palette, *colors)
            self.dialog.update_size_limits()
        self.dialog._theme_palette = theme
        if floating:
            if not is_wayland():
                self.dialog.setWindowOpacity(opacity)
            self.dialog.set_background_blur(blur, dark=colors[1].lightness() < 128)
        self.dialog.chart.up_color = self.window.up_color
        self.dialog.chart.down_color = self.window.down_color
        self.dialog.chart.neutral_color = self.window.neutral_color if floating else None
        self.dialog.order_book.neutral_color = self.window.neutral_color if floating else None
        self.dialog.market_overview.neutral_color = self.window.neutral_color if floating else None
        self.dialog.chart._emit_readout()
        self.dialog.update()
        self._update_details()

    def _options_changed(self):
        self.schedule_history_refresh()
        if not self.window.view_options.chart_enabled:
            self.click_timer.stop()
            if self.dialog:
                self.dialog.close()
        else:
            self.window.quotes.sync_market_overview(self.needs_market_overview())
            self.window.quotes.sync_market_breadth(self.needs_market_breadth())
            self._apply_theme()
            self._apply_chart_options()
            if self.needs_market_overview():
                self._start_next_background()
                if self.needs_market_breadth() and self.window.quotes.market_breadth() is None:
                    self.window.quotes.refresh()

    def _apply_chart_options(self):
        if self.dialog:
            options = self.window.view_options
            self.dialog.set_breadth_mode(is_cn_index(self.instrument))
            self.dialog.set_chart_options(options, self.window.unit_mode)
            self._update_details()

    def _update_details(self):
        if not self.dialog or not self.instrument:
            return
        if self.dialog.overview_mode:
            breadth = self.window.quotes.market_breadth()
            self.dialog.order_book.set_quote(
                self.instrument, None, self.window.unit_mode,
                self.window.up_color, self.window.down_color,
                market_breadth=breadth, force_breadth=True,
            )
            overview_quotes = self.window.quotes.market_overview_quotes()
            turnover_bars = {
                key: self._series[series_key].result.bars
                for key, inst in MARKET_TURNOVER_INDICES
                if (series_key := self._key(inst, "intraday")) in self._series
            }
            self.dialog.market_overview.set_overview(
                overview_quotes, turnover_bars, self.window.unit_mode,
                self.window.up_color, self.window.down_color,
            )
            return
        if (self.dialog.display_mode == "window" and self.dialog.isVisible()
                and not any((entry.get("market"), entry.get("code")) == (self.instrument.get("market"), self.instrument.get("code"))
                            for entry in self.window.checked_codes.values())):
            self._navigate(1)
            return
        quote = self.window.quotes.quote_for(self.instrument)
        breadth = self.window.quotes.market_breadth() if is_cn_index(self.instrument) else None
        self.dialog.order_book.set_quote(
            self.instrument, quote, self.window.unit_mode,
            self.window.up_color, self.window.down_color,
            market_breadth=breadth,
        )
        if self.dialog.display_mode == "window":
            daily = self._series_for(self.instrument, "daily")
            ratio = None if daily.result.stale else quote_volume_ratio(daily.result.bars, quote, self.instrument)
            self.dialog.quote_header.set_quote(self.instrument, quote, self.window.unit_mode,
                                               self.window.up_color, self.window.down_color, ratio,
                                               self.window.quotes.details_for(self.instrument))
            enabled = len(self.window.checked_codes) > 1
            self.dialog.ui.btn_previous.setEnabled(enabled)
            self.dialog.ui.btn_next.setEnabled(enabled)

    def _closed(self, *_args):
        self.click_timer.stop()
        self._clicked_instrument = None
        self._generation += 1
        self._pending = None
        self.window.quotes.set_detail_instrument(None)
        self.window.quotes.sync_market_overview(False)
        self.window.quotes.sync_market_breadth(False)

    def invalidate_quotes(self):
        """Watchlist changes invalidate both consumers of history."""
        self._closed()
        self._discard_before = self._generation
        self._history_pending.clear()
        self._series.clear()
        self._displayed_key = None
        if self.window.view_options.mini_chart_type == "trend":
            self.window.quotes.reproject()

    def clear_cache(self):
        cleared = self.cache.clear()
        self._discard_cached_history()
        return cleared

    def _discard_cached_history(self):
        self.click_timer.stop()
        self._clicked_instrument = None
        self._closed()
        self.window.quotes.clear_details()
        self._discard_before = self._generation
        self._series.clear()
        self._history_pending.clear()
        self._displayed_key = None
        if self.dialog:
            self.dialog.close()
            self.dialog.chart.set_data((), self.dialog.view, self.instrument or {})
            self.dialog.market_overview.set_overview(
                {}, {}, self.window.unit_mode, self.window.up_color, self.window.down_color
            )

    @staticmethod
    def _key(instrument, view):
        return (*(instrument.get(field) for field in ("market", "code", "type")),
                "daily" if view == "daily" else "five_day")

    def _series_for(self, instrument, view):
        key = self._key(instrument, view)
        if key not in self._series:
            self._series[key] = HistorySeries(instrument, view)
        series = self._series[key]
        quote = self.window.quotes.quote_for(instrument)
        if quote:
            series.update_quote(quote)
        return series

    def _quotes_updated(self):
        if self.dialog and self.dialog.isVisible() and self.dialog.display_mode == "window" and self.instrument:
            self.window.quotes.set_detail_instrument(self.instrument)
        for series in self._series.values():
            quote = self.window.quotes.quote_for(series.instrument)
            if quote:
                series.update_quote(quote)
        if self.dialog and self.dialog.isVisible() and self.instrument:
            self._update_details()
            if not self.dialog.overview_mode:
                series = self._series_for(self.instrument, self.dialog.view)
                if series.loaded_day is not None:
                    self._render(series, self.dialog.view)
                if not self._busy and series.needs_download(self.cache.clock()):
                    self.reload()
            self._start_next_background()
        self.window.quotes.reproject()
        self.schedule_history_refresh()

    def quote_metrics(self, instrument, quote):
        series = self._series.get(self._key(instrument, "daily"))
        if series is None or series.result.stale:
            return {}
        return {"volume_ratio": quote_volume_ratio(series.result.bars, quote, instrument),
                "average_price": quote_moving_average(series.result.bars, quote, instrument,
                                                      self.window.view_options.average_period)}

    def mini_chart_data(self, instrument):
        """Use the popup's minute series and live quote at the same cadence."""
        series = self._series_for(instrument, "intraday")
        bars = series.result.bars
        quote = self.window.quotes.quote_for(instrument)
        if quote:
            bars = merge_quote(bars, instrument, "five_day", quote)
        reference = series.reference_price_for(bars)
        bars = select_bars(bars, "intraday", instrument)
        return mini_trend_points(bars, instrument, self.cache.clock()), reference

    def schedule_history_refresh(self):
        if not self._history_scheduled:
            self._history_scheduled = True
            QTimer.singleShot(0, self, self._refresh_background)

    def _refresh_background(self):
        self._history_scheduled = False
        self._history_pending = {self._key(instrument, view): (instrument, view)
                              for view in ("intraday", "daily")
                              for instrument in self.window.quotes.history_instruments(view).values()}
        self._start_next_background()

    def _turnover_baseline_needed(self, instrument, now) -> bool:
        day = history_day(instrument, now)
        series = self._series.get(self._key(instrument, "intraday"))
        if series is None:
            return True
        if series.loaded_day == day and series.result.bars:
            return False
        return now.timestamp() >= series.retry_after

    def _start_next_background(self):
        if self._busy:
            return
        if self.needs_market_overview() and self.window.view_options.chart_volume_enabled:
            now = self.cache.clock()
            for _key, instrument in MARKET_TURNOVER_INDICES:
                if self._turnover_baseline_needed(instrument, now):
                    self._start((self._generation, dict(instrument), "intraday",
                                 history_day(instrument, now), False))
                    return
        if (self.dialog and self.dialog.isVisible() and self.dialog.display_mode == "window"
                and self.instrument and self.instrument.get("market") in {"sh", "sz", "bj"}
                and self.instrument.get("type") not in {"期", "指"}):
            now = self.cache.clock()
            series = self._series_for(self.instrument, "daily")
            if series.needs_download(now):
                self._start((self._generation, dict(self.instrument), "daily", history_day(self.instrument, now), False))
                return
        wanted = {self._key(entry, view) for view in ("intraday", "daily")
                  for entry in self.window.quotes.history_instruments(view).values()}
        while self._history_pending:
            key = next(iter(self._history_pending))
            instrument, view = self._history_pending.pop(key)
            if key not in wanted:
                continue
            now = self.cache.clock()
            series = self._series_for(instrument, view)
            if series.needs_download(now):
                self._start((self._generation, dict(instrument), view,
                             history_day(instrument, now), series.loaded_day == history_day(instrument, now)))
                return

    def reload(self):
        if (not self.dialog or not self.dialog.isVisible() or not self.instrument
                or not self.window.view_options.chart_enabled or self.dialog.overview_mode):
            return
        self._generation += 1
        view = self.dialog.view
        series = self._series_for(self.instrument, view)
        now = self.cache.clock()
        day = history_day(self.instrument, now)
        request = (self._generation, dict(self.instrument), view, day,
                   series.loaded_day == day)
        key = (self._key(self.instrument, view), view)
        if key != self._displayed_key:
            self.dialog.chart.set_data((), self.dialog.view, self.instrument)
        if not series.needs_download(now):
            self._pending = None
            self.dialog.chart.set_loading(False)
            self._render(series, view)
            self._start_next_background()
            return
        if not self.dialog.chart.bars:
            self.dialog.chart.set_loading(True)
            self.dialog.set_status("")
        if self._busy:
            self._pending = request  # At most one worker; latest selection wins.
        else:
            self._start(request)

    def _start(self, request):
        self._busy = True
        self._pending = None
        threading.Thread(target=self._worker, args=(request, self.cache.generation), daemon=True).start()

    def _worker(self, request, cache_generation):
        generation, instrument, view, day, repair = request
        try:
            result = self.cache.get(instrument, "daily" if view == "daily" else "five_day",
                                    repair=repair, generation=cache_generation)
        except Exception:
            result = BarResult(message="暂无可用历史数据 · 历史数据处理失败")
        try:
            self.result_ready.emit((*request, result))
        except RuntimeError:
            pass  # Parent was destroyed while the network request was running.

    def _accept(self, payload):
        self._busy = False
        generation, instrument, view, day, repair, result = payload
        if generation >= self._discard_before and self.window.widget_visible:
            series = self._series_for(instrument, view)
            series.set_history(result, day, self.cache.clock())
            if self.dialog and self.dialog.isVisible():
                self._update_details()
            if (generation == self._generation and self.dialog and self.dialog.isVisible() and self.instrument
                    and not self.dialog.overview_mode
                    and self._key(instrument, view) == self._key(self.instrument, self.dialog.view)):
                self._render(series, self.dialog.view)
            self.window.quotes.reproject()
        if self._pending is not None:
            pending, self._pending = self._pending, None
            if (pending[0] == self._generation and self.dialog
                    and self.dialog.isVisible() and not self.dialog.overview_mode):
                # The completed worker may have loaded the same minute family
                # while the user switched from intraday to five-day view.
                self.reload()
        self._start_next_background()

    def _render(self, series, view):
        result = series.result
        instrument = series.instrument
        bars = select_bars(result.bars, view, instrument)
        quote = self.window.quotes.quote_for(instrument) or {}
        limits = ((quote.get("low_price"), quote.get("high_price"))
                  if bars and quote.get("date") == trading_date(bars[-1], instrument) else ())
        self.dialog.chart.set_data(bars, view, instrument, reference_price=series.reference_price,
                                   session_range=limits, now=self.cache.clock())
        self._displayed_key = (self._key(instrument, view), view)
        pieces = [result.message]
        if result.stale:
            pieces.append("缓存（更新失败）")
        if bars and view == "daily" and len(bars) < 89:
            pieces.append("历史不足时，部分均线从满足周期处开始显示")
        self.dialog.set_status(" · ".join(piece for piece in pieces if piece))
