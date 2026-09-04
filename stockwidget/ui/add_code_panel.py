"""自选标的添加悬浮面板。"""

from PySide6.QtCore import (
    QModelIndex, QPoint, QPointF, QRect, QRectF, QSize, QSignalBlocker, Qt,
    QTimer, Signal,
)
from PySide6.QtGui import (
    QColor, QKeyEvent, QPainter, QPainterPath, QPalette, QPen,
    QStandardItem, QStandardItemModel,
)
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListView,
    QPushButton,
    QSizePolicy,
    QStyle,
    QStyledItemDelegate,
    QStyleOptionButton,
    QStyleOptionViewItem,
    QVBoxLayout,
    QWidget,
)

from stockwidget.core.code_search import query_search_index


ENTRY_ROLE = Qt.ItemDataRole.UserRole + 1
ADDED_ROLE = Qt.ItemDataRole.UserRole + 2
PAGE_SIZE = 10
RESULT_ROW_HEIGHT = 24
ADDED_BADGE_TEXT = "已添加"

CATEGORY_FILTER_ITEMS = (
    ("股票", "stock"),
    ("基金", "fund"),
    ("指数", "index"),
    ("期货", "futures"),
)
REGION_FILTER_ITEMS = (
    ("沪", "sh"),
    ("深", "sz"),
    ("京", "bj"),
    ("港", "hk"),
    ("美", "us"),
    ("其他", "other"),
)


def entry_display_text(entry: dict, *, added: bool = False) -> str:
    parts = (
        str(entry.get("type", "") or "").strip(),
        str(entry.get("code", "") or "").strip(),
        str(entry.get("name", "") or "").strip(),
    )
    label = "/".join(part for part in parts if part)
    return f"（已添加）{label}" if added else label


class FilledCheckBox(QCheckBox):
    """保留 Qt 原生布局和交互，仅覆盖指示器的状态填充。"""

    _INDICATOR_SIZE = 13
    _INDICATOR_TEXT_GAP = 3
    _HOVER_HORIZONTAL_PADDING = 4
    _HOVER_VERTICAL_PADDING = 3

    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self.set_theme(False)

    def set_theme(self, dark: bool):
        if dark:
            self._indicator_border = QColor(255, 255, 255, 115)
            self._unchecked_fill = QColor(255, 255, 255, 31)
            self._checked_fill = QColor(10, 132, 255)
            self._partial_fill = QColor(10, 132, 255, 148)
            self._hover_fill = QColor(255, 255, 255, 18)
        else:
            self._indicator_border = QColor(0, 0, 0, 97)
            self._unchecked_fill = QColor(0, 0, 0, 20)
            self._checked_fill = QColor(0, 122, 255)
            self._partial_fill = QColor(0, 122, 255, 148)
            self._hover_fill = QColor(0, 122, 255, 15)
        self.update()

    def sizeHint(self):
        metrics = self.fontMetrics()
        width = (
            2 * self._HOVER_HORIZONTAL_PADDING
            + self._INDICATOR_SIZE
            + self._INDICATOR_TEXT_GAP
            + metrics.horizontalAdvance(self.text())
        )
        height = (
            max(self._INDICATOR_SIZE, metrics.height())
            + 2 * self._HOVER_VERTICAL_PADDING
        )
        return QSize(width, height)

    def minimumSizeHint(self):
        return self.sizeHint()

    def _indicator_rect(self):
        return QRect(
            self._HOVER_HORIZONTAL_PADDING,
            (self.height() - self._INDICATOR_SIZE) // 2,
            self._INDICATOR_SIZE,
            self._INDICATOR_SIZE,
        )

    def _text_rect(self, indicator: QRect):
        left = indicator.right() + 1 + self._INDICATOR_TEXT_GAP
        return QRect(
            left,
            0,
            max(0, self.width() - self._HOVER_HORIZONTAL_PADDING - left),
            self.height(),
        )

    def paintEvent(self, event):
        option = QStyleOptionButton()
        self.initStyleOption(option)
        style = self.style()
        indicator = self._indicator_rect()
        contents = self._text_rect(indicator)
        state = self.checkState()
        if state == Qt.CheckState.Checked:
            fill = QColor(self._checked_fill)
        elif state == Qt.CheckState.PartiallyChecked:
            fill = QColor(self._partial_fill)
        else:
            fill = QColor(self._unchecked_fill)
        border = QColor(
            self._checked_fill
            if state != Qt.CheckState.Unchecked
            else self._indicator_border
        )
        if not self.isEnabled():
            fill.setAlpha(round(fill.alpha() * 0.45))
            border.setAlpha(round(border.alpha() * 0.45))

        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        if self.isEnabled() and self.underMouse():
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(self._hover_fill)
            hover_rect = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
            painter.drawRoundedRect(hover_rect, 4, 4)

        # 不调用原生 CE_CheckBox，彻底避免系统 indicator 的阴影边缘残留。
        painter.setPen(QPen(border, 1))
        painter.setBrush(fill)
        indicator_rect = QRectF(indicator).adjusted(0.5, 0.5, -0.5, -0.5)
        painter.drawRoundedRect(indicator_rect, 3, 3)

        mark_color = QColor(255, 255, 255)
        if not self.isEnabled():
            mark_color.setAlpha(130)
        mark_pen = QPen(mark_color, max(1.5, indicator_rect.height() * 0.13))
        mark_pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        mark_pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        painter.setPen(mark_pen)
        painter.setBrush(Qt.BrushStyle.NoBrush)
        left, top = indicator_rect.left(), indicator_rect.top()
        width, height = indicator_rect.width(), indicator_rect.height()
        if state == Qt.CheckState.Checked:
            path = QPainterPath(QPointF(left + width * 0.22, top + height * 0.52))
            path.lineTo(left + width * 0.43, top + height * 0.72)
            path.lineTo(left + width * 0.80, top + height * 0.30)
            painter.drawPath(path)
        elif state == Qt.CheckState.PartiallyChecked:
            painter.drawLine(
                QPointF(left + width * 0.25, top + height * 0.50),
                QPointF(left + width * 0.75, top + height * 0.50),
            )

        painter.setFont(self.font())
        style.drawItemText(
            painter,
            contents,
            Qt.AlignmentFlag.AlignLeft
            | Qt.AlignmentFlag.AlignVCenter
            | Qt.TextFlag.TextShowMnemonic
            | Qt.TextFlag.TextSingleLine,
            option.palette,
            self.isEnabled(),
            option.text,
            QPalette.ColorRole.WindowText,
        )


class SelectAllCheckBox(FilledCheckBox):
    """半选状态只用于展示；用户点击时切换为全选或全不选。"""

    def nextCheckState(self):
        state = (
            Qt.CheckState.Unchecked
            if self.checkState() == Qt.CheckState.Checked
            else Qt.CheckState.Checked
        )
        self.setCheckState(state)


class FilterCheckRow(QWidget):
    selection_changed = Signal()

    def __init__(self, object_prefix: str, title: str, items, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        # 相邻筛选项的间距大于 indicator 与自身文字的间距，避免视觉串组。
        layout.setSpacing(8)

        title_label = QLabel(title, self)
        title_label.setObjectName(f"{object_prefix}_filter_title")
        layout.addWidget(title_label)

        self.all_checkbox = SelectAllCheckBox("全选", self)
        self.all_checkbox.setObjectName(f"{object_prefix}_filter_all")
        self.all_checkbox.setTristate(True)
        self.all_checkbox.setCheckState(Qt.CheckState.Checked)
        layout.addWidget(self.all_checkbox)

        self.option_checkboxes = {}
        for label, value in items:
            checkbox = FilledCheckBox(label, self)
            checkbox.setObjectName(f"{object_prefix}_filter_{value}")
            checkbox.setChecked(True)
            self.option_checkboxes[value] = checkbox
            layout.addWidget(checkbox)
        layout.addStretch(1)

        self.all_checkbox.stateChanged.connect(self._set_all_options)
        for checkbox in self.option_checkboxes.values():
            checkbox.stateChanged.connect(self._sync_all_checkbox)

    def selected_values(self) -> frozenset[str]:
        return frozenset(
            value for value, checkbox in self.option_checkboxes.items()
            if checkbox.isChecked()
        )

    def _set_all_options(self, state):
        checked = getattr(state, "value", state) == Qt.CheckState.Checked.value
        blockers = [QSignalBlocker(box) for box in self.option_checkboxes.values()]
        for checkbox in self.option_checkboxes.values():
            checkbox.setChecked(checked)
        del blockers
        self.selection_changed.emit()

    def _sync_all_checkbox(self, _state=None):
        selected = len(self.selected_values())
        total = len(self.option_checkboxes)
        if selected == 0:
            state = Qt.CheckState.Unchecked
        elif selected == total:
            state = Qt.CheckState.Checked
        else:
            state = Qt.CheckState.PartiallyChecked
        blocker = QSignalBlocker(self.all_checkbox)
        self.all_checkbox.setCheckState(state)
        del blocker
        self.selection_changed.emit()


class SearchResultList(QListView):
    entry_requested = Signal(QModelIndex)

    def mouseDoubleClickEvent(self, event):
        index = self.indexAt(event.position().toPoint())
        super().mouseDoubleClickEvent(event)
        if index.isValid():
            self.entry_requested.emit(index)

    def keyPressEvent(self, event: QKeyEvent):
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            index = self.currentIndex()
            if index.isValid():
                self.entry_requested.emit(index)
            event.accept()
            return
        super().keyPressEvent(event)


class SearchResultDelegate(QStyledItemDelegate):
    """固定结果行高，并为已添加条目绘制状态徽标。"""

    _BADGE_HORIZONTAL_PADDING = 6
    _BADGE_HEIGHT = 18
    _CONTENT_MARGIN = 5
    _BADGE_GAP = 7

    def __init__(self, parent=None):
        super().__init__(parent)
        self.set_theme(False)

    def set_theme(self, dark: bool):
        if dark:
            self._badge_background = QColor(48, 209, 88, 58)
            self._badge_foreground = QColor(119, 235, 146)
        else:
            self._badge_background = QColor(52, 199, 89, 46)
            self._badge_foreground = QColor(24, 122, 63)

    def sizeHint(self, option, index):
        size = super().sizeHint(option, index)
        return QSize(size.width(), RESULT_ROW_HEIGHT)

    def paint(self, painter, option, index):
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        # 保留键盘焦点和回车操作，只去除原生的粗焦点框。
        opt.state &= ~QStyle.StateFlag.State_HasFocus

        if not bool(index.data(ADDED_ROLE)):
            super().paint(painter, opt, index)
            return

        label = opt.text
        opt.text = ""
        widget = opt.widget
        style = widget.style() if widget is not None else QApplication.style()
        style.drawControl(QStyle.ControlElement.CE_ItemViewItem, opt, painter, widget)

        content_rect = opt.rect.adjusted(
            self._CONTENT_MARGIN, 0, -self._CONTENT_MARGIN, 0
        )
        font_metrics = opt.fontMetrics
        badge_width = (
            font_metrics.horizontalAdvance(ADDED_BADGE_TEXT)
            + 2 * self._BADGE_HORIZONTAL_PADDING
        )
        badge_height = min(self._BADGE_HEIGHT, max(0, content_rect.height() - 4))
        badge_rect = QRect(
            content_rect.right() - badge_width + 1,
            content_rect.center().y() - badge_height // 2,
            badge_width,
            badge_height,
        )
        text_rect = QRect(
            content_rect.left(),
            content_rect.top(),
            max(0, badge_rect.left() - self._BADGE_GAP - content_rect.left()),
            content_rect.height(),
        )

        painter.save()
        painter.setFont(opt.font)
        painter.setPen(
            opt.palette.color(
                QPalette.ColorGroup.Disabled, QPalette.ColorRole.Text
            )
        )
        painter.drawText(
            text_rect,
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
            font_metrics.elidedText(
                label, Qt.TextElideMode.ElideRight, text_rect.width()
            ),
        )
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(self._badge_background)
        radius = badge_height / 2
        painter.drawRoundedRect(badge_rect, radius, radius)
        painter.setPen(self._badge_foreground)
        painter.drawText(badge_rect, Qt.AlignmentFlag.AlignCenter, ADDED_BADGE_TEXT)
        painter.restore()


class AddCodePanel(QFrame):
    """带多选筛选与分页的自选标的添加面板。"""

    entry_requested = Signal(object)

    def __init__(self, placeholder: str, parent=None):
        super().__init__(parent, Qt.WindowType.Popup)
        self.setObjectName("add_code_panel")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.setFixedWidth(430)

        self._search_index = ()
        self._existing_keys = set()
        self._page = 1
        self.current_result = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(7)

        self.search_input = QLineEdit(self)
        self.search_input.setObjectName("add_code_search_input")
        self.search_input.setPlaceholderText(placeholder)
        self.search_input.setToolTip(placeholder)
        layout.addWidget(self.search_input)

        self.category_filters = FilterCheckRow(
            "category", "类别", CATEGORY_FILTER_ITEMS, self
        )
        self.region_filters = FilterCheckRow(
            "region", "地区", REGION_FILTER_ITEMS, self
        )
        layout.addWidget(self.category_filters)
        layout.addWidget(self.region_filters)

        self.result_model = QStandardItemModel(self)
        self.result_list = SearchResultList(self)
        self.result_list.setObjectName("add_code_results")
        self.result_list.setModel(self.result_model)
        self.result_list.setEditTriggers(QListView.EditTrigger.NoEditTriggers)
        self.result_list.setSelectionMode(QListView.SelectionMode.SingleSelection)
        self.result_list.setUniformItemSizes(True)
        self.result_list.setSpacing(0)
        self.result_list.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.result_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.result_delegate = SearchResultDelegate(self.result_list)
        self.result_list.setItemDelegate(self.result_delegate)
        # 1px 上下边框 + 10 个固定高度条目，正好填满当前列表高度。
        self.result_list.setFixedHeight(PAGE_SIZE * RESULT_ROW_HEIGHT + 2)
        layout.addWidget(self.result_list)

        page_layout = QHBoxLayout()
        page_layout.setContentsMargins(0, 0, 0, 0)
        self.previous_button = QPushButton("上一页", self)
        self.previous_button.setObjectName("add_code_previous_page")
        self.previous_button.setAutoDefault(False)
        self.page_label = QLabel("0 / 0（共 0 条）", self)
        self.page_label.setObjectName("add_code_page_label")
        self.page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.next_button = QPushButton("下一页", self)
        self.next_button.setObjectName("add_code_next_page")
        self.next_button.setAutoDefault(False)
        page_layout.addWidget(self.previous_button)
        page_layout.addWidget(self.page_label, 1)
        page_layout.addWidget(self.next_button)
        layout.addLayout(page_layout)

        self.search_input.textChanged.connect(
            lambda _text: self.refresh_results(reset_page=True)
        )
        self.search_input.returnPressed.connect(self._activate_first_available)
        self.category_filters.selection_changed.connect(
            lambda: self.refresh_results(reset_page=True)
        )
        self.region_filters.selection_changed.connect(
            lambda: self.refresh_results(reset_page=True)
        )
        self.previous_button.clicked.connect(self._previous_page)
        self.next_button.clicked.connect(self._next_page)
        self.result_list.entry_requested.connect(self._activate_index)

        self.set_theme(False)
        self.refresh_results()

    def set_theme(self, dark: bool):
        if dark:
            background = "rgb(44, 44, 46)"
            foreground = "rgb(242, 242, 247)"
            border = "rgba(255, 255, 255, 0.28)"
            field = "rgb(58, 58, 60)"
            selected = "rgba(10, 132, 255, 0.42)"
            hovered = "rgba(255, 255, 255, 0.08)"
        else:
            background = "rgb(250, 250, 250)"
            foreground = "rgb(28, 28, 30)"
            border = "rgba(0, 0, 0, 0.28)"
            field = "rgb(255, 255, 255)"
            selected = "rgba(0, 122, 255, 0.20)"
            hovered = "rgba(0, 122, 255, 0.08)"
        self.result_delegate.set_theme(dark)
        for checkbox in self.findChildren(FilledCheckBox):
            checkbox.set_theme(dark)
        self.setStyleSheet(f"""
            QFrame#add_code_panel {{
                background-color: {background};
                color: {foreground};
                border: 1px solid {border};
                border-radius: 8px;
            }}
            QFrame#add_code_panel QLabel,
            QFrame#add_code_panel QCheckBox {{
                color: {foreground};
                border: none;
                background: transparent;
            }}
            QFrame#add_code_panel QListView {{
                color: {foreground};
                background-color: {field};
                border: 1px solid {border};
                border-radius: 5px;
                outline: none;
            }}
            QFrame#add_code_panel QListView::item {{
                padding: 0px 5px;
                border: none;
            }}
            QFrame#add_code_panel QListView::item:hover {{
                background-color: {hovered};
            }}
            QFrame#add_code_panel QListView::item:selected {{
                background-color: {selected};
                color: {foreground};
                border: none;
                outline: none;
            }}
        """)

    def set_context(self, search_index, existing_keys):
        self._search_index = tuple(search_index or ())
        self._existing_keys = {
            str(key or "").strip().casefold() for key in existing_keys or ()
        }
        self.refresh_results()

    def refresh_results(self, *, reset_page: bool = False):
        if reset_page:
            self._page = 1
        result = query_search_index(
            self._search_index,
            self.search_input.text(),
            categories=self.category_filters.selected_values(),
            regions=self.region_filters.selected_values(),
            page=self._page,
            page_size=PAGE_SIZE,
        )
        self.current_result = result
        self._page = result.page or 1
        self.result_model.clear()
        for entry in result.items:
            key = str(entry.get("key", "") or "").strip().casefold()
            added = key in self._existing_keys
            # “已添加”状态由 delegate 绘制为徽标，不再混入证券名称文本。
            item = QStandardItem(entry_display_text(entry))
            item.setEditable(False)
            item.setData(entry, ENTRY_ROLE)
            item.setData(added, ADDED_ROLE)
            if added:
                item.setEnabled(False)
                item.setSelectable(False)
            self.result_model.appendRow(item)

        self.page_label.setText(
            f"{result.page} / {result.page_count}（共 {result.total} 条）"
        )
        self.previous_button.setEnabled(result.page > 1)
        self.next_button.setEnabled(
            result.page_count > 0 and result.page < result.page_count
        )
        self._select_first_available()

    def show_for(self, anchor: QWidget):
        self.adjustSize()
        position = anchor.mapToGlobal(QPoint(0, anchor.height() + 2))
        screen = anchor.screen()
        available = screen.availableGeometry() if screen is not None else None
        if available is not None:
            max_x = max(available.left(), available.right() - self.width() + 1)
            x = min(max(position.x(), available.left()), max_x)
            below_y = position.y()
            above_y = anchor.mapToGlobal(QPoint(0, -self.height() - 2)).y()
            y = below_y if below_y + self.height() <= available.bottom() + 1 else above_y
            max_y = max(available.top(), available.bottom() - self.height() + 1)
            y = min(max(y, available.top()), max_y)
            position = QPoint(x, y)
        self.move(position)
        self.show()
        self.raise_()
        QTimer.singleShot(0, self._focus_search_input)

    def _focus_search_input(self):
        if self.isVisible():
            self.search_input.setFocus(Qt.FocusReason.PopupFocusReason)
            self.search_input.selectAll()

    def _previous_page(self):
        if self.current_result is not None and self.current_result.page > 1:
            self._page = self.current_result.page - 1
            self.refresh_results()

    def _next_page(self):
        if (self.current_result is not None
                and self.current_result.page < self.current_result.page_count):
            self._page = self.current_result.page + 1
            self.refresh_results()

    def _first_available_index(self) -> QModelIndex:
        for row in range(self.result_model.rowCount()):
            index = self.result_model.index(row, 0)
            if (index.flags() & Qt.ItemFlag.ItemIsEnabled
                    and index.flags() & Qt.ItemFlag.ItemIsSelectable):
                return index
        return QModelIndex()

    def _select_first_available(self):
        index = self._first_available_index()
        if index.isValid():
            self.result_list.setCurrentIndex(index)
        else:
            self.result_list.setCurrentIndex(QModelIndex())

    def _activate_first_available(self):
        index = self.result_list.currentIndex()
        if not index.isValid():
            index = self._first_available_index()
        self._activate_index(index)

    def _activate_index(self, index: QModelIndex):
        if not index.isValid():
            return
        flags = index.flags()
        if not (flags & Qt.ItemFlag.ItemIsEnabled
                and flags & Qt.ItemFlag.ItemIsSelectable):
            return
        if bool(index.data(ADDED_ROLE)):
            return
        entry = index.data(ENTRY_ROLE)
        if isinstance(entry, dict):
            self.entry_requested.emit(entry)

    def keyPressEvent(self, event: QKeyEvent):
        if event.key() == Qt.Key.Key_Escape:
            self.hide()
            event.accept()
            return
        super().keyPressEvent(event)
