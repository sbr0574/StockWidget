"""自选标的添加悬浮面板。"""

from PySide6.QtCore import QModelIndex, QPoint, QSignalBlocker, Qt, QTimer, Signal
from PySide6.QtGui import QKeyEvent, QStandardItem, QStandardItemModel
from PySide6.QtWidgets import (
    QCheckBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListView,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from stockwidget.core.code_search import query_search_index


ENTRY_ROLE = Qt.ItemDataRole.UserRole + 1
ADDED_ROLE = Qt.ItemDataRole.UserRole + 2
PAGE_SIZE = 10

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


class SelectAllCheckBox(QCheckBox):
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
        layout.setSpacing(6)

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
            checkbox = QCheckBox(label, self)
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
        self.result_list.setFixedHeight(242)
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
        else:
            background = "rgb(250, 250, 250)"
            foreground = "rgb(28, 28, 30)"
            border = "rgba(0, 0, 0, 0.28)"
            field = "rgb(255, 255, 255)"
            selected = "rgba(0, 122, 255, 0.20)"
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
            QFrame#add_code_panel QLineEdit,
            QFrame#add_code_panel QListView {{
                color: {foreground};
                background-color: {field};
                border: 1px solid {border};
                border-radius: 5px;
            }}
            QFrame#add_code_panel QListView::item {{ padding: 3px 5px; }}
            QFrame#add_code_panel QListView::item:selected {{
                background-color: {selected};
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
            item = QStandardItem(entry_display_text(entry, added=added))
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
