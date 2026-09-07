"""名称指标的显示设置悬浮面板。"""

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QSizePolicy,
    QVBoxLayout,
)

from stockwidget.ui.metric_settings_panel import MetricSettingsPanel

# 名称显示字数选项: 0=不显示, -1=全部显示, 1-4=前 N 个字
NAME_LENGTH_OPTIONS = (
    (0, "不显示"),
    (-1, "全部显示"),
    (1, "1个字"),
    (2, "2个字"),
    (3, "3个字"),
    (4, "4个字"),
)


class NameSettingsPanel(MetricSettingsPanel):
    """点击“名称”后的 ⓘ时弹出的设置面板（显示字数/代码/类型）。"""

    name_length_changed = Signal(int)
    code_visible_changed = Signal(bool)
    type_visible_changed = Signal(bool)
    # 面板关闭（含点击外部空白关闭）时发出，用于清除指标块的选中状态
    panel_closed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("name_settings_panel")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(7)

        length_row = QHBoxLayout()
        length_row.setContentsMargins(0, 0, 0, 0)
        length_row.setSpacing(6)
        self.length_label = QLabel("显示字数：", self)
        self.cmb_namelen = QComboBox(self)
        self.cmb_namelen.setObjectName("name_settings_namelen")
        for value, text in NAME_LENGTH_OPTIONS:
            self.cmb_namelen.addItem(text, userData=value)
        length_row.addWidget(self.length_label)
        length_row.addWidget(self.cmb_namelen, 1)
        layout.addLayout(length_row)

        flag_row = QHBoxLayout()
        flag_row.setContentsMargins(0, 0, 0, 0)
        flag_row.setSpacing(10)
        self.cb_code = QCheckBox("显示代码", self)
        self.cb_type = QCheckBox("显示类型", self)
        self.cb_code.setObjectName("name_settings_code")
        self.cb_type.setObjectName("name_settings_type")
        flag_row.addWidget(self.cb_code)
        flag_row.addWidget(self.cb_type)
        flag_row.addStretch(1)
        layout.addLayout(flag_row)

        self.cmb_namelen.currentIndexChanged.connect(self._on_name_length_changed)
        self.cb_code.toggled.connect(self.code_visible_changed)
        self.cb_type.toggled.connect(self.type_visible_changed)

        self.set_theme(False)
        self.sync_from(None)

    def sync_from(self, win):
        """按浮窗当前状态同步面板控件；win 为 None 时使用默认值。"""
        name_length = getattr(win, "name_length", -1)
        code_visible = bool(getattr(win, "code_visible", False))
        type_visible = bool(getattr(win, "type_visible", False))

        idx = self.cmb_namelen.findData(name_length)
        self.cmb_namelen.blockSignals(True)
        self.cmb_namelen.setCurrentIndex(idx if idx >= 0 else 1)
        self.cmb_namelen.blockSignals(False)

        self.cb_code.blockSignals(True)
        self.cb_code.setChecked(code_visible)
        self.cb_code.blockSignals(False)

        self.cb_type.blockSignals(True)
        self.cb_type.setChecked(type_visible)
        self.cb_type.blockSignals(False)

    def _on_name_length_changed(self, idx: int):
        value = self.cmb_namelen.itemData(idx)
        if isinstance(value, int):
            self.name_length_changed.emit(value)

    def set_theme(self, dark: bool):
        if dark:
            background = "rgb(44, 44, 46)"
            border = "rgba(255, 255, 255, 0.28)"
        else:
            background = "rgb(250, 250, 250)"
            border = "rgba(0, 0, 0, 0.28)"
        self.setStyleSheet(f"""
            QFrame#name_settings_panel {{
                background-color: {background};
                border: 1px solid {border};
                border-radius: 8px;
            }}
        """)

    def hideEvent(self, event):
        super().hideEvent(event)
        self.panel_closed.emit()
