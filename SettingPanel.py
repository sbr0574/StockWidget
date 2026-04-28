import os, re
from functools import partial

from PySide6.QtCore import Qt, QSize, QStringListModel
from PySide6.QtGui import QColor, QFontDatabase, QKeySequence
from PySide6.QtWidgets import (
    QWidget, QDialog, QVBoxLayout, QHBoxLayout, QGridLayout, QTabWidget, QPushButton, QSlider,
    QGroupBox, QLabel, QColorDialog, QComboBox, QAbstractItemView,
    QCheckBox, QListWidget, QListWidgetItem, QKeySequenceEdit, QFileDialog, QStyledItemDelegate,
    QLineEdit, QCompleter
)
from WidgetPanel import FloatLabel
from code_index import find_suggestions


class CodeCompleterDelegate(QStyledItemDelegate):
    def __init__(self, owner):
        super().__init__(owner)
        self.owner = owner

    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        completer = QCompleter(self.owner.suggestion_model, editor)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        completer.activated.connect(lambda text: self.owner._apply_suggestion(editor, text))
        editor.textEdited.connect(self.owner._update_suggestions)
        editor.setCompleter(completer)
        return editor

class SettingsDialog(QDialog):
    def __init__(self, win: FloatLabel, parent: QWidget, app=None):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.win = win
        self.app = app
        self.setModal(False)
        self.code_index = list(getattr(app, "code_index", []) or [])
        self.suggestion_model = QStringListModel(self)
        self._suggestion_map = {}

        main = QHBoxLayout(self)
        main.setContentsMargins(8, 8, 8, 8)
        main.setSpacing(8)
        self.tabs = QTabWidget()
        main.addWidget(self.tabs)

        self.tab_sizes = {
            0: QSize(300, 300),
            1: QSize(440, 420),
            2: QSize(360, 350),
            3: QSize(300, 220),
        }
        self._apply_tab_size(0)

        # ---- 第一页 ----
        tab_0 = QWidget()
        code_settings = QVBoxLayout(tab_0)

        # 1.自选列表
        g_codes = QGroupBox("自选列表")
        g_codes.setContentsMargins(3,12,3,6)
        lay_codes = QHBoxLayout(g_codes)
        lay_codes.setSpacing(6)
        # 1.1 代码列表
        self.list_codes = QListWidget()
        self.list_codes.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked | QAbstractItemView.EditKeyPressed)
        self.list_codes.setFixedWidth(220)
        self.list_codes.setItemDelegate(CodeCompleterDelegate(self))
        for c in self.win.codes:
            it = QListWidgetItem(c)
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsEditable | Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            it.setCheckState(Qt.Checked if c in getattr(self.win, 'checked_codes', []) else Qt.Unchecked)
            it.setData(Qt.UserRole, c)  # 记住上次有效值
            self.list_codes.addItem(it)
        # 1.2 操作按钮
        btn_col = QVBoxLayout()
        btn_col.setSpacing(4)
        self.btn_add = QPushButton("添加")
        self.btn_add.setFixedWidth(60)
        self.btn_del = QPushButton("删除")
        self.btn_del.setFixedWidth(60)
        self.btn_up  = QPushButton("上移")
        self.btn_up.setFixedWidth(60)
        self.btn_dn  = QPushButton("下移")
        self.btn_dn.setFixedWidth(60)
        for b in (self.btn_add, self.btn_del, self.btn_up, self.btn_dn):
            btn_col.addWidget(b)
        btn_col.addStretch(1)

        lay_codes.addWidget(self.list_codes, 1)
        lay_codes.addLayout(btn_col)
        code_settings.addWidget(g_codes)

        self.tabs.addTab(tab_0, "自选列表")

        # ---- 第二页 ----
        tab_1 = QWidget()
        data_settings = QVBoxLayout(tab_1)

        # 2.刷新间隔
        g_interval = QGroupBox("刷新间隔")
        g_interval.setContentsMargins(3,12,3,6)
        self.cmb_interval = QComboBox()
        self.cmb_interval.setFixedWidth(136)
        for s in [1,2,3,5,10,15,30,60]:
            self.cmb_interval.addItem(f"{s} 秒", userData=s)
        idx = self.cmb_interval.findData(self.win.refresh_seconds)
        self.cmb_interval.setCurrentIndex(idx if idx >= 0 else 1)
        v = QVBoxLayout(g_interval)
        v.setContentsMargins(6,6,6,6)
        v.addWidget(self.cmb_interval)
        data_settings.addWidget(g_interval)

        # 3.显示选项
        # 3.1复选框组
        g_flags = QGroupBox("显示指标")
        g_flags.setContentsMargins(3,12,3,6)
        gl_flags = QGridLayout(g_flags)
        self.cbs: list[QCheckBox] = []
        cb_texts = self.win.ALL_HEADERS

        g_flag_name = QGroupBox("名称")
        gl_flag_name = QGridLayout(g_flag_name)
        gl_flag_name.setHorizontalSpacing(6)
        gl_flag_name.setVerticalSpacing(6)
        # 代码、名称
        for i, h in enumerate(cb_texts[0:2]):
            cb = QCheckBox(h)
            cb.setChecked(self.win.header_is_visible(h))
            cb.stateChanged.connect(partial(self._on_cb_changed, h))
            self.cbs.append(cb)
            gl_flag_name.addWidget(cb, i, 0)
        self.cb_short_code = QCheckBox("仅显示数字")
        self.cb_short_code.setChecked(bool(self.win.short_code))
        self.cb_short_code.setEnabled(self.win.header_is_visible("代码"))
        gl_flag_name.addWidget(self.cb_short_code, 0, 1)
        self.cmb_namelength = QComboBox()
        self.cmb_namelength.setFixedWidth(80)
        for l in [0, 1, 2, 3, 4]:
            self.cmb_namelength.addItem(f"{l}个字" if l>0 else "完整", userData=l)
        idx_name = self.cmb_namelength.findData(self.win.name_length)
        self.cmb_namelength.setCurrentIndex(idx_name if idx_name>=0 else 1)
        self.cmb_namelength.setEnabled(self.win.header_is_visible("名称"))
        gl_flag_name.addWidget(self.cmb_namelength, 1, 1)
        gl_flags.addWidget(g_flag_name, 0, 0)

        g_flag_price = QGroupBox("价格")
        gl_flag_price = QGridLayout(g_flag_price)
        gl_flag_price.setHorizontalSpacing(6)
        gl_flag_price.setVerticalSpacing(6)
        # 现价、涨跌值、涨跌幅
        for i, h in enumerate(cb_texts[2:5]):
            cb = QCheckBox(h)
            cb.setChecked(self.win.header_is_visible(h))
            cb.stateChanged.connect(partial(self._on_cb_changed, h))
            self.cbs.append(cb)
            gl_flag_price.addWidget(cb, i, 0)
        gl_flags.addWidget(g_flag_price, 1, 0)

        g_flag_order = QGroupBox("盘口")
        gl_flag_order = QGridLayout(g_flag_order)
        gl_flag_order.setHorizontalSpacing(6)
        gl_flag_order.setVerticalSpacing(6)
        # 买一/卖一
        self.cb_b1s1 = QCheckBox("买一/卖一")
        self.cb_b1s1.setChecked(self.win.b1s1_visible)
        self.cb_b1s1.stateChanged.connect(self._on_b1s1_toggled)
        self.cbs.append(self.cb_b1s1)
        gl_flag_order.addWidget(self.cb_b1s1, 0, 0)
        
        # 委比
        cb_commi = QCheckBox("委比")
        cb_commi.setChecked(self.win.header_is_visible("委比"))
        cb_commi.stateChanged.connect(partial(self._on_cb_changed, "委比"))
        self.cbs.append(cb_commi)
        gl_flag_order.addWidget(cb_commi, 1, 0)
        
        # 买一/卖一显示模式：数量 / 价格 / 数量和价格
        self.cmb_b1s1_display = QComboBox()
        self.cmb_b1s1_display.setFixedWidth(100)
        self.cmb_b1s1_display.addItem("数量", userData="qty")
        self.cmb_b1s1_display.addItem("价格", userData="price")
        self.cmb_b1s1_display.addItem("数量和价格", userData="both")
        cur_mode = getattr(self.win, 'b1s1_display', 'qty')
        idx_mode = self.cmb_b1s1_display.findData(cur_mode)
        self.cmb_b1s1_display.setCurrentIndex(idx_mode if idx_mode>=0 else 0)
        self.cmb_b1s1_display.setEnabled(self.win.b1s1_visible)
        gl_flag_order.addWidget(self.cmb_b1s1_display, 0, 1)
        gl_flags.addWidget(g_flag_order, 0, 1)

        g_flag_deal = QGroupBox("成交")
        gl_flag_deal = QGridLayout(g_flag_deal)
        gl_flag_deal.setHorizontalSpacing(6)
        gl_flag_deal.setVerticalSpacing(6)
        for i in range(8,11):
            cb = QCheckBox(cb_texts[i])
            cb.setChecked(self.win.header_is_visible(cb_texts[i]))
            cb.stateChanged.connect(partial(self._on_cb_changed, cb_texts[i]))
            self.cbs.append(cb)
            gl_flag_deal.addWidget(cb, i-8, 0)
        gl_flags.addWidget(g_flag_deal, 1, 1)

        g_flag_other = QGroupBox("其他")
        gl_flag_other = QGridLayout(g_flag_other)
        gl_flag_other.setHorizontalSpacing(6)
        gl_flag_other.setVerticalSpacing(6)
        for i in range(11,12):
            cb = QCheckBox(cb_texts[i])
            cb.setChecked(self.win.header_is_visible(cb_texts[i]))
            cb.stateChanged.connect(partial(self._on_cb_changed, cb_texts[i]))
            self.cbs.append(cb)
            gl_flag_other.addWidget(cb, i-11, 0)
        gl_flags.addWidget(g_flag_other, 2, 0)

        data_settings.addWidget(g_flags)

        self.tabs.addTab(tab_1, "显示数据")

        # ---- 第三页 ----
        tab_2 = QWidget()
        appearance_settings = QVBoxLayout(tab_2)

        # 表格外观
        g_table = QGroupBox("表格外观")
        g_table.setContentsMargins(3,12,3,6)
        gl_table = QGridLayout(g_table)
        gl_table.setHorizontalSpacing(6)
        gl_table.setVerticalSpacing(6)
        # 复选框
        self.chk_table_header = QCheckBox("显示表头")
        self.chk_table_header.setChecked(self.win.header_visible)
        self.chk_table_grid = QCheckBox("显示网格")
        self.chk_table_grid.setChecked(self.win.grid_visible)

        gl_table.addWidget(self.chk_table_header,0,0)
        gl_table.addWidget(self.chk_table_grid,0,1)
        appearance_settings.addWidget(g_table)

        # 3.颜色/透明度
        g_color = QGroupBox("颜色与透明度")
        g_color.setContentsMargins(3,12,3,6)
        gl_color = QGridLayout(g_color)
        gl_color.setHorizontalSpacing(6)
        gl_color.setVerticalSpacing(6)
        # 3.1 复选框：默认颜色
        self.chk_default_color = QCheckBox("默认颜色")
        self.chk_default_color.setChecked(self.win.default_color)
        # 3.2 按钮：文字颜色
        self.btn_fg = QPushButton("文字颜色…")
        self.btn_fg.setFixedWidth(90)
        self.btn_fg.setEnabled(not self.win.default_color)
        # 3.3 按钮：背景颜色
        self.btn_bg = QPushButton("背景颜色…")
        self.btn_bg.setFixedWidth(90)
        # 3.4 滑块：背景不透明度
        self.slider_bg_alpha = QSlider(Qt.Horizontal)
        self.slider_bg_alpha.setRange(1, 100)
        self.slider_bg_alpha.setMinimumWidth(150)
        self.slider_bg_alpha.setValue(int(round(self.win.bg.alpha()/2.55)))
        self.lbl_bg_alpha = QLabel(f"{self.slider_bg_alpha.value()}%")
        # 3.5 滑块：整体不透明度
        self.slider_win_opacity = QSlider(Qt.Horizontal)
        self.slider_win_opacity.setRange(20, 100)
        self.slider_win_opacity.setMinimumWidth(150)
        self.slider_win_opacity.setValue(int(round(self.win.windowOpacity()*100)))
        self.lbl_win_opacity = QLabel(f"{self.slider_win_opacity.value()}%")

        gl_color.addWidget(self.chk_default_color,0,0,1,2)
        gl_color.addWidget(self.btn_fg,0,2,1,2)
        gl_color.addWidget(self.btn_bg,0,4,1,2)
        gl_color.addWidget(QLabel("背景不透明度："),1,0,1,2)
        gl_color.addWidget(self.slider_bg_alpha,1,2,1,3)
        gl_color.addWidget(self.lbl_bg_alpha,1,5,1,1)
        gl_color.addWidget(QLabel("整体不透明度："),2,0,1,2)
        gl_color.addWidget(self.slider_win_opacity,2,2,1,3)
        gl_color.addWidget(self.lbl_win_opacity,2,5,1,1)
        appearance_settings.addWidget(g_color)

        # 4.字体/行距
        g_font = QGroupBox("字体与行距")
        g_font.setContentsMargins(3,12,3,6)
        gl_font = QGridLayout(g_font)
        gl_font.setHorizontalSpacing(6)
        gl_font.setVerticalSpacing(6)
        # 4.1 选项：字体
        self.cmb_family = QComboBox()
        self.cmb_family.setFixedWidth(200)
        for fam in sorted(QFontDatabase.families()):
            self.cmb_family.addItem(fam)
        fi = self.cmb_family.findText(self.win.font.family())
        self.cmb_family.setCurrentIndex(fi if fi >= 0 else 0)
        # 4.2 滑块：字号
        self.slider_font = QSlider(Qt.Horizontal)
        self.slider_font.setRange(8, 15)
        self.slider_font.setMinimumWidth(150)
        self.slider_font.setValue(self.win.font.pointSize())
        self.lbl_font = QLabel(f"{self.slider_font.value()} pt")
        # 4.3 滑块：行间距
        self.slider_line = QSlider(Qt.Horizontal)
        self.slider_line.setRange(0, 20)
        self.slider_line.setMinimumWidth(150)
        self.slider_line.setValue(getattr(self.win,"line_extra_px",4))
        self.lbl_line = QLabel(f"+{self.slider_line.value()} px")

        gl_font.addWidget(QLabel("字体："),0,0,1,2)
        gl_font.addWidget(self.cmb_family,0,2,1,4)
        gl_font.addWidget(QLabel("字号："),1,0,1,2)
        gl_font.addWidget(self.slider_font,1,2,1,3)
        gl_font.addWidget(self.lbl_font,1,5,1,1)
        gl_font.addWidget(QLabel("行距："),2,0,1,2)
        gl_font.addWidget(self.slider_line,2,2,1,3)
        gl_font.addWidget(self.lbl_line,2,5,1,1)
        appearance_settings.addWidget(g_font)

        self.tabs.addTab(tab_2, "外观")

        # ---- 第四页 ----
        tab_3 = QWidget()
        other_settings = QVBoxLayout(tab_3)

        # 4.热键
        g_hotkey = QGroupBox("快捷键")
        g_hotkey.setContentsMargins(3,12,3,6)
        gl_hotkey = QGridLayout(g_hotkey)
        gl_hotkey.setHorizontalSpacing(6)
        gl_hotkey.setVerticalSpacing(6)
        gl_hotkey.addWidget(QLabel("隐藏/显示浮窗："),0,0,1,1)
        self.edit_hotkey = QKeySequenceEdit()
        self.edit_hotkey.setKeySequence(QKeySequence(self.win.hotkey))
        gl_hotkey.addWidget(self.edit_hotkey,0,1)
        # 开机启动复选框
        self.chk_start_on_boot = QCheckBox("开机启动")
        self.chk_start_on_boot.setChecked(bool(self.win.start_on_boot))
        other_settings.addWidget(self.chk_start_on_boot)
        other_settings.addWidget(g_hotkey)

        # 程序图标选择
        g_icon = QGroupBox("程序图标")
        g_icon.setContentsMargins(3,12,3,6)
        gl_icon = QHBoxLayout(g_icon)
        self.cmb_icon = QComboBox()
        icon_items = [
            ("默认", 'default'),
            ("系统：计算机", 'std:computer'),
            ("系统：网络", 'std:network'),
            ("系统：文件夹", 'std:folder'),
            ("系统：文件", 'std:file'),
            ("系统：回收站", 'std:trash'),
        ]
        for label, val in icon_items:
            self.cmb_icon.addItem(label, userData=val)
        self.btn_pick_icon = QPushButton("自定义图标…")
        self.btn_pick_icon.setFixedWidth(120)
        gl_icon.addWidget(self.cmb_icon)
        gl_icon.addWidget(self.btn_pick_icon)
        other_settings.addWidget(g_icon)

        self.tabs.addTab(tab_3, "常规")

        # ---- 连接 ----
        # 连接：代码列表
        self.list_codes.itemChanged.connect(self._on_codes_changed)
        self.btn_add.clicked.connect(self._add_code)
        self.btn_del.clicked.connect(self._del_code)
        self.btn_up.clicked.connect(self._move_up)
        self.btn_dn.clicked.connect(self._move_down)
        # 连接：其它设置
        self.cmb_interval.currentIndexChanged.connect(self._on_interval_changed)
        self.cmb_namelength.currentIndexChanged.connect(self._on_name_length_changed)
        self.chk_default_color.toggled.connect(self._on_default_color_toggled)
        self.btn_fg.clicked.connect(self.pick_fg)
        self.btn_bg.clicked.connect(self.pick_bg)
        self.slider_bg_alpha.valueChanged.connect(self.apply_bg_alpha)
        self.slider_win_opacity.valueChanged.connect(self.apply_win_opacity)
        self.cmb_family.currentTextChanged.connect(self._on_family_changed)
        self.slider_font.valueChanged.connect(self.apply_font_size)
        self.slider_line.valueChanged.connect(self._on_line_changed)
        self.edit_hotkey.editingFinished.connect(self._on_hotkey_changed)
        self.chk_start_on_boot.toggled.connect(self._on_start_on_boot_toggled)
        self.chk_table_header.toggled.connect(self._on_header_toggled)
        self.chk_table_grid.toggled.connect(self._on_grid_toggled)
        # icon controls
        try:
            # set current index based on app config if available
            cur_choice = None
            if hasattr(self, 'app') and self.app is not None:
                cur_choice = getattr(self.app, '_app_icon_choice', None)
            if cur_choice is None:
                cur_choice = 'default'
            # find index
            idx = self.cmb_icon.findData(cur_choice)
            if idx < 0:
                if isinstance(cur_choice, str) and os.path.exists(cur_choice):
                    self.cmb_icon.addItem('自定义', userData=cur_choice)
                    idx = self.cmb_icon.count()-1
            self.cmb_icon.setCurrentIndex(idx if idx >= 0 else 0)
        except Exception:
            pass
        self.cmb_icon.currentIndexChanged.connect(self._on_icon_changed)
        self.btn_pick_icon.clicked.connect(self._pick_custom_icon)
        self.tabs.currentChanged.connect(self._apply_tab_size)
        self.cmb_b1s1_display.currentIndexChanged.connect(self._on_b1s1_display_changed)
        self.cb_short_code.stateChanged.connect(self._on_short_code_toggled)

    def _on_start_on_boot_toggled(self, checked: bool):
        try:
            self.win.set_start_on_boot(bool(checked))
            if hasattr(self, 'app') and self.app is not None:
                try:
                    self.app.set_start_on_boot(bool(checked))
                except Exception:
                    pass
        except Exception:
            pass

    _re_full = re.compile(r'^(sh|sz|bj)\d{6}$')
    _re_6 = re.compile(r'^\d{6}$')

    def _to_prefixed_code(self, text: str) -> str | None:
        s = str(text or "").strip().lower()
        if not s:
            return None
        if self._re_full.match(s):
            return s
        if self._re_6.match(s):
            if s[0] in ("6", "5", "9"):
                return f"sh{s}"
            if s[0] in ("0", "1", "2", "3"):
                return f"sz{s}"
            if s[0] in ("4", "8"):
                return f"bj{s}"
        return None

    def _display_text(self, item: dict) -> str:
        return f"{item.get('code', '')} {item.get('name', '')}".strip()

    def _apply_suggestion(self, editor: QLineEdit, text: str):
        code = self._suggestion_map.get(text)
        if code:
            editor.setText(code)
            editor.selectAll()

    def _update_suggestions(self, text: str):
        query = str(text or "").strip()
        candidates = find_suggestions(self.code_index, query, limit=20)
        labels = []
        self._suggestion_map = {}
        for item in candidates:
            label = self._display_text(item)
            labels.append(label)
            self._suggestion_map[label] = item.get("code", "")
        self.suggestion_model.setStringList(labels)

    def _code_from_input(self, text: str) -> str | None:
        s = str(text or "").strip()
        if not s:
            return None
        if s in self._suggestion_map:
            return self._suggestion_map[s]
        token = s.split()[0]
        return self._to_prefixed_code(token)

    def _collect_codes_from_list(self):
        codes = []
        seen = set()
        for i in range(self.list_codes.count()):
            txt = self.list_codes.item(i).text()
            code = self._code_from_input(txt)
            if code:
                if code not in seen:
                    seen.add(code)
                    codes.append(code)
                it = self.list_codes.item(i)
                if it.text() != code:
                    self.list_codes.blockSignals(True)
                    it.setText(code)
                    it.setData(Qt.UserRole, code)
                    self.list_codes.blockSignals(False)
            else:
                it = self.list_codes.item(i)
                prev = it.data(Qt.UserRole)
                if prev:
                    self.list_codes.blockSignals(True)
                    it.setText(prev)
                    self.list_codes.blockSignals(False)
                else:
                    self.list_codes.takeItem(i)
                    return self._collect_codes_from_list()
        return codes

    def _on_codes_changed(self, _item):
        codes = self._collect_codes_from_list()
        self.win.set_codes(codes)
        checked_codes = [
            self.list_codes.item(i).text().split()[0]
            for i in range(self.list_codes.count())
            if self.list_codes.item(i).checkState() == Qt.Checked
        ]
        self.win.set_checked_codes(checked_codes)

    def _add_code(self):
        it = QListWidgetItem("sh000001")
        it.setFlags(it.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsEditable | Qt.ItemIsSelectable | Qt.ItemIsEnabled)
        it.setCheckState(Qt.Unchecked)
        it.setData(Qt.UserRole, "sh000001")
        self.list_codes.addItem(it)
        self.list_codes.setCurrentItem(it)
        self.list_codes.editItem(it)
        self._on_codes_changed(it)

    def _del_code(self):
        row = self.list_codes.currentRow()
        if row >= 0:
            self.list_codes.takeItem(row)
            self._on_codes_changed(None)

    def _move_up(self):
        row = self.list_codes.currentRow()
        if row > 0:
            it = self.list_codes.takeItem(row)
            self.list_codes.insertItem(row-1, it)
            self.list_codes.setCurrentRow(row-1)
            self._on_codes_changed(None)

    def _move_down(self):
        row = self.list_codes.currentRow()
        if 0 <= row < self.list_codes.count()-1:
            it = self.list_codes.takeItem(row)
            self.list_codes.insertItem(row+1, it)
            self.list_codes.setCurrentRow(row+1)
            self._on_codes_changed(None)

    # —— 其它槽 —— #
    def _on_interval_changed(self, idx):
        seconds = self.cmb_interval.currentData()
        if isinstance(seconds,int): 
            self.win.set_refresh_interval(seconds)

    def _on_default_color_toggled(self, checked: bool):
        self.btn_fg.setEnabled(not checked)
        self.win.set_default_color(bool(checked))
    
    def _on_grid_toggled(self, checked: bool):
        self.win.set_grid_visible(bool(checked))

    def _on_header_toggled(self, checked: bool):
        self.win.set_header_visible(bool(checked))

    def _on_cb_changed(self, header: str, state: bool):
        self.win.set_flag(header, state)
        if header == "代码":
            self.cb_short_code.setEnabled(state)
        elif header == "名称":
            self.cmb_namelength.setEnabled(state)
    
    def _on_short_code_toggled(self, checked: bool):
        self.win.set_code_type(checked)

    def _on_name_length_changed(self, length: int):
        self.win.set_name_length(length)

    def _on_b1s1_display_changed(self, idx: int):
        try:
            val = self.cmb_b1s1_display.itemData(idx)
            if not val:
                return
            self.win.set_b1s1_display(val)
        except Exception:
            pass

    def _on_b1s1_toggled(self, state: bool):
        self.win.set_flag("买一", state)
        self.cmb_b1s1_display.setEnabled(state)

    def _apply_tab_size(self, index: int):
        size = self.tab_sizes.get(index, QSize(400, 400))
        self.setFixedSize(size)

    def pick_fg(self):
        c = QColorDialog.getColor(self.win.fg, self, "选择文字颜色")
        if c.isValid(): self.win.set_fg_color(c)
    def pick_bg(self):
        base = QColor(self.win.bg)
        base.setAlpha(255)
        c = QColorDialog.getColor(base, self, "选择背景颜色")
        if c.isValid(): self.win.set_bg_rgb_keep_alpha(c)
    def apply_bg_alpha(self, v): 
        self.lbl_bg_alpha.setText(f"{v}%")
        self.win.set_bg_alpha_percent(v)
    def apply_win_opacity(self, v): 
        self.lbl_win_opacity.setText(f"{v}%")
        self.win.set_window_opacity_percent(v)
    def _on_family_changed(self, fam: str): 
        self.win.set_font_family(fam)
    def apply_font_size(self, v):
        self.lbl_font.setText(f"{v} pt")
        self.win.set_font_size(v)  # 同步 K 线缩放
    def _on_line_changed(self, v: int): 
        self.lbl_line.setText(f"+{v} px")
        self.win.set_line_extra(v)
    def _on_hotkey_changed(self):
        new_hotkey = self.edit_hotkey.keySequence().toString()
        try:
            self.win.update_hotkey(new_hotkey)
        except Exception:
            pass

    def _on_icon_changed(self, idx: int):
        try:
            val = self.cmb_icon.itemData(idx)
            if not val:
                return
            if hasattr(self, 'app') and self.app is not None:
                try:
                    self.app.set_app_icon(val)
                    # persist immediately
                    try:
                        self.app.save_now()
                    except Exception:
                        pass
                except Exception:
                    pass
        except Exception:
            pass

    def _pick_custom_icon(self):
        try:
            path, _ = QFileDialog.getOpenFileName(self, "选择图标文件", os.path.expanduser('~'), "图标文件 (*.ico);;All Files (*)")
            if path:
                # append or find existing custom entry
                idx = self.cmb_icon.findData(path)
                if idx < 0:
                    self.cmb_icon.addItem('自定义', userData=path)
                    idx = self.cmb_icon.count()-1
                self.cmb_icon.setCurrentIndex(idx)
                # trigger change handler will call app.set_app_icon
        except Exception:
            pass
