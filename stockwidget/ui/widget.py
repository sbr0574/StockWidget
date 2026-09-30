from functools import partial
import requests
import threading
import sys

from PySide6.QtCore import QPoint, Qt, QTimer, Signal
from PySide6.QtGui import QFont, QAction, QColor
from PySide6.QtWidgets import QApplication, QWidget, QMenu, QHeaderView, QStyledItemDelegate

from stockwidget.ui.table_model import (
    DEFAULT_DOWN_COLOR,
    DEFAULT_NEUTRAL_COLOR,
    DEFAULT_UP_COLOR,
    KLineDelegate,
    SimpleTableModel,
)
from stockwidget.ui.drag_mixin import DragBehaviorMixin
from stockwidget.ui.hide_controller import HideController
from stockwidget.ui.position_controller import PositionController
from stockwidget.core.hide_rules import normalize_hide_times
from stockwidget.core.view_options import DISPLAY_MODES, TASKBAR_STYLE_KEYS, ViewOptions, column_ranges, page_slice
from stockwidget.ui.table_header import SortIndicatorStyle
from stockwidget.platform.hotkeys import GlobalHotkeyManager, HotkeyResult
from stockwidget.data.quotes import request_quote
from stockwidget.data.network_errors import request_error_message
from stockwidget.core.quote_presentation import QuoteDisplayOptions, format_quote
from stockwidget.core.metric_layout import (
    METRIC_BY_ID,
    METRIC_SPECS,
    NAME_METRIC_ID,
    expand_metric_headers,
    legacy_visibility,
    normalize_visible_metrics,
    visible_metrics_from_config,
)
from stockwidget.core.watchlist import normalize_watchlist
from stockwidget.core.geometry import resolve_restore_position
from stockwidget.platform.capabilities import (
    is_wayland,
    hotkeys_supported, click_through_supported,
    opacity_supported, force_top_supported,
    default_font_family,
)
from stockwidget.platform.click_through import apply_click_through
from stockwidget.platform.windows import ensure_topmost


SORTABLE_HEADERS = (
    "现价", "涨跌", "涨幅", "浮盈", "委比", "成交量", "成交额", "均价",
)


def _config_color(value, default: QColor) -> QColor:
    color = QColor(value) if isinstance(value, (QColor, str)) else QColor()
    return color if color.isValid() else QColor(default)


class FloatLabel(DragBehaviorMixin, QWidget):
    widget_visibility_changed = Signal()
    view_options_changed = Signal()
    drag_started = Signal()
    drag_moved = Signal()
    drag_finished = Signal(bool)
    taskbar_options_changed = Signal()
    taskbar_status_changed = Signal(str)
    presentation_changed = Signal()
    hotkey_triggered = Signal()
    click_through_hotkey_triggered = Signal()
    click_through_changed = Signal(bool)
    topmost_changed = Signal()
    display_flags_changed = Signal()  # 显示指标/表头/网格/统一颜色等显示相关设置变化
    hide_options_changed = Signal()
    position_options_changed = Signal()
    data_ready = Signal(object)  # (ok, data, error)

    def __init__(self, cfg: dict, codes_list: dict):
        super().__init__()
        self._on_change = (lambda: None)
        self._open_settings_cb = None
        self.taskbar_status = "任务栏显示已关闭"
        self.widget_visible = True
        self.taskbar_preview_active = False
        self.float_page = self.taskbar_page = 0
        self._ordered_rows = []
        self._ordered_color_roles = []
        self._quote_generation = 0
        self._updating_topmost = False

        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setFocusPolicy(Qt.StrongFocus)
        if sys.platform == "darwin":
            self.setAttribute(Qt.WA_MacAlwaysShowToolWindow, True)

        self.codes_list: dict = codes_list
        # 加载自选标的配置（代码 -> {checked, cost, name, type}）
        watchlist_cfg = cfg.get("watchlist", {})
        self.watchlist: dict = normalize_watchlist(watchlist_cfg, self.codes_list)
        self._load_appearance_config(cfg)
        self._load_settings_config(cfg)
        self.setWindowFlag(Qt.WindowStaysOnTopHint, self.float_on_top)

        # 排序是浮窗运行时的视图状态，不改变或持久化自选列表顺序。
        self.sort_header = None
        self.sort_order = Qt.SortOrder.DescendingOrder
        self._last_full_rows = []
        self._last_color_roles = []
        self._last_sort_values = []

        # 平台能力限制:当前平台不支持时强制关闭对应功能
        # (如 Wayland 下无法实现全局快捷键/鼠标穿透,Linux 下强制置顶不可靠),
        # 并交由设置面板/托盘菜单将相关控件置为不可点按。
        if not hotkeys_supported():
            self.hotkey_enabled = False
            self.hotkey_click_through_enabled = False
        if not click_through_supported():
            self.click_through = False
            # 鼠标穿透不可用（如 macOS）时，其快捷键一并关闭，避免"按了没效果"
            self.hotkey_click_through_enabled = False
        if not force_top_supported():
            self.force_top = False

        # Wayland 会话下窗口位置由合成器接管,须用系统级拖动(startSystemMove)
        self._wayland_drag = is_wayland()

        self.hotkey_triggered.connect(self.toggle_win)
        self.click_through_hotkey_triggered.connect(self.toggle_click_through)
        # 全局快捷键管理器:Windows 用官方 RegisterHotKey,Linux/X11 用 XGrabKey,Wayland 不支持
        self._hotkeys = GlobalHotkeyManager(self)
        self._register_current()

        # UI
        from stockwidget.ui.generated.ui_floating_widget import Ui_FloatingWidget
        self.ui = Ui_FloatingWidget()
        self.ui.setupUi(self)
        self.panel = self.ui.panel
        self.vbox = self.ui.vbox
        self.table = self.ui.table
        self.right_table = self.ui.right_table
        self.split_separator = self.ui.split_separator
        self.float_tables = (self.table, self.right_table)
        self.message_label = self.ui.message_label
        self.hide_notice = self.ui.hide_notice
        self.hide_notice.hide()
        self.data_layout = self.ui.data_layout
        self.pager = self.ui.pager

        for table in self.float_tables:
            table.verticalHeader().setVisible(False)
            table.verticalHeader().setMinimumSectionSize(1)
            table.verticalHeader().setDefaultSectionSize(1)
            header = table.horizontalHeader()
            header.setVisible(self.header_visible)
            header.setStretchLastSection(False)
            header.setSectionResizeMode(QHeaderView.ResizeToContents)
            header.setStyle(SortIndicatorStyle(header))
            header.setSectionsClickable(True)
            header.setSortIndicatorShown(False)
            header.sectionClicked.connect(self._on_header_clicked)
            table.setFont(self.font)
            header.setFont(self.font)
        self.message_label.setVisible(False)
        self._message_kind = None
        self._refresh_thread = None  # 后台刷新线程（避免网络请求阻塞 UI）

        # 首次报价到达前只显示紧凑提示，避免所有列撑出临时长条。
        self.model = SimpleTableModel(headers=[], align_right_cols=[])
        self.right_model = SimpleTableModel(parent=self)
        self.taskbar_model = SimpleTableModel(parent=self)
        self.table.setModel(self.model)
        self.right_table.setModel(self.right_model)

        self.k_delegate = KLineDelegate(self.table, base_pt=12)
        self.right_k_delegate = KLineDelegate(self.right_table, base_pt=12)
        self.taskbar_k_delegate = KLineDelegate(self, base_pt=12)
        self._default_item_delegate = QStyledItemDelegate(self.table)
        self._right_default_delegate = QStyledItemDelegate(self.right_table)
        self._sync_colors_to_views()
        self.k_delegate.set_point_size(self.font.pointSize())
        self.k_column_visible_index = None

        self.pager.page_requested.connect(lambda delta: self.change_page("float", delta))
        self.page_timers = {}
        for surface in ("float", "taskbar"):
            timer = QTimer(self)
            timer.timeout.connect(partial(self.change_page, surface, 1, automatic=True))
            self.page_timers[surface] = timer
        self.table.hide()
        self.right_table.hide()
        self.split_separator.hide()
        self._show_message("加载中…", kind="loading")

        # 点击与拖动统一判定，表头只有在未发生拖动时才触发排序。
        header = self.table.horizontalHeader()
        self._init_drag(header)
        for w in (
            self.panel, self.message_label, self.hide_notice, self.table, self.table.viewport(),
            header.viewport(), self.right_table, self.right_table.viewport(), self.split_separator,
        ):
            self.register_drag_region(w)
        right_header = self.right_table.horizontalHeader()
        self.register_drag_region(right_header.viewport(),
                                  lambda pos: self._on_header_clicked(right_header.logicalIndexAt(pos)))
        self.register_drag_region(self.pager, self.pager.activate_at)
        self.position_controller = PositionController(self)

        self.apply_style()
        self.set_window_opacity_percent(self.opacity_pct)
        self._fit_to_contents()

        self._restore_position(cfg.get("pos"))
        self.hide_controller = HideController(self)

        # 定时刷新数据
        self.data_ready.connect(self._process_data)
        self.timer = QTimer(self)
        self.timer.setInterval(max(1, self.refresh_seconds)*1000)
        self.timer.timeout.connect(self._refresh_from_function)
        self.timer.start()
        self._refresh_from_function()
        self._defer_fit()

        # 强制置顶定时
        self._keep_top_timer = QTimer(self)
        self._keep_top_timer.setInterval(1000)  # 每 1000ms 检查一次
        self._keep_top_timer.timeout.connect(self._ensure_on_top)
        if self.force_top:
            self._keep_top_timer.start()

        self.set_click_through(self.click_through)

    def _load_appearance_config(self, cfg: dict):
        self.header_visible = bool(cfg.get("header_visible", False))
        self.grid_visible = bool(cfg.get("grid_visible", False))
        font_family = cfg.get("font_family", default_font_family())
        font_size = int(cfg.get("font_size", 10))
        self.font = QFont(font_family, max(5, min(15, font_size)))
        self.line_extra_px = int(cfg.get("line_extra_px", 1))
        self.fg = QColor(cfg.get("fg", "#FFFFFF"))
        bg = cfg.get("bg", {"r":0,"g":0,"b":0,"a":191})
        self.bg = QColor(bg["r"],bg["g"],bg["b"],bg["a"])
        self.opacity_pct = int(cfg.get("opacity_pct", 90))
        self.unicolor = bool(cfg.get("unicolor", True))
        self.up_color = _config_color(cfg.get("up_color"), DEFAULT_UP_COLOR)
        self.down_color = _config_color(cfg.get("down_color"), DEFAULT_DOWN_COLOR)
        self.neutral_color = _config_color(cfg.get("neutral_color"), DEFAULT_NEUTRAL_COLOR)

    def _load_settings_config(self, cfg: dict):
        # 加载面板配置
        self.code_visible = bool(cfg.get("code_visible", False))
        self.type_visible = bool(cfg.get("type_visible", False))
        self.name_length = int(cfg.get("name_length", -1))
        self.unit_mode = str(cfg.get("unit_mode", "auto"))
        if self.unit_mode not in ("cn", "en", "auto"):
            self.unit_mode = "auto"
        self.visible_metrics = visible_metrics_from_config(cfg)
        # 加载其他配置
        self.refresh_seconds = int(cfg.get("refresh_seconds", 2))
        self.data_source = str(cfg.get("data_source", "sina"))
        if self.data_source not in ("sina", "eastmoney"):
            self.data_source = "sina"
        self.scheduled_hide_enabled = bool(cfg.get("scheduled_hide_enabled", False))
        self.scheduled_hide_times = normalize_hide_times(cfg.get("scheduled_hide_times", []))
        self.auto_hide_enabled = bool(cfg.get("auto_hide_enabled", False))
        self.boundary_check_enabled = bool(cfg.get("boundary_check_enabled", False))
        self.edge_hide_enabled = self.boundary_check_enabled and bool(cfg.get("edge_hide_enabled", False))
        self.float_on_top = bool(cfg.get("float_on_top", True))
        self.force_top = self.float_on_top and bool(cfg.get("force_top", False))
        self.click_through = bool(cfg.get("click_through", False))
        self.hotkey_enabled = bool(cfg.get("hotkey_enabled", False))
        self.hotkey = cfg.get("hotkey", "Ctrl+Alt+F")
        self.hotkey_click_through_enabled = bool(cfg.get("hotkey_click_through_enabled", False))
        self.hotkey_click_through = cfg.get("hotkey_click_through", "Ctrl+Alt+C")
        self.start_on_boot = bool(cfg.get("start_on_boot", False))
        mode = cfg.get("display_mode", "float")
        self.display_mode = mode if sys.platform == "win32" and mode in ("float", "taskbar", "both") else "float"
        try:
            self.taskbar_offset = max(0, min(2000, int(cfg.get("taskbar_offset", 0))))
        except (TypeError, ValueError):
            self.taskbar_offset = 0
        self.view_options = ViewOptions.from_config(cfg, self.font.family(), self.visible_metrics)
        if not self.view_options.taskbar_enabled:
            self.display_mode = "float"
        elif self.display_mode != "float":
            self.display_mode = "both" if self.view_options.taskbar_dual_open else "taskbar"

    def reset_appearance(self):
        self._load_appearance_config({})
        defaults = ViewOptions.from_config({}, self.font.family())
        for key in TASKBAR_STYLE_KEYS:
            setattr(self.view_options, key, getattr(defaults, key))
        self.view_options_changed.emit()
        self._sync_colors_to_views()
        self.k_delegate.set_point_size(self.font.pointSize())
        for table in self.float_tables:
            table.horizontalHeader().setVisible(self.header_visible)
        self.apply_style()
        self.set_window_opacity_percent(self.opacity_pct)
        self.display_flags_changed.emit()

    def reset_settings(self):
        self._load_settings_config({key: getattr(self.view_options, key) for key in TASKBAR_STYLE_KEYS})
        self.position_controller.recheck()
        self.position_options_changed.emit()
        self._apply_float_on_top()
        self.topmost_changed.emit()
        self._quote_generation += 1
        self.hide_controller.cancel_countdown()
        self.hide_controller.configure()
        self.hide_options_changed.emit()
        self.float_page = self.taskbar_page = 0
        self._reproject_cached_data()
        self.view_options_changed.emit()
        self.taskbar_options_changed.emit()
        self.clear_sort()
        self._register_current()
        self._keep_top_timer.stop()
        apply_click_through(self, self.click_through)
        self.click_through_changed.emit(self.click_through)
        self.timer.setInterval(max(1, self.refresh_seconds) * 1000)
        self._refresh_from_function()
        self.display_flags_changed.emit()
        self._notify_change()

    # ----- 自选标的派生属性（由 watchlist 生成） -----

    @property
    def checked_codes(self) -> dict:
        return {c: dict(e) for c, e in self.watchlist.items() if e.get("checked")}

    # 与 App 连接
    def set_open_settings_callback(self, fn):
        self._open_settings_cb = fn

    def set_on_change(self, fn):
        self._on_change = fn or (lambda: None)

    def _notify_change(self):
        self.presentation_changed.emit()
        self._on_change()

    def current_config(self):
        return {
            "watchlist": {c: dict(e) for c, e in self.watchlist.items()},

            "code_visible": self.code_visible,
            "type_visible": self.type_visible,
            "name_length": self.name_length,
            "unit_mode": self.unit_mode,
            "visible_metrics": list(self.visible_metrics),
            **legacy_visibility(self.visible_metrics),

            "header_visible": self.header_visible,
            "grid_visible": self.grid_visible,
            "font_family": self.font.family(),
            "font_size": self.font.pointSize(),
            "line_extra_px": self.line_extra_px,
            "fg": self.fg.name(QColor.HexRgb),
            "bg": {"r": self.bg.red(), "g": self.bg.green(), "b": self.bg.blue(), "a": self.bg.alpha()},
            "opacity_pct": self.opacity_pct,
            "unicolor": self.unicolor,
            "up_color": self.up_color.name(QColor.HexRgb),
            "down_color": self.down_color.name(QColor.HexRgb),
            "neutral_color": self.neutral_color.name(QColor.HexRgb),

            "refresh_seconds": self.refresh_seconds,
            "data_source": self.data_source,
            "scheduled_hide_enabled": self.scheduled_hide_enabled,
            "scheduled_hide_times": list(self.scheduled_hide_times),
            "auto_hide_enabled": self.auto_hide_enabled,
            "boundary_check_enabled": self.boundary_check_enabled,
            "edge_hide_enabled": self.edge_hide_enabled,
            "float_on_top": self.float_on_top,
            "force_top": self.force_top,
            "click_through": self.click_through,
            "hotkey_enabled": self.hotkey_enabled,
            "hotkey": self.hotkey,
            "hotkey_click_through_enabled": self.hotkey_click_through_enabled,
            "hotkey_click_through": self.hotkey_click_through,
            "start_on_boot": self.start_on_boot,
            "display_mode": self.display_mode,
            "taskbar_offset": self.taskbar_offset,
            **self.view_options.to_config(),
            "pos": {"x": self.position_controller.full_geometry().x(),
                    "y": self.position_controller.full_geometry().y()},
        }

    # ----- 外观/尺寸 -----
    def _sync_colors_to_views(self):
        colors = (
            self.unicolor,
            self.fg,
            self.up_color,
            self.down_color,
            self.neutral_color,
        )
        self.model.set_colors(*colors)
        self.right_model.set_colors(*colors)
        self.k_delegate.set_colors(*colors)
        self.right_k_delegate.set_colors(*colors)
        taskbar_colors = colors if self.view_options.taskbar_sync_appearance else (
            self.view_options.taskbar_unicolor, _config_color(self.view_options.taskbar_color, self.fg),
            self.up_color, self.down_color, self.neutral_color,
        )
        self.taskbar_model.set_colors(*taskbar_colors)
        self.taskbar_k_delegate.set_colors(*taskbar_colors)
        if hasattr(self, "pager"):
            self.pager.set_page(self.get_page("float"), self.fg, self.font)
        for table in self.float_tables:
            table.viewport().update()
            table.horizontalHeader().viewport().update()

    def apply_style(self):
        # 先更新字体，再让样式表解析表头字重，避免尺寸计算与绘制使用不同字号。
        for table in self.float_tables:
            table.setFont(self.font)
            table.horizontalHeader().setFont(self.font)
        for delegate in (self.k_delegate, self.right_k_delegate):
            delegate.set_point_size(self.font.pointSize())
        self.hide_notice.setFont(self.font)
        self.hide_notice.setStyleSheet(f"color: {self.fg.name(QColor.HexRgb)}; padding: 2px 4px;")
        r,g,b,a = self.bg.red(), self.bg.green(), self.bg.blue(), self.bg.alpha()
        fg_r, fg_g, fg_b = self.fg.red(), self.fg.green(), self.fg.blue()
        line_col = f"rgba({fg_r},{fg_g},{fg_b},80)"
        self.panel.setStyleSheet(f"""
            QWidget#panel {{
                background: rgba({r},{g},{b},{a});
                border-radius: 5px;
            }}
            QTableView {{
                background: transparent;
                border: {f"1px solid {line_col}" if self.grid_visible else "none"};
                border-radius: 3px;
                color: {self.fg.name()};
                outline: none;
            }}
            QTableView::item {{
                border-right: {f"1px solid {line_col}" if self.grid_visible else "none"};
                border-bottom: {f"1px solid {line_col}" if self.grid_visible else "none"};
            }}
            QHeaderView {{
                background-color: transparent;
            }}
            QHeaderView::section {{
                background: transparent;
                border: none;
                border-bottom: 1px solid {line_col};
                font-weight: 600;
                color: {self.fg.name()};
                padding: 2px 4px;
            }}
            QFrame#split_separator {{
                color: {line_col};
            }}
        """)
        self._defer_fit()

    def _apply_row_heights(self):
        fm = self.table.fontMetrics()
        h = fm.height() + max(0, self.line_extra_px)
        for table in self.float_tables:
            table.verticalHeader().setDefaultSectionSize(h)
            for r in range(table.model().rowCount()):
                table.setRowHeight(r, h)

    def _fit_to_contents(self):
        split, _separator = self.get_split_settings("float")
        for table in self.float_tables:
            table.horizontalHeader().setStretchLastSection(False)
            table.horizontalHeader().setSectionResizeMode(QHeaderView.Fixed if split else QHeaderView.ResizeToContents)
            table.resizeColumnsToContents()
        self._apply_row_heights()
        widths = [max(table.columnWidth(c) for table in (self.float_tables if split else (self.table,)))
                  for c in range(self.model.columnCount())]
        for table in self.float_tables:
            table.verticalHeader().setFixedWidth(0)
            for c, width in enumerate(widths):
                table.setColumnWidth(c, width)
            hh = table.horizontalHeader().sizeHint().height() if self.header_visible else 0
            total_h = hh + 2 * table.frameWidth() + sum(
                self.table.rowHeight(r) for r in range(self.model.rowCount()))
            table.setFixedSize(max(1, sum(widths) + 2 * table.frameWidth()), max(1, total_h))
        # 固定表格尺寸后，嵌套布局的 sizeHint 仍可能缓存上一轮数据。
        # 同步重算布局再调整外框，避免浮窗大小慢一轮刷新。
        self.data_layout.invalidate()
        self.vbox.invalidate()
        self.vbox.activate()
        self.panel.adjustSize()
        self.position_controller.fit(self.panel.size())

    def _defer_fit(self):
        QTimer.singleShot(0, self.table, self._fit_to_contents)

    def _restore_position(self, pos_cfg):
        """默认原样恢复保存位置；启用边界检测时按完整浮窗尺寸调整。"""
        screens = QApplication.screens()
        rects = []
        for s in screens:
            g = s.availableGeometry()
            rects.append((g.left(), g.top(), g.width(), g.height()))
        pg = QApplication.primaryScreen().availableGeometry()
        primary = (pg.left(), pg.top(), pg.width(), pg.height())

        saved = None
        if isinstance(pos_cfg, dict) and "x" in pos_cfg and "y" in pos_cfg:
            saved = (int(pos_cfg["x"]), int(pos_cfg["y"]))

        x, y = saved if saved is not None else resolve_restore_position(
            None, rects, primary, self.width(), self.height())
        self.position_controller.restore_position(QPoint(x, y))

    # ----- 数据 & 投影 -----
    def _show_message(self, msg: str, is_error: bool = False, *, kind=None):
        """显示顶部提示；is_error=True 时用红色字体，否则用前景色"""
        text = str(msg) if msg is not None else ""
        color = "#ff6666" if is_error else self.fg.name(QColor.HexRgb)
        self.message_label.setStyleSheet(f"color: {color}; padding: 2px 4px;")
        self.message_label.setText(text)
        self.message_label.setVisible(True)
        self._message_kind = kind
        self._defer_fit()
        self.presentation_changed.emit()

    def _clear_message(self):
        """清除顶部提示"""
        self.message_label.setVisible(False)
        self.message_label.setText("")
        self._message_kind = None
        self.presentation_changed.emit()

    def _effective_visible_metrics(self):
        """排序期间强制显示名称，但不改写用户保存的指标显示配置。"""
        metrics = list(self.visible_metrics)
        if self.sort_header is not None and NAME_METRIC_ID not in metrics:
            metrics.insert(0, NAME_METRIC_ID)
        return metrics

    def _sorted_row_indices(self, sort_values: list[dict]):
        indices = list(range(len(sort_values)))
        if self.sort_header not in SORTABLE_HEADERS:
            return indices

        valid = [i for i in indices if sort_values[i].get(self.sort_header) is not None]
        missing = [i for i in indices if sort_values[i].get(self.sort_header) is None]
        reverse = self.sort_order == Qt.SortOrder.DescendingOrder
        valid.sort(key=lambda i: sort_values[i][self.sort_header], reverse=reverse)
        return valid + missing

    def _project_columns(self, full_rows: list[dict], color_roles: list[dict], sort_values=None):
        sort_values = sort_values or [{} for _ in full_rows]
        indices = self._sorted_row_indices(sort_values)
        self._ordered_rows = [full_rows[i] for i in indices]
        self._ordered_color_roles = [color_roles[i] for i in indices]
        self._project_float_page()
        self._project_taskbar_page()
        self._sync_page_timers()
        self.presentation_changed.emit()

    def get_page(self, surface):
        options = self.view_options
        limit = options.float_max_rows if surface == "float" else options.taskbar_rows
        if surface == "float" and not options.float_paging_enabled:
            limit = 0
        if self.get_split_settings(surface)[0]:
            limit *= 2
        mode, _interval = self.get_page_settings(surface)
        return page_slice(len(self._ordered_rows), limit, mode,
                          getattr(self, f"{surface}_page"))

    def get_page_settings(self, surface):
        options = self.view_options
        if surface == "float" or options.taskbar_sync_paging:
            return (options.float_page_mode if options.float_paging_enabled else "first",
                    options.float_page_interval)
        return options.taskbar_page_mode, options.taskbar_page_interval

    def get_split_settings(self, surface):
        options = self.view_options
        prefix = "float" if surface == "float" or options.taskbar_sync_split else "taskbar"
        return getattr(options, f"{prefix}_split_enabled"), getattr(options, f"{prefix}_split_separator")

    def get_taskbar_appearance(self):
        options = self.view_options
        if options.taskbar_sync_appearance:
            return self.font, self.fg, self.opacity_pct, self.unicolor
        return (QFont(options.taskbar_font_family, options.taskbar_font_size),
                _config_color(options.taskbar_color, self.fg), options.taskbar_opacity_pct, options.taskbar_unicolor)

    def _project_float_page(self):
        page = self.get_page("float")
        self.float_page = page.index
        full_rows = self._ordered_rows[page.start:page.stop]
        color_roles = self._ordered_color_roles[page.start:page.stop]

        # 名称与其余指标统一按用户配置顺序展开；排序期间名称被临时补到首列。
        headers = expand_metric_headers(self._effective_visible_metrics())

        proj_rows, projected_roles = [], []
        for r, row in enumerate(full_rows):
            proj_rows.append([row.get(h, "-") for h in headers])
            projected_roles.append([color_roles[r].get(h, "text") for h in headers])

        # 右对齐：名称、K线、卖一除外
        right_cols = [i for i, h in enumerate(headers) if h not in ("名称", "K线", "卖一")]
        split, separator = self.get_split_settings("float")
        ranges = column_ranges(len(proj_rows), split)
        for model, (start, stop) in zip((self.model, self.right_model), ranges):
            model.set_align_right_cols(right_cols)
            model.set_rows_headers(proj_rows[start:stop], headers, projected_roles[start:stop])
        if not split:
            self.right_model.set_rows_headers([], headers, [])
        self.table.setVisible(bool(headers))
        self.right_table.setVisible(bool(headers) and split)
        self.split_separator.setVisible(bool(headers) and split and separator)
        self._sync_colors_to_views()

        old_kline_col = self.k_column_visible_index
        new_kline_col = headers.index("K线") if "K线" in headers else None
        for table, delegate, default in ((self.table, self.k_delegate, self._default_item_delegate),
                                         (self.right_table, self.right_k_delegate, self._right_default_delegate)):
            if old_kline_col is not None and old_kline_col != new_kline_col:
                table.setItemDelegateForColumn(old_kline_col, default)
            if new_kline_col is not None:
                delegate.set_point_size(self.font.pointSize())
                table.setItemDelegateForColumn(new_kline_col, delegate)
            header = table.horizontalHeader()
            if self.sort_header in headers:
                header.setSortIndicator(headers.index(self.sort_header), self.sort_order)
                header.setSortIndicatorShown(True)
            else:
                header.setSortIndicatorShown(False)
        self.k_column_visible_index = new_kline_col

        if not headers:
            self._show_message(
                "请在设置面板中选择至少一个显示指标",
                is_error=True,
                kind="metrics",
            )
        elif headers and self._message_kind == "metrics":
            self._clear_message()

        self._fit_to_contents()

    def _project_taskbar_page(self):
        page = self.get_page("taskbar")
        self.taskbar_page = page.index
        metrics = self._effective_visible_metrics() if self.view_options.taskbar_sync_metrics else self.view_options.taskbar_metrics
        headers = expand_metric_headers(metrics)
        rows = [[row.get(h, "-") for h in headers] for row in self._ordered_rows[page.start:page.stop]]
        roles = [[row.get(h, "text") for h in headers] for row in self._ordered_color_roles[page.start:page.stop]]
        self.taskbar_model.set_align_right_cols([i for i, h in enumerate(headers) if h not in ("名称", "K线", "卖一")])
        self.taskbar_model.set_rows_headers(rows, headers, roles)
        self._sync_colors_to_views()

    def change_page(self, surface, delta, *, automatic=False):
        page = self.get_page(surface)
        if not page.controls:
            return
        setattr(self, f"{surface}_page", (page.index + delta) % page.count)
        if surface == "float":
            self._project_float_page()
        else:
            self._project_taskbar_page()
        self.presentation_changed.emit()
        if not automatic and self.page_timers[surface].isActive():
            self.page_timers[surface].start()

    def _sync_page_timers(self):
        if not hasattr(self, "page_timers"):
            return
        for surface, timer in self.page_timers.items():
            active = (self.isVisible() and not self.position_controller.collapsed) if surface == "float" else (
                self.widget_visible and (self.display_mode != "float" or self.taskbar_preview_active))
            mode, seconds = self.get_page_settings(surface)
            interval = seconds * 1000
            if timer.interval() != interval:
                timer.setInterval(interval)
            if active and self.get_page(surface).controls and mode == "auto":
                if not timer.isActive():
                    timer.start()
            else:
                timer.stop()

    def set_view_options(self, **changes):
        old = self.view_options.to_config()
        old_taskbar_page_mode = self.get_page_settings("taskbar")[0]
        old_split = {surface: self.get_split_settings(surface)[0] for surface in ("float", "taskbar")}
        options = ViewOptions.from_config({**old, **changes}, self.font.family())
        if options.to_config() == old:
            return
        if old["taskbar_enabled"] and not options.taskbar_enabled:
            self.finish_drag(False)
        self.view_options = options
        if self.get_page_settings("taskbar")[0] != old_taskbar_page_mode:
            self.taskbar_page = 0
        for surface in ("float", "taskbar"):
            limit = "float_max_rows" if surface == "float" else "taskbar_rows"
            if (getattr(options, limit) != old[limit]
                    or getattr(options, f"{surface}_page_mode") != old[f"{surface}_page_mode"]
                    or self.get_split_settings(surface)[0] != old_split[surface]
                    or (surface == "float" and options.float_paging_enabled != old["float_paging_enabled"])):
                setattr(self, f"{surface}_page", 0)
        if not options.taskbar_enabled:
            self.display_mode = "float"
            self.taskbar_preview_active = False
        elif self.display_mode != "float":
            self.display_mode = "both" if options.taskbar_dual_open else "taskbar"
        self._project_float_page()
        self._project_taskbar_page()
        self._sync_page_timers()
        self.view_options_changed.emit()
        self.taskbar_options_changed.emit()
        self._notify_change()

    def _reproject_cached_data(self):
        self._project_columns(
            self._last_full_rows,
            self._last_color_roles,
            self._last_sort_values,
        )

    def _refresh_from_function(self):
        """定时入口：将网络请求丢到后台线程执行，避免阻塞 UI。
        若上一轮请求尚未完成则跳过本次刷新，防止请求重叠。"""
        if self.position_controller.collapsed:
            return
        checked_codes = self.checked_codes
        if not checked_codes:
            self._process_data((True, {}, None))
            return
        if self._refresh_thread is not None and self._refresh_thread.is_alive():
            return
        self._refresh_thread = threading.Thread(
            target=self._fetch_data_worker,
            args=(checked_codes, self.data_source, self._quote_generation),
            daemon=True,
        )
        self._refresh_thread.start()

    def _fetch_data_worker(self, codes: dict, source: str, generation: int):
        """后台线程：执行网络请求，结果经 data_ready 信号回到主线程。"""
        source_name = {"sina": "新浪", "eastmoney": "东财"}.get(source, "行情")
        try:
            data = request_quote(codes, source=source)
            payload = (True, data, None)
        except requests.exceptions.RequestException as exc:
            payload = (False, None, f"{source_name}：{request_error_message(exc)}")
        except Exception as e:
            payload = (False, None, f"{source_name}：行情处理失败（{type(e).__name__}）")
        self.data_ready.emit((*payload, generation))

    def _process_data(self, payload):
        """主线程：处理请求结果并更新表格。payload = (ok, data, error)"""
        if self.position_controller.collapsed:
            return
        ok, data, error = payload[:3]
        if len(payload) > 3 and payload[3] != self._quote_generation:
            return
        checked_codes = self.checked_codes
        if not checked_codes:
            ok, data = True, {}
        elif ok:
            # 请求期间可能删除、取消勾选或调整顺序，以当前自选列表为准。
            data = {code: data[code] for code in checked_codes if code in data}
        if not ok:
            self.hide_controller.cancel_countdown()
            self._show_message(error or "请求失败", is_error=True)
            self._fit_to_contents()
            return

        full_rows = []
        full_color_roles = []
        sort_values = []
        options = QuoteDisplayOptions(
            name_length=self.name_length,
            code_visible=self.code_visible,
            type_visible=self.type_visible,
            unit_mode=self.unit_mode,
        )
        for c, d in data.items():
            entry = self.watchlist.get(c) or {}
            code_info = self.codes_list.get(c, {})
            type_ = entry.get("type") or code_info.get("type")
            market = entry.get("market") or code_info.get("market") or ""
            display_code = entry.get("code") or code_info.get("code") or c
            row, color_roles, row_sort_values = format_quote(
                d, type_, display_code, market=market,
                cost=entry.get("cost"), options=options,
                include_sort=True,
            )
            full_rows.append(row)
            full_color_roles.append(color_roles)
            sort_values.append(row_sort_values)

        self._last_full_rows = full_rows
        self._last_color_roles = full_color_roles
        self._last_sort_values = sort_values

        if data:
            self._clear_message()
        else:
            self._show_message("请在设置面板中添加自选股", is_error=True)
        self._project_columns(full_rows, full_color_roles, sort_values)
        self.hide_controller.quotes_refreshed(data)

    # ----- 排序 -----
    def set_sort(self, header: str, order):
        if header not in SORTABLE_HEADERS:
            return
        visible_headers = expand_metric_headers(self.visible_metrics)
        if header not in visible_headers:
            return
        order = Qt.SortOrder(order)
        if self.sort_header == header and self.sort_order == order:
            return
        self.sort_header = header
        self.sort_order = order
        self.float_page = self.taskbar_page = 0
        self._reproject_cached_data()

    def clear_sort(self):
        if self.sort_header is None:
            return
        self.sort_header = None
        self.float_page = self.taskbar_page = 0
        self.table.horizontalHeader().setSortIndicatorShown(False)
        self._reproject_cached_data()

    def _on_header_clicked(self, section: int):
        header_name = self.model.headerData(
            section, Qt.Orientation.Horizontal, Qt.ItemDataRole.DisplayRole
        )
        if header_name not in SORTABLE_HEADERS:
            return
        if self.sort_header != header_name:
            self.set_sort(header_name, Qt.SortOrder.DescendingOrder)
        elif self.sort_order == Qt.SortOrder.DescendingOrder:
            self.set_sort(header_name, Qt.SortOrder.AscendingOrder)
        else:
            self.clear_sort()

    # ----- 应用设置 -----
    def set_watchlist(self, watchlist: dict):
        """整体替换自选列表（key -> {code, market, checked, cost, name, type}）。"""
        self.watchlist = normalize_watchlist(watchlist, self.codes_list)
        self._quote_generation += 1
        self.hide_controller.cancel_countdown()
        self.float_page = self.taskbar_page = 0
        self._notify_change()
        self._refresh_from_function()

    def set_codes_list(self, codes_list: dict):
        """替换代码表，并用新代码表补齐自选项元数据。"""
        self.codes_list = codes_list or {}
        self.watchlist = normalize_watchlist(self.watchlist, self.codes_list)
        self._quote_generation += 1
        self.hide_controller.cancel_countdown()

    def set_type_visible(self, visible: bool):
        self.type_visible = bool(visible)
        self._notify_change()
        self._refresh_from_function()
        self.display_flags_changed.emit()

    def set_code_visible(self, visible: bool):
        self.code_visible = bool(visible)
        self._notify_change()
        self._refresh_from_function()
        self.display_flags_changed.emit()

    def set_visible_metrics(self, metric_ids):
        """一次性应用指标的显示状态和顺序。"""
        normalized = normalize_visible_metrics(metric_ids)
        if normalized == self.visible_metrics:
            return
        self.visible_metrics = normalized
        if self.sort_header is not None and self.sort_header not in expand_metric_headers(normalized):
            self.sort_header = None
            self.table.horizontalHeader().setSortIndicatorShown(False)
        self._reproject_cached_data()
        self._notify_change()
        self._refresh_from_function()
        self.display_flags_changed.emit()

    def set_metric_visible(self, metric_id: str, visible: bool):
        metric_id = str(metric_id or "")
        if metric_id not in METRIC_BY_ID:
            return
        updated = list(self.visible_metrics)
        if visible:
            if metric_id in updated:
                return
            updated.append(metric_id)
        else:
            if metric_id not in updated:
                return
            updated.remove(metric_id)
        self.set_visible_metrics(updated)

    def set_name_length(self, name_len: int):
        # -1 全部显示, 0 不显示, >0 显示前 N 个字
        if name_len == -1 or name_len >= 0:
            self.name_length = name_len
            self._notify_change()
            self._refresh_from_function()

    def set_unit_mode(self, mode: str):
        """设置成交量/成交额单位模式：cn=中文, en=英文, auto=自动。"""
        mode = str(mode or "").strip().lower()
        if mode not in ("cn", "en", "auto") or mode == self.unit_mode:
            return
        self.unit_mode = mode
        self._notify_change()
        self._refresh_from_function()
        self.display_flags_changed.emit()

    def set_header_visible(self, vis: bool):
        self.header_visible = bool(vis)
        for table in self.float_tables:
            table.horizontalHeader().setVisible(self.header_visible)
        self._notify_change()
        self._defer_fit()
        self.display_flags_changed.emit()

    def set_grid_visible(self, vis: bool):
        self.grid_visible = bool(vis)
        self.apply_style()
        self._notify_change()
        self.display_flags_changed.emit()

    def set_refresh_interval(self, seconds: int):
        try:
            seconds = max(1, int(seconds))
        except Exception:
            return
        self.refresh_seconds = seconds
        if self.timer is not None:
            self.timer.setInterval(seconds * 1000)
        self._notify_change()

    def set_data_source(self, source: str):
        """切换行情数据源：'sina'（新浪）或 'eastmoney'（东方财富）。"""
        source = str(source or "").strip().lower()
        if source not in ("sina", "eastmoney") or source == self.data_source:
            return
        self.data_source = source
        self._quote_generation += 1
        self.hide_controller.cancel_countdown()
        self._notify_change()
        self._refresh_from_function()

    def set_hide_options(self, *, scheduled_hide_enabled=None, scheduled_hide_times=None,
                         auto_hide_enabled=None):
        if scheduled_hide_enabled is not None:
            self.scheduled_hide_enabled = bool(scheduled_hide_enabled)
        if scheduled_hide_times is not None:
            self.scheduled_hide_times = normalize_hide_times(scheduled_hide_times)
        if auto_hide_enabled is not None:
            self.auto_hide_enabled = bool(auto_hide_enabled)
        self.hide_controller.configure()
        self.hide_options_changed.emit()
        self._notify_change()

    def set_fg_color(self, c: QColor):
        if isinstance(c, QColor) and c.isValid():
            self.fg = QColor(c)
            self._sync_colors_to_views()
            self.apply_style()
            self._notify_change()

    def set_bg_rgb_keep_alpha(self, c: QColor):
        if isinstance(c, QColor) and c.isValid():
            c2 = QColor(c)
            c2.setAlpha(self.bg.alpha())
            self.bg = c2
            self.apply_style()
            self._notify_change()

    def set_bg_alpha_percent(self, percent_0_100: int):
        p = max(0, min(100, int(percent_0_100)))
        self.bg.setAlpha(int(round(p*2.55)))
        self.apply_style()
        self._notify_change()

    def set_window_opacity_percent(self, percent_20_100: int):
        p = max(20, min(100, int(percent_20_100)))
        self.opacity_pct = p
        if opacity_supported():
            # Wayland 平台插件不支持设置窗口透明度,跳过以避免终端告警
            self.setWindowOpacity(p / 100.0)
        self._defer_fit()
        self._notify_change()

    def set_font_size(self, pt: int):
        pt = max(5, min(15, int(pt)))
        self.font.setPointSize(pt)
        self.k_delegate.set_point_size(pt)
        self.apply_style()
        self._notify_change()
        self.table.viewport().update()
        self._defer_fit()

    def set_font_family(self, family: str):
        if family and family != self.font.family():
            self.font.setFamily(family)
            self.apply_style()
            self._notify_change()

    def set_line_extra(self, px: int):
        self.line_extra_px = max(0, int(px))
        self.apply_style()
        self._defer_fit()
        self._notify_change()

    def _set_direction_color(self, attr: str, color: QColor):
        if not isinstance(color, QColor) or not color.isValid():
            return
        setattr(self, attr, QColor(color))
        self._sync_colors_to_views()
        self._notify_change()

    def set_up_color(self, color: QColor):
        self._set_direction_color("up_color", color)

    def set_down_color(self, color: QColor):
        self._set_direction_color("down_color", color)

    def set_neutral_color(self, color: QColor):
        self._set_direction_color("neutral_color", color)

    def set_unicolor(self, enabled: bool):
        enabled = bool(enabled)
        if self.unicolor == enabled:
            return
        self.unicolor = enabled
        self._sync_colors_to_views()
        self.apply_style()
        self._notify_change()
        self._defer_fit()
        self.display_flags_changed.emit()

    # ----- 鼠标穿透 / 浮窗置顶 / 强制置顶 / 快捷键开关 -----
    def set_display_mode(self, mode):
        if mode not in ("float", "taskbar", "both"):
            return
        if sys.platform != "win32" or not self.view_options.taskbar_enabled:
            mode = "float"
        if mode == self.display_mode:
            self.set_widget_visible(True)
            return
        self.display_mode = mode
        if mode != "float":
            self.view_options.taskbar_dual_open = mode == "both"
        was_hidden = not self.widget_visible
        self.widget_visible = True
        if was_hidden:
            self.widget_visibility_changed.emit()
        self.sync_refresh_timer()
        self.taskbar_options_changed.emit()
        self._notify_change()

    def set_taskbar_offset(self, value):
        self.taskbar_offset = max(0, min(2000, int(value)))
        self.taskbar_options_changed.emit()
        self._notify_change()

    def sync_refresh_timer(self):
        if not self.position_controller.collapsed and (
                self.isVisible() or (self.widget_visible and (self.display_mode != "float" or self.taskbar_preview_active))):
            if not self.timer.isActive():
                self.timer.start()
                self._refresh_from_function()
        else:
            self.timer.stop()
        self._sync_page_timers()

    def set_position_options(self, *, boundary_check_enabled=None, edge_hide_enabled=None):
        boundary = self.boundary_check_enabled if boundary_check_enabled is None else bool(boundary_check_enabled)
        edge_hide = self.edge_hide_enabled if edge_hide_enabled is None else bool(edge_hide_enabled)
        edge_hide = boundary and edge_hide
        if (boundary, edge_hide) == (self.boundary_check_enabled, self.edge_hide_enabled):
            return
        self.boundary_check_enabled, self.edge_hide_enabled = boundary, edge_hide
        self.position_controller.recheck()
        self.position_options_changed.emit()
        self._notify_change()

    def set_click_through(self, enable: bool):
        enable = bool(enable)
        if self.click_through == enable:
            return
        self.click_through = enable
        apply_click_through(self, self.click_through)
        self._ensure_on_top()
        self.click_through_changed.emit(self.click_through)
        self._notify_change()

    def toggle_click_through(self):
        self.set_click_through(not self.click_through)

    def _apply_float_on_top(self):
        if bool(self.windowFlags() & Qt.WindowStaysOnTopHint) == self.float_on_top:
            return
        visible = self.isVisible()
        geometry = self.geometry()
        show_without_activating = self.testAttribute(Qt.WA_ShowWithoutActivating)
        # Qt 修改窗口标志会暂时隐藏窗口；这不是用户隐藏，不能联动任务栏或隐藏倒计时。
        self._updating_topmost = True
        try:
            self.setWindowFlag(Qt.WindowStaysOnTopHint, self.float_on_top)
            self.setGeometry(geometry)
            if visible:
                self.setAttribute(Qt.WA_ShowWithoutActivating, True)
                self.show()
        finally:
            self.setAttribute(Qt.WA_ShowWithoutActivating, show_without_activating)
            self._updating_topmost = False

    def set_float_on_top(self, enabled: bool):
        enabled = bool(enabled)
        if self.float_on_top == enabled:
            return
        self.float_on_top = enabled
        if not enabled:
            self.force_top = False
            self._keep_top_timer.stop()
        self._apply_float_on_top()
        self.topmost_changed.emit()
        self._notify_change()

    def set_force_top(self, enabled: bool):
        enabled = bool(enabled)
        if not self.float_on_top or not force_top_supported():
            # 仅 Windows 支持强制置顶,其余平台忽略该开关
            enabled = False
        if self.force_top == enabled:
            return
        self.force_top = enabled
        if self.force_top:
            if self.isVisible() and self._keep_top_timer and not self._keep_top_timer.isActive():
                self._keep_top_timer.start()
            self._ensure_on_top()
        else:
            if self._keep_top_timer and self._keep_top_timer.isActive():
                self._keep_top_timer.stop()
        self.topmost_changed.emit()
        self._notify_change()

    def _set_hotkey_option(self, attr: str, value, *, register: bool = True) -> HotkeyResult:
        """保留用户设置；注册失败仍允许修改，并单独记录每个快捷键的生效状态。"""
        old = getattr(self, attr)
        key = "hotkey_click_through" if attr.startswith("hotkey_click_through") else "hotkey"
        if value == old and self.hotkey_results.get(key, HotkeyResult(True)):
            return HotkeyResult(True)
        setattr(self, attr, value)
        if register:
            self._register_current()
        if value != old:
            self._notify_change()
        return self.hotkey_results.get(key, HotkeyResult(True)) if register else HotkeyResult(True)

    def set_hotkey_enabled(self, enabled: bool) -> HotkeyResult:
        return self._set_hotkey_option("hotkey_enabled", bool(enabled))

    def set_click_through_hotkey_enabled(self, enabled: bool) -> HotkeyResult:
        return self._set_hotkey_option("hotkey_click_through_enabled", bool(enabled))

    def update_hotkey(self, new_hotkey: str) -> HotkeyResult:
        return self._set_hotkey_option(
            "hotkey", new_hotkey.strip(), register=self.hotkey_enabled,
        )

    def update_click_through_hotkey(self, new_hotkey: str) -> HotkeyResult:
        return self._set_hotkey_option(
            "hotkey_click_through", new_hotkey.strip(),
            register=self.hotkey_click_through_enabled,
        )

    # ----- 交互 -----
    def contextMenuEvent(self, event):
        self.show_context_menu(event.globalPos())

    def build_context_menu(self, surface="float"):
        menu = QMenu(self)
        sub_cols = QMenu("显示指标", menu)
        metrics = self.get_surface_metrics(surface)
        for spec in METRIC_SPECS:
            action = QAction(spec.label, sub_cols, checkable=True)
            action.setChecked(spec.metric_id in metrics)
            action.toggled.connect(partial(self.set_surface_metric_visible, surface, spec.metric_id))
            sub_cols.addAction(action)
        menu.addMenu(sub_cols)

        sort_menu = QMenu("排序", menu)
        visible_headers = set(expand_metric_headers(metrics))
        for header_name in SORTABLE_HEADERS:
            metric_menu = QMenu(header_name, sort_menu)
            metric_menu.setEnabled(header_name in visible_headers)
            asc = QAction("升序", metric_menu, checkable=True)
            desc = QAction("降序", metric_menu, checkable=True)
            asc.setChecked(
                self.sort_header == header_name
                and self.sort_order == Qt.SortOrder.AscendingOrder
            )
            desc.setChecked(
                self.sort_header == header_name
                and self.sort_order == Qt.SortOrder.DescendingOrder
            )
            asc.triggered.connect(
                partial(self.set_sort, header_name, Qt.SortOrder.AscendingOrder)
            )
            desc.triggered.connect(
                partial(self.set_sort, header_name, Qt.SortOrder.DescendingOrder)
            )
            metric_menu.addAction(asc)
            metric_menu.addAction(desc)
            sort_menu.addMenu(metric_menu)
        sort_menu.addSeparator()
        clear_action = QAction("恢复自选顺序", sort_menu)
        clear_action.setEnabled(self.sort_header is not None)
        clear_action.triggered.connect(self.clear_sort)
        sort_menu.addAction(clear_action)
        menu.addMenu(sort_menu)

        act_header = QAction("显示表头", menu, checkable=True)
        act_header.setChecked(self.header_visible)
        act_header.toggled.connect(self.set_header_visible)
        menu.addAction(act_header)

        act_grid = QAction("显示网格",menu, checkable=True)
        act_grid.setChecked(self.grid_visible)
        act_grid.toggled.connect(self.set_grid_visible)
        menu.addAction(act_grid)

        act_color = QAction("统一颜色", menu, checkable=True)
        act_color.setChecked(self.unicolor)
        act_color.toggled.connect(self.set_unicolor)
        menu.addAction(act_color)

        if sys.platform == "win32":
            modes = QMenu("显示位置", menu)
            menu.addMenu(modes)
            for mode, label in DISPLAY_MODES:
                action = modes.addAction(label)
                action.setCheckable(True)
                action.setChecked(self.display_mode == mode)
                action.setEnabled(mode == "float" or self.view_options.taskbar_enabled)
                if mode != "float" and not self.view_options.taskbar_enabled:
                    action.setToolTip("请先在设置的任务栏页启用任务栏模式")
                action.triggered.connect(lambda checked=False, value=mode: self.set_display_mode(value))

        menu.addSeparator()
        act_open_settings = QAction("设置…", menu)
        if callable(self._open_settings_cb):
            act_open_settings.triggered.connect(self._open_settings_cb)
        else:
            act_open_settings.setEnabled(False)
        menu.addAction(act_open_settings)

        menu.addSeparator()
        menu.addAction(QAction("隐藏", menu, triggered=self.hide_widget))
        return menu

    def get_surface_metrics(self, surface):
        if surface == "float" or self.view_options.taskbar_sync_metrics:
            return list(self.visible_metrics)
        return list(self.view_options.taskbar_metrics)

    def set_surface_metric_visible(self, surface, metric_id, visible):
        if surface == "float" or self.view_options.taskbar_sync_metrics:
            self.set_metric_visible(metric_id, visible)
        else:
            metrics = list(self.view_options.taskbar_metrics)
            if visible and metric_id not in metrics:
                metrics.append(metric_id)
            elif not visible and metric_id in metrics:
                metrics.remove(metric_id)
            self.set_view_options(taskbar_metrics=metrics)

    def show_context_menu(self, global_pos, surface="float"):
        menu = self.build_context_menu(surface)
        try:
            menu.exec(global_pos)
        finally:
            menu.deleteLater()

    def closeEvent(self, event):
        event.ignore()
        self.hide_widget()

    def showEvent(self, event):
        super().showEvent(event)
        if not self._updating_topmost and self.display_mode != "taskbar" and not self.widget_visible:
            self.widget_visible = True
            self.widget_visibility_changed.emit()
        self._sync_page_timers()
        self.sync_refresh_timer()
        if self.force_top and self._keep_top_timer and not self._keep_top_timer.isActive():
            self._keep_top_timer.start()
        apply_click_through(self, self.click_through)
        self._defer_fit()
        self._ensure_on_top()

    def hideEvent(self, event):
        super().hideEvent(event)
        if self._updating_topmost:
            return
        if self.display_mode != "taskbar" and self.widget_visible:
            self.widget_visible = False
            self.widget_visibility_changed.emit()
        self.sync_refresh_timer()
        if self._keep_top_timer and self._keep_top_timer.isActive():
            self._keep_top_timer.stop()

    def _ensure_on_top(self):
        if not self.float_on_top or not self.force_top or not self.isVisible():
            return
        try:
            aw = QApplication.activeWindow()
            popup = QApplication.activePopupWidget()
            if aw is not None and aw is not self and not self.isAncestorOf(aw):
                return
            if popup is not None and popup is not self and not self.isAncestorOf(popup):
                return
        except Exception:
            pass
        ensure_topmost(self)

    def _register_current(self) -> HotkeyResult:
        """分别注册两个快捷键，某个失败不妨碍另一个生效。"""
        self._hotkeys.unregister_all()
        self.hotkey_results = {}
        if self.hotkey_enabled:
            self.hotkey_results["hotkey"] = self._hotkeys.register(
                self.hotkey, self.hotkey_triggered.emit)
        if self.hotkey_click_through_enabled:
            self.hotkey_results["hotkey_click_through"] = self._hotkeys.register(
                self.hotkey_click_through, self.click_through_hotkey_triggered.emit,
            )
        for result in self.hotkey_results.values():
            if not result:
                return result
        return HotkeyResult(True)

    def set_widget_visible(self, visible):
        """Visibility belongs to the widget, independently of its display location."""
        visible = bool(visible)
        if visible == self.widget_visible:
            return
        self.widget_visible = visible
        if visible and self.display_mode != "taskbar":
            self.show()
            self.raise_()
            self.activateWindow()
            self.setFocus(Qt.ActiveWindowFocusReason)
        elif not visible:
            self.hide()
        self.widget_visibility_changed.emit()
        self.sync_refresh_timer()
        self._notify_change()

    def hide_widget(self):
        self.finish_drag(False)
        self.set_widget_visible(False)

    def toggle_win(self):
        self.set_widget_visible(not self.widget_visible)
