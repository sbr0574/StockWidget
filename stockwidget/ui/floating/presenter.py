"""行情请求生命周期、共享排序与浮窗 / 任务栏分页投影。

FloatLabel 持有配置和 Qt 视图，QuotePresenter 独占行情缓存、请求代数、
排序和页码。设置通过窗口入口修改，两个显示面共用此展示控制器。
"""

from functools import partial
import threading

from PySide6.QtCore import QObject, Qt, QTimer, Signal

from stockwidget.core.quote_presentation import (
    NAME_METRIC_ID,
    SORTABLE_HEADERS,
    QuoteDisplayOptions,
    format_quote,
    metric_headers,
    sorted_quote_indices,
    project_quote_columns,
)
from stockwidget.core.view_options import column_ranges, page_slice
from stockwidget.data.quotes import fetch_quote_result


class QuotePresenter(QObject):
    data_ready = Signal(object)
    quotes_updated = Signal()

    def __init__(self, window):
        super().__init__(window)
        self.window = window
        self.float_page = self.taskbar_page = 0
        self.sort_header = None
        self.sort_order = Qt.SortOrder.DescendingOrder
        self._ordered_rows = []
        self._ordered_color_roles = []
        self._last_full_rows = []
        self._last_color_roles = []
        self._last_sort_values = []
        self._last_keys = []
        self._latest_quotes = {}
        self._latest_source = None
        self._ordered_keys = []
        self._quote_generation = 0
        self._refresh_thread = None
        self.data_ready.connect(self.accept_result)
        self.page_timers = {}
        for surface in ("float", "taskbar"):
            timer = QTimer(self)
            timer.timeout.connect(partial(self.change_page, surface, 1, automatic=True))
            self.page_timers[surface] = timer

    def _effective_visible_metrics(self):
        """排序期间强制显示名称，但不改写用户保存的指标显示配置。"""
        metrics = list(self.window.visible_metrics)
        if self.sort_header is not None and NAME_METRIC_ID not in metrics:
            metrics.insert(0, NAME_METRIC_ID)
        return metrics

    def invalidate(self, *, reset_pages=False):
        """自选、数据源或配置更改后丢弃在途结果，停止旧行情倒计时。"""
        self._quote_generation += 1
        self._latest_quotes = {}
        self._latest_source = None
        if reset_pages:
            self.float_page = self.taskbar_page = 0
        self.window.hide_controller.cancel_countdown()

    def view_options_changed(self, previous):
        """同步设置只改变有效视图；两处的页码仍独立维护。"""
        options = self.window.view_options
        for surface in ("float", "taskbar"):
            limit = "float_max_rows" if surface == "float" else "taskbar_rows"
            if (getattr(previous, limit) != getattr(options, limit)
                    or previous.page_settings(surface)[0] != options.page_settings(surface)[0]
                    or getattr(previous, f"{surface}_page_mode") != getattr(options, f"{surface}_page_mode")
                    or previous.split_settings(surface)[0] != options.split_settings(surface)[0]
                    or (surface == "float" and previous.float_paging_enabled != options.float_paging_enabled)):
                setattr(self, f"{surface}_page", 0)
        self.reproject()

    def metrics_changed(self):
        """隐藏当前排序指标时清除排序，不改写自选顺序。"""
        if self.sort_header not in metric_headers(self.window.visible_metrics):
            self.sort_header = None
            self.window.table.horizontalHeader().setSortIndicatorShown(False)
        self.reproject()

    def project_rows(self, full_rows: list[dict], color_roles: list[dict], sort_values=None):
        sort_values = sort_values or [{} for _ in full_rows]
        indices = sorted_quote_indices(sort_values, self.sort_header,
                                       descending=self.sort_order == Qt.DescendingOrder)
        self._ordered_rows = [full_rows[i] for i in indices]
        self._ordered_color_roles = [color_roles[i] for i in indices]
        self._ordered_keys = [self._last_keys[i] if i < len(self._last_keys) else None for i in indices]
        self._project_float_page()
        self._project_taskbar_page()
        self.sync_page_timers()
        self.window.presentation_changed.emit()

    def get_page(self, surface):
        options = self.window.view_options
        limit = options.float_max_rows if surface == "float" else options.taskbar_rows
        if surface == "float" and not options.float_paging_enabled:
            limit = 0
        if options.split_settings(surface)[0]:
            limit *= 2
        mode, _interval = options.page_settings(surface)
        return page_slice(len(self._ordered_rows), limit, mode,
                          getattr(self, f"{surface}_page"))

    def instrument_at(self, surface, row, block=0):
        """Resolve identity through sorting, independent pages and split columns."""
        page = self.get_page(surface)
        ranges = column_ranges(page.stop - page.start, self.window.view_options.split_settings(surface)[0])
        if not 0 <= block < len(ranges):
            return None
        start, stop = ranges[block]
        if not 0 <= row < stop - start:
            return None
        offset = page.start + start + row
        key = self._ordered_keys[offset] if offset < len(self._ordered_keys) else None
        return self.window.checked_codes.get(key)

    def quote_for(self, instrument):
        """Return accepted raw data only from the currently selected source."""
        if self._latest_source != self.window.data_source:
            return None
        for key, entry in self.window.watchlist.items():
            if entry.get("checked") and all(entry.get(field) == instrument.get(field) for field in ("market", "code")):
                return self._latest_quotes.get(key)
        return None

    def _project_float_page(self):
        page = self.get_page("float")
        self.float_page = page.index
        full_rows = self._ordered_rows[page.start:page.stop]
        color_roles = self._ordered_color_roles[page.start:page.stop]

        # 名称与其余指标统一按用户配置顺序展开；排序期间名称被临时补到首列。
        headers = metric_headers(self._effective_visible_metrics())

        proj_rows, projected_roles, right_cols = project_quote_columns(full_rows, color_roles, headers)
        split, separator = self.window.view_options.split_settings("float")
        ranges = column_ranges(len(proj_rows), split)
        for model, (start, stop) in zip((self.window.model, self.window.right_model), ranges):
            model.set_align_right_cols(right_cols)
            model.set_rows_headers(proj_rows[start:stop], headers, projected_roles[start:stop])
        if not split:
            self.window.right_model.set_rows_headers([], headers, [])
        self.window.table.setVisible(bool(headers))
        self.window.right_table.setVisible(bool(headers) and split)
        self.window.split_separator.setVisible(bool(headers) and split and separator)
        self.window._sync_colors_to_views()

        old_kline_col = self.window.k_column_visible_index
        new_kline_col = headers.index("K线") if "K线" in headers else None
        for table, delegate, default in ((self.window.table, self.window.k_delegate, self.window._default_item_delegate),
                                         (self.window.right_table, self.window.right_k_delegate, self.window._right_default_delegate)):
            if old_kline_col is not None and old_kline_col != new_kline_col:
                table.setItemDelegateForColumn(old_kline_col, default)
            if new_kline_col is not None:
                delegate.set_point_size(self.window.font.pointSize())
                table.setItemDelegateForColumn(new_kline_col, delegate)
            header = table.horizontalHeader()
            if self.sort_header in headers:
                header.setSortIndicator(headers.index(self.sort_header), self.sort_order)
                header.setSortIndicatorShown(True)
            else:
                header.setSortIndicatorShown(False)
        self.window.k_column_visible_index = new_kline_col

        if not headers:
            self.window._show_message(
                "请在设置面板中选择至少一个显示指标",
                is_error=True,
                kind="metrics",
            )
        elif headers and self.window._message_kind == "metrics":
            self.window._clear_message()

        self.window._fit_to_contents()

    def _project_taskbar_page(self):
        page = self.get_page("taskbar")
        self.taskbar_page = page.index
        metrics = self._effective_visible_metrics() if self.window.view_options.taskbar_sync_metrics else self.window.view_options.taskbar_metrics
        headers = metric_headers(metrics)
        rows, roles, right_cols = project_quote_columns(
            self._ordered_rows[page.start:page.stop], self._ordered_color_roles[page.start:page.stop], headers)
        self.window.taskbar_model.set_align_right_cols(right_cols)
        self.window.taskbar_model.set_rows_headers(rows, headers, roles)
        self.window._sync_colors_to_views()

    def change_page(self, surface, delta, *, automatic=False):
        page = self.get_page(surface)
        if not page.controls:
            return
        setattr(self, f"{surface}_page", (page.index + delta) % page.count)
        if surface == "float":
            self._project_float_page()
        else:
            self._project_taskbar_page()
        self.window.presentation_changed.emit()
        if not automatic and self.page_timers[surface].isActive():
            self.page_timers[surface].start()

    def sync_page_timers(self):
        for surface, timer in self.page_timers.items():
            active = (self.window.isVisible() and not self.window.position_controller.collapsed) if surface == "float" else (
                self.window.widget_visible and (self.window.display_mode != "float" or self.window.taskbar_preview_active))
            mode, seconds = self.window.view_options.page_settings(surface)
            interval = seconds * 1000
            if timer.interval() != interval:
                timer.setInterval(interval)
            if active and self.get_page(surface).controls and mode == "auto":
                if not timer.isActive():
                    timer.start()
            else:
                timer.stop()

    def reproject(self):
        self.project_rows(
            self._last_full_rows,
            self._last_color_roles,
            self._last_sort_values,
        )

    def refresh(self):
        """定时入口：将网络请求丢到后台线程执行，避免阻塞 UI。
        若上一轮请求尚未完成则跳过本次刷新，防止请求重叠。"""
        if self.window.position_controller.collapsed:
            return
        checked_codes = self.window.checked_codes
        if not checked_codes:
            self.accept_result((True, {}, None))
            return
        if self._refresh_thread is not None and self._refresh_thread.is_alive():
            return
        self._refresh_thread = threading.Thread(
            target=self._fetch_worker,
            args=(checked_codes, self.window.data_source, self._quote_generation),
            daemon=True,
        )
        self._refresh_thread.start()

    def _fetch_worker(self, codes: dict, source: str, generation: int):
        """后台线程：执行网络请求，结果经 data_ready 信号回到主线程。"""
        payload = fetch_quote_result(codes, source)
        try:
            self.data_ready.emit((*payload, generation))
        except RuntimeError:
            # The owning window may have been destroyed while the request was running.
            pass

    def accept_result(self, payload):
        """主线程：处理请求结果并更新表格。payload = (ok, data, error)"""
        if self.window.position_controller.collapsed:
            return
        ok, data, error = payload[:3]
        if len(payload) > 3 and payload[3] != self._quote_generation:
            return
        checked_codes = self.window.checked_codes
        if not checked_codes:
            ok, data = True, {}
        elif ok:
            # 请求期间可能删除、取消勾选或调整顺序，以当前自选列表为准。
            data = {code: data[code] for code in checked_codes if code in data}
        if not ok:
            self.window.hide_controller.cancel_countdown()
            self.window._show_message(error or "请求失败", is_error=True)
            self.window._fit_to_contents()
            return

        full_rows = []
        full_color_roles = []
        sort_values = []
        options = QuoteDisplayOptions(
            name_length=self.window.name_length,
            code_visible=self.window.code_visible,
            type_visible=self.window.type_visible,
            unit_mode=self.window.unit_mode,
        )
        for c, d in data.items():
            entry = self.window.watchlist.get(c) or {}
            code_info = self.window.codes_list.get(c, {})
            type_ = entry.get("type") or code_info.get("type")
            market = entry.get("market") or code_info.get("market") or ""
            display_code = entry.get("code") or code_info.get("code") or c
            row, color_roles, row_sort_values = format_quote(
                d, type_, display_code, market=market,
                cost=entry.get("cost"), options=options,
            )
            full_rows.append(row)
            full_color_roles.append(color_roles)
            sort_values.append(row_sort_values)

        self._last_full_rows = full_rows
        self._last_color_roles = full_color_roles
        self._last_sort_values = sort_values
        self._last_keys = list(data)
        self._latest_quotes = data
        self._latest_source = self.window.data_source

        if data:
            self.window._clear_message()
        else:
            self.window._show_message("请在设置面板中添加自选股", is_error=True)
        self.project_rows(full_rows, full_color_roles, sort_values)
        self.window.hide_controller.quotes_refreshed(data)
        self.quotes_updated.emit()

    def set_sort(self, header: str, order):
        if header not in SORTABLE_HEADERS:
            return
        visible_headers = metric_headers(self.window.visible_metrics)
        if header not in visible_headers:
            return
        order = Qt.SortOrder(order)
        if self.sort_header == header and self.sort_order == order:
            return
        self.sort_header = header
        self.sort_order = order
        self.float_page = self.taskbar_page = 0
        self.reproject()

    def clear_sort(self):
        if self.sort_header is None:
            return
        self.sort_header = None
        self.float_page = self.taskbar_page = 0
        self.window.table.horizontalHeader().setSortIndicatorShown(False)
        self.reproject()

    def header_clicked(self, section: int):
        header_name = self.window.model.headerData(
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
