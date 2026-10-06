"""设置加载、主题、重置和同步组框的编辑状态。"""

from unittest.mock import Mock, patch
import os
import tempfile
import time

from PySide6.QtCore import QPoint, QPointF, QSize, Qt
from PySide6.QtGui import QColor, QIcon, QPalette, QPixmap, QTextDocument, QKeySequence, QWheelEvent
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QScrollArea, QGroupBox, QSlider

from stockwidget.core.view_options import APPEARANCE_OPTION_KEYS
from stockwidget.platform.hotkeys import HotkeyResult
from stockwidget.platform.taskbar import TaskbarArea
from stockwidget.ui.controls.style import COLOR_SWATCH_SIZE, color_swatch_icon
from stockwidget.ui.floating.presenter import QuotePresenter
from stockwidget.ui.settings.dialog import SettingsDialog

from tests.support import SettingsTestCase, SignalRecorder


class SettingsDialogTests(SettingsTestCase):
    def test_taskbar_adjustments_do_not_restyle_pages_without_a_theme_change(self):
        dialog, window = self._make_dialog(taskbar_enabled=True, taskbar_sync_appearance=False)
        with patch.object(dialog, "_apply_theme_stylesheet", wraps=dialog._apply_theme_stylesheet) as restyle:
            for opacity in (90, 80, 70):
                dialog.taskbar_style_settings.opacity.setValue(opacity)
            dialog.taskbar_settings.rows.setValue(3)
            dialog.taskbar_settings.offset.setValue(10)
            self.qt_app.processEvents()
            restyle.assert_not_called()
            self.assertEqual(window.view_options.taskbar_opacity_pct, 70)
            dialog.ui.cmb_color_mode.setCurrentIndex(dialog.ui.cmb_color_mode.findData("dark"))
            self.assertEqual(restyle.call_count, 1)
            self.assertLess(dialog.palette().color(QPalette.Window).lightness(), 128)

    def test_taskbar_font_limit_tracks_rows_height_and_dpi_without_erasing_independent_preferences(self):
        with patch("stockwidget.ui.settings.groups.find_taskbar", return_value=TaskbarArea(1, 2560, 104, 2200, 192)) as find:
            dialog, window = self._make_dialog(taskbar_enabled=True, taskbar_sync_appearance=False,
                                              taskbar_rows=4, taskbar_font_size=20)
            style = dialog.taskbar_style_settings
            self.assertEqual((style.font_size.maximum(), style.font_size.value()), (7, 7))
            window.set_view_options(taskbar_rows=3)
            self.assertEqual((style.font_size.maximum(), style.font_size.value()), (10, 10))
            self.assertEqual(window.view_options.taskbar_font_size, 20)
            style.font_size.setValue(8)
            self.assertEqual(window.view_options.taskbar_font_size, 8)
            find.return_value = TaskbarArea(1, 1920, 48, 1700, 96)
            dialog._sync_view_settings()
            self.assertEqual(style.font_size.maximum(), 8)
            find.return_value = None
            dialog._sync_view_settings()
            self.assertEqual(style.font_size.maximum(), 8)

    def test_fixed_size_and_taskbar_category_only_on_windows(self):
        for platform in ("win32", "linux", "darwin"):
            with self.subTest(platform=platform), patch("stockwidget.ui.settings.dialog.sys.platform", platform):
                dialog, _ = self._make_dialog()
                dialog.resize(900, 700)
                self.assertEqual(dialog.size(), QSize(704, 496))
                dialog.resize(600, 350)
                self.assertEqual(dialog.size(), QSize(704, 496))
                self.assertEqual(dialog.ui.settings_navigation.item(5).isHidden(), platform != "win32")

    def test_inline_metric_options_preserve_configuration_and_apply_to_both_surfaces(self):
        dialog, window = self._make_dialog(name_length=3, unit_mode="cn")
        self.assertEqual(dialog.ui.cmb_name_length.currentData(), 3)
        dialog.ui.cb_name_visible.click()
        self.assertEqual(window.name_length, 0)
        self.assertFalse(dialog.ui.cmb_name_length.isEnabled())
        dialog.ui.cb_name_visible.click()
        self.assertEqual(window.name_length, 3)
        dialog.ui.cb_code_visible.click()
        dialog.ui.cb_type_visible.click()
        dialog.ui.cmb_unit_mode.setCurrentIndex(dialog.ui.cmb_unit_mode.findData("en"))
        self.assertEqual((window.code_visible, window.type_visible, window.unit_mode), (True, True, "en"))
        window.set_name_length(2)
        self.assertEqual(dialog.ui.cmb_name_length.currentData(), 2)
        self.assertEqual(window.current_config()["unit_mode"], "en")

    def test_wheel_scrolls_page_without_changing_slider_or_combo_values(self):
        dialog, _ = self._make_dialog(chart_enabled=True)
        dialog.show()
        for page, controls in ((dialog.ui.data, (dialog.ui.cmb_chart_display_mode,
                                               dialog.ui.cmb_name_length, dialog.ui.cmb_unit_mode)),
                               (dialog.ui.floating, (dialog.ui.cmb_font, dialog.ui.slider_font_size))):
            dialog.ui.settings_pages.setCurrentWidget(page)
            scroll = page.findChild(QScrollArea)
            for control in controls:
                self.qt_app.processEvents()
                scroll.ensureWidgetVisible(control)
                self.qt_app.processEvents()
                value = control.value() if isinstance(control, QSlider) else control.currentText()
                point = control.rect().center()
                before = scroll.verticalScrollBar().value()
                upwards = before == scroll.verticalScrollBar().maximum()
                event = QWheelEvent(QPointF(point), QPointF(control.mapToGlobal(point)), QPoint(),
                                    QPoint(0, 120 if upwards else -120), Qt.NoButton, Qt.NoModifier, Qt.ScrollUpdate, False)
                QApplication.sendEvent(control, event)
                self.assertEqual(control.value() if isinstance(control, QSlider) else control.currentText(), value)
                if upwards:
                    self.assertLess(scroll.verticalScrollBar().value(), before)
                else:
                    self.assertGreater(scroll.verticalScrollBar().value(), before)

    def test_manual_color_mode_overrides_system_and_can_return_to_system(self):
        dialog, window = self._make_dialog()
        for mode, dark in (("light", False), ("dark", True)):
            dialog.ui.cmb_color_mode.setCurrentIndex(dialog.ui.cmb_color_mode.findData(mode))
            self.assertEqual(window.view_options.color_mode, mode)
            self.assertEqual(dialog.palette().color(QPalette.Window).lightness() < 128, dark)
        dialog.ui.cmb_color_mode.setCurrentIndex(0)
        self.assertEqual(dialog.palette(), QApplication.palette())

    def test_hide_status_describes_the_active_conditions(self):
        dialog, window = self._make_dialog()
        cases = ((False, [], False, "程序未启用自动隐藏"),
                 (True, ["15:00", "16:00"], False, "程序将在15:00，16:00自动隐藏"),
                 (True, [], True, "程序将在无数据更新时自动隐藏"),
                 (True, ["15:00", "16:00"], True, "程序将在15:00，16:00及无数据更新时自动隐藏"))
        for enabled, times, automatic, text in cases:
            window.set_hide_options(hide_enabled=enabled, scheduled_hide_enabled=enabled,
                                    scheduled_hide_times=times, auto_hide_enabled=automatic)
            self.assertEqual(dialog.ui.label_hide_status.text(), text)

    def test_about_status_text_follows_light_and_dark_palettes(self):
        dialog, _ = self._make_dialog()
        original = self.qt_app.palette()
        try:
            for foreground, background in (("#eeeeee", "#222222"), ("#222222", "#eeeeee")):
                palette = QPalette(original)
                palette.setColor(QPalette.WindowText, QColor(foreground))
                palette.setColor(QPalette.Window, QColor(background))
                self.qt_app.setPalette(palette)
                self.qt_app.processEvents()
                dialog.refresh_about()
                dialog.refresh_data_state()
                for label in (dialog.ui.label_version_state, dialog.ui.label_data_state,
                              dialog.ui.label_about_info):
                    self.assertEqual(label.palette().color(QPalette.WindowText), QColor(foreground))
        finally:
            self.qt_app.setPalette(original)

    def test_about_has_blank_line_after_copyright_and_fits_label(self):
        dialog, _ = self._make_dialog()
        dialog.ui.settings_pages.setCurrentWidget(dialog.ui.about)
        dialog.show()
        self.qt_app.processEvents()
        label = dialog.ui.label_about_info
        document = QTextDocument()
        document.setDefaultFont(label.font())
        document.setDocumentMargin(0)
        document.setHtml(label.text())
        document.setTextWidth(label.contentsRect().width())
        self.assertIn("Copyright 2026 sbr0574\n\n仓库地址", document.toPlainText())
        self.assertLessEqual(document.size().height(), label.contentsRect().height())

    def test_linux_about_uses_smaller_font_and_fits_scroll_content(self):
        with patch("stockwidget.ui.settings.dialog.sys.platform", "linux"):
            dialog, _ = self._make_dialog()
        dialog.ui.settings_pages.setCurrentWidget(dialog.ui.about)
        dialog.show()
        self.qt_app.processEvents()
        label = dialog.ui.label_about_info
        self.assertEqual(label.font().pixelSize(), 12)
        self.assertTrue(dialog.ui.about.findChildren(QScrollArea))
        self.assertIn("问题反馈", label.text())
        document = QTextDocument()
        document.setDefaultFont(label.font())
        document.setDocumentMargin(0)
        document.setHtml(label.text())
        document.setTextWidth(label.contentsRect().width())
        self.assertLessEqual(document.size().height(), label.contentsRect().height())
        self.assertLess(label.geometry().bottom(), dialog.ui.btn_open_cache_dir.y())

    def test_linux_fonts_override_desktop_styles_for_controls_and_rich_text(self):
        original_stylesheet = self.qt_app.styleSheet()
        # 模拟桌面主题对子控件指定大字号：仅在父窗口 setFont 不会覆盖它。
        self.qt_app.setStyleSheet("QWidget { font-size: 20px; }")
        try:
            with patch("stockwidget.ui.settings.dialog.sys.platform", "linux"):
                dialog, _ = self._make_dialog({"sh600519": {"checked": True}})
                dialog.show()
                self.qt_app.processEvents()
                editor = self._start_code_editor(dialog)
                controls = (
                    dialog.ui.btn_add, dialog.ui.cb_auto_start, dialog.ui.rb_sina,
                    dialog.ui.label_source, dialog.ui.sb_interval,
                    dialog.ui.cmb_font, dialog.ui.keyseq_hide,
                    dialog.ui.list_codes, dialog.ui.list_codes.horizontalHeader(),
                    dialog.metric_pool.displayed_pool, editor,
                    dialog.ui.cmb_name_length, dialog.ui.cb_code_visible,
                    dialog.watchlist_editor.add_code_panel.search_input,
                    dialog.watchlist_editor.add_code_panel.result_list,
                    dialog.watchlist_editor.add_code_panel.next_button,
                )
                for _ in range(2):
                    # 切换主题后，以及延迟创建的表格编辑器，也必须使用同一字号。
                    dialog._apply_theme_stylesheet()
                    self.qt_app.processEvents()
                    for control in controls:
                        with self.subTest(control=control.objectName() or type(control).__name__):
                            self.assertEqual(control.font().pixelSize(), 12)
                    for control in (dialog.ui.label_about_info, dialog.ui.label_version_state,
                                    dialog.ui.label_data_state, dialog.ui.btn_open_cache_dir):
                        with self.subTest(about_control=control.objectName()):
                            self.assertEqual(control.font().pixelSize(), 12)
                label = dialog.ui.label_about_info
                dialog.ui.settings_pages.setCurrentWidget(dialog.ui.about)
                self.qt_app.processEvents()
                document = QTextDocument()
                document.setDefaultFont(label.font())
                document.setDocumentMargin(0)
                document.setHtml(label.text())
                document.setTextWidth(label.contentsRect().width())
                self.assertLessEqual(document.size().height(), label.height())
                self.assertTrue(dialog.ui.about.findChildren(QScrollArea))
                # 具有独立设计字号的图形按钮、空列表提示仍保留原样。
                self.assertEqual(dialog.ui.btn_icon_custom.font().pixelSize(), 22)
                self.assertEqual(dialog.watchlist_editor.empty_watchlist_hint.font().pixelSize(), 18)
        finally:
            self.qt_app.setStyleSheet(original_stylesheet)

    def test_linux_font_rules_do_not_override_other_platforms_or_float_window(self):
        original_stylesheet = self.qt_app.styleSheet()
        self.qt_app.setStyleSheet("QWidget { font-size: 20px; }")
        try:
            for system in ("win32", "darwin", "linux"):
                with self.subTest(system=system), patch(
                    "stockwidget.ui.settings.dialog.sys.platform", system
                ):
                    dialog, window = self._make_dialog()
                    dialog.show()
                    self.qt_app.processEvents()
                    self.assertEqual(window.table.font().pixelSize(), 20)
                    if system != "linux":
                        self.assertEqual(dialog.ui.btn_add.font().pixelSize(), 20)
                        self.assertEqual(dialog.ui.label_about_info.font().pixelSize(), 20)
        finally:
            self.qt_app.setStyleSheet(original_stylesheet)

    def test_failed_hotkeys_stay_editable_and_show_inline_result(self):
        with patch("stockwidget.ui.settings.dialog.hotkeys_supported", return_value=True), patch(
            "stockwidget.ui.settings.dialog.click_through_supported", return_value=True
        ):
            dialog, window = self._make_dialog()
        for key, checkbox, editor in dialog._hotkey_rows:
            with self.subTest(key=key), patch.object(window._hotkeys, "register") as register, patch(
                "stockwidget.ui.settings.dialog.QMessageBox.warning"
            ) as warning:
                register.return_value = HotkeyResult(False, "conflict")
                checkbox.setChecked(True)
                status = dialog._hotkey_status[key]
                self.assertTrue(checkbox.isChecked())
                self.assertTrue(editor.isEnabled())
                self.assertFalse(status.isHidden())
                self.assertFalse(status.active)
                self.assertIn("占用", status.toolTip())
                register.return_value = HotkeyResult(True)
                editor.setKeySequence(QKeySequence("Ctrl+Alt+X"))
                editor.editingFinished.emit()
                self.assertEqual(getattr(window, key), "Ctrl+Alt+X")
                self.assertTrue(status.active)
                checkbox.setChecked(False)
                self.assertTrue(status.isHidden())
                self.assertFalse(editor.isEnabled())
                warning.assert_not_called()

    def test_reset_appearance_preserves_settings_and_watchlist(self):
        app = self._make_icon_app(choice="dark")
        dialog, window = self._make_dialog(
            {"sh600519": {"checked": False, "cost": 100}}, app=app,
            fg="#123456", bg={"r": 30, "g": 40, "b": 50, "a": 100},
            up_color="#123456", down_color="#abcdef", neutral_color="#654321",
            opacity_pct=45, font_family="Arial", font_size=14, line_extra_px=9,
            unicolor=False, header_visible=True, grid_visible=True,
            refresh_seconds=10, data_source="eastmoney", visible_metrics=["price"],
            hotkey="Ctrl+Alt+G",
            taskbar_sync_appearance=False, taskbar_font_family="Arial", taskbar_font_size=17,
            taskbar_color="#aabbcc", color_mode="dark",
        )
        _, defaults = self._make_dialog()
        before = window.current_config()
        appearance_keys = (
            "fg", "bg", "up_color", "down_color", "neutral_color", "opacity_pct",
            "font_family", "font_size", "line_extra_px", "unicolor", "header_visible", "grid_visible",
        ) + APPEARANCE_OPTION_KEYS
        dialog.ui.btn_reset_appearance.click()
        expected = {**before, **{key: defaults.current_config()[key] for key in appearance_keys}}
        self.assertEqual(window.current_config(), expected)
        self.assertEqual(dialog.ui.slider_font_size.value(), 10)
        self.assertEqual(dialog.ui.slider_bg_alpha.value(), 75)
        self.assertEqual(dialog.ui.slider_all_alpha.value(), 90)
        self.assertTrue(dialog.ui.btn_icon_default.isChecked())
        self.assertEqual(app._icon_choice, "default")
        self.assertEqual(app._custom_icon_path, "")
        app.save_now.assert_called()

    def test_reset_settings_preserves_appearance_and_watchlist(self):
        app = self._make_icon_app(choice="dark")
        dialog, window = self._make_dialog(
            {"sh600519": {"checked": False, "cost": 100}}, app=app,
            fg="#123456", opacity_pct=45, font_size=14, header_visible=True,
            refresh_seconds=10, data_source="eastmoney", visible_metrics=["volume", "price"],
            code_visible=True, type_visible=True, name_length=3, unit_mode="en",
            start_on_boot=True, hotkey="Ctrl+Alt+G", hotkey_click_through="Ctrl+Alt+D",
            taskbar_rows=3, taskbar_page_mode="auto", taskbar_metrics=["price"],
            taskbar_color="#aabbcc", float_max_rows=4, color_mode="dark", hide_tray_icon=True,
        )
        _, defaults = self._make_dialog()
        before = window.current_config()
        # 模拟已经启用的系统功能，避免测试实际注册全局快捷键或改变鼠标穿透。
        window.hotkey_enabled = window.hotkey_click_through_enabled = True
        window.force_top = window.click_through = True
        window._keep_top_timer.start()
        with (
            patch.object(window._hotkeys, "unregister_all") as unregister,
            patch("stockwidget.ui.floating.widget.apply_click_through") as click_through,
        ):
            dialog.ui.btn_reset_settings.click()
        settings_keys = (
            "refresh_seconds", "data_source", "visible_metrics", "code_visible", "type_visible",
            "name_length", "unit_mode", "start_on_boot", "force_top", "click_through",
            "hotkey", "hotkey_click_through", "hotkey_enabled", "hotkey_click_through_enabled",
        )
        expected = {**before, **{key: defaults.current_config()[key] for key in settings_keys}}
        expected.update({key: defaults.current_config()[key] for key in defaults.view_options.to_config()
                         if key not in APPEARANCE_OPTION_KEYS})
        from stockwidget.core.quote_presentation import legacy_visibility
        expected.update(legacy_visibility(defaults.visible_metrics))
        self.assertEqual(window.current_config(), expected)
        self.assertEqual(window.timer.interval(), 2000)
        self.assertFalse(window._keep_top_timer.isActive())
        self.assertFalse(dialog.ui.keyseq_hide.isEnabled())
        self.assertFalse(dialog.ui.cb_auto_start.isChecked())
        self.assertTrue(dialog.ui.rb_sina.isChecked())
        self.assertEqual(dialog.metric_pool.visible_metrics, list(defaults.visible_metrics))
        self.assertEqual(app._icon_choice, "dark")
        unregister.assert_called_once()
        click_through.assert_called_once_with(window, False)
        app.set_start_on_boot.assert_called_once_with(False)
        app.save_now.assert_called()

    def test_reloading_settings_does_not_save_or_round_opacity(self):
        dialog, window = self._make_dialog(bg={"r": 0, "g": 0, "b": 0, "a": 127})
        changes = Mock()
        window.set_on_change(changes)
        for _ in range(2):
            dialog._load_settings()
        self.assertEqual(window.bg.alpha(), 127)
        changes.assert_not_called()

    def test_float_topmost_checkbox_controls_force_top_and_stays_synced(self):
        with patch("stockwidget.ui.settings.dialog.force_top_supported", return_value=True), patch(
            "stockwidget.ui.floating.widget.force_top_supported", return_value=True
        ):
            dialog, window = self._make_dialog(force_top=True)
            self.assertTrue(dialog.ui.cb_float_on_top.isChecked())
            self.assertTrue(dialog.ui.cb_force_top.isChecked())
            self.assertTrue(dialog.ui.cb_force_top.isEnabled())
            dialog.show()
            dialog.ui.settings_pages.setCurrentWidget(dialog.ui.floating)
            self.qt_app.processEvents()
            dialog.ui.floating_scroll.ensureWidgetVisible(dialog.ui.gb_fcn)
            QTest.mouseClick(dialog.ui.cb_float_on_top, Qt.LeftButton)
            self.assertFalse(window.float_on_top)
            self.assertFalse(window.force_top)
            self.assertFalse(dialog.ui.cb_force_top.isChecked())
            self.assertFalse(dialog.ui.cb_force_top.isEnabled())
            QTest.mouseClick(dialog.ui.cb_force_top, Qt.LeftButton)
            self.assertFalse(window.force_top)
            dialog._load_settings()
            dialog._apply_theme_stylesheet()
            self.assertFalse(dialog.ui.cb_force_top.isEnabled())
            QTest.mouseClick(dialog.ui.cb_float_on_top, Qt.LeftButton)
            self.assertTrue(window.float_on_top)
            self.assertTrue(dialog.ui.cb_force_top.isEnabled())
            self.assertFalse(dialog.ui.cb_force_top.isChecked())
            QTest.mouseClick(dialog.ui.cb_force_top, Qt.LeftButton)
            self.assertTrue(window.force_top)
            window.set_float_on_top(False)
            self.assertFalse(dialog.ui.cb_float_on_top.isChecked())
            self.assertFalse(dialog.ui.cb_force_top.isChecked())
            self.assertFalse(dialog.ui.cb_force_top.isEnabled())
            dialog.ui.btn_reset_settings.click()
            self.assertTrue(dialog.ui.cb_float_on_top.isChecked())
            self.assertTrue(dialog.ui.cb_force_top.isEnabled())
            self.assertFalse(dialog.ui.cb_force_top.isChecked())

    def test_disabled_float_topmost_loads_with_force_top_unchecked_and_disabled(self):
        dialog, window = self._make_dialog(float_on_top=False, force_top=True)
        self.assertFalse(window.force_top)
        self.assertFalse(dialog.ui.cb_float_on_top.isChecked())
        self.assertFalse(dialog.ui.cb_force_top.isChecked())
        self.assertFalse(dialog.ui.cb_force_top.isEnabled())

    def test_float_topmost_does_not_enable_unsupported_force_top(self):
        with patch("stockwidget.ui.settings.dialog.force_top_supported", return_value=False):
            dialog, window = self._make_dialog()
            for enabled in (False, True):
                window.set_float_on_top(enabled)
                self.assertFalse(dialog.ui.cb_force_top.isEnabled())

    def _make_icon_app(self, choice="default", custom_path=""):
        app = Mock()
        app._icon_choice = choice
        app._custom_icon_path = custom_path
        app.app_version = "1.0.0"
        app._has_update = False
        app._latest_version = None
        app.code_data_state.return_value = ("cached", "")
        app.code_data_error.return_value = ""

        def set_app_icon(icon_choice):
            app._icon_choice = icon_choice

        def set_custom_icon(path):
            icon = QIcon(path)
            if icon.isNull():
                return False
            app._custom_icon_path = os.path.abspath(path)
            app._icon_choice = "custom"
            return True

        def clear_custom_icon():
            app._custom_icon_path = ""
            app._icon_choice = "default"

        app.set_app_icon.side_effect = set_app_icon
        app.set_custom_icon.side_effect = set_custom_icon
        app.clear_custom_icon.side_effect = clear_custom_icon
        return app

    @patch.object(QuotePresenter, "refresh")
    def test_data_state_tooltip_explains_failed_download_and_clears_on_success(self, refresh):
        app = self._make_icon_app()
        app.code_data_error.return_value = "stock_hk.json 更新失败；Gitee：HTTP 451"
        dialog, _window = self._make_dialog(app=app)
        self.assertIn("stock_hk.json", dialog.ui.label_data_state.toolTip())
        self.assertIn("30 分钟", dialog.ui.label_data_state.toolTip())
        app.code_data_state.return_value = ("current", "2026-09-09")
        app.code_data_error.return_value = ""
        dialog.refresh_data_state()
        self.assertEqual(dialog.ui.label_data_state.toolTip(), "")
        self.assertIn("最新", dialog.ui.label_data_state.text())

    def test_color_swatch_renders_at_device_pixel_ratio(self):
        icon = color_swatch_icon(QColor("#123456"), 2.0)
        pixmap = icon.pixmap(QSize(COLOR_SWATCH_SIZE, COLOR_SWATCH_SIZE), 2.0)

        self.assertEqual(pixmap.devicePixelRatio(), 2.0)
        self.assertEqual(pixmap.width(), COLOR_SWATCH_SIZE * 2)
        self.assertEqual(pixmap.height(), COLOR_SWATCH_SIZE * 2)
        image = pixmap.toImage()
        center = image.pixelColor(image.width() // 2, image.height() // 2)
        self.assertEqual(center.name(), "#123456")
        self.assertEqual(image.pixelColor(0, 0).alpha(), 0)

    def test_shared_paging_and_taskbar_style_groups_preserve_control_behavior(self):
        dialog, window = self._make_dialog()
        floating, taskbar, style = (dialog.float_row_settings,
                                   dialog.taskbar_paging_settings, dialog.taskbar_style_settings)
        row_limit = dialog.float_row_settings
        self.assertFalse(floating.auto.isEnabled())
        row_limit.setChecked(True)
        dialog.taskbar_settings.enabled.setChecked(True)
        dialog.taskbar_settings.metrics.sync_toggle.setChecked(False)
        row_limit.rows.setValue(3)
        floating.auto.setChecked(True)
        floating.interval.setValue(10)
        taskbar.sync_toggle.setChecked(False)
        dialog.taskbar_settings.rows.setValue(1)
        taskbar.mode.setCurrentIndex(taskbar.mode.findData("manual"))
        self.assertEqual(window.view_options.float_max_rows, 3)
        self.assertEqual(window.view_options.float_page_mode, "auto")
        self.assertEqual(window.view_options.float_page_interval, 10)
        self.assertEqual(window.view_options.taskbar_rows, 1)
        self.assertTrue(floating.interval.isEnabled())
        self.assertFalse(taskbar.interval.isEnabled())
        self.assertEqual(dialog.taskbar_settings.rows.maximum(), 4)
        self.assertEqual(dialog.taskbar_settings.rows.minimum(), 1)
        self.assertFalse(style.font_size.isEnabled())
        style.sync_toggle.setChecked(False)
        style.font_size.setValue(14)
        self.assertEqual(window.view_options.taskbar_font_size, 14)
        self.assertEqual(window.font.pointSize(), 10)
        style.sync_toggle.setChecked(False)
        self.assertTrue(style.color.isEnabled())
        window.set_view_options(taskbar_color="#aabbcc", taskbar_metrics=["price", "volume"])
        self.assertEqual(style.color.text().strip(), "#AABBCC")
        self.assertIn("#AABBCC", style.color.toolTip())
        self.assertIsInstance(style.font_size, QSlider)
        self.assertEqual(style.font_size_label.text(), "14 pt")
        self.assertEqual(dialog.taskbar_settings.metric_pool.visible_metrics, ["price", "volume"])
        row_limit.setChecked(False)
        self.assertFalse(row_limit.rows.isEnabled())
        self.assertFalse(floating.auto.isEnabled())
        self.assertFalse(floating.interval.isEnabled())

    def test_row_limit_preserves_saved_paging_mode_and_interval_until_auto_is_changed(self):
        for mode in ("first", "manual", "auto"):
            with self.subTest(mode=mode):
                dialog, window = self._make_dialog(float_paging_enabled=True,
                                                   float_page_mode=mode, float_page_interval=17)
                rows = dialog.float_row_settings
                self.assertEqual(rows.paging.isChecked(), mode != "first")
                self.assertEqual(rows.auto.isChecked(), mode == "auto")
                self.assertEqual(rows.interval.isEnabled(), mode == "auto")
                self.assertEqual(rows.interval.value(), 17)
                rows.setChecked(False)
                dialog._load_settings()
                self.assertFalse(rows.auto.isEnabled())
                self.assertEqual(window.view_options.float_page_mode, mode)
                self.assertEqual(window.view_options.float_page_interval, 17)
                rows.setChecked(True)
                rows.paging.setChecked(True)
                rows.auto.setChecked(True)
                self.assertEqual(window.view_options.float_page_mode, "auto")
                self.assertTrue(rows.interval.isEnabled())
                rows.auto.setChecked(False)
                self.assertEqual(window.view_options.float_page_mode, "manual")
                self.assertFalse(rows.interval.isEnabled())
                self.assertEqual(window.view_options.float_page_interval, 17)

    def test_nested_paging_switch_controls_first_manual_auto_and_synced_taskbar(self):
        dialog, window = self._make_dialog(float_paging_enabled=True, float_page_mode="first")
        dialog.ui.settings_pages.setCurrentWidget(dialog.ui.floating)
        dialog.show()
        self.qt_app.processEvents()
        rows = dialog.float_row_settings
        paging = rows.paging
        dialog.ui.floating_scroll.ensureWidgetVisible(paging)
        self.assertTrue(rows.rows.isEnabled())
        self.assertFalse(rows.auto.isEnabled())

        def click_paging():
            dialog.ui.floating_scroll.ensureWidgetVisible(paging)
            QTest.mouseClick(paging, Qt.LeftButton)
            self.qt_app.processEvents()

        click_paging()
        self.assertTrue(rows.auto.isEnabled())
        self.assertFalse(rows.interval.isEnabled())
        self.assertEqual(window.view_options.page_settings("float")[0], "manual")
        rows.auto.click()
        rows.interval.setValue(60)
        self.assertEqual(window.view_options.page_settings("float"), ("auto", 60))
        self.assertEqual(window.view_options.page_settings("taskbar"), ("auto", 60))
        click_paging()
        self.assertTrue(rows.rows.isEnabled())
        self.assertFalse(rows.auto.isEnabled())
        self.assertFalse(rows.interval.isEnabled())
        self.assertEqual(window.view_options.page_settings("float"), ("first", 60))
        self.assertEqual(window.view_options.page_settings("taskbar"), ("first", 60))
        dialog._apply_theme_stylesheet()
        self.qt_app.processEvents()
        self.assertFalse(rows.auto.isEnabled())
        click_paging()
        self.assertTrue(rows.auto.isEnabled())
        self.assertEqual(window.view_options.page_settings("float"), ("manual", 60))

    def test_taskbar_switches_hide_independent_settings_and_preserve_them_when_reopened(self):
        dialog, window = self._make_dialog(taskbar_enabled=True, taskbar_sync_metrics=False,
                                           taskbar_sync_appearance=False, taskbar_sync_paging=False,
                                           taskbar_sync_split=False)
        dialog.ui.settings_pages.setCurrentWidget(dialog.ui.taskbar)
        dialog.show()
        self.qt_app.processEvents()
        groups = (dialog.taskbar_settings.metrics, dialog.taskbar_style_settings,
                  dialog.taskbar_paging_settings, dialog.taskbar_split_settings)
        for group in groups:
            with self.subTest(group=group.title()):
                self.assertTrue(group.body.isVisible())
                dialog.ui.taskbar_scroll.ensureWidgetVisible(group.sync_toggle)
                QTest.mouseClick(group.sync_toggle, Qt.LeftButton)
                self.qt_app.processEvents()
                self.assertTrue(group.body.isHidden())
                self.assertTrue(dialog.taskbar_settings.offset.isVisible())
                if group is dialog.taskbar_style_settings:
                    self.assertFalse(dialog.taskbar_settings.rows.isVisible())
                QTest.mouseClick(group.sync_toggle, Qt.LeftButton)
                self.qt_app.processEvents()
                self.assertTrue(group.body.isVisible())
        style = dialog.taskbar_style_settings
        style.opacity.setValue(43)
        self.assertIsInstance(style.opacity, QSlider)
        self.assertEqual(style.opacity_label.text(), "43%")
        style.unicolor.setChecked(False)
        dialog.taskbar_settings.rows.setValue(4)
        dialog.taskbar_settings.offset.setValue(35)
        self.assertEqual(window.view_options.taskbar_opacity_pct, 43)
        self.assertFalse(window.view_options.taskbar_unicolor)
        self.assertEqual(window.view_options.taskbar_rows, 4)
        dialog.taskbar_settings.enabled.setChecked(False)
        self.assertTrue(dialog.taskbar_settings.body.isHidden())
        self.assertTrue(all(not group.isVisible() for group in groups))
        dialog.taskbar_settings.enabled.setChecked(True)
        self.assertTrue(all(group.body.isVisible() for group in groups))
        self.assertEqual((style.opacity.value(), dialog.taskbar_settings.rows.value(),
                          dialog.taskbar_settings.offset.value()), (43, 4, 35))
        self.assertFalse(style.unicolor.isChecked())

    def test_taskbar_automatic_color_locks_manual_controls_and_restores_saved_preferences(self):
        dialog, window = self._make_dialog(taskbar_enabled=True, taskbar_sync_appearance=False,
                                           taskbar_color="#123456", taskbar_unicolor=False)
        style = dialog.taskbar_style_settings
        style.auto_color.click()
        self.assertTrue(window.view_options.taskbar_auto_color)
        self.assertFalse(style.color.isEnabled())
        self.assertFalse(style.unicolor.isEnabled())
        self.assertEqual(window.view_options.taskbar_color, "#123456")
        self.assertFalse(window.view_options.taskbar_unicolor)
        style.sync_toggle.click()
        dialog._apply_theme_stylesheet()
        self.assertTrue(style.body.isHidden())
        self.assertTrue(window.view_options.taskbar_auto_color)
        style.sync_toggle.click()
        style.auto_color.click()
        self.assertTrue(style.color.isEnabled())
        self.assertTrue(style.unicolor.isEnabled())
        self.assertEqual(style.color.text().strip(), "#123456")
        self.assertFalse(style.unicolor.isChecked())

    def test_split_controls_use_native_groups_and_preserve_independent_taskbar_settings(self):
        dialog, window = self._make_dialog(taskbar_enabled=True, taskbar_sync_split=False,
                                           taskbar_split_enabled=True, taskbar_split_separator=False)
        floating = dialog.float_split_settings
        taskbar = dialog.taskbar_split_settings
        for group in (dialog.float_row_settings, floating,
                      dialog.taskbar_settings.metrics,
                      dialog.taskbar_style_settings, dialog.taskbar_paging_settings, taskbar):
            self.assertIs(type(group), QGroupBox)
        self.assertFalse(floating.isChecked())
        self.assertFalse(floating.separator.isEnabled())
        floating.setChecked(True)
        self.assertTrue(window.view_options.float_split_enabled)
        floating.separator.setChecked(True)
        self.assertTrue(taskbar.enabled.isEnabled())
        self.assertTrue(taskbar.enabled.isChecked())
        self.assertFalse(taskbar.separator.isChecked())
        taskbar.sync_toggle.setChecked(True)
        self.assertFalse(taskbar.enabled.isEnabled())
        self.assertTrue(taskbar.separator.isChecked())
        floating.setChecked(False)
        self.assertFalse(taskbar.enabled.isChecked())
        for _ in range(2):
            dialog._apply_theme_stylesheet()
            self.qt_app.processEvents()
            self.assertFalse(taskbar.body.isEnabled())
        dialog.taskbar_settings.enabled.setChecked(False)
        dialog.taskbar_settings.enabled.setChecked(True)
        self.assertFalse(taskbar.body.isEnabled())
        taskbar.sync_toggle.setChecked(False)
        self.assertTrue(taskbar.enabled.isEnabled())
        self.assertTrue(taskbar.enabled.isChecked())
        self.assertFalse(taskbar.separator.isChecked())
        taskbar.enabled.setChecked(False)
        self.assertFalse(taskbar.separator.isEnabled())
        taskbar.enabled.setChecked(True)
        self.assertTrue(taskbar.separator.isEnabled())
        dialog.show()
        for page, group in ((dialog.ui.floating, floating), (dialog.ui.taskbar, taskbar)):
            dialog.ui.settings_pages.setCurrentWidget(page)
            self.qt_app.processEvents()
            scroll = page.findChild(QScrollArea)
            scroll.ensureWidgetVisible(group)
            viewport = scroll.viewport()
            bounds = group.rect().translated(group.mapTo(viewport, QPoint()))
            self.assertTrue(viewport.rect().contains(bounds))

    def test_sidebar_navigation_preserves_edits_and_reaches_scrolled_settings(self):
        dialog, window = self._make_dialog()
        dialog.resize(640, 400)
        dialog.show()
        self.qt_app.processEvents()
        navigation = dialog.ui.settings_navigation
        self.assertTrue(dialog.ui.watchlist.isVisible())
        pages = (dialog.ui.watchlist, dialog.ui.data, dialog.ui.general, dialog.ui.shortcuts,
                 dialog.ui.floating, dialog.ui.taskbar, dialog.ui.about)
        for index, page in enumerate(pages):
            item = navigation.item(index)
            QTest.mouseClick(navigation.viewport(), Qt.LeftButton,
                             pos=navigation.visualItemRect(item).center())
            self.qt_app.processEvents()
            self.assertTrue(page.isVisible())
            self.assertFalse(item.icon().isNull())
            self.assertEqual(dialog.ui.settings_pages.currentIndex(), index)
        navigation.setCurrentRow(1)
        dialog.ui.rb_em.click()
        dialog.ui.cb_head.setChecked(True)
        navigation.setCurrentRow(4)
        self.qt_app.processEvents()
        scroll = dialog.ui.floating_scroll
        self.assertGreater(scroll.verticalScrollBar().maximum(), 0)
        rows = dialog.float_row_settings
        scroll.ensureWidgetVisible(rows)
        rows.setChecked(True)
        rows.rows.setValue(20)
        rows.paging.setChecked(True)
        rows.auto.setChecked(True)
        rows.interval.setValue(60)
        offset = scroll.verticalScrollBar().value()
        navigation.setCurrentRow(0)
        navigation.setFocus()
        QTest.keyClick(navigation, Qt.Key_Down)
        self.assertTrue(dialog.ui.data.isVisible())
        self.assertEqual(window.data_source, "eastmoney")
        self.assertTrue(window.header_visible)
        navigation.setCurrentRow(4)
        self.assertEqual(scroll.verticalScrollBar().value(), offset)
        self.assertEqual(window.view_options.float_max_rows, 20)
        self.assertEqual(window.view_options.page_settings("float"), ("auto", 60))

    def test_header_switch_supports_the_whole_track_keyboard_and_disabled_state(self):
        dialog, window = self._make_dialog(header_visible=False)
        dialog.ui.settings_pages.setCurrentWidget(dialog.ui.data)
        dialog.show()
        self.qt_app.processEvents()
        switch = dialog.ui.cb_head
        dialog.ui.data_scroll.ensureWidgetVisible(switch)
        point = QPoint(switch.width() - 3, switch.height() // 2)
        QTest.mouseClick(switch, Qt.LeftButton, pos=point)
        self.assertTrue(window.header_visible)
        switch.setFocus()
        QTest.keyClick(switch, Qt.Key_Space)
        self.assertFalse(window.header_visible)
        dialog._apply_theme_stylesheet()
        switch.setEnabled(False)
        QTest.mouseClick(switch, Qt.LeftButton, pos=point)
        QTest.keyClick(switch, Qt.Key_Space)
        self.assertFalse(window.header_visible)
        switch.setEnabled(True)
        switch.setFocus()
        QTest.keyClick(switch, Qt.Key_Space)
        self.assertTrue(window.header_visible)

    def test_custom_icon_can_be_selected_and_deleted_from_hover_action(self):
        app = self._make_icon_app()
        dialog, _window = self._make_dialog(app=app)

        with tempfile.TemporaryDirectory() as temp_dir:
            icon_path = os.path.join(temp_dir, "custom.png")
            pixmap = QPixmap(24, 24)
            pixmap.fill(QColor("#3578e5"))
            self.assertTrue(pixmap.save(icon_path))

            with patch(
                "stockwidget.ui.settings.dialog.QFileDialog.getOpenFileName",
                return_value=(icon_path, ""),
            ):
                dialog.ui.btn_icon_custom.click()

            self.assertEqual(app._icon_choice, "custom")
            self.assertEqual(app._custom_icon_path, os.path.abspath(icon_path))
            self.assertTrue(dialog.ui.btn_icon_custom.has_custom_icon())
            self.assertEqual(dialog.ui.btn_icon_custom.text(), "")
            self.assertTrue(dialog.ui.btn_icon_custom.isChecked())
            app.set_custom_icon.assert_called_once_with(icon_path)
            app.save_now.assert_called_once_with()

            dialog.ui.settings_pages.setCurrentWidget(dialog.ui.general)
            dialog.show()
            self.qt_app.processEvents()
            QTest.mouseMove(dialog.ui.btn_icon_custom, QPoint(35, 5))
            # Windows delivers the native hover event asynchronously.
            deadline = time.monotonic() + 1
            while not dialog.ui.btn_icon_custom.underMouse() and time.monotonic() < deadline:
                QTest.qWait(10)
            self.assertTrue(dialog.ui.btn_icon_custom.underMouse())

            QTest.mouseClick(
                dialog.ui.btn_icon_custom,
                Qt.MouseButton.LeftButton,
                pos=QPoint(35, 5),
            )

        self.assertEqual(app._icon_choice, "default")
        self.assertEqual(app._custom_icon_path, "")
        self.assertFalse(dialog.ui.btn_icon_custom.has_custom_icon())
        self.assertEqual(dialog.ui.btn_icon_custom.text(), "+")
        self.assertTrue(dialog.ui.btn_icon_default.isChecked())
        app.clear_custom_icon.assert_called_once_with()
        self.assertEqual(app.save_now.call_count, 2)

    def test_cancelling_custom_icon_picker_restores_previous_choice(self):
        app = self._make_icon_app(choice="dark")
        dialog, _window = self._make_dialog(app=app)

        with patch(
            "stockwidget.ui.settings.dialog.QFileDialog.getOpenFileName",
            return_value=("", ""),
        ):
            dialog.ui.btn_icon_custom.click()

        self.assertTrue(dialog.ui.btn_icon_dark.isChecked())
        self.assertEqual(app._icon_choice, "dark")
        app.set_custom_icon.assert_not_called()
        app.save_now.assert_not_called()

    def test_manual_update_check_shows_result(self):
        dialog, _window = self._make_dialog()

        with patch(
            "stockwidget.ui.settings.dialog.QMessageBox.information"
        ) as info:
            dialog._on_update_check_finished((True, "9.9.9"))
            info.assert_called_once()

        with patch(
            "stockwidget.ui.settings.dialog.QMessageBox.information"
        ) as info:
            dialog._on_update_check_finished((False, "1.4.1"))
            info.assert_called_once()

        with patch(
            "stockwidget.ui.settings.dialog.QMessageBox.warning"
        ) as warn:
            dialog._on_update_check_finished((False, None))
            warn.assert_called_once()

    def test_open_cache_dir_button_opens_folder(self):
        from stockwidget.core.config_store import config_paths

        dialog, _window = self._make_dialog()
        with patch(
            "stockwidget.ui.settings.dialog.QDesktopServices.openUrl"
        ) as open_url:
            dialog._open_cache_dir()
            open_url.assert_called_once()
            url = open_url.call_args[0][0]
            self.assertTrue(url.isLocalFile())
            self.assertEqual(os.path.normpath(url.toLocalFile()), os.path.normpath(config_paths()))

    def test_unicolor_defaults_on_and_controls_direction_colors(self):
        dialog, window = self._make_dialog()

        self.assertTrue(window.unicolor)
        self.assertEqual(window.up_color.name(), "#dd2100")
        self.assertEqual(window.down_color.name(), "#019933")
        self.assertEqual(window.neutral_color.name(), "#494949")
        self.assertTrue(dialog.ui.cb_unicolor.isChecked())
        self.assertEqual(
            dialog.ui.cb_unicolor.minimumWidth(), dialog.ui.cb_unicolor.maximumWidth()
        )
        self.assertGreaterEqual(
            dialog.ui.cb_unicolor.minimumWidth(), dialog.ui.cb_unicolor.sizeHint().width()
        )
        self.assertTrue(dialog.ui.btn_bg_color.isEnabled())
        self.assertTrue(dialog.ui.btn_fg_color.isEnabled())
        self.assertFalse(dialog.ui.btn_up_color.isEnabled())
        self.assertFalse(dialog.ui.btn_down_color.isEnabled())
        self.assertFalse(dialog.ui.btn_neutral_color.isEnabled())

        color_buttons = (
            (dialog.ui.btn_fg_color, "文字", window.fg),
            (dialog.ui.btn_bg_color, "背景", window.bg),
            (dialog.ui.btn_up_color, "上涨", window.up_color),
            (dialog.ui.btn_down_color, "下跌", window.down_color),
            (dialog.ui.btn_neutral_color, "中性", window.neutral_color),
        )
        for button, text, color in color_buttons:
            self.assertTrue(button.isFlat())
            self.assertEqual(button.text().strip(), QColor(color).name().upper())
            self.assertEqual(button.styleSheet(), "")
            self.assertLess(button.iconSize().width(), 20)
            self.assertGreater(button.maximumWidth(), 20)
            image = button.icon().pixmap(button.iconSize()).toImage()
            center = image.pixelColor(image.width() // 2, image.height() // 2)
            self.assertEqual(center.name(), QColor(color).name())
            self.assertEqual(image.pixelColor(0, 0).alpha(), 0)

        disabled_icon = dialog.ui.btn_up_color.icon().pixmap(
            dialog.ui.btn_up_color.iconSize(), QIcon.Mode.Disabled
        ).toImage()
        disabled_center = disabled_icon.pixelColor(
            disabled_icon.width() // 2, disabled_icon.height() // 2
        )
        self.assertEqual(disabled_center.name(), window.up_color.name())

        for label_name in (
            "label_fg_color",
            "label_bg_color",
            "label_up_color",
            "label_down_color",
            "label_neutral_color",
        ):
            self.assertFalse(hasattr(dialog.ui, label_name))

        with patch(
            "stockwidget.ui.settings.dialog.QColorDialog.getColor",
            return_value=QColor("#123456"),
        ):
            dialog.ui.btn_fg_color.click()
        updated_icon = dialog.ui.btn_fg_color.icon().pixmap(dialog.ui.btn_fg_color.iconSize()).toImage()
        updated_center = updated_icon.pixelColor(
            updated_icon.width() // 2, updated_icon.height() // 2
        )
        self.assertEqual(updated_center.name(), "#123456")
        self.assertIn("#123456", dialog.ui.btn_fg_color.toolTip())

        dialog.ui.cb_unicolor.setChecked(False)
        self.qt_app.processEvents()

        self.assertFalse(window.unicolor)
        self.assertTrue(dialog.ui.btn_up_color.isEnabled())
        self.assertTrue(dialog.ui.btn_down_color.isEnabled())
        self.assertTrue(dialog.ui.btn_neutral_color.isEnabled())
        config = window.current_config()
        self.assertFalse(config["unicolor"])
        self.assertEqual(config["up_color"], window.up_color.name())
        self.assertEqual(config["down_color"], window.down_color.name())
        self.assertEqual(config["neutral_color"], window.neutral_color.name())

    def test_source_toggle_only_updates_for_selected_button(self):
        calls = []
        owner = type(
            "Owner",
            (),
            {"win": type("Window", (), {"set_data_source": calls.append})()},
        )()

        SettingsDialog._on_source_toggled(owner, "sina", False)
        SettingsDialog._on_source_toggled(owner, "eastmoney", True)

        self.assertEqual(calls, ["eastmoney"])

    def test_github_check_does_not_block_dialog_thread(self):
        owner = type("Owner", (), {"github_check_finished": SignalRecorder()})()

        def slow_check(timeout):
            time.sleep(0.2)
            return False

        with patch(
            "stockwidget.ui.settings.dialog.github_available",
            side_effect=slow_check,
        ):
            started = time.perf_counter()
            SettingsDialog._start_github_check(owner)
            elapsed = time.perf_counter() - started
            self.assertLess(elapsed, 0.1)
            self.assertTrue(owner.github_check_finished.event.wait(1))

        self.assertTrue(owner.github_check_finished.value)
