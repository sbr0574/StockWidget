import os
import re
import tempfile
import threading
import time
import unittest
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QPoint, QSize, Qt
from PySide6.QtGui import QColor, QIcon, QPalette, QPixmap
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QAbstractItemDelegate

from stockwidget.ui.add_code_panel import (
    ADDED_ROLE,
    ENTRY_ROLE,
    FilledCheckBox,
    PAGE_SIZE,
    RESULT_ROW_HEIGHT,
    SearchResultDelegate,
)
from stockwidget.ui.settings_dialog import (
    CodeSearchEditor,
    SettingsDialog,
    _COLOR_SWATCH_SIZE,
    _SEARCH_PLACEHOLDER,
    _build_settings_stylesheet,
    _color_swatch_icon,
)
from stockwidget.ui.widget import FloatLabel


CODES = {
    "sh600519": {
        "code": "600519",
        "market": "sh",
        "name": "贵州茅台",
        "type": "沪",
        "py": "guizhoumaotai",
        "abbr": "gzmt",
    },
    "sh501001": {
        "code": "501001",
        "market": "sh",
        "name": "财通精选混合LOF",
        "type": "基",
        "py": "caitongjingxuanhunhelof",
        "abbr": "ctjxhhlof",
    },
    "sh000001": {
        "code": "000001",
        "market": "sh",
        "name": "上证指数",
        "type": "指",
        "py": "shangzhengzhishu",
        "abbr": "szzs",
    },
    "ad0": {
        "code": "ad0",
        "market": "",
        "name": "铸造铝合金连续",
        "type": "期",
        "py": "zhuzaolvhejinlianxu",
        "abbr": "zzlhjlx",
    },
    "longname": {
        "code": "longname",
        "market": "us",
        "name": "用于验证候选列表根据条目内容自动扩展宽度的特别长证券名称",
        "type": "美",
        "py": "yongyuyanzhengchaotezhengquanmingcheng",
        "abbr": "cctm",
    },
}


class _SignalRecorder:
    def __init__(self):
        self.event = threading.Event()
        self.value = None

    def emit(self, value):
        self.value = value
        self.event.set()


class SettingsDialogTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt_app = QApplication.instance() or QApplication([])
        cls.qt_app.setQuitOnLastWindowClosed(False)

    def setUp(self):
        self._windows = []

    def tearDown(self):
        for dialog, window in reversed(self._windows):
            for editor in dialog.list_codes.findChildren(CodeSearchEditor):
                editor.setProperty("_code_editor_committed", True)
                editor._code_completer.popup().hide()
                dialog.list_codes.closeEditor(
                    editor, QAbstractItemDelegate.EndEditHint.NoHint
                )
            self.qt_app.processEvents()
            dialog.close()
            window.close()
            dialog.deleteLater()
            window.deleteLater()
        self.qt_app.processEvents()

    def _make_dialog(self, watchlist=None, codes=None, app=None):
        cfg = {"watchlist": watchlist or {}}
        window = FloatLabel(cfg, CODES if codes is None else codes)
        with patch.object(SettingsDialog, "_start_github_check"):
            dialog = SettingsDialog(window, window, app=app)
        self._windows.append((dialog, window))
        return dialog, window

    def _click_info(self, dialog, pool, item):
        dialog.show()
        self.qt_app.processEvents()
        pool.scrollToItem(item)
        QTest.mouseClick(pool.viewport(), Qt.LeftButton, pos=pool.info_rect(item).center())

    def _make_icon_app(self, choice="default", custom_path=""):
        app = Mock()
        app._icon_choice = choice
        app._custom_icon_path = custom_path
        app.app_version = "1.0.0"
        app._has_update = False
        app._latest_version = None
        app.code_data_state.return_value = ("cached", "")

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

    def _start_code_editor(self, dialog):
        dialog._start_quick_add()
        self.qt_app.processEvents()
        editor = dialog.list_codes.findChild(CodeSearchEditor)
        self.assertIsNotNone(editor)
        return editor

    def test_watchlist_action_buttons_share_color_button_style(self):
        regular = _build_settings_stylesheet(dark=False)
        dark = _build_settings_stylesheet(dark=True)

        for stylesheet in (regular, dark):
            for selector in ("btn_add", "btn_del", "btn_top"):
                self.assertIn(f"QPushButton#{selector}", stylesheet)
                self.assertIn(f"QPushButton#{selector}:hover", stylesheet)
                self.assertIn(f"QPushButton#{selector}:pressed", stylesheet)
                self.assertIn(f"QPushButton#{selector}:disabled", stylesheet)
            self.assertIn("QPushButton#btn_icon_default", stylesheet)
            self.assertIn("QPushButton#btn_icon_default:hover", stylesheet)
            self.assertIn("QPushButton#btn_icon_default:pressed", stylesheet)
            self.assertIn("QPushButton#btn_icon_default:checked", stylesheet)
            self.assertIn("QPushButton#btn_icon_default:checked:hover", stylesheet)
            self.assertIn("QPushButton#btn_icon_default:checked:pressed", stylesheet)
            self.assertIn("QPushButton#btn_icon_custom", stylesheet)
            self.assertIn("QPushButton#btn_fg_color:hover", stylesheet)
            self.assertIn("QPushButton#btn_fg_color:pressed", stylesheet)
            self.assertIn("QPushButton#btn_fg_color:disabled", stylesheet)
            self.assertNotIn("QPushButton#btn_fg_color:checked", stylesheet)
            self.assertIn(self.qt_app.palette().color(QPalette.Accent).name(), stylesheet)

            color_rules = []
            for state in ("", ":hover", ":pressed", ":disabled"):
                rule = re.search(
                    rf"QPushButton#btn_neutral_color{state}(?:,\s*[^{{]+)? \{{([^}}]+)", stylesheet
                ).group(1)
                background = next(
                    line.strip()
                    for line in rule.splitlines()
                    if line.strip().startswith("background-color:")
                )
                self.assertNotIn("transparent", background)
                color_rules.append(background)

            self.assertEqual(len(set(color_rules)), 4)
            color_base_rule = re.search(
                r"QPushButton#btn_neutral_color(?:,\s*[^{}]+)? \{([^}]+)", stylesheet
            ).group(1)
            self.assertIn("border: none", color_base_rule)
            self.assertIn("border-radius: 6px", color_base_rule)
            self.assertIn("padding: 3px", color_base_rule)

    def test_color_swatch_renders_at_device_pixel_ratio(self):
        icon = _color_swatch_icon(QColor("#123456"), 2.0)
        pixmap = icon.pixmap(QSize(_COLOR_SWATCH_SIZE, _COLOR_SWATCH_SIZE), 2.0)

        self.assertEqual(pixmap.devicePixelRatio(), 2.0)
        self.assertEqual(pixmap.width(), _COLOR_SWATCH_SIZE * 2)
        self.assertEqual(pixmap.height(), _COLOR_SWATCH_SIZE * 2)
        image = pixmap.toImage()
        center = image.pixelColor(image.width() // 2, image.height() // 2)
        self.assertEqual(center.name(), "#123456")
        self.assertEqual(image.pixelColor(0, 0).alpha(), 0)

    def test_buttons_and_metric_pools_follow_palette_changes(self):
        dialog, _window = self._make_dialog()
        original = self.qt_app.palette()
        try:
            palette = QPalette(original)
            palette.setColor(QPalette.Accent, QColor("#a23b61"))
            self.qt_app.setPalette(palette)
            self.qt_app.processEvents()
            self.assertIn("#a23b61", dialog.styleSheet())
            self.assertIn("rgba(162, 59, 97, 0.08)", dialog.styleSheet())
            for pool in (dialog.metric_pool.displayed_pool, dialog.metric_pool.available_pool):
                self.assertEqual(pool._drop_color.name(), "#a23b61")
                self.assertIn("rgba(162, 59, 97, 0.08)", pool.styleSheet())
        finally:
            self.qt_app.setPalette(original)

    def test_drop_restores_dragged_row_selection_after_cleanup(self):
        for target in (0, 1):
            with self.subTest(target=target):
                with patch.object(FloatLabel, "_refresh_from_function"):
                    dialog, _window = self._make_dialog({
                        "sh600519": {"checked": True},
                        "sh000001": {"checked": True},
                    })
                dialog.show()
                self.qt_app.processEvents()
                table = dialog.list_codes
                table.setCurrentCell(0, 1)
                dragged = table.item(0, 1)
                event = Mock()
                event.source.return_value = table
                event.position.return_value.toPoint.return_value = table.visualItemRect(table.item(target, 1)).center()
                dialog._handle_drop(event)
                self.assertEqual(table.selectedItems(), [])
                event.setDropAction.assert_called_once_with(Qt.CopyAction)
                self.qt_app.processEvents()
                self.assertIs(table.currentItem(), dragged)
                self.assertIn(dragged, table.selectedItems())
                self.assertEqual(table.currentRow(), target)
                self.assertEqual(table.rowCount(), 2)
                self.assertTrue(table.hasFocus())
                self.assertTrue(dialog.btn_del.isEnabled())
                self.assertEqual(dialog.btn_top.isEnabled(), target > 0)

    def test_general_layout_and_about_tab(self):
        dialog, _window = self._make_dialog()

        self.assertEqual(dialog.minimumSize(), dialog.maximumSize())
        self.assertEqual(dialog.size(), dialog.minimumSize())
        self.assertFalse(dialog.isSizeGripEnabled())
        self.assertEqual(
            [dialog.ui.tab_widget.tabText(i)
             for i in range(dialog.ui.tab_widget.count())],
            ["数据", "通用", "关于"],
        )
        self.assertIs(dialog.ui.gb_about.parentWidget(), dialog.ui.about)

        self.assertLess(dialog.ui.gb_icon.geometry().bottom(), dialog.ui.gb_fcn.y())
        self.assertLess(dialog.ui.gb_fcn.geometry().bottom(), dialog.ui.gb_hotkeys.y())
        self.assertLess(dialog.ui.gb_color.geometry().bottom(), dialog.ui.gb_text.y())
        self.assertEqual(dialog.ui.gb_fcn.x(), dialog.ui.gb_hotkeys.x())
        for button in dialog.icon_buttons.values():
            self.assertTrue(button.isFlat())
            self.assertEqual(button.styleSheet(), "")
        self.assertEqual(dialog.ui.btn_icon_custom.text(), "+")
        self.assertFalse(dialog.ui.btn_icon_custom.has_custom_icon())

    def test_custom_icon_can_be_selected_and_deleted_from_hover_action(self):
        app = self._make_icon_app()
        dialog, _window = self._make_dialog(app=app)

        with tempfile.TemporaryDirectory() as temp_dir:
            icon_path = os.path.join(temp_dir, "custom.png")
            pixmap = QPixmap(24, 24)
            pixmap.fill(QColor("#3578e5"))
            self.assertTrue(pixmap.save(icon_path))

            with patch(
                "stockwidget.ui.settings_dialog.QFileDialog.getOpenFileName",
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

            dialog.ui.tab_widget.setCurrentWidget(dialog.ui.general)
            dialog.show()
            self.qt_app.processEvents()
            QTest.mouseMove(dialog.ui.btn_icon_custom, QPoint(35, 5))
            self.qt_app.processEvents()
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
            "stockwidget.ui.settings_dialog.QFileDialog.getOpenFileName",
            return_value=("", ""),
        ):
            dialog.ui.btn_icon_custom.click()

        self.assertTrue(dialog.ui.btn_icon_dark.isChecked())
        self.assertEqual(app._icon_choice, "dark")
        app.set_custom_icon.assert_not_called()
        app.save_now.assert_not_called()

    def test_metric_pool_controls_order_and_name_visibility(self):
        dialog, window = self._make_dialog()
        save = Mock()
        window.set_on_change(save)

        self.assertEqual(
            dialog.metric_pool.visible_metrics,
            ["name", "price", "change_pct"],
        )
        self.assertTrue(window.name_visible)
        displayed_ids = [
            dialog.metric_pool.displayed_pool.item(row).data(
                Qt.ItemDataRole.UserRole
            )
            for row in range(dialog.metric_pool.displayed_pool.count())
        ]
        self.assertEqual(displayed_ids, ["name", "price", "change_pct"])

        with patch.object(window, "_refresh_from_function") as refresh:
            dialog.metric_pool.move_metric(
                "available", "displayed", "kline", 0
            )

        self.assertEqual(
            window.visible_metrics,
            ["kline", "name", "price", "change_pct"],
        )
        self.assertEqual(
            dialog.metric_pool.visible_metrics,
            window.visible_metrics,
        )
        save.assert_called_once_with()
        refresh.assert_called_once_with()

        # 名称与其他指标一样可从显示池移出，从而关闭名称列
        with patch.object(window, "_refresh_from_function") as refresh:
            dialog.metric_pool.move_metric(
                "displayed", "available", "name", 0
            )
        self.assertFalse(window.name_visible)
        self.assertEqual(
            window.visible_metrics,
            ["kline", "price", "change_pct"],
        )

    def test_name_metric_click_opens_settings_panel(self):
        dialog, window = self._make_dialog()
        pool = dialog.metric_pool.displayed_pool
        name_item = next(
            pool.item(row)
            for row in range(pool.count())
            if pool.item(row).data(Qt.ItemDataRole.UserRole) == "name"
        )

        with patch.object(window, "_refresh_from_function"):
            self._click_info(dialog, pool, name_item)
        self.qt_app.processEvents()

        panel = dialog.name_settings_panel
        self.assertTrue(panel.isVisible())
        self.assertEqual(panel.cmb_namelen.currentData(), window.name_length)
        self.assertFalse(panel.cb_code.isChecked())
        self.assertFalse(panel.cb_type.isChecked())

        with patch.object(window, "_refresh_from_function"):
            panel.cb_code.setChecked(True)
            panel.cb_type.setChecked(True)
            panel.cmb_namelen.setCurrentIndex(
                panel.cmb_namelen.findData(2)
            )
        self.assertTrue(window.code_visible)
        self.assertTrue(window.type_visible)
        self.assertEqual(window.name_length, 2)

        panel.hide()

    def test_volume_metric_click_opens_shared_unit_settings_panel(self):
        dialog, window = self._make_dialog()
        pool = dialog.metric_pool.available_pool
        volume_item = next(
            pool.item(row)
            for row in range(pool.count())
            if pool.item(row).data(Qt.ItemDataRole.UserRole) == "volume"
        )

        with patch.object(window, "_refresh_from_function"):
            self._click_info(dialog, pool, volume_item)
        self.qt_app.processEvents()

        panel = dialog.unit_settings_panel
        self.assertTrue(panel.isVisible())
        self.assertTrue(panel.radio_buttons["auto"].isChecked())
        self.assertEqual(window.unit_mode, "auto")

        # 成交量/成交额共享同一面板与同一设置
        with patch.object(window, "_refresh_from_function"):
            panel.radio_buttons["en"].setChecked(True)
        self.assertEqual(window.unit_mode, "en")
        self.assertEqual(
            window.current_config()["unit_mode"], "en"
        )
        panel.sync_from(window)
        self.assertTrue(panel.radio_buttons["en"].isChecked())

        panel.hide()

    def test_delete_button_disabled_without_watchlist_selection(self):
        watchlist = {
            "sh600519": {
                "checked": True,
                "name": "贵州茅台",
                "type": "沪",
                "market": "sh",
                "code": "600519",
            },
            "sh000001": {
                "checked": True,
                "name": "上证指数",
                "type": "指",
                "market": "sh",
                "code": "000001",
            },
        }
        dialog, _window = self._make_dialog(watchlist)

        self.assertFalse(dialog.btn_del.isEnabled())
        self.assertFalse(dialog.btn_top.isEnabled())

        dialog.list_codes.setCurrentCell(1, 1)
        self.qt_app.processEvents()
        self.assertTrue(dialog.btn_del.isEnabled())
        self.assertTrue(dialog.btn_top.isEnabled())

        # 点击空白处后取消选中，删除与置顶按钮均禁用
        dialog.list_codes.setCurrentCell(-1, -1)
        self.qt_app.processEvents()
        self.assertFalse(dialog.btn_del.isEnabled())
        self.assertFalse(dialog.btn_top.isEnabled())

    def test_outside_press_clears_watchlist_and_metric_selections(self):
        watchlist = {
            "sh600519": {
                "checked": True,
                "name": "贵州茅台",
                "type": "沪",
                "market": "sh",
                "code": "600519",
            },
            "sh000001": {
                "checked": True,
                "name": "上证指数",
                "type": "指",
                "market": "sh",
                "code": "000001",
            },
        }
        dialog, _window = self._make_dialog(watchlist)
        displayed = dialog.metric_pool.displayed_pool

        dialog.list_codes.setCurrentCell(1, 1)
        displayed.setCurrentItem(displayed.item(1))
        self.qt_app.processEvents()
        self.assertEqual(dialog.list_codes.currentRow(), 1)
        self.assertTrue(dialog.btn_del.isEnabled())
        self.assertIsNotNone(displayed.currentItem())

        # 点击设置页其他区域（空白处）→ 自选列表与指标池选中全部清除
        dialog._handle_outside_press(dialog.ui.gb_color)
        self.qt_app.processEvents()

        self.assertEqual(dialog.list_codes.currentRow(), -1)
        self.assertFalse(dialog.btn_del.isEnabled())
        self.assertFalse(dialog.btn_top.isEnabled())
        self.assertIsNone(displayed.currentItem())

    def test_press_inside_pool_clears_watchlist_and_other_pool(self):
        watchlist = {
            "sh600519": {
                "checked": True,
                "name": "贵州茅台",
                "type": "沪",
                "market": "sh",
                "code": "600519",
            },
        }
        dialog, _window = self._make_dialog(watchlist)
        displayed = dialog.metric_pool.displayed_pool
        available = dialog.metric_pool.available_pool

        dialog.list_codes.setCurrentCell(0, 1)
        displayed.setCurrentItem(displayed.item(1))
        self.qt_app.processEvents()

        dialog._handle_outside_press(available.viewport())
        self.qt_app.processEvents()

        self.assertEqual(dialog.list_codes.currentRow(), -1)
        self.assertIsNone(displayed.currentItem())
        self.assertFalse(dialog.btn_del.isEnabled())

    def test_press_on_delete_or_top_button_keeps_selection(self):
        watchlist = {
            "sh600519": {
                "checked": True,
                "name": "贵州茅台",
                "type": "沪",
                "market": "sh",
                "code": "600519",
            },
            "sh000001": {
                "checked": True,
                "name": "上证指数",
                "type": "指",
                "market": "sh",
                "code": "000001",
            },
        }
        dialog, _window = self._make_dialog(watchlist)
        dialog.list_codes.setCurrentCell(1, 1)
        self.qt_app.processEvents()

        dialog._handle_outside_press(dialog.btn_del)
        dialog._handle_outside_press(dialog.btn_top)
        self.qt_app.processEvents()

        self.assertEqual(dialog.list_codes.currentRow(), 1)
        self.assertTrue(dialog.btn_del.isEnabled())

    def test_clickable_metric_chips_have_i_badge_and_tooltip(self):
        dialog, _window = self._make_dialog()
        displayed = dialog.metric_pool.displayed_pool
        available = dialog.metric_pool.available_pool

        name_item = displayed.item(0)
        self.assertEqual(
            name_item.data(Qt.ItemDataRole.UserRole), "name"
        )
        self.assertIn("ⓘ", name_item.text())
        self.assertIn("单击", name_item.toolTip())
        price_item = displayed.item(1)
        self.assertNotIn("ⓘ", price_item.text())
        self.assertEqual(price_item.toolTip(), "")
        # 可单击块的 ⓘ 提示计入文本宽度
        self.assertGreater(
            name_item.sizeHint().width(), price_item.sizeHint().width()
        )

        volume_item = next(
            available.item(row)
            for row in range(available.count())
            if available.item(row).data(Qt.ItemDataRole.UserRole) == "volume"
        )
        self.assertIn("ⓘ", volume_item.text())
        self.assertIn("单位", volume_item.toolTip())

    def test_plain_metric_chip_does_not_keep_selection(self):
        dialog, _window = self._make_dialog()
        displayed = dialog.metric_pool.displayed_pool
        available = dialog.metric_pool.available_pool

        # 单击没有独立面板的指标块：不保留选中状态
        displayed.itemClicked.emit(displayed.item(1))
        self.qt_app.processEvents()
        self.assertIsNone(displayed.currentItem())

        # 双击/拖动切换显示后，普通指标同样不保留选中
        dialog.metric_pool.move_metric("available", "displayed", "kline", 0)
        self.qt_app.processEvents()
        self.assertIsNone(displayed.currentItem())
        self.assertIsNone(available.currentItem())

    def test_double_click_moves_clickable_metric_without_opening_panel(self):
        dialog, window = self._make_dialog()
        displayed = dialog.metric_pool.displayed_pool
        name_item = next(
            displayed.item(row)
            for row in range(displayed.count())
            if displayed.item(row).data(Qt.ItemDataRole.UserRole) == "name"
        )

        # 点击文字立即结束，不会打开面板，也没有待执行的延迟。
        dialog.show()
        self.qt_app.processEvents()
        pos = displayed.visualItemRect(name_item).center() - QPoint(8, 0)
        QTest.mouseClick(displayed.viewport(), Qt.LeftButton, pos=pos)
        self.assertIsNone(displayed.currentItem())
        self.assertFalse(dialog.name_settings_panel.isVisible())

        # 双击文字快速移动到另一池。
        with patch.object(window, "_refresh_from_function"):
            QTest.mouseDClick(displayed.viewport(), Qt.LeftButton, pos=pos)
        self.qt_app.processEvents()

        self.assertFalse(dialog.name_settings_panel.isVisible())
        self.assertFalse(window.name_visible)
        self.assertEqual(displayed.currentItem(), None)

    def test_clickable_chip_loses_selection_when_panel_closes(self):
        dialog, _window = self._make_dialog()
        displayed = dialog.metric_pool.displayed_pool
        name_item = next(
            displayed.item(row)
            for row in range(displayed.count())
            if displayed.item(row).data(Qt.ItemDataRole.UserRole) == "name"
        )

        self._click_info(dialog, displayed, name_item)
        self.qt_app.processEvents()
        panel = dialog.name_settings_panel
        self.assertTrue(panel.isVisible())

        displayed.setCurrentItem(name_item)
        self.assertIsNotNone(displayed.currentItem())

        # 面板关闭（含点击外部空白关闭）后指标块失焦
        panel.hide()
        self.qt_app.processEvents()
        self.assertIsNone(displayed.currentItem())

    def test_info_click_toggles_immediately_and_outside_click_closes(self):
        dialog, _window = self._make_dialog()
        pool = dialog.metric_pool.displayed_pool
        item = pool.item(0)
        self._click_info(dialog, pool, item)
        panel = dialog.name_settings_panel
        self.assertTrue(panel.isVisible())
        self.assertIs(pool.currentItem(), item)
        # QTest 直接投递到指标池时，也应切换关闭。
        QTest.mouseClick(pool.viewport(), Qt.LeftButton, pos=pool.info_rect(item).center())
        self.assertFalse(panel.isVisible())
        self.assertIsNone(pool.currentItem())

        self._click_info(dialog, pool, item)
        # 原生 Popup 把外部点击交给弹窗：同一个 ⓘ 禁止重放，避免重开。
        global_pos = pool.viewport().mapToGlobal(pool.info_rect(item).center())
        with patch.object(panel, "setAttribute", wraps=panel.setAttribute) as set_attribute:
            QTest.mouseClick(panel, Qt.LeftButton, pos=panel.mapFromGlobal(global_pos))
            set_attribute.assert_any_call(Qt.WA_NoMouseReplay, True)
        self.assertFalse(panel.isVisible())
        self.assertIsNone(pool.currentItem())

        self._click_info(dialog, pool, item)
        global_pos = dialog.ui.gb_color.mapToGlobal(QPoint(4, 4))
        QTest.mouseClick(panel, Qt.LeftButton, pos=panel.mapFromGlobal(global_pos))
        self.assertFalse(panel.isVisible())
        self.assertFalse(panel.testAttribute(Qt.WA_NoMouseReplay))
        self.assertIsNone(pool.currentItem())

    def test_switching_info_entry_preserves_new_highlight(self):
        dialog, _window = self._make_dialog()
        name_pool = dialog.metric_pool.displayed_pool
        self._click_info(dialog, name_pool, name_pool.item(0))
        pool = dialog.metric_pool.available_pool
        item = next(pool.item(r) for r in range(pool.count())
                    if pool.item(r).data(Qt.UserRole) == "volume")
        QTest.mouseClick(pool.viewport(), Qt.LeftButton, pos=pool.info_rect(item).center())
        self.assertFalse(dialog.name_settings_panel.isVisible())
        self.assertTrue(dialog.unit_settings_panel.isVisible())
        self.assertIs(pool.currentItem(), item)
        self.assertIsNone(name_pool.currentItem())
        self.assertEqual(dialog.unit_settings_panel._metric_id, "volume")

    def test_manual_update_check_shows_result(self):
        dialog, _window = self._make_dialog()

        with patch(
            "stockwidget.ui.settings_dialog.QMessageBox.information"
        ) as info:
            dialog._on_update_check_finished((True, "9.9.9"))
            info.assert_called_once()

        with patch(
            "stockwidget.ui.settings_dialog.QMessageBox.information"
        ) as info:
            dialog._on_update_check_finished((False, "1.4.0"))
            info.assert_called_once()

        with patch(
            "stockwidget.ui.settings_dialog.QMessageBox.warning"
        ) as warn:
            dialog._on_update_check_finished((False, None))
            warn.assert_called_once()

    def test_open_cache_dir_button_opens_folder(self):
        from stockwidget.core.config_store import config_paths

        dialog, _window = self._make_dialog()
        with patch(
            "stockwidget.ui.settings_dialog.QDesktopServices.openUrl"
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
        self.assertTrue(dialog.cb_unicolor.isChecked())
        self.assertEqual(
            dialog.cb_unicolor.minimumWidth(), dialog.cb_unicolor.maximumWidth()
        )
        self.assertGreaterEqual(
            dialog.cb_unicolor.minimumWidth(), dialog.cb_unicolor.sizeHint().width()
        )
        self.assertTrue(dialog.btn_bg.isEnabled())
        self.assertTrue(dialog.btn_fg.isEnabled())
        self.assertFalse(dialog.btn_up.isEnabled())
        self.assertFalse(dialog.btn_down.isEnabled())
        self.assertFalse(dialog.btn_neutral.isEnabled())

        color_buttons = (
            (dialog.btn_fg, "文字", window.fg),
            (dialog.btn_bg, "背景", window.bg),
            (dialog.btn_up, "上涨", window.up_color),
            (dialog.btn_down, "下跌", window.down_color),
            (dialog.btn_neutral, "中性", window.neutral_color),
        )
        for button, text, color in color_buttons:
            self.assertTrue(button.isFlat())
            self.assertEqual(button.text(), text)
            self.assertEqual(button.styleSheet(), "")
            self.assertLess(button.iconSize().width(), 20)
            self.assertGreater(button.maximumWidth(), 20)
            image = button.icon().pixmap(button.iconSize()).toImage()
            center = image.pixelColor(image.width() // 2, image.height() // 2)
            self.assertEqual(center.name(), QColor(color).name())
            self.assertEqual(image.pixelColor(0, 0).alpha(), 0)

        disabled_icon = dialog.btn_up.icon().pixmap(
            dialog.btn_up.iconSize(), QIcon.Mode.Disabled
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
            "stockwidget.ui.settings_dialog.QColorDialog.getColor",
            return_value=QColor("#123456"),
        ):
            dialog.pick_fg()
        updated_icon = dialog.btn_fg.icon().pixmap(dialog.btn_fg.iconSize()).toImage()
        updated_center = updated_icon.pixelColor(
            updated_icon.width() // 2, updated_icon.height() // 2
        )
        self.assertEqual(updated_center.name(), "#123456")
        self.assertIn("#123456", dialog.btn_fg.toolTip())

        dialog.cb_unicolor.setChecked(False)
        self.qt_app.processEvents()

        self.assertFalse(window.unicolor)
        self.assertTrue(dialog.btn_up.isEnabled())
        self.assertTrue(dialog.btn_down.isEnabled())
        self.assertTrue(dialog.btn_neutral.isEnabled())
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
        owner = type("Owner", (), {"github_check_finished": _SignalRecorder()})()

        def slow_check(timeout):
            time.sleep(0.2)
            return False

        with patch(
            "stockwidget.ui.settings_dialog.github_available",
            side_effect=slow_check,
        ):
            started = time.perf_counter()
            SettingsDialog._start_github_check(owner)
            elapsed = time.perf_counter() - started
            self.assertLess(elapsed, 0.1)
            self.assertTrue(owner.github_check_finished.event.wait(1))

        self.assertTrue(owner.github_check_finished.value)

    def test_quick_editor_has_no_category_selector_and_searches_all_types(self):
        dialog, _window = self._make_dialog()
        editor = self._start_code_editor(dialog)

        self.assertEqual(editor.placeholderText(), _SEARCH_PLACEHOLDER)
        self.assertFalse(hasattr(editor, "category_combo"))

        expected = {
            "茅台": "sh600519",
            "财通": "sh501001",
            "上证": "sh000001",
            "铝合金": "ad0",
        }
        for query, key in expected.items():
            dialog._update_suggestions(editor, query)
            self.assertEqual(
                dialog.suggestion_model.item(0).data(ENTRY_ROLE)["key"], key
            )

    def test_empty_hint_is_visible_and_does_not_block_double_click(self):
        dialog, _window = self._make_dialog()
        dialog.show()
        self.qt_app.processEvents()

        hint = dialog.empty_watchlist_hint
        self.assertEqual(hint.text(), "双击空白处添加条目")
        self.assertFalse(hint.isHidden())
        self.assertTrue(
            hint.testAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        )

        QTest.mouseDClick(
            dialog.list_codes.viewport(),
            Qt.MouseButton.LeftButton,
            pos=QPoint(20, 80),
        )
        self.qt_app.processEvents()

        self.assertEqual(dialog.list_codes.rowCount(), 1)
        self.assertTrue(hint.isHidden())
        self.assertIsNotNone(dialog.list_codes.findChild(CodeSearchEditor))

    def test_add_button_opens_panel_without_creating_a_row(self):
        dialog, _window = self._make_dialog()
        dialog.move(20, 20)
        dialog.show()
        self.qt_app.processEvents()

        dialog.btn_add.click()
        self.qt_app.processEvents()

        panel = dialog.add_code_panel
        self.assertEqual(dialog.list_codes.rowCount(), 0)
        self.assertTrue(panel.isVisible())
        self.assertEqual(panel.search_input.placeholderText(), _SEARCH_PLACEHOLDER)
        self.assertEqual(
            [box.text() for box in panel.category_filters.option_checkboxes.values()],
            ["股票", "基金", "指数", "期货"],
        )
        self.assertEqual(
            [box.text() for box in panel.region_filters.option_checkboxes.values()],
            ["沪", "深", "京", "港", "美", "其他"],
        )
        button_bottom = dialog.btn_add.mapToGlobal(
            QPoint(0, dialog.btn_add.height())
        ).y()
        self.assertGreaterEqual(panel.y(), button_bottom)
        self.assertEqual(panel.search_input.styleSheet(), "")
        self.assertNotIn("QFrame#add_code_panel QLineEdit", panel.styleSheet())
        checkboxes = panel.findChildren(FilledCheckBox)
        self.assertEqual(len(checkboxes), 12)
        self.assertTrue(
            all(box._unchecked_fill != box._checked_fill for box in checkboxes)
        )
        self.assertTrue(all(box._hover_fill.alpha() > 0 for box in checkboxes))
        self.assertTrue(
            all(
                box.sizeHint().height()
                >= box.fontMetrics().height()
                + 2 * box._HOVER_VERTICAL_PADDING
                for box in checkboxes
            )
        )
        self.assertGreater(
            panel.category_filters.layout().spacing(),
            FilledCheckBox._INDICATOR_TEXT_GAP,
        )
        self.assertIn("QListView::item:hover", panel.styleSheet())
        self.assertIn("color: rgb(28, 28, 30);", panel.styleSheet())

    def test_top_button_moves_selected_row_to_top_and_persists_order(self):
        watchlist = {
            "sh600519": {
                "checked": True,
                "cost": 123,
                "name": "贵州茅台",
                "type": "沪",
                "market": "sh",
                "code": "600519",
            },
            "sh501001": {
                "checked": False,
                "cost": 12.5,
                "name": "财通精选混合LOF",
                "type": "基",
                "market": "sh",
                "code": "501001",
            },
            "sh000001": {
                "checked": True,
                "name": "上证指数",
                "type": "指",
                "market": "sh",
                "code": "000001",
            },
        }
        dialog, window = self._make_dialog(watchlist)

        self.assertEqual(dialog.btn_top.text(), "置顶")
        self.assertIs(dialog.btn_top.parentWidget(), dialog.ui.gb_list)
        self.assertLessEqual(
            dialog.btn_top.geometry().right(), dialog.ui.gb_list.width()
        )
        self.assertLess(dialog.btn_top.iconSize().width(), 20)
        # 没有选中条目时按钮禁用
        self.assertFalse(dialog.btn_top.isEnabled())

        dialog.list_codes.setCurrentCell(2, 1)
        self.assertTrue(dialog.btn_top.isEnabled())

        dialog.btn_top.click()

        self.assertEqual(
            dialog.list_codes.item(0, 1).data(Qt.ItemDataRole.UserRole),
            "sh000001",
        )
        self.assertEqual(
            dialog.list_codes.item(1, 1).data(Qt.ItemDataRole.UserRole),
            "sh600519",
        )
        self.assertEqual(
            list(window.watchlist), ["sh000001", "sh600519", "sh501001"]
        )
        # 置顶后当前行回到顶部，按钮随之禁用
        self.assertEqual(dialog.list_codes.currentRow(), 0)
        self.assertFalse(dialog.btn_top.isEnabled())

        # 已在顶部的行再点置顶不产生变化
        dialog._top_code()
        self.assertEqual(
            list(window.watchlist), ["sh000001", "sh600519", "sh501001"]
        )

    def test_opening_panel_cancels_unfinished_quick_add_row(self):
        dialog, _window = self._make_dialog()
        self._start_code_editor(dialog)
        self.assertEqual(dialog.list_codes.rowCount(), 1)

        dialog._show_add_code_panel()
        self.qt_app.processEvents()

        self.assertEqual(dialog.list_codes.rowCount(), 0)
        self.assertTrue(dialog.add_code_panel.isVisible())
        self.assertFalse(dialog.empty_watchlist_hint.isHidden())

    def test_filter_select_all_checkbox_uses_three_states(self):
        dialog, _window = self._make_dialog()
        filters = dialog.add_code_panel.category_filters
        fund = filters.option_checkboxes["fund"]

        fund.setChecked(False)
        self.assertEqual(
            filters.all_checkbox.checkState(), Qt.CheckState.PartiallyChecked
        )
        filters.all_checkbox.click()
        self.assertEqual(filters.all_checkbox.checkState(), Qt.CheckState.Checked)
        self.assertTrue(all(box.isChecked() for box in filters.option_checkboxes.values()))
        filters.all_checkbox.click()
        self.assertEqual(filters.all_checkbox.checkState(), Qt.CheckState.Unchecked)
        self.assertFalse(any(box.isChecked() for box in filters.option_checkboxes.values()))

    def test_panel_paginates_ten_results_and_resets_after_filter_change(self):
        codes = {
            f"sh{code:06d}": {
                "code": f"{code:06d}",
                "market": "sh",
                "name": f"测试股票{code}",
                "type": "沪",
            }
            for code in range(21)
        }
        dialog, _window = self._make_dialog(codes=codes)
        panel = dialog.add_code_panel
        dialog._show_add_code_panel()

        self.assertEqual(panel.current_result.total, 21)
        self.assertEqual(panel.current_result.page_count, 3)
        self.assertEqual(panel.result_model.rowCount(), PAGE_SIZE)
        self.qt_app.processEvents()
        self.assertIsInstance(panel.result_list.itemDelegate(), SearchResultDelegate)
        self.assertEqual(
            panel.result_list.viewport().height(), PAGE_SIZE * RESULT_ROW_HEIGHT
        )
        self.assertEqual(
            [panel.result_list.sizeHintForRow(row) for row in range(PAGE_SIZE)],
            [RESULT_ROW_HEIGHT] * PAGE_SIZE,
        )
        self.assertEqual(
            panel.result_list.visualRect(
                panel.result_model.index(PAGE_SIZE - 1, 0)
            ).bottom(),
            panel.result_list.viewport().rect().bottom(),
        )
        panel.next_button.click()
        self.assertEqual(panel.current_result.page, 2)
        self.assertEqual(panel.result_model.rowCount(), PAGE_SIZE)

        panel.category_filters.option_checkboxes["stock"].setChecked(False)
        self.assertEqual(panel.current_result.page, 0)
        self.assertEqual(panel.current_result.total, 0)
        self.assertEqual(panel.result_model.rowCount(), 0)

    def test_panel_adds_complete_entry_and_marks_it_added(self):
        dialog, window = self._make_dialog()
        save_callback = Mock()
        window.set_on_change(save_callback)
        dialog._show_add_code_panel()
        panel = dialog.add_code_panel

        panel._activate_index(panel.result_model.index(0, 0))

        self.assertEqual(dialog.list_codes.rowCount(), 1)
        self.assertIn("sh600519", window.watchlist)
        self.assertEqual(window.watchlist["sh600519"]["market"], "sh")
        self.assertEqual(window.watchlist["sh600519"]["code"], "600519")
        self.assertTrue(window.watchlist["sh600519"]["checked"])
        self.assertEqual(save_callback.call_count, 1)
        self.assertTrue(panel.result_model.item(0).data(ADDED_ROLE))
        self.assertFalse(
            panel.result_model.item(0).flags() & Qt.ItemFlag.ItemIsEnabled
        )
        self.assertEqual(panel.result_model.item(0).text(), "沪/600519/贵州茅台")
        self.assertNotIn("已添加", panel.result_model.item(0).text())

    def test_panel_adds_new_entry_at_start_of_watchlist(self):
        watchlist = {
            "sh501001": {
                "checked": False,
                "cost": 12.5,
                "name": "财通精选混合LOF",
                "type": "基",
                "market": "sh",
                "code": "501001",
            }
        }
        dialog, window = self._make_dialog(watchlist)
        dialog._show_add_code_panel()
        panel = dialog.add_code_panel

        panel._activate_index(panel.result_model.index(0, 0))

        self.assertEqual(
            dialog.list_codes.item(0, 1).data(Qt.ItemDataRole.UserRole),
            "sh600519",
        )
        self.assertEqual(
            dialog.list_codes.item(1, 1).data(Qt.ItemDataRole.UserRole),
            "sh501001",
        )
        self.assertEqual(list(window.watchlist), ["sh600519", "sh501001"])
        self.assertFalse(window.watchlist["sh501001"]["checked"])
        self.assertEqual(window.watchlist["sh501001"]["cost"], 12.5)

    def test_added_result_has_prefix_and_cannot_be_selected(self):
        watchlist = {
            "sh600519": {
                "checked": True,
                "cost": 123,
                "name": "贵州茅台",
                "type": "沪",
                "market": "sh",
                "code": "600519",
            }
        }
        dialog, _window = self._make_dialog(watchlist)
        editor = self._start_code_editor(dialog)

        dialog._update_suggestions(editor, "茅台")
        item = dialog.suggestion_model.item(0)
        self.assertEqual(item.text(), "（已添加）沪/600519/贵州茅台")
        self.assertTrue(item.data(ADDED_ROLE))
        self.assertFalse(item.flags() & Qt.ItemFlag.ItemIsEnabled)
        self.assertFalse(item.flags() & Qt.ItemFlag.ItemIsSelectable)

        completion_index = editor._code_completer.completionModel().index(0, 0)
        self.assertFalse(dialog._apply_suggestion(editor, completion_index))
        self.assertIsNone(editor.property("_selected_entry"))
        self.assertFalse(editor._code_completer.popup().currentIndex().isValid())

    def test_duplicate_manual_commit_preserves_original_row(self):
        watchlist = {
            "sh600519": {
                "checked": True,
                "cost": 123,
                "name": "贵州茅台",
                "type": "沪",
                "market": "sh",
                "code": "600519",
            }
        }
        dialog, _window = self._make_dialog(watchlist)
        editor = self._start_code_editor(dialog)
        editor.setText("600519")

        dialog.list_codes.itemDelegateForColumn(1)._commit_editor(editor)

        self.assertEqual(dialog.list_codes.rowCount(), 1)
        self.assertEqual(
            dialog.list_codes.item(0, 1).data(Qt.ItemDataRole.UserRole),
            "sh600519",
        )
        self.assertEqual(dialog.list_codes.item(0, 2).text(), "123")

    def test_popup_uses_table_width_and_expands_for_long_result(self):
        dialog, _window = self._make_dialog()
        editor = self._start_code_editor(dialog)
        base_width = (
            dialog.list_codes.columnWidth(1)
            + dialog.list_codes.columnWidth(2)
        )

        dialog._update_suggestions(editor, "茅台")
        popup = editor._code_completer.popup()
        self.assertEqual(popup.minimumWidth(), base_width)
        self.assertEqual(popup.maximumWidth(), base_width)

        dialog._update_suggestions(editor, "特别长证券名称")
        self.assertGreater(popup.minimumWidth(), base_width)
        self.assertEqual(popup.minimumWidth(), popup.maximumWidth())


if __name__ == "__main__":
    unittest.main()
