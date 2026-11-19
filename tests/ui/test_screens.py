"""多屏任务栏手势与设置窗口选屏；原生坐标由平台层提供。"""

from types import SimpleNamespace
from unittest.mock import Mock, patch

from PySide6.QtCore import QPoint, QRect

from stockwidget.app import App
from stockwidget.platform.taskbar import TaskbarArea
from tests.support import PagingTestCase, QtTestCase


PRIMARY = r"\\.\DISPLAY1"
SECONDARY = r"\\.\DISPLAY2"


class SettingsScreenTests(QtTestCase):
    def test_settings_follow_float_or_active_taskbar_and_recenter_when_open(self):
        primary, secondary = Mock(), Mock()
        for screen, name, geometry in ((primary, PRIMARY, QRect(0, 0, 1920, 1080)),
                                       (secondary, SECONDARY, QRect(-1280, 0, 1280, 1024))):
            screen.name.return_value = name
            screen.geometry.return_value = geometry
            screen.availableGeometry.return_value = geometry.adjusted(0, 0, 0, -48)
        win = SimpleNamespace(display_mode="float", taskbar_screen=PRIMARY, position_controller=Mock())
        win.position_controller.full_geometry.return_value = QRect(-1100, 100, 400, 100)
        dialog = Mock()
        dialog.width.return_value, dialog.height.return_value = 704, 496
        app = SimpleNamespace(win=win, taskbar=SimpleNamespace(_active=True), settings_dlg=None)
        app.taskbar.screen = lambda: next((s for s in (primary, secondary) if s.name() == win.taskbar_screen), None) \
            if app.taskbar._active else None
        with patch("stockwidget.app.QApplication.screens", return_value=[primary, secondary]), \
                patch("stockwidget.app.QApplication.primaryScreen", return_value=primary), \
                patch("stockwidget.app.SettingsDialog", return_value=dialog) as create:
            for mode, active, name, expected in (("float", True, PRIMARY, secondary),
                                                 ("both", True, PRIMARY, secondary),
                                                 ("taskbar", True, PRIMARY, primary),
                                                 ("taskbar", True, SECONDARY, secondary),
                                                 ("taskbar", False, PRIMARY, secondary),
                                                 ("taskbar", True, "removed-display", secondary)):
                with self.subTest(mode=mode, active=active, name=name):
                    win.display_mode, app.taskbar._active, win.taskbar_screen = mode, active, name
                    App.open_settings(app)
                    dialog.windowHandle().setScreen.assert_called_with(expected)
                    rect = expected.availableGeometry()
                    dialog.move.assert_called_with(QPoint(rect.x() + (rect.width() - 704) // 2,
                                                          rect.y() + (rect.height() - 496) // 2))
            create.assert_called_once()  # An already-open dialog moves without discarding edits.


class MultiscreenDockingTests(PagingTestCase):
    def test_native_monitor_origin_maps_qt_friendly_names_on_scaled_screens(self):
        controller = self.make_controller()
        controller._active = True
        controller._area = TaskbarArea(2, 3840, 96, 3600, 192, SECONDARY, (2560, 0, 3840, 2160))
        primary, secondary = Mock(), Mock()
        primary.name.return_value, secondary.name.return_value = PRIMARY, "LG HDR 4K"
        primary.geometry.return_value = QRect(0, 0, 1280, 800)
        secondary.geometry.return_value = QRect(2560, 0, 1920, 1080)
        with patch("stockwidget.ui.floating.taskbar.QApplication.screens", return_value=[primary, secondary]):
            self.assertIs(controller.screen(), secondary)
        controller._active = False
        self.assertIsNone(controller.screen())

    def multiscreen_controller(self):
        controller = self.make_controller()
        areas = (TaskbarArea(1, 1920, 48, 1700, 96, PRIMARY),
                 TaskbarArea(2, 2560, 72, 2300, 144, SECONDARY, left=-2560))
        self.finder = self.enterContext(patch("stockwidget.ui.floating.taskbar.find_taskbar",
            side_effect=lambda **kw: areas[1] if kw["hwnd"] == 2 or kw["device_name"] == SECONDARY else areas[0]))
        return controller

    def gesture(self, controller, mode, end):
        press = QPoint(700, 760)
        target = QPoint(-1727, 760)
        self.native_mouse("press", 100, 20, press, 1)
        self.native_mouse("move", 150, 20, target, 2, screen_x=-1310)
        self.assertEqual(self.win.taskbar_screen, PRIMARY)
        self.assertEqual(controller._area.device_name, SECONDARY)
        self.assertEqual(controller._dpi, 144)
        self.assertEqual(self.native.present.call_args.kwargs["x"], 1100)
        self.assertTrue(controller._active)
        self.assertEqual(self.win.isVisible(), mode == "both")
        self.native_mouse(end, 150, 20, target, 2, screen_x=-1310)

    def test_drag_between_taskbars_commits_the_target_display_and_escape_restores(self):
        controller = self.multiscreen_controller()
        for mode in ("taskbar", "both"):
            with self.subTest(mode=mode):
                self.win.taskbar_x = 600
                self.win.set_display_mode(mode, taskbar_screen=PRIMARY)
                origin = self.win.pos()
                self.gesture(controller, mode, "cancel")
                self.assertEqual(self.win.taskbar_screen, PRIMARY)
                self.assertEqual(controller._area.device_name, PRIMARY)
                self.assertEqual(self.native.present.call_args.kwargs["x"], 600)
                self.assertEqual(self.win.display_mode, mode)
                self.assertFalse(controller.input_timer.isActive())
                self.gesture(controller, mode, "release")
                self.assertEqual(self.win.current_config()["taskbar_screen"], SECONDARY)
                self.assertEqual(self.win.current_config()["taskbar_x"], 1100 * 96 / 144)
                self.assertEqual(self.native.present.call_args.kwargs["x"], 1100)
                self.assertEqual(self.win.display_mode, mode)
                self.assertEqual(self.win.isVisible(), mode == "both")
                self.assertEqual(self.win.pos(), origin)

    def test_dual_mode_keeps_last_taskbar_across_gap_and_float_moves_independently(self):
        controller = self.multiscreen_controller()
        self.win.taskbar_x = 600
        self.win.set_display_mode("both")
        self.win.move(100, 100)
        origin = self.win.pos()
        press = QPoint(700, 760)
        gap = QPoint(-1727, 550)
        self.native_mouse("press", 100, 20, press, 1)
        self.native_mouse("move", 150, -190, gap, screen_x=-1310)
        self.assertEqual(controller._area.device_name, PRIMARY)
        self.native_mouse("move", 150, 20, QPoint(-1727, 760), 2, screen_x=-1310)
        self.native_mouse("move", 150, -190, gap, screen_x=-1310)
        controller.refresh()
        self.assertTrue(controller._active)
        self.assertEqual(controller._area.device_name, SECONDARY)
        self.native_mouse("release", 150, -190, gap, screen_x=-1310)
        self.assertEqual(self.win.taskbar_screen, SECONDARY)
        self.assertEqual(self.win.pos(), origin)
        self.native.hide.assert_not_called()

        self.over.return_value = 1
        start = origin + QPoint(5, 5)
        self.win.begin_drag(start)
        self.win.move_drag(start + QPoint(2000, 50))
        self.win.finish_drag()
        self.assertEqual(self.win.pos(), origin + QPoint(2000, 50))
        self.assertTrue(self.win.isVisible())
        self.assertEqual(self.win.display_mode, "both")
        self.assertEqual(self.win.taskbar_screen, SECONDARY)
        self.assertTrue(controller._active)

    def test_dual_mode_retains_active_bar_when_target_cannot_attach(self):
        controller = self.multiscreen_controller()
        self.win.taskbar_x = 600
        self.win.set_display_mode("both")

        def present(area, *_args, **_kwargs):
            if area.device_name == SECONDARY:
                raise OSError("模拟嵌入失败")

        self.native.present.side_effect = present
        self.native_mouse("press", 100, 20, QPoint(700, 760), 1)
        target = QPoint(-1727, 760)
        self.native_mouse("move", 150, 20, target, 2, screen_x=-1310)
        self.assertTrue(controller._active)
        self.assertEqual(controller._area.device_name, PRIMARY)
        self.native_mouse("release", 150, 20, target, 2, screen_x=-1310)
        self.assertEqual(self.win.taskbar_screen, PRIMARY)
        self.assertEqual(self.win.display_mode, "both")
        self.assertTrue(self.win.isVisible())
        self.native.hide.assert_not_called()

    def test_enable_from_float_selects_its_native_monitor_and_retains_saved_target_on_refresh(self):
        controller = self.multiscreen_controller()
        self.win.taskbar_screen = SECONDARY
        self.win.set_display_mode("taskbar")
        self.finder.assert_called_with(window=int(self.win.winId()), device_name="", hwnd=None)
        self.assertEqual(self.win.taskbar_screen, PRIMARY)
        self.win.taskbar_screen = SECONDARY
        controller.refresh()
        self.finder.assert_called_with(window=int(self.win.winId()), device_name=SECONDARY, hwnd=None)
        self.assertEqual(self.win.taskbar_screen, SECONDARY)
        self.win.set_view_options(taskbar_enabled=False)
        self.win.set_view_options(taskbar_enabled=True)
        self.assertEqual(self.win.taskbar_screen, PRIMARY)
