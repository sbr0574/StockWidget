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
                 TaskbarArea(2, 2560, 72, 2300, 144, SECONDARY))
        self.finder = self.enterContext(patch("stockwidget.ui.floating.taskbar.find_taskbar",
            side_effect=lambda **kw: areas[1] if kw["hwnd"] == 2 or kw["device_name"] == SECONDARY else areas[0]))
        return controller

    def gesture(self, controller, end="release"):
        press = QPoint(700, 760)
        controller._dispatch_pointer("press", 100, 20, press, 1)
        controller._dispatch_pointer("move", 150, 20, press + QPoint(200, 0), 2)
        self.assertEqual(self.win.taskbar_screen, SECONDARY)
        self.assertEqual(controller._dpi, 144)
        self.assertTrue(controller._active)
        controller._dispatch_pointer(end, 150, 20, press + QPoint(200, 0), 2)

    def test_drag_between_taskbars_commits_the_target_display_and_escape_restores(self):
        controller = self.multiscreen_controller()
        self.win.set_display_mode("taskbar")
        self.gesture(controller, "cancel")
        self.assertEqual(self.win.taskbar_screen, PRIMARY)
        self.assertEqual(self.win.display_mode, "taskbar")
        self.assertFalse(controller.input_timer.isActive())
        self.gesture(controller)
        self.assertEqual(self.win.current_config()["taskbar_screen"], SECONDARY)
        self.assertEqual(self.win.display_mode, "taskbar")
        self.assertFalse(self.win.isVisible())

    def test_native_dual_drag_changes_display_without_moving_the_float(self):
        controller = self.multiscreen_controller()
        self.win.set_display_mode("both")
        self.win.move(100, 100)
        origin = self.win.pos()
        self.gesture(controller)
        self.assertEqual(self.win.pos(), origin)
        self.assertTrue(self.win.isVisible())
        self.assertEqual(self.win.display_mode, "both")
        self.assertEqual(self.win.taskbar_screen, SECONDARY)
        self.gesture(controller, "cancel")
        self.assertEqual(self.win.pos(), origin)
        self.assertEqual(self.win.taskbar_screen, SECONDARY)

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
