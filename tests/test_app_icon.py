import os
import tempfile
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtGui import QColor, QIcon, QPixmap
from PySide6.QtWidgets import QApplication

from stockwidget.app import App, _load_custom_icon
from stockwidget.constants import CONFIG_FILE


class AppIconTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt_app = QApplication.instance() or QApplication([])

    def _write_icon(self, directory: str) -> str:
        path = os.path.join(directory, "custom.png")
        pixmap = QPixmap(24, 24)
        pixmap.fill(QColor("#3578e5"))
        self.assertTrue(pixmap.save(path))
        return path

    def test_load_custom_icon_accepts_image_and_rejects_invalid_path(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = self._write_icon(temp_dir)
            normalized, icon = _load_custom_icon(path)

            self.assertEqual(normalized, os.path.abspath(path))
            self.assertFalse(icon.isNull())

        normalized, icon = _load_custom_icon(path)
        self.assertEqual(normalized, "")
        self.assertTrue(icon.isNull())

    def test_set_custom_icon_applies_it_immediately(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = self._write_icon(temp_dir)
            app = SimpleNamespace(
                _icon_choice="default",
                _custom_icon_path="",
                setWindowIcon=Mock(),
                tray=Mock(),
            )

            self.assertTrue(App.set_custom_icon(app, path))

        self.assertEqual(app._icon_choice, "custom")
        self.assertEqual(app._custom_icon_path, os.path.abspath(path))
        app.setWindowIcon.assert_called_once()
        app.tray.setIcon.assert_called_once()

    def test_save_now_writes_custom_choice_and_path(self):
        app = SimpleNamespace(
            win=Mock(),
            _icon_choice="custom",
            _custom_icon_path=r"C:\icons\mine.png",
            _start_on_boot=False,
        )
        app.win.current_config.return_value = {"watchlist": {}}

        with patch("stockwidget.app.save_file") as save:
            App.save_now(app)

        saved_config, file_name = save.call_args.args
        self.assertEqual(file_name, CONFIG_FILE)
        self.assertEqual(saved_config["app_icon"], "custom")
        self.assertEqual(saved_config["custom_icon_path"], r"C:\icons\mine.png")

    def test_invalid_custom_path_falls_back_to_default(self):
        app = SimpleNamespace(
            _icon_choice="custom",
            _custom_icon_path=r"C:\missing\icon.png",
            find_icon=Mock(return_value=QIcon()),
            setWindowIcon=Mock(),
            tray=Mock(),
        )

        App.set_app_icon(app, "custom")

        self.assertEqual(app._icon_choice, "default")
        self.assertEqual(app._custom_icon_path, "")
        app.setWindowIcon.assert_called_once()
        app.tray.setIcon.assert_called_once()


if __name__ == "__main__":
    unittest.main()
