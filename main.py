# filename: StockWidget.py
# python3 -m PyInstaller -F -w .\StockWidget.py --name StockWidget --icon .\StockWidget.ico --add-data ".\StockWidget.ico;."
import sys
import platform
from src.App import App, APP_NAME

if __name__ == "__main__":
    system = platform.system()

    if system == "win32":
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(f"{APP_NAME}.1")

    app = App(sys.argv)
    sys.exit(app.exec())
