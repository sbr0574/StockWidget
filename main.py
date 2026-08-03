# python3 -m PyInstaller -F -w .\main.py --name StockWidget --icon .\resources\StockWidget.ico
import sys
import platform
from src.App import App, APP_NAME

if __name__ == "__main__":

    if platform.system() == "Windows":
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(f"{APP_NAME}.1")

    app = App(sys.argv)
    sys.exit(app.exec())
