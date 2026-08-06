
# -*- coding: utf-8 -*-
"""
StockWidget (Project by @sbr0574)
极简透明盯盘 Widget 浮窗, 按指定股票代码实时显示行情表格

Repository: https://github.com/sbr0574/StockWidget
Update  : 2026-08-06
Version : 1.3.0
License : Apache-2.0 license
"""

import sys
import platform
from src.App import App, APP_NAME

if __name__ == "__main__":

    if platform.system() == "Windows":
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(f"{APP_NAME}.1")

    app = App(sys.argv)
    sys.exit(app.exec())
