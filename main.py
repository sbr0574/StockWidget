
# -*- coding: utf-8 -*-
"""
StockWidget
极简透明盯盘 Widget 浮窗, 按指定股票代码实时显示行情表格

Copyright © 2026 sbr0574

官方仓库地址 Official website:
https://github.com/sbr0574/StockWidget

邮箱 Email:
sbr0574@qq.com

本项目基于 Apache License 2.0 开源协议发布, 请遵守相关协议
Licensed under Apache License 2.0
"""

import sys

from stockwidget.app import App
from stockwidget.constants import APP_NAME

if __name__ == "__main__":

    if sys.platform == "win32":
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(f"{APP_NAME}.1")

    app = App(sys.argv)
    sys.exit(app.exec())
