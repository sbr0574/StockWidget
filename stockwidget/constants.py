# -*- coding: utf-8 -*-
"""全局常量：应用名/版本、配置文件、市场代码列表文件与下载地址。"""

APP_NAME = "StockWidget"
APP_VERSION = "1.4.0"
CONFIG_FILE = "stock_widget_config.json"

# 全市场代码列表：三个独立 JSON（GitHub Action 每日更新后由程序下载）
LIST_FILES = ("stock_codes_list.json", "stock_codes_global.json", "stock_codes_futures.json")
CODES_RAW_URL = "https://raw.githubusercontent.com/sbr0574/StockWidget/{branch}/resources/{name}"
CODES_BRANCHES = ("main", "master")
