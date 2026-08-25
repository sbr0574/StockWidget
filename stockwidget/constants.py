# -*- coding: utf-8 -*-
"""全局常量：应用名/版本、配置文件、市场代码列表文件与下载地址。"""

APP_NAME = "StockWidget"
APP_VERSION = "1.4.0"
CONFIG_FILE = "stock_widget_config.json"

# 全市场代码列表：证券与期货两个 JSON（GitHub Action 更新到独立数据分支）
LIST_FILES = ("stock_codes_list.json", "futures_codes_list.json")
CODES_RAW_URL = "https://raw.githubusercontent.com/sbr0574/StockWidget/{branch}/resources/{name}"
CODES_RAW_BACKUP = "https://raw.giteeusercontent.com/sbr0574/StockWidget/raw/{branch}/resources/{name}"
CODES_BRANCHES = ("codes-data",)
