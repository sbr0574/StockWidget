# -*- coding: utf-8 -*-
"""全局常量：应用名/版本、配置文件、市场代码列表文件与下载地址。"""

APP_NAME = "StockWidget"
APP_VERSION = "1.4.0"
CONFIG_FILE = "stock_widget_config.json"

# 分类代码列表（服务器更新到独立数据分支，客户端按文件独立选择版本）。
CODE_LIST_FILES = (
    "stock_sh.json",
    "stock_sz.json",
    "stock_bj.json",
    "fund_cn.json",
    "stock_hk.json",
    "stock_us.json",
    "index_cn.json",
    "index_global.json",
    "futures_sh.json",
)
CODES_STATUS_FILE = "codes_update_status.json"
CODES_RAW_URLS = (
    "https://raw.githubusercontent.com/sbr0574/StockWidget/codes-data/resources/{name}",
    "https://raw.giteeusercontent.com/sbr0574/StockWidget/raw/codes-data/resources/{name}",
)
