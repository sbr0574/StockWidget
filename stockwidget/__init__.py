# -*- coding: utf-8 -*-
"""
StockWidget
极简透明盯盘 Widget 浮窗，按指定股票代码实时显示行情表格。

包结构（分层职责清晰分离）：
- ``stockwidget.ui``       界面层：所有 Qt 组件与显示（浮窗、设置面板、表格 Model/Delegate）。
- ``stockwidget.data``     数据层：行情请求与解析、代码列表下载、更新检查。
- ``stockwidget.core``     功能函数层：纯业务逻辑（格式化、代码搜索、配置读写、市场代码约定）。
- ``stockwidget.platform`` 平台适配层：能力探测、鼠标穿透、开机自启、全局快捷键。
- ``stockwidget.app``      应用装配：创建各层对象并连接信号，不含界面/数据细节。

Copyright © 2026 sbr0574 | Apache License 2.0
"""
from stockwidget.constants import APP_NAME, APP_VERSION

__all__ = ["APP_NAME", "APP_VERSION"]
__version__ = APP_VERSION
