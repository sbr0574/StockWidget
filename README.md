# StockWidget

一个在 **Windows** 上运行的极简透明盯盘 Widget 浮窗，按指定股票代码实时显示行情表格，可选展示迷你 **K 线**（当天），支持沪深京三市股票、基金、指数。支持拖拽、右键菜单、设置面板、全局快捷键隐藏浮窗、鼠标穿透、自动保存配置。

> 适合贴在屏幕一角随时查看 👀

[![Release](https://img.shields.io/badge/下载-Releases-blue?style=flat-square&logo=github)](https://github.com/sbr0574/StockWidget/releases) ![Version](https://img.shields.io/badge/版本-1.3.0-green?style=flat-square) ![License](https://img.shields.io/badge/License-Apache--2.0-lightgrey?style=flat-square)

---

## 声明

* **作者**：`sbr0574`
* **官方仓库**：https://github.com/sbr0574/StockWidget
* 本项目的**唯一官方发布渠道**为上方 GitHub 仓库的 [Releases](https://github.com/sbr0574/StockWidget/releases)，其它网站提供的下载链接均非官方，请谨慎使用。
* 本项目基于 **Apache License 2.0** 开源。任何人对本项目进行**再分发**（转载源码、镜像下载、打包发布等）时，**必须**：

  * 保留 `LICENSE` 与 `NOTICE` 文件；
  * 保留版权与署名信息（详见 `NOTICE`）；
  * 在显著位置注明原始作者与仓库地址：https://github.com/sbr0574/StockWidget
* 如发现第三方网站转载时未注明上述信息，可先联系对方补充；若对方拒绝，其行为已违反 Apache 2.0 第 4 节（Redistribution）的条款，作者有权要求其停止分发。

---

## ✨ 功能概览

* **透明无框浮窗**：

  * 置顶显示（默认置顶策略在任务栏等位置存在置顶冲突，可在设置中开启**强制置顶**）
  * 拖拽任意区域即可移动
  * **双击**浮窗可隐藏
  * 右键展示设置菜单
  * 可选鼠标穿透
* **系统托盘**：左键切换显示/隐藏；右键菜单含“设置 / 退出”。
* **全局快捷键**：`Ctrl+Alt+F` 显示/隐藏浮窗，`Ctrl+Alt+C` 切换鼠标穿透（均可在设置中自定义）。
* **表格展示**（可选列）：`名称 | 现价（默认） | 涨跌值 | 涨跌幅（默认） | 浮盈 | 买一卖一数量 | 委比 | 成交量 | 成交额 | 均价 | K线`

  * **现价触及当日最高/最低**时显示 `↑ / ↓`
* **默认颜色**：开启后自动 **红涨绿跌**；关闭则为 **单色模式**（按自定义的文字颜色）。
* **浮窗颜色**：**背景可透明**，且可单独设置**整体不透明度**。
* **列开关与表头显示**：右键浮窗 → “显示列”“显示表头”即时生效。
* **字体与行距**：字号 **5–15 pt**；行距为额外像素（行高 = 字高 + 行距），**K 线尺寸随字号同步缩放**。
* **刷新间隔**：可选 **1-15** 秒。
* **股票代码管理**：设置面板内用列表**增加/删除/上移/下移/置顶**，每日首次启动自动从 GitHub/Gitee 下载由 Actions 更新的全市场代码列表，自选股可通过增加按钮或双击空白区域添加，双击条目可修改，支持输入**数字代码、拼音、首字母、中文名**进行匹配搜索。
* **自动保存**：所有设置即时保存至配置文件（`%APPDATA%\StockWidget\stock_widget_config.json`）；浮窗隐藏时**暂停刷新**，显示时自动恢复。
* **自动检查更新**：启动时优先检查 GitHub Releases，GitHub 不可用时切换 Gitee，有新版本时提示下载。

---

## 🖼️ 界面一览

> 📷 **极简显示效果**
<img width="127" height="90" alt="image" src="https://github.com/user-attachments/assets/69ce08a7-6b55-41e8-bbab-e47b64d6e8a0" />

> 📷 **设置面板**
<img width="360" height="259" alt="image" src="https://github.com/user-attachments/assets/99e9e5e9-79fb-4708-a75a-3210923d64f4" />
<img width="360" height="259" alt="image" src="https://github.com/user-attachments/assets/ec371889-3046-423f-b07c-9a34df018a6f" />

> 📷 **显示全部指标+默认颜色**
<img width="426" height="57" alt="image" src="https://github.com/user-attachments/assets/95fa0e99-0b0c-4fed-803d-aa3b78324b2f" />

---

## 📁 项目结构

代码按职责分层，前端（界面）与后端（数据/逻辑）分离：

```
main.py                      # 程序入口
StockWidget.spec             # PyInstaller 打包配置
resources/                   # 静态资源（图标、内置代码列表、Qt 资源）
tests/                       # 单元测试（python -m unittest discover -s tests）
stockwidget/
  app.py                     # 应用装配：连接各层、托盘、后台任务
  constants.py               # 全局常量（名称/版本/文件/地址）
  ui/                        # 界面层（前端）：所有 Qt 组件与显示
    widget.py                #   盯盘浮窗主面板
    settings_dialog.py       #   设置面板
    table_model.py           #   表格 Model 与 K 线 Delegate
    drag_mixin.py            #   拖拽 / 双击隐藏交互（混入）
    tray.py                  #   系统托盘（平台差异的点击行为）
    generated/               #   Qt Designer / pyside6-uic 生成文件（勿手改）
  data/                      # 数据层（后端）：行情请求与整理
    quotes.py                #   行情请求与解析（新浪 / 东财）
    code_lists.py            #   代码列表下载 / 缓存 / 兜底
    update_check.py          #   版本更新检查
    # 代码列表生成已合并到 .github/scripts/update_codes.py（仅 CI 使用）
  core/                      # 功能函数层：纯业务逻辑（无 Qt，可单元测试）
    formatters.py            #   成交量 / 成交额格式化
    code_search.py           #   代码搜索 / 建议
    watchlist.py             #   自选列表规范化
    config_store.py          #   配置读写
    geometry.py              #   多显示器位置恢复
  platform/                  # 平台适配层：跨平台原生实现
    capabilities.py          #   能力探测（X11/Wayland 等）
    click_through.py         #   鼠标穿透
    autostart.py             #   开机自启
    hotkeys.py               #   全局快捷键
```

分层原则：`ui` 只负责显示与交互，`data` 只负责取数与解析，`core` 是可独立测试的纯函数，`platform` 隔离平台差异；各层通过 `app.py` 装配连接，避免职责互相缠绕。

## 🧰 运行环境

右侧 [Releases](https://github.com/sbr0574/StockWidget/releases) 已有打包好的程序（`StockWidget-windows.zip`），**直接下载，解压后运行 StockWidget.exe 使用**。

若要通过代码脚本形式运行，则需要：
* Windows 10/11
* Python **3.12+** （其他 Python 版本暂未测试）
* 依赖见 `requirements.txt`（`PySide6` 界面、`requests` 拉取行情等）
```powershell
pip install -r requirements.txt
```
脚本式运行：
```powershell
python main.py
```

---

## 🖱️ 操作速览

* **拖动窗口**：按住窗口任意位置拖动。
* **双击浮窗**：隐藏。
* **右键浮窗**：快捷菜单。
* **全局快捷键**：

  * `Ctrl+Alt+F`：显示/隐藏浮窗
  * `Ctrl+Alt+C`：切换鼠标穿透
* **系统托盘**：

  * 左键：显示/隐藏浮窗
  * 右键：设置 / 退出

---

## 🌐 数据来源 & 网络

* 行情通过 `requests` 从 **新浪财经**接口（`hq.sinajs.cn`）获取；股票代码列表通过 **AkShare** 更新。
* 程序仅发起 GET 请求，不包含任何账户/交易操作；请根据自身网络环境决定是否使用代理或更换数据源。
* 浮窗隐藏时会暂停刷新，显示后自动恢复，减少不必要的请求。

---

## 📜 许可

本项目基于 **Apache License 2.0** 发布，完整文本见 [LICENSE](LICENSE)，署名要求见 [NOTICE](NOTICE)。

* 个人/学习用途自由使用；涉及第三方数据源时请遵守其使用条款。
* **再分发（转载/镜像/打包发布）必须保留 `LICENSE` 与 `NOTICE` 文件，并注明原始作者与仓库地址**（https://github.com/sbr0574/StockWidget），否则视为违反 Apache 2.0 第 4 节的再分发条款。
* 官方唯一发布渠道：https://github.com/sbr0574/StockWidget
