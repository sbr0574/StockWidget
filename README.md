# StockWidget

极简透明盯盘 Widget 浮窗，实时显示自选股行情，支持沪深京股票基金、港股、美股、全球主要指数、期货（上期所），兼容**Windows、macOS、Linux**多平台。支持调整浮窗样式、双击隐藏浮窗、鼠标穿透，配置自动保存。

[![Release](https://img.shields.io/badge/下载-Releases-blue?style=flat-square&logo=github)](https://github.com/sbr0574/StockWidget/releases) [![Version](https://img.shields.io/github/v/tag/sbr0574/StockWidget?sort=semver&label=版本&style=flat-square)](https://github.com/sbr0574/StockWidget/tags) ![License](https://img.shields.io/badge/License-Apache--2.0-lightgrey?style=flat-square)

[![Windows](https://img.shields.io/badge/Windows-supported-0078D4?style=flat-square&logo=windows&logoColor=white)](https://github.com/sbr0574/StockWidget/releases) [![macOS](https://img.shields.io/badge/macOS-supported-000000?style=flat-square&logo=apple&logoColor=white)](https://github.com/sbr0574/StockWidget/releases) [![Linux](https://img.shields.io/badge/Linux-supported-FCC624?style=flat-square&logo=linux&logoColor=black)](https://github.com/sbr0574/StockWidget/releases)

---

## 声明

* **作者**：`sbr0574`
* **仓库地址**：https://github.com/sbr0574/StockWidget
* 本项目的**唯一发布渠道**为上方 GitHub 仓库的 [Releases](https://github.com/sbr0574/StockWidget/releases)，其它网站提供的下载链接请谨慎使用。
* 本项目基于 **Apache License 2.0** 开源。任何人对本项目进行**再分发**（转载源码、镜像下载、打包发布等）时，**必须**：

  * 保留 `LICENSE` 与 `NOTICE` 文件；
  * 保留版权与署名信息（详见 `NOTICE`）；
  * 注明原作者与仓库地址：https://github.com/sbr0574/StockWidget
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
* **统一颜色**：默认开启，所有内容使用文字颜色；关闭后可分别设置**上涨、下跌和中性颜色**。
* **浮窗颜色**：可分别设置背景、文字、上涨、下跌和中性颜色；**背景可透明**，且可单独设置**整体不透明度**。
* **列开关与表头显示**：右键浮窗 → “显示列”“显示表头”即时生效。
* **字体与行距**：字号 **5–15 pt**；行距为额外像素（行高 = 字高 + 行距），**K 线尺寸随字号同步缩放**。
* **刷新间隔**：可选 **1-15** 秒。
* **股票代码管理**：设置面板内用列表**增加/删除/上移/下移/置顶**，每日首次启动自动从 GitHub/Gitee 下载由 Actions 更新的全市场代码列表；双击列表空白区域可在全部标的中快速搜索，点击增加按钮可按**股票、基金、指数、期货**及**沪、深、京、港、美、其他**组合筛选并分页添加；搜索支持由空格分隔的**数字代码、名称、拼音或缩写**关键词，双击已有条目可修改。
* **自动保存**：所有设置即时保存至配置文件（Windows：`%APPDATA%\StockWidget\stock_widget_config.json`；macOS/Linux：`~/StockWidget/stock_widget_config.json`）；浮窗隐藏时**暂停刷新**，显示时自动恢复。
* **自动检查更新**：启动时优先检查 GitHub Releases，GitHub 不可用时切换 Gitee，有新版本时提示下载。

---

## 🖼️ 界面一览

> 📷 **极简显示效果**
<img width="127" height="90" alt="image" src="https://github.com/user-attachments/assets/69ce08a7-6b55-41e8-bbab-e47b64d6e8a0" />

> 📷 **设置面板**
<img width="360" height="259" alt="image" src="https://github.com/user-attachments/assets/99e9e5e9-79fb-4708-a75a-3210923d64f4" />
<img width="360" height="259" alt="image" src="https://github.com/user-attachments/assets/ec371889-3046-423f-b07c-9a34df018a6f" />

> 📷 **显示全部指标+涨跌颜色**
<img width="426" height="57" alt="image" src="https://github.com/user-attachments/assets/95fa0e99-0b0c-4fed-803d-aa3b78324b2f" />

---

## 📁 项目结构

代码按职责分层，前端（界面）与后端（数据/逻辑）分离：

```
main.py                      # 程序入口
StockWidget.spec             # PyInstaller 打包配置
resources/                   # 静态资源
stockwidget/
  app.py                     # 应用装配：连接各层、托盘、后台任务
  constants.py               # 全局常量（名称/版本/文件/地址）
  ui/                        # 界面层：所有 Qt 组件与显示
    widget.py                #   盯盘浮窗主面板
    settings_dialog.py       #   设置面板
    table_model.py           #   表格 Model 与 K 线 Delegate
    drag_mixin.py            #   拖拽 / 双击隐藏交互
    tray.py                  #   系统托盘
    generated/               #   Qt Designer / pyside6-uic 生成文件
  data/                      # 数据层：行情请求与整理
    quotes.py                #   行情请求与解析（新浪 / 东财）
    code_lists.py            #   代码列表下载 / 缓存 / 兜底
    update_check.py          #   版本更新检查
  core/                      # 功能函数层：纯业务逻辑
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

## 🧰 下载与运行

右侧 [Releases](https://github.com/sbr0574/StockWidget/releases) 提供按版本号命名的三平台压缩包：

| 平台 | 发布包 | 使用方式 |
| --- | --- | --- |
| Windows 10/11 | `StockWidget-Windows-<版本号>.zip` | 解压后运行 `StockWidget.exe` |
| macOS | `StockWidget-macOS-<版本号>.zip` | 解压后将 `StockWidget.app` 拖入“应用程序” |
| Linux（Ubuntu/Debian） | `StockWidget-Linux-<版本号>.zip` | 解压后安装其中的 `.deb` 包 |

### 源码运行

支持 Windows、macOS 和 Linux，发布构建与 CI 使用 Python **3.13**。安装依赖并生成 Qt 资源文件：

```bash
pip install -r requirements.txt
pyside6-rcc resources/resources.qrc -o resources/resources_rc.py
```

脚本式运行：

```bash
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

* 行情通过 `requests` 从 **新浪财经**接口（`hq.sinajs.cn`）获取；沪深股票和基金代码列表直接从交易所官方接口更新，其他市场沿用东财或新浪接口。
* 程序仅发起 GET 请求，不包含任何账户/交易操作；请根据自身网络环境决定是否使用代理或更换数据源。
* 浮窗隐藏时会暂停刷新，显示后自动恢复，减少不必要的请求。

---

## 📜 许可

本项目基于 **Apache License 2.0** 发布，完整文本见 [LICENSE](LICENSE)，署名要求见 [NOTICE](NOTICE)。

* 个人/学习用途自由使用；涉及第三方数据源时请遵守其使用条款。
* **再分发（转载/镜像/打包发布）必须保留 `LICENSE` 与 `NOTICE` 文件，注明原始作者与仓库地址**（https://github.com/sbr0574/StockWidget）。
