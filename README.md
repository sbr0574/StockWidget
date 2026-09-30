# StockWidget

简洁实时行情显示工具，浮窗支持自定义样式、双击隐藏、鼠标穿透。覆盖沪深京市场、港股、美股股票基金，全球主要指数和期货（上期所）数据，兼容**Windows、macOS、Linux**多平台。

[![Release](https://img.shields.io/badge/下载-Releases-blue?style=flat-square&logo=github)](https://github.com/sbr0574/StockWidget/releases) [![Version](https://img.shields.io/github/v/tag/sbr0574/StockWidget?sort=semver&label=版本&style=flat-square)](https://github.com/sbr0574/StockWidget/tags) ![License](https://img.shields.io/badge/License-Apache--2.0-lightgrey?style=flat-square)

[![Windows](https://img.shields.io/badge/Windows-supported-0078D4?style=flat-square&logo=windows&logoColor=white)](https://github.com/sbr0574/StockWidget/releases) [![macOS](https://img.shields.io/badge/macOS-supported-000000?style=flat-square&logo=apple&logoColor=white)](https://github.com/sbr0574/StockWidget/releases) [![Linux](https://img.shields.io/badge/Linux-supported-FCC624?style=flat-square&logo=linux&logoColor=black)](https://github.com/sbr0574/StockWidget/releases)

---

## 声明

* **作者**：`sbr0574`
* **仓库地址**：https://github.com/sbr0574/StockWidget
* **镜像仓库**: https://gitee.com/sbr0574/StockWidget
* StockWidget的**唯一发布渠道**为上方 GitHub/Gitee 仓库的 [Releases](https://github.com/sbr0574/StockWidget/releases)，其它渠道提供的下载请谨慎使用。
* 本项目基于 **Apache License 2.0** 开源。任何人对本项目进行**再分发**（转载源码、镜像下载、打包发布等）时，**必须**：

  * 保留 `LICENSE` 与 `NOTICE` 文件；
  * 保留版权与署名信息（详见 `NOTICE`）；
  * 注明原作者与仓库地址：https://github.com/sbr0574/StockWidget
* 如发现第三方网站转载时未注明上述信息，可要求对方补充。

---

## 🖼️ 界面一览

> **极简显示效果**

<img width="127" height="90" alt="image" src="https://github.com/user-attachments/assets/69ce08a7-6b55-41e8-bbab-e47b64d6e8a0" />

> **设置面板———支持自动浅色/深色模式切换**

<img width="393" height="279" alt="image" src="https://github.com/user-attachments/assets/5ec0170a-bc7a-495d-be36-d62812d4b155" />
<img width="393" height="279" alt="image" src="https://github.com/user-attachments/assets/ea60a6c2-970a-4e29-83d8-600847c21fee" />

> **自选股添加**

<img width="393" height="345" alt="image" src="https://github.com/user-attachments/assets/1083c49d-6ec1-41c9-8abd-44ecc97737c6" />

> **快捷添加自选（双击表格空白/双击条目编辑）**

<img width="393" height="332" alt="image" src="https://github.com/user-attachments/assets/cde49b06-2e14-40b8-9ff5-afe828781fbf" />

> **浮窗样式与功能设置**

<img width="393" height="279" alt="image" src="https://github.com/user-attachments/assets/16e94332-97ee-4707-a917-f55b440a94a8" />


## ✨ 功能概览

* **透明无框浮窗**：

  * 置顶显示
  * 拖拽移动
  * **双击**可隐藏
  * 快速排序
  * 右键快捷设置菜单
  * 可开启鼠标穿透
* **全局快捷键**（需要在设置中启用，可自定义修改）：

  * `Ctrl+Alt+F` 显示/隐藏浮窗
  * `Ctrl+Alt+C` 开启/关闭鼠标穿透
* **设置分类**：「外观」管理浮窗颜色、不透明度、字体行距和图标；「功能」管理常用开关、快捷键、行数限制、翻页和隐藏规则；「任务栏」独立管理任务栏模式、指标、样式和分页。
* **定时与自动隐藏**：在「设置 → 功能」启用「定时隐藏」，可添加或删除最多 3 个 `HH:mm` 时间，每日按电脑本地时间隐藏浮窗与任务栏，关闭开关保留时间。「自动隐藏」支持新浪和东财（`f124` 行情时间戳）；仅当刷新得到的所有勾选标的都有有效行情时间、且均超过 30 秒时触发，缺失数据或请求失败不触发。手动呼出后再次触发时显示 5 秒倒计时，新行情到达或关闭自动隐藏会取消倒计时；定时隐藏同时生效。隐藏后呼出恢复原位置与显示模式。
* **Windows 任务栏行情**：先在「设置 → 任务栏」组框标题勾选「启用任务栏模式」。关闭此总开关时，拖到任务栏或使用显示位置菜单都不会开启任务栏模式；开启后可拖入 / 拖出，托盘也可切换显示位置。拖出到浮窗后总开关仍保持开启。
  * 使用与 TrafficMonitor 相同思路的原生子窗口嵌入，无需安装 TrafficMonitor。
  * 在任务栏页选择 **1–4 行**，支持只显示前 N 行、手动翻页或自动翻页，可调整自动翻页间隔。翻页控件位于数据左侧，同列显示 `︿`、当前页/总页数、`﹀`，首尾循环；按住箭头或页码也可拖动，拖动不会翻页。
  * 指标数量不限，复用浮窗的双池选择、拖拽排序和双击移入 / 移出操作。「同步浮窗指标」小组框会实时跟随浮窗，关闭同步则恢复任务栏独立配置。「同步浮窗外观」跟随字体、字号、文字颜色、不透明度和统一颜色；取消勾选后可独立设置（包括涨跌颜色和 K 线）。背景透明，字号适应任务栏高度，过宽内容省略显示。
  * **单击数据不操作，双击隐藏**。浮窗与任务栏属于同一个行情组件，隐藏不会改变显示位置；托盘或全局快捷键会在原位置恢复。默认两种位置不共存，勾选「浮窗与任务栏同时显示」后可同时显示，双击任一处会一起隐藏和恢复。两处复用右键菜单，支持指标、排序、设置及隐藏操作。
  * **独占模式可在桌面与主屏任务栏间拖动**：浮窗进入任务栏会立即显示预览，移出即取消；在任务栏上松开鼠标则停靠。从任务栏按住拖出，会在鼠标附近显示浮窗；松开后改为浮窗位置，拖回任务栏仍可停靠。**双开模式任务栏固定显示，浮窗可在全屏自由拖动，经过任务栏区域也不会改变模式**。拖动中按 Esc 可取消并恢复原位置。
  * 可调整向左偏移，避开任务栏图标。此方式不为行情预留系统布局空间，任务栏拥挤时可能重叠；仅支持主屏横向任务栏。
  * Explorer 重启后自动重新嵌入；失败时暂时显示浮窗并自动重试。Windows 更新可能影响这种非官方任务栏扩展方式。
* **浮窗分页**：在「设置 → 功能」勾选「限制行数」，设置最大行数；在独立的「翻页」组框设置不翻页 / 手动 / 自动及间隔。关闭行数限制时显示全部行、不分页、不翻页，保留之前的参数供重新启用。任务栏显示行数可设为 1–4 行；「同步浮窗翻页」跟随浮窗的翻页方式和自动间隔，取消勾选后可独立设置不翻页 / 手动 / 自动及间隔。浮窗与任务栏分别记录页码，采用相同的标的排序；行情更新保留当前页，重新排序或调整自选列表时回到第一页。
* **左右分栏**：在「设置 → 功能」勾选「分栏」，将当前页标的按顺序均分，前半放左栏、后半放右栏；奇数时左栏多一条。「显示分隔线」控制中间竖线。行数限制按每栏计算，例如 3 行可显示最多 6 个标的，两栏一起手动或自动翻页。关闭行数限制后仍将全部标的均分。任务栏的「同步浮窗分栏」可跟随浮窗，取消同步后可独立设置分栏和分隔线，独立参数在同步期间保留。
* **显示指标**（在设置中拖动指标，可自定义显示内容和顺序）：

  * 名称（点击ⓘ设置市场、代码显示及名称显示字数）
  * 现价（默认，**现价触及当日最高/最低**时显示 `↑ / ↓`）
  * 涨跌值
  * 涨跌幅（默认）
  * 浮盈
  * 买一卖一数量
  * 委比
  * 成交量（点击ⓘ设置数字单位）
  * 成交额（点击ⓘ设置数字单位）
  * 均价
  * K线（当日）
* **指标排序**：显示表头后，单击表头开启排序，重复点击按“**降序-升序-不排序**”循环，支持排序的列有**现价、涨跌、涨幅、浮盈、成交量、成交额、均价、委比**。右键菜单也可快捷排序。
* **颜色设置**：可分别设置背景、文字、上涨、下跌和中性颜色；**背景可调不透明度**，且可单独设置**窗口整体不透明度**。
* **统一颜色**：默认开启，所有内容使用文字颜色；关闭后可分别设置**上涨、下跌和中性颜色**。
* **自选列表管理**：程序内置全市场代码数据，并定期更新；双击列表空白区域可快速添加，点击增加按钮可按**股票、基金、指数、期货**及**沪、深、京、港、美、其他**组合筛选并分页添加；搜索支持由空格分隔的**数字代码、名称、拼音或缩写**关键词，双击已有条目可修改。
* **程序图标**：自带4个图标可选，选择即时生效；也可上传自定义图片、图标作为程序图标。


## 常见问题

* **Windows下任务栏与浮窗置顶冲突**：在设置中打开**强制置顶**即可，开启后浮窗同时会阻挡截屏等其他置顶窗口。

## 🧰 下载与运行

右侧 [Releases](https://github.com/sbr0574/StockWidget/releases) 提供按版本号命名的三平台压缩包：

| 平台 | 发布包 | 使用方式 |
| --- | --- | --- |
| Windows 10/11 | `StockWidget-Windows-<版本号>.zip` | 解压后运行 `StockWidget.exe` |
| macOS | `StockWidget-macOS-<版本号>.zip` | 解压后将 `StockWidget.app` 拷贝至“应用程序”文件夹并运行 |
| Linux | `StockWidget-Linux-<版本号>.zip` | 解压后在文件夹内终端运行`./StockWidget`；Linux还提供`.deb`和`.rpm`安装包，按需使用 |

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

设置页、添加标的面板、指标池、名称/单位设置和浮窗布局均提供 Qt Designer `.ui` 文件；修改后运行 `python scripts/generate_ui.py` 生成对应代码。具体文件和调整方法见 [手动调整界面](docs/UI_EDITING.md)。


## 🌐 数据来源 & 网络

* 行情信息通过 `requests` 从 **新浪财经**或**东方财富**接口获取，数据不可避免存在一定延迟，仅供参考。
* 程序仅显示行情数据，不包含任何账户/交易操作。
* 请根据自身网络环境决定是否使用代理或更换数据源。
* 行情组件隐藏时暂停刷新和自动翻页，显示后自动恢复；仅隐藏浮窗外观、仍在任务栏显示时继续刷新。


## 📜 许可

本项目基于 **Apache License 2.0** 发布，完整文本见 [LICENSE](LICENSE)，署名要求见 [NOTICE](NOTICE)。

* 个人/学习用途自由使用；涉及第三方数据源时请遵守其使用条款。
* **再分发（转载/镜像/打包发布）必须保留 `LICENSE` 与 `NOTICE` 文件，注明原始作者与仓库地址**（https://github.com/sbr0574/StockWidget）。
