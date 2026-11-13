<p align="center">
  <img src="resources/icons/StockWidget.png" width="100" alt="MyProject Logo">
</p>

<h1 align="center">StockWidget</h1>

<div align="center">

简洁的桌面行情显示工具，用置顶浮窗查看自选行情

支持**A股、港股、美股**的股票、基金、指数和上期所期货

兼容 **Windows、macOS 和 Linux**

<img width="303" height="140" alt="image" src="https://github.com/user-attachments/assets/e6fa4447-68a5-4dd6-a124-9e14bc8c8f4b" />

</div><br>

<div align="center">

[![Release](https://img.shields.io/badge/下载-Releases-blue?style=flat-square&logo=github)](https://github.com/sbr0574/StockWidget/releases) [![Downloads](https://img.shields.io/github/downloads/sbr0574/StockWidget/total?label=下载量&style=flat-square)](https://github.com/sbr0574/StockWidget/releases) [![Version](https://img.shields.io/github/v/tag/sbr0574/StockWidget?sort=semver&label=版本&style=flat-square)](https://github.com/sbr0574/StockWidget/tags) [![Stars](https://img.shields.io/github/stars/sbr0574/StockWidget?style=flat-square&label=Star)](https://github.com/sbr0574/StockWidget) ![License](https://img.shields.io/badge/License-Apache--2.0-lightgrey?style=flat-square)

[![Windows](https://img.shields.io/badge/Windows-supported-0078D4?style=flat-square&logo=windows&logoColor=white)](https://github.com/sbr0574/StockWidget/releases) [![macOS](https://img.shields.io/badge/macOS-supported-000000?style=flat-square&logo=apple&logoColor=white)](https://github.com/sbr0574/StockWidget/releases) [![Linux](https://img.shields.io/badge/Linux-supported-FCC624?style=flat-square&logo=linux&logoColor=black)](https://github.com/sbr0574/StockWidget/releases)

</div>

作者：[sbr0574](https://github.com/sbr0574)。官方仓库：[GitHub](https://github.com/sbr0574/StockWidget) · [Gitee 镜像](https://gitee.com/sbr0574/StockWidget)。如果这个工具对你有帮助，欢迎在 GitHub 仓库右上角点一个 **Star ⭐**，支持后续维护。

## 功能

### 自选与行情指标

自选列表支持添加 ***A 股***（沪深京）、***港股***、***美股*** 的个股、基金和指数，以及上期所**期货**。可按类别、地区**筛选添加**；支持**代码**、**名称**、**拼音**、**缩写**搜索，可用空格分隔多个**关键词**。

- 自选列表支持**双击空白快速添加**，**双击已有条目编辑**，**拖动调整顺序**；可选择在浮窗中显示的条目。

    <img width="469" height="349" alt="image" src="https://github.com/user-attachments/assets/54c3a42a-8496-475d-8ccf-3762e254f6cf" />

- 浮窗可选显示指标：**名称、现价、涨跌、涨幅、浮盈、一档数量、委比、成交量、成交额、均值、量比、迷你图**。在“数据 → 指标显示”添加、移除或排序。

    <img width="330" height="141" alt="image" src="https://github.com/user-attachments/assets/41207091-8124-4f2f-8794-9c17ce1b0102" />

- 现价触及当日最高或最低时显示 `↑` / `↓`。

- 单击表头**排序**，按“降序 → 升序 → 不排序”循环，右键菜单也可排序。

在“数据 → 指标设置”中可设置：

- **名称**字数可选择显示字数，可选是否显示代码、市场。

- **均值**为当日均值或 MA5 / 10 / 20 / 30 / 60；

- **迷你图**为日 K 线、走势或幅值图。走势图为五分钟粒度当日分时；幅值图用横线表示当日最低至最高涨幅，小圆点表示当前涨幅，多股同列。

    <img width="327" height="117" alt="image" src="https://github.com/user-attachments/assets/71a17194-8008-4245-a15a-1c0d8a71b994" />

### 行情图表

单击浮窗中的个股，可打开行情图表，展示 **当日分时** / **5日分时** / **30日K线**，可选**成交量**、**均线**、**五档盘口**数据显示。

- 图表颜色与主浮窗一致，不透明度可随主浮窗调整，也可独立设置；

    <img width="300" height="300" alt="image" src="https://github.com/user-attachments/assets/b3ad942d-d682-4b08-9ded-7ee17a59456c" />

- 图表默认临时显示，可拖动改变位置和大小，点击外部或按 ESC 可关闭；也可点击右上角**固定按钮**将图表**固定在屏幕上显示**；

- 点击图表右上角第二个按钮可**在独立窗口打开图表**，并展示包含个股换手率、量比、市盈率、市值、股本等主要指标，可通过上下按钮循环切换自选股；

    <img width="520" height="345" alt="image" src="https://github.com/user-attachments/assets/5b96e8d5-16df-400f-8901-ef5d17c5bf91" />


### 浮窗外观与操作

浮窗用于展示自选股行情指标，支持全区域**拖动**、**双击隐藏**、**鼠标穿透**。

- 支持**自定义颜色、不透明度、字体、字号、行距**，支持自定义涨跌色；

    <img width="336" height="264" alt="image" src="https://github.com/user-attachments/assets/941f157d-d9a7-4696-b4ed-f423ae9c9a2c" />

- 默认**置顶**显示；

- 右键打开**快捷菜单**；

- 支持**双栏**显示，支持**浮窗高度限制**和**翻页**；

    <img width="341" height="123" alt="image" src="https://github.com/user-attachments/assets/e3e3adc0-6ee3-4ded-aa69-2cce16e42b4b" />

- Windows 支持单击任务栏托盘图标显示 / 隐藏浮窗，也可在设置中隐藏托盘图标；

- Windows 支持将浮窗固定在任务栏，可单独设置外观和显示指标等。

    <img width="172" height="74" alt="image" src="https://github.com/user-attachments/assets/7cdc5c21-0f44-4035-a712-2df3aadb97fb" />

**全局快捷键**默认关闭，可在“快捷键”中启用，支持组合键或功能键。

| 默认快捷键 | 操作 |
| --- | --- |
| `Ctrl+Alt+F` | 显示 / 隐藏行情组件 |
| `Ctrl+Alt+C` | 开启 / 关闭鼠标穿透 |

### 数据目录与缓存

缓存目录用于储存软件配置、代码列表和图表缓存。默认位置：

- Windows : `%APPDATA%\StockWidget`，无权限时自动回退到程序所在文件夹；

- macOS : `~/Library/Application Support/StockWidget`；

- Linux : `~/StockWidget`。

目录位置可在“设置 - 常规页”中更改。

### 翻页与分栏

浮窗**可限制每栏行数**，超出限制后，可选不翻页、手动翻页或自动翻页（间隔 **1–60 秒**）。

开启**分栏**后，当前页按顺序分为**左右两栏**。

### Windows 任务栏显示

在“设置 → 任务栏”开启“任务栏行情模式”。

- 默认显示在浮窗所在屏幕**任务栏右侧空白处**，可设置**偏移值**，偏移正值向左、负值向右；

- 限制最多显示 **4 行**，支持翻页、分栏；指标、翻页和分栏可分别同步浮窗，或独立设置；

- 任务栏文字大小、颜色与不透明度独立设置；开启**自动文字颜色**，随系统深浅模式切换；

- 可将浮窗拖出任务栏，自由浮窗可再拖入任务栏停靠；也可在多显示器的任务栏之间拖动；

- 开启“**允许浮窗与任务栏同时显示**”后，任务栏固定不可拖动，浮窗可自由移动；

- 可隐藏托盘图标，随后通过行情右键菜单或快捷键操作程序。

### 自动隐藏与贴边收起

支持每天最多三个**定时隐藏**时间，也可在全部行情超过 30 秒未更新时，延迟 5 秒后**自动隐藏**。

“常规 → 窗口行为 → 保持在屏幕内”会将越界浮窗调整回屏幕内。“贴边收起”可在屏幕四边将浮窗缩为窄条，鼠标移入展开，移出收起；下侧收起会避开任务栏或 Dock。

Windows 需先开启“保持在屏幕内”。Linux - Wayland 下无法可靠手动定位，贴边收起不可用。

## 下载与使用

前往 [Releases 下载](https://github.com/sbr0574/StockWidget/releases)，选择适合系统的发布包。请使用官方仓库发布的文件。

| 平台与处理器 | 系统版本要求 | ZIP 发布包 |
| --- | --- | --- |
| Windows，Intel / AMD 64位 | Windows 10 1809及以上，或 Windows 11 | `StockWidget-Windows-amd64.zip` |
| Windows，ARM64 | Windows 11 ARM64 | `StockWidget-Windows-arm64.zip` |
| macOS，Intel | macOS 13 Ventura及以上 | `StockWidget-macOS-amd64.zip` |
| macOS，Apple Silicon（M系列） | macOS 13 Ventura及以上 | `StockWidget-macOS-arm64.zip` |
| Linux，Intel / AMD 64位 | Ubuntu 22.04及以上，或提供 glibc 2.35及以上的同类发行版 | `StockWidget-Linux-amd64.zip` |
| Linux，ARM64 | Ubuntu 24.04及以上，或提供 glibc 2.39及以上的同类发行版 | `StockWidget-Linux-arm64.zip` |

`amd64`、`x64`、`x86_64` 指同一种架构；`arm64` 是另一种架构，请同时匹配操作系统和处理器类型。

- Windows：解压后运行 `StockWidget.exe`，或在桌面和开始菜单创建快捷链接启动。
- macOS：解压后将 `StockWidget.app` 拷贝到“应用程序”并运行。
- Linux：解压后在该文件夹的终端运行 `./StockWidget`；两种架构同时提供同名的 `.deb`、`.rpm` 安装包，例如 `StockWidget-Linux-arm64.deb`。安装包会检查对应的 glibc 最低版本。

**若设备未在上表中支持，可尝试下载源码使用**，需要 python 3.12 以上版本：

先安装所需依赖：`pip install -r requirements.txt`

然后编译资源文件：`pyside6-rcc resources/resources.qrc -o resources/resources_rc.py`

最后直接运行 `python main.py` 或打包使用 `pyinstaller --clean --noconfirm StockWidget.spec`

### 第一次打开的安全提示

**Windows**：首次运行可能出现“Windows 已保护你的电脑”，这是 SmartScreen 对尚未建立信誉的下载程序的提示。确认文件来自上方官方 Releases 且你信任该程序后，点击“更多信息”→“仍要运行”。如果没有“仍要运行”，可能受 Smart App Control 或组织策略限制，请按系统提示处理；不要为此关闭整个系统的安全保护。可参阅 [Microsoft 的 SmartScreen 说明](https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/smartscreen-reputation)。

**macOS**：如果提示“无法验证开发者”或“Apple 无法检查 App 是否包含恶意软件”，先尝试打开一次，再进入“系统设置”→“隐私与安全性”，向下找到 StockWidget 的拦截提示，点击“仍要打开”，按提示确认。旧版 macOS 对应“系统偏好设置”→“安全性与隐私”。这只为该 App 添加例外，后续可正常双击打开。具体步骤见 [Apple 的打开 App 说明](https://support.apple.com/zh-cn/102445)。若提示 App 已损坏或将损坏电脑，先从官方 Releases 重新下载，仍有问题时提交反馈并附上系统版本和提示内容。

### 开始使用

1. 从托盘菜单打开“设置”，在“自选列表”中添加并勾选需要显示的标的。
2. 在“数据”中选择显示指标，拖动调整顺序；在“外观”中调整颜色、字体和透明度。
3. 按住浮窗任意区域拖动，双击隐藏；通过托盘菜单或已启用的快捷键恢复。
4. 自选较多时，开启行数限制、翻页或分栏。Windows 用户还可在“任务栏”中启用任务栏显示。

## 常见问题

**Windows下任务栏遮挡浮窗**

在“常规 → 窗口行为”中开启强制置顶，或打开任务栏行情模式开关，将行情显示固定在任务栏上。

**隐藏或开启鼠标穿透后如何操作？**

通过托盘菜单或已启用的快捷键恢复显示、关闭穿透。关闭托盘图标显示和快捷键时，不建议使用鼠标穿透

**为什么部分设置没有显示或无法使用？**

Linux X11 支持全局快捷键和鼠标穿透，而 Wayland 下这些功能及窗口不透明度不可用。以当前界面的选项和提示为准。

**行情无法刷新怎么办？**

查看浮窗错误提示，检查网络或已有代理设置。新浪支持的市场范围及实时、历史数据可用情况可能不同。反馈问题时请附上系统、程序版本、标的代码和错误提示。

## 数据说明

基本行情与图表来自 **新浪财经**，沪深北个股内外盘统计来自 **腾讯**，可能存在延迟、缺失或暂时不可用，仅供参考。程序只展示行情，不接入证券账户，不执行交易。使用第三方数据源时请遵守其使用条款。

## 许可与署名

本项目基于 **Apache License 2.0** 发布，完整文本见 [LICENSE](LICENSE)，署名信息见 [NOTICE](NOTICE)。再分发时必须保留许可、版权与署名信息，注明原作者 **sbr0574** 和[原始仓库](https://github.com/sbr0574/StockWidget)。
