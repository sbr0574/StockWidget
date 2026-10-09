# StockWidget

简洁的桌面行情显示工具，用透明浮窗查看自选行情，也可在 Windows 任务栏中显示。支持股票、基金、全球主要指数和上期所期货，兼容 **Windows、macOS 和 Linux**。

[![Release](https://img.shields.io/badge/下载-Releases-blue?style=flat-square&logo=github)](https://github.com/sbr0574/StockWidget/releases) [![Downloads](https://img.shields.io/github/downloads/sbr0574/StockWidget/total?label=下载量&style=flat-square)](https://github.com/sbr0574/StockWidget/releases) [![Version](https://img.shields.io/github/v/tag/sbr0574/StockWidget?sort=semver&label=版本&style=flat-square)](https://github.com/sbr0574/StockWidget/tags) [![Stars](https://img.shields.io/github/stars/sbr0574/StockWidget?style=flat-square&label=Star)](https://github.com/sbr0574/StockWidget) ![License](https://img.shields.io/badge/License-Apache--2.0-lightgrey?style=flat-square)

[![Windows](https://img.shields.io/badge/Windows-supported-0078D4?style=flat-square&logo=windows&logoColor=white)](https://github.com/sbr0574/StockWidget/releases) [![macOS](https://img.shields.io/badge/macOS-supported-000000?style=flat-square&logo=apple&logoColor=white)](https://github.com/sbr0574/StockWidget/releases) [![Linux](https://img.shields.io/badge/Linux-supported-FCC624?style=flat-square&logo=linux&logoColor=black)](https://github.com/sbr0574/StockWidget/releases)

作者：[sbr0574](https://github.com/sbr0574)。官方仓库：[GitHub](https://github.com/sbr0574/StockWidget) · [Gitee 镜像](https://gitee.com/sbr0574/StockWidget)。如果这个工具对你有帮助，欢迎在 GitHub 仓库右上角点一个 **Star ⭐**，支持后续维护。

## 下载与使用

前往 [Releases 下载](https://github.com/sbr0574/StockWidget/releases)，选择适合系统的发布包。请使用官方仓库发布的文件。

| 平台 | 发布包 | 使用方式 |
| --- | --- | --- |
| Windows 10 / 11 | `StockWidget-Windows-<版本号>.zip` | 完整解压后运行 `StockWidget.exe` |
| macOS | `StockWidget-macOS-<版本号>.zip` | 解压后将 `StockWidget.app` 拷贝到“应用程序”并运行 |
| Linux | `StockWidget-Linux-<版本号>.zip` | 解压后在该文件夹的终端运行 `./StockWidget`；也可选择 `.deb` 或 `.rpm` 安装包 |

### 第一次打开的安全提示

**Windows**：首次运行可能出现“Windows 已保护你的电脑”，这是 SmartScreen 对尚未建立信誉的下载程序的提示。确认文件来自上方官方 Releases 且你信任该程序后，点击“更多信息”→“仍要运行”。如果没有“仍要运行”，可能受 Smart App Control 或组织策略限制，请按系统提示处理；不要为此关闭整个系统的安全保护。可参阅 [Microsoft 的 SmartScreen 说明](https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/smartscreen-reputation)。

**macOS**：如果提示“无法验证开发者”或“Apple 无法检查 App 是否包含恶意软件”，先尝试打开一次，再进入“系统设置”→“隐私与安全性”，向下找到 StockWidget 的拦截提示，点击“仍要打开”，按提示确认。旧版 macOS 对应“系统偏好设置”→“安全性与隐私”。这只为该 App 添加例外，后续可正常双击打开。具体步骤见 [Apple 的打开 App 说明](https://support.apple.com/zh-cn/102445)。若提示 App 已损坏或将损坏电脑，先从官方 Releases 重新下载，仍有问题时提交反馈并附上系统版本和提示内容。

### 开始使用

1. 从托盘菜单打开“设置”，在“自选列表”中添加并勾选需要显示的标的。
2. 在“数据”中选择显示指标，拖动调整顺序；在“浮窗”中调整颜色、字体和透明度。
3. 按住浮窗任意区域拖动，双击隐藏；通过托盘菜单或已启用的快捷键恢复。
4. 自选较多时，开启行数限制、翻页或分栏。Windows 用户还可在“任务栏”中启用任务栏显示。

## 功能

### 自选与行情指标

支持沪、深、京、港、美等市场的股票、基金，以及指数和上期所期货。添加时可按类别、地区筛选；搜索支持代码、名称、拼音、缩写和空格分隔的多个关键词。

- 双击自选列表空白处快速添加，双击已有条目编辑；拖动调整自选顺序。
- 内置市场代码列表并自动检查更新。
- 可选指标：名称、现价、涨跌、涨幅、浮盈、买一 / 卖一数量、委比、成交量、成交额、均价、当日日 K 线。
- 拖动或双击“可用指标”添加；拖动“已显示指标”调整顺序，双击或拖回目录移除。
- 在“数据 → 指标显示”设置名称字数、代码、类型与成交单位，以及浮窗表头。
- 现价触及当日最高或最低时显示 `↑` / `↓`。
- 单击数值表头按“降序 → 升序 → 不排序”循环，右键菜单也可排序；取消后恢复原自选顺序。

### 行情图表

在“数据 → 分时 / K线图”开启图表后，单击浮窗或任务栏的一行即可查看该标的。支持当日分时、5日分时和最近30个交易日的日K线，鼠标移入图表可查看数值。

可选择独立窗口或大、中、小半透明浮窗。图表浮窗靠近行情显示，点击外部、按 `Esc` 或点击关闭按钮即可关闭。分时均线、成交量和 MA5 / 10 / 20 / 30 / 60 均线可分别开关。

图表跟随行情刷新，并保存在本地以便再次查看。选择东方财富时保持东财来源；选择新浪时，历史图表必要时使用备用源。数据均为未复权，休市时显示最近可用交易日。部分标的历史可能不可用或不足5日，图表会提示；历史不足时，均线从满足周期的位置开始显示。

在“关于”点击“清理数据缓存”，可清除图表缓存，下次查看时重新加载。

### 浮窗外观与操作

透明无框浮窗支持全区域拖动、双击隐藏、右键快捷设置和鼠标穿透，大小随行情内容调整。

- 可设置颜色、不透明度、字体、字号、行距、表头和网格；关闭“统一颜色”后显示各自的涨跌颜色。
- “浮窗置顶”默认开启，Windows 可进一步开启“强制置顶”。
- 在“常规”选择跟随系统、浅色或深色模式，设置开机启动或图标。
- 浮窗与任务栏右键菜单可调整指标、排序、网格、统一颜色和分栏，并提供设置与退出入口。任务栏菜单保持紧凑，不提供表头开关。
- 托盘菜单可显示 / 隐藏行情、关闭鼠标穿透、打开帮助或提交问题；Windows 下也可切换任务栏行情。

全局快捷键需在“快捷键”中启用，支持自定义组合，并显示注册结果。

| 默认快捷键 | 操作 |
| --- | --- |
| `Ctrl+Alt+F` | 显示 / 隐藏行情组件 |
| `Ctrl+Alt+C` | 开启 / 关闭鼠标穿透 |

### 翻页与分栏

浮窗可限制每栏 **1–20 行**，选择只显示前 N 行、手动翻页或自动翻页（间隔 **1–60 秒**）。手动翻页使用数据左侧的箭头，首尾循环。关闭行数限制后显示全部标的，原设置保留。

开启分栏后，当前页按顺序分为左右两栏，可选择分隔线。例如每栏限制3行时，一页最多显示6个标的。

### Windows 任务栏显示

在“设置 → 任务栏”开启“任务栏行情模式”，立即切换到任务栏显示。

- 每栏 **1–4 行**，支持翻页、分栏、网格和向左偏移。
- 指标、外观、翻页和分栏可分别同步浮窗，或保留独立设置。
- 独立外观支持字体、字号、文字颜色、不透明度与统一颜色；“自动文字颜色”随系统深浅模式切换。字号上限随任务栏高度、行数与显示缩放调整。
- 将浮窗拖入任务栏停靠，拖出恢复浮窗；松开确认，`Esc` 取消。开启“允许浮窗与任务栏同时显示”后，两处一起显示。
- 可隐藏托盘图标，随后通过行情右键菜单或快捷键操作程序。

目前仅支持 Windows 主屏横向任务栏。拥挤时可能与图标重叠，可调整偏移；Explorer 重启后会尝试恢复，嵌入失败时暂时回到浮窗并重试。

### 自动隐藏与贴边收起

“常规 → 自动隐藏”支持每天最多三个定时隐藏时间，以及全部行情超过30秒未更新时自动隐藏；手动呼出后，旧行情再次触发时会先显示5秒倒计时。

“浮窗 → 贴边收起”可在屏幕左、上、右边将浮窗缩为窄条，鼠标移入展开，移出收起。Windows 需先开启“边界检测”；macOS 和 Linux X11 可直接开启。Wayland 下无法可靠手动定位，贴边收起不可用。

行情组件完全隐藏后暂停刷新和翻页，恢复后继续；仅浮窗贴边收起、任务栏仍显示时继续刷新行情。

## 常见问题

**Windows下任务栏遮挡浮窗怎么办**

在设置-浮窗中开启强制置顶，或打开任务栏行情模式开关，将行情显示固定在任务栏上。

**隐藏或开启鼠标穿透后如何操作？**

通过托盘菜单或已启用的快捷键恢复显示、关闭穿透。关闭托盘图标显示和快捷键时，不建议使用鼠标穿透

**为什么部分设置没有显示或无法使用？**

任务栏、强制置顶和边界检测设置仅在 Windows 显示；macOS 不提供鼠标穿透。Linux X11 支持全局快捷键和鼠标穿透，Wayland 下这些功能及窗口不透明度不可用。以当前界面的选项和提示为准。

**行情无法刷新怎么办？**

查看浮窗错误提示，检查网络或已有代理设置，也可切换数据源。不同数据源支持的市场和可用情况可能不同。反馈问题时请附上系统、程序版本、数据源和错误提示。

## 界面一览

以下为已有版本截图，具体设置以当前版本界面为准。

**极简浮窗**

<img width="127" height="90" alt="极简行情浮窗" src="https://github.com/user-attachments/assets/69ce08a7-6b55-41e8-bbab-e47b64d6e8a0" />

**任务栏行情模式（仅Windows）**

<img width="361" height="54" alt="任务栏行情模式" src="https://github.com/user-attachments/assets/24302d2a-897f-4dca-b008-80be07aea84d" />

**分时/日K显示**

<img width="330" height="188" alt="分时浮窗" src="https://github.com/user-attachments/assets/a5488528-6f3a-4523-a649-f64db8089e53" />
<img width="290" height="188" alt="日K浮窗" src="https://github.com/user-attachments/assets/69cb3089-ccd6-4f51-abf5-85f4b3d94d3a" />


**浅色与深色设置面板**

<img width="470" height="350" alt="浅色设置面板" src="https://github.com/user-attachments/assets/1dd6f7c7-0b51-4e9d-981d-0ab6d71fbbb6" />
<img width="470" height="350" alt="深色设置面板" src="https://github.com/user-attachments/assets/0a8caf56-1a16-45e2-a2de-c600de0ed099" />

**添加与编辑自选**

<img width="470" height="350" alt="添加自选" src="https://github.com/user-attachments/assets/7eb04bea-3f4a-494f-8735-533fd297dbd6" />
<img width="470" height="350" alt="快捷添加与编辑" src="https://github.com/user-attachments/assets/0d93bba7-71fc-4195-b250-4f09aeffa57c" />

## 数据说明

行情来自 **新浪财经**或**东方财富**，可能存在延迟、缺失或暂时不可用，仅供参考。程序只展示行情，不接入证券账户，不执行交易。使用第三方数据源时请遵守其使用条款。

## 许可与署名

本项目基于 **Apache License 2.0** 发布，完整文本见 [LICENSE](LICENSE)，署名信息见 [NOTICE](NOTICE)。再分发时必须保留许可、版权与署名信息，注明原作者 **sbr0574** 和[原始仓库](https://github.com/sbr0574/StockWidget)。
