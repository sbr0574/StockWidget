# StockWidget 开发约定

StockWidget 是 Windows、macOS 和 Linux 桌面行情工具，只展示行情，不接入证券账户或执行交易。版本号统一维护于 `stockwidget/constants.py`。

## 架构与状态所有权

| 层 / 模块 | 职责 |
| --- | --- |
| `main.py`、`stockwidget/app.py` | 程序入口、应用装配、托盘、设置、后台任务与配置保存 |
| `core/config_store.py`、`watchlist.py` | 配置路径与原子保存，自选元数据、校验及搜索 |
| `core/quote_presentation.py` | 指标目录、显示值 / 颜色角色 / 原始排序值，统一列投影 |
| `core/view_options.py`、`window_rules.py` | 显示偏好、同步、分页与分栏，屏幕位置与隐藏的纯规则 |
| `data/quotes.py`、`network_errors.py` | 行情请求、统一数据结构、共享请求策略与安全错误结果 |
| `data/bars.py`、`bar_cache.py` | 历史适配、交易日与实时补充规则，同源 SQLite 缓存 |
| `data/code_lists.py`、`update_check.py` | 市场代码列表同步与应用更新检查 |
| `platform/capabilities.py` | 平台 / 实际 Qt 会话能力及默认值 |
| `platform/window.py`、`taskbar.py`、`hotkeys.py`、`autostart.py` | 原生窗口、任务栏嵌入与输入、全局快捷键及自启 |
| `ui/floating/widget.py` | `FloatLabel` 装配、配置、外观与共用窗口操作入口 |
| `ui/floating/presenter.py` | `QuotePresenter` 独占请求代数、行情缓存、排序、行身份及两处页码 |
| `ui/floating/interaction.py` | 共用拖动、完整位置 / 贴边状态与自动隐藏控制器 |
| `ui/floating/taskbar.py` | 消费共享行情模型的任务栏绘制、原生输入适配与显示模式协调 |
| `ui/controls/` | 共用行情模型、代理、网格、指标池、设置控件、主题及历史图表绘制 |
| `ui/settings/dialog.py`、`groups.py` | 设置页装配与信号绑定，参数同步与编辑状态 |
| `ui/watchlist/`、`ui/history.py`、`ui/menus.py` | 自选编辑、图表请求 / 窗口生命周期与各显示面的菜单 |
| `ui/generated/` | Designer `.ui` 源文件及自动生成适配层 |
| `tests/core`、`data`、`ui`、`native` | 纯规则、数据、Qt 行为及真实平台集成；场景共用于 `tests/support.py` |

新增纯规则放入 `core`，网络适配放入 `data`，系统调用放入 `platform`，界面和绑定放入 `ui`。同类职责优先归入已有模块；按状态所有权组合较大的类，避免碎片文件、无用转发和全项目重新组织。手写源码与测试尽量不超过每文件1000行，生成文件除外。

- 浮窗和任务栏共用 `QuotePresenter`、统一指标定义及行情模型，不各自维护请求、缓存或排序逻辑。
- `format_quote()` 返回显示值、颜色角色和原始排序值；一个指标对应一个表头。两个显示面共用单元格尺寸、买一 / 卖一、K线和网格描边函数。
- 设置和菜单通过 `FloatLabel` 的共用入口修改配置；按当前显示面及同步状态选择有效设置，不能覆盖同步期间保留的独立偏好。
- 设置绑定订阅 `configuration_changed`，与行情绘制变化分开；`TaskbarSettings` 统一同步子组，避免父子重复订阅、重复刷新。
- 移动内部模块时更新全部调用方、测试补丁目标和 Designer 提升路径；删除无调用的旧接口，不保留空转发层。

## 配置、数据与请求

- 用户配置与自选读取 / 保存兼容必须保留；删除旧配置迁移需有明确迁移方案，不能把仍在使用的兼容字段当作死代码。
- 配置路径由 `config_store.py` 管理：Windows 为 `%APPDATA%/StockWidget`，其他平台为 `~/StockWidget`；市场代码在 `data/`，历史在 `cache/history.sqlite3`。不擅自更换目录或覆盖用户文件。
- 服务器代码列表发布到 `codes-data` 分支，业务代码在 `main`；更新脚本支持 `CODES_OUTPUT_DIR`。
- 请求错误使用共用分类及安全文本，不展示原始 URL、代理地址、凭据或长异常栈。沿用代理与证书校验，仅重试已分类的临时错误，不伪造空行情成功。
- 明确选择东财时保持东财来源，包括同源备用节点及缓存；不得悄悄切换新浪。缓存记录实际来源，不跨源拼接。
- 历史图表复用现有行情更新，没有独立刷新定时器；分钟 / 日K请求按标的、源和类别共享。关闭或清理缓存时使在途请求失效，迟到结果不得重开窗口或写回已清除缓存。
- 缓存按事务保存，只有确认损坏的数据库才允许重建；锁定、只读和未知版本不能触发删除。清理历史缓存不影响自选、配置和市场代码表。

## 显示与交互

- 拖动、隐藏、呼出、刷新暂停均走共用入口，Qt 浮窗和原生任务栏复用 `begin_drag / move_drag / finish_drag` 及 `hide_widget / toggle_win`。
- 所有可见浮窗区域接入 `register_drag_region()`，包括表头、分页、分隔线、留白、提示和新增组件。未越过系统拖动阈值才执行单击，拖回起点也不能误触单击；双击隐藏，Esc 取消并恢复原状态。
- 任务栏总开关与当前显示位置分开保存；启动尊重保存位置，嵌入失败回退浮窗，Explorer 重启后允许重建。双开时不因浮窗经过任务栏改变模式。
- 行情组件完全隐藏时暂停行情刷新和自动翻页；任务栏仍显示时继续刷新。分页各自记录页码，行情更新保留当前页，修改自选或排序后回到第一页。
- 贴边收起保存完整几何，不能持久化窄条尺寸。拖动、菜单和预览期间避免收起；内容变化立即适配尺寸并保持已贴边位置。
- Qt 窗口标志变化要保留几何、可见性、模式、计时器和鼠标穿透，不抢焦点；X11 输入透明标志与置顶标志一起维护，避免直接设置输入区域后被 Qt 重建覆盖。
- 图表点击通过排序、分页、分栏后的 canonical key 定位标的，不按名称或显示行号索引自选。任何隐藏路径都要拒收迟到图表请求。

## 平台能力

通过 `capabilities.py` 和实际 Qt 插件判断能力，不仅依赖系统名称或环境变量。

- 任务栏分类、强制置顶和边界检测只在 Windows 显示。对实际平台不存在的专属功能隐藏整行或分类；图标选择与自定义图标在三个平台均保留。
- Windows 贴边依赖边界检测；macOS / Linux X11 可独立开启贴边收起。Wayland 无可靠手动定位时禁用贴边并说明原因。
- 会话相关或有解释价值的不可用功能保留禁用提示，例如 Wayland 全局快捷键、鼠标穿透、窗口透明度和 macOS 鼠标穿透。
- Windows 原生置顶不激活窗口并避让设置 / 弹出菜单；原生任务栏窗口重建后重新应用穿透与外观状态。
- 平台模拟测试只能验证策略，不能替代真实桌面输入、窗口嵌入或 DPI 证据。

## 设置 UI 与新增条目

设置按“自选列表 / 数据 / 常规 / 快捷键 / 浮窗 / 任务栏 / 关于”分类：数据负责来源、指标与图表，常规负责应用偏好，浮窗和任务栏分别管理显示面。新增参数按用户用途归类，复用已有控件和行样式。

- 所有设置控件及静态布局维护在 `ui/generated/settings.ui`；指标池、添加面板和浮窗对应各自 `.ui`。绑定代码只做信号、配置同步、能力与编辑状态，不动态创建另一套布局。
- 修改 `.ui` 后运行 `python scripts/generate_ui.py`，用 `--check` 验证同步；禁止手改 `ui_*.py`。保留 `objectName`、提升类及枚举与绑定的对应关系。
- 设置窗口固定为704×496，分类页单栏纵向滚动，无横向滚动；组框使用原生 non-flat 外观。说明列宽统一310逻辑像素，左 / 顶对齐；数值、下拉框、字体与开关右对齐。
- 新增设置行沿用现有 Designer 行结构、`ToggleSwitch`、`settingDescription`、颜色按钮和主题样式；不要为单条设置新增独立配色或硬编码浅深背景。
- 组框首行距标题至少25px、距左侧至少10px，组件间隔至少5px；检查实际坐标和页面滚动，不只看布局边距。
- 同类组件保持尺寸一致；滑块和下拉框不响应滚轮改值，继续滚动设置页。独立配置在同步或暂时禁用期间保留。
- 主题统一由 `ui/controls/style.py` 管理；所选调色板应用于实际子页与控件，检查文字和背景可读性。只有主题 / 调色板变化才刷新整页样式，普通参数变化只同步值。

## 测试与验证

使用安装依赖的 Python 环境，本地通常为 `.venv/Scripts/python.exe`；CI 与构建使用 Python 3.13。

```powershell
python -m pip install -r requirements.txt
pyside6-rcc resources/resources.qrc -o resources/resources_rc.py
python scripts/generate_ui.py --check
python -m unittest discover -s tests -t .
git diff --check
```

- 优先验证用户行为、配置迁移、错误恢复、请求代数和平台边界；同一规则用参数化场景，初始化 / 清理使用共用夹具。不要重复测试 Qt 固有属性、控件别名、常量字符串或源码样式表内容。
- 浮窗改动检查左右数据、表头、分页、提示及留白的单击 / 拖动 / 松开 / Esc / 双击 / 恢复；分页与分栏覆盖末页、自动翻页、独占和双开。
- 绘制和主题检查实际输出、可读性及设备像素比；保留有意义的 UI 行为测试，不靠删除失败用例获得通过。
- 无头测试使用 `QT_QPA_PLATFORM=offscreen`。Windows 原生验证另设 `STOCKWIDGET_TEST_WINDOWS=1`、`QT_QPA_PLATFORM=windows`；X11 依据 CI 用 Xvfb、`QT_QPA_PLATFORM=xcb`、`STOCKWIDGET_TEST_X11=1`。
- 分别报告离线 / 全量 / 原生 / 高 DPI结果；跳过、基线失败或未验证平台不能描述为通过。真实网络可用性与解析测试分别报告。

## 文档与版本记录

- `README.md` 面向使用者，说明下载、首次运行、操作与可见限制；不写源码结构、原生 API、接口字段、缓存算法和开发验证细节。
- 本文件只保留重要且持久的架构、维护及行为约定，不积累单次功能要求、实现过程和每条参数的产品规格。
- 每次修改更新本地 `CHANGELOG.md`，用版本标题分隔，按新增 / 优化 / 修复记录实际最终成果；该文件保持 Git 忽略，不在 README 中链接。
- 保留 `LICENSE`、`NOTICE`、作者署名；不把未实现建议写成成果。功能工作使用独立分支，区分本地提交与已验证推送。
