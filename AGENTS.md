# StockWidget 开发约定

StockWidget 是 Windows、macOS 和 Linux 桌面行情工具，只展示行情，不接入证券账户或执行交易。版本号统一维护在 [stockwidget/constants.py](stockwidget/constants.py)。

## 分支、提交与记录

- `main`分支是主分支，主分支不允许直接修改与提交，只允许通过 PR 方式合并，合并方式默认为 Squash merge。
- `dev` 为开发分支，新任务从 `dev` 创建独立分支，完成开发后合并回 `dev`。
- `codes-data` 分支专用于发布代码列表，与其他分支互相保持独立。
- 若无说明，分支不允许自动创建、自动合并，若当前开发内容与当前分支用途不符，则询问是否转至其他分支或者创建新分支编辑。
- 分支名用 `<type>/<description>`，描述用小写英文和连字符，可含数字或 Issue 编号，如 `fix/taskbar-drag`。发布准备用 `release/<version>`，紧急修复用 `hotfix/<description>`。
- 每次修改完成后自动提交，一次提交只解决一个问题。标题用 `<type>(<scope>): <subject>`，范围可省略。优先用中文写明实际改动，尽量不超过 50 字，末尾不加标点，如 `fix(taskbar): 修复拖离任务栏后无法继续移动的问题`。
- 类型按用途选 `feat`、`fix`、`docs`、`refactor`、`perf`、`test`、`style`、`chore`、`ci` 或 `revert`。复杂改动在正文交代原因与影响，必要时用 `Closes #123` 或 `BREAKING CHANGE: ...` 作尾注。
- 每次修改 `/stockwidget` 下软件核心内容时，更新本地忽略的 `CHANGELOG.md`。按版本归入新增、优化、修复，描述相对上一版的最终结果。本版新增功能的后续调整并回原条目，保持简洁，不涉及细碎修改。开发验证放在折叠区，README 不链接此文件。
- [README.md](README.md) 面向使用者，功能或运行要求变化时再更新。本文件只保留跨任务仍适用的维护约定。

## 环境与常用命令

源码运行要求 Python 3.12 或更高版本，开发、CI 和打包优先用 3.13。从仓库根目录运行。Windows PowerShell 首次准备环境如下。

```powershell
py -3.13 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt -r scripts/requirements.txt
pyside6-rcc resources/resources.qrc -o resources/resources_rc.py
```

macOS / Linux 用 `python3.13 -m venv .venv` 和 `source .venv/bin/activate`，再执行相同的安装与资源生成命令。桌面程序使用根目录依赖，代码列表脚本使用 `scripts/requirements.txt`，全量测试需要两份。

- **运行程序**　在桌面会话中执行 `python main.py`。离屏测试使用单独的终端，以免桌面程序继承 `offscreen`。
- **生成界面**　编辑 `stockwidget/ui/generated/` 下的 Designer `.ui` 后运行 `python scripts/generate_ui.py`。用 `python scripts/generate_ui.py --check` 检查提交的适配层。该检查逐字比较生成结果，应使用与已提交文件头一致的 Qt UIC 版本。资源清单或其中的文件变化后，重跑上面的 `pyside6-rcc` 命令。两种生成文件均不手改。
- **运行测试**　先运行受影响的模块。以下 PowerShell 示例依次给出单模块、全量离屏和差异检查。macOS / Linux 的测试终端用 `export QT_QPA_PLATFORM=offscreen`。

```powershell
$env:QT_QPA_PLATFORM = "offscreen"
python -m unittest tests.ui.test_settings_layout -v
python -m unittest discover -s tests -t .
git diff --check
```

## 按改动找入口

程序的纯规则放在 `stockwidget/core/`，网络适配放在 `data/`，平台能力与系统调用放在 `platform/`，界面放在 `ui/`（这四个目录均在 `stockwidget/` 下）。复用已有模块，按状态归属拆分。手写源码与测试以每文件约 1000 行为维护参考，避免只做转发的中间层。

| 改动 | 入口与需要一起检查的内容 |
| --- | --- |
| 新增行情指标 | 在 [quote_presentation.py](stockwidget/core/quote_presentation.py) 定义指标与表头，`format_quote()` 提供显示值、颜色角色和原始排序值。[指标池](stockwidget/ui/controls/metrics.py) 读取同一目录，浮窗和任务栏共用列投影及绘制。 |
| 设置、外观与菜单 | 在 [settings.ui](stockwidget/ui/generated/settings.ui) 按现有分类和行样式添加控件，保留 `objectName` 和提升类。默认值与兼容读取归配置所有者，[groups.py](stockwidget/ui/settings/groups.py) 和 [dialog.py](stockwidget/ui/settings/dialog.py) 负责绑定。设置和菜单通过 `FloatLabel` 修改配置，同步期间保留各自独立偏好。 |
| 浮窗、任务栏与交互 | [widget.py](stockwidget/ui/floating/widget.py) 装配窗口。[presenter.py](stockwidget/ui/floating/presenter.py) 的 `QuotePresenter` 独占行情请求、缓存、排序、证券行身份和两个页码。[interaction.py](stockwidget/ui/floating/interaction.py) 管拖动及贴边，[taskbar.py](stockwidget/ui/floating/taskbar.py) 管嵌入和原生输入。 |
| 基本行情与图表 | 基本行情适配在 [quotes.py](stockwidget/data/quotes.py)，历史与缓存分别在 [bars.py](stockwidget/data/bars.py)、[bar_cache.py](stockwidget/data/bar_cache.py)。[history.py](stockwidget/ui/history.py) 管图表窗口和在途请求，点击按排序、分页、分栏后的证券键定位。 |
| 平台功能 | 在 [capabilities.py](stockwidget/platform/capabilities.py) 结合实际 Qt 插件判断能力。原生调用归对应平台模块。设置中隐藏平台专属项，对随会话而不可用的功能保留禁用原因。 |

## 容易引入回归的边界

- **配置与迁移**　[config_store.py](stockwidget/core/config_store.py) 统一管理数据路径与原子保存。保留旧配置、自选和位置记录的读取兼容。仅 Windows 在默认目录不可写时回退程序目录，用独立临时文件探测写入。
- 更改数据目录时保留源文件、备份目标已有数据，重启后切换。macOS 旧目录迁移优先保留新版数据，完整迁移或备份成功后才清理旧目录。写入失败合并提示，设置页与退出仍须可用。
- **行情与缓存**　基本行情、历史和美股中文名称使用新浪，沪深北个股的内外盘使用腾讯快照。保留统计来源与时间，不用总量补差，不新增数据源配置项。沿用共享错误分类、代理与证书校验，只重试临时错误，错误提示不暴露原始 URL、凭据或异常栈。
- 图表复用行情刷新，分钟和日 K 请求按标的、类别共享。关闭窗口或清理缓存后，迟到结果不得重开窗口或写回已清数据。历史缓存按事务写入，仅确认数据库损坏才重建，清缓存只影响历史数据。
- **代码列表**　服务端直接运行 [scripts/update_codes.py](scripts/update_codes.py)，以 `CODES_OUTPUT_DIR` 指定发布目录。采用交易所官方名录，港股名称发布前转简体。分类或补充表失败时保留旧文件、记入状态清单，其余分类继续保存。
- 在 `codes-data` 分支的 `resources/index_cn.json` 和 `resources/index_global.json` 手工维护指数及 `last_update`，更新脚本只读。新增国内指数须核对官方身份及新浪实时报价。海外指数的 `timezone` 随证券元数据传递，缺失时不猜测日 K 交易日期。
- **拖动与可见性**　浮窗新增可见区域接入 `register_drag_region()`。两种显示面复用拖动与显隐入口，越过拖动阈值后不触发单击，松开后才做边界检查，Esc 恢复拖动前状态。贴边收起时保存完整几何。
- 双开时任务栏固定，浮窗经过任务栏不停靠。任务栏启用状态与显示位置分开保存，嵌入失败回退浮窗并允许重建。根据两个显示面的实际可见性暂停或恢复行情刷新，各自保留页码。
- **窗口与设置**　更改窗口标志时保留几何、可见性、计时器与鼠标穿透，避免抢焦点。设置布局沿用 Designer 的单栏滚动、说明列宽和控件对齐，检查实际可达性与横向溢出。主题集中在 [style.py](stockwidget/ui/controls/style.py)，普通配置变化通过 `configuration_changed` 同步值，避免重复刷新整页。

## 验证与报告

- 优先覆盖行为与错误恢复，复用 [tests/support.py](tests/support.py) 的夹具。存储改动检查[迁移、回滚及重启](tests/core/test_storage_migration.py)。窗口改动检查[拖动、取消及双开](tests/ui/test_position.py)，设置页检查[布局与控件可达性](tests/ui/test_settings_layout.py)。绘制改动检查实际输出、可读性及缩放效果。
- 离屏测试验证 Qt 行为。改动原生操作时，在对应平台另跑下表用例。Linux 系统依赖与 Xvfb 步骤见 [test.yml](.github/workflows/test.yml)。打包改动按平台工作流检查产物架构，macOS 还要核验包内原生库的最低系统版本。
- 报告实际运行的范围、失败和跳过项。疑似既有失败用同环境的基线提交复核。分别说明离屏、原生桌面、高 DPI 与真实网络的验证范围。

| 平台 | 原生验证示例（已激活虚拟环境） |
| --- | --- |
| Windows 桌面（PowerShell） | 分别执行 `$env:QT_QPA_PLATFORM = "windows"`、`$env:STOCKWIDGET_TEST_WINDOWS = "1"`，然后执行 `python -m unittest tests.native.test_taskbar -v`。 |
| macOS 桌面 | `QT_QPA_PLATFORM=cocoa STOCKWIDGET_TEST_MACOS=1 python -m unittest tests.native.test_macos_integration -v` |
| Linux X11 | `QT_QPA_PLATFORM=xcb STOCKWIDGET_TEST_X11=1 xvfb-run -a python -m unittest tests.native.test_x11_integration -v` |
