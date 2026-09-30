# 手动调整界面

所有可用 Qt Designer 调整的界面放在 `stockwidget/ui/generated/`。现在程序直接使用这些文件生成的控件和布局；Python 负责连接信号、同步设置与显示数据，不再在设置页另建一套布局。

| 文件 | 可调整内容 |
| --- | --- |
| `settings.ui` | 五个设置页；浮窗行数限制、翻页、分栏、自动隐藏与定时隐藏；任务栏全部组框、字体和不透明度滑块、颜色按钮、指标区域、分栏同步；快捷键状态位置 |
| `metric_pool.ui` | 浮窗和任务栏共用的指标双池、标题、边距和间距 |
| `add_code_panel.ui` | 添加标的面板：搜索框、类别和地区筛选、结果列表、分页按钮 |
| `name_settings.ui` | 名称指标弹出的字数、代码和类型设置 |
| `unit_settings.ui` | 成交量、成交额弹出的单位设置 |
| `floating_widget.ui` | 浮窗背景容器、提示与隐藏倒计时区域、分页区域、左右行情表格和中间分隔线的位置与间距 |

## 修改与生成

1. 用 Qt Designer 打开需要调整的 `.ui`，修改布局、间距、边距、控件尺寸和静态文字，然后保存。
2. 在项目根目录运行：

   ```powershell
   .venv/Scripts/python.exe scripts/generate_ui.py
   ```

   只生成一个界面也可以：

   ```powershell
   .venv/Scripts/python.exe scripts/generate_ui.py settings
   ```

3. 重新运行源码程序查看效果；已经打包的 EXE 需要重新打包才会更新。

检查生成文件是否最新：

```powershell
.venv/Scripts/python.exe scripts/generate_ui.py --check
```

如果尚未安装 Designer，可安装项目依赖后运行 `.venv/Scripts/pyside6-designer.exe`。

## 哪些属性会由运行时控制

- 保留控件的 `objectName` 和提升类名称，业务绑定靠这些名称找到控件；移动、调整大小和布局不需要改名。
- 设置项的当前勾选状态、数值、字体和颜色来自用户配置；修改 Designer 的初始值不会覆盖已保存的配置。默认业务参数在 `core/view_options.py` 和 `ui/widget.py` 中。
- 下拉框的选项位置对应业务枚举；可以改显示文字，增删或重新排序选项需同步修改绑定代码。
- `MetricPoolWidget` 等提升控件在 Designer 中通常显示为基础控件占位区域。要调整指标池内部，打开 `metric_pool.ui`；其外部位置和尺寸在 `settings.ui` 中调整。不需要撤销控件提升。
- 共用 flat 组框及颜色按钮主题在 `ui/settings_style.py` 中。`settings.ui` 根窗口的自定义 `styleSheet` 会追加在主题之后，可作为个人覆盖样式；程序会随系统主题更新通用样式。
- 设置页的全部组框直接使用 Qt 原生 `QGroupBox`，不再提升为自定义组框。`ui/view_settings.py` 中的 `QObject` 绑定对象只负责配置与编辑状态同步；同步组框勾选时禁用独立编辑，取消勾选时恢复编辑。
- 指标池按内容和可用空间动态分配列表高度，浮窗按行情行列和字体动态调整尺寸。`.ui` 的边距和间距会生效，但 Designer 中的预览窗口大小不会锁死浮窗尺寸。
- 添加标的面板默认每页 10 条，条目高度 24；若大幅修改结果列表高度，需要同时考虑 `add_code_panel.py` 的 `PAGE_SIZE` 和 `RESULT_ROW_HEIGHT`。

原生任务栏窗口、行情/K 线绘制、翻页箭头绘制、拖动捕获和动态右键菜单继续由 Python/Windows API 实现，不能直接用 Qt Designer 排版。浮窗中的分页占位区域可以在 `floating_widget.ui` 调整尺寸与位置。

修改浮窗布局后，按 `AGENTS.md` 检查数据、表头、分页箭头、页码、留白和各类提示区域都能拖动，并检查双击隐藏及恢复。
