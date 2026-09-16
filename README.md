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


## 🌐 数据来源 & 网络

* 行情信息通过 `requests` 从 **新浪财经**或**东方财富**接口获取，数据不可避免存在一定延迟，仅供参考。
* 程序仅显示行情数据，不包含任何账户/交易操作。
* 请根据自身网络环境决定是否使用代理或更换数据源。
* 浮窗隐藏时会暂停刷新，显示后自动恢复，减少不必要的请求。


## 📜 许可

本项目基于 **Apache License 2.0** 发布，完整文本见 [LICENSE](LICENSE)，署名要求见 [NOTICE](NOTICE)。

* 个人/学习用途自由使用；涉及第三方数据源时请遵守其使用条款。
* **再分发（转载/镜像/打包发布）必须保留 `LICENSE` 与 `NOTICE` 文件，注明原始作者与仓库地址**（https://github.com/sbr0574/StockWidget）。
