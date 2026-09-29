@REM conda create -n stockWidget python=3.13
conda activate stockWidget
python -V

@REM 安装依赖并生成 Qt 资源文件
@REM pip install -r requirements.txt
@REM pyside6-rcc resources/resources.qrc -o resources/resources_rc.py

@REM 脚本式运行
@REM python main.py

@REM 打包exe（产出 dist\StockWidget\StockWidget.exe）
@REM 1. 生成 Qt 资源文件（app.py 会 import resources.resources_rc）
@REM pyside6-rcc resources/resources.qrc -o resources/resources_rc.py
@REM 2. 用 spec 文件打包
@REM pyinstaller --clean --noconfirm StockWidget.spec
