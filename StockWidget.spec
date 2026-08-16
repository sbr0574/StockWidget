# -*- mode: python ; coding: utf-8 -*-
#
# StockWidget 打包配置（Windows / macOS / Linux 通用）。
#
# - Windows / Linux：EXE + COLLECT，产出 one-dir 目录（dist/StockWidget/）。
# - macOS：额外执行 BUNDLE 阶段，产出可直接拖入「应用程序」的
#   StockWidget.app 应用包（dist/StockWidget/StockWidget.app）。
#
# 版本号请与 stockwidget/constants.py 中的 APP_VERSION、version_info.txt 保持一致。
import sys

APP_VERSION = "1.3.1"

datas = [
    ('README.md', '.'),
    ('LICENSE', '.'),
    ('NOTICE', '.'),
]


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='StockWidget',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=(['resources/StockWidget.ico'] if sys.platform == 'win32' else []),
    version=('version_info.txt' if sys.platform == 'win32' else None),
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='StockWidget',
)

# macOS：把 one-dir 产物封装为 .app 应用包（在其它平台为 no-op）。
if sys.platform == 'darwin':
    app = BUNDLE(
        coll,
        name='StockWidget.app',
        icon='resources/StockWidget.icns',
        bundle_identifier='com.sbr0574.StockWidget',
        version=APP_VERSION,
        info_plist={
            'NSHighResolutionCapable': True,
            'NSHumanReadableCopyright': 'Copyright © 2026 sbr0574',
            'LSApplicationCategoryType': 'public.app-category.finance',
        },
    )
