# -*- mode: python ; coding: utf-8 -*-
#
# StockWidget 打包配置
# - Windows / Linux：EXE + COLLECT，产出 one-dir 目录（dist/StockWidget/）。
# - macOS：额外执行 BUNDLE 阶段，产出可直接拖入「应用程序」的
#   StockWidget.app 应用包（dist/StockWidget.app）。

import sys
from pathlib import Path
from runpy import run_path

metadata = run_path(str(Path(SPECPATH) / 'stockwidget/constants.py'))
APP_NAME = metadata['APP_NAME']
APP_VERSION = metadata['APP_VERSION']
COPYRIGHT = 'Copyright © 2026 sbr0574'

# Windows 版本资源
windows_version = None
if sys.platform == 'win32':
    import re
    from PyInstaller.utils.win32.versioninfo import (
        FixedFileInfo, StringFileInfo, StringStruct, StringTable,
        VarFileInfo, VarStruct, VSVersionInfo,
    )

    # Windows 数值版本
    if not re.fullmatch(r'[0-9]+(?:\.[0-9]+){0,3}', APP_VERSION):
        raise ValueError('APP_VERSION must contain 1 to 4 dot-separated integers')
    version_parts = tuple(int(part) for part in APP_VERSION.split('.'))
    if any(part > 65535 for part in version_parts):
        raise ValueError('Each APP_VERSION component must be between 0 and 65535')
    numeric_version = version_parts + (0,) * (4 - len(version_parts))
    windows_version = VSVersionInfo(
        ffi=FixedFileInfo(
            filevers=numeric_version,
            prodvers=numeric_version,
            mask=0x3f,
            flags=0x0,
            OS=0x40004,
            fileType=0x1,
            subtype=0x0,
            date=(0, 0),
        ),
        kids=[
            StringFileInfo([StringTable('040904B0', [
                StringStruct('CompanyName', 'sbr0574'),
                StringStruct('FileDescription', f'{APP_NAME} - Desktop Stock Monitor'),
                StringStruct('FileVersion', APP_VERSION),
                StringStruct('InternalName', APP_NAME),
                StringStruct('LegalCopyright', COPYRIGHT),
                StringStruct('OriginalFilename', f'{APP_NAME}.exe'),
                StringStruct('ProductName', APP_NAME),
                StringStruct('ProductVersion', APP_VERSION),
                StringStruct('Website', 'https://github.com/sbr0574/StockWidget'),
            ])]),
            VarFileInfo([VarStruct('Translation', [1033, 1200])]),
        ],
    )

datas = []
if sys.platform.startswith('linux'):
    datas.append(('resources/icons/StockWidget.png', 'icons'))


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
    name=APP_NAME,
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
    icon=(['resources/icons/StockWidget.ico'] if sys.platform == 'win32' else []),
    version=windows_version,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name=APP_NAME,
)

if sys.platform == 'darwin':
    app = BUNDLE(
        coll,
        name=f'{APP_NAME}.app',
        icon='resources/icons/StockWidget.icns',
        bundle_identifier='com.sbr0574.StockWidget',
        version=APP_VERSION,
        info_plist={
            'NSHighResolutionCapable': True,
            'LSUIElement': True,
            'NSHumanReadableCopyright': COPYRIGHT,
            'LSApplicationCategoryType': 'public.app-category.finance',
        },
    )
