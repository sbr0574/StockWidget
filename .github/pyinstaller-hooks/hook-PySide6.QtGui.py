from pathlib import Path

from PyInstaller.utils.hooks.qt import add_qt6_dependencies


hiddenimports, binaries, datas = add_qt6_dependencies(__file__)
# Exclude unused plugins before their PDF and Qt Quick dependencies are collected.
binaries = [
    item for item in binaries
    if Path(item[0]).stem.lower() not in {
        "qpdf", "libqpdf", "qtvirtualkeyboardplugin", "libqtvirtualkeyboardplugin",
    }
]
