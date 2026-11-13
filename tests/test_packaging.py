"""Release architecture, deployment baseline and Qt bundle collection checks."""

from pathlib import Path
import runpy
import struct
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch


ROOT = Path(__file__).resolve().parents[1]
BUILD_CHECK = runpy.run_path(str(ROOT / ".github/scripts/verify-build.py"))


class BuildVerificationTests(unittest.TestCase):
    def setUp(self):
        self.directory = Path(self.enterContext(tempfile.TemporaryDirectory()))

    def test_qt_plugin_filter_preserves_desktop_input_and_supported_images(self):
        hook = ROOT / ".github/pyinstaller-hooks/hook-PySide6.QtGui.py"
        for prefix, suffix, platform_plugins, input_plugins in (
                ("", ".dll", ("qwindows", "qoffscreen"), ()),
                ("lib", ".so", ("qxcb", "qwayland-generic", "qoffscreen"),
                 ("ibusplatforminputcontextplugin", "fcitx5platforminputcontextplugin")),
                ("lib", ".dylib", ("qcocoa", "qoffscreen"), ("composeplatforminputcontextplugin",))):
            with self.subTest(suffix=suffix):
                plugins = (
                    *(f"platforms/{name}" for name in platform_plugins),
                    *(f"platforminputcontexts/{name}" for name in input_plugins),
                    "imageformats/qjpeg", "imageformats/qico", "imageformats/qsvg",
                    "imageformats/qpdf", "platforminputcontexts/qtvirtualkeyboardplugin",
                )
                binaries = [
                    (str(self.directory / Path(plugin).parent / f"{prefix}{Path(plugin).name}{suffix}"),
                     str(Path("PySide6/plugins") / Path(plugin).parent))
                    for plugin in plugins
                ]
                dependencies = (["PySide6.QtCore"], binaries, [("qt.conf", "PySide6")])
                with patch("PyInstaller.utils.hooks.qt.add_qt6_dependencies", return_value=dependencies):
                    result = runpy.run_path(str(hook))
                self.assertEqual(result["binaries"], binaries[:-2])
                self.assertEqual(result["hiddenimports"], dependencies[0])
                self.assertEqual(result["datas"], dependencies[2])

    def test_spec_retains_chinese_qt_translations_and_other_package_data(self):
        names = (
            "PySide6/translations/qtbase_zh_CN.qm", "PySide6/translations/qtbase_de.qm",
            "PySide6/Qt/translations/qtbase_zh_CN.qm", "PySide6/Qt/translations/qtbase_ja.qm",
            "PySide6\\translations\\qtbase_zh_CN.qm", "PySide6\\translations\\qtbase_fr.qm",
            "other/translations/fr.qm", "tzdata/zoneinfo/America/New_York",
        )
        entries = [(name, str(self.directory / "data"), "DATA") for name in names]
        analysis = SimpleNamespace(datas=entries, binaries=[], scripts=[], pure=[])
        result = runpy.run_path(str(ROOT / "StockWidget.spec"), init_globals={
            "SPECPATH": str(ROOT), "Analysis": Mock(return_value=analysis),
            "PYZ": Mock(), "EXE": Mock(), "COLLECT": Mock(), "BUNDLE": Mock(),
        })
        self.assertEqual(result["a"].datas, [entries[index] for index in (0, 2, 4, 6, 7)])

    def test_native_headers_accept_matching_architecture_and_reject_wrong_labels(self):
        for system, machines in (("win32", {"amd64": 0x8664, "arm64": 0xAA64}),
                                 ("linux", {"amd64": 62, "arm64": 183})):
            for architecture, machine in machines.items():
                with self.subTest(system=system, architecture=architecture):
                    header = bytearray(128)
                    if system == "win32":
                        header[:2] = b"MZ"
                        struct.pack_into("<I", header, 0x3C, 64)
                        header[64:68] = b"PE\0\0"
                        struct.pack_into("<H", header, 68, machine)
                    else:
                        header[:6] = b"\x7fELF\x02\x01"
                        struct.pack_into("<H", header, 18, machine)
                    binary = self.directory / "StockWidget"
                    binary.write_bytes(header)
                    BUILD_CHECK["verify_binary"](binary, system, architecture)
                    other = "arm64" if architecture == "amd64" else "amd64"
                    with self.assertRaisesRegex(ValueError, "expected"):
                        BUILD_CHECK["verify_binary"](binary, system, other)

    def test_invalid_or_truncated_native_executables_are_rejected(self):
        binary = self.directory / "StockWidget"
        for contents in (b"", b"not executable", b"MZ" + bytes(62), b"\x7fELF\x01\x01" + bytes(58)):
            for system in ("win32", "linux"):
                with self.subTest(contents=contents[:6], system=system), self.assertRaises(ValueError):
                    binary.write_bytes(contents)
                    BUILD_CHECK["verify_binary"](binary, system, "amd64")

    def test_macos_architecture_check_rejects_mislabeled_or_universal_outputs(self):
        binary = self.directory / "StockWidget"
        for output, architecture, accepted in (("arm64", "arm64", True), ("x86_64", "amd64", True),
                                                ("arm64", "amd64", False), ("x86_64 arm64", "amd64", False)):
            with self.subTest(output=output, architecture=architecture), \
                    patch("subprocess.check_output", return_value=output):
                if accepted:
                    BUILD_CHECK["verify_binary"](binary, "darwin", architecture)
                else:
                    with self.assertRaises(ValueError):
                        BUILD_CHECK["verify_binary"](binary, "darwin", architecture)

    def test_macos_baseline_ignores_sdk_and_source_versions(self):
        commands = """Load command 1
              cmd LC_BUILD_VERSION
            minos 13.0
              sdk 26.0
Load command 2
              cmd LC_SOURCE_VERSION
          version 999.0
Load command 3
              cmd LC_VERSION_MIN_MACOSX
          version 12.3
              sdk 15.0
"""
        self.assertEqual(BUILD_CHECK["macos_minimum"](commands), (13, 0, 0))

    def test_bundled_library_can_raise_the_required_macos_version(self):
        library = self.directory / "Frameworks/QtGui"
        library.parent.mkdir()
        library.write_bytes(b"\xcf\xfa\xed\xfe")
        (self.directory / "README.md").write_text("ordinary data", encoding="utf-8")
        for required, accepted in (("13.0", True), ("14.1", False)):
            with self.subTest(required=required), patch("subprocess.check_output",
                    return_value=f"Load command 1\ncmd LC_BUILD_VERSION\nminos {required}\nsdk 26.0\n") as inspect:
                if accepted:
                    BUILD_CHECK["verify_macos_minimum"](self.directory, "13.0")
                else:
                    with self.assertRaisesRegex(ValueError, "QtGui requires macOS 14.1.0"):
                        BUILD_CHECK["verify_macos_minimum"](self.directory, "13.0")
                inspect.assert_called_once()
