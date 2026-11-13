"""Check native package architecture, macOS deployment baseline and Qt loading."""

import argparse
import os
from pathlib import Path
import platform
import re
import struct
import subprocess
import sys


ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def verify_binary(path, system, arch):
    if system == "darwin":
        actual = subprocess.check_output(["lipo", "-archs", str(path)], text=True).split()
        expected = "x86_64" if arch == "amd64" else "arm64"
        if actual != [expected]:
            raise ValueError(f"{path.name}: expected {expected}, found {actual}")
        return
    with path.open("rb") as binary:
        header = binary.read(64)
        if system == "win32":
            if len(header) < 64 or header[:2] != b"MZ":
                raise ValueError(f"{path.name}: invalid PE header")
            binary.seek(struct.unpack_from("<I", header, 0x3C)[0])
            header = binary.read(24)
            if len(header) < 24 or header[:4] != b"PE\0\0":
                raise ValueError(f"{path.name}: invalid PE signature")
            machine = struct.unpack_from("<H", header, 4)[0]
            expected = {"amd64": 0x8664, "arm64": 0xAA64}[arch]
        elif system == "linux":
            if len(header) < 20 or header[:6] != b"\x7fELF\x02\x01":
                raise ValueError(f"{path.name}: expected a 64-bit little-endian ELF executable")
            machine = struct.unpack_from("<H", header, 18)[0]
            expected = {"amd64": 62, "arm64": 183}[arch]
        else:
            raise ValueError(f"Unsupported build platform: {system}")
    if machine != expected:
        raise ValueError(f"{path.name}: expected {arch} (0x{expected:x}), found machine 0x{machine:x}")


def macos_minimum(text):
    versions = [(0, 0, 0)]
    for command in re.split(r"Load command \d+", text):
        if re.search(r"\bcmd LC_BUILD_VERSION\b", command):
            match = re.search(r"^\s*minos ([0-9.]+)\s*$", command, re.MULTILINE)
        elif re.search(r"\bcmd LC_VERSION_MIN_MACOSX\b", command):
            match = re.search(r"^\s*version ([0-9.]+)\s*$", command, re.MULTILINE)
        else:
            continue
        if match:
            parts = tuple(int(part) for part in match[1].split("."))
            versions.append(parts + (0,) * (3 - len(parts)))
    return max(versions)


def verify_macos_minimum(bundle, version):
    target = tuple(int(part) for part in version.split("."))
    target += (0,) * (3 - len(target))
    magic = {b"\xcf\xfa\xed\xfe", b"\xfe\xed\xfa\xcf", b"\xca\xfe\xba\xbe", b"\xca\xfe\xba\xbf"}
    checked = set()
    for path in bundle.rglob("*"):
        if not path.is_file() or path.resolve() in checked:
            continue
        checked.add(path.resolve())
        with path.open("rb") as binary:
            if binary.read(4) not in magic:
                continue
        commands = subprocess.check_output(["otool", "-l", str(path)], text=True)
        required = macos_minimum(commands)
        if required > target:
            minimum = ".".join(str(part) for part in required)
            raise ValueError(f"{path.relative_to(bundle)} requires macOS {minimum}, above release baseline {version}")


def smoke_qt():
    os.environ["QT_QPA_PLATFORM"] = "offscreen"
    from PySide6.QtCore import qVersion
    from PySide6.QtWidgets import QApplication, QWidget
    import stockwidget.app  # noqa: F401

    application = QApplication([])
    preview = QWidget()
    preview.resize(64, 64)
    if preview.grab().isNull():
        raise RuntimeError("Qt offscreen rendering failed")
    preview.close()
    application.processEvents()
    print(f"Qt {qVersion()} import and offscreen rendering passed")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--arch", choices=("amd64", "arm64"), required=True)
    parser.add_argument("--macos-minimum", default="13.0")
    args = parser.parse_args()
    aliases = {"amd64": {"amd64", "x86_64"}, "arm64": {"arm64", "aarch64"}}
    if platform.machine().lower() not in aliases[args.arch] or struct.calcsize("P") != 8:
        parser.error(f"Python architecture {platform.machine()} does not match {args.arch}")
    if sys.platform == "darwin":
        bundle = ROOT / "dist/StockWidget.app"
        binary = bundle / "Contents/MacOS/StockWidget"
        verify_binary(binary, sys.platform, args.arch)
        verify_macos_minimum(bundle, args.macos_minimum)
    else:
        name = "StockWidget.exe" if sys.platform == "win32" else "StockWidget"
        verify_binary(ROOT / "dist/StockWidget" / name, sys.platform, args.arch)
    smoke_qt()
    print(f"Verified {sys.platform} {args.arch} build")


if __name__ == "__main__":
    main()
