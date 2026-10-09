"""Regenerate the Python adapters for every editable Qt Designer form."""

import argparse
from importlib.util import find_spec
from pathlib import Path
import subprocess
import sys


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("forms", nargs="*", help="Optional form names, e.g. settings or metric_pool")
    parser.add_argument("--check", action="store_true", help="Check generated files without writing")
    args = parser.parse_args()
    directory = Path(__file__).resolve().parents[1] / "stockwidget/ui/generated"
    if find_spec("PySide6") is None:
        parser.error("找不到 PySide6，请在安装了项目依赖的 Python 环境中运行")
    # Invoke the installed console entry point with this interpreter. A moved
    # virtual environment can leave Windows .exe launchers pointing elsewhere.
    compiler = [sys.executable, "-c", "from PySide6.scripts.pyside_tool import uic; uic()"]
    forms = sorted(directory.glob("*.ui"))
    if args.forms:
        names = {Path(name).stem for name in args.forms}
        available = {path.stem for path in forms}
        if names - available:
            parser.error("未知界面：" + ", ".join(sorted(names - available)))
        forms = [path for path in forms if path.stem in names]
    stale = []
    for source in forms:
        result = subprocess.run([*compiler, str(source)], check=True, capture_output=True)
        generated = result.stdout.decode("utf-8").replace("\r\n", "\n")
        target = source.with_name("ui_" + source.stem + ".py")
        if args.check:
            if not target.is_file() or target.read_text(encoding="utf-8") != generated:
                stale.append(target.name)
        else:
            target.write_text(generated, encoding="utf-8")
            print(source.name + " -> " + target.name)
    if stale:
        print("需要重新生成：" + ", ".join(stale), file=sys.stderr)
        return 1
    if args.check:
        print(f"{len(forms)} 个界面的生成文件已同步")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
