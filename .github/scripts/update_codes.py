"""拉取全市场代码并写入 resources/ 下三个 JSON（供 GitHub Action 调用）。

用法: python .github/scripts/update_codes.py
"""
import os
import sys

# 脚本位于 <项目根>/.github/scripts/，向上三层即项目根
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)

from stockwidget.data import code_index


def main() -> int:
    resources_dir = os.path.join(ROOT, "resources")
    groups = code_index.write_codes_groups(resources_dir)
    if groups is None:
        print("::error::拉取代码列表失败")
        return 1
    for fname, data in groups.items():
        print(f"更新 {fname}: {len(data.get('codes', {}))} 条")
    print("完成")
    return 0


if __name__ == "__main__":
    sys.exit(main())
