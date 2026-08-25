# -*- coding: utf-8 -*-
"""每日市场代码工作流的数据分支约束测试。"""

import unittest
from pathlib import Path


WORKFLOW = (
    Path(__file__).parents[1] / ".github" / "workflows" / "update-codes.yml"
)


class TestUpdateWorkflow(unittest.TestCase):
    def test_only_json_files_are_committed_to_data_branch(self):
        text = WORKFLOW.read_text(encoding="utf-8")
        self.assertIn("DATA_BRANCH: codes-data", text)
        self.assertIn('git push origin "HEAD:$DATA_BRANCH"', text)
        self.assertIn(
            "git add resources/stock_codes_list.json resources/futures_codes_list.json",
            text,
        )
        self.assertNotIn("resources/resources_rc.py", text)


if __name__ == "__main__":
    unittest.main()
