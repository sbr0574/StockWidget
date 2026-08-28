# -*- coding: utf-8 -*-
"""代码数据工作流与 PR 测试工作流约束测试。"""

import unittest
from pathlib import Path


ROOT = Path(__file__).parents[1]
UPDATE_WORKFLOW = ROOT / ".github" / "workflows" / "update-codes.yml"
TEST_WORKFLOW = ROOT / ".github" / "workflows" / "test.yml"
UPDATER = ROOT / "scripts" / "updater.sh"


class UpdateWorkflowTests(unittest.TestCase):
    def test_only_classified_json_files_are_committed(self):
        text = UPDATE_WORKFLOW.read_text(encoding="utf-8")
        self.assertIn("workflow_dispatch:", text)
        self.assertIn("CODES_OUTPUT_DIR=\"$DATA_DIR/resources\"", text)
        self.assertIn("DATA_BRANCH: codes-data", text)
        self.assertIn('git push origin "HEAD:$DATA_BRANCH"', text)
        for filename in (
            "stock_sh.json",
            "stock_sz.json",
            "stock_bj.json",
            "fund_cn.json",
            "stock_hk.json",
            "stock_us.json",
            "index_cn.json",
            "index_global.json",
            "futures_sh.json",
            "codes_update_status.json",
        ):
            self.assertIn(f"resources/{filename}", text)
        self.assertNotIn("resources/resources_rc.py", text)
        self.assertNotIn("git add -A", text)

    def test_server_updater_scopes_credentials_and_added_files(self):
        text = UPDATER.read_text(encoding="utf-8")
        self.assertIn("GITHUB_TOKEN_FILE", text)
        self.assertIn(".github_token", text)
        self.assertIn("AUTHORIZATION: basic", text)
        self.assertIn('git -C "$DATA_DIR" add --', text)
        self.assertNotIn("git add -A", text)
        self.assertNotIn("resources/resources_rc.py", text)


class PullRequestTestWorkflowTests(unittest.TestCase):
    def test_tests_are_tracked_and_run_for_main_pull_requests(self):
        ignored = {
            line.strip()
            for line in (ROOT / ".gitignore").read_text(encoding="utf-8").splitlines()
            if line.strip() and not line.lstrip().startswith("#")
        }
        workflow = TEST_WORKFLOW.read_text(encoding="utf-8")

        self.assertNotIn("tests/", ignored)
        self.assertIn("pull_request:", workflow)
        self.assertIn("- main", workflow)
        self.assertIn("python -m unittest discover -s tests -v", workflow)


if __name__ == "__main__":
    unittest.main()
