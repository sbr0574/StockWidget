# -*- coding: utf-8 -*-
"""代码托管源选择和链接生成测试。"""

import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

from stockwidget.constants import CODES_BRANCHES
from stockwidget.data.update_check import GITEE, github_available, project_links
from stockwidget.ui.settings_dialog import SettingsDialog


class TestRemoteSource(unittest.TestCase):
    def test_gitee_links_contain_no_github_urls(self):
        links = project_links(GITEE)
        self.assertEqual(links["repository_label"], "Gitee仓库")
        self.assertTrue(all("github.com" not in value for value in links.values()))

    @patch("stockwidget.data.update_check.requests.get")
    def test_github_probe_requires_success(self, get):
        get.return_value = Mock(status_code=200)
        self.assertTrue(github_available())
        self.assertIn(CODES_BRANCHES[0], get.call_args.args[0])
        self.assertEqual(get.call_args.kwargs["timeout"], (2.5, 2.5))
        self.assertTrue(get.call_args.kwargs["stream"])
        get.return_value = Mock(status_code=503)
        self.assertFalse(github_available())

    def test_about_uses_selected_links_and_always_shows_both_repositories(self):
        label = Mock()
        dialog = SimpleNamespace(
            app=SimpleNamespace(
                app_version="1.4.0",
                _has_update=False,
                _latest_version=None,
                _latest_release_url=None,
                _remote_source=GITEE,
            ),
            ui=SimpleNamespace(label_about_info=label),
        )

        SettingsDialog._setup_about(dialog)

        html = label.setText.call_args.args[0]
        self.assertIn("https://gitee.com/sbr0574/StockWidget/releases", html)
        self.assertIn("https://gitee.com/sbr0574/StockWidget/blob/main/LICENSE", html)
        self.assertIn(">GitHub仓库</a>", html)
        self.assertIn(">Gitee仓库</a>", html)


if __name__ == "__main__":
    unittest.main()
