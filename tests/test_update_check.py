# -*- coding: utf-8 -*-
"""代码托管源选择和链接生成测试。"""

import unittest
from unittest.mock import Mock, patch

from stockwidget.data.update_check import GITEE, github_available, project_links


class TestRemoteSource(unittest.TestCase):
    def test_gitee_links_contain_no_github_urls(self):
        links = project_links(GITEE)
        self.assertEqual(links["repository_label"], "Gitee仓库")
        self.assertTrue(all("github.com" not in value for value in links.values()))

    @patch("stockwidget.data.update_check.requests.get")
    def test_github_probe_requires_success(self, get):
        get.return_value = Mock(status_code=200)
        self.assertTrue(github_available())
        get.return_value = Mock(status_code=503)
        self.assertFalse(github_available())


if __name__ == "__main__":
    unittest.main()
