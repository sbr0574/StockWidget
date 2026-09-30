import unittest
from unittest.mock import Mock, patch

import requests

from stockwidget.data import update_check


class UpdateCheckTests(unittest.TestCase):
    def test_failed_release_sources_return_specific_diagnostics(self):
        errors = []
        with patch.object(update_check.requests, "get", side_effect=(
            requests.ConnectTimeout(), requests.exceptions.ProxyError()
        )):
            self.assertEqual(update_check.get_update_info("1.0.0", errors=errors), (False, None))
        self.assertEqual(errors, ["GitHub：连接超时", "Gitee：代理请求失败"])

    def test_release_fallback_succeeds_after_a_request_error(self):
        response = Mock(status_code=200)
        response.json.return_value = {"tag_name": "v1.5.0"}
        errors = []
        with patch.object(update_check.requests, "get", side_effect=(requests.ReadTimeout(), response)):
            self.assertEqual(update_check.get_update_info("1.0.0", errors=errors), (True, "1.5.0"))
        self.assertEqual(errors, ["GitHub：读取超时"])

    def test_http_and_response_errors_are_reported_without_changing_default_return(self):
        response = requests.Response()
        response.status_code = 403
        invalid = Mock(status_code=200)
        invalid.json.side_effect = requests.exceptions.JSONDecodeError("bad", "", 0)
        errors = []
        with patch.object(update_check.requests, "get", side_effect=(response, invalid)):
            self.assertIsNone(update_check.get_latest_release(errors=errors))
        self.assertEqual(errors, ["GitHub：HTTP 403：访问被拒绝", "Gitee：响应不是有效 JSON"])
        with patch.object(update_check.requests, "get", side_effect=requests.ConnectTimeout()):
            self.assertIsNone(update_check.get_latest_release())

    def test_release_check_falls_back_to_gitee(self):
        with patch.object(
            update_check,
            "_release_version",
            side_effect=(None, "1.5.0"),
        ) as fetch:
            result = update_check.get_latest_release()

        self.assertEqual(result, "1.5.0")
        self.assertIn("api.github.com", fetch.call_args_list[0].args[0])
        self.assertIn("gitee.com", fetch.call_args_list[1].args[0])

    def test_project_links_need_no_source_constants(self):
        self.assertIn("github.com", update_check.project_links()["project"])
        self.assertIn(
            "gitee.com",
            update_check.project_links(use_gitee=True)["project"],
        )


if __name__ == "__main__":
    unittest.main()
