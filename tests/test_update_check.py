import unittest
from unittest.mock import patch

from stockwidget.data import update_check


class UpdateCheckTests(unittest.TestCase):
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
