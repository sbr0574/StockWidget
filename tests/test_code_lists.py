import unittest
from datetime import datetime, timedelta, timezone
from unittest.mock import patch

from stockwidget.constants import CODES_STATUS_FILE, CODE_LIST_FILES
from stockwidget.data.code_lists import CodeListManager, expected_status_date


UTC8 = timezone(timedelta(hours=8))


def _payload(filename: str, update_date: str) -> dict:
    code = filename.removesuffix(".json")
    return {
        "last_update": update_date,
        "codes": {
            code: {
                "code": code,
                "market": "",
                "type": "",
                "name": filename,
                "name_en": "",
                "py": "",
                "abbr": "",
            }
        },
    }


class ExpectedStatusDateTests(unittest.TestCase):
    def test_weekday_before_nine_uses_previous_workday(self):
        monday = datetime(2026, 8, 31, 8, 59, tzinfo=UTC8)
        self.assertEqual(expected_status_date(monday), "2026-08-28")

    def test_weekday_after_nine_uses_today(self):
        monday = datetime(2026, 8, 31, 9, 0, tzinfo=UTC8)
        self.assertEqual(expected_status_date(monday), "2026-08-31")

    def test_weekend_uses_friday(self):
        saturday = datetime(2026, 8, 29, 12, 0, tzinfo=UTC8)
        self.assertEqual(expected_status_date(saturday), "2026-08-28")


class CodeListManagerTests(unittest.TestCase):
    def test_sync_downloads_only_files_with_different_content_date(self):
        old_date = "2026-08-27"
        expected = "2026-08-28"
        resources = {name: _payload(name, old_date) for name in CODE_LIST_FILES}
        files = {
            name: {
                "last_checked": expected,
                "last_update": expected if name == "stock_hk.json" else old_date,
                "updated": name == "stock_hk.json",
                "error": False,
            }
            for name in CODE_LIST_FILES
        }
        status = {"run_date": expected, "files": files}
        new_hk = _payload("stock_hk-new.json", expected)
        fetched = []

        def fetcher(url):
            fetched.append(url)
            if url.endswith("/" + CODES_STATUS_FILE):
                return status
            if url.endswith("/stock_hk.json"):
                return new_hk
            return None

        with (
            patch(
                "stockwidget.data.code_lists.load_json_from_resource",
                side_effect=lambda path: resources.get(path[2:], {}),
            ),
            patch("stockwidget.data.code_lists.load_file", return_value={}),
            patch("stockwidget.data.code_lists.save_file") as save,
        ):
            manager = CodeListManager(fetcher=fetcher)
            manager.load_local()
            result = manager.sync_remote(
                now=datetime(2026, 8, 28, 10, 0, tzinfo=UTC8)
            )

        self.assertEqual(result["state"], "current")
        self.assertEqual(result["updated"], ("stock_hk.json",))
        self.assertEqual(sum(url.endswith(".json") for url in fetched), 2)
        save.assert_any_call(status, CODES_STATUS_FILE)
        save.assert_any_call(new_hk, "stock_hk.json")

    def test_remote_file_falls_back_from_github_to_gitee(self):
        fetched = []

        def fetcher(url):
            fetched.append(url)
            return {"ok": True} if "gitee" in url else None

        manager = CodeListManager(fetcher=fetcher)
        self.assertEqual(manager._fetch_remote("test.json"), {"ok": True})
        self.assertIn("raw.githubusercontent.com", fetched[0])
        self.assertIn("gitee", fetched[1])


if __name__ == "__main__":
    unittest.main()
