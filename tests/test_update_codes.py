import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

from scripts import update_codes


def _frame(code: str, name: str) -> pd.DataFrame:
    return pd.DataFrame(
        [(code, name, "", "美", "us")],
        columns=update_codes._DF_COLUMNS,
    )


class UpdateCodeFilesTests(unittest.TestCase):
    def test_long_running_tasks_are_scheduled_first(self):
        filenames = [task[0] for task in update_codes._tasks()]
        self.assertEqual(
            filenames[:3],
            ["stock_us.json", "stock_hk.json", "futures_sh.json"],
        )

    def test_szse_download_relies_on_task_level_retry(self):
        with (
            patch.object(
                update_codes.requests,
                "get",
                side_effect=update_codes.requests.ConnectionError("temporary"),
            ) as request,
            patch.object(update_codes.time, "sleep") as sleep,
        ):
            with self.assertRaises(update_codes.requests.ConnectionError):
                update_codes._szse_xlsx("1110", "tab1", "https://www.szse.cn/")

        request.assert_called_once()
        sleep.assert_not_called()

    def test_successful_unchanged_and_failed_files_are_recorded_independently(self):
        tasks = [
            ("same.json", "相同", None, (), {}),
            ("changed.json", "变化", None, (), {}),
            ("failed.json", "失败", None, (), {}),
        ]
        same_codes = update_codes._df_to_dict(_frame("same", "Same"))
        failed_codes = update_codes._df_to_dict(_frame("failed", "Failed"))

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "same.json").write_text(
                json.dumps({"last_update": "2026-08-27", "codes": same_codes}),
                encoding="utf-8",
            )
            failed_path = root / "failed.json"
            failed_path.write_text(
                json.dumps({"last_update": "2026-08-25", "codes": failed_codes}),
                encoding="utf-8",
            )
            (root / update_codes.STATUS_FILE).write_text(
                json.dumps(
                    {
                        "files": {
                            "failed.json": {
                                "last_checked": "2026-08-26",
                                "last_update": "2026-08-25",
                            }
                        }
                    }
                ),
                encoding="utf-8",
            )
            failed_before = failed_path.read_bytes()

            with (
                patch.object(update_codes, "_tasks", return_value=tasks),
                patch.object(
                    update_codes,
                    "_run_tasks",
                    return_value={
                        "same.json": (_frame("same", "Same"), None),
                        "changed.json": (_frame("changed", "Changed"), None),
                        "failed.json": (None, "network error"),
                    },
                ),
                patch.object(
                    update_codes,
                    "_iso_now",
                    side_effect=(
                        "2026-08-28T09:00:00+08:00",
                        "2026-08-28T09:01:00+08:00",
                    ),
                ),
            ):
                status = update_codes.update_code_files(directory)

            self.assertFalse(status["files"]["same.json"]["updated"])
            self.assertEqual(status["files"]["same.json"]["last_update"], "2026-08-27")
            self.assertTrue(status["files"]["changed.json"]["updated"])
            self.assertEqual(status["files"]["changed.json"]["last_update"], "2026-08-28")
            self.assertTrue(status["files"]["failed.json"]["error"])
            self.assertEqual(status["files"]["failed.json"]["last_checked"], "2026-08-26")
            self.assertEqual(failed_path.read_bytes(), failed_before)

    def test_task_gets_two_retries(self):
        calls = 0

        def flaky():
            nonlocal calls
            calls += 1
            if calls < 3:
                raise RuntimeError("temporary")
            return _frame("ok", "OK")

        task = ("test.json", "测试", flaky, (), {})
        with patch.object(update_codes.time, "sleep"):
            frame, error = update_codes._run_task(task)
        self.assertIsNone(error)
        self.assertEqual(len(frame), 1)
        self.assertEqual(calls, 3)


if __name__ == "__main__":
    unittest.main()
