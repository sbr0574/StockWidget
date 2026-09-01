import json
import tempfile
import unittest
from datetime import date
from io import BytesIO
from pathlib import Path
from unittest.mock import Mock, patch

import pandas as pd
from openpyxl import Workbook

from scripts import update_codes


def _frame(code: str, name: str) -> pd.DataFrame:
    return pd.DataFrame(
        [(code, name, "", "美", "us")],
        columns=update_codes._DF_COLUMNS,
    )


def _xlsx_bytes(headers: tuple[str, ...], rows: list[tuple]) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "ListOfSecurities"
    sheet.append(("title",))
    sheet.append(("updated",))
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def _xlsx_response(content: bytes) -> Mock:
    response = Mock()
    response.content = content
    response.headers = {
        "Content-Type": (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    }
    return response


def _text_response(text: str) -> Mock:
    response = Mock()
    response.content = text.encode("utf-8")
    return response


def _us_directory_texts() -> tuple[str, str]:
    nasdaq = "\n".join(
        (
            "Symbol|Security Name|Market Category|Test Issue|Financial Status|Round Lot Size|ETF|NextShares",
            "AAPL|Apple Inc. - Common Stock|Q|N|N|100|N|N",
            "QQQ|Invesco QQQ Trust, Series 1|Q|N|N|100|Y|N",
            "ZXIET|Test Security|Q|Y|N|100|N|N",
            "File Creation Time: 0831202621:31|||||||",
        )
    )
    other = "\n".join(
        (
            "ACT Symbol|Security Name|Exchange|CQS Symbol|ETF|Round Lot Size|Test Issue|NASDAQ Symbol",
            "BRK.B|Berkshire Hathaway Class B|N|BRK.B|N|40|N|BRK.B",
            "AAC.W|Ares Acquisition Warrant|N|AAC.WS|N|100|N|AAC+",
            "ABR$D|Arbor Realty Preferred D|N|ABRpD|N|100|N|ABR-D",
            "SPY|SPDR S&P 500 ETF Trust|P|SPY|Y|100|N|SPY",
            "TEST|Exchange Test Security|N|TEST|N|100|Y|TEST",
            "File Creation Time: 0831202621:31||||||",
        )
    )
    return nasdaq, other


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

    def test_hkex_list_includes_equities_funds_and_reits(self):
        english = _xlsx_bytes(
            ("Stock Code", "Name of Securities", "Category"),
            [
                ("00001", "CKH HOLDINGS", "Equity"),
                ("02800", "TRACKER FUND", "Exchange Traded Products"),
                ("00405", "YUEXIU REIT", "Real Estate Investment Trusts"),
                ("04000", "TEST BOND", "Debt Securities"),
            ],
        )
        chinese = _xlsx_bytes(
            ("股份代號", "股份名稱"),
            [
                ("00001", "長和"),
                ("02800", "盈富基金"),
                ("00405", "越秀房產信託基金"),
                ("04000", "測試債券"),
            ],
        )

        with patch.object(
            update_codes.requests,
            "get",
            side_effect=[_xlsx_response(english), _xlsx_response(chinese)],
        ) as request:
            frame = update_codes._stock_hk_name_code()

        self.assertEqual(frame["code"].tolist(), ["00001", "02800", "00405"])
        by_code = frame.set_index("code")
        self.assertEqual(by_code.at["00001", "name"], "長和")
        self.assertEqual(by_code.at["00001", "name_en"], "CKH HOLDINGS")
        self.assertEqual(by_code.at["00001", "type"], "港")
        self.assertEqual(by_code.at["02800", "type"], "基")
        self.assertEqual(by_code.at["00405", "type"], "基")
        self.assertNotIn("04000", by_code.index)
        self.assertEqual(
            [call.args[0] for call in request.call_args_list],
            [
                update_codes._HKEX_SECURITIES_EN_URL,
                update_codes._HKEX_SECURITIES_ZH_URL,
            ],
        )

    def test_nasdaq_official_lists_keep_stocks_and_etfs_as_us(self):
        nasdaq, other = _us_directory_texts()
        with tempfile.TemporaryDirectory() as directory:
            cache_path = str(Path(directory) / "aliases.json")
            Path(cache_path).write_text(
                json.dumps(
                    {"last_update": "2026-08-03", "aliases": {"old": "旧证券"}},
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            with (
                patch.object(
                    update_codes.requests,
                    "get",
                    side_effect=[_text_response(nasdaq), _text_response(other)],
                ) as request,
                patch.object(update_codes, "_fetch_us_cn_aliases") as fetch_aliases,
            ):
                frame = update_codes._stock_us_name_code(
                    today=date(2026, 8, 31),
                    output_dir=directory,
                    cache_path=cache_path,
                )

        self.assertEqual(
            frame["code"].tolist(),
            ["aapl", "qqq", "brk_b", "aac_ws", "abr_d", "spy"],
        )
        self.assertEqual(set(frame["type"]), {"美"})
        self.assertEqual(set(frame["market"]), {"us"})
        self.assertNotIn("zxiet", set(frame["code"]))
        self.assertNotIn("test", set(frame["code"]))
        fetch_aliases.assert_not_called()
        self.assertEqual(
            [call.args[0] for call in request.call_args_list],
            [
                update_codes._NASDAQ_LISTED_URL,
                update_codes._NASDAQ_OTHER_LISTED_URL,
            ],
        )

    def test_us_names_use_git_cached_aliases_without_daily_eastmoney_request(self):
        nasdaq, other = _us_directory_texts()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            cache_path = root / "cache_us_cn_aliases.json"
            cache_path.write_text(
                json.dumps(
                    {
                        "last_update": "2026-08-03",
                        "aliases": {"aapl": "苹果"},
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            with (
                patch.object(
                    update_codes.requests,
                    "get",
                    side_effect=[_text_response(nasdaq), _text_response(other)],
                ),
                patch.object(update_codes, "_fetch_us_cn_aliases") as fetch_aliases,
            ):
                frame = update_codes._stock_us_name_code(
                    today=date(2026, 8, 31),
                    output_dir=directory,
                    cache_path=str(cache_path),
                )

            by_code = frame.set_index("code")
            self.assertEqual(by_code.at["aapl", "name"], "苹果")
            self.assertEqual(
                by_code.at["aapl", "name_en"], "Apple Inc. - Common Stock"
            )
            fetch_aliases.assert_not_called()
            self.assertTrue(cache_path.exists())

    def test_us_aliases_refresh_when_month_changes(self):
        nasdaq, other = _us_directory_texts()
        with tempfile.TemporaryDirectory() as directory:
            cache_path = str(Path(directory) / "aliases.json")
            Path(cache_path).write_text(
                json.dumps(
                    {"last_update": "2026-08-03", "aliases": {"aapl": "旧苹果"}},
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            with (
                patch.object(
                    update_codes.requests,
                    "get",
                    side_effect=[_text_response(nasdaq), _text_response(other)],
                ),
                patch.object(
                    update_codes,
                    "_fetch_us_cn_aliases",
                    return_value={
                        "aapl": "苹果",
                        "spy": "标普500ETF",
                        "not_listed": "已退市测试",
                    },
                ) as fetch_aliases,
            ):
                frame = update_codes._stock_us_name_code(
                    today=date(2026, 9, 1),
                    output_dir=directory,
                    cache_path=cache_path,
                )

            by_code = frame.set_index("code")
            self.assertEqual(by_code.at["aapl", "name"], "苹果")
            self.assertEqual(by_code.at["spy", "name"], "标普500ETF")
            self.assertEqual(by_code.at["qqq", "name"], "Invesco QQQ Trust, Series 1")
            self.assertEqual(set(frame["type"]), {"美"})
            fetch_aliases.assert_called_once_with()
            cache = json.loads(Path(cache_path).read_text(encoding="utf-8"))
            self.assertEqual(cache["last_update"], "2026-09-01")
            self.assertNotIn("not_listed", cache["aliases"])

    def test_us_aliases_are_not_fetched_again_in_same_month(self):
        nasdaq, other = _us_directory_texts()
        with tempfile.TemporaryDirectory() as directory:
            cache_path = Path(directory) / "aliases.json"
            cache_path.write_text(
                json.dumps(
                    {"last_update": "2026-09-01", "aliases": {"aapl": "苹果"}},
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            with (
                patch.object(
                    update_codes.requests,
                    "get",
                    side_effect=[_text_response(nasdaq), _text_response(other)],
                ),
                patch.object(update_codes, "_fetch_us_cn_aliases") as fetch_aliases,
            ):
                frame = update_codes._stock_us_name_code(
                    today=date(2026, 9, 1),
                    output_dir=directory,
                    cache_path=str(cache_path),
                )

            self.assertEqual(frame.set_index("code").at["aapl", "name"], "苹果")
            fetch_aliases.assert_not_called()

    def test_us_alias_refresh_failure_does_not_block_official_list(self):
        nasdaq, other = _us_directory_texts()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            cache_path = root / "aliases.json"
            cache_path.write_text(
                json.dumps(
                    {"last_update": "2026-08-03", "aliases": {"aapl": "苹果"}},
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            with (
                patch.object(
                    update_codes.requests,
                    "get",
                    side_effect=[_text_response(nasdaq), _text_response(other)],
                ),
                patch.object(
                    update_codes,
                    "_fetch_us_cn_aliases",
                    side_effect=RuntimeError("eastmoney unavailable"),
                ),
            ):
                frame = update_codes._stock_us_name_code(
                    today=date(2026, 9, 1),
                    output_dir=directory,
                    cache_path=str(cache_path),
                )

            self.assertEqual(frame.set_index("code").at["aapl", "name"], "苹果")
            self.assertEqual(len(frame), 6)

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
