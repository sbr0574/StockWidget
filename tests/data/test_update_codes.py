"""服务端代码抓取的分页恢复、完整性、缓存保护与失败状态。"""

from concurrent.futures import ThreadPoolExecutor
from contextlib import redirect_stdout
from datetime import date
from io import StringIO
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch

import requests

from scripts import update_codes as updater


def _official_us_responses():
    nasdaq = Mock(content=(
        b"Symbol|Security Name|Test Issue\nAAPL|Apple Inc.|N\n"
    ))
    other = Mock(content=(
        b"CQS Symbol|Security Name|Test Issue\nSPY|SPDR S&P 500 ETF|N\n"
    ))
    return [nasdaq, other]


def _bse_page(page=0, total=3, size=2):
    records = [
        {"xxzqdm": f"{920000 + number:06}", "xxzqjc": f"北交股票{number}",
         "xxywjc": f"BSE STOCK {number}", "xxfcbj": "2", "xxzqjb": "T"}
        for number in range(page * size, min((page + 1) * size, total))
    ]
    return {"number": page, "numberOfElements": len(records), "size": size,
            "totalElements": total, "totalPages": (total + size - 1) // size,
            "content": records}


def _bse_response(payload, callback="null"):
    body = json.dumps([payload], ensure_ascii=False)
    return Mock(text=f"{callback}({body});" if callback else body)


def _write_manual_indices(root):
    for filename in updater.MANUAL_FILES:
        (root / filename).write_text(json.dumps({
            "last_update": "2026-10-21",
            "codes": {"manual_index": {"name": "手工添加的指数"}},
        }, ensure_ascii=False, indent=3) + "\n", encoding="utf-8")




class SinaAliasBatchTests(unittest.TestCase):
    def setUp(self):
        self.clock = 0.0
        self.enterContext(patch.object(updater.time, "sleep", side_effect=self.advance))
        self.enterContext(patch.object(updater.time, "monotonic", side_effect=lambda: self.clock))
        self.enterContext(patch.object(updater, "_last_request_started", None))
        self.output = self.enterContext(redirect_stdout(StringIO()))

    def advance(self, seconds):
        self.clock += seconds

    @staticmethod
    def response(codes):
        return Mock(text='\n'.join(f'var hq_str_gb_{code}="证券{code},10,0";' for code in codes))

    def test_failed_batch_resumes_without_repeating_completed_codes(self):
        codes = {f"a{n:03}" for n in range(152)}
        with patch.object(updater.requests, "get", side_effect=[
            self.response(sorted(codes)[:150]), requests.ConnectionError("private proxy"),
            self.response(sorted(codes)[150:]),
        ]) as get:
            aliases = updater._fetch_us_cn_aliases(codes)
        self.assertEqual(set(aliases), codes)
        self.assertEqual([len(c.args[0].split('list=')[1].split(',')) for c in get.call_args_list], [150,2,2])
        self.assertEqual(get.call_args_list[1].args[0], get.call_args_list[2].args[0])
        self.assertNotIn("private proxy", self.output.getvalue())

    def test_aliases_use_requested_symbols_and_chinese_names_only(self):
        response=Mock(text='var hq_str_gb_aapl="苹果,10,0";\nvar hq_str_gb_spy="SPDR,10,0";\nvar hq_str_gb_other="无关,10,0";')
        with patch.object(updater.requests,"get",return_value=response) as get:
            aliases=updater._fetch_us_cn_aliases({"spy","aapl"})
        self.assertEqual(aliases,{"aapl":"苹果"})
        self.assertTrue(get.call_args.args[0].endswith("list=gb_aapl,gb_spy"))
        self.assertEqual(response.encoding,"gbk")
        response.close.assert_called_once()
        with patch.object(updater.requests,"get",return_value=self.response(["aapl"])):
            with self.assertRaises(updater.CodeListError):
                updater._fetch_us_cn_aliases({"aapl","nvda"})

    def test_denials_limits_and_bad_payloads_do_not_restart_task(self):
        failures=[requests.exceptions.SSLError("private certificate")]
        for status in (401,403,404,429):
            response=requests.Response();response.status_code=status
            failures.append(requests.HTTPError("private URL",response=response))
        for error in failures:
            with self.subTest(error=type(error).__name__),patch.object(updater.requests,"get",side_effect=error) as get:
                with self.assertRaises(updater.CodeListError):updater._fetch_us_cn_aliases({"aapl"})
                self.assertEqual(get.call_count,1)
        for body in ("<html>denied</html>",'var hq_str_gb_spy="无关,10,0";'):
            with self.subTest(body=body),patch.object(updater.requests,"get",return_value=Mock(text=body)) as get:
                with self.assertRaises(updater.CodeListError):updater._fetch_us_cn_aliases({"aapl"})
                self.assertEqual(get.call_count,1)

    def test_transient_retries_are_bounded_and_do_not_expose_raw_error(self):
        task=("aliases.json","美股中文名",updater._fetch_us_cn_aliases,({"aapl"},),{})
        with patch.object(updater.requests,"get",side_effect=requests.ReadTimeout("private URL")) as get:
            frame,error=updater._run_task(task)
        self.assertIsNone(frame)
        self.assertEqual(get.call_count,3)
        self.assertIn("第 1 批",error)
        self.assertIn("超时",error)
        self.assertNotIn("private URL",error)

    def test_batches_share_rate_limit_and_preserve_requests_environment(self):
        starts=[]
        def get(url,**kwargs):
            starts.append(self.clock)
            self.assertNotIn("verify",kwargs)
            self.assertNotIn("proxies",kwargs)
            return self.response([x.removeprefix("gb_") for x in url.split("list=")[1].split(",")])
        with patch.object(updater.requests,"get",side_effect=get):
            with ThreadPoolExecutor(max_workers=2) as pool:
                rows=list(pool.map(updater._fetch_us_cn_aliases,({"aapl"},{"nvda"})))
        self.assertEqual([len(r) for r in rows],[1,1])
        self.assertGreaterEqual(starts[1]-starts[0],.5)


class OfficialCodeListTests(unittest.TestCase):
    def setUp(self):
        self.enterContext(patch.object(updater.time, "sleep"))
        self.enterContext(redirect_stdout(StringIO()))

    def test_bse_official_jsonp_and_json_pages_keep_bj_identity_and_names(self):
        responses = [_bse_response(_bse_page(), "jQuery_123"),
                     _bse_response(_bse_page(1), None)]
        with patch.object(updater.requests, "post", side_effect=responses) as post:
            frame = updater._stock_bj_name_code()
        codes = updater._df_to_dict(frame)
        self.assertEqual(set(codes), {"bj920000", "bj920001", "bj920002"})
        self.assertEqual(codes["bj920000"]["name"], "北交股票0")
        self.assertEqual(codes["bj920000"]["name_en"], "BSE STOCK 0")
        self.assertTrue(all(v["market"] == "bj" and v["type"] == "京" for v in codes.values()))
        self.assertEqual([c.kwargs["data"]["page"] for c in post.call_args_list], ["0", "1"])
        self.assertTrue(all("bse.cn" in c.args[0] for c in post.call_args_list))
        for response in responses:
            response.close.assert_called_once()

    def test_bse_retry_resumes_failed_page_without_restarting(self):
        server = requests.Response()
        server.status_code = 502
        failed = _bse_response(_bse_page(1))
        failed.raise_for_status.side_effect = requests.HTTPError("private URL", response=server)
        with (
            patch.object(updater.requests, "post", side_effect=[
                _bse_response(_bse_page()), failed, _bse_response(_bse_page(1)),
            ]) as post,
            patch.object(updater.requests, "get") as get,
        ):
            records = updater.fetch_bse_list()
        self.assertEqual(len(records), 3)
        self.assertEqual([c.kwargs["data"]["page"] for c in post.call_args_list], ["0", "1", "1"])
        get.assert_not_called()
        failed.close.assert_called_once()

    def test_bse_denials_and_invalid_jsonp_are_not_retried(self):
        denied = requests.Response()
        denied.status_code = 403
        cases = [
            requests.HTTPError("https://user:password@host", response=denied),
            Mock(text="<html>please verify</html>"),
            Mock(text="callback([{}]); extra_code()"),
            Mock(text="callback(not_json);"),
            Mock(text="callback([]);"),
        ]
        for result in cases:
            with self.subTest(result=type(result).__name__):
                get_kwargs = {"side_effect": result} if isinstance(result, Exception) else {"return_value": result}
                with patch.object(updater.requests, "post", **get_kwargs) as post:
                    with self.assertRaises(updater.CodeListError) as caught:
                        updater.fetch_bse_list()
                self.assertEqual(post.call_count, 1)
                self.assertNotIn("password", str(caught.exception))

    def test_bse_incomplete_changed_duplicate_and_wrong_market_lists_are_rejected(self):
        variants = []
        missing = _bse_page(1)
        missing["content"] = []
        variants.append(missing)
        variants.append(_bse_page(1, total=4))
        repeated = _bse_page(1)
        repeated["content"][0] = _bse_page()["content"][0]
        variants.append(repeated)
        for field, value in (("xxfcbj", "1"), ("xxzqjb", "A"), ("xxzqdm", "6000000"), ("xxzqjc", "")):
            invalid = _bse_page(1)
            invalid["content"][0][field] = value
            variants.append(invalid)
        for last in variants:
            with self.subTest(last=last):
                with patch.object(updater.requests, "post", side_effect=[
                    _bse_response(_bse_page()), _bse_response(last),
                ]):
                    with self.assertRaises(updater.CodeListError):
                        updater.fetch_bse_list()

    def test_bse_failure_keeps_previous_file_and_manual_indices(self):
        tasks = [t for t in updater._tasks() if t[0] in {"stock_bj.json", "index_cn.json"}]
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            _write_manual_indices(root)
            old = '{"last_update":"2026-09-30","codes":{"bj920000":{}}}'
            (root / "stock_bj.json").write_text(old, encoding="utf-8")
            with (
                patch.object(updater, "_tasks", return_value=tasks),
                patch.object(updater.requests, "post", side_effect=requests.ReadTimeout("private proxy")) as post,
                patch.object(updater.requests, "get") as get,
            ):
                status = updater.update_code_files(directory)
            self.assertEqual((root / "stock_bj.json").read_text(encoding="utf-8"), old)
            self.assertTrue(status["files"]["stock_bj.json"]["error"])
            self.assertFalse(status["files"]["index_cn.json"]["error"])
            self.assertEqual(post.call_count, 3)
            get.assert_not_called()

    def test_manual_indices_keep_exact_bytes_and_report_manually_changed_dates(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            _write_manual_indices(root)
            originals = {name: (root / name).read_bytes() for name in updater.MANUAL_FILES}
            with (
                patch.object(updater, "_tasks", return_value=[]),
                patch.object(updater, "_run_tasks", return_value={}),
                patch.object(updater.requests, "get") as get,
                patch.object(updater.requests, "post") as post,
            ):
                status = updater.update_code_files(directory)
                changed = json.loads((root / "index_cn.json").read_text(encoding="utf-8"))
                changed["last_update"] = "2026-10-22"
                changed["codes"]["new_index"] = {"name": "维护者新增"}
                (root / "index_cn.json").write_text(json.dumps(changed), encoding="utf-8")
                originals["index_cn.json"] = (root / "index_cn.json").read_bytes()
                updated_status = updater.update_code_files(directory)
            for name, original in originals.items():
                self.assertEqual((root / name).read_bytes(), original)
                self.assertFalse(status["files"][name]["updated"])
                self.assertFalse(updated_status["files"][name]["error"])
            self.assertEqual(updated_status["files"]["index_cn.json"]["last_update"], "2026-10-22")
            self.assertEqual(updated_status["files"]["index_global.json"]["last_update"], "2026-10-21")
            get.assert_not_called()
            post.assert_not_called()

    def test_single_script_runs_from_a_directory_without_sibling_modules(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            script = root / "update_codes.py"
            script.write_bytes(Path(updater.__file__).read_bytes())
            output = root / "data"
            output.mkdir()
            _write_manual_indices(output)
            with patch.dict("os.environ", {"CODES_OUTPUT_DIR": str(output)}):
                spec = importlib.util.spec_from_file_location("standalone_codes", script)
                standalone = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(standalone)
            task = ("stock_sh.json", "沪市股票", lambda: standalone._rows_frame([
                ("600000", "浦发银行", "", "沪", "sh"),
            ]), (), {})
            with (
                patch.object(standalone, "_tasks", return_value=[task]),
                patch.object(standalone.requests, "get", side_effect=AssertionError("unexpected network")),
                patch.object(standalone.requests, "post", side_effect=AssertionError("unexpected network")),
            ):
                self.assertEqual(standalone.main(), 0)
            payload = json.loads((output / "stock_sh.json").read_text(encoding="utf-8"))
            self.assertIn("sh600000", payload["codes"])
            status = json.loads((output / standalone.STATUS_FILE).read_text(encoding="utf-8"))
            self.assertFalse(status["has_errors"])

    def test_missing_or_invalid_manual_indices_are_reported_without_replacement(self):
        for invalid in (None, "{broken json", '{"last_update":"2026-10-22","codes":{}}'):
            with self.subTest(invalid=invalid), tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / "index_cn.json"
                if invalid is not None:
                    path.write_text(invalid, encoding="utf-8")
                with (
                    patch.object(updater, "_tasks", return_value=[]),
                    patch.object(updater, "_run_tasks", return_value={}),
                ):
                    status = updater.update_code_files(directory)
                self.assertTrue(status["files"]["index_cn.json"]["error"])
                if invalid is None:
                    self.assertFalse(path.exists())
                else:
                    self.assertEqual(path.read_text(encoding="utf-8"), invalid)





class AliasUpdateStatusTests(unittest.TestCase):
    def test_refresh_failure_preserves_cached_names_and_reports_auxiliary_failure(self):
        for failure in (requests.ConnectionError("https://user:password@private"), {}):
            with self.subTest(failure=type(failure).__name__):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    cache_path = root / updater.US_CN_ALIAS_CACHE_FILE
                    cache_path.write_text(json.dumps({
                        "last_update": "2026-09-01", "aliases": {"aapl": "苹果"},
                    }, ensure_ascii=False), encoding="utf-8")
                    original = cache_path.read_bytes()
                    task = ("stock_us.json", "美股", updater._stock_us_name_code, (), {
                        "today": date(2026, 10, 22), "output_dir": directory,
                    })
                    fetch = Mock(side_effect=failure) if isinstance(failure, Exception) else Mock(return_value=failure)
                    output = StringIO()
                    with (
                        redirect_stdout(output),
                        patch.object(updater, "_tasks", return_value=[task]) as tasks,
                        patch.object(updater.requests, "get", side_effect=_official_us_responses()),
                        patch.object(updater, "_fetch_us_cn_aliases", fetch),
                        patch.object(updater, "_iso_now", return_value="2026-10-22T09:00:00+08:00"),
                    ):
                        status = updater.update_code_files(directory)
                    tasks.assert_called_once_with(output_dir=directory)
                    codes = json.loads((root / "stock_us.json").read_text(encoding="utf-8"))["codes"]
                    self.assertEqual(codes["usaapl"]["name"], "苹果")
                    self.assertEqual(cache_path.read_bytes(), original)
                    self.assertFalse(status["files"]["stock_us.json"]["error"])
                    auxiliary = status["auxiliary"][updater.US_CN_ALIAS_CACHE_FILE]
                    self.assertTrue(auxiliary["error"])
                    self.assertEqual(auxiliary["last_update"], "2026-09-01")
                    self.assertEqual(auxiliary["last_attempt"], "2026-10-22")
                    self.assertTrue(status["has_errors"])
                    self.assertNotIn("password", output.getvalue())
                    self.assertNotIn("private", json.dumps(status))

    def test_cache_write_failure_keeps_previous_names_for_this_run(self):
        with tempfile.TemporaryDirectory() as directory:
            cache = Path(directory) / updater.US_CN_ALIAS_CACHE_FILE
            cache.write_text(json.dumps({
                "last_update": "2026-09-01", "aliases": {"aapl": "旧苹果"},
            }, ensure_ascii=False), encoding="utf-8")
            with (
                redirect_stdout(StringIO()),
                patch.object(updater.requests, "get", side_effect=_official_us_responses()),
                patch.object(updater, "_fetch_us_cn_aliases", return_value={"aapl": "新苹果"}),
                patch.object(updater, "_save_json", side_effect=PermissionError("private path")),
            ):
                frame = updater._stock_us_name_code(date(2026, 10, 22), directory)
            self.assertEqual(frame.set_index("code").at["aapl", "name"], "旧苹果")
            self.assertTrue(frame.attrs["us_cn_aliases"]["error"])

    def test_unavailable_requested_names_do_not_create_aliases(self):
        with patch.object(updater.requests, "get", return_value=Mock(text='var hq_str_gb_aapl="";')):
            with self.assertRaises(updater.CodeListError):
                updater._fetch_us_cn_aliases({"aapl"})

    def test_direct_run_keeps_successful_results_when_a_category_fails(self):
        with patch.object(updater, "update_code_files", return_value={"has_errors": True}) as update:
            self.assertEqual(updater.main(), 0)
        update.assert_called_once()

    def test_fatal_errors_are_safe_and_nonzero(self):
        output = StringIO()
        with (
            redirect_stdout(output),
            patch.object(updater, "update_code_files", side_effect=requests.ConnectionError("https://user:password@host")),
        ):
            self.assertEqual(updater.main(), 1)
        self.assertNotIn("password", output.getvalue())


if __name__ == "__main__":
    unittest.main()
