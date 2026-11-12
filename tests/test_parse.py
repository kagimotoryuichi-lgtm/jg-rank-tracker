import json
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from tracker.parse import is_own, parse_dataforseo, parse_yahoo_jp  # noqa: E402
import run  # noqa: E402

FX = ROOT / "tests" / "fixtures"
BRANDS = ["japan green vietnam", "ジャパン・グリーン・ベトナム"]


class ParseTest(unittest.TestCase):
    def test_is_own(self):
        self.assertTrue(is_own("https://www.japangreen.vn/x/", "japangreen.vn"))
        self.assertTrue(is_own("japangreen.vn", "japangreen.vn"))
        self.assertFalse(is_own("https://notjapangreen.vn.example.com/", "japangreen.vn"))
        self.assertFalse(is_own("https://japangreen.vn.evil.com/", "japangreen.vn"))
        self.assertFalse(is_own("", "japangreen.vn"))

    def test_dataforseo(self):
        t = json.loads((FX / "dataforseo_task.json").read_text(encoding="utf-8"))
        p = parse_dataforseo(t, "japangreen.vn", BRANDS)
        self.assertEqual(p["rank"], 3)  # 広告・地図枠は数えず、organic の 3 番目
        self.assertEqual(p["url"], "https://www.japangreen.vn/outpatient-consultation/")
        self.assertEqual(p["local_rank"], 2)  # 地図枠は名前で一致
        self.assertTrue(p["has_local_pack"])
        self.assertEqual(p["top10"][:3], ["dym.asia", "www.vn.emb-japan.go.jp", "www.japangreen.vn"])

    def test_dataforseo_empty_and_error(self):
        p = parse_dataforseo({"status_code": 40000, "result": None}, "japangreen.vn", BRANDS)
        self.assertEqual(p["rank"], "")
        self.assertEqual(p["local_rank"], "")
        self.assertFalse(p["has_local_pack"])

    def test_yahoo(self):
        items = json.loads((FX / "apify_yahoo_jp.json").read_text(encoding="utf-8"))
        p = parse_yahoo_jp(items, "ジャパングリーンクリニック", "japangreen.vn")
        self.assertEqual(p["rank"], 1)
        q = parse_yahoo_jp(items, "ハノイ 病院", "japangreen.vn")
        self.assertEqual(q["rank"], "")
        self.assertEqual(q["top10"], ["dym.asia", "lotus-clinic.com"])
        r = parse_yahoo_jp(items, "存在しない語", "japangreen.vn")
        self.assertEqual(r["top10"], [])


class JobsTest(unittest.TestCase):
    def test_counts(self):
        cfg = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
        jobs = run.build_jobs(cfg)
        by = {}
        for j in jobs:
            by[j["engine"]] = by.get(j["engine"], 0) + 1
        self.assertEqual(by, {"google": 156, "bing": 52, "yahoo_jp": 17})  # google = PC 52 + iPhone 52 + Android 52（2026-11-12）
        self.assertEqual(len({run.tag_of(j) for j in jobs}), len(jobs))  # tag が一意
        self.assertEqual(len(run.build_jobs(cfg, limit=2, yahoo=False)), 4 * 2 * 2 * 2)
        prof = {}
        for j in jobs:
            prof[j["profile"]] = prof.get(j["profile"], 0) + 1
        self.assertEqual(prof, {"chrome-pc": 52, "edge-pc": 52, "iphone": 52, "android": 52, "yahoo-pc": 17})
        android = [j for j in jobs if j["profile"] == "android"][0]
        self.assertEqual((android["engine"], android["device"], android["os"]), ("google", "mobile", "android"))
        iphone = [j for j in jobs if j["profile"] == "iphone"][0]
        self.assertEqual((iphone["engine"], iphone["device"], iphone["os"]), ("google", "mobile", "ios"))


class SheetTestModeTest(unittest.TestCase):
    def test_writes_sample_rows_to_ranks_test(self):
        import argparse
        import os
        from unittest import mock
        cfg = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
        args = argparse.Namespace(dry_run=True, limit=1, no_yahoo=False, sheet_test=True)
        env = {"SHEET_WEBAPP_URL": "https://example.invalid/exec", "SHEET_TOKEN": "dummy"}
        with mock.patch.dict(os.environ, env), mock.patch("tracker.sheets.append_rows") as ap, \
                mock.patch.object(run, "ROOT", ROOT):
            rows, errors = run.run(cfg, args, log=lambda *_: None)
        ap.assert_called_once()
        self.assertEqual(ap.call_args.kwargs["worksheet"], "ranks_test")
        self.assertEqual(len(rows), 4 * 4 + 1)  # chrome-pc・edge-pc・iphone・android × 4 地点 + Yahoo! JAPAN 1
        self.assertEqual(set(rows[0]), set(run.HEADER))  # 新しい列も全部ある
        self.assertEqual({r["browser_profile"] for r in rows}, {"chrome-pc", "edge-pc", "iphone", "android", "yahoo-pc"})
        self.assertEqual(errors, [])
        self.assertTrue(all(r["status"] == "dry-run" for r in rows))


class AppsScriptClientTest(unittest.TestCase):
    def _fake(self, text):
        from unittest import mock
        res = mock.MagicMock()
        res.__enter__.return_value.read.return_value = text.encode("utf-8")
        res.__enter__.return_value.geturl.return_value = "https://script.googleusercontent.com/x"
        return mock.patch("urllib.request.urlopen", return_value=res)

    def test_posts_token_in_body_and_checks_ok(self):
        from tracker import sheets
        with self._fake('{"ok": true, "appended": 1, "sheet": "ranks_test"}') as uo:
            out = sheets.append_rows("https://example.invalid/exec", "tok", [{"date": "2026-11-05", "rank": 3}],
                                     worksheet="ranks_test")
        self.assertEqual(out["appended"], 1)
        req = uo.call_args.args[0]
        self.assertNotIn("tok", req.full_url)  # 合言葉は URL に載せない
        sent = json.loads(req.data.decode("utf-8"))
        self.assertEqual(sent["token"], "tok")
        self.assertEqual(sent["header"], sheets.HEADER)
        self.assertEqual(sent["rows"][0]["rank"], 3)
        self.assertEqual(sent["rows"][0]["url"], "")

    def test_raises_on_error_or_html(self):
        from tracker import sheets
        with self._fake('{"ok": false, "error": "bad token"}'):
            with self.assertRaises(sheets.SheetWriteError):
                sheets.append_rows("https://example.invalid/exec", "x", [])
        with self._fake("<html>sign in</html>"):
            with self.assertRaises(sheets.SheetWriteError):
                sheets.append_rows("https://example.invalid/exec", "x", [])


class RetryNoResultsTest(unittest.TestCase):
    def test_bing_40102_is_retried_until_ok(self):
        """Bing の 40102 は送り直し、2 回目で取れたら ok として記録する。"""
        import argparse
        import os
        from unittest import mock
        cfg = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
        ok = json.loads((FX / "dataforseo_task.json").read_text(encoding="utf-8"))
        posted = []

        class FakeDFS:
            def __init__(self, *a):
                self.n = 0

            def post_tasks(self, eng, tasks, out=None, failed=None):
                out = [] if out is None else out
                for t in tasks:
                    self.n += 1
                    posted.append((eng, t["tag"]))
                    out.append((t["tag"], f"{eng}-{self.n}", 0.0018))
                return out

            def wait_all(self, ids, *a):
                # 最初に送った bing の 1 件目だけ 40102、ほかは成功
                return {tid: ({"status_code": 40102, "status_message": "No Search Results."}
                              if tid == first_bing[0] else ok) for tid in ids}

        first_bing = []
        orig_post = FakeDFS.post_tasks

        def post(self, eng, tasks, out=None, failed=None):
            out = orig_post(self, eng, tasks, out, failed)
            if eng == "bing" and not first_bing:
                first_bing.append(out[0][1])
            return out
        FakeDFS.post_tasks = post
        args = argparse.Namespace(dry_run=False, limit=1, no_yahoo=True, sheet_test=False)
        import tempfile
        with tempfile.TemporaryDirectory() as tmp, mock.patch("tracker.clients.DataForSEO", FakeDFS), \
                mock.patch.dict(os.environ, {"SHEET_WEBAPP_URL": ""}), mock.patch.object(run, "ROOT", Path(tmp)):
            rows, errors = run.run(cfg, args, log=lambda *_: None)
        self.assertEqual(errors, [])
        self.assertEqual(len(rows), 4 * 4)
        self.assertEqual(sum(1 for e, _ in posted if e == "bing"), 4 + 1)  # 1 件だけ送り直し


class SecretSafetyTest(unittest.TestCase):
    """2026-11-05: 改行入りの APIFY_TOKEN でヘッダーが不正になり、例外文に鍵がそのまま入ってシートに書かれた。"""

    def test_env_strips_newline_and_redact_masks(self):
        import os
        from unittest import mock
        fake = "apify_api_" + "Ab12" * 8
        with mock.patch.dict(os.environ, {"APIFY_TOKEN": fake + "\n", "DFS_PASSWORD": "pw-secret-123"}):
            self.assertEqual(run.env("APIFY_TOKEN"), fake)
            msg = run.redact(f"Invalid header value b'Bearer {fake}\\n' / pw-secret-123")
        self.assertNotIn(fake, msg)
        self.assertNotIn("pw-secret-123", msg)
        self.assertNotIn("Ab12Ab12", run.redact("token apify_api_Ab12Ab12Ab12 leaked"))

    def test_yahoo_error_row_has_no_token(self):
        import argparse
        import os
        import tempfile
        from unittest import mock
        cfg = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
        fake = "apify_api_" + "Zz98" * 8

        def boom(token, *a, **k):
            raise ValueError(f"Invalid header value b'Bearer {token}\\n'")
        args = argparse.Namespace(dry_run=False, limit=2, no_yahoo=False, sheet_test=False, only_yahoo=True)
        with tempfile.TemporaryDirectory() as tmp, mock.patch("tracker.clients.apify_yahoo_jp", boom), \
                mock.patch.dict(os.environ, {"APIFY_TOKEN": fake, "SHEET_WEBAPP_URL": ""}), \
                mock.patch.object(run, "ROOT", Path(tmp)):
            rows, errors = run.run(cfg, args, log=lambda *_: None)
        self.assertEqual(len(rows), 2)
        self.assertTrue(all(r["engine"] == "yahoo_jp" for r in rows))
        self.assertTrue(all(fake not in r["status"] for r in rows))


    def test_rows_recorded_counts_only_ok_rows_of_the_day(self):
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "ranks.csv"
            lines = ["date,status", "2026-11-07,ok", "2026-11-07,error 40102", "2026-11-06,ok", "2026-11-07,ok"]
            p.write_text("\n".join(lines) + "\n", encoding="utf-8")
            self.assertEqual(run.rows_recorded("2026-11-07", p), 2)
            self.assertEqual(run.rows_recorded("2026-11-08", p), 0)
            self.assertEqual(run.rows_recorded("2026-11-07", Path(tmp) / "none.csv"), 0)


if __name__ == "__main__":
    unittest.main()


COMP = [{"id": "sakura", "domains": ["sakurahanoi.com"]}, {"id": "dym", "domains": ["jp-dymmedicalcenter.com.vn", "dym.asia"]}]


def _task(items):
    return {"status_code": 20000, "result": [{"items": items, "se_results_count": 1, "check_url": "x"}]}


class FeaturesTest(unittest.TestCase):
    """2026-11-12 追加: 競合の順位・AI 概要・強調スニペット・他の人はこちらも質問"""

    def test_competitors_best_rank(self):
        items = [{"type": "organic", "rank_group": 1, "url": "https://dym.asia/a"},
                 {"type": "organic", "rank_group": 2, "url": "https://www.japangreen.vn/"},
                 {"type": "organic", "rank_group": 4, "url": "https://jp-dymmedicalcenter.com.vn/b"}]
        p = parse_dataforseo(_task(items), "japangreen.vn", BRANDS, COMP, "google")
        self.assertEqual(p["rank"], 2)
        self.assertEqual(p["rank_dym"], 1)  # 2 つのドメインのうち上の方
        self.assertEqual(p["rank_sakura"], "")  # 圏外は空欄

    def test_ai_overview_cites_own(self):
        items = [{"type": "ai_overview", "items": [{"type": "ai_overview_element", "references": [
                    {"type": "ai_overview_reference", "domain": "www.japangreen.vn", "url": "https://www.japangreen.vn/x/"}]}]},
                 {"type": "organic", "rank_group": 1, "url": "https://sakurahanoi.com/"}]
        p = parse_dataforseo(_task(items), "japangreen.vn", BRANDS, COMP, "google")
        self.assertEqual((p["ai_overview"], p["ai_overview_own"]), ("TRUE", "TRUE"))
        self.assertEqual(p["rank_sakura"], 1)
        self.assertEqual((p["featured_snippet_own"], p["paa_own"]), ("", ""))  # 枠が出ていないときは空欄

    def test_ai_overview_without_own_and_other_boxes(self):
        items = [{"type": "ai_overview", "references": [{"url": "https://dym.asia/"}]},
                 {"type": "featured_snippet", "domain": "lotus-clinic.com", "url": "https://lotus-clinic.com/"},
                 {"type": "people_also_ask", "items": [{"type": "people_also_ask_element", "expanded_element": [
                     {"type": "people_also_ask_expanded_element", "url": "https://www.japangreen.vn/q/"}]}]}]
        p = parse_dataforseo(_task(items), "japangreen.vn", BRANDS, COMP, "google")
        self.assertEqual((p["ai_overview"], p["ai_overview_own"]), ("TRUE", "FALSE"))
        self.assertEqual((p["featured_snippet_own"], p["paa_own"]), ("FALSE", "TRUE"))

    def test_no_ai_overview_google_vs_bing(self):
        items = [{"type": "organic", "rank_group": 1, "url": "https://www.japangreen.vn/"}]
        g = parse_dataforseo(_task(items), "japangreen.vn", BRANDS, COMP, "google")
        b = parse_dataforseo(_task(items), "japangreen.vn", BRANDS, COMP, "bing")
        self.assertEqual((g["ai_overview"], g["ai_overview_own"]), ("FALSE", ""))
        self.assertEqual((b["ai_overview"], b["ai_overview_own"]), ("", ""))  # Bing は対象外

    def test_yahoo_competitors(self):
        items = json.loads((FX / "apify_yahoo_jp.json").read_text(encoding="utf-8"))
        q = parse_yahoo_jp(items, "ハノイ 病院", "japangreen.vn", COMP)
        self.assertEqual(q["rank_dym"], 1)


class MigrateCsvTest(unittest.TestCase):
    def test_old_header_is_extended(self):
        import csv
        import tempfile
        from tracker.sheets import BASE_HEADER
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / "ranks.csv"
            with path.open("w", encoding="utf-8", newline="") as f:
                w = csv.writer(f)
                w.writerow(BASE_HEADER)
                w.writerow(["2026-11-11"] + ["x"] * (len(BASE_HEADER) - 1))
            self.assertTrue(run.migrate_csv(path))
            rows = list(csv.reader(path.open(encoding="utf-8", newline="")))
            self.assertEqual(rows[0], run.HEADER)
            self.assertEqual(len(rows[1]), len(run.HEADER))
            self.assertEqual(rows[1][0], "2026-11-11")
            self.assertFalse(run.migrate_csv(path))  # 2 回目は何もしない

    def test_task_has_async_ai_overview_only_for_google(self):
        cfg = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
        self.assertTrue(cfg["ai_overview_async"])
