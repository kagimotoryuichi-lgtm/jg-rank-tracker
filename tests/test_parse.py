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
        self.assertEqual(by, {"google": 52, "bing": 52, "yahoo_jp": 17})
        self.assertEqual(len({run.tag_of(j) for j in jobs}), len(jobs))  # tag が一意
        self.assertEqual(len(run.build_jobs(cfg, limit=2, yahoo=False)), 2 * 2 * 2 * 2)


class SheetTestModeTest(unittest.TestCase):
    def test_writes_sample_rows_to_ranks_test(self):
        import argparse
        import os
        from unittest import mock
        cfg = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
        args = argparse.Namespace(dry_run=True, limit=1, no_yahoo=False, sheet_test=True)
        env = {"GOOGLE_SA_JSON": "{}", "SHEET_ID": "dummy"}
        with mock.patch.dict(os.environ, env), mock.patch("tracker.sheets.append_rows") as ap, \
                mock.patch.object(run, "ROOT", ROOT):
            rows, errors = run.run(cfg, args, log=lambda *_: None)
        ap.assert_called_once()
        self.assertEqual(ap.call_args.kwargs["worksheet"], "ranks_test")
        self.assertEqual(len(rows), 4 * 2 + 1)  # google・bing × 4 地点 + Yahoo! JAPAN 1
        self.assertEqual(errors, [])
        self.assertTrue(all(r["status"] == "dry-run" for r in rows))


if __name__ == "__main__":
    unittest.main()
