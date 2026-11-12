"""2026-11-10 の不具合の再現: Bing の登録が 402 で断られても、登録済み（課金済み）の Google の結果は記録する。"""
import datetime as dt
import json
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import run  # noqa: E402
import tracker.clients as clients  # noqa: E402

FX = json.loads((ROOT / "tests" / "fixtures" / "dataforseo_task.json").read_text(encoding="utf-8"))
CFG = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
CFG["poll"] = {"interval_sec": 0, "timeout_min": 1}


class FakeDFS:
    """Google は登録でき、Bing は 402 で断る。"""
    ready = []

    def __init__(self, login, password):
        self.n = 0

    def post_tasks(self, engine, tasks, posted=None, failed=None):
        posted = [] if posted is None else posted
        if engine == "bing":
            raise RuntimeError("HTTP 402 Payment Required.")
        for t in tasks:
            self.n += 1
            posted.append((t["tag"], f"{engine}-{self.n}", 0.0025))
        return posted

    def wait_all(self, engine_ids, interval_sec, timeout_min, log=print):
        return {tid: FX for tid in engine_ids}

    def tasks_ready(self, engine):
        return [r for r in self.ready if r["se"] == engine]


def args(out, **kw):
    a = dict(dry_run=False, limit=None, no_yahoo=True, only_yahoo=False, sheet_test=False, out=out)
    a.update(kw)
    return types.SimpleNamespace(**a)


class EngineIsolationTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.out = Path(self.tmp.name) / "ranks.csv"
        self.env = mock.patch.dict("os.environ", {"SHEET_WEBAPP_URL": "", "SHEET_TOKEN": ""})
        self.env.start()

    def tearDown(self):
        self.env.stop()
        self.tmp.cleanup()

    def test_bing_402_keeps_google(self):
        with mock.patch.object(clients, "DataForSEO", FakeDFS):
            rows, errors = run.run(CFG, args(self.out), log=lambda *_: None)
        google = [r for r in rows if r["engine"] == "google"]
        bing = [r for r in rows if r["engine"] == "bing"]
        self.assertEqual(len(google), 156)  # PC・iPhone・Android（2026-11-12）
        self.assertTrue(all(r["status"] == "ok" for r in google))
        self.assertEqual(len(bing), 52)
        self.assertTrue(all("402" in r["status"] for r in bing))
        self.assertEqual(len(errors), 52)
        with self.out.open(encoding="utf-8") as f:
            self.assertEqual(sum(1 for _ in f), 1 + 156 + 52)  # 見出し + Google 156 + Bing のエラー 52

    def test_recover_picks_todays_latest_per_tag(self):
        ict = dt.timezone(dt.timedelta(hours=7))
        now = dt.datetime.now(ict).astimezone(dt.timezone.utc)
        fmt = "%Y-%m-%d %H:%M:%S +00:00"
        jobs = [j for j in run.build_jobs(CFG, yahoo=False) if j["engine"] == "google"]
        ready = []
        for i, j in enumerate(jobs):
            tag = run.tag_of(j)
            ready.append({"id": f"old-{i}", "se": "google", "tag": tag, "date_posted": (now - dt.timedelta(minutes=20)).strftime(fmt)})
            ready.append({"id": f"new-{i}", "se": "google", "tag": tag, "date_posted": now.strftime(fmt)})
        ready.append({"id": "yesterday", "se": "google", "tag": run.tag_of(jobs[0]),
                      "date_posted": (now - dt.timedelta(days=1)).strftime(fmt)})
        ready.append({"id": "other", "se": "google", "tag": "not-ours", "date_posted": now.strftime(fmt)})
        seen = {}

        class F(FakeDFS):
            def wait_all(self, engine_ids, *a, **k):
                seen.update(engine_ids)
                return {tid: FX for tid in engine_ids}
        F.ready = ready
        with mock.patch.object(clients, "DataForSEO", F):
            rows, errors = run.recover(CFG, args(self.out), log=lambda *_: None)
        self.assertEqual(len(rows), 156)
        self.assertEqual(errors, [])
        self.assertTrue(all(t.startswith("new-") for t in seen))


if __name__ == "__main__":
    unittest.main()
