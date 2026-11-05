"""検索順位の日次計測（japangreen.vn）。

使い方:
  python run.py                 # 本番（環境変数のキーで API を呼び、CSV とスプレッドシートに追記）
  python run.py --dry-run       # API を呼ばず tests/fixtures の見本応答で全工程を通す
  python run.py --limit 2       # キーワードを各言語 2 語に絞る（費用を抑えた試運転）
  python run.py --no-yahoo      # Yahoo! JAPAN（Apify）を外す
  python run.py --sheet-test    # API を呼ばず見本の行をスプレッドシートの ranks_test シートに書く
                                #（SHEET_WEBAPP_URL と SHEET_TOKEN だけで、書き込みの準備を確かめる）

環境変数: DFS_LOGIN, DFS_PASSWORD, APIFY_TOKEN, SHEET_WEBAPP_URL, SHEET_TOKEN
（GitHub Actions の Secrets から渡す。キーをファイルに書かない）
"""
import argparse
import csv
import datetime as dt
import json
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
from tracker.parse import parse_dataforseo, parse_yahoo_jp  # noqa: E402
from tracker.sheets import HEADER  # noqa: E402

JST = dt.timezone(dt.timedelta(hours=9))


def build_jobs(cfg, limit=None, yahoo=True):
    """計測する組み合わせ（engine, locale, keyword）の一覧。"""
    jobs = []
    for loc in cfg["locales"]:
        kws = cfg["keywords"][loc["keyword_set"]]
        if limit:
            kws = kws[:limit]
        for eng in loc["engines"]:
            if eng == "yahoo_jp" and not yahoo:
                continue
            for k in kws:
                jobs.append({"engine": eng, "locale": loc, "kw": k["kw"], "group": k["group"]})
    return jobs


def tag_of(job):
    return json.dumps([job["engine"], job["locale"]["id"], job["kw"]], ensure_ascii=False)


def row_of(date, job, parsed, status):
    loc = job["locale"]
    return {
        "date": date, "engine": job["engine"], "locale": loc["id"], "locale_label": loc["label"],
        "keyword": job["kw"], "group": job["group"], "rank": parsed.get("rank", ""),
        "url": parsed.get("url", ""), "local_rank": parsed.get("local_rank", ""),
        "has_local_pack": "TRUE" if parsed.get("has_local_pack") else "FALSE",
        "top10": " ".join(parsed.get("top10") or []), "results_count": parsed.get("results_count", ""),
        "status": status, "check_url": parsed.get("check_url", ""),
    }


def run(cfg, args, log=print):
    date = dt.datetime.now(JST).strftime("%Y-%m-%d")
    jobs = build_jobs(cfg, args.limit, yahoo=not args.no_yahoo)
    dfs_jobs = [j for j in jobs if j["engine"] in ("google", "bing")]
    yj_jobs = [j for j in jobs if j["engine"] == "yahoo_jp"]
    log(f"{date} 計測 {len(jobs)} 件（DataForSEO {len(dfs_jobs)}・Yahoo! JAPAN {len(yj_jobs)}）")
    rows, cost = [], 0.0

    # --- Google / Bing（DataForSEO） ---
    if args.dry_run:
        fx = json.loads((ROOT / "tests/fixtures/dataforseo_task.json").read_text(encoding="utf-8"))
        for j in dfs_jobs:
            rows.append(row_of(date, j, parse_dataforseo(fx, cfg["site_domain"], cfg["brand_names"]), "dry-run"))
    elif dfs_jobs:
        from tracker.clients import DataForSEO
        d = DataForSEO(os.environ["DFS_LOGIN"], os.environ["DFS_PASSWORD"])
        by_tag = {tag_of(j): j for j in dfs_jobs}
        ids = {}
        for eng in ("google", "bing"):
            tasks = [{"keyword": j["kw"], "location_code": j["locale"]["location_code"],
                      "language_code": j["locale"]["language_code"], "device": cfg["device"],
                      "depth": cfg["depth"], "tag": tag_of(j)} for j in dfs_jobs if j["engine"] == eng]
            if not tasks:
                continue
            for tag, tid, c in d.post_tasks(eng, tasks):
                ids[tid] = (eng, tag)
                cost += float(c or 0)
        log(f"  DataForSEO に {len(ids)} 件を登録（費用 ${cost:.4f}）")
        got = d.wait_all({tid: e for tid, (e, _) in ids.items()}, cfg["poll"]["interval_sec"], cfg["poll"]["timeout_min"], log)
        for tid, (eng, tag) in ids.items():
            t = got.get(tid, {})
            st = "ok" if t.get("status_code") == 20000 else f"error {t.get('status_code')} {t.get('status_message')}"
            parsed = parse_dataforseo(t, cfg["site_domain"], cfg["brand_names"]) if st == "ok" else {}
            rows.append(row_of(date, by_tag[tag], parsed, st))

    # --- Yahoo! JAPAN（Apify） ---
    if yj_jobs:
        kws = [j["kw"] for j in yj_jobs]
        try:
            if args.dry_run:
                items = json.loads((ROOT / "tests/fixtures/apify_yahoo_jp.json").read_text(encoding="utf-8"))
            else:
                from tracker.clients import apify_yahoo_jp
                items = apify_yahoo_jp(os.environ["APIFY_TOKEN"], cfg["yahoo_jp"]["actor"], kws, cfg["yahoo_jp"]["max_items"])
            st = "dry-run" if args.dry_run else "ok"
            for j in yj_jobs:
                p = parse_yahoo_jp(items, j["kw"], cfg["site_domain"])
                rows.append(row_of(date, j, p, st if p["top10"] else "no-results"))
        except Exception as e:  # Yahoo! JAPAN が失敗しても Google/Bing は保存する
            log(f"  Yahoo! JAPAN 失敗: {e}")
            for j in yj_jobs:
                rows.append(row_of(date, j, {}, f"error {type(e).__name__}"))

    # --- 保存 ---
    out = ROOT / ("data/ranks_dryrun.csv" if args.dry_run else "data/ranks.csv")
    new = not out.exists()
    out.parent.mkdir(exist_ok=True)
    with out.open("a", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=HEADER)
        if new:
            w.writeheader()
        w.writerows(rows)
    log(f"  CSV に {len(rows)} 行を追記: {out.relative_to(ROOT)}")
    if getattr(args, "sheet_test", False):
        from tracker.sheets import append_rows
        append_rows(os.environ["SHEET_WEBAPP_URL"], os.environ["SHEET_TOKEN"], rows, worksheet="ranks_test")
        log(f"  スプレッドシートの ranks_test シートに見本の {len(rows)} 行を追記（書き込みテスト）")
    elif not args.dry_run and os.environ.get("SHEET_WEBAPP_URL") and os.environ.get("SHEET_TOKEN"):
        from tracker.sheets import append_rows
        append_rows(os.environ["SHEET_WEBAPP_URL"], os.environ["SHEET_TOKEN"], rows)
        log("  スプレッドシートに追記")
    errors = [r for r in rows if r["status"].startswith(("error", "timeout"))]
    own = [r for r in rows if r["rank"] not in ("", None)]
    log(f"  完了: 当院が {len(own)}/{len(rows)} 件で 30 位以内・エラー {len(errors)} 件")
    return rows, errors


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--limit", type=int)
    ap.add_argument("--no-yahoo", action="store_true")
    ap.add_argument("--sheet-test", action="store_true")
    args = ap.parse_args()
    if args.sheet_test:  # 見本応答で 1 語ずつだけ（API は呼ばない）
        args.dry_run = True
        args.limit = args.limit or 1
    cfg = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
    rows, errors = run(cfg, args)
    # 全件エラーのときだけ失敗扱い（GitHub Actions の通知メールが届く）
    sys.exit(1 if rows and len(errors) == len(rows) else 0)


if __name__ == "__main__":
    main()
