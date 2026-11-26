"""検索順位の日次計測（japangreen.vn）。

使い方:
  python run.py                 # 本番（環境変数のキーで API を呼び、CSV とスプレッドシートに追記）
  python run.py --dry-run       # API を呼ばず tests/fixtures の見本応答で全工程を通す
  python run.py --limit 2       # キーワードを各言語 2 語に絞る（費用を抑えた試運転）
  python run.py --no-yahoo      # Yahoo! JAPAN（Apify）を外す
  python run.py --recover       # 登録済みで未回収の今日の Google・Bing の結果を回収（新たな課金なし）
  python run.py --to-test --limit 5  # 本番と同じ API で少数を測り、ranks_test シートと data/ranks_test.csv に書く（試運転）
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
from tracker.parse import parse_dataforseo, parse_yahoo_jp, serp_matches  # noqa: E402
from tracker.sheets import HEADER, FEATURE_HEADER  # noqa: E402

ICT = dt.timezone(dt.timedelta(hours=7))  # 記録の日付はベトナム時間（2026-11-05 から）
SECRET_ENVS = ("DFS_LOGIN", "DFS_PASSWORD", "APIFY_TOKEN", "SHEET_TOKEN", "SHEET_WEBAPP_URL")


def env(name):
    """Secrets を読む。貼り付けで入りがちな前後の空白・改行を取り除く（2026-11-05: 改行入りの APIFY_TOKEN で失敗）。"""
    return (os.environ.get(name) or "").strip()


def redact(text):
    """エラー文から Secrets の値と Apify トークンの形を伏せる（シート・CSV・ログに鍵を残さない）。"""
    import re
    s = str(text)
    for n in SECRET_ENVS:
        v = env(n)
        if len(v) >= 6:
            s = s.replace(v, "***")
    s = re.sub(r"apify_api_[A-Za-z0-9]+", "apify_api_***", s)
    s = re.sub(r"(Bearer|Basic)\s+[A-Za-z0-9+/=._-]+", r"\1 ***", s)
    return s


def build_jobs(cfg, limit=None, yahoo=True):
    """計測する組み合わせ（ブラウザ想定プロファイル × 地点・言語 × 検索語）の一覧。"""
    jobs = []
    for loc in cfg["locales"]:
        kws = cfg["keywords"][loc["keyword_set"]]
        if limit:
            kws = kws[:limit]
        for pid in loc["profiles"]:
            prof = cfg["profiles"][pid]
            if prof["engine"] == "yahoo_jp" and not yahoo:
                continue
            for k in kws:
                jobs.append({"engine": prof["engine"], "profile": pid, "device": prof["device"], "os": prof["os"],
                             "locale": loc, "kw": k["kw"], "group": k["group"]})
    return jobs


def tag_of(job):
    return json.dumps([job["profile"], job["locale"]["id"], job["kw"]], ensure_ascii=False)


def row_of(date, job, parsed, status):
    loc = job["locale"]
    return {
        "date": date, "engine": job["engine"], "browser_profile": job["profile"], "device": job["device"],
        "os": job["os"], "locale": loc["id"], "locale_label": loc["label"],
        "keyword": job["kw"], "group": job["group"], "rank": parsed.get("rank", ""),
        "url": parsed.get("url", ""), "local_rank": parsed.get("local_rank", ""),
        "has_local_pack": "TRUE" if parsed.get("has_local_pack") else "FALSE",
        "top10": " ".join(parsed.get("top10") or []), "results_count": parsed.get("results_count", ""),
        "status": status, "check_url": parsed.get("check_url", ""),
        **{k: parsed.get(k, "") for k in HEADER if k in FEATURE_HEADER or k.startswith("rank_")},
    }


def save(rows, args, log=print):
    """CSV（data/ranks.csv）とスプレッドシートに追記する。"""
    to_test = getattr(args, "to_test", False)
    out = getattr(args, "out", None) or ROOT / ("data/ranks_dryrun.csv" if args.dry_run else
                                                "data/ranks_test.csv" if to_test else "data/ranks.csv")
    new = not out.exists()
    out.parent.mkdir(exist_ok=True)
    if not new:
        migrate_csv(out)
    with out.open("a", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=HEADER)
        if new:
            w.writeheader()
        w.writerows(rows)
    log(f"  CSV に {len(rows)} 行を追記: {out}")
    if getattr(args, "sheet_test", False) or (to_test and not args.dry_run):
        from tracker.sheets import append_rows
        append_rows(env("SHEET_WEBAPP_URL"), env("SHEET_TOKEN"), rows, worksheet="ranks_test")
        log(f"  スプレッドシートの ranks_test シートに {len(rows)} 行を追記（試運転）")
    elif not args.dry_run and env("SHEET_WEBAPP_URL") and env("SHEET_TOKEN"):
        from tracker.sheets import append_rows
        append_rows(env("SHEET_WEBAPP_URL"), env("SHEET_TOKEN"), rows)
        log("  スプレッドシートに追記")


def migrate_csv(path):
    """CSV の見出しが今の列より短い（列を足す前の）ときは、見出しを今の列にして書き直す。古い行の新しい列は空欄。"""
    with open(path, encoding="utf-8", newline="") as f:
        r = csv.reader(f)
        head = next(r, [])
        if head == HEADER:
            return False
        if head != HEADER[:len(head)]:
            raise RuntimeError(f"{path} の列の並びが想定と違う（先頭の一致しない列あり）。書き足さずに止める")
        body = list(r)
    tmp = Path(str(path) + ".tmp")
    with tmp.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(HEADER)
        w.writerows(row + [""] * (len(HEADER) - len(row)) for row in body)
    tmp.replace(path)
    return True


MISMATCH = "error serp-mismatch（DataForSEO が別の検索語の結果ページを返した）"


def mismatched(t, cfg, engine):
    """2026-11-13: 取得できた（20000）のに、結果ページが検索語に合っていない（serp_check の対象エンジンだけ見る）。"""
    sc = cfg.get("serp_check") or {}
    return (t.get("status_code") == 20000 and engine in sc.get("engines", [])
            and not serp_matches(t, sc.get("terms", []), sc.get("top", 5)))


def needs_retry(t, cfg, engine):
    """送り直す結果: 結果なし（40102）か、結果ページの不一致。"""
    return t.get("status_code") == 40102 or mismatched(t, cfg, engine)


def dfs_row_parts(t, cfg, engine):
    """DataForSEO の結果 1 件 → (status, parsed)。不一致は順位を入れず、確かめ用に上位のドメインと check_url だけ残す。"""
    if t.get("status_code") != 20000:
        return f"error {t.get('status_code')} {t.get('status_message')}", {}
    parsed = parse_dfs(t, cfg, engine)
    if mismatched(t, cfg, engine):
        return MISMATCH, {k: parsed.get(k, "") for k in ("top10", "results_count", "check_url")}
    return "ok", parsed


def parse_dfs(t, cfg, engine):
    return parse_dataforseo(t, cfg["site_domain"], cfg["brand_names"], cfg.get("competitors"), engine)


def utc_now():
    return dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def result_keyword(t):
    """task_get の結果側の検索語（result[0].keyword）。DataForSEO が返した結果ページの検索語"""
    res = (t.get("result") or [None])[0] or {}
    return res.get("keyword")


def same_keyword(a, b):
    norm = lambda x: " ".join(str(x or "").replace("+", " ").split()).lower()
    return norm(a) == norm(b)


def bing_evidence(date, sent, got, task_of, cfg, log=print):
    """2026-11-26: DataForSEO サポートの依頼（2026-11-26 の返信）で、取れなかった Bing のタスクの資料を残す。
    送り直し前のものも含め、結果なし（40102）・結果の不一致になった Bing のタスクごとに、
    タスク ID・依頼の中身・登録時刻・取得時刻・task_get の応答そのものを data/evidence/bing_failed_<日付>.json に書く
    （Actions の保存ファイルとして残す。リポジトリには入れない。鍵は含まない）。
    あわせて、取れた Bing の全部について、結果側の検索語（result.keyword）が依頼と一致するかをログに出す。"""
    records, n_ok, n_diff, n_diff_mis = [], 0, 0, 0
    fetched_at = utc_now()
    for tid, (eng, tag, attempt, posted_at) in sent.items():
        if eng != "bing":
            continue
        t = got.get(tid)
        if t is None:
            continue
        payload = task_of[tag]
        rk = result_keyword(t)
        if t.get("status_code") == 20000:
            n_ok += 1
            if rk is not None and not same_keyword(rk, payload["keyword"]):
                n_diff += 1
                n_diff_mis += mismatched(t, cfg, eng)
        if not needs_retry(t, cfg, eng):
            continue
        records.append({
            "task_id": tid, "tag": tag, "attempt": attempt,
            "kind": "40102" if t.get("status_code") == 40102 else "serp-mismatch",
            "posted_at_utc": posted_at, "fetched_by_utc": fetched_at,
            "fetched_times": 1,  # この仕組みは 1 つのタスクを 1 回だけ取りに行く
            "task_post_payload": payload,
            "result_keyword": rk, "result_keyword_matches": None if rk is None else same_keyword(rk, payload["keyword"]),
            "task_get_advanced_response": t,
        })
    log(f"  Bing の結果側の検索語（result.keyword）: 取れた {n_ok} 件のうち依頼と違う {n_diff} 件（うち結果の不一致 {n_diff_mis} 件）")
    if records:
        out = ROOT / "data" / "evidence" / f"bing_failed_{date}.json"
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(records, ensure_ascii=False, indent=1), encoding="utf-8")
        log(f"  Bing の取れなかったタスクの資料 {len(records)} 件を保存: {out.relative_to(ROOT)}")


def dfs_task(j, cfg):
    """DataForSEO に送る 1 件。2026-11-26: Google には resolve_poisoned_results を付ける（config の設定・鍵本さんの決定）。
    Google が自動の検索にだけ本来の結果を差し替えたページ（poisoned SERP）を返すため。費用は 1 件 +$0.0006（標準キュー）。"""
    t = {"keyword": j["kw"], "location_code": j["locale"]["location_code"],
         "language_code": j["locale"]["language_code"], "device": j["device"], "os": j["os"],
         "depth": cfg["depth"], "tag": tag_of(j)}
    if j["engine"] == "google" and cfg.get("resolve_poisoned_results"):
        t["resolve_poisoned_results"] = True
    return t


def run(cfg, args, log=print):
    date = dt.datetime.now(ICT).strftime("%Y-%m-%d")
    jobs = build_jobs(cfg, args.limit, yahoo=not args.no_yahoo)
    if getattr(args, "only_yahoo", False):
        jobs = [j for j in jobs if j["engine"] == "yahoo_jp"]
    dfs_jobs = [j for j in jobs if j["engine"] in ("google", "bing")]
    yj_jobs = [j for j in jobs if j["engine"] == "yahoo_jp"]
    log(f"{date} 計測 {len(jobs)} 件（DataForSEO {len(dfs_jobs)}・Yahoo! JAPAN {len(yj_jobs)}）")
    rows, cost = [], 0.0

    # --- Google / Bing（DataForSEO） ---
    if args.dry_run:
        fx = json.loads((ROOT / "tests/fixtures/dataforseo_task.json").read_text(encoding="utf-8"))
        for j in dfs_jobs:
            rows.append(row_of(date, j, parse_dfs(fx, cfg, j["engine"]), "dry-run"))
    elif dfs_jobs:
        from tracker.clients import DataForSEO
        by_tag = {tag_of(j): j for j in dfs_jobs}
        task_of = {tag_of(j): dfs_task(j, cfg) for j in dfs_jobs}  # 2026-11-12: AI 概要は取得しない（鍵本さんの決定）
        ids = {}
        sent = {}  # 2026-11-26: task_id → (engine, tag, 何回目, 登録時刻 UTC)。DataForSEO への報告の資料用
        d = DataForSEO(env("DFS_LOGIN"), env("DFS_PASSWORD"))
        # 2026-11-10: Bing の登録が 402 で断られたとき、登録済み（課金済み）の Google まで捨てていた。
        # エンジンごとに失敗を閉じ込め、登録できた分は必ず結果を取りに行く。
        for eng in ("google", "bing"):
            eng_jobs = [j for j in dfs_jobs if j["engine"] == eng]
            if not eng_jobs:
                continue
            posted, failed = [], []
            try:
                d.post_tasks(eng, [task_of[tag_of(j)] for j in eng_jobs], posted, failed)
            except Exception as e:
                log(f"  DataForSEO {eng} の登録に失敗: {redact(e)}")
                done = {tag for tag, _, _ in posted} | {tag for tag, _ in failed}
                failed += [(tag_of(j), redact(e)) for j in eng_jobs if tag_of(j) not in done]
            for tag, tid, c in posted:
                ids[tid] = (eng, tag)
                sent[tid] = (eng, tag, 0, utc_now())
                cost += float(c or 0)
            for tag, msg in failed:
                rows.append(row_of(date, by_tag[tag], {}, f"error {msg}"[:200]))
        if ids:
            log(f"  DataForSEO に {len(ids)} 件を登録（費用 ${cost:.4f}）")
            got = {}
            try:
                got = d.wait_all({tid: e for tid, (e, _) in ids.items()}, cfg["poll"]["interval_sec"], cfg["poll"]["timeout_min"], log)
                # Bing は同じ条件でも「No Search Results（40102）」が不規則に返る（2026-11-05 の診断）→ config の retry_no_results 回まで送り直す
                # （2026-11-26 から 1 回。それまでは 2 回）
                # 2026-11-13: 結果ページが検索語に合っていない組（serp_check）も同じく送り直す
                for attempt in range(cfg.get("retry_no_results", 2)):
                    retry = [(tid, eng, tag) for tid, (eng, tag) in ids.items() if needs_retry(got.get(tid, {}), cfg, eng)]
                    if not retry:
                        break
                    n_mis = sum(1 for tid, eng, _ in retry if mismatched(got.get(tid, {}), cfg, eng))
                    log(f"  送り直し {attempt + 1} 回目: {len(retry)} 件（結果なし 40102 {len(retry) - n_mis}・結果の不一致 {n_mis}）")
                    new = {}
                    for eng in ("google", "bing"):
                        ts = [task_of[tag] for _, e, tag in retry if e == eng]
                        if ts:
                            posted2 = []
                            try:
                                d.post_tasks(eng, ts, posted2, [])
                            except Exception as e:  # 送り直しの失敗は元の 40102 の結果を残す
                                log(f"  送り直しの登録に失敗（{eng}）: {redact(e)}")
                            for tag, tid2, c in posted2:
                                new[tid2] = (eng, tag)
                                sent[tid2] = (eng, tag, attempt + 1, utc_now())
                                cost += float(c or 0)
                    got2 = d.wait_all({tid: e for tid, (e, _) in new.items()}, cfg["poll"]["interval_sec"], cfg["poll"]["timeout_min"], log)
                    resent = {t for _, t in new.values()}
                    for tid, _, tag in retry:
                        if tag in resent:
                            ids.pop(tid)
                    ids.update(new)
                    got.update(got2)
            except Exception as e:  # 取得途中の失敗: 取れた分は残し、残りをエラーにする
                log(f"  DataForSEO の結果取得に失敗: {redact(e)}")
            log(f"  DataForSEO の費用（送り直しを含む）${cost:.4f}")
            for tid, (eng, tag) in ids.items():
                t = got.get(tid)
                if t is None:
                    rows.append(row_of(date, by_tag[tag], {}, f"error 結果未取得 task_id={tid}"))
                    continue
                st, parsed = dfs_row_parts(t, cfg, eng)
                rows.append(row_of(date, by_tag[tag], parsed, st))
            n_mis = sum(1 for r in rows if r["status"] == MISMATCH)
            if n_mis:
                log(f"  送り直しても結果が不一致のまま: {n_mis} 件（圏外に数えず error serp-mismatch で記録）")
            # 2026-11-26: DataForSEO への報告用に、最後まで取れなかった Bing のタスク ID をログに残す（ID は秘密情報ではない）
            bad = [(tid, tag) for tid, (eng, tag) in ids.items() if eng == "bing" and needs_retry(got.get(tid) or {}, cfg, eng)]
            if bad:
                log(f"  Bing の取れなかったタスク {len(bad)} 件（task_id | tag）:")
                for tid, tag in bad:
                    log(f"    {tid} | {tag}")
            bing_evidence(date, sent, got, task_of, cfg, log)

    # --- Yahoo! JAPAN（Apify） ---
    if yj_jobs:
        kws = [j["kw"] for j in yj_jobs]
        try:
            if args.dry_run:
                items = json.loads((ROOT / "tests/fixtures/apify_yahoo_jp.json").read_text(encoding="utf-8"))
            else:
                from tracker.clients import apify_yahoo_jp
                items = apify_yahoo_jp(env("APIFY_TOKEN"), cfg["yahoo_jp"]["actor"], kws, cfg["yahoo_jp"]["max_items"])
            st = "dry-run" if args.dry_run else "ok"
            for j in yj_jobs:
                p = parse_yahoo_jp(items, j["kw"], cfg["site_domain"], cfg.get("competitors"))
                rows.append(row_of(date, j, p, st if p["top10"] else "no-results"))
        except Exception as e:  # Yahoo! JAPAN が失敗しても Google/Bing は保存する
            log(f"  Yahoo! JAPAN 失敗: {redact(e)}")
            for j in yj_jobs:
                rows.append(row_of(date, j, {}, f"error {redact(e)}"[:200]))

    save(rows, args, log)
    errors = [r for r in rows if r["status"].startswith(("error", "timeout"))]
    own = [r for r in rows if r["rank"] not in ("", None)]
    log(f"  完了: 当院が {len(own)}/{len(rows)} 件で 30 位以内・エラー {len(errors)} 件")
    return rows, errors


def recover(cfg, args, log=print):
    """登録済み（課金済み）なのに取りに行けなかった今日の Google・Bing の結果を回収して記録する（2026-11-10 追加）。
    DataForSEO の tasks_ready（結果ができていて未回収のタスク・直近 3 日）から、今日（ベトナム時間）に登録した分を
    検索条件（tag）ごとに最新の 1 件だけ取り、task_get で結果を読む。新たな課金は無い。"""
    from tracker.clients import DataForSEO
    date = dt.datetime.now(ICT).strftime("%Y-%m-%d")
    by_tag = {tag_of(j): j for j in build_jobs(cfg, yahoo=False)}
    d = DataForSEO(env("DFS_LOGIN"), env("DFS_PASSWORD"))
    pick = {}  # tag -> (posted, id, engine)
    for eng in ("google", "bing"):
        ready = d.tasks_ready(eng)
        log(f"  {eng}: 未回収のタスク {len(ready)} 件")
        for r in ready:
            tag = r.get("tag")
            if tag not in by_tag or by_tag[tag]["engine"] != eng:
                continue
            try:
                posted = dt.datetime.strptime(r.get("date_posted", ""), "%Y-%m-%d %H:%M:%S %z")
            except ValueError:
                continue
            if posted.astimezone(ICT).strftime("%Y-%m-%d") != date:
                continue
            if tag not in pick or posted > pick[tag][0]:
                pick[tag] = (posted, r["id"], eng)
    log(f"{date} 回収 {len(pick)} 件（今日の登録分・検索条件ごとに最新 1 件）")
    if not pick:
        return [], []
    got = d.wait_all({tid: eng for _, tid, eng in pick.values()}, cfg["poll"]["interval_sec"], cfg["poll"]["timeout_min"], log)
    rows = []
    for tag, (_, tid, eng) in pick.items():
        t = got.get(tid, {})
        st, parsed = dfs_row_parts(t, cfg, eng)
        rows.append(row_of(date, by_tag[tag], parsed, st))
    save(rows, args, log)
    errors = [r for r in rows if r["status"].startswith(("error", "timeout"))]
    own = [r for r in rows if r["rank"] not in ("", None)]
    log(f"  回収完了: {len(rows)} 行（当院が 30 位以内 {len(own)} 件・エラー {len(errors)} 件）")
    return rows, errors


def diagnose_bing(cfg, log=print):
    """Bing の「No Search Results（40102）」の原因を切り分ける。パラメータを変えた少数のタスクを送り、結果を表で出す（記録はしない）。"""
    from tracker.clients import DataForSEO
    d = DataForSEO(env("DFS_LOGIN"), env("DFS_PASSWORD"))
    base = [("ジャパングリーンクリニック ハノイ", 2392, "ja"), ("ジャパングリーンクリニック", 2392, "ja"),
            ("Japan Green Clinic Hanoi", 2704, "en"), ("ハノイ 病院", 2704, "ja")]
    variants = [("今の設定", {"device": "desktop", "os": "windows", "depth": 30}),
                ("os なし", {"device": "desktop", "depth": 30}),
                ("depth 10", {"device": "desktop", "os": "windows", "depth": 10}),
                ("location_name", {"device": "desktop", "depth": 10, "use_name": True})]
    tasks, meta = [], {}
    for kw, loc, lang in base:
        for vname, v in variants:
            t = {"keyword": kw, "language_code": lang, "tag": json.dumps([vname, kw, loc], ensure_ascii=False)}
            t.update({k: val for k, val in v.items() if k != "use_name"})
            if v.get("use_name"):
                t["location_name"] = {2392: "Japan", 2704: "Vietnam"}[loc]
            else:
                t["location_code"] = loc
            tasks.append(t)
    ids = {tid: (tag, c) for tag, tid, c in d.post_tasks("bing", tasks)}
    got = d.wait_all({tid: "bing" for tid in ids}, 30, 30, log)
    log("variant | keyword | loc | status | items | organic | cost")
    for tid, (tag, c) in ids.items():
        vname, kw, loc = json.loads(tag)
        t = got.get(tid, {})
        res = (t.get("result") or [None])[0] or {}
        items = res.get("items") or []
        org = [i for i in items if i.get("type") == "organic"]
        log(f"{vname} | {kw} | {loc} | {t.get('status_code')} {t.get('status_message')} | {len(items)} | {len(org)} | {c}")


GROUP_DOMAINS = ("japangreen.vn", "japan-green.com.sg", "japangreen.co.uk", "greenclinic.com.cn")


def from_web_count(task):
    """resolve_poisoned_results で差し替えを戻した通常の検索結果の数（checks に from_web が付いた項目）。"""
    res = (task.get("result") or [None])[0] or {}
    n = 0
    for i in res.get("items") or []:
        if i.get("type") != "organic":
            continue
        ch = i.get("checks")
        if (isinstance(ch, dict) and ch.get("from_web")) or (isinstance(ch, (list, tuple)) and "from_web" in ch):
            n += 1
    return n


def morning_of(date, path=None):
    """その日の本番の Google の行 → {(profile, locale, keyword): row}（試運転の比べる相手）"""
    path = path or ROOT / "data/ranks.csv"
    out = {}
    if Path(path).exists():
        with open(path, encoding="utf-8", newline="") as f:
            for r in csv.DictReader(f):
                if r.get("date") == date and r.get("engine") == "google":
                    out[(r["browser_profile"], r["locale"], r["keyword"])] = r
    return out


def diagnose_poison(cfg, log=print):
    """2026-11-26: Google の指名語の異常（当院もグループのサイトも上位に無い）が、DataForSEO の言う poisoned SERP か確かめる。
    指名語だけを resolve_poisoned_results: true で取り、同じ日の朝の本番と比べて表で出す（記録はしない）。
    費用は 1 件 $0.0018（30 位まで）＋ $0.0006（突き合わせ）。指名語 12 組 × Google 2 本＝24 件で約 $0.06。"""
    from tracker.clients import DataForSEO
    date = dt.datetime.now(ICT).strftime("%Y-%m-%d")
    jobs = [j for j in build_jobs(cfg, None, yahoo=False) if j["engine"] == "google" and j["group"] == "brand"]
    tasks = [dict(dfs_task(j, cfg), resolve_poisoned_results=True) for j in jobs]
    by_tag = {tag_of(j): j for j in jobs}
    d = DataForSEO(env("DFS_LOGIN"), env("DFS_PASSWORD"))
    posted, failed = [], []
    d.post_tasks("google", tasks, posted, failed)
    for tag, why in failed:
        log(f"  登録できず: {tag} {redact(str(why))}")
    cost = sum(float(c or 0) for _, _, c in posted)
    log(f"{date} poisoned SERP の試し: {len(posted)} 件を登録（費用 ${cost:.4f}）")
    got = d.wait_all({tid: "google" for _, tid, _ in posted}, 30, 40, log)
    morning = morning_of(date)
    log("profile | locale | keyword | status | rank 朝→試し | 公式サイトが上位10件に 朝→試し | from_web | 上位3件（試し）")
    odd_m = odd_t = n = 0
    for tag, tid, _ in posted:
        j, t = by_tag[tag], got.get(tid, {})
        key = (j["profile"], j["locale"]["id"], j["kw"])
        m = morning.get(key, {})
        m_grp = any(g in (m.get("top10") or "") for g in GROUP_DOMAINS) if m else None
        if t.get("status_code") != 20000:
            log(f"{key[0]} | {key[1]} | {key[2]} | {t.get('status_code')} {t.get('status_message')} | - | - | - | -")
            continue
        p = parse_dfs(t, cfg, "google")
        t_grp = any(g in " ".join(p["top10"]) for g in GROUP_DOMAINS)
        n += 1
        odd_m += m_grp is False
        odd_t += not t_grp
        log(f"{key[0]} | {key[1]} | {key[2]} | ok | {m.get('rank') or '圏外'}→{p['rank'] or '圏外'} | "
            f"{'あり' if m_grp else ('なし' if m_grp is False else '?')}→{'あり' if t_grp else 'なし'} | "
            f"{from_web_count(t)} | {' '.join(p['top10'][:3])}")
    log(f"まとめ: 取れた {n} 組のうち、公式サイトが上位 10 件に無い組＝朝 {odd_m}・試し {odd_t}（費用 ${cost:.4f}）")


def rows_recorded(date, path=None):
    """data/ranks.csv に、その日付の status=ok の行が何件あるか（予備の定時実行で二重に測らないため）"""
    path = path or ROOT / "data/ranks.csv"
    if not Path(path).exists():
        return 0
    with open(path, encoding="utf-8", newline="") as f:
        return sum(1 for r in csv.DictReader(f) if r.get("date") == date and r.get("status") == "ok")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--limit", type=int)
    ap.add_argument("--no-yahoo", action="store_true")
    ap.add_argument("--sheet-test", action="store_true")
    ap.add_argument("--diagnose-bing", action="store_true")
    ap.add_argument("--only-yahoo", action="store_true")
    ap.add_argument("--diagnose-poison", action="store_true",
                    help="Google の指名語を resolve_poisoned_results 付きで取り、朝の本番と比べて表で出す（記録はしない・約 $0.06）")
    ap.add_argument("--to-test", action="store_true",
                    help="本番と同じ API で測り、ranks_test シートと data/ranks_test.csv に書く（本番の記録に混ぜない試運転）")
    ap.add_argument("--recover", action="store_true",
                    help="登録済みで未回収の今日の Google・Bing の結果を回収して記録する（新たな課金なし）")
    ap.add_argument("--skip-if-done", type=int, metavar="N",
                    help="今日（ベトナム時間）の取得済みの行が N 件以上あれば何もしない（予備の定時実行用）")
    args = ap.parse_args()
    if args.skip_if_done:
        today = dt.datetime.now(ICT).strftime("%Y-%m-%d")
        n = rows_recorded(today)
        if n >= args.skip_if_done:
            print(f"{today} は計測済み（{n} 件）なので何もしない")
            return
    if args.recover:
        rows, errors = recover(json.loads((ROOT / "config.json").read_text(encoding="utf-8")), args)
        sys.exit(1 if not rows or len(errors) == len(rows) else 0)
    if args.diagnose_bing:
        diagnose_bing(json.loads((ROOT / "config.json").read_text(encoding="utf-8")))
        return
    if args.diagnose_poison:
        diagnose_poison(json.loads((ROOT / "config.json").read_text(encoding="utf-8")))
        return
    if args.sheet_test:  # 見本応答で 1 語ずつだけ（API は呼ばない）
        args.dry_run = True
        args.limit = args.limit or 1
    cfg = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
    rows, errors = run(cfg, args)
    # 全件エラーのときだけ失敗扱い（GitHub Actions の通知メールが届く）
    sys.exit(1 if rows and len(errors) == len(rows) else 0)


if __name__ == "__main__":
    main()
