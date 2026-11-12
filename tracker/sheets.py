"""Google スプレッドシートへの追記（スプレッドシートの Apps Script ウェブアプリに POST する）。

受け口のスクリプトは apps-script/Code.gs。Google Cloud のサービスアカウントは使わない。
"""
import json
import re
import urllib.parse
import urllib.request

BASE_HEADER = ["date", "engine", "browser_profile", "device", "os", "locale", "locale_label", "keyword", "group", "rank",
               "url", "local_rank", "has_local_pack", "top10", "results_count", "status", "check_url"]
# 2026-11-12 追加（列は末尾に足すだけ。既存の列の順は変えない＝スプレッドシート・Looker Studio が位置で読むため）
# 同日、AI 概要（ai_overview・ai_overview_own）は取得しないことになり列ごと削除（鍵本さんの決定）。シートの R・S 列も削除済み
FEATURE_HEADER = ["featured_snippet_own", "paa_own"]


def header_of(cfg):
    """列の並び: 既存 17 列 → 枠 2 列 → 競合の順位（config の competitors の順に rank_<id>）。"""
    return BASE_HEADER + FEATURE_HEADER + ["rank_" + c["id"] for c in cfg.get("competitors", [])]


def _load_cfg():
    from pathlib import Path
    return json.loads((Path(__file__).resolve().parents[1] / "config.json").read_text(encoding="utf-8"))


HEADER = header_of(_load_cfg())


class SheetWriteError(RuntimeError):
    pass


def append_rows(webapp_url, token, rows, worksheet="ranks", timeout=120):
    """rows（dict のリスト）を worksheet に追記する。合言葉は URL ではなく本文で送る。"""
    body = json.dumps({"token": token, "worksheet": worksheet, "header": HEADER,
                       "rows": [{k: r.get(k, "") for k in HEADER} for r in rows]},
                      ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(webapp_url, data=body, method="POST",
                                 headers={"Content-Type": "application/json"})
    # Apps Script は 302 で結果の URL へ転送する。urllib は POST の 302 を GET で追う（本文は送り直さない）
    with urllib.request.urlopen(req, timeout=timeout) as res:
        text = res.read().decode("utf-8", "replace")
        final_host = urllib.parse.urlsplit(res.geturl()).netloc
    try:
        out = json.loads(text)
    except ValueError:
        hint = ("Google のログイン画面に転送された。SHEET_WEBAPP_URL が /exec で終わる公開版の URL か"
                "（/dev のテスト用ではないか）、アクセスできるユーザー＝全員か確認"
                if final_host.startswith("accounts.") else "公開設定を確認")
        m = re.search(r"<title>(.*?)</title>", text, re.S | re.I)
        shape = urllib.parse.urlsplit(webapp_url.strip())
        url_check = (f"host={shape.netloc} /exec終わり={shape.path.endswith('/exec')} "
                     f"/u/含む={'/u/' in shape.path} 前後の空白={webapp_url != webapp_url.strip()}")
        raise SheetWriteError(f"Apps Script の応答が JSON ではない（{hint}）。転送先: {final_host}・"
                              f"ページ名: {(m.group(1).strip() if m else '不明')[:80]}・URL の形: {url_check}")
    if not out.get("ok"):
        raise SheetWriteError(f"Apps Script がエラーを返した: {out.get('error')}")
    return out
