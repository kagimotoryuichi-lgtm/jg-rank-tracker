"""Google スプレッドシートへの追記（スプレッドシートの Apps Script ウェブアプリに POST する）。

受け口のスクリプトは apps-script/Code.gs。Google Cloud のサービスアカウントは使わない。
"""
import json
import urllib.request

HEADER = ["date", "engine", "locale", "locale_label", "keyword", "group", "rank", "url",
          "local_rank", "has_local_pack", "top10", "results_count", "status", "check_url"]


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
    try:
        out = json.loads(text)
    except ValueError:
        raise SheetWriteError(f"Apps Script の応答が JSON ではない（公開設定を確認）: {text[:200]}")
    if not out.get("ok"):
        raise SheetWriteError(f"Apps Script がエラーを返した: {out.get('error')}")
    return out
