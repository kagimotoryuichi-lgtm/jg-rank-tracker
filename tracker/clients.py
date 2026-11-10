"""DataForSEO と Apify への問い合わせ（標準ライブラリの urllib だけで書く）。"""
import base64
import json
import time
import urllib.error
import urllib.parse
import urllib.request

DFS = "https://api.dataforseo.com/v3/serp"


def _req(url, method="GET", body=None, headers=None, timeout=120):
    data = json.dumps(body).encode("utf-8") if body is not None else None
    h = {"Content-Type": "application/json"}
    h.update(headers or {})
    r = urllib.request.Request(url, data=data, method=method, headers=h)
    try:
        with urllib.request.urlopen(r, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:  # 本文の status_message を添えて原因を分かりやすくする
        try:
            msg = json.loads(e.read().decode("utf-8", "replace")).get("status_message", "")
        except Exception:
            msg = ""
        hint = ""
        if e.code == 401:
            hint = ("（DFS_LOGIN／DFS_PASSWORD が API Access の値か確認）" if "dataforseo" in url
                    else "（APIFY_TOKEN が登録されているか確認）" if "apify" in url else "")
        raise RuntimeError(f"HTTP {e.code} {msg}{hint}".strip()) from None


class DataForSEO:
    def __init__(self, login, password):
        tok = base64.b64encode(f"{login}:{password}".encode()).decode()
        self.h = {"Authorization": "Basic " + tok}

    def post_tasks(self, engine, tasks, posted=None, failed=None):
        """tasks: [{keyword, location_code, language_code, device, depth, tag}]
        登録できた分を posted に (tag, id, cost)、タスク単位で断られた分を failed に (tag, 理由) で順に足す。
        通信や HTTP の失敗は例外にする（それまでに登録できた分は posted に残る＝課金済みの結果を取りに行ける）。"""
        posted = [] if posted is None else posted
        failed = [] if failed is None else failed
        for i in range(0, len(tasks), 100):  # 1 回 100 件まで
            j = _req(f"{DFS}/{engine}/organic/task_post", "POST", tasks[i:i + 100], self.h)
            if j.get("status_code") != 20000:
                raise RuntimeError(f"task_post {engine}: {j.get('status_code')} {j.get('status_message')}")
            for t in j.get("tasks", []):
                tag = (t.get("data") or {}).get("tag")
                if t.get("status_code") == 20100:
                    posted.append((tag, t["id"], t.get("cost", 0)))
                elif tag:
                    failed.append((tag, f"{t.get('status_code')} {t.get('status_message')}"))
        return posted

    def tasks_ready(self, engine):
        """結果ができていて、まだ取りに来ていないタスク（直近 3 日・最大 1000 件）。[{id, tag, date_posted, …}]"""
        j = _req(f"{DFS}/{engine}/organic/tasks_ready", "GET", None, self.h)
        out = []
        for t in j.get("tasks") or []:
            out += t.get("result") or []
        return out

    def get_task(self, engine, task_id):
        j = _req(f"{DFS}/{engine}/organic/task_get/advanced/{task_id}", "GET", None, self.h)
        t = (j.get("tasks") or [{}])[0]
        return t

    def wait_all(self, engine_ids, interval_sec, timeout_min, log=print):
        """engine_ids: {task_id: engine}。完了したものから task_get。戻り値 {task_id: task}"""
        done, deadline = {}, time.time() + timeout_min * 60
        pending = dict(engine_ids)
        while pending and time.time() < deadline:
            for tid, eng in list(pending.items()):
                t = self.get_task(eng, tid)
                code = t.get("status_code")
                if code == 20000:
                    done[tid] = t
                    pending.pop(tid)
                elif code in (40601, 40602, 20100):  # 処理中・キュー待ち
                    continue
                else:  # エラーは結果無しとして記録
                    done[tid] = t
                    pending.pop(tid)
            if pending:
                log(f"  待機中 {len(pending)} 件…")
                time.sleep(interval_sec)
        for tid in pending:
            done[tid] = {"status_code": "timeout", "status_message": "timeout", "result": None}
        return done


def apify_yahoo_jp(token, actor, keywords, max_items, timeout=900):
    """Apify のアクターを同期実行し、データセットの行（list）を返す。"""
    start = [{"url": "https://search.yahoo.co.jp/search?" + urllib.parse.urlencode({"p": k})} for k in keywords]
    url = f"https://api.apify.com/v2/acts/{actor}/run-sync-get-dataset-items?timeout={timeout}"
    return _req(url, "POST", {"startUrls": start, "maxItems": max_items},
                {"Authorization": "Bearer " + token}, timeout=timeout + 60)
