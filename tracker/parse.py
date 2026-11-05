"""検索結果（API の応答）から当院の順位などを取り出す。外部ライブラリを使わない純粋関数だけ置く。"""
from urllib.parse import urlparse

OUT_OF_RANGE = ""  # 圏外は空欄（スプレッドシート・Looker Studio で数値として扱いやすい）


def _host(url_or_domain):
    if not url_or_domain:
        return ""
    s = url_or_domain if "://" in url_or_domain else "http://" + url_or_domain
    return (urlparse(s).hostname or "").lower()


def is_own(url_or_domain, site_domain):
    h = _host(url_or_domain)
    d = site_domain.lower()
    return h == d or h.endswith("." + d)


def _brand_match(text, brand_names):
    t = (text or "").lower()
    return any(b.lower() in t for b in brand_names)


def parse_dataforseo(task, site_domain, brand_names):
    """DataForSEO の task_get/advanced の task 1 件 → 結果 dict。

    rank: 通常の検索結果（organic）だけで数えた順位（rank_group）。広告・地図枠は数えない
    local_rank: 地図枠（local_pack）の中の順位
    """
    res = (task.get("result") or [None])[0] or {}
    items = res.get("items") or []
    organic = [i for i in items if i.get("type") == "organic"]
    rank, url = OUT_OF_RANGE, ""
    for i in organic:
        if is_own(i.get("url") or i.get("domain"), site_domain):
            rank, url = i.get("rank_group"), i.get("url", "")
            break
    local_rank = OUT_OF_RANGE
    local_items = [i for i in items if i.get("type") in ("local_pack", "map")]
    for n, i in enumerate(local_items, 1):
        if is_own(i.get("url") or i.get("domain"), site_domain) or _brand_match(i.get("title"), brand_names):
            local_rank = i.get("rank_group") or n
            break
    top10 = [_host(i.get("url") or i.get("domain")) for i in organic[:10]]
    return {
        "rank": rank,
        "url": url,
        "local_rank": local_rank,
        "has_local_pack": bool(local_items),
        "top10": top10,
        "results_count": res.get("se_results_count", ""),
        "check_url": res.get("check_url", ""),
    }


def parse_yahoo_jp(items, keyword, site_domain):
    """Apify（search.yahoo.co.jp）の結果のうち、keyword に当たる行 → 結果 dict。"""
    rows = sorted((i for i in items if (i.get("query") or "") == keyword), key=lambda i: i.get("position") or 9999)
    rank, url = OUT_OF_RANGE, ""
    for i in rows:
        if is_own(i.get("url"), site_domain):
            rank, url = i.get("position"), i.get("url", "")
            break
    top10 = [_host(i.get("url")) for i in rows[:10]]
    return {
        "rank": rank,
        "url": url,
        "local_rank": OUT_OF_RANGE,
        "has_local_pack": False,
        "top10": top10,
        "results_count": "",
        "check_url": rows[0].get("searchUrl", "") if rows else "",
    }
