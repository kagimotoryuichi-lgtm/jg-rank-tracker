"""検索順位の Excel ダッシュボードを作る。

使い方: python -X utf8 dashboard/build_excel_dashboard.py <ranks.csv または .tsv> <出力.xlsx> [日付 YYYY-MM-DD]
- 日付を省くと、データの最新日を使う
- 同じ日・同じ（ブラウザ・地点・検索語）が複数あるときは、取得できた（ok）最後の行を使う。無ければ最後の行
- 2026-11-07 にデザインを刷新: 日付と数字のカード・前日比・主な変化・グラフ 2 つ・色分け表（矢印つき）・地図の枠の表
"""
import csv
import datetime as dt
import json
import sys
from collections import Counter, OrderedDict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.worksheet.pagebreak import Break
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
CFG = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))

FONT = "Meiryo UI"
NAVY, GREEN = "2F4078", "249835"
CARD_BG = "F3F6FA"
UP, DOWN = "1E8E3E", "D0342C"
BANDS = [("1〜3 位", "2E9E5B", "FFFFFF"), ("4〜10 位", "9ED8B2", "1C4A2E"), ("11〜30 位", "FFE38F", "5C4700")]
OUT_BG, OUT_FG = "EEF0F3", "9AA0A8"
FAIL_BG, FAIL_FG = "F9D9C9", "A0441C"
MISS_BG, MISS_FG = "FCE8D6", "B35A12"
LINE_COLORS = ["2F4078", "0F9D8A", "7B61C9", "E0533D"]  # profiles の順（Chrome PC・Edge PC・iPhone・Yahoo）

thin = Side(style="thin", color="D5D9E0")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

PROFILE_LABEL = {pid: p["label"] for pid, p in CFG["profiles"].items()}
LOCALE_LABEL = {l["id"]: l["label"] for l in CFG["locales"]}
YAHOO_NOTE_1105 = "・Yahoo! JAPAN は初回の全件計測で設定の不備（トークン）により失敗し、その後の確認で日本語 2 語だけ取得した。残りは「取得失敗」。翌日から全語を計測する。"
OUT, FAIL, NOTYET = "圏外", "取得失敗", "未計測"


def F(**kw):
    kw.setdefault("name", FONT)
    return Font(**kw)


def fill(c):
    return PatternFill("solid", fgColor=c)


H_FILL = fill(NAVY)


def load(path):
    text = Path(path).read_text(encoding="utf-8-sig")
    delim = "\t" if str(path).endswith(".tsv") else ","
    return list(csv.DictReader(text.splitlines(), delimiter=delim))


def pick_latest(rows, day):
    best = OrderedDict()
    for r in rows:
        if r["date"] != day:
            continue
        k = (r["browser_profile"], r["locale"], r["keyword"])
        if r["status"] == "ok" or k not in best or best[k]["status"] != "ok":
            best[k] = r
    return best


def style_header(ws, row, c1, c2):
    for c in range(c1, c2 + 1):
        x = ws.cell(row, c)
        x.fill, x.font, x.alignment, x.border = H_FILL, F(bold=True, color="FFFFFF", size=9), CENTER, BOX


def rank_of(r):
    """ok の行なら順位（int）か圏外（None）。ok でなければ False"""
    if not r or r["status"] != "ok":
        return False
    v = (r.get("rank") or "").strip()
    return int(float(v)) if v else None


def band_of(n):
    return 0 if n <= 3 else 1 if n <= 10 else 2


def combos_all():
    return [(p, l["id"], k["kw"]) for l in CFG["locales"] for p in l["profiles"]
            for k in CFG["keywords"][l["keyword_set"]]]


def metrics(best):
    m = Counter()
    for key in combos_all():
        m["total"] += 1
        r = best.get(key)
        n = rank_of(r)
        if n is False:
            continue
        m["got"] += 1
        if n is not None:
            m["top30"] += 1
            m["top10"] += n <= 10
            m["first"] += n == 1
        if (r.get("local_rank") or "").strip():
            m["local"] += 1
    return m


def short(loc, p):
    return f"{LOCALE_LABEL[loc]}・{PROFILE_LABEL[p]}"


def movers(best, prev):
    """前日から順位が変わった組。(上がった, 下がった) を変化の大きい順に"""
    up, down = [], []
    for key in combos_all():
        a, b = rank_of(prev.get(key)), rank_of(best.get(key))
        if a is False or b is False or a == b:
            continue
        a2, b2 = (a or 31), (b or 31)  # 圏外は 31 位とみなして差を出す
        (up if b2 < a2 else down).append((abs(a2 - b2), key, a, b))
    return sorted(up, key=lambda x: -x[0]), sorted(down, key=lambda x: -x[0])


def rank_text(n):
    return OUT if n is None else f"{n} 位"


def card(ws, row, c1, label, value, unit, delta, sub=""):
    """数字のカード（4 行 × 2 列）。delta は前日比（None なら sub を表示）"""
    c2 = c1 + 1
    for r in range(row, row + 4):
        for c in (c1, c2):
            ws.cell(r, c).fill = fill(CARD_BG)
    for c in (c1, c2):
        ws.cell(row, c).fill = fill(GREEN)  # 上の色帯
    ws.merge_cells(start_row=row + 1, start_column=c1, end_row=row + 1, end_column=c2)
    x = ws.cell(row + 1, c1, label)
    x.font, x.alignment = F(size=9, color="555B66", bold=True), CENTER
    ws.merge_cells(start_row=row + 2, start_column=c1, end_row=row + 2, end_column=c2)
    x = ws.cell(row + 2, c1)
    x.value = CellRichText([TextBlock(InlineFont(rFont=FONT, sz=26, b=True, color=NAVY), str(value)),
                            TextBlock(InlineFont(rFont=FONT, sz=10, color="555B66"), f" {unit}")])
    x.alignment = CENTER
    ws.merge_cells(start_row=row + 3, start_column=c1, end_row=row + 3, end_column=c2)
    x = ws.cell(row + 3, c1)
    if delta is None:
        x.value, x.font = sub or "前日のデータなし", F(size=9, color="8A8F98")
    else:
        mark = "▲" if delta > 0 else "▼" if delta < 0 else "±"
        col = UP if delta > 0 else DOWN if delta < 0 else "8A8F98"
        x.value = f"前日比 {mark}{abs(delta)}" + (f"（{sub}）" if sub else "")
        x.font = F(size=9, bold=True, color=col)
    x.alignment = CENTER


def legend(ws, row, col, items):
    """色見本を並べる: items = [(文字, 背景, 文字色), ...]"""
    for i, (t, bg, fg) in enumerate(items):
        x = ws.cell(row, col + i, t)
        x.fill, x.font, x.alignment, x.border = fill(bg), F(size=8, color=fg, bold=True), CENTER, BOX


def paint_rank(x, rec, prv):
    n = rank_of(rec)
    if n is False:
        x.value = FAIL if rec else NOTYET
        x.fill, x.font = fill(FAIL_BG), F(size=8, color=FAIL_FG)
        return
    if n is None:
        bg, fg, txt = OUT_BG, OUT_FG, OUT
    else:
        bg, fg = BANDS[band_of(n)][1:]
        txt = str(n)
    a = rank_of(prv)
    tail = ""
    if a is not False:
        if n is not None and a is None:
            tail = " NEW"
        elif n is None and a is not None:
            tail = " ▼"
        elif n is not None and a is not None and n != a:
            tail = f" ▲{a - n}" if n < a else f" ▼{n - a}"
    x.fill = fill(bg)
    size = 10 if n else 8
    if not tail:
        x.value, x.font = txt, F(size=size, bold=bool(n), color=fg)
        return
    dark = n is not None and band_of(n) == 0
    tc = "FFFFFF" if dark else (UP if ("▲" in tail or "NEW" in tail) else DOWN)
    x.value = CellRichText([TextBlock(InlineFont(rFont=FONT, sz=size, b=bool(n), color=fg), txt),
                            TextBlock(InlineFont(rFont=FONT, sz=8, b=True, color=tc), tail)])


def paint_local(x, p, rec):
    if CFG["profiles"][p]["engine"] == "yahoo_jp":
        x.value, x.font = "対象外", F(size=8, color="B0B4BB")
        return
    if not rec or rec["status"] != "ok":
        x.value, x.fill, x.font = FAIL if rec else NOTYET, fill(FAIL_BG), F(size=8, color=FAIL_FG)
        return
    lr = (rec.get("local_rank") or "").strip()
    if lr:
        n = int(float(lr))
        bg, fg = BANDS[band_of(n)][1:]
        x.value, x.fill, x.font = str(n), fill(bg), F(size=10, bold=True, color=fg)
    elif rec.get("has_local_pack") == "TRUE":
        x.value, x.fill, x.font = "当院なし", fill(MISS_BG), F(size=8, bold=True, color=MISS_FG)
    else:
        x.value, x.font = "—", F(size=8, color="C0C4CB")


def heat(ws, top, title, keywords, cols, best, prev, field="rank"):
    """検索語 × (地点, ブラウザ) の色分け表。戻り値は次に使える行"""
    ws.cell(top, 1, title).font = F(bold=True, size=11, color=NAVY)
    ws.cell(top + 1, 1, "検索語")
    ws.cell(top + 1, 2, "種類")
    ws.merge_cells(start_row=top + 1, start_column=1, end_row=top + 2, end_column=1)
    ws.merge_cells(start_row=top + 1, start_column=2, end_row=top + 2, end_column=2)
    c = 3
    for loc, profs in cols:
        ws.cell(top + 1, c, LOCALE_LABEL[loc])
        if len(profs) > 1:
            ws.merge_cells(start_row=top + 1, start_column=c, end_row=top + 1, end_column=c + len(profs) - 1)
        for i, p in enumerate(profs):
            ws.cell(top + 2, c + i, PROFILE_LABEL[p])
        c += len(profs)
    last = c - 1
    style_header(ws, top + 1, 1, last)
    style_header(ws, top + 2, 1, last)
    ws.row_dimensions[top + 2].height = 26
    r = top + 3
    for k in keywords:
        x = ws.cell(r, 1, k["kw"])
        x.font, x.alignment, x.border = F(size=9), LEFT, BOX
        x = ws.cell(r, 2, "ブランド" if k["group"] == "brand" else "一般")
        x.font, x.alignment, x.border = F(size=8, color="555B66"), CENTER, BOX
        c = 3
        for loc, profs in cols:
            for p in profs:
                x = ws.cell(r, c)
                x.alignment, x.border = CENTER, BOX
                key = (p, loc, k["kw"])
                if field == "local_rank":
                    paint_local(x, p, best.get(key))
                else:
                    paint_rank(x, best.get(key), prev.get(key))
                c += 1
        ws.row_dimensions[r].height = 17
        r += 1
    return r + 1


def build_dashboard(wb, rows, days, day, best, prev_day, prev):
    kw_ja, kw_en = CFG["keywords"]["ja"], CFG["keywords"]["en"]
    ws = wb.active
    ws.title = "ダッシュボード"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 8
    for c in range(3, 13):
        ws.column_dimensions[get_column_letter(c)].width = 11.5

    # --- 見出し ---
    for c in range(1, 13):
        ws.cell(1, c).fill = H_FILL
    ws.row_dimensions[1].height = 32
    x = ws.cell(1, 1, "  japangreen.vn 検索順位ダッシュボード")
    x.font, x.alignment = F(bold=True, size=16, color="FFFFFF"), Alignment(vertical="center")
    ws.merge_cells("A2:L2")
    x = ws.cell(2, 1, "当院（japangreen.vn）が通常の検索結果で何位に出たか（30 位まで）。"
                      "「組」＝検索語 × ブラウザ × 地点。ブラウザ＝Chrome（PC）・Edge（PC）・Safari（iPhone）・Yahoo! JAPAN（PC）")
    x.font, x.alignment = F(size=9, color="555B66"), Alignment(vertical="center")
    ws.row_dimensions[2].height = 20

    # --- 日付と数字のカード（4〜7 行目） ---
    cur, pm = metrics(best), (metrics(prev) if prev else None)
    for c in (1, 2):
        ws.cell(4, c).fill = fill(NAVY)
        for r in range(5, 8):
            ws.cell(r, c).fill = fill(CARD_BG)
    for r, txt, font in (
        (5, "計測日（ベトナム時間）", F(size=9, color="555B66", bold=True)),
        (6, f"{int(day[5:7])}/{int(day[8:])}（{'月火水木金土日'[dt.date.fromisoformat(day).weekday()]}）",
         F(size=24, bold=True, color=NAVY)),
        (7, f"{day[:4]} 年 ／ 比べる日 {prev_day}" if prev_day else f"{day[:4]} 年 ／ 初回", F(size=9, color="8A8F98")),
    ):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        x = ws.cell(r, 1, txt)
        x.font, x.alignment = font, CENTER
    ws.row_dimensions[4].height = 5
    ws.row_dimensions[6].height = 42

    def d(k):
        return None if pm is None else cur[k] - pm[k]
    tot = cur["total"]
    card(ws, 4, 3, "30 位以内に入った組", cur["top30"], f"/ {tot} 組", d("top30"))
    card(ws, 4, 5, "10 位以内（1 ページ目）", cur["top10"], "組", d("top10"))
    card(ws, 4, 7, "1 位", cur["first"], "組", d("first"))
    card(ws, 4, 9, "地図の枠に当院", cur["local"], "組", d("local"), sub="Google・Bing")
    rate = round(100 * cur["got"] / tot) if tot else 0
    card(ws, 4, 11, "取得できた割合", rate, "%", None, sub=f"{cur['got']} / {tot} 組（失敗 {tot - cur['got']}）")

    # --- 前日からの主な変化（9〜14 行目） ---
    ws.cell(9, 1, "前日からの主な変化").font = F(bold=True, size=12, color=NAVY)
    ws.cell(9, 2, "変化の大きい順に 4 件（圏外は 31 位とみなす）").font = F(size=8, color="8A8F98")
    up, down = movers(best, prev) if prev else ([], [])
    for j, (head, items, col, c0) in enumerate((("上がった", up, UP, 1), ("下がった", down, DOWN, 7))):
        ws.merge_cells(start_row=10, start_column=c0, end_row=10, end_column=c0 + 5)
        x = ws.cell(10, c0, f" {'▲' if j == 0 else '▼'} {head}（{len(items)} 組）")
        x.font, x.alignment = F(bold=True, size=10, color="FFFFFF"), Alignment(vertical="center")
        for c in range(c0, c0 + 6):
            ws.cell(10, c).fill = fill(col)
        for i in range(4):
            r = 11 + i
            ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c0 + 5)
            for c in range(c0, c0 + 6):
                ws.cell(r, c).fill = fill("FAFBFC")
            x = ws.cell(r, c0)
            if i < len(items):
                _, (p, loc, kw), a, b = items[i]
                x.value = CellRichText([
                    TextBlock(InlineFont(rFont=FONT, sz=9, b=True, color="222222"), kw),
                    TextBlock(InlineFont(rFont=FONT, sz=8, color="8A8F98"), f"　{short(loc, p)}　"),
                    TextBlock(InlineFont(rFont=FONT, sz=9, b=True, color=col), f"{rank_text(a)} → {rank_text(b)}")])
            elif i == 0:
                x.value, x.font = ("前日のデータなし" if not prev else "なし"), F(size=9, color="8A8F98")
            x.alignment = Alignment(vertical="center", indent=1)
            ws.row_dimensions[r].height = 18

    # --- グラフの元データ（別シート。非表示の列はグラフに出ないため） ---
    g = wb.create_sheet("グラフ用")
    g["A1"] = "ダッシュボードのグラフの元データ（自動で作る。編集不要）"
    g.append([])
    g.append(["地点・ブラウザ"] + [b[0] for b in BANDS] + [OUT, FAIL])
    band_head = g.max_row
    for l in CFG["locales"]:
        for p in l["profiles"]:
            cnt = Counter()
            for k in CFG["keywords"][l["keyword_set"]]:
                n = rank_of(best.get((p, l["id"], k["kw"])))
                cnt["fail" if n is False else "out" if n is None else band_of(n)] += 1
            g.append([short(l["id"], p), cnt[0], cnt[1], cnt[2], cnt["out"], cnt["fail"]])
    band_last = g.max_row
    g.append([])
    g.append(["日付"] + [PROFILE_LABEL[p] for p in CFG["profiles"]])
    trend_head = g.max_row
    for dd in [x for x in days if x <= day][-30:]:
        b = pick_latest(rows, dd)
        g.append([f"{int(dd[5:7])}/{int(dd[8:])}"] +
                 [sum(1 for key in combos_all() if key[0] == p and rank_of(b.get(key)) not in (False, None))
                  for p in CFG["profiles"]])
    trend_last = g.max_row
    g.column_dimensions["A"].width = 32

    ch = BarChart()
    ch.type, ch.grouping, ch.overlap = "bar", "stacked", 100
    ch.height, ch.width = 11.5, 15.8  # A〜F 列の幅
    ch.add_data(Reference(g, min_col=2, max_col=5, min_row=band_head, max_row=band_last), titles_from_data=True)
    ch.set_categories(Reference(g, min_col=1, min_row=band_head + 1, max_row=band_last))
    for s_, colr in zip(ch.series, [BANDS[0][1], BANDS[1][1], BANDS[2][1], "D5D9E0"]):
        s_.graphicalProperties.solidFill = colr
        s_.graphicalProperties.line.solidFill = "FFFFFF"
    ch.x_axis.scaling.orientation = "maxMin"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.majorGridlines = None
    ch.y_axis.scaling.min = 0
    ch.y_axis.crosses = "max"  # 項目を上から並べても数値の軸は下に
    ch.legend = None  # 色の見方は表の上に出す
    ch.gapWidth = 40
    ws.cell(16, 1, "順位の内訳（検索語の数）").font = F(bold=True, size=11, color=NAVY)
    legend(ws, 16, 2, [(BANDS[0][0], BANDS[0][1], BANDS[0][2]), (BANDS[1][0], BANDS[1][1], BANDS[1][2]),
                       (BANDS[2][0], BANDS[2][1], BANDS[2][2]), (OUT, "D5D9E0", "555B66")])
    ws.add_chart(ch, "A17")

    lc = LineChart()
    lc.height, lc.width = 11.5, 11.2  # H〜L 列の幅
    lc.add_data(Reference(g, min_col=2, max_col=1 + len(CFG["profiles"]), min_row=trend_head, max_row=trend_last),
                titles_from_data=True)
    lc.set_categories(Reference(g, min_col=1, min_row=trend_head + 1, max_row=trend_last))
    for s_, colr in zip(lc.series, LINE_COLORS):
        s_.graphicalProperties.line.solidFill = colr
        s_.graphicalProperties.line.width = 28000
        s_.marker.symbol = "circle"
        s_.marker.size = 7
        s_.marker.graphicalProperties.solidFill = colr
        s_.marker.graphicalProperties.line.solidFill = colr
        s_.smooth = False
    lc.x_axis.delete = False
    lc.y_axis.delete = False
    lc.y_axis.scaling.min = 0
    lc.legend.position = "t"
    ws.cell(16, 8, "30 位以内に入った組数の推移（ブラウザ別）").font = F(bold=True, size=11, color=NAVY)
    ws.add_chart(lc, "H17")
    note = "※ ベトナムは 11/6 からハノイ市で計測（11/5 は国単位）。11/5 の Yahoo! JAPAN は設定の不備で 2 語のみ"
    ws.cell(44, 1, note).font = F(size=8, color="8A8F98")

    # --- 色分け表 ---
    nxt = 46
    ws.cell(nxt, 1, "色の見方（表の中の ▲▼ は前日からの変化）").font = F(bold=True, size=9, color="555B66")
    legend(ws, nxt, 3, [(BANDS[0][0], BANDS[0][1], BANDS[0][2]), (BANDS[1][0], BANDS[1][1], BANDS[1][2]),
                        (BANDS[2][0], BANDS[2][1], BANDS[2][2]), (OUT, OUT_BG, OUT_FG), (FAIL, FAIL_BG, FAIL_FG),
                        ("▲ 上がった", "FFFFFF", UP), ("▼ 下がった", "FFFFFF", DOWN), ("NEW 圏外から", "FFFFFF", UP)])
    nxt += 2
    ja_cols = [(l["id"], l["profiles"]) for l in CFG["locales"] if l["keyword_set"] == "ja"]
    en_cols = [(l["id"], l["profiles"]) for l in CFG["locales"] if l["keyword_set"] == "en"]
    nxt = heat(ws, nxt, "日本語の検索語 — 通常の検索結果の順位", kw_ja, ja_cols, best, prev)
    ws.row_breaks.append(Break(id=nxt - 1))  # 英語の表からは 2 ページ目
    nxt = heat(ws, nxt, "英語の検索語 — 通常の検索結果の順位", kw_en, en_cols, best, prev)

    ws.cell(nxt, 1, "地図の枠（検索結果に出る Google マップ・Bing の地図の枠）").font = F(bold=True, size=12, color=NAVY)
    nxt += 1
    ws.cell(nxt, 1, "色の見方").font = F(bold=True, size=9, color="555B66")
    legend(ws, nxt, 3, [("数字＝位置", BANDS[0][1], BANDS[0][2]), ("当院なし", MISS_BG, MISS_FG),
                        ("— 枠が出ない", "FFFFFF", "A0A4AB")])
    ws.merge_cells(start_row=nxt + 1, start_column=1, end_row=nxt + 1, end_column=12)
    ws.cell(nxt + 1, 1, "「当院なし」＝地図の枠は出たが当院が入っていない語。Google ビジネスプロフィールの改善（MEO）で狙える").font = \
        F(size=8, bold=True, color=MISS_FG)
    nxt = heat(ws, nxt + 3, "日本語の検索語 — 地図の枠での当院の位置", kw_ja, ja_cols, best, prev, field="local_rank")
    nxt = heat(ws, nxt, "英語の検索語 — 地図の枠での当院の位置", kw_en, en_cols, best, prev, field="local_rank")

    ws.print_area = f"A1:L{nxt}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.print_options.horizontalCentered = True


def main(src, out, day=None):
    rows = load(src)
    days = sorted({r["date"] for r in rows})
    day = day or days[-1]
    best = pick_latest(rows, day)
    before = [x for x in days if x < day]
    prev_day = before[-1] if before else None
    prev = pick_latest(rows, prev_day) if prev_day else {}
    kw_ja, kw_en = CFG["keywords"]["ja"], CFG["keywords"]["en"]

    wb = Workbook()
    build_dashboard(wb, rows, days, day, best, prev_day, prev)

    # --- 1 位のサイト ---
    t = wb.create_sheet("1位のサイト", 1)
    t.sheet_view.showGridLines = False
    t.column_dimensions["A"].width = 34
    t["A1"] = f"各検索語で 1 位に出たサイト（{day}）— 競合・まとめ記事の把握用"
    t["A1"].font = F(bold=True, size=12, color=NAVY)
    combos = [(l["id"], p) for l in CFG["locales"] for p in l["profiles"]]
    t.cell(3, 1, "検索語")
    for i, (loc, p) in enumerate(combos, 2):
        t.cell(3, i, f"{LOCALE_LABEL[loc]}\n{PROFILE_LABEL[p]}")
        t.column_dimensions[get_column_letter(i)].width = 17
    style_header(t, 3, 1, len(combos) + 1)
    t.row_dimensions[3].height = 32
    rr = 4
    for kset, kws in (("ja", kw_ja), ("en", kw_en)):
        for k in kws:
            x = t.cell(rr, 1, k["kw"])
            x.border, x.font = BOX, F(size=9)
            for i, (loc, p) in enumerate(combos, 2):
                lset = next(l["keyword_set"] for l in CFG["locales"] if l["id"] == loc)
                x = t.cell(rr, i)
                x.border, x.alignment, x.font = BOX, CENTER, F(size=8)
                if lset != kset:
                    x.fill = fill("F4F4F4")
                    continue
                rec = best.get((p, loc, k["kw"]))
                if rec and rec["status"] == "ok":
                    top = (rec["top10"] or "").split(" ")[0].replace("www.", "")
                    x.value = top
                    if "japangreen.vn" in top:
                        x.font, x.fill = F(size=8, bold=True, color="FFFFFF"), fill(BANDS[0][1])
                else:
                    x.value = FAIL if rec else NOTYET
                    x.font, x.fill = F(size=8, color=FAIL_FG), fill(FAIL_BG)
            rr += 1
    t.freeze_panes = "B4"

    # --- データ ---
    d = wb.create_sheet("データ", 2)
    cols = ["date", "browser_profile", "engine", "device", "os", "locale", "keyword", "group", "rank", "url",
            "local_rank", "has_local_pack", "top10", "status", "check_url"]
    d.append(cols)
    style_header(d, 1, 1, len(cols))
    for rec in best.values():
        d.append([rec.get(c, "") for c in cols])
    for i, w in enumerate([11, 12, 9, 9, 9, 8, 32, 8, 6, 40, 8, 9, 60, 14, 40], 1):
        d.column_dimensions[get_column_letter(i)].width = w
    d.freeze_panes = "A2"
    d.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{d.max_row}"

    # --- 説明 ---
    n = wb.create_sheet("説明", 3)
    n.column_dimensions["A"].width = 120
    st = Counter((r["browser_profile"], r["status"] if r["status"] == "ok" else "失敗") for r in best.values())
    lines = [
        ("この表の見方", True),
        (f"・計測日 {day}。毎朝ベトナム時間 7 時台に自動で計測し、Google スプレッドシート「japangreen_検索順位」の ranks シートに追記している（このファイルはその抜粋）。", False),
        ("・順位は「通常の検索結果」だけを数えた当院（japangreen.vn）の順位。広告・地図の枠・AI の概要は数えない。30 位までに無ければ「圏外」。", False),
        ("・「組」は検索語 × ブラウザ × 地点の組み合わせ（1 日 173 組）。前日比は、その前にデータがある日と比べている。", False),
        ("・ブラウザは検索エンジン×端末で代表させている: Chrome（PC）＝Google パソコン、Edge（PC）＝Bing パソコン、Safari（iPhone）＝Google スマホ、Yahoo! JAPAN（PC）。Firefox・Mac の Safari は Chrome（PC）と同じ結果。", False),
        ("・地点は日本＝国単位、ベトナム＝ハノイ市（2026-11-06 から。それ以前のベトナムは国単位なので比べない）。"
         "ログインや閲覧履歴の影響は含まない（各自の画面とは多少ずれる）。", False),
        ("・地図の枠: 数字＝枠の中の当院の位置、「当院なし」＝枠は出たが当院がいない、「—」＝枠が出ない。Yahoo! JAPAN は対象外。", False),
        ("", False),
        ("この日の注意", True),
        ("・Bing は同じ条件でも「結果なし」が不規則に返る。2 回まで自動で送り直し、それでも取れなかったものは「取得失敗」。", False),
        ("・同じ日に複数回計測した組は、取得できた最後の結果を使っている。", False),
        ("", False),
        ("取得の状況（ブラウザ別）", True),
    ]
    if day == "2026-11-05":
        lines.insert(lines.index(("・同じ日に複数回計測した組は、取得できた最後の結果を使っている。", False)), (YAHOO_NOTE_1105, False))
    lines += [(f"・{PROFILE_LABEL[p]}: 取得 {st[(p, 'ok')]} 件／失敗 {st[(p, '失敗')]} 件", False) for p in CFG["profiles"]]
    for i, (txt, b) in enumerate(lines, 1):
        c = n.cell(i, 1, txt)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.font = F(bold=True, color=NAVY) if b else F(size=10)

    wb.active = 0
    wb.save(out)
    print("saved", out, "day", day, "rows", len(best))


if __name__ == "__main__":
    main(*sys.argv[1:])
