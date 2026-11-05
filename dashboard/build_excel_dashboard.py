"""検索順位の Excel ダッシュボードを作る。

使い方: python -X utf8 dashboard/build_excel_dashboard.py <ranks.csv または .tsv> <出力.xlsx> [日付 YYYY-MM-DD]
- 日付を省くと、データの最新日を使う
- 同じ日・同じ（ブラウザ・地点・検索語）が複数あるときは、取得できた（ok）最後の行を使う。無ければ最後の行
"""
import csv
import json
import sys
from collections import Counter, OrderedDict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
CFG = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))

NAVY, GREEN = "2F4078", "249835"
H_FILL = PatternFill("solid", fgColor=NAVY)
H_FONT = Font(bold=True, color="FFFFFF")
SUB_FILL = PatternFill("solid", fgColor="E8ECF6")
thin = Side(style="thin", color="C8CCD6")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

PROFILE_LABEL = {pid: p["label"] for pid, p in CFG["profiles"].items()}
LOCALE_LABEL = {l["id"]: l["label"] for l in CFG["locales"]}
YAHOO_NOTE_1105 = "・Yahoo! JAPAN は初回の全件計測で設定の不備（トークン）により失敗し、その後の確認で日本語 2 語だけ取得した。残りは「取得失敗」。翌日から全語を計測する。"
OUT, FAIL, NOTYET = "圏外", "取得失敗", "未計測"


def load(path):
    text = Path(path).read_text(encoding="utf-8-sig")
    delim = "\t" if path.endswith(".tsv") else ","
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


def cell_value(r, field="rank"):
    if r is None:
        return NOTYET
    if r["status"] != "ok":
        return FAIL
    v = (r.get(field) or "").strip()
    return int(float(v)) if v else (OUT if field == "rank" else "")


def style_header(ws, row, c1, c2):
    for c in range(c1, c2 + 1):
        x = ws.cell(row, c)
        x.fill, x.font, x.alignment, x.border = H_FILL, H_FONT, CENTER, BOX


def _bg(c):
    return PatternFill(start_color=c, end_color=c, fill_type="solid")  # 条件付き書式は bgColor が必要


def add_rank_colors(ws, rng):
    ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["1", "3"],
                                  fill=_bg("2E9E5B"), font=Font(color="FFFFFF", bold=True)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["4", "10"],
                                  fill=_bg("BFE6CC")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=["11", "30"],
                                  fill=_bg("FFF1B8")))
    first = rng.split(":")[0]
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first}="{OUT}"'], fill=_bg("EEEEEE"),
                                                   font=Font(color="888888")))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'OR({first}="{FAIL}",{first}="{NOTYET}")'],
                                                   fill=_bg("FBE3D6"), font=Font(color="B4532A")))


def matrix(ws, top, title, keywords, cols, best, field="rank", note=""):
    """keywords × (地点, ブラウザ) の表。戻り値は次に使える行。"""
    ws.cell(top, 1, title).font = Font(bold=True, size=12, color=NAVY)
    if note:
        ws.cell(top, 3, note).font = Font(size=9, color="666666")
    # 2 段の見出し（地点 / ブラウザ）
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
    last_col = c - 1
    style_header(ws, top + 1, 1, last_col)
    style_header(ws, top + 2, 1, last_col)
    r = top + 3
    for k in keywords:
        ws.cell(r, 1, k["kw"]).alignment = LEFT
        ws.cell(r, 2, "ブランド" if k["group"] == "brand" else "一般").alignment = CENTER
        c = 3
        for loc, profs in cols:
            for p in profs:
                if field == "local_rank" and CFG["profiles"][p]["engine"] == "yahoo_jp":
                    x = ws.cell(r, c, "対象外")
                    x.font = Font(color="999999")
                else:
                    x = ws.cell(r, c, cell_value(best.get((p, loc, k["kw"])), field))
                x.alignment = CENTER
                c += 1
        for cc in range(1, last_col + 1):
            ws.cell(r, cc).border = BOX
        r += 1
    add_rank_colors(ws, f"C{top + 3}:{get_column_letter(last_col)}{r - 1}")
    return r + 1


def main(src, out, day=None):
    rows = load(src)
    day = day or max(r["date"] for r in rows)
    best = pick_latest(rows, day)
    kw_ja, kw_en = CFG["keywords"]["ja"], CFG["keywords"]["en"]
    loc_profiles = {l["id"]: l["profiles"] for l in CFG["locales"]}

    wb = Workbook()
    ws = wb.active
    ws.title = "ダッシュボード"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 19
    for c in range(3, 12):
        ws.column_dimensions[get_column_letter(c)].width = 13.5

    ws["A1"] = "japangreen.vn 検索順位ダッシュボード"
    ws["A1"].font = Font(bold=True, size=16, color=NAVY)
    ws["A2"] = (f"計測日: {day}（ベトナム時間）／ 通常の検索結果での当院（japangreen.vn）の順位・30 位まで ／ "
                "ブラウザ＝検索エンジン×端末（Chrome PC＝Google、Edge PC＝Bing、Safari iPhone＝Google スマホ、Yahoo! JAPAN PC）")
    ws["A2"].font = Font(size=9, color="555555")
    ws.merge_cells("A2:K2")

    # --- 要約（ブラウザ × 地点） ---
    ws["A4"] = "まとめ：30 位以内に入った検索語の数"
    ws["A4"].font = Font(bold=True, size=12, color=NAVY)
    heads = ["地点 × 言語", "ブラウザ", "計測した語", "取得できた語", "30 位以内", "10 位以内", "ブランド語で 1 位", "地図の枠に当院"]
    for i, h in enumerate(heads, 1):
        ws.cell(5, i, h)
    style_header(ws, 5, 1, len(heads))
    r = 6
    chart_rows = []
    for l in CFG["locales"]:
        kws = CFG["keywords"][l["keyword_set"]]
        for p in l["profiles"]:
            recs = [best.get((p, l["id"], k["kw"])) for k in kws]
            got = [x for x in recs if x and x["status"] == "ok"]
            ranks = [int(float(x["rank"])) for x in got if (x["rank"] or "").strip()]
            brand1 = sum(1 for x in got if x["group"] == "brand" and (x["rank"] or "").strip() and int(float(x["rank"])) == 1)
            local = sum(1 for x in got if (x.get("local_rank") or "").strip())
            vals = [l["label"], PROFILE_LABEL[p], len(kws), len(got), len(ranks), sum(1 for v in ranks if v <= 10), brand1, local]
            ws.cell(r, 26, f"{l['label']}・{PROFILE_LABEL[p]}")  # Z 列: グラフの項目名
            for i, v in enumerate(vals, 1):
                x = ws.cell(r, i, v)
                x.border, x.alignment = BOX, Alignment(horizontal="left" if i in (1, 2) else "center", vertical="center")
            if len(got) < len(kws):
                ws.cell(r, 4).fill = PatternFill("solid", fgColor="FBE3D6")
            chart_rows.append(r)
            r += 1
    sum_first, sum_last = 6, r - 1

    ws.column_dimensions["Z"].width = 30  # グラフの項目名
    ch = BarChart()
    ch.type = "bar"
    ch.title = "30 位以内に入った検索語の数（地点・ブラウザ別）"
    ch.height, ch.width = 9.5, 17
    ch.varyColors = False
    ch.add_data(Reference(ws, min_col=5, min_row=5, max_row=sum_last), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=26, min_row=6, max_row=sum_last))
    ch.series[0].graphicalProperties.solidFill = GREEN
    ch.series[0].graphicalProperties.line.solidFill = GREEN
    from openpyxl.chart.label import DataLabelList
    ch.dataLabels = DataLabelList()
    ch.dataLabels.showVal = True
    ch.dataLabels.showSerName = False
    ch.dataLabels.showCatName = False
    ch.dataLabels.showLegendKey = False
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.x_axis.scaling.orientation = "maxMin"  # 表と同じ上から下の順
    ch.y_axis.majorGridlines = None
    ch.legend = None
    ws.add_chart(ch, "K4")
    ws.cell(5, 26).value = "項目名（グラフ用）"

    # --- 表 ---
    nxt = sum_last + 3
    ja_cols = [(l["id"], l["profiles"]) for l in CFG["locales"] if l["keyword_set"] == "ja"]
    en_cols = [(l["id"], l["profiles"]) for l in CFG["locales"] if l["keyword_set"] == "en"]
    legend = "色: 1〜3 位＝濃い緑、4〜10 位＝薄い緑、11〜30 位＝黄、圏外＝灰、取得失敗・未計測＝橙"
    nxt = matrix(ws, nxt, "日本語の検索語（通常の検索結果の順位）", kw_ja, ja_cols, best, note=legend)
    nxt = matrix(ws, nxt, "英語の検索語（通常の検索結果の順位）", kw_en, en_cols, best)
    nxt = matrix(ws, nxt, "地図の枠（ローカルパック）での順位 — 日本語", kw_ja, ja_cols, best, field="local_rank",
                 note="空欄＝地図の枠に当院なし（Yahoo! JAPAN は取得対象外）")
    ws.freeze_panes = "B4"
    ws.print_area = f"A1:T{nxt}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # --- 1 位のサイト ---
    t = wb.create_sheet("1位のサイト")
    t.sheet_view.showGridLines = False
    t.column_dimensions["A"].width = 34
    t["A1"] = f"各検索語で 1 位に出たサイト（{day}）— 競合・まとめ記事の把握用"
    t["A1"].font = Font(bold=True, size=12, color=NAVY)
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
            t.cell(rr, 1, k["kw"]).border = BOX
            for i, (loc, p) in enumerate(combos, 2):
                lset = next(l["keyword_set"] for l in CFG["locales"] if l["id"] == loc)
                x = t.cell(rr, i)
                x.border, x.alignment = BOX, CENTER
                if lset != kset:
                    x.fill = PatternFill("solid", fgColor="F4F4F4")
                    continue
                rec = best.get((p, loc, k["kw"]))
                if rec and rec["status"] == "ok":
                    top = (rec["top10"] or "").split(" ")[0].replace("www.", "")
                    x.value = top
                    if "japangreen.vn" in top:
                        x.font = Font(bold=True, color=GREEN)
                else:
                    x.value = FAIL if rec else NOTYET
                    x.font = Font(color="B4532A")
            rr += 1
    t.freeze_panes = "B4"

    # --- データ ---
    d = wb.create_sheet("データ")
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
    n = wb.create_sheet("説明")
    n.column_dimensions["A"].width = 120
    st = Counter((r["browser_profile"], r["status"] if r["status"] == "ok" else "失敗") for r in best.values())
    lines = [
        ("この表の見方", True),
        (f"・計測日 {day}。毎日ベトナム時間 8:00 に自動で計測し、Google スプレッドシート「japangreen_検索順位」の ranks シートに追記している（このファイルはその抜粋）。", False),
        ("・順位は「通常の検索結果」だけを数えた当院（japangreen.vn）の順位。広告・地図の枠・AI の概要は数えない。30 位までに無ければ「圏外」。", False),
        ("・ブラウザは検索エンジン×端末で代表させている: Chrome（PC）＝Google パソコン、Edge（PC）＝Bing パソコン、Safari（iPhone）＝Google スマホ、Yahoo! JAPAN（PC）。Firefox・Mac の Safari は Chrome（PC）と同じ結果。", False),
        ("・地点は日本＝国単位、ベトナム＝ハノイ市（2026-11-06 から。それ以前のベトナムは国単位なので比べない）。"
         "ログインや閲覧履歴の影響は含まない（各自の画面とは多少ずれる）。", False),
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
        if b:
            c.font = Font(bold=True, color=NAVY)

    wb.save(out)
    print("saved", out, "day", day, "rows", len(best))


if __name__ == "__main__":
    main(*sys.argv[1:])
