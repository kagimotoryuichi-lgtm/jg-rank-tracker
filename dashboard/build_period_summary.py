"""検索順位の「期間まとめ」Excel を作る（毎日のダッシュボードとは別のファイル）。

使い方: python -X utf8 dashboard/build_period_summary.py <ranks.csv> <出力.xlsx> [開始日] [終了日]
- 開始日・終了日を省くと、データの最初の日〜最新日
- 色・書体・順位の数え方は build_excel_dashboard.py と同じ（そこから読み込む）
- 圏外（30 位以内に無し）は平均を出すとき 31 位とみなす。取得失敗の日は平均に入れない
- 2026-11-14 作成（鍵本さんの依頼。要素＝語句ごとの推移・上昇と下降・指名語と一般語の折れ線・競合・地域と端末・検索エンジン）
"""
import sys
from collections import OrderedDict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_excel_dashboard as B  # noqa: E402
from build_excel_dashboard import (BANDS, BOX, CENTER, CFG, DOWN, FAIL_BG, FAIL_FG, GREEN, LEFT, LOCALE_LABEL,  # noqa: E402
                                   NAVY, OUT_BG, OUT_FG, PROFILE_LABEL, UP, F, _cell, band_of, fill, legend,
                                   paint_rank, pick_latest, rank_of, short, style_header)
from openpyxl import Workbook  # noqa: E402
from openpyxl.chart import LineChart, Reference  # noqa: E402
from openpyxl.styles import Alignment  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

HANOI_FROM = "2026-11-06"      # ベトナムをハノイ市で計測し始めた日（11/5 は国単位）
COMP_FROM = "2026-11-13"       # 競合 4 院の順位を取り始めた日
BING_CHECK_FROM = "2026-11-14"  # Bing の「別の語の結果」を取得失敗として数え始めた日
GOOGLE = ("chrome-pc", "iphone")
GRAY = "8A8F98"
SERIES_COLORS = ["2F4078", "249835", "E0533D", "C98A1E", "7B61C9", "0F9D8A"]


def md(d):
    return f"{int(d[5:7])}/{int(d[8:])}"


def loc_of(lid):
    return next(l for l in CFG["locales"] if l["id"] == lid)


def keys_where(profiles=None, locales=None, group=None, kws=None):
    out = []
    for l in CFG["locales"]:
        if locales and l["id"] not in locales:
            continue
        for p in l["profiles"]:
            if profiles and p not in profiles:
                continue
            for k in CFG["keywords"][l["keyword_set"]]:
                if group and k["group"] != group:
                    continue
                if kws and k["kw"] not in kws:
                    continue
                out.append((p, l["id"], k["kw"]))
    return out


class Data:
    def __init__(self, rows, days):
        self.days = days
        self.picks = {d: pick_latest(rows, d) for d in days}

    def rank(self, d, key):
        """順位（int）／圏外（None）／取得失敗・未計測・比べられない日（False）"""
        if key[1].startswith("VN") and d < HANOI_FROM:
            return False
        return rank_of(self.picks[d].get(key))

    def vals(self, d, keys):
        return [n for n in (self.rank(d, k) for k in keys) if n is not False]

    def avg(self, ds, keys):
        v = [(n or 31) for d in ds for n in self.vals(d, keys)]
        return (sum(v) / len(v)) if v else None

    def count(self, d, keys, limit=30):
        v = self.vals(d, keys)
        return sum(1 for n in v if n is not None and n <= limit), len(v)


def title(ws, text, sub, width=14):
    for c in range(1, width + 1):
        ws.cell(1, c).fill = fill(NAVY)
    ws.row_dimensions[1].height = 30
    x = ws.cell(1, 1, "  " + text)
    x.font, x.alignment = F(bold=True, size=15, color="FFFFFF"), Alignment(vertical="center")
    x = ws.cell(2, 1, sub)
    x.font = F(size=9, color="555B66")
    ws.sheet_view.showGridLines = False


def paint_avg(x, v):
    if v is None:
        x.value, x.font = "—", F(size=8, color="C0C4CB")
        return
    if v > 30:
        x.value, x.fill, x.font = "圏外", fill(OUT_BG), F(size=8, color=OUT_FG)
        return
    bg, fg = BANDS[band_of(int(v))][1:]
    x.value, x.fill, x.font = round(v, 1), fill(bg), F(size=9, bold=True, color=fg)
    x.number_format = "0.0"


def paint_diff(x, a, b):
    """a＝前半の平均、b＝後半の平均（小さいほど上位）"""
    if a is None or b is None:
        x.value, x.font = "—", F(size=8, color="C0C4CB")
        return
    d = round(a - b, 1)
    if abs(d) < 0.05:
        x.value, x.font = "±0", F(size=9, color=GRAY)
    else:
        x.value = f"▲{abs(d):.1f}" if d > 0 else f"▼{abs(d):.1f}"
        x.font = F(size=9, bold=True, color=UP if d > 0 else DOWN)


def line_chart(g, head, last, ncols, w, h, ranks=True):
    lc = LineChart()
    lc.height, lc.width = h, w
    lc.add_data(Reference(g, min_col=2, max_col=1 + ncols, min_row=head, max_row=last), titles_from_data=True)
    lc.set_categories(Reference(g, min_col=1, min_row=head + 1, max_row=last))
    for s_, colr in zip(lc.series, SERIES_COLORS):
        s_.graphicalProperties.line.solidFill = colr
        s_.graphicalProperties.line.width = 24000
        s_.marker.symbol = "circle"
        s_.marker.size = 6
        s_.marker.graphicalProperties.solidFill = colr
        s_.marker.graphicalProperties.line.solidFill = colr
        s_.smooth = False
    lc.x_axis.delete = False
    lc.y_axis.delete = False
    lc.legend.position = "t"
    lc.legend.overlay = False
    lc.display_blanks = "gap"
    if ranks:
        lc.y_axis.scaling.orientation = "maxMin"  # 1 位を上に
        lc.y_axis.scaling.min, lc.y_axis.scaling.max = 1, 31
        lc.y_axis.majorUnit = 5
        lc.x_axis.crosses = "max"                 # 日付の軸は下に
    else:
        lc.y_axis.scaling.min = 0
    return lc


def graph_block(g, heading, header, rows_):
    """グラフ用シートに 1 ブロック書き、(見出し行, 最終行) を返す"""
    g.append([])
    g.append([heading])
    g.append(header)
    head = g.max_row
    for r in rows_:
        g.append(r)
    return head, g.max_row


def main(src, out, start=None, end=None):
    rows = B.load(src)
    all_days = sorted({r["date"] for r in rows})
    start, end = start or all_days[0], end or all_days[-1]
    days = [d for d in all_days if start <= d <= end]
    D = Data(rows, days)
    gkeys = keys_where(profiles=GOOGLE)
    # Google が取れた日（半分以上）だけを「計測できた日」とする
    gdays = [d for d in days if D.count(d, gkeys)[1] * 2 >= len(keys_where(profiles=GOOGLE,
             locales=None if d >= HANOI_FROM else ("JP-ja", "JP-en")))]
    cmp_days = [d for d in gdays if d >= HANOI_FROM]
    first_half, last_half = cmp_days[:2], cmp_days[-3:]
    base, last = cmp_days[0], cmp_days[-1]
    missing = [d for d in days if d not in gdays]
    period = f"{start}〜{end}"
    fh, lh = "・".join(md(d) for d in first_half), "・".join(md(d) for d in last_half)

    wb = Workbook()
    g = wb.create_sheet("グラフ用")
    g["A1"] = "グラフの元データ（自動で作る。編集不要）。順位の 31＝圏外、空白＝取得できなかった日"
    g.column_dimensions["A"].width = 12

    # ================= 1. まとめ =================
    ws = wb.active
    ws.title = "まとめ"
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 8
    for c in range(3, 13):
        ws.column_dimensions[get_column_letter(c)].width = 11.5
    title(ws, f"japangreen.vn 検索順位 期間まとめ（{md(start)}〜{md(end)}）",
          f"期間 {period}（{len(days)} 日）。Google を計測できた日は {len(gdays)} 日"
          + (f"（{'・'.join(md(d) for d in missing)} は残高切れで Google・Bing が取れていない）" if missing else "")
          + "。数え方は「説明」シート", 12)

    B.CMP = md(base)
    cur = {k: D.count(last, gkeys, k)[0] for k in (30, 10, 1)}
    old = {k: D.count(base, gkeys, k)[0] for k in (30, 10, 1)}
    tot = len(gkeys)
    for c in (1, 2):
        ws.cell(4, c).fill = fill(NAVY)
        for r in range(5, 8):
            ws.cell(r, c).fill = fill(B.CARD_BG)
    for r, txt, font in ((5, "Google（PC・iPhone）の最新日", F(size=9, color="555B66", bold=True)),
                         (6, md(last), F(size=24, bold=True, color=NAVY)),
                         (7, f"比べる日 {md(base)}（ハノイ市での計測の初日）", F(size=9, color=GRAY))):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        x = ws.cell(r, 1, txt)
        x.font, x.alignment = font, CENTER
    ws.row_dimensions[4].height = 5
    ws.row_dimensions[6].height = 42
    B.card(ws, 4, 3, "30 位以内に入った組", cur[30], f"/ {tot} 組", cur[30] - old[30])
    B.card(ws, 4, 5, "10 位以内（1 ページ目）", cur[10], "組", cur[10] - old[10])
    B.card(ws, 4, 7, "1 位", cur[1], "組", cur[1] - old[1])
    ba, la = D.avg(first_half, keys_where(GOOGLE, group="brand")), D.avg(last_half, keys_where(GOOGLE, group="brand"))
    B.card(ws, 4, 9, "指名語の平均順位", f"{la:.1f}", "位", None, sub=f"前半（{fh}）は {ba:.1f} 位")
    ga, gl = D.avg(first_half, keys_where(GOOGLE, group="general")), D.avg(last_half, keys_where(GOOGLE, group="general"))
    B.card(ws, 4, 11, "一般語の平均順位", f"{gl:.1f}", "位", None, sub=f"前半（{fh}）は {ga:.1f} 位")
    ws.cell(8, 3, f"※ 上の 3 つは Google の {tot} 組（検索語 × 地点 × PC・iPhone）。平均順位は後半（{lh}）の Google の平均で、圏外は 31 位とみなす").font = F(size=8, color=GRAY)

    # --- 推移のグラフ（Google の組数） ---
    head, lastr = graph_block(g, "Google（PC・iPhone）の組数の推移", ["日付", "30 位以内", "10 位以内", "1 位"], [
        [md(d)] + ([D.count(d, gkeys, k)[0] for k in (30, 10, 1)] if d in gdays else [None] * 3) for d in days])
    ws.cell(10, 1, "Google で順位に入った組数の推移（日別）").font = F(bold=True, size=11, color=NAVY)
    ws.cell(10, 5, f"※ {md(days[0])} のベトナムは国単位の計測（ハノイ市は {md(HANOI_FROM)} から）。空白＝計測できなかった日").font = F(size=8, color=GRAY)
    lc = line_chart(g, head, lastr, 3, 27.5, 8.5, ranks=False)
    ws.add_chart(lc, "A11")

    # --- 上昇・下降 ---
    r0 = 32   # グラフ（高さ 8.5cm）は 11〜30 行目あたりまで
    ws.cell(r0, 1, f"上がった組・下がった組（前半 {fh} の平均 → 後半 {lh} の平均）").font = F(bold=True, size=12, color=NAVY)
    ws.cell(r0 + 1, 1, "Google と Yahoo! JAPAN の組が対象（Bing は取得の不具合があるので外した）。圏外は 31 位とみなす。変化の大きい順に 8 件").font = F(size=8, color=GRAY)
    mv = []
    for key in keys_where(profiles=("chrome-pc", "iphone", "yahoo-pc")):
        a, b_ = D.avg(first_half, [key]), D.avg(last_half, [key])
        if a is None or b_ is None or abs(a - b_) < 0.05:
            continue
        mv.append((a - b_, key, a, b_))
    ups = sorted([m for m in mv if m[0] > 0], key=lambda m: -m[0])
    downs = sorted([m for m in mv if m[0] < 0], key=lambda m: m[0])

    def rt(v):
        return "圏外" if v > 30 else f"{v:.1f} 位"
    for j, (headtxt, items, col, c0) in enumerate((("上がった", ups, UP, 1), ("下がった", downs, DOWN, 7))):
        ws.merge_cells(start_row=r0 + 2, start_column=c0, end_row=r0 + 2, end_column=c0 + 5)
        x = ws.cell(r0 + 2, c0, f" {'▲' if j == 0 else '▼'} {headtxt}（{len(items)} 組）")
        x.font, x.alignment = F(bold=True, size=10, color="FFFFFF"), Alignment(vertical="center")
        for c in range(c0, c0 + 6):
            ws.cell(r0 + 2, c).fill = fill(col)
        for i in range(8):
            r = r0 + 3 + i
            ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c0 + 5)
            for c in range(c0, c0 + 6):
                ws.cell(r, c).fill = fill("FAFBFC")
            x = ws.cell(r, c0)
            if i < len(items):
                _, (p, loc, kw), a, b_ = items[i]
                x.value = B.CellRichText([
                    B.TextBlock(B.InlineFont(rFont=B.FONT, sz=9, b=True, color="222222"), kw),
                    B.TextBlock(B.InlineFont(rFont=B.FONT, sz=8, color=GRAY), f"　{short(loc, p)}　"),
                    B.TextBlock(B.InlineFont(rFont=B.FONT, sz=9, b=True, color=col), f"{rt(a)} → {rt(b_)}")])
            x.alignment = Alignment(vertical="center", indent=1)
            ws.row_dimensions[r].height = 18

    # --- 入った・外れた ---
    r = r0 + 12
    ws.cell(r, 1, "30 位以内への出入り（Google・Yahoo! JAPAN）").font = F(bold=True, size=12, color=NAVY)
    new_in = [(k, b_) for d_, k, a, b_ in mv if a > 30 and b_ <= 30]
    gone = [(k, a) for d_, k, a, b_ in mv if a <= 30 and b_ > 30]
    for j, (lab, items, col) in enumerate(((f"前半は圏外 → 後半に 30 位以内（{len(new_in)} 組）", new_in, UP),
                                           (f"前半は 30 位以内 → 後半は 3 日とも圏外（{len(gone)} 組）", gone, DOWN))):
        rr = r + 1 + j
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=12)
        txt = "、".join(f"{k[2]}（{short(k[1], k[0])}）" for k, _ in items[:12]) or "なし"
        x = ws.cell(rr, 1)
        x.value = B.CellRichText([B.TextBlock(B.InlineFont(rFont=B.FONT, sz=9, b=True, color=col), lab + "："),
                                  B.TextBlock(B.InlineFont(rFont=B.FONT, sz=8, color="333333"), txt
                                              + (f" ほか {len(items) - 12} 組" if len(items) > 12 else ""))])
        x.alignment = Alignment(vertical="top", wrap_text=True)
        ws.row_dimensions[rr].height = 44

    # ================= 2. 語句ごとの推移 =================
    t = wb.create_sheet("語句ごとの推移")
    title(t, "語句ごとの順位の推移（日別）", f"期間 {period}。前半＝{fh} の平均、後半＝{lh} の平均（圏外は 31 位）。"
          "変化の ▲ は上がった・▼ は下がった", 6 + len(days))
    t.column_dimensions["A"].width = 30
    t.column_dimensions["B"].width = 7
    t.column_dimensions["C"].width = 27
    hr = 5
    t.cell(4, 1, "色の見方").font = F(bold=True, size=9, color="555B66")
    legend(t, 4, 4, [(BANDS[0][0], BANDS[0][1], BANDS[0][2]), (BANDS[1][0], BANDS[1][1], BANDS[1][2]),
                     (BANDS[2][0], BANDS[2][1], BANDS[2][2]), ("圏外", OUT_BG, OUT_FG), ("取得失敗", FAIL_BG, FAIL_FG)])
    heads = ["検索語", "種類", "地点・ブラウザ"] + [md(d) for d in days] + ["前半", "後半", "変化"]
    for i, h in enumerate(heads, 1):
        t.cell(hr, i, h)
        if i > 3:
            t.column_dimensions[get_column_letter(i)].width = 8.5
    style_header(t, hr, 1, len(heads))
    r = hr + 1
    for l in CFG["locales"]:
        for k in CFG["keywords"][l["keyword_set"]]:
            for p in l["profiles"]:
                key = (p, l["id"], k["kw"])
                _cell(t, r, 1, k["kw"], align=LEFT)
                _cell(t, r, 2, "ブランド" if k["group"] == "brand" else "一般", size=8, color="555B66")
                _cell(t, r, 3, short(l["id"], p), size=8, color="555B66", align=LEFT)
                for i, d in enumerate(days):
                    x = t.cell(r, 4 + i)
                    x.alignment, x.border = CENTER, BOX
                    if l["id"].startswith("VN") and d < HANOI_FROM:
                        x.value, x.font = "国単位", F(size=8, color="B0B4BB")
                    else:
                        paint_rank(x, D.picks[d].get(key), None)
                a, b_ = D.avg(first_half, [key]), D.avg(last_half, [key])
                c0 = 4 + len(days)
                for c, v in ((c0, a), (c0 + 1, b_)):
                    x = t.cell(r, c)
                    x.alignment, x.border = CENTER, BOX
                    paint_avg(x, v)
                x = t.cell(r, c0 + 2)
                x.alignment, x.border = CENTER, BOX
                paint_diff(x, a, b_)
                t.row_dimensions[r].height = 16
                r += 1
    t.freeze_panes = t.cell(hr + 1, 4)
    t.auto_filter.ref = f"A{hr}:{get_column_letter(len(heads))}{r - 1}"

    # ================= 3. 折れ線（指名語・一般語） =================
    c = wb.create_sheet("指名語・一般語のグラフ")
    title(c, "指名語・一般語の順位の推移（Google の PC と iPhone の平均）",
          "縦軸は順位（上ほど上位）。31＝圏外。線が切れている日は計測できなかった日。グラフの元の数字は「グラフ用」シート", 16)
    for i in range(1, 17):
        c.column_dimensions[get_column_letter(i)].width = 11

    def series_rows(defs):
        """defs = [(線の名前, keys), ...] → 日別の平均順位の行"""
        return [[md(d)] + [(round(v, 1) if (v := D.avg([d], ks)) is not None else None) for _, ks in defs] for d in days]

    brand_defs = [(l["label"], keys_where(GOOGLE, (l["id"],), "brand")) for l in CFG["locales"]]
    head, lastr = graph_block(g, "指名語の平均順位（地点別・Google）", ["日付"] + [n for n, _ in brand_defs], series_rows(brand_defs))
    c.cell(4, 1, "1. 指名語（ジャパングリーン 系）の平均順位 — 地点別").font = F(bold=True, size=11, color=NAVY)
    c.add_chart(line_chart(g, head, lastr, len(brand_defs), 27, 9), "A5")

    def top_general(lids, n):
        ks = [k["kw"] for k in CFG["keywords"][loc_of(lids[0])["keyword_set"]] if k["group"] == "general"]
        scored = sorted(((D.avg(cmp_days, keys_where(GOOGLE, lids, kws=(kw,))) or 99, kw) for kw in ks))
        return [kw for s_, kw in scored[:n] if s_ <= 30.9]

    row = 27  # グラフ 1 つ（高さ 9cm）で約 21 行使う
    for no, (lids, lab) in enumerate(((("VN-ja",), "ハノイ×日本語"), (("JP-ja",), "日本×日本語"),
                                       (("VN-en",), "ハノイ×英語"), (("JP-en",), "日本×英語")), 2):
        kws = top_general(lids, 6)
        defs = [(kw, keys_where(GOOGLE, lids, kws=(kw,))) for kw in kws]
        c.cell(row, 1, f"{no}. 一般語の順位 — {lab}（期間の平均が上位の {len(kws)} 語）").font = F(bold=True, size=11, color=NAVY)
        if not defs:
            c.cell(row + 1, 1, "期間中に 30 位以内に入った一般語なし").font = F(size=9, color=GRAY)
            row += 3
            continue
        head, lastr = graph_block(g, f"一般語の順位（{lab}・Google）", ["日付"] + kws, series_rows(defs))
        c.add_chart(line_chart(g, head, lastr, len(defs), 27, 9), f"A{row + 1}")
        row += 23

    # ================= 4. 競合 =================
    k4 = wb.create_sheet("競合4院")
    comps = CFG.get("competitors", [])
    cdays = [d for d in gdays if d >= COMP_FROM]
    title(k4, "競合 4 院との比較（Google・通常の検索結果）",
          f"競合の順位は {md(COMP_FROM)} から取り始めたので、比べられるのは {('・'.join(md(d) for d in cdays)) or 'なし'} の {len(cdays)} 日だけ", 14)
    k4.column_dimensions["A"].width = 32
    for i in range(2, 16):
        k4.column_dimensions[get_column_letter(i)].width = 9.5
    names = [("当院", "rank")] + [(x["label"], "rank_" + x["id"]) for x in comps]

    def cval(d, key, col):
        rec = D.picks[d].get(key)
        if not rec or rec.get("status") != "ok":
            return False
        v = (rec.get(col) or "").strip()
        return int(float(v)) if v else None

    r = 4
    k4.cell(r, 1, "1. 順位に入った組数（Google の PC・iPhone、全部の検索語）").font = F(bold=True, size=11, color=NAVY)
    r += 1
    k4.cell(r, 1, "")
    cc = 2
    for d in cdays:
        k4.cell(r, cc, md(d))
        k4.merge_cells(start_row=r, start_column=cc, end_row=r, end_column=cc + 2)
        cc += 3
    style_header(k4, r, 1, cc - 1)
    r += 1
    cc = 2
    for d in cdays:
        for h in ("30 位以内", "10 位以内", "1 位"):
            k4.cell(r, cc, h)
            cc += 1
    style_header(k4, r, 1, cc - 1)
    r += 1
    for nm, col in names:
        own = nm == "当院"
        _cell(k4, r, 1, nm, bold=own, color=GREEN if own else "222222", align=LEFT)
        cc = 2
        for d in cdays:
            v = [n for n in (cval(d, key, col) for key in gkeys) if n not in (False, None)]
            for n in (len(v), sum(x <= 10 for x in v), sum(x == 1 for x in v)):
                _cell(k4, r, cc, n, bold=own)
                cc += 1
        r += 1
    r += 1
    k4.cell(r, 1, "2. 一般語ごとの順位（Chrome（PC）。「—」＝30 位以内に無し）").font = F(bold=True, size=11, color=NAVY)
    r += 2
    for l in CFG["locales"]:
        k4.cell(r, 1, l["label"]).font = F(bold=True, size=10, color=NAVY)
        r += 1
        k4.cell(r, 1, "検索語")
        cc = 2
        for nm, _ in names:
            k4.cell(r, cc, nm)
            if len(cdays) > 1:
                k4.merge_cells(start_row=r, start_column=cc, end_row=r, end_column=cc + len(cdays) - 1)
            cc += len(cdays)
        style_header(k4, r, 1, cc - 1)
        r += 1
        cc = 2
        for _ in names:
            for d in cdays:
                k4.cell(r, cc, md(d))
                cc += 1
        style_header(k4, r, 1, cc - 1)
        r += 1
        for k in CFG["keywords"][l["keyword_set"]]:
            if k["group"] == "brand":
                continue
            _cell(k4, r, 1, k["kw"], align=LEFT)
            cc = 2
            for nm, col in names:
                for d in cdays:
                    x = _cell(k4, r, cc, "")
                    n = cval(d, ("chrome-pc", l["id"], k["kw"]), col)
                    if n is False:
                        x.value, x.fill, x.font = "取得失敗", fill(FAIL_BG), F(size=8, color=FAIL_FG)
                    elif n is None:
                        x.value, x.font = "—", F(size=8, color="C0C4CB")
                    else:
                        bg, fg = BANDS[band_of(n)][1:]
                        x.value, x.fill, x.font = n, fill(bg), F(size=9, bold=True, color=fg)
                    cc += 1
            r += 1
        r += 1

    # ================= 5. 地域・端末・検索エンジン =================
    s = wb.create_sheet("地域・端末・エンジン別")
    title(s, "地域・端末・検索エンジン別の推移", "各マスは「30 位以内の組数 / 取得できた組数」。空欄の「—」は計測できなかった日。"
          f"Edge（Bing）は {md(BING_CHECK_FROM)} より前、別の語の結果が混じって低めに出ている", 5 + len(days))
    s.column_dimensions["A"].width = 42
    s.column_dimensions["B"].width = 34
    for i in range(3, 6 + len(days)):
        s.column_dimensions[get_column_letter(i)].width = 9
    ja, en = ("JP-ja", "VN-ja"), ("JP-en", "VN-en")
    groups = [
        ("検索エンジン（日本×日本語・PC・同じ 17 語）", [
            ("Google（Chrome）", keys_where(("chrome-pc",), ("JP-ja",))),
            ("Bing（Edge）", keys_where(("edge-pc",), ("JP-ja",))),
            ("Yahoo! JAPAN", keys_where(("yahoo-pc",), ("JP-ja",)))]),
        ("検索エンジン（全地点・PC）", [
            ("Google（Chrome）", keys_where(("chrome-pc",))), ("Bing（Edge）", keys_where(("edge-pc",)))]),
        ("地域（Google の PC・iPhone）", [
            ("日本から検索", keys_where(GOOGLE, ("JP-ja", "JP-en"))), ("ハノイから検索", keys_where(GOOGLE, ("VN-ja", "VN-en")))]),
        ("端末（Google）", [("PC（Chrome）", keys_where(("chrome-pc",))), ("スマホ（iPhone）", keys_where(("iphone",)))]),
        ("検索語の言語（Google の PC・iPhone）", [("日本語の検索語", keys_where(GOOGLE, ja)), ("英語の検索語", keys_where(GOOGLE, en))]),
        ("検索語の種類（Google の PC・iPhone）", [("指名語（ブランド）", keys_where(GOOGLE, group="brand")),
                                               ("一般語", keys_where(GOOGLE, group="general"))]),
        ("地点 × ブラウザ（13 通り）", [(short(l["id"], p), keys_where((p,), (l["id"],)))
                                    for l in CFG["locales"] for p in l["profiles"]]),
    ]
    r = 4
    for limit, lab in ((30, "30 位以内"), (10, "10 位以内（1 ページ目）")):
        s.cell(r, 1, f"■ {lab}に入った組数").font = F(bold=True, size=12, color=NAVY)
        r += 1
        heads = ["区分", "内訳"] + [md(d) for d in days] + ["前半の平均", "後半の平均", "変化"]
        for i, h in enumerate(heads, 1):
            s.cell(r, i, h)
        style_header(s, r, 1, len(heads))
        r += 1
        for gname, items in groups:
            for j, (nm, ks) in enumerate(items):
                _cell(s, r, 1, gname if j == 0 else "", size=8, bold=True, color=NAVY, align=LEFT)
                _cell(s, r, 2, f"{nm}（{len(ks)} 組）", size=9, align=LEFT)
                for i, d in enumerate(days):
                    n, got = D.count(d, ks, limit)
                    x = _cell(s, r, 3 + i, "")
                    if not got:
                        x.value, x.font = "—", F(size=8, color="C0C4CB")
                        continue
                    x.value = f"{n}/{got}"
                    ratio = n / got
                    bg, fg = (BANDS[0][1:] if ratio >= 0.6 else BANDS[1][1:] if ratio >= 0.3 else BANDS[2][1:] if ratio > 0 else (OUT_BG, OUT_FG))
                    x.fill, x.font = fill(bg), F(size=9, bold=True, color=fg)
                a, b_ = D.avg(first_half, ks), D.avg(last_half, ks)
                c0 = 3 + len(days)
                for cc, v in ((c0, a), (c0 + 1, b_)):
                    x = _cell(s, r, cc, "")
                    if v is None:
                        x.value, x.font = "—", F(size=8, color="C0C4CB")
                    else:
                        x.value, x.number_format = round(v, 1), "0.0"
                paint_diff(_cell(s, r, c0 + 2, ""), a, b_)
                r += 1
        r += 1
    s.cell(r, 1, f"※ 平均＝平均順位（圏外は 31 位）。前半＝{fh}、後半＝{lh}。「地点 × ブラウザ」より上の行はどれも順位の平均で、組数の平均ではない").font = F(size=8, color=GRAY)
    s.freeze_panes = "C4"

    # ================= 6. 説明 =================
    n = wb.create_sheet("説明")
    n.column_dimensions["A"].width = 125
    fail = OrderedDict()
    for d in days:
        for p in CFG["profiles"]:
            ks = keys_where((p,))
            ok = sum(1 for k in ks if rank_of(D.picks[d].get(k)) is not False)
            fail.setdefault(d, []).append(f"{PROFILE_LABEL[p]} {ok}/{len(ks)}")
    lines = [
        ("この表の見方", True),
        (f"・期間 {period}。元のデータは毎朝の自動計測（data/ranks.csv）。同じ日に複数回計測した組は、取得できた最後の結果を使う。", False),
        ("・順位は通常の検索結果での当院（japangreen.vn）の順位。30 位までに無ければ「圏外」。平均を出すときは圏外を 31 位とみなし、取得できなかった日は平均に入れない。", False),
        ("・「組」＝検索語 × ブラウザ × 地点。Chrome（PC）＝Google パソコン、Safari（iPhone）＝Google スマホ、Edge（PC）＝Bing、Yahoo! JAPAN（PC）。", False),
        (f"・前半＝{fh}、後半＝{lh}。日ごとの上下が大きいので、初日と最終日ではなく前半と後半の平均で比べている。", False),
        ("", False),
        ("数字を読むときの注意", True),
        (f"・{('・'.join(md(d) for d in missing)) or 'なし'} は DataForSEO の残高切れで Google・Bing が取れていない（Yahoo! JAPAN だけある）。グラフは空白にしている。", False),
        (f"・ベトナムの地点は {md(HANOI_FROM)} からハノイ市（{md(days[0])} は国単位）。{md(days[0])} のベトナムの数字は比べる対象から外した。", False),
        (f"・Edge（PC）＝Bing は、計測会社が別の検索語の結果を返すことがある。{md(BING_CHECK_FROM)} からはそれを取得失敗として数えるが、それより前は圏外として混じっている。Bing の日ごとの上下は実際の順位の動きとは限らない。上昇・下降の一覧からは Bing を外した。", False),
        (f"・{md(last)} の Bing は 52 組のうち取得できたのが少ない（下の一覧）。Microsoft の障害（11/12〜13）と時期が重なるが、原因は確認できていない。", False),
        (f"・競合 4 院の順位は {md(COMP_FROM)} から取り始めた。それより前の比較はできない。", False),
        ("・Google の順位は日によって大きく動く（同じ語で 1 位の翌日に 20 位台など）。1 日だけの数字で判断しない。", False),
        ("", False),
        ("取得できた組数（日別・ブラウザ別）", True),
    ] + [(f"・{md(d)}: " + "／".join(v), False) for d, v in fail.items()]
    for i, (txt, b_) in enumerate(lines, 1):
        x = n.cell(i, 1, txt)
        x.alignment = Alignment(wrap_text=True, vertical="top")
        x.font = F(bold=True, color=NAVY) if b_ else F(size=10)

    wb.move_sheet("グラフ用", offset=len(wb.sheetnames))
    wb.active = 0
    wb.save(out)
    print("saved", out, "days", len(days), "google days", len(gdays), "first", first_half, "last", last_half)
    print("cards", cur, old, round(ba, 2), round(la, 2), round(ga, 2), round(gl, 2), "ups", len(ups), "downs", len(downs),
          "new_in", len(new_in), "gone", len(gone))


if __name__ == "__main__":
    main(*sys.argv[1:])
