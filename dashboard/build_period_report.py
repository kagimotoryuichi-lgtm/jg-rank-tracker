"""検索順位の「期間報告書」Excel を作る（SEO の初心者が読む報告書の体裁。A4 縦で印刷・PDF にできる）。

使い方: python -X utf8 dashboard/build_period_report.py <ranks.csv> <出力.xlsx> [開始日] [終了日]
- 毎日のダッシュボード（build_excel_dashboard.py）・期間まとめ（build_period_summary.py）とは別のファイル
- 各章に「狙い・見方・注意・読み取れること」を書く。先頭に結論、2 枚目に SEO の基本と用語
- Google は GOOGLE_ODD_FROM 以降を「要確認」として分ける（計測会社が返す結果ページの異常が疑われるため。
  数字は消さずに載せるが、良くなった・悪くなったの判断には使わない）
- 平均順位は圏外を 31 位とみなす。取得できなかった日は平均に入れない
- 2026-11-17 作成（鍵本さんの依頼。構成は Codex の意見を参考にした）
"""
import datetime as dt
import math
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_excel_dashboard as B  # noqa: E402
from build_excel_dashboard import (BANDS, BOX, CENTER, CFG, FAIL_BG, FAIL_FG, GREEN, LEFT, NAVY, OUT_BG, OUT_FG,  # noqa: E402
                                   PROFILE_LABEL, F, _cell, band_of, fill, rank_of, short, style_header)
from build_period_summary import COMP_FROM, GOOGLE, GRAY, Data, graph_block, keys_where, md  # noqa: E402
from openpyxl import Workbook  # noqa: E402
from openpyxl.chart import LineChart, Reference  # noqa: E402
from openpyxl.styles import Alignment, Border, Side  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402
from openpyxl.worksheet.page import PageMargins  # noqa: E402
from openpyxl.worksheet.pagebreak import Break  # noqa: E402
from openpyxl.worksheet.properties import PageSetupProperties  # noqa: E402

GOOGLE_ODD_FROM = "2026-11-14"   # Google の結果ページの異常が疑われ始めた日（これ以降は「要確認」）
GROUP_SITES = ("japangreen.vn", "japan-green.com.sg", "japangreen.co.uk", "greenclinic.com.cn",
               "japangreen.co.jp", "fkmc.or.jp")
NCOL, W_A, W = 13, 24, 6.3       # 本文の列: A ＋ 12 列
ORANGE, WARN_BG, WARN_FG = "C77700", "FFF4E5", "7A4A00"
SOFT = "F3F6FA"
# 実際の検索画面で確かめたこと（日付ごと。文章にそのまま入れる）
YAHOO_NOTE = {"2026-11-17": "同じ日に実際の Yahoo! JAPAN を開き、2 つの検索語（「ハノイ クリニック」で 3 番目、「ハノイ 婦人科」で 2 番目）に当院が出ることを確認しました。"}
LIGHT = {"2F4078": "9AA6CF", "249835": "8FD09A", "E0533D": "F0A99E"}
LABELS = {"狙い": NAVY, "見方": GREEN, "注意": ORANGE, "読み取れること": "555B66"}


# ---------- 体裁の部品 ----------
def setup(ws, landscape=False, widths=None):
    ws.sheet_view.showGridLines = False
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    else:
        ws.column_dimensions["A"].width = W_A
        for c in range(2, NCOL + 1):
            ws.column_dimensions[get_column_letter(c)].width = W
    ps = ws.page_setup
    ps.paperSize, ps.orientation, ps.fitToWidth, ps.fitToHeight = 9, ("landscape" if landscape else "portrait"), 1, 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.55, bottom=0.6)
    ws.print_options.horizontalCentered = True
    ws.oddFooter.center.text = "japangreen.vn 検索順位 期間報告書　&P / &N"
    ws.oddFooter.center.size = 8


def units(ws, c1, c2):
    return sum(ws.column_dimensions[get_column_letter(c)].width or 8.43 for c in range(c1, c2 + 1))


def vlen(s):
    return sum(1 if ord(ch) > 0x2E7F or ch in "「」（）、。・〜—※■▲▼→×" else 0.55 for ch in s)


def text(ws, r, c1, c2, s, size=10, bold=False, color="222222", bg=None, valign="top", indent=0, min_h=0):
    """結合セルに折り返しの文を置き、文字数から行の高さを決める。次の行番号を返す"""
    if c2 > c1:
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    x = ws.cell(r, c1, s)
    x.font = F(size=size, bold=bold, color=color)
    x.alignment = Alignment(wrap_text=True, vertical=valign, indent=indent)
    cpl = max(4, units(ws, c1, c2) * 0.97 / (2.0 * size / 10) - indent * 1.5)
    lines = sum(max(1, math.ceil(vlen(p) / cpl)) for p in s.split("\n"))
    h = max(min_h, lines * size * 1.42 + 5)
    ws.row_dimensions[r].height = max(ws.row_dimensions[r].height or 0, h)
    if bg:
        for c in range(c1, c2 + 1):
            ws.cell(r, c).fill = fill(bg)
    return r + 1


def chapter(ws, no, title_, lead):
    for c in range(1, ws.max_column_used + 1):
        ws.cell(1, c).fill = fill(NAVY)
    x = ws.cell(1, 1, f"  {no}　{title_}" if no else f"  {title_}")
    x.font, x.alignment = F(bold=True, size=14, color="FFFFFF"), Alignment(vertical="center")
    ws.row_dimensions[1].height = 30
    return text(ws, 2, 1, ws.max_column_used, lead, size=9, color="555B66") + 1


def h2(ws, r, s):
    ws.blocks.append(r)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ws.max_column_used)
    x = ws.cell(r, 1, s)
    x.font, x.alignment = F(bold=True, size=12, color=NAVY), Alignment(vertical="center")
    for c in range(1, ws.max_column_used + 1):
        ws.cell(r, c).border = Border(bottom=Side(style="medium", color=GREEN))
    ws.row_dimensions[r].height = 24
    return r + 1


def note(ws, r, label, s):
    """左に「狙い／見方／注意／読み取れること」の札、右に説明文"""
    x = ws.cell(r, 1, label)
    x.fill, x.font = fill(LABELS[label]), F(bold=True, size=10, color="FFFFFF")
    x.alignment = Alignment(horizontal="center", vertical="center")
    bg = WARN_BG if label == "注意" else SOFT
    nr = text(ws, r, 2, ws.max_column_used, s, size=10, bg=bg, valign="center", indent=1, min_h=22)
    return nr


def paginate(ws, capacity=800):
    """章の中の区切り（h2）が、ページをまたがないように改ページを入れる。1 ページより長い区切りはそのまま流す。
    capacity＝A4 縦 1 ページに入る行の高さの合計（pt。幅を 1 ページに合わせたときの縮小を見込んだ値）"""
    last = ws.max_row
    hs = {r: (ws.row_dimensions[r].height or 15) for r in range(1, last + 1)}
    starts = sorted(set(ws.blocks))
    used, page_top = 0.0, 1
    bounds = starts + [last + 1]
    pos = 1
    for i, st in enumerate(starts):
        used += sum(hs[r] for r in range(pos, st))
        size = sum(hs[r] for r in range(st, bounds[i + 1]))
        if used + size > capacity and st > page_top and size <= capacity:
            ws.row_breaks.append(Break(id=st - 1))
            used, page_top = 0.0, st
        used += size
        while used > capacity:
            used -= capacity
        pos = bounds[i + 1]


def bullets(ws, r, items, size=10, bg=None):
    for s in items:
        r = text(ws, r, 1, ws.max_column_used, "・" + s, size=size, bg=bg, indent=1)
    return r


def gap(ws, r, h=8):
    ws.row_dimensions[r].height = h
    return r + 1


def reserve(ws, r, h_cm):
    """グラフを置く行を確保する（行の高さ 15pt＝約 0.53cm）"""
    n = math.ceil(h_cm / 0.529) + 1
    for i in range(n):
        ws.row_dimensions[r + i].height = 15
    return r + n


def day_header(ws, r, c0, days, odd_days, heads_before=(), heads_after=()):
    """見出し行。Google が要確認の日は橙色にする"""
    hs = list(heads_before) + [md(d) for d in days] + list(heads_after)
    for i, h in enumerate(hs):
        ws.cell(r, c0 + i, h)
    style_header(ws, r, c0, c0 + len(hs) - 1)
    for i, d in enumerate(days):
        if d in odd_days:
            ws.cell(r, c0 + len(heads_before) + i).fill = fill(ORANGE)
    ws.row_dimensions[r].height = 18
    return r + 1


def rank_cell(x, n):
    """n＝順位（int）／圏外（None）／取得失敗（False）"""
    x.alignment, x.border = CENTER, BOX
    if n is False:
        x.value, x.fill, x.font = "失敗", fill(FAIL_BG), F(size=8, color=FAIL_FG)
    elif n is None:
        x.value, x.fill, x.font = "圏外", fill(OUT_BG), F(size=8, color=OUT_FG)
    else:
        bg, fg = BANDS[band_of(n)][1:]
        x.value, x.fill, x.font = n, fill(bg), F(size=10, bold=True, color=fg)


def avg_cell(x, v, digits=1):
    x.alignment, x.border = CENTER, BOX
    if v is None:
        x.value, x.font = "—", F(size=8, color="C0C4CB")
    elif v > 30:
        x.value, x.fill, x.font = "圏外", fill(OUT_BG), F(size=8, color=OUT_FG)
    else:
        bg, fg = BANDS[band_of(int(v))][1:]
        x.value, x.fill, x.font = round(v, digits), fill(bg), F(size=9, bold=True, color=fg)
        x.number_format = "0.0"


def color_legend(ws, r, with_odd=True):
    items = [(BANDS[0][0], BANDS[0][1], BANDS[0][2]), (BANDS[1][0], BANDS[1][1], BANDS[1][2]),
             (BANDS[2][0], BANDS[2][1], BANDS[2][2]), ("圏外", OUT_BG, OUT_FG), ("失敗", FAIL_BG, FAIL_FG)]
    x = ws.cell(r, 1, "色の見方")
    x.font, x.alignment = F(bold=True, size=9, color="555B66"), Alignment(horizontal="right", vertical="center")
    c = 2
    for t, bg, fg in items:
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 1)
        y = ws.cell(r, c, t)
        y.fill, y.font, y.alignment = fill(bg), F(size=8, bold=True, color=fg), CENTER
        ws.cell(r, c + 1).fill = fill(bg)
        c += 2
    if with_odd:
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 1)
        y = ws.cell(r, c, "日付が橙＝要確認")
        y.fill, y.font, y.alignment = fill(ORANGE), F(size=7, bold=True, color="FFFFFF"), CENTER
        ws.cell(r, c + 1).fill = fill(ORANGE)
    ws.row_dimensions[r].height = 16
    return r + 1


def rng(vals, unit=""):
    vals = [v for v in vals if v is not None]
    if not vals:
        return "—"
    lo, hi = min(vals), max(vals)
    return f"{lo}{unit}" if lo == hi else f"{lo}〜{hi}{unit}"


def two_phase_chart(g, heading, names, colors, days, ok_days, fns, w=19.5, h=7.5, ymax=None):
    """判断に使える日は実線、要確認の日は同じ色の薄い点線にする折れ線"""
    has_odd = any(d not in ok_days for d in days)
    rows_ = []
    for d in days:
        vals = [f(d) for f in fns]
        solid = [v if d in ok_days else None for v in vals]
        dash = [v if (d not in ok_days or (ok_days and d == ok_days[-1])) else None for v in vals]
        rows_.append([md(d)] + solid + (dash if has_odd else []))
    header = ["日付"] + list(names) + ([n + "（要確認）" for n in names] if has_odd else [])
    head, last = graph_block(g, heading, header, rows_)
    lc = LineChart()
    lc.height, lc.width = h, w
    lc.add_data(Reference(g, min_col=2, max_col=len(header), min_row=head, max_row=last), titles_from_data=True)
    lc.set_categories(Reference(g, min_col=1, min_row=head + 1, max_row=last))
    n = len(names)
    for i, s_ in enumerate(lc.series):
        base = colors[i % n]
        colr = base if i < n else LIGHT.get(base, "B0B4BB")
        s_.graphicalProperties.line.solidFill = colr
        s_.graphicalProperties.line.width = 26000 if i < n else 20000
        if i >= n:
            s_.graphicalProperties.line.dashStyle = "sysDash"
        s_.marker.symbol = "circle"
        s_.marker.size = 6
        s_.marker.graphicalProperties.solidFill = colr if i < n else "FFFFFF"
        s_.marker.graphicalProperties.line.solidFill = colr
        s_.smooth = False
    lc.x_axis.delete = False
    lc.y_axis.delete = False
    lc.legend.position = "b"
    lc.legend.overlay = False
    lc.display_blanks = "gap"
    lc.y_axis.scaling.min = 0
    if ymax:
        lc.y_axis.scaling.max = ymax
    return lc


class Sheet:
    """ws に「使う列の数」を持たせるだけの薄い包み"""
    def __init__(self, wb, name, ncol=NCOL, first=False, **kw):
        self.ws = wb.active if first else wb.create_sheet(name)
        self.ws.title = name
        self.ws.max_column_used = ncol
        self.ws.blocks = []
        setup(self.ws, **kw)


def main(src, out, start=None, end=None):
    rows = B.load(src)
    all_days = sorted({r["date"] for r in rows})
    start, end = start or all_days[0], end or all_days[-1]
    days = [d for d in all_days if start <= d <= end]
    D = Data(rows, days)
    ok_days = [d for d in days if d < GOOGLE_ODD_FROM]      # Google の数字を判断に使える日
    odd_days = [d for d in days if d >= GOOGLE_ODD_FROM]    # Google は要確認の日
    okp = f"{md(ok_days[0])}〜{md(ok_days[-1])}" if ok_days else "なし"
    oddp = f"{md(odd_days[0])}〜{md(odd_days[-1])}" if odd_days else "なし"
    period = f"{start}〜{end}"
    today = dt.date.today().isoformat()

    gkeys = keys_where(profiles=GOOGLE)
    brand = keys_where(GOOGLE, group="brand")
    general = keys_where(GOOGLE, group="general")
    ykeys = keys_where(profiles=("yahoo-pc",))
    bkeys = keys_where(profiles=("edge-pc",))
    n_all = len(keys_where())
    n_kw = len({k["kw"] for s_ in CFG["keywords"].values() for k in s_})

    def cnt(keys, limit=30):
        return lambda d: D.count(d, keys, limit)[0]

    def odd_rows(d):
        """指名語（Google）のうち、当院もグループの公式サイトも上位 10 件に無い組の数"""
        n = 0
        for key in brand:
            rec = D.picks[d].get(key)
            if rec and rec.get("status") == "ok" and not any(s_ in (rec.get("top10") or "") for s_ in GROUP_SITES):
                n += 1
        return n

    def ok_avg(keys):
        return D.avg(ok_days, keys)

    # --- 文章に使う数字 ---
    b30 = [cnt(brand)(d) for d in ok_days]
    b3 = [cnt(brand, 3)(d) for d in ok_days]
    b30_odd = [cnt(brand)(d) for d in odd_days]
    g30 = [cnt(general)(d) for d in ok_days]
    g10 = [cnt(general, 10)(d) for d in ok_days]
    g30_odd = [cnt(general)(d) for d in odd_days]
    y30 = [cnt(ykeys)(d) for d in days]
    y10 = [cnt(ykeys, 10)(d) for d in days]
    bing_got = [D.count(d, bkeys)[1] for d in days]
    odd_n = {d: odd_rows(d) for d in days}
    b_avg, g_avg = ok_avg(brand), ok_avg(general)
    y_prev = y30[:-1]
    y_jump = len(days) > 1 and y30[-1] >= max(y_prev) + 5
    if y_jump:
        y_text = (f"Yahoo! JAPAN（日本から・日本語 {len(ykeys)} 語）は、{md(days[0])}〜{md(days[-2])} は 30 位以内が {rng(y_prev)} 語でしたが、"
                  f"最終日の {md(end)} に {y30[-1]} 語（上位 10 件は {y10[-1]} 語）へ急に増えました。" + YAHOO_NOTE.get(end, "")
                  + "ただし 1 日だけの記録なので、続くかどうかを確認します。")
    else:
        y_text = f"Yahoo! JAPAN（日本から・日本語 {len(ykeys)} 語）は期間を通して安定しています。30 位以内は毎日 {rng(y30)} 語、上位 10 件は {rng(y10)} 語でした。"

    wb = Workbook()
    g = wb.create_sheet("グラフ用")
    g["A1"] = "グラフの元データ（自動で作る。編集不要）。空白＝その線では使わない日"
    g.column_dimensions["A"].width = 12

    # ================= 表紙・結論 =================
    ws = Sheet(wb, "報告書", first=True).ws
    for r in (1, 2, 3):
        for c in range(1, NCOL + 1):
            ws.cell(r, c).fill = fill(NAVY)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=NCOL)
    x = ws.cell(2, 1, "  japangreen.vn　検索順位　期間報告書")
    x.font, x.alignment = F(bold=True, size=20, color="FFFFFF"), Alignment(vertical="center")
    ws.row_dimensions[1].height, ws.row_dimensions[2].height, ws.row_dimensions[3].height = 8, 40, 8
    r = 4
    r = text(ws, r, 1, NCOL, f"計測期間　{start} 〜 {end}（{len(days)} 日間）　　／　　作成日　{today}", size=11, bold=True, color=NAVY)
    r = text(ws, r, 1, NCOL, "この報告書は、当院のホームページ（japangreen.vn）が Google などの検索結果で何番目に出るかを毎朝自動で調べた記録を、"
             "SEO（検索で見つけてもらうための取り組み）に詳しくない方にも読めるようにまとめたものです。"
             "各章のはじめに「狙い（何を確かめるための表・グラフか）」、あとに「見方」「注意」「読み取れること」を書いています。", size=10)
    r = gap(ws, r)

    r = h2(ws, r, "はじめに読むところ ― 今回の結論")
    r = text(ws, r, 1, NCOL,
             f"■ {md(GOOGLE_ODD_FROM)} 以降の Google の数字は「要確認」です。順位が下がったとは判断しません。\n"
             f"{md(GOOGLE_ODD_FROM)} から、Google の自動計測で「当院のサイトが見つからない」という結果が急に増えました。"
             "しかし、同じ条件で実際に Google を開くと当院は 1〜2 位に出ていました（11/16 に 3 つの検索語で確認）。"
             "当院が消えた計測結果では、グループの公式サイト（シンガポールなど）も一緒に消えており、計測を頼んでいる会社が返す結果のほうに異常がある疑いが強い状態です。"
             "サーチコンソール（Google 公式の集計）での裏づけはまだ取れていません。", size=10, bg=WARN_BG, color=WARN_FG, indent=1) if odd_days else r
    r = gap(ws, r)

    r = text(ws, r, 1, NCOL, "1. 確認できたこと", size=11, bold=True, color=NAVY)
    facts = []
    if ok_days:
        facts.append(f"院名で探す人には、よく見つけてもらえています。「ジャパングリーン」を含む検索（指名語）は、{okp} の Google で {len(brand)} 通りのうち "
                     f"{rng(b30)} 通りが 30 位以内、そのうち {rng(b3)} 通りが 1〜3 位でした（平均 {b_avg:.1f} 位）。")
        facts.append(f"診療内容や地域で探す人には、まだ見つかりにくい状態です。「ハノイ 病院」「ハノイ 健康診断」などの検索（一般語）は、{okp} の Google で {len(general)} 通りのうち "
                     f"30 位以内が {rng(g30)} 通り、上位 10 件に入ったのは {rng(g10)} 通りでした。詳しくは 4 章。")
    facts.append(y_text)
    r = bullets(ws, r, facts)
    r = gap(ws, r)
    r = text(ws, r, 1, NCOL, "2. まだ判断できないこと", size=11, bold=True, color=NAVY)
    unk = []
    if odd_days:
        unk.append(f"{oddp} の Google の順位（指名語の 30 位以内は {rng(b30_odd)} 通りと記録されていますが、実際より低く出ている疑いがあります）。")
    unk.append(f"Bing（Edge）の順位。計測会社が別の検索語の結果を返すことがあり、毎日 {len(bkeys)} 通りのうち取得できたのは {rng(bing_got)} 通りでした。日ごとの上下は実際の動きとは限りません。")
    unk.append(f"ホームページの改善の効果。{len(days)} 日間という短い期間では、日々の揺れ・計測の異常・改善の効果を区別できません。")
    r = bullets(ws, r, unk)
    r = gap(ws, r)
    r = text(ws, r, 1, NCOL, "3. 次に確認すること", size=11, bold=True, color=NAVY)
    r = bullets(ws, r, [
        "サーチコンソールで、院名を含む検索の「掲載順位」を日別に見る（11/14 以降に落ちていなければ、計測側の異常と確定できます）。",
        "計測の異常が続くかどうかを数日見る。続く場合は、異常な結果を自動で見分けて取り直す仕組みを入れる／計測会社に報告する。",
        "一般語のうち「あと一歩」（11〜30 位）の語を、内容を充実させる候補として検討する（4 章の表）。"]
        + ([f"Yahoo! JAPAN の {md(end)} の急な上昇が、明日以降も続くかを見る。"] if y_jump else []))
    r = gap(ws, r)

    r = h2(ws, r, "数字で見る今回の期間")
    for i, h in enumerate(("項目", "数字", "説明")):
        c1, c2 = ((1, 1), (2, 4), (5, NCOL))[i]
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        ws.cell(r, c1, h)
    style_header(ws, r, 1, NCOL)
    r += 1
    table = [
        ("計測した日数", f"{len(days)} 日", f"{md(start)}〜{md(end)}。毎朝 1 回、{n_all} 通り（検索語 {n_kw} 語 × 地点・言語 × ブラウザ）を調べています。"),
        ("Google を判断に使える日", f"{len(ok_days)} 日", f"{okp}。{oddp} は要確認（上の結論を参照）。" if odd_days else "全期間。"),
        ("指名語・Google・30 位以内", f"{rng(b30)} / {len(brand)}", f"{okp}。院名を含む検索で当院が 30 位以内に出た通り数。"),
        ("指名語・Google・平均順位", f"{b_avg:.1f} 位" if b_avg else "—", f"{okp}。1 に近いほど良い。圏外は 31 位として計算。"),
        ("一般語・Google・30 位以内", f"{rng(g30)} / {len(general)}", f"{okp}。診療内容・地域の検索で当院が 30 位以内に出た通り数。"),
        ("一般語・Google・上位 10 件", f"{rng(g10)} / {len(general)}", f"{okp}。検索結果の最初のほうに出た通り数。"),
        ("Yahoo! JAPAN・30 位以内", (f"{rng(y_prev)} → {y30[-1]} / {len(ykeys)}" if y_jump else f"{rng(y30)} / {len(ykeys)}"),
         (f"{md(days[0])}〜{md(days[-2])} → {md(end)}。最終日に急増（1 日だけの記録）。" if y_jump else "全期間。") + "日本から日本語で検索した場合。"),
        ("Bing で取得できた数", f"{rng(bing_got)} / {len(bkeys)}", "全期間。取得できなかった分は順位が分かりません（圏外とは別）。"),
    ]
    for a, b_, c_ in table:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        _cell(ws, r, 1, a, bold=True, color=NAVY, align=LEFT)
        _cell(ws, r, 2, b_, size=12, bold=True, color=NAVY)
        for c in (3, 4):
            ws.cell(r, c).border = BOX
        text(ws, r, 5, NCOL, c_, size=9, valign="center", indent=1, min_h=24)
        for c in range(5, NCOL + 1):
            ws.cell(r, c).border = BOX
        r += 1
    r = gap(ws, r)
    r = h2(ws, r, "この報告書の構成")
    r = bullets(ws, r, [
        "1 章　SEO の基本と用語 ― はじめての方はここから",
        "2 章　全体の推移 ― 当院のサイトはどれくらい見つかるか",
        "3 章　院名で探す人（指名語）",
        "4 章　診療内容・地域で探す人（一般語）",
        "5 章　他院との比較",
        "6 章　条件別（地域・端末・検索エンジン）",
        "7 章　データの注意点と、言えること・言えないこと",
        "付録　検索語ごとの毎日の順位（全 " + str(n_all) + " 通り）"], size=10)

    # ================= 1 章 SEO の基本 =================
    ws = Sheet(wb, "1 SEOの基本").ws
    r = chapter(ws, "1 章", "SEO の基本と用語", "はじめての方向けに、この報告書を読むのに必要なことだけをまとめました。")
    r = h2(ws, r, "SEO とは")
    r = text(ws, r, 1, NCOL, "SEO（エス・イー・オー）は、Google などで検索した人に、自分たちのホームページを見つけてもらいやすくする取り組みです。"
             "広告とは違い、お金を払って上に出すのではなく、ページの内容や作りを良くして、検索結果の上のほうに出ることを目指します。")
    r = text(ws, r, 1, NCOL, "検索結果は、上にあるほどよく見られます。多くの人は最初の数件しか見ないので、1〜3 位と 11 位以下では、見てもらえる回数に大きな差が出ます。"
             "そこで「当院のページが何番目に出るか（順位）」を毎日記録し、変化を追いかけています。")
    r = gap(ws, r)
    r = h2(ws, r, "検索する人は 2 種類いる")
    r = text(ws, r, 1, NCOL, "① 院名で探す人（指名語）　例:「ジャパングリーンクリニック ハノイ」\n"
             "　すでに当院を知っていて、場所や電話番号、診療時間を調べたい人です。ここで 1 位付近に出ないと、知っている人にもたどり着いてもらえません。まず守るべき土台です。")
    r = text(ws, r, 1, NCOL, "② 診療内容や地域で探す人（一般語）　例:「ハノイ 健康診断」「ハノイ 日系クリニック」\n"
             "　まだ当院を知らず、どこに行くかを探している人です。ここで上位に出ると、新しい患者さんに知ってもらえます。他院と競争になるので、上げるのに時間がかかります。")
    r = gap(ws, r)
    r = h2(ws, r, "順位から来院までの流れと、この報告書で分かる範囲")
    r = text(ws, r, 1, NCOL, "検索結果に出る（順位）→ クリックしてホームページを見る（訪問）→ 予約・問い合わせ → 来院、という流れになります。"
             "この報告書で分かるのは、最初の「順位」だけです。訪問の数や予約の数は、サーチコンソールやアクセス解析など別の記録で見ます。順位から来院数を計算することはできません。")
    r = gap(ws, r)
    r = h2(ws, r, "どうやって調べているか")
    r = text(ws, r, 1, NCOL, f"毎朝 1 回、計測サービス（DataForSEO など）を使って、決めた検索語で実際に検索したときの結果を取り寄せ、当院のページが何番目にあるかを記録しています。"
             f"検索結果は「どこから」「何語の設定で」「どのブラウザ・端末で」検索するかで変わるので、次の組み合わせで調べています（1 日 {n_all} 通り）。")
    r = bullets(ws, r, [
        f"検索語: {n_kw} 語（日本語 {len(CFG['keywords']['ja'])} 語・英語 {len(CFG['keywords']['en'])} 語。うち指名語は日本語 4 語・英語 2 語）",
        "地点と言語: 日本から・ハノイから × 日本語・英語",
        "検索エンジンと端末: Google（パソコン・iPhone）、Bing（パソコンの Edge）、Yahoo! JAPAN（日本から・日本語のみ）",
        "調べる深さ: 30 位まで。30 位までに当院が無ければ「圏外」と書きます"])
    r = gap(ws, r)
    r = h2(ws, r, "用語")
    glossary = [
        ("検索語", "検索の窓に入れる言葉。キーワードとも言います。"),
        ("指名語", "院名を含む検索語。当院を知っている人が使います。"),
        ("一般語", "院名を含まない検索語（診療内容・地域など）。当院を知らない人が使います。"),
        ("検索エンジン", "検索の仕組みそのもの。Google・Bing・Yahoo! JAPAN など。ブラウザ（Chrome・Edge・Safari）ごとに最初に使われる検索エンジンが違います。"),
        ("順位", "通常の検索結果の中で、当院のページが上から何番目か。広告や地図の枠は数えません。数字が小さいほど良い。"),
        ("通り（組）", "「検索語 × 地点・言語 × ブラウザ」の 1 つの組み合わせ。この報告書の数字の多くは「何通りのうち何通り」で書いています。"),
        ("30 位以内／上位 10 件", "30 位以内＝3 ページ目くらいまでに出る。上位 10 件＝最初に目に入る範囲で、実際に見てもらえるのは主にここです。"),
        ("圏外", "30 位までに当院が見つからなかったこと。「サイトが消えた」という意味ではありません。"),
        ("取得失敗（失敗）", "その日は検索結果そのものを取れなかったこと。順位は不明で、圏外とは別です。"),
        ("要確認", "数字は取れたが、計測の異常が疑われるもの。良くなった・悪くなったの判断には使いません。"),
        ("平均順位", "複数の通り・複数の日の順位を平均したもの。圏外は 31 位として計算します。"),
        ("サーチコンソール", "Google が無料で提供する公式の集計。実際に検索した人に当院が何位で表示され、何回クリックされたかが分かります。"),
        ("地図の枠", "検索結果に出る Google マップの枠。通常の順位とは別の仕組みで、この報告書では扱いません（毎日のダッシュボードにあります）。"),
    ]
    ws.cell(r, 1, "用語")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=NCOL)
    ws.cell(r, 2, "意味")
    style_header(ws, r, 1, NCOL)
    r += 1
    for t_, d_ in glossary:
        _cell(ws, r, 1, t_, bold=True, color=NAVY, align=LEFT)
        text(ws, r, 2, NCOL, d_, size=9, valign="center", indent=1, min_h=20)
        for c in range(2, NCOL + 1):
            ws.cell(r, c).border = BOX
        r += 1

    # ================= 2 章 全体の推移 =================
    ws = Sheet(wb, "2 全体の推移").ws
    r = chapter(ws, "2 章", "全体の推移", "問い: 検索したとき、当院のサイトはどれくらい見つかるか。")
    r = h2(ws, r, f"Google で当院が見つかった通り数（{len(gkeys)} 通りのうち・日別）")
    r = note(ws, r, "狙い", "全部の検索語をまとめて、当院が検索結果に出ている広さをつかみます。細かい語ごとの話の前に、全体が増えているか減っているかを見るためのグラフです。")
    ws.add_chart(two_phase_chart(g, "Google の通り数", ["30 位以内", "上位 10 件", "1 位"], [NAVY, GREEN, "E0533D"], days, ok_days,
                                 [cnt(gkeys, 30), cnt(gkeys, 10), cnt(gkeys, 1)]), f"A{r}")
    r = reserve(ws, r, 7.5)
    r = note(ws, r, "見方", "横が日付、縦が通り数です。線が上にあるほど、多くの検索で当院が見つかっています。"
             "「1 位」は「上位 10 件」に、「上位 10 件」は「30 位以内」に含まれます（別々の数ではありません）。")
    if odd_days:
        r = note(ws, r, "注意", f"薄い点線（{oddp}）は「要確認」の数字です。{md(GOOGLE_ODD_FROM)} の落ち込みは、計測の異常による見かけの低下の疑いが強く、実際に順位が下がったとは判断しません。")
    r = gap(ws, r)

    r = h2(ws, r, "検索エンジン別の通り数（日別）")
    r = note(ws, r, "狙い", "Google だけでなく、Yahoo! JAPAN・Bing でも見つかるかを確かめます。日本から調べる人は Yahoo! JAPAN も多く使います。")
    r = day_header(ws, r, 1, days, odd_days, heads_before=("区分",))
    lines = [
        ("Google　30 位以内", lambda d: f"{cnt(gkeys)(d)}/{D.count(d, gkeys)[1]}", True),
        ("Google　上位 10 件", lambda d: f"{cnt(gkeys, 10)(d)}/{D.count(d, gkeys)[1]}", True),
        ("Google　1 位", lambda d: f"{cnt(gkeys, 1)(d)}/{D.count(d, gkeys)[1]}", True),
        ("　うち指名語　30 位以内", lambda d: f"{cnt(brand)(d)}/{D.count(d, brand)[1]}", True),
        ("　うち一般語　30 位以内", lambda d: f"{cnt(general)(d)}/{D.count(d, general)[1]}", True),
        ("Yahoo! JAPAN　30 位以内", lambda d: f"{cnt(ykeys)(d)}/{D.count(d, ykeys)[1]}", False),
        ("Yahoo! JAPAN　上位 10 件", lambda d: f"{cnt(ykeys, 10)(d)}/{D.count(d, ykeys)[1]}", False),
        ("Bing　30 位以内", lambda d: f"{cnt(bkeys)(d)}/{D.count(d, bkeys)[1]}", False),
        ("Bing　取得できた数", lambda d: f"{D.count(d, bkeys)[1]}/{len(bkeys)}", False),
    ]
    for lab, fn, is_g in lines:
        _cell(ws, r, 1, lab, bold=not lab.startswith("　"), color=NAVY if not lab.startswith("　") else "333333", align=LEFT)
        for i, d in enumerate(days):
            x = _cell(ws, r, 2 + i, fn(d), size=8)
            if is_g and d in odd_days:
                x.fill, x.font = fill(WARN_BG), F(size=8, color=WARN_FG)
        r += 1
    r = gap(ws, r, 4)
    r = note(ws, r, "見方", "各マスは「当院が見つかった通り数 / その日に結果を取れた通り数」です。例「44/104」＝104 通りのうち 44 通りで 30 位以内。"
             "右側の数（分母）が小さい日は、取得に失敗した通りが多い日です。")
    r = note(ws, r, "注意", "日付が橙色・マスが薄い橙色のところは Google の「要確認」の数字です。Bing は取得できた数が日によって大きく違うので、通り数の増減をそのまま比べられません。")
    reads = []
    if ok_days:
        reads.append(f"{okp} の Google は、{len(gkeys)} 通りのうち 30 位以内が {rng([cnt(gkeys)(d) for d in ok_days])} 通り、上位 10 件が {rng([cnt(gkeys, 10)(d) for d in ok_days])} 通りで、大きな変化はありませんでした。")
    reads.append(y_text)
    if odd_days:
        reads.append(f"{oddp} の Google の落ち込みは指名語に集中しています（{rng(b30)} → {rng(b30_odd)} 通り）。院名の検索で急に見つからなくなるのは不自然で、計測の異常を疑う理由の 1 つです。")
    r = note(ws, r, "読み取れること", "\n".join("・" + s_ for s_ in reads))

    # ================= 3 章 指名語 =================
    ws = Sheet(wb, "3 院名で探す人").ws
    r = chapter(ws, "3 章", "院名で探す人（指名語）", "問い: 当院を知っている人が院名で検索したとき、ちゃんと当院のページにたどり着けるか。")
    r = h2(ws, r, f"指名語で当院が見つかった通り数（Google・{len(brand)} 通りのうち）")
    r = note(ws, r, "狙い", "院名で検索した人が当院にたどり着けるかは、SEO のいちばんの土台です。ここが 1〜3 位にあるのが普通の状態で、下がったらすぐ原因を調べます。")
    ws.add_chart(two_phase_chart(g, "指名語の通り数（Google）", ["30 位以内", "1〜3 位", "1 位"], [NAVY, GREEN, "E0533D"], days, ok_days,
                                 [cnt(brand, 30), cnt(brand, 3), cnt(brand, 1)], h=7, ymax=len(brand)), f"A{r}")
    r = reserve(ws, r, 7)
    r = note(ws, r, "見方", f"縦の最大は {len(brand)} 通り（指名語 6 語 × 日本・ハノイ × パソコン・iPhone）。線が {len(brand)} に近いほど良い状態です。")
    if odd_days:
        r = note(ws, r, "注意", f"薄い点線（{oddp}）は要確認。この期間、計測結果の上位 10 件に当院もグループの公式サイトも 1 つも出ていない通りが "
                 f"1 日 {rng([odd_n[d] for d in odd_days])} 通りありました（{okp} は {rng([odd_n[d] for d in ok_days])} 通り）。実際の Google では起こりにくい結果です。")
    r = gap(ws, r)

    r = h2(ws, r, "指名語ごとの毎日の順位（Google）")
    r = color_legend(ws, r, bool(odd_days))
    r = gap(ws, r, 4)
    r = day_header(ws, r, 1, days, odd_days, heads_before=("検索語", "地点", "", "端末"), heads_after=("平均",))
    ws.merge_cells(start_row=r - 1, start_column=2, end_row=r - 1, end_column=3)
    for l in CFG["locales"]:
        for k in CFG["keywords"][l["keyword_set"]]:
            if k["group"] != "brand":
                continue
            for p in GOOGLE:
                key = (p, l["id"], k["kw"])
                _cell(ws, r, 1, k["kw"], size=8, align=LEFT)
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
                _cell(ws, r, 2, l["label"], size=8, color="555B66")
                ws.cell(r, 3).border = BOX
                _cell(ws, r, 4, "PC" if p == "chrome-pc" else "iPhone", size=8, color="555B66")
                for i, d in enumerate(days):
                    rank_cell(ws.cell(r, 5 + i), D.rank(d, key))
                avg_cell(ws.cell(r, 5 + len(days)), ok_avg([key]))
                ws.row_dimensions[r].height = 16
                r += 1
    r = gap(ws, r, 4)
    r = note(ws, r, "見方", f"横に読むと、その検索語の毎日の順位が分かります。緑が濃いほど上位。右端の「平均」は {okp}（判断に使える日）だけの平均です。")
    r = note(ws, r, "注意", "Google の順位は、同じ語でも日によって 1 位になったり 10 位台になったりします。1 日だけの数字で良し悪しを決めないでください。")
    if ok_days:
        weak = []
        for l in CFG["locales"]:
            for k in CFG["keywords"][l["keyword_set"]]:
                if k["group"] == "brand":
                    v = ok_avg(keys_where(GOOGLE, (l["id"],), kws=(k["kw"],)))
                    if v is not None and v > 5:
                        weak.append((v, f"「{k['kw']}」（{l['label']}・平均 {v:.1f} 位）"))
        reads = [f"{okp} は {len(brand)} 通りのうち {rng(b30)} 通りが 30 位以内、{rng(b3)} 通りが 1〜3 位で、平均 {b_avg:.1f} 位。院名での検索は、おおむね当院にたどり着ける状態です。"]
        if weak:
            reads.append("ただし、次の語は平均が 5 位より下で、院名の検索としては弱めです: " + "、".join(s_ for _, s_ in sorted(weak, reverse=True)[:5]) + "。")
        r = note(ws, r, "読み取れること", "\n".join("・" + s_ for s_ in reads))

    # ================= 4 章 一般語 =================
    ws = Sheet(wb, "4 診療内容で探す人").ws
    r = chapter(ws, "4 章", "診療内容・地域で探す人（一般語）", "問い: 当院をまだ知らない人が、診療内容や地域で検索したとき、当院を見つけてもらえるか。")
    r = h2(ws, r, f"一般語で当院が見つかった通り数（Google・{len(general)} 通りのうち）")
    r = note(ws, r, "狙い", "新しい患者さんに知ってもらえるかは、一般語の順位で決まります。ここを少しずつ増やすのが SEO の本番です。")
    ws.add_chart(two_phase_chart(g, "一般語の通り数（Google）", ["30 位以内", "上位 10 件"], [NAVY, GREEN], days, ok_days,
                                 [cnt(general, 30), cnt(general, 10)], h=6.5), f"A{r}")
    r = reserve(ws, r, 6.5)
    r = note(ws, r, "見方", f"縦は通り数（最大 {len(general)}＝一般語 20 語 × 日本・ハノイ × パソコン・iPhone）。上位 10 件の線が、実際に見てもらえる範囲です。")
    if odd_days:
        r = note(ws, r, "注意", f"薄い点線（{oddp}）は要確認。一般語にも計測の異常が混じっている可能性がありますが、どの通りが異常かは見分けられていません。")
    r = gap(ws, r)

    r = h2(ws, r, "一般語ごとの現在地（Google・パソコンと iPhone の平均）")
    r = note(ws, r, "狙い", "どの検索語が強く、どの語が「あと一歩」で、どの語がまだ出ていないかを分けます。力を入れる語を選ぶための表です。")
    heads = [("検索語", 1, 1), ("地点", 2, 3), (f"平均順位\n{okp}", 4, 5), ("現在地", 6, 7), ("30 位以内\nだった回数", 8, 9),
             (f"参考: 計測値\n{oddp}", 10, 11), (f"Yahoo!\n{md(end)}", 12, 13)]
    for h, c1, c2 in heads:
        if c2 > c1:
            ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        ws.cell(r, c1, h)
    style_header(ws, r, 1, NCOL)
    ws.cell(r, 10).fill = ws.cell(r, 11).fill = fill(ORANGE)
    ws.row_dimensions[r].height = 44
    r += 1
    cls_count = {"上位（1〜10 位）": [], "あと一歩（11〜30 位）": [], "圏外": []}
    for l in CFG["locales"]:
        items = []
        for k in CFG["keywords"][l["keyword_set"]]:
            if k["group"] != "general":
                continue
            ks = keys_where(GOOGLE, (l["id"],), kws=(k["kw"],))
            items.append((ok_avg(ks) if ok_avg(ks) is not None else 99, k["kw"], ks))
        for v, kw, ks in sorted(items):
            hits = sum(1 for d in ok_days for n in D.vals(d, ks) if n is not None)
            tot = sum(len(D.vals(d, ks)) for d in ok_days)
            cls = "上位（1〜10 位）" if v <= 10 else "あと一歩（11〜30 位）" if v <= 30 else "圏外"
            if v != 99:
                cls_count[cls].append(f"{kw}（{l['label']}）")
            _cell(ws, r, 1, kw, size=8, align=LEFT)
            for c1, c2 in ((2, 3), (4, 5), (6, 7), (8, 9), (10, 11), (12, 13)):
                ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
                ws.cell(r, c2).border = BOX
            _cell(ws, r, 2, l["label"], size=8, color="555B66")
            avg_cell(ws.cell(r, 4), None if v == 99 else v)
            x = _cell(ws, r, 6, cls.split("（")[0] if v != 99 else "—", size=8, bold=v <= 30)
            if v <= 10:
                x.fill, x.font = fill(BANDS[0][1]), F(size=8, bold=True, color=BANDS[0][2])
            elif v <= 30:
                x.fill, x.font = fill(BANDS[2][1]), F(size=8, bold=True, color=BANDS[2][2])
            _cell(ws, r, 8, f"{hits} / {tot}" if tot else "—", size=9)
            x = ws.cell(r, 10)
            avg_cell(x, D.avg(odd_days, ks) if odd_days else None)
            if x.value == "圏外" or x.value == "—":
                pass
            else:
                x.fill, x.font = fill(WARN_BG), F(size=9, color=WARN_FG)
            x = ws.cell(r, 12)
            if "yahoo-pc" in l["profiles"]:
                rank_cell(x, D.rank(end, ("yahoo-pc", l["id"], kw)))
            else:
                x.value, x.font, x.alignment, x.border = "対象外", F(size=8, color="B0B4BB"), CENTER, BOX
            ws.row_dimensions[r].height = 16
            r += 1
    r = gap(ws, r, 4)
    r = note(ws, r, "見方", f"「平均順位」は {okp} の Google（パソコン・iPhone）の平均で、圏外は 31 位として計算しています。"
             "「現在地」は、その平均が 1〜10 位なら「上位」、11〜30 位なら「あと一歩」、それより下なら「圏外」。"
             "「30 位以内だった回数」は、その期間に調べた回数（日数 × 2 端末）のうち何回 30 位以内だったか。平均が良くても回数が少ない語は、たまたま出ただけかもしれません。"
             f"右端の「Yahoo!」は、最終日（{md(end)}）の Yahoo! JAPAN での順位です（日本から日本語で検索した場合だけ調べています）。")
    r = note(ws, r, "注意", f"橙色の「参考: 計測値」は {oddp} の要確認の数字です。左の平均と比べて良くなった・悪くなったと読まないでください。")
    top_, near, out_ = cls_count["上位（1〜10 位）"], cls_count["あと一歩（11〜30 位）"], cls_count["圏外"]
    reads = [f"上位（平均 1〜10 位）は {len(top_)} 件" + (": " + "、".join(top_[:8]) + (" ほか" if len(top_) > 8 else "") if top_ else "") + "。",
             f"あと一歩（平均 11〜30 位）は {len(near)} 件" + (": " + "、".join(near[:10]) + (" ほか" if len(near) > 10 else "") if near else "") + "。ページの内容を充実させる候補です。",
             f"圏外（平均が 30 位より下）は {len(out_)} 件。その語に対応するページが無い・弱い可能性があります。"]
    r = note(ws, r, "読み取れること", "\n".join("・" + s_ for s_ in reads))

    # ================= 5 章 競合 =================
    ws = Sheet(wb, "5 他院との比較").ws
    comps = CFG.get("competitors", [])
    cdays = [d for d in days if d >= COMP_FROM]
    c_ok = [d for d in cdays if d in ok_days]
    names = [("当院", "rank")] + [(x["label"] if len(x["label"]) < 8 else x["label"].split(".")[0], "rank_" + x["id"]) for x in comps]
    r = chapter(ws, "5 章", "他院との比較", "問い: 同じ検索語で、ハノイの他の日系クリニックと比べて当院はどの位置にいるか。")

    def cval(d, key, col):
        rec = D.picks[d].get(key)
        if not rec or rec.get("status") != "ok":
            return False
        v = (rec.get(col) or "").strip()
        return int(float(v)) if v else None

    pc_general = keys_where(("chrome-pc",), group="general")
    r = h2(ws, r, f"一般語で上位に入った語の数（Google・パソコン・{len(pc_general)} 通りのうち）")
    r = note(ws, r, "狙い", "当院の順位だけでは、良いのか悪いのか分かりません。同じ検索語で他院が何通り上位にいるかを並べて、当院の立ち位置をつかみます。院名の検索は比べても意味がないので、一般語だけで比べます。")
    ws.cell(r, 1, "クリニック")
    c = 2
    for d in cdays:
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 1)
        ws.cell(r, c, md(d))
        c += 2
    style_header(ws, r, 1, c - 1)
    for i, d in enumerate(cdays):
        if d in odd_days:
            ws.cell(r, 2 + 2 * i).fill = ws.cell(r, 3 + 2 * i).fill = fill(ORANGE)
    r += 1
    c = 2
    for d in cdays:
        ws.cell(r, c, "30 位内")
        ws.cell(r, c + 1, "10 件内")
        c += 2
    style_header(ws, r, 1, c - 1)
    ws.row_dimensions[r].height = 28
    r += 1
    comp_tot = {}
    for nm, col in names:
        own = nm == "当院"
        _cell(ws, r, 1, nm, bold=own, color=GREEN if own else "222222", align=LEFT)
        c = 2
        for d in cdays:
            v = [n for n in (cval(d, key, col) for key in pc_general) if n is not False and n is not None]
            comp_tot[(nm, d)] = (len(v), sum(n <= 10 for n in v))
            for n in comp_tot[(nm, d)]:
                x = _cell(ws, r, c, n, bold=own)
                if d in odd_days:
                    x.fill, x.font = fill(WARN_BG), F(size=9, bold=own, color=WARN_FG)
                c += 1
        r += 1
    r = gap(ws, r, 4)
    r = note(ws, r, "見方", "「30 位内」＝30 位以内に出た通り数、「10 件内」＝上位 10 件に出た通り数。数が大きいほど、多くの検索語で見つかっています。")
    r = note(ws, r, "注意", f"他院の順位は {md(COMP_FROM)} から取り始めたので、判断に使える日は {('・'.join(md(d) for d in c_ok)) or 'ありません'}"
             + (f" の {len(c_ok)} 日だけです" if c_ok else "") + "。橙色の日は要確認で、他院の数字にも計測の異常が混じっている可能性があります。この表だけで優劣を決めないでください。")
    r = gap(ws, r)

    base_day = c_ok[-1] if c_ok else (cdays[0] if cdays else None)
    if base_day:
        r = h2(ws, r, f"一般語ごとの順位（Google・パソコン・{md(base_day)}）")
        r = note(ws, r, "狙い", "どの検索語で、どの院が上にいるかを見ます。当院が圏外で他院が上位の語は、他院にあって当院に無いページがある可能性があります。")
        for l in CFG["locales"]:
            ws.cell(r, 1, "検索語（" + l["label"] + "）")
            c = 2
            for nm, _ in names:
                ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 1)
                ws.cell(r, c, nm)
                c += 2
            style_header(ws, r, 1, c - 1)
            r += 1
            for k in CFG["keywords"][l["keyword_set"]]:
                if k["group"] == "brand":
                    continue
                _cell(ws, r, 1, k["kw"], size=8, align=LEFT)
                c = 2
                for nm, col in names:
                    ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 1)
                    rank_cell(ws.cell(r, c), cval(base_day, ("chrome-pc", l["id"], k["kw"]), col))
                    ws.cell(r, c + 1).border = BOX
                    c += 2
                ws.row_dimensions[r].height = 16
                r += 1
            r = gap(ws, r, 6)
        r = note(ws, r, "見方", "横に読むと、その検索語での各院の順位が分かります。緑が濃いほど上位、灰色の「圏外」は 30 位までに無し。")
        r = note(ws, r, "注意", f"{md(base_day)} の 1 日だけの記録です。順位は日によって動くので、傾向として読むには日数が足りません。")
        own30, own10 = comp_tot[("当院", base_day)]
        others = "、".join(f"{nm} {comp_tot[(nm, base_day)][1]}" for nm, _ in names[1:])
        r = note(ws, r, "読み取れること", f"・{md(base_day)} の時点で、上位 10 件に入った一般語は当院 {own10} 通り（30 位以内は {own30} 通り）。他院は {others} 通りでした。\n"
                 "・これは 1 日だけの記録なので、「この条件・この日にこの順位だった」という事実にとどまります。")

    # ================= 6 章 条件別 =================
    ws = Sheet(wb, "6 条件別").ws
    r = chapter(ws, "6 章", "条件別（地域・端末・検索エンジン）", "問い: どこから・何で検索する人に、当院は見つかりやすいか。")
    r = h2(ws, r, "条件ごとの「30 位以内の通り数」（日別）")
    r = note(ws, r, "狙い", "患者さんには、ハノイに住んでいる人も、赴任前に日本から調べる人もいます。パソコンの人もスマホの人もいます。どの条件で弱いかが分かると、手を入れる場所を絞れます。")
    ja, en = ("JP-ja", "VN-ja"), ("JP-en", "VN-en")
    groups = [
        ("地域（Google）", [("日本から検索", keys_where(GOOGLE, ("JP-ja", "JP-en")), True), ("ハノイから検索", keys_where(GOOGLE, ("VN-ja", "VN-en")), True)]),
        ("端末（Google）", [("パソコン", keys_where(("chrome-pc",)), True), ("スマホ（iPhone）", keys_where(("iphone",)), True)]),
        ("検索語の言語（Google）", [("日本語の検索語", keys_where(GOOGLE, ja), True), ("英語の検索語", keys_where(GOOGLE, en), True)]),
        ("検索エンジン（日本から・日本語・パソコン）", [("Google", keys_where(("chrome-pc",), ("JP-ja",)), True),
                                                    ("Yahoo! JAPAN", keys_where(("yahoo-pc",), ("JP-ja",)), False),
                                                    ("Bing", keys_where(("edge-pc",), ("JP-ja",)), False)]),
    ]
    r = day_header(ws, r, 1, days, odd_days, heads_before=("条件",), heads_after=("平均順位", ""))
    ws.merge_cells(start_row=r - 1, start_column=2 + len(days), end_row=r - 1, end_column=3 + len(days))
    cond = {}
    for gname, items in groups:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3 + len(days))
        x = ws.cell(r, 1, gname)
        x.font, x.fill = F(bold=True, size=9, color=NAVY), fill(SOFT)
        for c in range(1, 4 + len(days)):
            ws.cell(r, c).fill, ws.cell(r, c).border = fill(SOFT), BOX
        r += 1
        for nm, ks, is_g in items:
            _cell(ws, r, 1, f"　{nm}（{len(ks)} 通り）", size=9, align=LEFT)
            for i, d in enumerate(days):
                n, got = D.count(d, ks)
                x = _cell(ws, r, 2 + i, f"{n}/{got}" if got else "—", size=9)
                if is_g and d in odd_days:
                    x.fill, x.font = fill(WARN_BG), F(size=9, color=WARN_FG)
            v = D.avg(ok_days if is_g else days, ks)
            cond[nm] = v
            ws.merge_cells(start_row=r, start_column=2 + len(days), end_row=r, end_column=3 + len(days))
            x = _cell(ws, r, 2 + len(days), f"{v:.1f} 位" if v is not None else "—", size=9, bold=True)
            ws.cell(r, 3 + len(days)).border = BOX
            r += 1
    r = gap(ws, r, 4)
    r = note(ws, r, "見方", f"各マスは「30 位以内の通り数 / 取得できた通り数」。右端の「平均順位」は、Google は {okp}（判断に使える日）、Yahoo! JAPAN・Bing は全期間の平均で、小さいほど良い（圏外は 31 位）。"
             "条件によって通り数（かっこ内）が違うので、マスの数どうしではなく、割合や平均順位で比べてください。")
    r = note(ws, r, "注意", "日付が橙色の日の Google は要確認です。Bing は取得できた通りだけの数字で、取得できた通りの顔ぶれが毎日違うため、日ごとの比較には向きません。")
    reads = []

    def cmp_(a, b):
        va, vb = cond.get(a), cond.get(b)
        if va is None or vb is None:
            return None
        better = a if va < vb else b
        return f"{a} {va:.1f} 位・{b} {vb:.1f} 位（{better}のほうが上位）"
    for lab, a, b in (("地域", "日本から検索", "ハノイから検索"), ("端末", "パソコン", "スマホ（iPhone）"), ("言語", "日本語の検索語", "英語の検索語")):
        s_ = cmp_(a, b)
        if s_:
            reads.append(f"{lab}: {s_}。")
    if cond.get("Google") is not None and cond.get("Yahoo! JAPAN") is not None:
        reads.append(f"検索エンジン（日本から・日本語）: Google {cond['Google']:.1f} 位・Yahoo! JAPAN {cond['Yahoo! JAPAN']:.1f} 位。")
    reads.append("差が小さい条件は、日々の揺れの範囲かもしれません。大きな差がある条件だけを手がかりにしてください。")
    r = note(ws, r, "読み取れること", "\n".join("・" + s_ for s_ in reads))

    # ================= 7 章 データの注意 =================
    ws = Sheet(wb, "7 データの注意").ws
    r = chapter(ws, "7 章", "データの注意点と、言えること・言えないこと", "数字を読む前に知っておくこと。自動計測には取りこぼしや異常があります。")
    r = h2(ws, r, "その日に結果を取れた通り数")
    r = note(ws, r, "狙い", "毎日ちゃんと計測できているかを確かめます。取れなかった通りが多い日は、その日の数字をそのまま信じられません。")
    r = day_header(ws, r, 1, days, odd_days, heads_before=("検索エンジン",))
    for lab, ks in (("Google（パソコン）", keys_where(("chrome-pc",))), ("Google（iPhone）", keys_where(("iphone",))),
                    ("Bing（Edge）", bkeys), ("Yahoo! JAPAN", ykeys)):
        _cell(ws, r, 1, f"{lab}　全 {len(ks)} 通り", size=9, align=LEFT)
        for i, d in enumerate(days):
            got = D.count(d, ks)[1]
            x = _cell(ws, r, 2 + i, got, size=9, bold=got < len(ks))
            if got < len(ks) * 0.8:
                x.fill, x.font = fill(FAIL_BG), F(size=9, bold=True, color=FAIL_FG)
        r += 1
    _cell(ws, r, 1, f"Google の指名語で異常が疑われる数（{len(brand)} 通り中）", size=8, align=LEFT)
    for i, d in enumerate(days):
        x = _cell(ws, r, 2 + i, odd_n[d], size=9, bold=odd_n[d] >= 5)
        if odd_n[d] >= 5:
            x.fill, x.font = fill(WARN_BG), F(size=9, bold=True, color=WARN_FG)
    ws.row_dimensions[r].height = 26
    r += 1
    r = gap(ws, r, 4)
    r = note(ws, r, "見方", "上の 4 行は、その日に検索結果を取れた通り数です（全部取れていれば左の「全○通り」と同じ数）。赤いマスは 2 割以上取れなかった日。"
             "いちばん下の行は、Google の指名語のうち「上位 10 件に当院もグループの公式サイトも出ていない」通りの数で、多い日は計測の異常が疑われます。")
    r = gap(ws, r)

    r = h2(ws, r, "注意 1 ― Google の計測結果の異常（" + (md(GOOGLE_ODD_FROM) + " 以降" if odd_days else "この期間は該当なし") + "）")
    if odd_days:
        r = bullets(ws, r, [
            f"起きていること: {md(GOOGLE_ODD_FROM)} から、指名語の計測結果で当院が圏外になる通りが急に増えました（30 位以内 {rng(b30)} → {rng(b30_odd)} 通り）。",
            "実際の検索では: 11/16 に、計測で圏外だった 3 つの検索語を同じ条件で実際の Google で開いたところ、当院は 1〜2 位に出ていました。ただし確認したのは 3 例で、全部の条件を保証するものではありません。",
            "異常を疑う理由: 当院が消えた結果には、いつも出ているグループの公式サイトも出ていません。代わりに SNS のページが上位に並ぶなど、結果全体がふだんと違います。",
            "ホームページ側: 検索に出さない設定（noindex）や、Google の巡回への応答に問題は見つかりませんでした（11/16 確認）。",
            "この報告書での扱い: 数字は消さずに載せ、「要確認」として色と線の種類を分けました。良くなった・悪くなったの判断には使っていません。",
            "確定するには: サーチコンソールで院名の検索の掲載順位を日別に見ます（未実施）。"])
    r = gap(ws, r)
    r = h2(ws, r, "注意 2 ― Bing（Edge）の取りこぼし")
    r = bullets(ws, r, [
        "計測会社が、頼んだ検索語とは別の検索語の結果を返すことがあります。11/14 からは、そういう結果を見分けて「取得失敗」として扱い、圏外には数えていません。",
        f"この期間、Bing は {len(bkeys)} 通りのうち取得できたのが 1 日 {rng(bing_got)} 通りでした。11/13 までは見分ける仕組みが無かったので、圏外の中に別の語の結果が混じっています。",
        "そのため Bing は「取得できた通りの中でどうだったか」しか言えず、日ごとの増減は実際の順位の動きとは限りません。"])
    r = gap(ws, r)
    r = h2(ws, r, "注意 3 ― 順位はもともと日によって動く")
    r = bullets(ws, r, [
        "Google の順位は、同じ検索語・同じ条件でも日によって変わります（1 位の翌日に 10 位台、など）。1 日だけの上下ではなく、数日の平均や回数で見てください。",
        "検索結果は、検索する人の場所・過去の検索・端末によっても変わります。この計測は「決めた条件で調べたらこうだった」という記録で、すべての人に同じ順位で見えているわけではありません。"])
    r = gap(ws, r)
    r = h2(ws, r, f"この {len(days)} 日間の記録から、言えること・言えないこと")
    r = text(ws, r, 1, NCOL, "言えること", size=10, bold=True, color=GREEN)
    r = bullets(ws, r, [
        "「この条件で・この日に調べたら・この順位だった」という事実。",
        "院名での検索は、判断に使える日には、おおむね当院にたどり着ける状態だったこと。",
        "一般語には、上位の語・あと一歩の語・圏外の語があり、どれがどれか（4 章の表）。",
        "計測の結果に、確認が必要な異常があること。"])
    r = text(ws, r, 1, NCOL, "言えないこと", size=10, bold=True, color=ORANGE)
    r = bullets(ws, r, [
        "ホームページの改善が成功した・失敗した（期間が短く、日々の揺れと区別できません）。",
        "他院に患者さんを取られている・勝っている（順位は来院数ではありません）。",
        "この先も上がる・下がる（傾向を言うには、少なくとも数週間の記録が必要です）。"])

    # ================= 付録 =================
    ncol_a = 3 + len(days) + 1
    ws = Sheet(wb, "付録 毎日の順位", ncol=ncol_a, landscape=True, widths=[30, 8, 28] + [8] * len(days) + [9]).ws
    r = chapter(ws, "付録", f"検索語ごとの毎日の順位（全 {n_all} 通り）",
                f"期間 {period}。順位は通常の検索結果での当院の位置。右端の「平均」は、Google は {okp}、Yahoo! JAPAN・Bing は全期間の平均（圏外は 31 位、取得失敗の日は除く）。")
    r = color_legend(ws, r, bool(odd_days))
    r = text(ws, r, 1, ncol_a, "※ 日付が橙色の日は、Google（Chrome・Safari）の数字が要確認です。Edge（Bing）の「失敗」は、結果を取れなかった・別の検索語の結果が返ってきた日です。", size=8, color=GRAY)
    hr = r
    r = day_header(ws, r, 1, days, odd_days, heads_before=("検索語", "種類", "地点・ブラウザ"), heads_after=("平均",))
    for l in CFG["locales"]:
        for k in CFG["keywords"][l["keyword_set"]]:
            for p in l["profiles"]:
                key = (p, l["id"], k["kw"])
                _cell(ws, r, 1, k["kw"], align=LEFT)
                _cell(ws, r, 2, "指名語" if k["group"] == "brand" else "一般語", size=8, color="555B66")
                _cell(ws, r, 3, short(l["id"], p), size=8, color="555B66", align=LEFT)
                for i, d in enumerate(days):
                    rank_cell(ws.cell(r, 4 + i), D.rank(d, key))
                avg_cell(ws.cell(r, 4 + len(days)), D.avg(ok_days if p in GOOGLE else days, [key]))
                ws.row_dimensions[r].height = 16
                r += 1
    ws.freeze_panes = ws.cell(hr + 1, 4)
    ws.print_title_rows = f"{hr}:{hr}"
    ws.auto_filter.ref = f"A{hr}:{get_column_letter(ncol_a)}{r - 1}"

    for name in wb.sheetnames:
        if name not in ("グラフ用", "付録 毎日の順位"):
            paginate(wb[name])
    wb.move_sheet("グラフ用", offset=len(wb.sheetnames))
    wb.active = 0
    wb.save(out)
    print("saved", out, "days", len(days), "ok_days", okp, "odd_days", oddp)
    print("brand in30", b30, b30_odd, "top3", b3, "avg", b_avg and round(b_avg, 2), "| general in30", g30, g30_odd, "top10", g10,
          "avg", g_avg and round(g_avg, 2))
    print("yahoo in30", y30, "top10", y10, "| bing got", bing_got, "| odd", list(odd_n.values()))
    print("classes", {k: len(v) for k, v in cls_count.items()}, "| cond", {k: v and round(v, 1) for k, v in cond.items()})
    print("PROFILE", len(PROFILE_LABEL))


if __name__ == "__main__":
    main(*sys.argv[1:])
