"""手動の順位記録表（xlsx）を作る。Google スプレッドシートに「インポート→新しいシートを挿入」で取り込む前提。"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

OUT = sys.argv[1]

BROWSERS = [  # 表示名, 検索エンジン, 端末
    ("Chrome（PC）", "Google", "PC"),
    ("Edge（PC）", "Bing", "PC"),
    ("Safari（iPhone）", "Google", "iPhone"),
]
OPTIONAL = [("Chrome（Android）", "Google", "Android"), ("Yahoo! JAPAN アプリ", "Yahoo! JAPAN", "iPhone")]
KEYWORDS = [
    ("ジャパングリーンクリニック ハノイ", "ブランド"),
    ("ジャパングリーン ベトナム クリニック", "ブランド"),
    ("Japan Green Vietnam Clinic", "ブランド"),
    ("ハノイ 日系クリニック", "一般"),
    ("ハノイ 日本人医師", "一般"),
    ("ハノイ 病院", "一般"),
    ("ハノイ 健康診断", "一般"),
    ("ハノイ ワークパーミット 健康診断", "一般"),
    ("ハノイ 歯医者 日本人", "一般"),
    ("Japanese clinic Hanoi", "一般"),
]
HEAD = ["日付", "場所", "ブラウザ", "検索エンジン", "端末", "状態", "検索語", "種類",
        "順位（1〜30・圏外は空欄）", "地図の枠の順位（1〜3）", "AIの概要", "当院のページ（URL）",
        "上位3件（サイト名）", "スクショのファイル名", "担当", "メモ"]
WIDTH = [12, 8, 18, 12, 8, 11, 34, 8, 14, 12, 11, 30, 30, 30, 8, 30]

HFILL = PatternFill("solid", fgColor="2F4078")
HFONT = Font(bold=True, color="FFFFFF")
INFILL = PatternFill("solid", fgColor="FFF6D5")  # 入力する列
thin = Side(style="thin", color="BBBBBB")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def rows_block():
    for b, eng, dev in BROWSERS:
        for kw, kind in KEYWORDS:
            yield [None, "ハノイ", b, eng, dev, "シークレット", kw, kind] + [None] * 8


def header(ws):
    ws.append(HEAD)
    for i, w in enumerate(WIDTH, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
        c = ws.cell(1, i)
        c.fill, c.font, c.border = HFILL, HFONT, BOX
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "A2"


def body(ws, n_rows):
    for r in range(2, 2 + n_rows):
        for c in range(1, len(HEAD) + 1):
            cell = ws.cell(r, c)
            cell.border = BOX
            if c in (1, 9, 10, 11, 12, 13, 14, 15, 16):
                cell.fill = INFILL
        ws.cell(r, 1).number_format = "yyyy-mm-dd"


def validations(ws, last):
    def add(dv, col):
        dv.error, dv.errorTitle = "一覧から選ぶか、範囲内の数字を入れてください", "入力の確認"
        dv.showErrorMessage = True
        ws.add_data_validation(dv)
        dv.add(f"{col}2:{col}{last}")
    add(DataValidation(type="list", formula1='"ハノイ,日本"', allow_blank=True), "B")
    names = ",".join(b for b, _, _ in BROWSERS + OPTIONAL)
    add(DataValidation(type="list", formula1=f'"{names}"', allow_blank=True), "C")
    add(DataValidation(type="list", formula1='"シークレット,普段"', allow_blank=True), "F")
    add(DataValidation(type="list", formula1='"ブランド,一般"', allow_blank=True), "H")
    add(DataValidation(type="whole", operator="between", formula1="1", formula2="30", allow_blank=True), "I")
    add(DataValidation(type="whole", operator="between", formula1="1", formula2="3", allow_blank=True), "J")
    add(DataValidation(type="list", formula1='"当院を引用,当院なし,概要なし"', allow_blank=True), "K")
    add(DataValidation(type="date", operator="greaterThan", formula1="46000", allow_blank=True), "A")


wb = Workbook()

# --- 使い方 ---
u = wb.active
u.title = "使い方"
u.column_dimensions["A"].width = 110
LINES = [
    ("検索順位の手動記録表（ジャパングリーン ベトナム クリニック）", "title"),
    ("実際のブラウザでの見え方を確かめるときに使います（毎日の自動計測との照合、「ブラウザで違う」という報告や指示があったとき）。ハノイから 3 つのブラウザ × 10 語＝30 回検索して記録します。所要 30〜45 分。Cốc Cốc は集計の対象外です（2026-11-05 の決定）。", ""),
    ("", ""),
    ("■ 手順", "h"),
    ("1. 「ひな形」シートの 2〜31 行目（30 行）をコピーし、「記録」シートのいちばん下（空いている最初の行）に貼り付ける。", ""),
    ("2. 貼り付けた 30 行の A 列（日付）に今日の日付を入れる（1 つ入れて下へコピー）。", ""),
    ("3. ブラウザごとにシークレットウィンドウを開く（Chrome＝シークレット、Edge＝InPrivate、iPhone の Safari＝プライベート）。ログインはしない。", ""),
    ("4. 10 語を順に検索し、黄色の列を埋める。", ""),
    ("   ・I 列「順位」: ふつうの検索結果だけを上から数えた、当院（japangreen.vn）の順位。30 位までに無ければ空欄（＝圏外）。", ""),
    ("   ・J 列「地図の枠の順位」: 地図と施設の一覧に当院があれば、その中の順番（1〜3）。無ければ空欄。", ""),
    ("   ・K 列「AIの概要」: いちばん上の AI の回答の引用元に当院のサイトがあれば「当院を引用」、無ければ「当院なし」、AI の回答が出なければ「概要なし」。", ""),
    ("   ・L〜P 列: 当院のページの URL、上位 3 件のサイト名（例: DYM、大使館、ロータス）、スクショのファイル名、担当、気づいたこと。", ""),
    ("5. スクリーンショット: 1 語につき 1 枚（当院が出た画面、または 1 ページ目）を OneDrive\\japangreen-HP\\順位確認\\YYYY-MM\\ に保存。", ""),
    ("   ファイル名の例: 2026-11-10_Edge_ハノイ病院.png", ""),
    ("6. 「集計」シートで今回と前回の順位を確認する（日付は自動で最新と 1 つ前を選ぶ）。", ""),
    ("", ""),
    ("■ 数え方（全ブラウザ共通）", "h"),
    ("・数えない: 広告（「スポンサー」「広告」）、地図の枠、AI による概要、「他の人はこちらも質問」、画像・動画のまとまり。", ""),
    ("・当院のサイトリンク（下に並ぶ小見出し）は親の 1 件として数える。", ""),
    ("・当院の Instagram・Facebook は公式サイトではないので順位に入れない（上位にあれば M 列に書く）。", ""),
    ("・スマホは下へ続けて読み込まれるので、30 件目まで数える。", ""),
    ("・検索エンジンは画面左上のロゴで確かめる（Edge が Google になっている等、既定が変えられていることがある）。", ""),
    ("", ""),
    ("■ 任意", "h"),
    ("・Chrome（Android）、Yahoo! JAPAN アプリでも測る場合は、「記録」に行を足して C 列で選ぶ。", ""),
    ("・普段の状態（ログインしたまま）での見え方も残す場合は、F 列を「普段」にした行を足す（集計はシークレットの行だけを使う）。", ""),
    ("・日本からの見え方は、ハノイからは再現できない → 月 1 回、Google サーチコンソールと Bing Webmaster Tools で国＝日本の平均順位を見る。", ""),
    ("", ""),
    ("■ 月 1 回: 無料の公式ツール", "h"),
    ("・Google サーチコンソール → 検索パフォーマンス → 期間「過去 28 日間」→ ＋新規 → 国（日本／ベトナム）→「クエリ」で 10 語の平均掲載順位・表示回数・クリック数を見る。", ""),
    ("・Bing Webmaster Tools → Search Performance → 検索語ごとの平均順位を見る（Edge の利用者の分。国で絞り込める場合は日本・ベトナム別に）。", ""),
    ("・平均の順位なので、手作業で数えた順位とはずれる。当院が表示されなかった語は出てこない。", ""),
]
for text, kind in LINES:
    u.append([text])
    c = u.cell(u.max_row, 1)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    if kind == "title":
        c.font = Font(bold=True, size=14)
    elif kind == "h":
        c.font = Font(bold=True, color="2F4078")

# --- 記録 ---
rec = wb.create_sheet("記録")
header(rec)
for row in rows_block():
    rec.append(row)
body(rec, 30)
validations(rec, 5000)
rec.auto_filter.ref = f"A1:{get_column_letter(len(HEAD))}5000"

# --- ひな形 ---
tpl = wb.create_sheet("ひな形")
header(tpl)
for row in rows_block():
    tpl.append(row)
body(tpl, 30)
validations(tpl, 31)

# --- 集計 ---
s = wb.create_sheet("集計")
s.column_dimensions["A"].width = 36
for i in range(2, 2 + len(BROWSERS)):
    s.column_dimensions[get_column_letter(i)].width = 16
R = "'記録'"
s["A1"], s["A2"] = "今回の日付（最新）", "前回の日付"
s["B1"] = f"=MAX({R}!A2:A5000)"
s["B2"] = f"=SUMPRODUCT(MAX(({R}!A2:A5000<B1)*{R}!A2:A5000))"
for c in ("B1", "B2"):
    s[c].number_format = "yyyy-mm-dd"
    s[c].font = Font(bold=True)
s["C1"] = "※ シークレットの行だけで集計。「未記録」＝その日の行が無い、「圏外」＝30 位までに無い"


def table(top, title, date_cell, col_letter, empty="圏外"):
    s.cell(top, 1, title).font = Font(bold=True, color="2F4078")
    s.cell(top + 1, 1, "検索語 ＼ ブラウザ").font = Font(bold=True)
    for j, (b, _, _) in enumerate(BROWSERS, 2):
        h = s.cell(top + 1, j, b)
        h.fill, h.font = HFILL, HFONT
    for i, (kw, _) in enumerate(KEYWORDS, top + 2):
        s.cell(i, 1, kw)
        for j in range(2, 2 + len(BROWSERS)):
            bl = get_column_letter(j)
            cond = (f"{R}!$A$2:$A$5000,{date_cell},{R}!$C$2:$C$5000,{bl}${top + 1},"
                    f"{R}!$G$2:$G$5000,$A{i},{R}!$F$2:$F$5000,\"シークレット\"")
            f = (f'=IF(COUNTIFS({cond})=0,"未記録",'
                 f'IFERROR(AVERAGEIFS({R}!${col_letter}$2:${col_letter}$5000,{cond}),"{empty}"))')
            c = s.cell(i, j, f)
            c.alignment = Alignment(horizontal="center")
            c.border = BOX
    # ブランド語で 10 位以内の数（順位の表だけ）
    if col_letter == "I":
        r = top + 2 + len(KEYWORDS)
        s.cell(r, 1, "ブランド語で 10 位以内に入った数（3 語中）").font = Font(bold=True)
        for j in range(2, 2 + len(BROWSERS)):
            bl = get_column_letter(j)
            c = s.cell(r, j, (f'=COUNTIFS({R}!$A$2:$A$5000,{date_cell},{R}!$C$2:$C$5000,{bl}${top + 1},'
                              f'{R}!$H$2:$H$5000,"ブランド",{R}!$F$2:$F$5000,"シークレット",'
                              f'{R}!$I$2:$I$5000,"<=10")'))
            c.alignment = Alignment(horizontal="center")
    return top + 4 + len(KEYWORDS)


nxt = table(4, "今回の順位（ふつうの検索結果）", "$B$1", "I")
nxt = table(nxt, "前回の順位（ふつうの検索結果）", "$B$2", "I")
nxt = table(nxt, "今回の地図の枠の順位（なし＝地図の枠に出ていない）", "$B$1", "J", empty="なし")

wb.active = 0
wb.save(OUT)
print("saved", OUT)
