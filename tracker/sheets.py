"""Google スプレッドシートへの追記（GitHub Actions で gspread を入れて使う）。"""
import json

HEADER = ["date", "engine", "locale", "locale_label", "keyword", "group", "rank", "url",
          "local_rank", "has_local_pack", "top10", "results_count", "status", "check_url"]


def append_rows(sa_json, sheet_id, rows, worksheet="ranks"):
    import gspread  # 遅延 import（ローカルのテストでは不要）
    gc = gspread.service_account_from_dict(json.loads(sa_json))
    sh = gc.open_by_key(sheet_id)
    try:
        ws = sh.worksheet(worksheet)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=worksheet, rows=1000, cols=len(HEADER))
        ws.append_row(HEADER)
    if not ws.row_values(1):
        ws.append_row(HEADER)
    ws.append_rows([[r.get(k, "") for k in HEADER] for r in rows], value_input_option="USER_ENTERED")
