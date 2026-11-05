# jg-rank-tracker

japangreen.vn の検索順位を毎日計測する（Google・Bing＝DataForSEO、Yahoo! JAPAN＝Apify）。
結果は `data/ranks.csv` と Google スプレッドシート（`ranks` シート。受け口は `apps-script/Code.gs`）に追記し、Looker Studio で見る。

- 準備の手順: [SETUP.md](SETUP.md)
- 設定（キーワード・地点・言語）: `config.json`
- 試運転（API を呼ばない）: `python -X utf8 run.py --dry-run`
- テスト: `python -X utf8 -m unittest discover -s tests`

## 2026-11-05 休止中（手動集計へ）
- API キー（DataForSEO・Apify）は購入せず、当面は手動で集計する方針（鍵本さんの決定）。毎日の自動実行（schedule）は無効のまま。
- 手動の記録表は `manual/build_manual_sheet.py` で作る（出力: `OneDrive\japangreen-HP\順位確認\japangreen_順位_手動記録表.xlsx` → Google スプレッドシート「japangreen_検索順位」にインポート）。
- 再開するときは SETUP.md の 1・2（DataForSEO・Apify と Secrets 3 つ）→ `mode=production`・`limit=2` で試運転 → `daily.yml` の schedule のコメントを外す。
