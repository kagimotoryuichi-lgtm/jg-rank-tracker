# jg-rank-tracker

japangreen.vn の検索順位を毎日計測する（Google・Bing＝DataForSEO、Yahoo! JAPAN＝Apify）。
結果は `data/ranks.csv` と Google スプレッドシート（`ranks` シート。受け口は `apps-script/Code.gs`）に追記し、Looker Studio で見る。

- 準備の手順: [SETUP.md](SETUP.md)
- 設定（キーワード・地点・言語）: `config.json`
- 試運転（API を呼ばない）: `python -X utf8 run.py --dry-run`
- テスト: `python -X utf8 -m unittest discover -s tests`

## 2026-11-05 の方針
- 同日に一度「API は買わず手動」としたが、**DataForSEO・Apify に課金して自動計測を再開する**ことになった（鍵本さんの決定）。Cốc Cốc は集計の対象外。
- 計測はブラウザ想定プロファイル（`config.json` の `profiles`）ごと: Chrome（PC）＝Google PC、Edge（PC）＝Bing PC、Safari（iPhone）＝Google mobile/ios、Yahoo! JAPAN（PC）＝Apify。1 日 173 件。
- 手動の記録表（実ブラウザでの照合用・3 ブラウザ×10 語）は `manual/build_manual_sheet.py` で作る（出力: `OneDrive\japangreen-HP\順位確認\japangreen_順位_手動記録表.xlsx`）。
- 2026-11-05: 試運転（Google・iPhone・Bing・Yahoo! JAPAN）が済み、毎日の自動実行を有効化（同日、ベトナム時間 8:00＝01:00 UTC に変更。記録の日付もベトナム時間）。2026-11-06: 8:00 の定時実行が起動しなかったため（毎時 0 分は GitHub が混む）、7:47（00:47 UTC）にずらした。2026-11-07: 7:47 も起動しなかったため、毎朝の起動を Apps Script の時間主導トリガー（`apps-script/Trigger.gs`、7 時台 VN、workflow_dispatch を呼ぶ）に移した。GitHub の schedule は 9:47 の予備で、その日の取得済みが 100 件以上なら `--skip-if-done` で何もしない。Bing の 40102 は最大 2 回送り直す。Secrets は読み込み時に strip し、エラー文からは伏せる。
- 2026-11-12: Android（Chrome）と AI 概要（AI による概要）は計測しない（鍵本さんの決定）。`config.json` の `ai_overview: false`・プロファイルから android を外した。競合の順位・強調スニペット・他の人はこちらも質問は続ける。1 日 173 件
