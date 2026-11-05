# jg-rank-tracker

japangreen.vn の検索順位を毎日計測する（Google・Bing＝DataForSEO、Yahoo! JAPAN＝Apify）。
結果は `data/ranks.csv` と Google スプレッドシート（`ranks` シート）に追記し、Looker Studio で見る。

- 準備の手順: [SETUP.md](SETUP.md)
- 設定（キーワード・地点・言語）: `config.json`
- 試運転（API を呼ばない）: `python -X utf8 run.py --dry-run`
- テスト: `python -X utf8 -m unittest discover -s tests`
