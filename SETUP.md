# 検索順位の日次計測 — 準備の手順（鍵本さん用）

設計: `C:\Vault\10_Projects\HPリニューアル\japangreen_SEO_順位計測_設計案_2026-11-05.md`
このフォルダ: `C:\Dev\jg-rank-tracker`（Claude が作成。ローカルの試運転は合格済み）

アカウント作成・支払い・キーの入力は、AI が代行できないため鍵本さんが行います。
**API キーやパスワードは、チャット・Vault・このフォルダのファイルに書かないでください**（GitHub の Secrets にだけ入れる）。

## 1. DataForSEO（Google・Bing の順位）
1. https://dataforseo.com/ で「Try for free」から登録（登録時に $1 分の無料クレジット）
2. ダッシュボードの **API Access** で「API login」と「API password」を確認（ログイン用のパスワードとは別物）
3. 試運転は無料クレジットで足ります。本運用の前に **$50 を入金**（最低入金額。月 $3〜10 の想定なので数か月もちます）

## 2. Apify（Yahoo! JAPAN の順位）
1. https://apify.com/ で無料プランに登録（毎月 $5 分の無料枠）
2. **Settings → API & Integrations** で「Personal API token」をコピー
3. 使うアクター: `rainminer/yahoo-co-jp-search-scraper`（第三者製。利用者が少ないので、止まったら代わりを探す）

## 3. Google スプレッドシートと書き込み用アカウント
1. Google ドライブで新しいスプレッドシートを作る（名前の例: `japangreen 検索順位`）。URL の `/d/` と `/edit` の間の文字列が **SHEET_ID**
2. https://console.cloud.google.com/ でプロジェクトを作る（例: `jg-rank-tracker`）
3. 「API とサービス → ライブラリ」で **Google Sheets API** を有効にする
4. 「IAM と管理 → サービス アカウント」で作成（名前の例: `rank-bot`）→「鍵 → 鍵を追加 → JSON」でファイルをダウンロード
5. スプレッドシートの「共有」で、サービスアカウントのメールアドレス（`rank-bot@…iam.gserviceaccount.com`）を **編集者** で追加
6. ダウンロードした JSON は、GitHub に登録したら PC から削除してよい

## 4. GitHub（毎日の自動実行）
1. GitHub で **非公開（Private）** のリポジトリを作る（名前の例: `jg-rank-tracker`。README などは追加しない）
2. URL を Claude に伝える → Claude がこのフォルダの内容を push します
3. リポジトリの **Settings → Secrets and variables → Actions → New repository secret** で 5 つ登録

| 名前 | 中身 |
|---|---|
| `DFS_LOGIN` | DataForSEO の API login |
| `DFS_PASSWORD` | DataForSEO の API password |
| `APIFY_TOKEN` | Apify の Personal API token |
| `GOOGLE_SA_JSON` | サービスアカウントの JSON ファイルの中身をまるごと貼る |
| `SHEET_ID` | スプレッドシートの ID |

## 4.5 書き込みテスト（API のキーが無くてもできる）
1. Secrets に `GOOGLE_SA_JSON` と `SHEET_ID` の 2 つだけ登録した段階で、**Actions → daily-rank-check → Run workflow**、`mode` は `sheet-test` のまま実行
2. 1〜2 分で終わる → スプレッドシートに `ranks_test` シートができ、見本の 9 行（status＝dry-run）が入っていれば成功。このシートは後で消してよい
3. 毎日の自動実行（schedule）は、5 の試運転が済むまで無効にしてある（Claude が有効にする）

## 5. 試運転
1. リポジトリの **Actions → daily-rank-check → Run workflow**。`mode` を `production`、`limit` に `2` を入れて実行（各言語 2 語だけ。費用は数セント）
2. 10〜20 分ほどで終わる → スプレッドシートの `ranks` シートに行が増えていれば成功
3. Claude が結果を確認し、問題がなければ全キーワードで 1 回実行 → 以後は毎朝 6 時（日本時間）に自動実行

## 6. ダッシュボード（Looker Studio）
データが数日分たまったら、Claude と一緒に作ります（https://lookerstudio.google.com/ ・無料）。

## 費用の目安
- DataForSEO: 月 $3〜10（1 日 104 回 × 30 位まで）
- Apify: 月 $1〜2（無料枠に収まる見込み）
- GitHub Actions・Google スプレッドシート・Looker Studio: 無料

## 止めたいとき
リポジトリの Actions 画面で daily-rank-check を「Disable workflow」。
