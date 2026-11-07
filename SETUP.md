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

## 3. Google スプレッドシートと受け口（Apps Script。Google Cloud は使わない）
1. Google ドライブで新しいスプレッドシートを作る（名前の例: `japangreen 検索順位`）※2026-11-05 作成済み
2. スプレッドシートの **拡張機能 → Apps Script** を開く → 最初からある `function myFunction() {}` を消し、このフォルダの `apps-script/Code.gs` の中身をまるごと貼る → 💾 保存（プロジェクト名は `rank-receiver` など）
3. 上部の関数の選択で **setupToken** を選び「実行」→ 初回は「権限を確認」→ 自分のアカウント →「詳細」→「（安全ではないページ）に移動」→「許可」
4. 下の「実行ログ」に出た合言葉をコピー → GitHub の Secrets `SHEET_TOKEN` に貼る（チャット・Vault には書かない）
5. 右上の **デプロイ → 新しいデプロイ** → 種類の歯車で **ウェブアプリ** → 次のユーザーとして実行＝**自分**・アクセスできるユーザー＝**全員** →「デプロイ」
6. 表示された **ウェブアプリの URL**（`https://script.google.com/macros/s/…/exec`）をコピー → Secrets `SHEET_WEBAPP_URL` に貼る
7. スクリプトを直したときは「デプロイ → デプロイを管理 → 鉛筆 → バージョン＝新しいバージョン → デプロイ」（URL は変わらない）

## 4. GitHub（毎日の自動実行）
1. GitHub で **非公開（Private）** のリポジトリを作る（名前の例: `jg-rank-tracker`。README などは追加しない）
2. URL を Claude に伝える → Claude がこのフォルダの内容を push します
3. リポジトリの **Settings → Secrets and variables → Actions → New repository secret** で 5 つ登録

| 名前 | 中身 |
|---|---|
| `DFS_LOGIN` | DataForSEO の API login |
| `DFS_PASSWORD` | DataForSEO の API password |
| `APIFY_TOKEN` | Apify の Personal API token |
| `SHEET_WEBAPP_URL` | Apps Script のウェブアプリの URL（3-6） |
| `SHEET_TOKEN` | setupToken が出した合言葉（3-4） |

## 4.5 書き込みテスト（API のキーが無くてもできる）
1. Secrets に `SHEET_WEBAPP_URL` と `SHEET_TOKEN` の 2 つだけ登録した段階で、**Actions → daily-rank-check → Run workflow**、`mode` は `sheet-test` のまま実行
2. 1〜2 分で終わる → スプレッドシートに `ranks_test` シートができ、見本の 13 行（status＝dry-run、browser_profile 列あり）が入っていれば成功。このシートは後で消してよい
   - 2026-11-05 に列（browser_profile・device・os）を足したので、**前の `ranks_test` シートは先に削除**してから実行する（見出しが古いまま残るため）
3. 毎日の自動実行（schedule）は、5 の試運転が済むまで無効にしてある（Claude が有効にする）

## 5. 試運転
1. リポジトリの **Actions → daily-rank-check → Run workflow**。`mode` を `production`、`limit` に `2` を入れて実行（各言語 2 語だけ。費用は数セント）
2. 10〜20 分ほどで終わる → スプレッドシートの `ranks` シートに行が増えていれば成功
3. Claude が結果を確認し、問題がなければ全キーワードで 1 回実行 → 以後は毎朝 7 時台（ベトナム時間）に自動実行（5.5 の Apps Script から起動）

## 5.5 毎朝の起動を Apps Script から行う（2026-11-07 から）
GitHub の定時実行（schedule）が 11/6・11/7 と起動しなかったため、Google 側の時計から毎朝「実行して」と頼む。GitHub の定時実行は 9:47（VN）の予備として残し、その日の計測が済んでいれば何もしない。
1. GitHub 右上のアイコン → **Settings → Developer settings → Personal access tokens → Fine-grained tokens → Generate new token**
   - Token name: `jg-rank-tracker 朝の起動`／Expiration: 1 年（期限の日を Vault に書いておく）
   - Repository access: **Only select repositories** → `jg-rank-tracker`
   - Permissions → Repository permissions → **Actions: Read and write**（ほかは触らない）
   - Generate token → 表示された鍵をコピー（チャット・Vault には書かない）
2. スプレッドシート「japangreen_検索順位」→ 拡張機能 → Apps Script → 左の「＋」→ スクリプト → 名前 `Trigger` → `apps-script/Trigger.gs` の中身を貼って保存
3. 左の歯車「プロジェクトの設定」→ 下の **スクリプト プロパティ** → 「スクリプト プロパティを追加」→ プロパティ `GH_TOKEN`、値に 1 の鍵 → 保存（既にある `TOKEN` は消さない）
4. エディタに戻り、上の関数で `testDispatch` を選んで実行 → 承認（前回と同じ「詳細」→「安全ではないページに移動」）→ 実行ログに「sheet-test の実行を頼みました」。GitHub の Actions に sheet-test の実行が出れば成功（費用なし）
5. 関数 `installDailyTrigger` を実行 → 左の時計「トリガー」に dispatchDaily（日付ベースのタイマー・午前 7〜8 時）が 1 つあれば完了
- 失敗すると Google から「実行に失敗しました」のメールが届く。鍵の期限切れもこれで分かる

## 6. ダッシュボード（Looker Studio）
データが数日分たまったら、Claude と一緒に作ります（https://lookerstudio.google.com/ ・無料）。

## 費用の目安
- DataForSEO: 月 $6〜9 程度（1 日 156 回 × 30 位まで。Chrome（PC）＝Google 52・Edge（PC）＝Bing 52・Safari（iPhone）＝Google スマホ 52。最大で 1 回 $0.0018＝10 件ずつ 3 枚）
- Apify: 月 $1〜2（無料枠に収まる見込み）
- GitHub Actions・Google スプレッドシート・Looker Studio: 無料

## 止めたいとき
リポジトリの Actions 画面で daily-rank-check を「Disable workflow」。
