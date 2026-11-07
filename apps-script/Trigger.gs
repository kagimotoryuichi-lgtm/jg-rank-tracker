/**
 * 毎朝の自動計測を GitHub Actions に頼む（2026-11-07 追加）
 *
 * GitHub の定時実行（schedule）が 11/6・11/7 と起動しなかったため、
 * Google 側の時計（Apps Script の時間主導トリガー）から毎朝「実行して」と頼む。
 *
 * 準備（鍵本さんが 1 回だけ行う。鍵はチャット・Vault に書かない）:
 *   1. GitHub で fine-grained token を作る（対象は jg-rank-tracker だけ、権限は Actions: Read and write）
 *   2. このプロジェクトの「プロジェクトの設定」→「スクリプト プロパティ」に
 *      プロパティ GH_TOKEN、値に 1 の鍵を貼って保存
 *   3. 関数「testDispatch」を実行 → 承認 → GitHub に sheet-test の実行が出れば鍵は正しい（費用なし）
 *   4. 関数「installDailyTrigger」を実行 → 毎朝 7 時台（ベトナム時間）に dispatchDaily が動く
 */

var GH_REPO = 'kagimotoryuichi-lgtm/jg-rank-tracker';
var GH_WORKFLOW = 'daily.yml';
var TRIGGER_TZ = 'Asia/Ho_Chi_Minh';

/** 毎朝トリガーから呼ばれる: 全件の本番計測を頼む */
function dispatchDaily() {
  dispatch_('production');
}

/** 鍵の確認用: API を呼ばない sheet-test を頼む（ranks_test シートに見本 1 行ずつ） */
function testDispatch() {
  dispatch_('sheet-test');
}

/** 毎朝 7 時台（ベトナム時間）のトリガーを作る。何度実行しても 1 つだけになる */
function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dispatchDaily') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dispatchDaily')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .nearMinute(30)
    .inTimezone(TRIGGER_TZ)
    .create();
  Logger.log('毎朝 7 時台（' + TRIGGER_TZ + '）に dispatchDaily を実行するトリガーを作りました');
}

function dispatch_(mode) {
  var token = PropertiesService.getScriptProperties().getProperty('GH_TOKEN');
  if (!token) throw new Error('スクリプト プロパティ GH_TOKEN がありません');
  var url = 'https://api.github.com/repos/' + GH_REPO + '/actions/workflows/' + GH_WORKFLOW + '/dispatches';
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token.trim(),
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28'
    },
    payload: JSON.stringify({ ref: 'main', inputs: { mode: mode, limit: '' } }),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  if (code !== 204) {
    // 失敗すると Google から実行失敗のメールが届く（応答本文に鍵は含まれない）
    throw new Error('GitHub への依頼に失敗: HTTP ' + code + ' ' + res.getContentText().slice(0, 300));
  }
  Logger.log('GitHub に ' + mode + ' の実行を頼みました');
}
