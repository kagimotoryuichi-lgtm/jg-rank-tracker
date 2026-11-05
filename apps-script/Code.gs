/**
 * 検索順位の受け口（japangreen 検索順位 スプレッドシートの Apps Script に貼る）
 *
 * GitHub Actions の run.py が、計測結果を JSON で POST する。
 * 合言葉（スクリプト プロパティ TOKEN）が一致したときだけ、指定のシートに行を追記する。
 *
 * 初回だけ: エディタ上部の関数で「setupToken」を選んで実行 → 実行ログに出た合言葉を
 * GitHub の Secrets「SHEET_TOKEN」に貼る（合言葉はチャット・Vault に書かない）。
 */

var ALLOWED_SHEETS = ['ranks', 'ranks_test'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var body = JSON.parse(e.postData.contents);
    var token = PropertiesService.getScriptProperties().getProperty('TOKEN');
    if (!token || body.token !== token) {
      return reply_({ ok: false, error: 'bad token' });
    }
    var name = body.worksheet || 'ranks';
    if (ALLOWED_SHEETS.indexOf(name) < 0) {
      return reply_({ ok: false, error: 'sheet not allowed: ' + name });
    }
    var header = body.header;
    var rows = body.rows || [];
    if (!header || !header.length) {
      return reply_({ ok: false, error: 'no header' });
    }
    lock.waitLock(30000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(header);
      sh.setFrozenRows(1);
    }
    if (rows.length) {
      var values = rows.map(function (r) {
        return header.map(function (k) { return r[k] === undefined || r[k] === null ? '' : r[k]; });
      });
      sh.getRange(sh.getLastRow() + 1, 1, values.length, header.length).setValues(values);
    }
    return reply_({ ok: true, appended: rows.length, sheet: name });
  } catch (err) {
    return reply_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/** 初回に 1 回だけ実行: 合言葉を作って保存し、実行ログに表示する（作り直すと前の合言葉は使えなくなる） */
function setupToken() {
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('TOKEN', token);
  Logger.log('SHEET_TOKEN に貼る合言葉: ' + token);
}

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
