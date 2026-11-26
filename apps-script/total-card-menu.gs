/**
 * 集客・報酬カード作成ツール（total-card）をシートから開くメニュー
 *
 * 置き場所：2026毎週公演管理（スプレッドシート）に付けたApps Script
 * 使い方：「N月」シートで、開きたい週のブロック内のどこかを選んで、
 *         メニュー「集客・報酬カード」→「この週のカードを作る」
 *
 * ツールに渡すもの：開催日／出演者名／報酬プラン（D列）／ブロマイド会場（K列）／ブロマイド発送（L列）
 * 渡し方：URLの「#d=...」（#より後ろはインターネット上に送られない）
 */

var TC_CONFIG = {
  TOOL_URL: 'https://misakikaku33.github.io/total-card/',   // 公開後のURL（リポジトリ名を変えたらここも直す）
  NAME_COL: 3,        // C列：名前
  PLAN_COL: 4,        // D列：報酬プラン
  VENUE_COL: 11,      // K列：ブロマイド 会場販売数
  SHIP_COL: 12,       // L列：ブロマイド 発送
  MAX_ROWS: 12,       // 1つの週ブロックで読む出演者の最大行数（今のシートは8行）
  MONTH_SHEET: /^\d{1,2}月$/
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('集客・報酬カード')
    .addItem('この週のカードを作る', 'openTotalCard')
    .addToUi();
}

function openTotalCard() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  if (!TC_CONFIG.MONTH_SHEET.test(sheet.getName())) {
    ui.alert('「11月」「12月」のような月のシートを開いてから実行してください。');
    return;
  }
  var cell = sheet.getActiveCell();
  var found = tcFindBlock_(sheet, cell.getRow());
  if (!found) {
    ui.alert('選んでいる位置から、週の日付（例：2026-12-02）が見つかりませんでした。\n週のブロックの中のセルを選んでから、もう一度実行してください。');
    return;
  }
  var payload = {date: found.date, sheet: sheet.getName(), rows: found.rows};
  var url = TC_CONFIG.TOOL_URL + '#d=' + encodeURIComponent(JSON.stringify(payload));

  var names = found.rows.length;
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;font-size:14px;line-height:1.7">' +
    '<p>' + found.date + ' の分（出演者 ' + names + '名）を渡して、ツールを開きます。</p>' +
    '<p>自動で開かない場合は、下のボタンを押してください。</p>' +
    '<p><a href="' + url.replace(/&/g, '&amp;').replace(/"/g, '&quot;') + '" target="_blank" ' +
    'style="display:inline-block;padding:10px 18px;background:#3c6230;color:#fff;border-radius:99px;text-decoration:none;font-weight:bold">ツールを開く</a></p>' +
    '<script>window.open(' + JSON.stringify(url).replace(/</g, '\\x3c') + ', "_blank");</script>' +
    '</div>')
    .setWidth(360).setHeight(190);
  ui.showModalDialog(html, '集客・報酬カード');
}

/**
 * 選んでいる行から上にさかのぼって、週の日付の行（C列が yyyy-mm-dd）を探し、
 * そのブロックの出演者（名前・D列・K列・L列）を読む。
 * 次の週の日付の行までを範囲とし、選んだ行が次の週に入っていたらそちらを使う。
 */
function tcFindBlock_(sheet, activeRow) {
  var last = sheet.getLastRow();
  var colC = sheet.getRange(1, TC_CONFIG.NAME_COL, last, 1).getDisplayValues();
  var isDate = function (v) { return /^\d{4}-\d{2}-\d{2}$/.test(String(v).trim()); };

  var dateRow = 0;
  for (var r = Math.min(activeRow, last); r >= 1; r--) {
    if (isDate(colC[r - 1][0])) { dateRow = r; break; }
  }
  if (!dateRow) return null;

  // 見出し行（C列が「名前」）を探す。次の週の日付に当たったら打ち切り
  var headerRow = 0;
  for (var r2 = dateRow + 1; r2 <= last; r2++) {
    var v = String(colC[r2 - 1][0]).trim();
    if (isDate(v)) break;
    if (v === '名前') { headerRow = r2; break; }
  }
  if (!headerRow) return null;

  var rows = [];
  var lim = Math.min(last, headerRow + TC_CONFIG.MAX_ROWS);
  if (lim > headerRow) {
    var vals = sheet.getRange(headerRow + 1, 1, lim - headerRow, TC_CONFIG.SHIP_COL).getDisplayValues();
    for (var i = 0; i < vals.length; i++) {
      var name = String(vals[i][TC_CONFIG.NAME_COL - 1]).trim();
      if (isDate(name) || name === '収入の部') break;
      if (!name) continue;
      rows.push({
        name: name,
        plan: String(vals[i][TC_CONFIG.PLAN_COL - 1]).trim(),
        bVenue: tcToNum_(vals[i][TC_CONFIG.VENUE_COL - 1]),
        bShip: tcToNum_(vals[i][TC_CONFIG.SHIP_COL - 1])
      });
    }
  }
  return {date: String(colC[dateRow - 1][0]).trim(), rows: rows};
}

function tcToNum_(s) {
  var n = Number(String(s == null ? '' : s).replace(/[^\d.\-]/g, ''));
  return isFinite(n) && n > 0 ? n : 0;
}
