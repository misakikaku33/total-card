/**
 * 集客リスト取り込み
 *
 * ドライブの専用フォルダ（「集客リスト取り込み」）に、来場者リスト・配信購入リストのExcelを置くと、
 * 数分おきに見つけて、月シートの該当の週ブロックに集客数を入れる。
 *
 *   会場 → 同じ名前の行の F列    配信 → 同じ名前の行の G列
 *   出演者の指定がない集客（扱いなし・指定なし）→ 「団体扱い」の人数欄（会場 H列 / 配信 L列）
 *
 * 決めたこと（岬さん確認済み 2026-11-26）
 *  - ファイルの種類は列名で判別（来場者: 購入者・扱い・購入枚数 / 配信: 購入者名・アンケート回答）
 *  - どの週かは、ファイルの中の日付（来場者「予約日」/ 配信「配信日時/イベント日」）で決める
 *  - すでに入っている数字は、ファイルの数字で上書きする（上書きは「取り込みログ」に 前→後 で残す）
 *  - ブロックにいない名前は入れず、「取り込みログ」に残す（空き行への追加はしない）
 *  - 数分おきに自動で動く
 *
 * 使うもの：ScriptApp / DriveApp / SpreadsheetApp / UrlFetchApp（Drive APIでExcelをシートに変換）
 */

var IA = {
  FOLDER_NAME: '集客リスト取り込み',
  PARENT_NAME: '01_ミサキカク毎週公演',      // この中にフォルダを作る（なければマイドライブ直下）
  LOG_SHEET: '取り込みログ',
  MONTH_SHEET: /^\d{1,2}月$/,
  NAME_COL: 3,          // C列：名前（週の日付もここ）
  VENUE_COL: 6,         // F列：会場
  STREAM_COL: 7,        // G列：配信
  MAX_ROWS: 12,         // 1つの週ブロックで読む出演者の最大行数
  GROUP_LABEL: '団体扱い',
  GROUP_VENUE_COL: 8,   // H列：団体扱いの会場人数（ラベルはE列）
  GROUP_STREAM_COL: 12, // L列：団体扱いの配信人数（ラベルはK列）
  LABEL_VENUE_COL: 5,
  LABEL_STREAM_COL: 11,
  UNASSIGNED: ['（扱いなし）', '指定なし', '扱いなし'],
  TRIGGER_MINUTES: 5,
  P_FOLDER: 'IA_FOLDER_ID',
  P_DONE: 'IA_PROCESSED',
  P_DRY: 'IA_DRY_RUN'   // '1' のときは、シートに書かずにログだけ残す（動作確認用）
};

/* =========================================================
 * 1. 初期設定（1回だけ実行）
 * ======================================================= */

/** フォルダを作る（すでにあれば、そのまま使う）。フォルダのURLを取り込みログに残す */
function iaSetupFolder() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(IA.P_FOLDER);
  var folder = null;
  if (id) { try { folder = DriveApp.getFolderById(id); if (folder.isTrashed()) folder = null; } catch (e) { folder = null; } }
  if (!folder) {
    var parent = null;
    var it = DriveApp.getFoldersByName(IA.PARENT_NAME);
    if (it.hasNext()) parent = it.next();
    if (parent) {
      var ex = parent.getFoldersByName(IA.FOLDER_NAME);
      folder = ex.hasNext() ? ex.next() : parent.createFolder(IA.FOLDER_NAME);
    } else {
      folder = DriveApp.createFolder(IA.FOLDER_NAME);
    }
    props.setProperty(IA.P_FOLDER, folder.getId());
  }
  var where = '';
  try { var ps = folder.getParents(); where = ps.hasNext() ? ps.next().getName() : 'マイドライブ'; } catch (e) {}
  iaAppendLog_([[new Date(), '', '', '', '設定', 'フォルダ「' + folder.getName() + '」(' + where + ' の中): ' + folder.getUrl()]]);
  return folder.getUrl();
}

/** 数分おきの自動実行を設定する（すでにあれば作り直す） */
function iaSetupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'iaImportNewFiles') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('iaImportNewFiles').timeBased().everyMinutes(IA.TRIGGER_MINUTES).create();
  iaAppendLog_([[new Date(), '', '', '', '設定', IA.TRIGGER_MINUTES + '分おきの自動取り込みを設定しました']]);
}

/** 自動実行をやめる */
function iaRemoveTrigger() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'iaImportNewFiles') { ScriptApp.deleteTrigger(t); n++; }
  });
  iaAppendLog_([[new Date(), '', '', '', '設定', '自動取り込みを停止しました(' + n + '件)']]);
}

/** 書き込まないモードの切り替え（動作確認用） */
function iaSetDryRun(on) {
  PropertiesService.getScriptProperties().setProperty(IA.P_DRY, on ? '1' : '0');
}

/* =========================================================
 * 2. 取り込みの本体
 * ======================================================= */

/** メニュー「集客リストを今すぐ取り込む」。エラーで止まったファイルも、もう一度読む */
function iaImportNow() {
  var props = PropertiesService.getScriptProperties();
  var done = iaLoadDone_();
  Object.keys(done).forEach(function (k) { if (!done[k].ok) delete done[k]; });
  props.setProperty(IA.P_DONE, JSON.stringify(done));
  var summary = iaImportNewFiles();
  try { SpreadsheetApp.getUi().alert(summary); } catch (e) {}
}

/** 数分おきに動く入口。新しい（または更新された）ファイルだけを取り込む */
function iaImportNewFiles() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return '別の取り込みが動いているため、今回は見送りました。';
  try {
    var props = PropertiesService.getScriptProperties();
    var folderId = props.getProperty(IA.P_FOLDER);
    if (!folderId) return 'フォルダが未設定です。先に iaSetupFolder を実行してください。';
    var dry = props.getProperty(IA.P_DRY) === '1';
    var done = iaLoadDone_();
    var files = iaListTargets_(DriveApp.getFolderById(folderId));
    var handled = 0, msgs = [];
    files.forEach(function (f) {
      var id = f.getId(), stamp = f.getLastUpdated().getTime();
      if (done[id] && done[id].t === stamp) return;
      var logRows = [];
      var ok = false;
      try {
        ok = iaProcessFile_(f, dry, logRows);
      } catch (e) {
        logRows.push([new Date(), f.getName(), '', '', 'エラー', String(e && e.message || e)]);
      }
      iaAppendLog_(logRows);
      done[id] = {t: stamp, ok: ok};
      handled++;
      msgs.push(f.getName() + ' → ' + (ok ? '取り込みました' : '取り込めませんでした(取り込みログを見てください)') + (dry ? '(試行・書き込みなし)' : ''));
    });
    props.setProperty(IA.P_DONE, JSON.stringify(done));
    return handled ? msgs.join('\n') : '新しいファイルはありませんでした。';
  } finally {
    lock.releaseLock();
  }
}

function iaLoadDone_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(IA.P_DONE) || '{}'); } catch (e) { return {}; }
}

/** フォルダ内のExcel（またはGoogleスプレッドシート）を、更新の古い順に返す */
function iaListTargets_(folder) {
  var out = [];
  var it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    var mt = f.getMimeType();
    if (mt === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
        mt === 'application/vnd.ms-excel' ||
        mt === 'application/vnd.google-apps.spreadsheet') out.push(f);
  }
  out.sort(function (a, b) { return a.getLastUpdated().getTime() - b.getLastUpdated().getTime(); });
  return out;
}

/** 1ファイルを読んで、シートに入れる。入れられたら true */
function iaProcessFile_(file, dry, logRows) {
  var now = new Date(), name = file.getName();
  var read = iaReadFile_(file);                     // {kind, date, rows} または null
  if (!read) { logRows.push([now, name, '', '', 'エラー', '来場者リスト・配信購入リストの形式として読めませんでした(列名を確認してください)']); return false; }
  var kindLabel = read.kind === 'venue' ? '会場' : '配信';
  if (!read.date) { logRows.push([now, name, kindLabel, '', 'エラー', '開催日(' + (read.kind === 'venue' ? '予約日' : '配信日時/イベント日') + ')が読めませんでした']); return false; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var month = Number(read.date.slice(5, 7)) + '月';
  var sheet = ss.getSheetByName(month);
  if (!sheet) { logRows.push([now, name, kindLabel, read.date, 'エラー', '「' + month + '」のシートがありません']); return false; }
  var block = iaFindBlock_(sheet, read.date);
  if (!block) { logRows.push([now, name, kindLabel, read.date, 'エラー', '「' + month + '」シートに、' + read.date + ' の週のブロックが見つかりませんでした']); return false; }

  var agg = iaAggregate_(read.rows);
  var plan = iaPlan_(block, agg, read.kind);
  var tag = dry ? '(試行・書き込みなし) ' : '';

  if (!dry) iaApplyPlan_(sheet, plan);

  var matched = plan.writes.length + plan.unchanged;
  logRows.push([now, name, kindLabel, read.date, dry ? '試行' : 'OK',
    tag + month + 'シート ' + read.date + ' の週に取り込みました(出演者 ' + block.performers.length + '名のうち、ファイルに名前がある人 ' + plan.found + '名、書き換え ' + plan.writes.length + '件)']);
  plan.writes.forEach(function (w) {
    logRows.push([now, name, kindLabel, read.date, w.old === '' ? '入力' : '上書き',
      tag + w.label + ' ' + kindLabel + ' ' + (w.old === '' ? '(空欄)' : w.old) + ' → ' + w.value]);
  });
  plan.missing.forEach(function (m) {
    logRows.push([now, name, kindLabel, read.date, '確認', 'ブロックにいない名前のため入れませんでした: 「' + m.name + '」 ' + m.qty + '名']);
  });
  plan.dupNames.forEach(function (n) {
    logRows.push([now, name, kindLabel, read.date, '確認', 'ブロックに同じ名前が複数あるため、1行目だけに入れました: 「' + n + '」']);
  });
  if (plan.groupNote) logRows.push([now, name, kindLabel, read.date, plan.groupNote.status, tag + plan.groupNote.text]);
  return true;
}

/* =========================================================
 * 3. Excelの読み取り（ツールと同じ数え方）
 * ======================================================= */

/** ファイルを読んで、種別・日付・{performer, qty} の行を返す。判別できなければ null */
function iaReadFile_(file) {
  var mt = file.getMimeType();
  var ssId = file.getId(), temp = false;
  if (mt !== 'application/vnd.google-apps.spreadsheet') { ssId = iaConvertToSheet_(file); temp = true; }
  try {
    var ss = SpreadsheetApp.openById(ssId);
    var matrix = ss.getSheets()[0].getDataRange().getValues();
    return iaParse_(matrix, ss.getSpreadsheetTimeZone());
  } finally {
    if (temp) { try { DriveApp.getFileById(ssId).setTrashed(true); } catch (e) {} }
  }
}

/** Excel → 一時のGoogleスプレッドシート（読み終わったらごみ箱へ） */
function iaConvertToSheet_(file) {
  var token = ScriptApp.getOAuthToken();
  var name = 'tmp-import-' + file.getId();
  var res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + file.getId() + '/copy', {
    method: 'post', contentType: 'application/json',
    headers: {Authorization: 'Bearer ' + token},
    payload: JSON.stringify({name: name, mimeType: 'application/vnd.google-apps.spreadsheet'}),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() < 300) return JSON.parse(res.getContentText()).id;

  // うまくいかない場合は、アップロードで変換する
  var boundary = 'ia' + new Date().getTime();
  var head = Utilities.newBlob('--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify({name: name, mimeType: 'application/vnd.google-apps.spreadsheet'}) +
    '\r\n--' + boundary + '\r\nContent-Type: ' + file.getMimeType() + '\r\n\r\n').getBytes();
  var tail = Utilities.newBlob('\r\n--' + boundary + '--').getBytes();
  var body = head.concat(file.getBlob().getBytes()).concat(tail);
  var res2 = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', {
    method: 'post', contentType: 'multipart/related; boundary=' + boundary,
    headers: {Authorization: 'Bearer ' + token}, payload: body, muteHttpExceptions: true
  });
  if (res2.getResponseCode() >= 300) throw new Error('Excelの変換に失敗しました: ' + res2.getContentText().slice(0, 200));
  return JSON.parse(res2.getContentText()).id;
}

function iaClean_(v) { return v == null ? '' : String(v).replace(/_x000D_/g, '').replace(/\r/g, '').trim(); }
function iaNorm_(s) { return String(s == null ? '' : s).normalize('NFKC').replace(/\s+/g, '').toLowerCase(); }
function iaPad2_(n) { return ('0' + n).slice(-2); }

/** セルの値から yyyy-MM-dd を取り出す(日付セルでも、文字でも)。読めなければ '' */
function iaToYmd_(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var m = String(v == null ? '' : v).match(/(\d{4})\s*[\/\-年.]\s*(\d{1,2})\s*[\/\-月.]\s*(\d{1,2})/);
  return m ? m[1] + '-' + iaPad2_(m[2]) + '-' + iaPad2_(m[3]) : '';
}

function iaFindHeader_(matrix, keys) {
  for (var i = 0; i < Math.min(matrix.length, 15); i++) {
    var row = (matrix[i] || []).map(iaClean_);
    if (keys.every(function (k) { return row.indexOf(k) >= 0; })) return i;
  }
  return -1;
}

/**
 * 表（二次元配列）から、種別・開催日・行を取り出す。
 *  来場者: 購入者・扱い・購入枚数 / キャンセルは0枚 / 日付は「予約日」の最初の日付
 *  配信:   購入者名・アンケート回答 / ステータスに「販売」を含まない行は除く / 日付は「配信日時/イベント日」の最初の日付
 */
function iaParse_(matrix, tz) {
  var hi = iaFindHeader_(matrix, ['購入者', '扱い', '購入枚数']);
  if (hi >= 0) {
    var H = matrix[hi].map(iaClean_);
    var c = {buyer: H.indexOf('購入者'), perf: H.indexOf('扱い'), qty: H.indexOf('購入枚数'), date: H.indexOf('予約日')};
    var rows = [], date = '';
    for (var i = hi + 1; i < matrix.length; i++) {
      var r = matrix[i] || [];
      if (!iaClean_(r[c.buyer])) continue;
      if (!date && c.date >= 0) date = iaToYmd_(r[c.date], tz);
      var rawQty = iaClean_(r[c.qty]);
      var cancelled = /キャンセル/i.test(rawQty);
      rows.push({performer: iaClean_(r[c.perf]) || '（扱いなし）', qty: cancelled ? 0 : (Number(r[c.qty]) || 1)});
    }
    return {kind: 'venue', date: date, rows: rows};
  }
  hi = iaFindHeader_(matrix, ['購入者名', 'アンケート回答']);
  if (hi >= 0) {
    var H2 = matrix[hi].map(iaClean_);
    var d = {name: H2.indexOf('購入者名'), perf: H2.indexOf('アンケート回答'), qty: H2.indexOf('購入個数'),
      status: H2.indexOf('ステータス'), date: H2.indexOf('配信日時/イベント日')};
    var rows2 = [], date2 = '';
    for (var j = hi + 1; j < matrix.length; j++) {
      var r2 = matrix[j] || [];
      if (!iaClean_(r2[d.name])) continue;
      if (d.status >= 0) { var st = iaClean_(r2[d.status]); if (st && st.indexOf('販売') < 0) continue; }
      if (!date2 && d.date >= 0) date2 = iaToYmd_(r2[d.date], tz);
      rows2.push({performer: iaClean_(r2[d.perf]) || '（扱いなし）', qty: d.qty >= 0 ? (Number(r2[d.qty]) || 1) : 1});
    }
    return {kind: 'stream', date: date2, rows: rows2};
  }
  return null;
}

/** 出演者名ごとの人数。指定なしは別にまとめる */
function iaAggregate_(rows) {
  var unassigned = {};
  IA.UNASSIGNED.forEach(function (n) { unassigned[iaNorm_(n)] = true; });
  var byName = {}, group = 0;
  rows.forEach(function (r) {
    var key = iaNorm_(r.performer);
    if (unassigned[key]) { group += r.qty; return; }
    if (!byName[key]) byName[key] = {name: r.performer, qty: 0};
    byName[key].qty += r.qty;
  });
  return {byName: byName, group: group};
}

/* =========================================================
 * 4. シートの週ブロックを探して、入れる内容を決める
 * ======================================================= */

/** 日付（yyyy-MM-dd）の週ブロックを探す。出演者の行・「団体扱い」の行を返す。なければ null */
function iaFindBlock_(sheet, ymd) {
  var last = sheet.getLastRow();
  if (last < 1) return null;
  var colC = sheet.getRange(1, IA.NAME_COL, last, 1).getDisplayValues();
  var isDate = function (v) { return /^\d{4}-\d{2}-\d{2}$/.test(String(v).trim()); };
  var dateRow = 0;
  for (var r = 1; r <= last; r++) { if (String(colC[r - 1][0]).trim() === ymd) { dateRow = r; break; } }
  if (!dateRow) return null;
  var endRow = last;
  for (var r2 = dateRow + 1; r2 <= last; r2++) { if (isDate(colC[r2 - 1][0])) { endRow = r2 - 1; break; } }
  var headerRow = 0;
  for (var r3 = dateRow + 1; r3 <= endRow; r3++) { if (String(colC[r3 - 1][0]).trim() === '名前') { headerRow = r3; break; } }
  if (!headerRow) return null;

  var width = IA.LABEL_STREAM_COL > IA.GROUP_STREAM_COL ? IA.LABEL_STREAM_COL : IA.GROUP_STREAM_COL;
  var n = endRow - headerRow;
  var disp = n > 0 ? sheet.getRange(headerRow + 1, 1, n, width).getDisplayValues() : [];
  var vals = n > 0 ? sheet.getRange(headerRow + 1, 1, n, width).getValues() : [];

  var performers = [], groupRow = 0;
  var lim = Math.min(n, IA.MAX_ROWS);
  for (var i = 0; i < lim; i++) {
    var nm = String(disp[i][IA.NAME_COL - 1]).trim();
    if (nm === '収入の部') break;
    if (!nm) continue;
    performers.push({row: headerRow + 1 + i, name: nm, key: iaNorm_(nm),
      oldVenue: vals[i][IA.VENUE_COL - 1], oldStream: vals[i][IA.STREAM_COL - 1]});
  }
  for (var k = 0; k < n; k++) {
    if (String(disp[k][IA.LABEL_VENUE_COL - 1]).trim() === IA.GROUP_LABEL) {
      groupRow = headerRow + 1 + k;
      return {dateRow: dateRow, headerRow: headerRow, endRow: endRow, performers: performers, groupRow: groupRow,
        groupLabelStreamOk: String(disp[k][IA.LABEL_STREAM_COL - 1]).trim() === IA.GROUP_LABEL,
        oldGroupVenue: vals[k][IA.GROUP_VENUE_COL - 1], oldGroupStream: vals[k][IA.GROUP_STREAM_COL - 1]};
    }
  }
  return {dateRow: dateRow, headerRow: headerRow, endRow: endRow, performers: performers, groupRow: 0, groupLabelStreamOk: false,
    oldGroupVenue: '', oldGroupStream: ''};
}

function iaNum_(v) { return v === '' || v == null ? '' : Number(v); }

/**
 * 入れる内容を決める（書き込みはしない）。
 *  - ブロックの名前がファイルにあれば、その人数。なければ0（ただし、空欄のままの人に0は書かない）
 *  - 人数が今の数字と同じなら、何もしない
 *  - 指定なし(団体扱い)は、1名以上のときだけ書く（0名のときは手入力の数字を消さないよう、触らない）
 */
function iaPlan_(block, agg, kind) {
  var col = kind === 'venue' ? IA.VENUE_COL : IA.STREAM_COL;
  var kindLabel = kind === 'venue' ? '会場' : '配信';
  var writes = [], unchanged = 0, found = 0, dupNames = [], seen = {}, usedKeys = {};
  block.performers.forEach(function (p) {
    if (seen[p.key]) { if (dupNames.indexOf(p.name) < 0) dupNames.push(p.name); return; }
    seen[p.key] = true;
    var hit = agg.byName[p.key];
    if (hit) { found++; usedKeys[p.key] = true; }
    var value = hit ? hit.qty : 0;
    var oldRaw = kind === 'venue' ? p.oldVenue : p.oldStream;
    var old = iaNum_(oldRaw);
    if (old === '' && value === 0) return;           // 空欄のまま0にはしない
    if (old !== '' && old === value) { unchanged++; return; }
    writes.push({row: p.row, col: col, value: value, old: old, label: p.name});
  });
  var missing = [];
  Object.keys(agg.byName).forEach(function (k) {
    if (!usedKeys[k] && !seen[k]) missing.push(agg.byName[k]);
  });

  var groupNote = null, groupWrite = null;
  if (agg.group > 0) {
    var gCol = kind === 'venue' ? IA.GROUP_VENUE_COL : IA.GROUP_STREAM_COL;
    var okLabel = kind === 'venue' ? block.groupRow > 0 : (block.groupRow > 0 && block.groupLabelStreamOk);
    if (!okLabel) {
      groupNote = {status: '確認', text: '指定なしの集客 ' + agg.group + '名は、「団体扱い」の欄が見つからないため入れませんでした'};
    } else {
      var oldG = iaNum_(kind === 'venue' ? block.oldGroupVenue : block.oldGroupStream);
      if (oldG === agg.group) {
        groupNote = {status: 'OK', text: '指定なしの集客 ' + agg.group + '名(団体扱い ' + kindLabel + ')は、すでに同じ数字です'};
      } else {
        groupWrite = {row: block.groupRow, col: gCol, value: agg.group, old: oldG, label: '団体扱い'};
        groupNote = {status: oldG === '' ? '入力' : '上書き',
          text: '指定なしの集客 ' + agg.group + '名を、団体扱い(' + kindLabel + ')に入れました(前: ' + (oldG === '' ? '空欄' : oldG) + ')'};
      }
    }
  }
  if (groupWrite) writes.push(groupWrite);
  return {writes: writes, unchanged: unchanged, found: found, missing: missing, dupNames: dupNames, groupNote: groupNote};
}

function iaApplyPlan_(sheet, plan) {
  plan.writes.forEach(function (w) { sheet.getRange(w.row, w.col).setValue(w.value); });
}

/* =========================================================
 * 5. 取り込みログ
 * ======================================================= */

function iaAppendLog_(rows) {
  if (!rows || !rows.length) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(IA.LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(IA.LOG_SHEET, ss.getNumSheets());
    sh.getRange(1, 1, 1, 6).setValues([['日時', 'ファイル', '種別', '開催日', '結果', '内容']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 220); sh.setColumnWidth(6, 560);
  }
  var out = rows.map(function (r) {
    var d = r[0];
    var stamp = Object.prototype.toString.call(d) === '[object Date]' ? Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss') : d;
    return [stamp, r[1], r[2], r[3], r[4], r[5]];
  });
  sh.getRange(sh.getLastRow() + 1, 1, out.length, 6).setValues(out);
}

/** 初期設定のやり直し用：保存してあるフォルダIDと処理済みの記録を消す（別アカウントで作り直すとき） */
function iaResetSetup() {
  var p = PropertiesService.getScriptProperties();
  p.deleteProperty(IA.P_FOLDER);
  p.deleteProperty(IA.P_DONE);
}

/** 試行モードON：シートには書かず、取り込みログだけに結果を残す */
function iaDryRunOn() { iaSetDryRun(true); }

/** 試行モードOFF：本番（シートに書く） */
function iaDryRunOff() { iaSetDryRun(false); }

/** 本番に切り替える：試行モードをOFFにして、処理済みの記録を消し、フォルダ内のファイルをもう一度取り込む */
function iaGoLive() {
  iaSetDryRun(false);
  PropertiesService.getScriptProperties().deleteProperty(IA.P_DONE);
  return iaImportNewFiles();
}

/* =========================================================
 * 6. カード作成ツールに渡す購入者一覧（メニュー「この週のカードを作る」から呼ばれる）
 *    取り込みフォルダとその中のフォルダ（保管用など）から、ファイル名の先頭が開催日(yyyymmdd)のExcelを探す
 *    一覧はシートには保存しない（開くたびに読む）
 * ======================================================= */

/** 開催日(yyyy-MM-dd)の購入者一覧を読む。見つからなくても止まらず、notes に理由を入れて返す */
function iaReadLists(ymd) {
  var out = {venue: null, stream: null, eventName: '', nameFallback: '', notes: []};
  var folderId = PropertiesService.getScriptProperties().getProperty(IA.P_FOLDER);
  if (!folderId) { out.notes.push('取り込みフォルダが未設定のため、購入者一覧は読めませんでした。'); return out; }
  var files = iaFindFilesByDate_(DriveApp.getFolderById(folderId), ymd);
  if (!files.length) {
    out.notes.push('この週のExcel(ファイル名が ' + ymd.replace(/-/g, '') + ' で始まるもの)が取り込みフォルダに見つかりませんでした。購入者一覧は、Excelをここへ入れると載ります。');
    return out;
  }
  files.forEach(function (f) {
    var m;
    try { m = iaReadMatrix_(f); } catch (e) { out.notes.push('「' + f.getName() + '」を読めませんでした: ' + String(e && e.message || e)); return; }
    var p = iaParseFull_(m.matrix, m.sheetName);
    if (!p) { out.notes.push('「' + f.getName() + '」は、来場者リスト・配信購入リストの形式として読めませんでした。'); return; }
    var label = p.kind === 'venue' ? '来場者リスト' : '配信購入リスト';
    if (out[p.kind] !== null) { out.notes.push(label + 'が複数ありました。新しい「' + out[p.kind + 'File'] + '」を使いました。'); return; }
    out[p.kind] = p.rows;
    out[p.kind + 'File'] = f.getName();
    if (p.eventName && !out.eventName) out.eventName = p.eventName;
    if (p.nameFallback && !out.nameFallback) out.nameFallback = p.nameFallback;
    out.notes.push(label + ': ' + f.getName() + ' を読みました。');
  });
  delete out.venueFile; delete out.streamFile;
  if (out.venue === null && out.stream !== null) out.notes.push('来場者リストは見つかりませんでした。');
  if (out.stream === null && out.venue !== null) out.notes.push('配信購入リストは見つかりませんでした。');
  return out;
}

/** フォルダ以下（中のフォルダも）から、名前が ymd8桁で始まるExcel/スプレッドシートを、更新の新しい順に返す */
function iaFindFilesByDate_(folder, ymd) {
  var prefix = ymd.replace(/-/g, '');
  var found = [];
  (function walk(fd) {
    var it = fd.getFiles();
    while (it.hasNext()) {
      var f = it.next(), mt = f.getMimeType();
      if (f.getName().indexOf(prefix) !== 0) continue;
      if (mt === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
          mt === 'application/vnd.ms-excel' || mt === 'application/vnd.google-apps.spreadsheet') found.push(f);
    }
    var sub = fd.getFolders();
    while (sub.hasNext()) walk(sub.next());
  })(folder);
  found.sort(function (a, b) { return b.getLastUpdated().getTime() - a.getLastUpdated().getTime(); });
  return found;
}

function iaReadMatrix_(file) {
  var ssId = file.getId(), temp = false;
  if (file.getMimeType() !== 'application/vnd.google-apps.spreadsheet') { ssId = iaConvertToSheet_(file); temp = true; }
  try {
    var sh = SpreadsheetApp.openById(ssId).getSheets()[0];
    return {matrix: sh.getDataRange().getValues(), sheetName: sh.getName()};
  } finally {
    if (temp) { try { DriveApp.getFileById(ssId).setTrashed(true); } catch (e) {} }
  }
}

/**
 * 表から、ツールに渡す購入者一覧の行を取り出す（ツール側の読み方と同じ）
 *  venue: [扱い, 購入者, フリガナ, 枚数, キャンセル(1/0), 備考]
 *  stream: [扱い, 個数, 購入者ID, 購入者名, コメント]
 */
function iaParseFull_(matrix, sheetName) {
  var hi = iaFindHeader_(matrix, ['購入者', '扱い', '購入枚数']);
  if (hi >= 0) {
    var H = matrix[hi].map(iaClean_);
    var c = {buyer: H.indexOf('購入者'), kana: H.indexOf('フリガナ'), perf: H.indexOf('扱い'), qty: H.indexOf('購入枚数'), note: H.indexOf('備考')};
    var rows = [];
    for (var i = hi + 1; i < matrix.length; i++) {
      var r = matrix[i] || [];
      var buyer = iaClean_(r[c.buyer]);
      if (!buyer) continue;
      var cancelled = /キャンセル/i.test(iaClean_(r[c.qty]));
      rows.push([iaClean_(r[c.perf]) || '（扱いなし）', buyer, c.kana >= 0 ? iaClean_(r[c.kana]) : '',
        cancelled ? 0 : (Number(r[c.qty]) || 1), cancelled ? 1 : 0, c.note >= 0 ? iaClean_(r[c.note]) : '']);
    }
    var sn = iaClean_(sheetName).split('_')[0];
    return {kind: 'venue', rows: rows, eventName: '', nameFallback: (sn && !/^(sheet|soldlist)/i.test(sn)) ? sn : ''};
  }
  hi = iaFindHeader_(matrix, ['購入者名', 'アンケート回答']);
  if (hi >= 0) {
    var H2 = matrix[hi].map(iaClean_);
    var d = {status: H2.indexOf('ステータス'), product: H2.indexOf('商品名'), qty: H2.indexOf('購入個数'),
      id: H2.indexOf('購入者ID'), name: H2.indexOf('購入者名'), comment: H2.indexOf('コメント'), perf: H2.indexOf('アンケート回答')};
    var rows2 = [], product = '';
    for (var j = hi + 1; j < matrix.length; j++) {
      var r2 = matrix[j] || [];
      var bn = iaClean_(r2[d.name]);
      if (!bn) continue;
      if (d.status >= 0) { var st = iaClean_(r2[d.status]); if (st && st.indexOf('販売') < 0) continue; }
      if (d.product >= 0 && !product) product = iaClean_(r2[d.product]);
      rows2.push([iaClean_(r2[d.perf]) || '（扱いなし）', Number(r2[d.qty]) || 1,
        d.id >= 0 ? iaClean_(r2[d.id]) : '', bn, d.comment >= 0 ? iaClean_(r2[d.comment]) : '']);
    }
    return {kind: 'stream', rows: rows2, eventName: product, nameFallback: ''};
  }
  return null;
}
