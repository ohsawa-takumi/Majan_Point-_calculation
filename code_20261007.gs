// =====================================================
// 雀卓精算:画面の表示・記録の保存・月度順位の集計
// ルール:Mリーグ準拠
//   25000点持ち / 30000点返し / オカ+20(トップ取り)
//   ウマ 10-30 / 同点は起家に近い人が上位 / 残った供託はトップが取得
//
// シート構成
//   月度順位                … 1枚。月ごとのブロックを下へ追加していく(古い月が上)
//   YYYY年MM月 半荘結果     … 月ごと。1対局=4行
//   YYYY年MM月 対局記録     … 月ごと。1局=1行
// =====================================================

// メンバー(固定の4人)。画面(index.html)の PLAYERS と同じにしておく
var PLAYERS = ['オオサワ', 'サクモト', 'トリイ', 'ヤマダ'];

var RETURN_POINT = 30000;              // 返し点
var RANK_POINTS  = [50, 10, -10, -30]; // 1〜4位の順位点(オカ+20込み)

var SHEET_MONTHLY  = '月度順位';
var RESULT_SUFFIX  = ' 半荘結果';      // 例:2026年10月 半荘結果
var RECORD_SUFFIX  = ' 対局記録';      // 例:2026年10月 対局記録
var MONTH_RESULT_RE = /^(\d{4}年\d{2}月) 半荘結果$/;
var MONTH_RECORD_RE = /^(\d{4}年\d{2}月) 対局記録$/;

var LEGACY_RECORDS = '対局記録';       // 以前の1枚形式(自動で月別に移行)
var LEGACY_RESULTS = '半荘結果';       // 以前の1枚形式(自動で月別に移行)

var RECORD_HEADER_KEY = '対局ID';
var WINDS = ['東', '南', '西', '北'];
var RESULT_HEADERS = ['対局ID', '日付', '月度', 'プレイヤー', '席順',
                      '最終持ち点', '順位', '素点pt', '順位点', '合計pt'];

// ウェブアプリURL(.../exec)を開いたときに画面を表示
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('雀卓精算')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// スプレッドシートを開いたときにメニューを追加
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('麻雀集計')
    .addItem('月度順位を更新', 'updateMonthlyRanking')
    .addItem('シートの色分けを更新', 'refreshSheetColors')
    .addItem('シートにロックをかける', 'protectAllSheets')
    .addItem('月度順位の自動更新をオンにする', 'enableAutoUpdate')
    .addItem('月度順位の自動更新をオフにする', 'disableAutoUpdate')
    .addSeparator()
    .addItem('過去分入力シートを開く', 'openImportSheet')
    .addItem('過去分を取り込む', 'importPastGames')
    .addToUi();
}

// ---------- 共通 ----------
// 月度の切り替わり時刻:毎月1日の朝8時。それより前(0:00〜7:59)の対局は前の月に入れる
var MONTH_START_HOUR = 8;

// 年月日・時刻から月度("2026年10月")を決める
function businessMonth_(y, mo, d, h, mi) {
  var t = new Date(y, mo - 1, d, h, mi || 0);
  t.setHours(t.getHours() - MONTH_START_HOUR); // 8時間戻した日付の月 = 月度
  return t.getFullYear() + '年' + ('0' + (t.getMonth() + 1)).slice(-2) + '月';
}

// "2026/10/05 23:38" や Date から月度を作る(時刻がなければ日中扱い)
function toMonth_(v) {
  if (v instanceof Date) {
    return businessMonth_(v.getFullYear(), v.getMonth() + 1, v.getDate(), v.getHours(), v.getMinutes());
  }
  var m = String(v).match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (!m) return '';
  var h = (m[4] !== undefined) ? Number(m[4]) : 12;
  return businessMonth_(+m[1], +m[2], +m[3], h, Number(m[5] || 0));
}

// 対局ID("G20261101-013000")の開始時刻から月度を作る
function gameIdMonth_(id) {
  var m = String(id).match(/^G(\d{4})(\d{2})(\d{2})-(\d{2})(\d{2})/);
  return m ? businessMonth_(+m[1], +m[2], +m[3], +m[4], +m[5]) : '';
}

// 月別シートを取得(なければ作る)。並び:月度順位 → 新しい月の半荘結果・対局記録 → 古い月…
function getMonthSheet_(month, suffix, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = month + suffix;
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  var ranking = getRankingSheet_();
  var pos = ranking.getIndex(); // 月度順位のすぐ右(0始まりの挿入位置)
  if (suffix === RECORD_SUFFIX) {
    var res = ss.getSheetByName(month + RESULT_SUFFIX);
    if (res) pos = res.getIndex();          // 同じ月の半荘結果の右
  } else {
    var rec = ss.getSheetByName(month + RECORD_SUFFIX);
    if (rec) pos = rec.getIndex() - 1;      // 同じ月の対局記録の左
  }
  sheet = ss.insertSheet(name, pos);
  protectSheet_(sheet);
  if (headers) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1c4531').setFontColor('#f2eee1');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getRankingSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MONTHLY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_MONTHLY, 0);
    protectSheet_(sheet);
  }
  return sheet;
}

function listMonths_(re) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .map(function (sh) { var m = sh.getName().match(re); return m ? m[1] : null; })
    .filter(function (m) { return m; })
    .sort();
}

// ---------- シートのロック ----------
// 編集しようとすると「本当に編集しますか?」と警告が出る保護をかける
// (警告のみなので、スクリプトからの書き込みは今までどおり動く)
var PROTECT_DESCRIPTION = '雀卓精算:自動で管理しているシート';

function protectSheet_(sheet) {
  if (!sheet) return;
  if (sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).length > 0) return; // 設定済み
  sheet.protect().setDescription(PROTECT_DESCRIPTION).setWarningOnly(true);
}

// 月度順位・月別の半荘結果・対局記録すべてにロックをかける
function protectAll_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  protectSheet_(ss.getSheetByName(SHEET_MONTHLY));
  listMonths_(MONTH_RESULT_RE).forEach(function (m) { protectSheet_(ss.getSheetByName(m + RESULT_SUFFIX)); });
  listMonths_(MONTH_RECORD_RE).forEach(function (m) { protectSheet_(ss.getSheetByName(m + RECORD_SUFFIX)); });
}

// メニュー「麻雀集計 > シートにロックをかける」
function protectAllSheets() {
  protectAll_();
  SpreadsheetApp.getActiveSpreadsheet().toast('月度順位・半荘結果・対局記録のシートにロックをかけました。', '麻雀集計', 5);
}

function renameToOld_(sheet, baseName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var oldName = baseName + '(旧)';
  var n = 2;
  while (ss.getSheetByName(oldName)) { oldName = baseName + '(旧' + n + ')'; n++; }
  sheet.setName(oldName);
}

// ---------- 1局ごとの記録(画面から自動で呼ばれる) ----------
function saveRecord(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    migrateLegacySheets_();
    var rows = data.rows || [data.row];
    var month = gameIdMonth_(rows[0][0]) || toMonth_(rows[0][1]); // 対局の開始時刻で月度を決める
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var existing = ss.getSheetByName(month + RECORD_SUFFIX);
    // 列の構成が変わった古いシートは「(旧)」として残し、新しい見出しで作り直す
    if (existing && existing.getLastRow() > 0 && data.headers) {
      var cur = existing.getRange(1, 1, 1, existing.getLastColumn()).getValues()[0];
      if (cur.join('|') !== data.headers.join('|')) renameToOld_(existing, month + RECORD_SUFFIX);
    }
    var sheet = getMonthSheet_(month, RECORD_SUFFIX, data.headers);
    writeRecordRows_(sheet, rows);
    return 'ok';
  } finally {
    lock.releaseLock();
  }
}

function writeRecordRows_(sheet, rows) {
  if (rows.length === 0) return;
  var start = sheet.getLastRow() + 1;
  var cols = rows[0].length;
  sheet.getRange(start, 1, rows.length, 2).setNumberFormat('@'); // 対局ID・開始日時を文字のまま保持
  sheet.getRange(start, 1, rows.length, cols).setValues(rows);
  // 対局ID → No の順に並べ替え
  if (sheet.getLastRow() > 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
      .sort([{column: 1, ascending: true}, {column: 3, ascending: true}]);
  }
  drawGameDividers_(sheet, 1);
}

// ---------- 対局終了時の結果保存(画面の「対局を終了して結果を保存」) ----------
function saveGameResult(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    migrateLegacySheets_();
    var results = saveResult_(data);
    updateMonthlyRanking_();
    return results;
  } finally {
    lock.releaseLock();
  }
}

// 1対局分を計算して、その月の半荘結果シートに書き込む
function saveResult_(data) {
  data.names.forEach(function (n) {
    if (PLAYERS.indexOf(String(n).trim()) < 0) throw new Error('メンバー外の名前があります:' + n);
  });
  var results = calcGameResult(data);
  var datePart = String(data.startedAt).split(' ')[0];  // 2026/10/05(実際の開始日)
  var month = toMonth_(data.startedAt);                  // 月度(朝8時切り替え)
  var rows = results.map(function (r) {
    return [data.gameId, datePart, month, r.name, WINDS[r.seat - 1],
            r.score, r.rank, r.soten, r.rankPoint, r.total];
  });
  writeResultRows_(getMonthSheet_(month, RESULT_SUFFIX, RESULT_HEADERS), rows, true);
  return results;
}

// 半荘結果シートに行を書き込む(replaceSameId=true なら同じ対局IDを置き換え、false なら重複を飛ばす)
function writeResultRows_(sheet, rows, replaceSameId) {
  if (rows.length === 0) return;
  var last = sheet.getLastRow();
  var existing = {};
  if (last > 1) {
    var ids = sheet.getRange(2, 1, last - 1, 1).getValues(); // A列:対局ID
    for (var i = ids.length - 1; i >= 0; i--) {
      var id = String(ids[i][0]);
      if (replaceSameId && id === String(rows[0][0])) sheet.deleteRow(i + 2);
      else existing[id] = true;
    }
  }
  if (!replaceSameId) {
    rows = rows.filter(function (r) { return !existing[String(r[0])]; });
    if (rows.length === 0) return;
  }

  var start = sheet.getLastRow() + 1;
  sheet.getRange(start, 3, rows.length, 1).setNumberFormat('@'); // C列:月度を文字のまま保持
  sheet.getRange(start, 1, rows.length, RESULT_HEADERS.length).setValues(rows);

  var n = sheet.getLastRow() - 1;
  sheet.getRange(2, 6, n, 1).setNumberFormat('#,##0');
  sheet.getRange(2, 8, n, 3).setNumberFormat('+0.0;-0.0;0.0');
  if (n > 1) {
    sheet.getRange(2, 1, n, RESULT_HEADERS.length)
      .sort([{column: 1, ascending: true}, {column: 7, ascending: true}]); // 対局ID → 順位
  }
  drawGameDividers_(sheet, 1);
  sheet.autoResizeColumns(1, RESULT_HEADERS.length);
}

// Mリーグ準拠の順位・ポイント計算
function calcGameResult(data) {
  var players = data.names.map(function (name, i) {
    return {
      name: String(name).trim(),
      score: Number(data.scores[i]),
      seat: ((i - data.chiicha) % 4 + 4) % 4 + 1 // 起家=1, 南家=2, …
    };
  });

  // 持ち点の高い順、同点なら起家に近い順
  players.sort(function (a, b) {
    return (b.score - a.score) || (a.seat - b.seat);
  });

  // 残った供託はトップへ
  players[0].score += Number(data.pot) || 0;

  return players.map(function (p, idx) {
    // 小数の誤差を避けるため 0.1pt 単位の整数で計算
    var sotenTenth = Math.round((p.score - RETURN_POINT) / 100);
    var rankTenth  = RANK_POINTS[idx] * 10;
    return {
      name: p.name,
      seat: p.seat,
      score: p.score,
      rank: idx + 1,
      soten: sotenTenth / 10,
      rankPoint: RANK_POINTS[idx],
      total: (sotenTenth + rankTenth) / 10
    };
  });
}

// ---------- 月度順位の集計 ----------
// メニュー「麻雀集計 > 月度順位を更新」からも実行できる(半荘結果を手で直した時など)
function updateMonthlyRanking() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    migrateLegacySheets_();
    updateMonthlyRanking_();
  } finally {
    lock.releaseLock();
  }
}

// 「月度順位」シートを作り直す。古い月から順に、月ごとのブロックを下へ並べる
function updateMonthlyRanking_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = getRankingSheet_();
  out.clear();

  var months = listMonths_(MONTH_RESULT_RE);
  var header = ['順位', 'プレイヤー', '合計pt', '半荘数', '平均順位',
                '1位', '2位', '3位', '4位', 'トップ率', 'ラス回避率', '最高持ち点'];
  var row = 1;

  months.forEach(function (month) {
    var src = ss.getSheetByName(month + RESULT_SUFFIX);
    if (!src || src.getLastRow() < 2) return;

    var values = src.getRange(2, 1, src.getLastRow() - 1, RESULT_HEADERS.length).getValues();
    var players = {};
    values.forEach(function (v) {
      var name = String(v[3]);
      var score = Number(v[5]), rank = Number(v[6]), total = Number(v[9]);
      if (!name) return;
      var p = players[name] = players[name] ||
        {name: name, total: 0, games: 0, rankSum: 0, ranks: [0, 0, 0, 0], best: null};
      p.total = Math.round((p.total + total) * 10) / 10;
      p.games++;
      p.rankSum += rank;
      if (rank >= 1 && rank <= 4) p.ranks[rank - 1]++;
      p.best = (p.best === null) ? score : Math.max(p.best, score);
    });
    var list = Object.keys(players).map(function (k) { return players[k]; });
    if (list.length === 0) return;
    list.sort(function (a, b) {
      return (b.total - a.total) || (a.rankSum / a.games - b.rankSum / b.games);
    });

    out.getRange(row, 1).setValue(month + '度').setFontWeight('bold').setFontSize(13);
    row++;
    out.getRange(row, 1, 1, header.length).setValues([header])
      .setFontWeight('bold').setBackground('#1c4531').setFontColor('#f2eee1');
    row++;

    var rows = [], prevTotal = null, prevRank = 0;
    list.forEach(function (p, i) {
      var r = (p.total === prevTotal) ? prevRank : i + 1; // 同ポイントは同順位
      prevTotal = p.total; prevRank = r;
      rows.push([r, p.name, p.total, p.games,
                 Math.round(p.rankSum / p.games * 100) / 100,
                 p.ranks[0], p.ranks[1], p.ranks[2], p.ranks[3],
                 p.ranks[0] / p.games, 1 - p.ranks[3] / p.games, p.best]);
    });
    out.getRange(row, 1, rows.length, header.length).setValues(rows);
    out.getRange(row, 3, rows.length, 1).setNumberFormat('+0.0;-0.0;0.0').setFontWeight('bold');
    out.getRange(row, 5, rows.length, 1).setNumberFormat('0.00');
    out.getRange(row, 10, rows.length, 2).setNumberFormat('0.0%');
    out.getRange(row, 12, rows.length, 1).setNumberFormat('#,##0');
    row += rows.length + 1; // 月度の間に1行空ける
  });

  if (row === 1) out.getRange(1, 1).setValue('まだ対局結果がありません');
  protectAll_(); // ロックが外れているシートがあればかけ直す
  out.autoResizeColumns(1, header.length);
}

// ---------- 以前の形式からの移行(該当シートがある時だけ自動で動く) ----------
function migrateLegacySheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1枚形式の「半荘結果」→ 月別に振り分け、元は「半荘結果(旧)」に改名
  var legacyRes = ss.getSheetByName(LEGACY_RESULTS);
  if (legacyRes) {
    migrateSeatColumn_(legacyRes);
    var last = legacyRes.getLastRow();
    if (last > 1) {
      var values = legacyRes.getRange(2, 1, last - 1, RESULT_HEADERS.length).getValues();
      var byMonth = {};
      var oldOrder = legacyRes.getRange(1, 1).getValue() !== '対局ID'; // 旧並び:日付,月度,対局ID
      values.forEach(function (v) {
        var id = oldOrder ? v[2] : v[0];
        var date = oldOrder ? v[0] : v[1];
        var month = gameIdMonth_(id) || String(oldOrder ? v[1] : v[2]) || toMonth_(date);
        if (!month || !id) return;
        var row = [id, date, month].concat(v.slice(3));
        (byMonth[month] = byMonth[month] || []).push(row);
      });
      Object.keys(byMonth).forEach(function (month) {
        writeResultRows_(getMonthSheet_(month, RESULT_SUFFIX, RESULT_HEADERS), byMonth[month], false);
      });
    }
    renameToOld_(legacyRes, LEGACY_RESULTS);
  }

  // 1枚形式の「対局記録」→ 月別に振り分け、元は「対局記録(旧)」に改名
  var legacyRec = ss.getSheetByName(LEGACY_RECORDS);
  if (legacyRec) {
    var lastR = legacyRec.getLastRow(), cols = legacyRec.getLastColumn();
    if (lastR > 1 && legacyRec.getRange(1, 1).getValue() === RECORD_HEADER_KEY) {
      var headers = legacyRec.getRange(1, 1, 1, cols).getValues()[0];
      var recs = legacyRec.getRange(2, 1, lastR - 1, cols).getValues();
      var recByMonth = {};
      recs.forEach(function (v) {
        var month = gameIdMonth_(v[0]) || toMonth_(v[1]);
        if (!month || !v[0]) return;
        if (v[1] instanceof Date) {
          v[1] = Utilities.formatDate(v[1], Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
        }
        (recByMonth[month] = recByMonth[month] || []).push(v);
      });
      Object.keys(recByMonth).forEach(function (month) {
        writeRecordRows_(getMonthSheet_(month, RECORD_SUFFIX, headers), recByMonth[month]);
      });
    }
    renameToOld_(legacyRec, LEGACY_RECORDS);
  }
}

// ---------- 見た目の調整 ----------
// 対局IDごとに順番に使う背景色(隣の対局と必ず違う色になる)
var GAME_COLORS = ['#e3f2e8', '#fdf3d7', '#e2ecf8', '#f8e3e3', '#ece5f6'];

// 対局IDが変わる位置に太線を引き、対局ごとに背景色を交互に付ける
function drawGameDividers_(sheet, idCol) {
  var last = sheet.getLastRow();
  if (last < 2) return;
  var cols = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, last - 1, cols);
  data.setBorder(false, false, false, false, false, false);

  var ids = sheet.getRange(2, idCol, last - 1, 1).getValues();
  var backgrounds = [];
  var band = 0;
  for (var i = 0; i < ids.length; i++) {
    if (i > 0 && String(ids[i][0]) !== String(ids[i - 1][0])) band++;
    var color = GAME_COLORS[band % GAME_COLORS.length];
    var row = [];
    for (var c = 0; c < cols; c++) row.push(color);
    backgrounds.push(row);
    // 次の行で対局が変わる(または最終行)なら下に太線
    if (i === ids.length - 1 || String(ids[i][0]) !== String(ids[i + 1][0])) {
      sheet.getRange(i + 2, 1, 1, cols)
        .setBorder(null, null, true, null, null, null, '#1c4531', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }
  data.setBackgrounds(backgrounds);
}

// 以前の形式(席順(起家=1) に 1〜4)を「席順」+ 東南西北 に置き換える
function migrateSeatColumn_(sheet) {
  var head = sheet.getRange(1, 5);
  if (head.getValue() === '席順') return;
  head.setValue('席順');
  var last = sheet.getLastRow();
  if (last < 2) return;
  var range = sheet.getRange(2, 5, last - 1, 1);
  var vals = range.getValues().map(function (v) {
    var n = Number(v[0]);
    return [(n >= 1 && n <= 4) ? WINDS[n - 1] : v[0]];
  });
  range.setValues(vals);
}

// メニュー「麻雀集計 > シートの色分けを更新」:今あるデータに色分けと区切り線を付け直す
function refreshSheetColors() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  listMonths_(MONTH_RESULT_RE).forEach(function (m) {
    drawGameDividers_(ss.getSheetByName(m + RESULT_SUFFIX), 1);
  });
  listMonths_(MONTH_RECORD_RE).forEach(function (m) {
    drawGameDividers_(ss.getSheetByName(m + RECORD_SUFFIX), 1);
  });
}

// =====================================================
// 過去分の取り込み
//   「過去分入力」シートに1対局=1行で最終持ち点を入力 →
//   メニュー「麻雀集計 > 過去分を取り込む」で半荘結果・月度順位に反映
// =====================================================
var SHEET_IMPORT = '過去分入力';
var IMPORT_HEADERS = ['日付', '開始時刻(任意)',
                      '東家(起家)', '持ち点', '南家', '持ち点', '西家', '持ち点', '北家', '持ち点',
                      '残り供託(任意)', '取込状況'];
var TOTAL_POINTS = 100000; // 25000点×4人

// メニュー「過去分入力シートを開く」:なければ作る
function openImportSheet() {
  var sheet = getImportSheet_();
  sheet.activate();
}

function getImportSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_IMPORT);
  if (sheet) { applyNameDropdown_(sheet); return sheet; }
  sheet = ss.insertSheet(SHEET_IMPORT, ss.getSheets().length);
  sheet.getRange(1, 1, 1, IMPORT_HEADERS.length).setValues([IMPORT_HEADERS])
    .setFontWeight('bold').setBackground('#1c4531').setFontColor('#f2eee1');
  sheet.setFrozenRows(1);
  sheet.getRange('A2:A').setNumberFormat('yyyy/mm/dd');
  sheet.getRange('B2:B').setNumberFormat('@');          // 時刻は「23:30」のように文字で
  sheet.getRange('D2:D').setNumberFormat('#,##0');
  sheet.getRange('F2:F').setNumberFormat('#,##0');
  sheet.getRange('H2:H').setNumberFormat('#,##0');
  sheet.getRange('J2:J').setNumberFormat('#,##0');
  sheet.getRange(1, 1).setNote('例:2026/09/20\n開始時刻を空欄にすると昼12:00開始として扱います。\n0:00〜7:59開始は前日の対局(月初なら前月分)になります。');
  sheet.getRange(1, 3).setNote('起家(東1局の親)から順に、東家・南家・西家・北家の名前と最終持ち点を入力します。\n同点の順位決めに使います。');
  sheet.getRange(1, 11).setNote('終局時に卓に残ったリーチ棒の点数(例:1000)。トップに加算されます。');
  sheet.getRange(1, 12).setNote('取り込むとスクリプトが書き込みます。「✅」の行は次回以降スキップされます。');
  applyNameDropdown_(sheet);
  sheet.autoResizeColumns(1, IMPORT_HEADERS.length);
  return sheet;
}

// 名前の列(東家・南家・西家・北家)を4人のプルダウンにする
function applyNameDropdown_(sheet) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(PLAYERS, true)
    .setAllowInvalid(false)
    .setHelpText('メンバー(' + PLAYERS.join('・') + ')から選んでください')
    .build();
  [3, 5, 7, 9].forEach(function (col) {
    sheet.getRange(2, col, sheet.getMaxRows() - 1, 1).setDataValidation(rule);
  });
}

// メニュー「過去分を取り込む」
function importPastGames() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    migrateLegacySheets_();
    var sheet = getImportSheet_();
    var last = sheet.getLastRow();
    if (last < 2) {
      ss.toast('「過去分入力」シートに対局を入力してから実行してください。', '麻雀集計', 6);
      sheet.activate();
      return;
    }

    var values = sheet.getRange(2, 1, last - 1, IMPORT_HEADERS.length).getValues();
    var usedIds = collectGameIds_();
    var ok = 0, ng = 0;

    values.forEach(function (v, i) {
      var statusCell = sheet.getRange(i + 2, 12);
      if (String(v[11]).indexOf('✅') === 0) return;                 // 取込済み
      var names = [v[2], v[4], v[6], v[8]].map(function (x) { return String(x).trim(); });
      var hasAny = v[0] || names.some(function (n) { return n; });
      if (!hasAny) return;                                           // 空行

      var err = [];
      var date = parseImportDate_(v[0]);
      if (!date) err.push('日付');
      var time = parseImportTime_(v[1]);
      if (!time) err.push('開始時刻');
      if (names.some(function (n) { return PLAYERS.indexOf(n) < 0; })) err.push('名前(メンバー外)');
      else if (names.filter(function (n, k) { return names.indexOf(n) !== k; }).length) err.push('名前(重複)');
      var scores = [v[3], v[5], v[7], v[9]].map(function (x) {
        return (x === '' || x === null) ? NaN : Number(String(x).replace(/,/g, ''));
      });
      if (scores.some(function (x) { return isNaN(x); })) err.push('持ち点');
      var pot = (v[10] === '' || v[10] === null) ? 0 : Number(String(v[10]).replace(/,/g, ''));
      if (isNaN(pot)) err.push('残り供託');
      if (err.length) {
        statusCell.setValue('⚠ 入力を確認してください:' + err.join('・'));
        ng++;
        return;
      }
      var sum = scores.reduce(function (a, b) { return a + b; }, 0) + pot;
      if (sum !== TOTAL_POINTS) {
        statusCell.setValue('⚠ 持ち点+残り供託の合計が ' + sum.toLocaleString() + ' 点です(100,000点になるように)');
        ng++;
        return;
      }

      // 対局IDは「G + 開始日時」。同じ時刻が重なったら秒をずらす
      var base = 'G' + date.y + pad2_(date.m) + pad2_(date.d) + '-' + pad2_(time.h) + pad2_(time.mi);
      var sec = 0, id = base + pad2_(sec);
      while (usedIds[id]) { sec++; id = base + pad2_(sec); }
      usedIds[id] = true;

      saveResult_({
        gameId: id,
        startedAt: date.y + '/' + pad2_(date.m) + '/' + pad2_(date.d) + ' ' + pad2_(time.h) + ':' + pad2_(time.mi),
        names: names,
        scores: scores,
        chiicha: 0,   // 東家=起家
        pot: pot
      });
      statusCell.setValue('✅ 取込済み(' + id + ')');
      ok++;
    });

    if (ok > 0) updateMonthlyRanking_();
    ss.toast('取り込み ' + ok + ' 件' + (ng ? ' / 要確認 ' + ng + ' 件(取込状況の列を見てください)' : ''), '麻雀集計', 8);
  } finally {
    lock.releaseLock();
  }
}

function pad2_(n) { return ('0' + n).slice(-2); }

function parseImportDate_(v) {
  if (v instanceof Date && !isNaN(v)) return {y: v.getFullYear(), m: v.getMonth() + 1, d: v.getDate()};
  var m = String(v).trim().match(/^(\d{4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})/);
  return m ? {y: +m[1], m: +m[2], d: +m[3]} : null;
}

function parseImportTime_(v) {
  if (v === '' || v === null) return {h: 12, mi: 0};     // 未入力は昼12:00扱い
  if (v instanceof Date && !isNaN(v)) return {h: v.getHours(), mi: v.getMinutes()};
  var m = String(v).trim().match(/^(\d{1,2})[:：時](\d{1,2})?/);
  if (!m) return null;
  var h = +m[1], mi = +(m[2] || 0);
  return (h < 24 && mi < 60) ? {h: h, mi: mi} : null;
}

// 既存の全対局IDを集める(取り込み時のID重複防止)
function collectGameIds_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ids = {};
  listMonths_(MONTH_RESULT_RE).forEach(function (m) {
    var sh = ss.getSheetByName(m + RESULT_SUFFIX);
    if (sh.getLastRow() < 2) return;
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
      .forEach(function (r) { ids[String(r[0])] = true; });
  });
  return ids;
}

// =====================================================
// シートを直接編集したときの自動再計算
//   「YYYY年MM月 半荘結果」で プレイヤー・席順・最終持ち点 を書き換えると、
//   その対局の 順位・素点pt・順位点・合計pt を計算し直し、月度順位も更新する
//   (順位〜合計ptを手で直した場合は、そのまま月度順位だけ更新する)
// =====================================================
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (!MONTH_RESULT_RE.test(sheet.getName())) return;

  var r1 = e.range.getRow(), r2 = e.range.getLastRow();
  var c1 = e.range.getColumn(), c2 = e.range.getLastColumn();
  if (r2 < 2) return;                       // 見出し行だけの編集は無視
  r1 = Math.max(r1, 2);

  var ss = sheet.getParent();
  try {
    var touchesInput = c1 <= 6 && c2 >= 4;  // D:プレイヤー / E:席順 / F:最終持ち点
    var msgs = [];
    if (touchesInput) {
      var ids = sheet.getRange(r1, 1, r2 - r1 + 1, 1).getValues()
        .map(function (r) { return String(r[0]); })
        .filter(function (id, i, arr) { return id && arr.indexOf(id) === i; });
      ids.forEach(function (id) {
        var m = recalcGame_(sheet, id);
        if (m) msgs.push(m);
      });
    }
    updateMonthlyRanking_();
    ss.toast(msgs.length ? msgs.join('\n') : '順位・ptと月度順位を更新しました。', '麻雀集計', msgs.length ? 10 : 4);
  } catch (err) {
    ss.toast('自動再計算でエラー:' + err.message + '\nメニュー「麻雀集計 > 月度順位を更新」を試してください。', '麻雀集計', 10);
  }
}

// 指定した対局IDの4行を、シート上の 席順・最終持ち点 から計算し直す
function recalcGame_(sheet, gameId) {
  var last = sheet.getLastRow();
  var values = sheet.getRange(2, 1, last - 1, RESULT_HEADERS.length).getValues();
  var rows = [];
  values.forEach(function (v, i) {
    if (String(v[0]) === gameId) rows.push({row: i + 2, v: v});
  });
  if (rows.length !== 4) return '⚠ ' + gameId + ':行が4つではないため計算し直せませんでした。';

  var players = rows.map(function (r) {
    return {
      row: r.row,
      name: String(r.v[3]).trim(),
      seat: WINDS.indexOf(String(r.v[4]).trim()) + 1,
      score: Number(String(r.v[5]).replace(/,/g, ''))
    };
  });
  if (players.some(function (p) { return p.seat < 1; })) return '⚠ ' + gameId + ':席順は 東・南・西・北 で入力してください。';
  if (players.some(function (p) { return isNaN(p.score) || String(p.score) === ''; })) return '⚠ ' + gameId + ':最終持ち点が数値ではありません。';
  if (players.some(function (p) { return PLAYERS.indexOf(p.name) < 0; })) return '⚠ ' + gameId + ':メンバー外の名前があります。';

  // 持ち点の高い順、同点なら起家に近い順(最終持ち点には残り供託が含まれている前提)
  players.sort(function (a, b) { return (b.score - a.score) || (a.seat - b.seat); });
  players.forEach(function (p, idx) {
    var sotenTenth = Math.round((p.score - RETURN_POINT) / 100);
    sheet.getRange(p.row, 7, 1, 4).setValues([[
      idx + 1, sotenTenth / 10, RANK_POINTS[idx], (sotenTenth + RANK_POINTS[idx] * 10) / 10
    ]]);
  });

  // 対局ID → 順位 で並べ直し、色分けし直す
  sheet.getRange(2, 1, last - 1, RESULT_HEADERS.length)
    .sort([{column: 1, ascending: true}, {column: 7, ascending: true}]);
  drawGameDividers_(sheet, 1);

  var sum = players.reduce(function (a, p) { return a + p.score; }, 0);
  if (sum !== TOTAL_POINTS) {
    return '⚠ ' + gameId + ':計算し直しましたが、4人の持ち点合計が ' + sum.toLocaleString() + ' 点です(通常は100,000点)。';
  }
  return '';
}

// =====================================================
// 月度順位の自動更新(インストール型トリガー)
//   ・行の削除/追加、シートの削除・追加など、onEdit では拾えない変更を検知して更新
//   ・毎朝8時(月度の切り替わり時刻)にも念のため更新
//   メニュー「月度順位の自動更新をオンにする」を1回実行すると設定される
// =====================================================
var AUTO_HANDLERS = ['autoUpdateOnChange_', 'autoUpdateDaily_'];

function enableAutoUpdate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  removeAutoTriggers_();
  ScriptApp.newTrigger('autoUpdateOnChange_').forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger('autoUpdateDaily_').timeBased().everyDays(1).atHour(MONTH_START_HOUR).create();
  updateMonthlyRanking();
  ss.toast('自動更新をオンにしました。行や対局を削除したときや、毎朝8時に月度順位が自動で更新されます。', '麻雀集計', 8);
}

function disableAutoUpdate() {
  removeAutoTriggers_();
  SpreadsheetApp.getActiveSpreadsheet().toast('自動更新をオフにしました。', '麻雀集計', 5);
}

function removeAutoTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (AUTO_HANDLERS.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
}

// 構造の変更(行の削除・挿入、シートの削除・追加など)を検知
// セルの値の編集は onEdit が担当するので、ここでは扱わない(二重実行を防ぐ)
function autoUpdateOnChange_(e) {
  var type = e && e.changeType;
  if (type === 'EDIT' || type === 'FORMAT') return;
  runRankingUpdate_();
}

function autoUpdateDaily_() {
  runRankingUpdate_();
}

function runRankingUpdate_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return; // 他の更新が動いている時は任せる
  try {
    updateMonthlyRanking_();
  } finally {
    lock.releaseLock();
  }
}
