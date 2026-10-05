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

var RETURN_POINT = 30000;              // 返し点
var RANK_POINTS  = [50, 10, -10, -30]; // 1〜4位の順位点(オカ+20込み)

var SHEET_MONTHLY  = '月度順位';
var RESULT_SUFFIX  = ' 半荘結果';      // 例:2026年10月 半荘結果
var RECORD_SUFFIX  = ' 対局記録';      // 例:2026年10月 対局記録
var MONTH_RESULT_RE = /^(\d{4}年\d{2}月) 半荘結果$/;
var MONTH_RECORD_RE = /^(\d{4}年\d{2}月) 対局記録$/;
var MONTH_RANKING_RE = /^\d{4}年\d{2}月 順位$/; // 前のバージョンで作った月別順位シート(削除対象)

var LEGACY_RECORDS = '対局記録';       // 以前の1枚形式(自動で月別に移行)
var LEGACY_RESULTS = '半荘結果';       // 以前の1枚形式(自動で月別に移行)

var RECORD_HEADER_KEY = '対局ID';
var WINDS = ['東', '南', '西', '北'];
var RESULT_HEADERS = ['日付', '月度', '対局ID', 'プレイヤー', '席順',
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
  if (headers) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1c4531').setFontColor('#f2eee1');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getRankingSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_MONTHLY) || ss.insertSheet(SHEET_MONTHLY, 0);
}

function listMonths_(re) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .map(function (sh) { var m = sh.getName().match(re); return m ? m[1] : null; })
    .filter(function (m) { return m; })
    .sort();
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
    var month = gameIdMonth_(data.row[0]) || toMonth_(data.row[1]); // 対局の開始時刻で月度を決める
    var sheet = getMonthSheet_(month, RECORD_SUFFIX, data.headers);
    writeRecordRows_(sheet, [data.row]);
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
    var results = calcGameResult(data);

    var datePart = String(data.startedAt).split(' ')[0];  // 2026/10/05(実際の開始日)
    var month = toMonth_(data.startedAt);                  // 月度(朝8時切り替え)

    var rows = results.map(function (r) {
      return [datePart, month, data.gameId, r.name, WINDS[r.seat - 1],
              r.score, r.rank, r.soten, r.rankPoint, r.total];
    });
    writeResultRows_(getMonthSheet_(month, RESULT_SUFFIX, RESULT_HEADERS), rows, true);
    updateMonthlyRanking_();
    return results;
  } finally {
    lock.releaseLock();
  }
}

// 半荘結果シートに行を書き込む(replaceSameId=true なら同じ対局IDを置き換え、false なら重複を飛ばす)
function writeResultRows_(sheet, rows, replaceSameId) {
  if (rows.length === 0) return;
  var last = sheet.getLastRow();
  var existing = {};
  if (last > 1) {
    var ids = sheet.getRange(2, 3, last - 1, 1).getValues();
    for (var i = ids.length - 1; i >= 0; i--) {
      var id = String(ids[i][0]);
      if (replaceSameId && id === String(rows[0][2])) sheet.deleteRow(i + 2);
      else existing[id] = true;
    }
  }
  if (!replaceSameId) {
    rows = rows.filter(function (r) { return !existing[String(r[2])]; });
    if (rows.length === 0) return;
  }

  var start = sheet.getLastRow() + 1;
  sheet.getRange(start, 2, rows.length, 1).setNumberFormat('@'); // 月度を文字のまま保持
  sheet.getRange(start, 1, rows.length, RESULT_HEADERS.length).setValues(rows);

  var n = sheet.getLastRow() - 1;
  sheet.getRange(2, 6, n, 1).setNumberFormat('#,##0');
  sheet.getRange(2, 8, n, 3).setNumberFormat('+0.0;-0.0;0.0');
  if (n > 1) {
    sheet.getRange(2, 1, n, RESULT_HEADERS.length)
      .sort([{column: 3, ascending: true}, {column: 7, ascending: true}]);
  }
  drawGameDividers_(sheet, 3);
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
      values.forEach(function (v) {
        var month = gameIdMonth_(v[2]) || String(v[1]) || toMonth_(v[0]);
        if (!month || !v[2]) return;
        v[1] = month;
        (byMonth[month] = byMonth[month] || []).push(v);
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

  // 前のバージョンで作った「YYYY年MM月 順位」シートは「月度順位」1枚にまとめたので削除
  ss.getSheets().forEach(function (sh) {
    if (MONTH_RANKING_RE.test(sh.getName())) ss.deleteSheet(sh);
  });
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
    drawGameDividers_(ss.getSheetByName(m + RESULT_SUFFIX), 3);
  });
  listMonths_(MONTH_RECORD_RE).forEach(function (m) {
    drawGameDividers_(ss.getSheetByName(m + RECORD_SUFFIX), 1);
  });
}
