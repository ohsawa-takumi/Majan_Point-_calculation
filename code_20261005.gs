// =====================================================
// 雀卓精算:画面の表示・記録の保存・月度順位の集計
// ルール:Mリーグ準拠
//   25000点持ち / 30000点返し / オカ+20(トップ取り)
//   ウマ 10-30 / 同点は起家に近い人が上位 / 残った供託はトップが取得
// =====================================================

var RETURN_POINT = 30000;              // 返し点
var RANK_POINTS  = [50, 10, -10, -30]; // 1〜4位の順位点(オカ+20込み)

var SHEET_RECORDS = '対局記録';
var SHEET_RESULTS = '半荘結果';
var SHEET_MONTHLY = '月度順位';

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

// ---------- 1局ごとの記録(画面から自動で呼ばれる) ----------
function saveRecord(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_RECORDS);

    // 対局IDの列がない古い形式のシートは「対局記録(旧)」として残し、新しく作り直す
    if (sheet && sheet.getLastRow() > 0 &&
        sheet.getRange(1, 1).getValue() !== RECORD_HEADER_KEY) {
      var oldName = SHEET_RECORDS + '(旧)';
      var n = 2;
      while (ss.getSheetByName(oldName)) { oldName = SHEET_RECORDS + '(旧' + n + ')'; n++; }
      sheet.setName(oldName);
      sheet = null;
    }
    if (!sheet) sheet = ss.insertSheet(SHEET_RECORDS);

    if (sheet.getLastRow() === 0 && data.headers) {
      sheet.appendRow(data.headers);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow(data.row);

    // 対局ID → No の順に並べ替え(複数の対局が混ざらないように)
    if (sheet.getLastRow() > 2) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
        .sort([{column: 1, ascending: true}, {column: 3, ascending: true}]);
    }
    drawGameDividers_(sheet, 1);
    return 'ok';
  } finally {
    lock.releaseLock();
  }
}

// ---------- 対局終了時の結果保存(画面の「対局を終了して結果を保存」) ----------
function saveGameResult(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var results = calcGameResult(data);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_RESULTS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_RESULTS);
      sheet.appendRow(RESULT_HEADERS);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, RESULT_HEADERS.length).setFontWeight('bold');
    }
    migrateSeatColumn_(sheet);

    // 同じ対局IDが既にあれば消してから書き直す(上書き保存)
    var last = sheet.getLastRow();
    if (last > 1) {
      var ids = sheet.getRange(2, 3, last - 1, 1).getValues();
      for (var i = ids.length - 1; i >= 0; i--) {
        if (String(ids[i][0]) === data.gameId) sheet.deleteRow(i + 2);
      }
    }

    var datePart = String(data.startedAt).split(' ')[0];      // 2026/10/05
    var ym = datePart.split('/');
    var month = ym[0] + '年' + ym[1] + '月';                   // 2026年10月

    var rows = results.map(function (r) {
      return [datePart, month, data.gameId, r.name, WINDS[r.seat - 1],
              r.score, r.rank, r.soten, r.rankPoint, r.total];
    });
    var start = sheet.getLastRow() + 1;
    sheet.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
    sheet.getRange(start, 2, rows.length, 1).setNumberFormat('@'); // 月度を文字のまま保持
    sheet.getRange(2, 8, sheet.getLastRow() - 1, 3).setNumberFormat('+0.0;-0.0;0.0');

    // 対局ID → 順位 の順に並べ替えて、対局ごとに区切り線を引く
    if (sheet.getLastRow() > 2) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, RESULT_HEADERS.length)
        .sort([{column: 3, ascending: true}, {column: 7, ascending: true}]);
    }
    drawGameDividers_(sheet, 3);

    updateMonthlyRanking_();
    return results;
  } finally {
    lock.releaseLock();
  }
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
    updateMonthlyRanking_();
  } finally {
    lock.releaseLock();
  }
}

function updateMonthlyRanking_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = ss.getSheetByName(SHEET_RESULTS);
  var out = ss.getSheetByName(SHEET_MONTHLY) || ss.insertSheet(SHEET_MONTHLY);
  out.clear();

  if (!src || src.getLastRow() < 2) {
    out.getRange(1, 1).setValue('まだ対局結果がありません');
    return;
  }

  // 月度 → プレイヤー → 集計値
  var values = src.getRange(2, 1, src.getLastRow() - 1, RESULT_HEADERS.length).getValues();
  var months = {};
  values.forEach(function (v) {
    var month = String(v[1]), name = String(v[3]);
    var score = Number(v[5]), rank = Number(v[6]), total = Number(v[9]);
    if (!month || !name) return;
    months[month] = months[month] || {};
    var p = months[month][name] = months[month][name] ||
      {name: name, total: 0, games: 0, rankSum: 0, ranks: [0, 0, 0, 0], best: null};
    p.total = Math.round((p.total + total) * 10) / 10;
    p.games++;
    p.rankSum += rank;
    if (rank >= 1 && rank <= 4) p.ranks[rank - 1]++;
    p.best = (p.best === null) ? score : Math.max(p.best, score);
  });

  var header = ['順位', 'プレイヤー', '合計pt', '半荘数', '平均順位',
                '1位', '2位', '3位', '4位', 'トップ率', 'ラス回避率', '最高持ち点'];
  var monthKeys = Object.keys(months).sort().reverse(); // 新しい月度が上
  var row = 1;

  monthKeys.forEach(function (month) {
    var list = Object.keys(months[month]).map(function (k) { return months[month][k]; });
    list.sort(function (a, b) {
      return (b.total - a.total) || (a.rankSum / a.games - b.rankSum / b.games);
    });

    out.getRange(row, 1).setValue(month + '度')
      .setFontWeight('bold').setFontSize(13);
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

  out.autoResizeColumns(1, header.length);
  out.setFrozenRows(0);
}

// ---------- 見た目の調整 ----------
// 対局IDごとに順番に使う背景色(隣の対局と必ず違う色になる)
var GAME_COLORS = ['#e3f2e8', '#fdf3d7', '#e2ecf8', '#f8e3e3', '#ece5f6'];

// メニュー「麻雀集計 > シートの色分けを更新」:今あるデータに色分けと区切り線を付け直す
function refreshSheetColors() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = ss.getSheetByName(SHEET_RESULTS);
  if (results) drawGameDividers_(results, 3);
  var records = ss.getSheetByName(SHEET_RECORDS);
  if (records) drawGameDividers_(records, 1);
}
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
