// 麻雀点数画面(雀卓精算)を表示し、記録をこのスプレッドシートに保存する

// ウェブアプリURL(.../exec)を開いたときに画面を表示
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('雀卓精算')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// 画面から呼ばれ、1局分の記録を「対局記録」シートに追加する
function saveRecord(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('対局記録');
    if (!sheet) {
      sheet = ss.insertSheet('対局記録');
    }
    if (sheet.getLastRow() === 0 && data.headers) {
      sheet.appendRow(data.headers);
    }
    sheet.appendRow(data.row);
    if (sheet.getLastRow() > 2) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
        .sort({column: 1, ascending: true});
    }
    return 'ok';
  } finally {
    lock.releaseLock();
  }
}
