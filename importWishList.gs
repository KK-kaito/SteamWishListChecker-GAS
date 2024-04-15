// SteamWishList取り込み用
function importWishListMain() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('sheetID');
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheetSteamWishList = spreadsheet.getSheetByName('SteamWishList');
  const sheetPasteWishList = spreadsheet.getSheetByName('PasteWishList');

  if (!headerCheck(sheetSteamWishList, 'appid')) {
    // ヘッダー行のデータを定義
    const headers = ['appid', 'priority', 'added', 'status', 'タイトル', '比較用タイトル', 'ストアURL'];
    sheetSteamWishList.appendRow(headers);
  }
  const appidColumn = getColumnName(sheetSteamWishList, 'appid');
  const priorityColumn = getColumnName(sheetSteamWishList, 'priority');
  const addedColumn = getColumnName(sheetSteamWishList, 'added');
  const statusColumn = getColumnName(sheetSteamWishList, 'status');

  // PasteWishListシートのデータを取得
  const range = sheetPasteWishList.getDataRange();
  var values = range.getValues();
  var pasteData = values[0][0] // 左上セルしか使わない場合
  if (pasteData.slice(-1) === ';') {
    pasteData = pasteData.slice(0, -1);
  }

  pasteData = JSON.parse(pasteData);
  Logger.log(pasteData);

  pasteData.forEach((obj) => {
    var lastRow = sheetSteamWishList.getLastRow() + 1;
    // appidが重複している場合はすでにインポート済みとしてスルー
    if (!isValueInColumn(sheetSteamWishList, obj['appid'], 1)) {
      Object.keys(obj).forEach((key, index) => {
        switch (key) {
          case 'appid':
            sheetSteamWishList.getRange(lastRow, appidColumn).setNumberFormat('@');
            sheetSteamWishList.getRange(lastRow, appidColumn).setValue(obj[key]);
            Logger.log(`Import appid = ${obj[key]}`);
            break;
          case 'priority':
            sheetSteamWishList.getRange(lastRow, priorityColumn).setNumberFormat('@');
            sheetSteamWishList.getRange(lastRow, priorityColumn).setValue(obj[key]);
            break;
          case 'added':
            sheetSteamWishList.getRange(lastRow, addedColumn).setNumberFormat('@');
            sheetSteamWishList.getRange(lastRow, addedColumn).setValue(obj[key]);
            break;
          default:
            Logger.log(`Import failed. Invalid object key. appid = ${obj['appid']}`);
        }
      });
      // 最後に「未実行」列を追加
      sheetSteamWishList.getRange(lastRow, statusColumn).setValue('未実行');
    }
  });

  // 指定された列に特定の値が存在するかどうかをチェック
  function isValueInColumn(sheet, value, columnIndex) {
    var data = sheet.getRange(1, columnIndex, sheet.getLastRow(), 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0].toString() === value.toString()) {
        Logger.log(`check true ${value}`);
        return true; // 値が見つかった場合はtrue
      }
    }
    Logger.log(`check false ${value}`);
    return false; // 値が見つからなかった場合はfalse
  }

  // ヘッダー行の有無チェック
  function headerCheck(sheet, str) {
    if (sheet.getLastColumn() === 0) {
      return false;
    }
    const headerValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    for (var i = 0; i < headerValues.length; i++) {
      if (headerValues[i].toString() === str) {
        return true;
      }
    }
  }

  // カラム名取得
  function getColumnName(sheet, columnName) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var columnIndex = headers.indexOf(columnName) + 1;
    return columnIndex;
  }
}
