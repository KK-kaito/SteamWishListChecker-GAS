// 文字列比較用に記号など削除
function removeSymbolPlayniteMain() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('sheetID');
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheetPlayniteLibrary = spreadsheet.getSheetByName('PlayniteLibrary');

  if (!headerCheck(sheetPlayniteLibrary, '比較用タイトル')) {
    // ヘッダー行になければ追加
    const lastColumn = sheetPlayniteLibrary.getLastColumn();
    sheetPlayniteLibrary.getRange(1, lastColumn + 1).setValue('比較用タイトル');
  }

  const titleColumn = getColumnName(sheetPlayniteLibrary, 'タイトル');
  const renamedPNColumn = getColumnName(sheetPlayniteLibrary, '比較用タイトル');

  const playniteRange = sheetPlayniteLibrary.getRange(2, titleColumn, sheetPlayniteLibrary.getLastRow());
  var playniteValueArray = playniteRange.getValues();
  var playniteNumRows = playniteValueArray.length;
  var targetCount = 0;
  var renameValue = '';

  for (var i = 0; i < playniteNumRows; i++) {
    var playniteValue = playniteValueArray[i][0].toString();
    renameValue = sheetPlayniteLibrary.getRange(i + 2, renamedPNColumn).getValue().toString();
    // タイトルが空か、すでに本処理を実行後ならスキップ
    if (!playniteValue || renameValue) {
      continue;
    }
    playniteValue = removeSymbol(playniteValue);
    sheetPlayniteLibrary.getRange(i + 2, renamedPNColumn).setNumberFormat('@');
    sheetPlayniteLibrary.getRange(i + 2, renamedPNColumn).setValue(playniteValue);
    targetCount++;
  }
  Logger.log(`Check completed. playnite targetCount = ${targetCount}`);

  function removeSymbol(str) {
    return str.replace(/[^0-9A-Za-z\u3041-\u3096\u30A1-\u30FA\u4E00-\u9FFF\uFF41-\uFF5A\uFF21-\uFF3A&]/g, '');
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
