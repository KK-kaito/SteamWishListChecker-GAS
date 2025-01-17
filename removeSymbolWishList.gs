// 文字列比較用に記号など削除
function removeSymbolWishListMain() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('sheetID');
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheetSteamWishList = spreadsheet.getSheetByName('SteamWishList');

  const titleColumn = getColumnName(sheetSteamWishList, 'タイトル');
  const renamedWLColumn = getColumnName(sheetSteamWishList, '比較用タイトル');

  const wishListRange = sheetSteamWishList.getRange(2, titleColumn, sheetSteamWishList.getLastRow());
  var wishListValueArray = wishListRange.getValues();
  var wishListNumRows = wishListValueArray.length;
  var targetCount = 0;
  var renameValue = '';

  for (var i = 0; i < wishListNumRows; i++) {
    var wishListValue = wishListValueArray[i][0].toString();
    renameValue = sheetSteamWishList.getRange(i + 2, renamedWLColumn).getValue().toString();
    // タイトルが空か、すでに本処理を実行後ならスキップ
    if (!wishListValue || renameValue) {
      continue;
    }
    wishListValue = removeSymbol(wishListValue);
    sheetSteamWishList.getRange(i + 2, renamedWLColumn).setNumberFormat('@');
    sheetSteamWishList.getRange(i + 2, renamedWLColumn).setValue(wishListValue);
    targetCount++;
  }
  Logger.log(`Check completed. wishlist targetCount = ${targetCount}`);

  function removeSymbol(str) {
    return str.replace(/[^0-9A-Za-z\u3041-\u3096\u30A1-\u30FA\u4E00-\u9FFF\uFF41-\uFF5A\uFF21-\uFF3A&]/g, '');
  }

  // カラム名取得
  function getColumnName(sheet, columnName) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var columnIndex = headers.indexOf(columnName) + 1;
    return columnIndex;
  }
}
