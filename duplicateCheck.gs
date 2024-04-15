// 別プラットフォームで所持済みのゲームがwishlistに存在するかをチェック
function duplicateCheckMain() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('sheetID');
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheetSteamWishList = spreadsheet.getSheetByName('SteamWishList');
  const sheetPlayniteLibrary = spreadsheet.getSheetByName('PlayniteLibrary');
  const appidColumn = getColumnName(sheetSteamWishList,'appid');
  const titleColumn = getColumnName(sheetSteamWishList,'タイトル');
  const renamedWLColumn = getColumnName(sheetSteamWishList, '比較用タイトル');
  const storeURLColumn = getColumnName(sheetSteamWishList, 'ストアURL');
  const renamedPNColumn = getColumnName(sheetPlayniteLibrary, '比較用タイトル');

  // 記号などを除外した文字列（比較用）を使用
  const wishListRange = sheetSteamWishList.getRange(2, renamedWLColumn, sheetSteamWishList.getLastRow());
  const playniteRange = sheetPlayniteLibrary.getRange(2, renamedPNColumn, sheetPlayniteLibrary.getLastRow());
  var wishListValueArray = wishListRange.getValues();
  var playniteValueArray = playniteRange.getValues();
  var wishListNumRows = wishListValueArray.length;
  var playniteNumRows = playniteValueArray.length;
  var targetCount = 0;

  // 2つのシートのゲームタイトルの値を比較し、重複を検出
  for (var i = 0; i < wishListNumRows; i++) {
    // タイトルが空ならスキップ
    if (!wishListValueArray[i][0]) {
      continue;
    }
    var wishListValue = wishListValueArray[i][0].toString();
    for (var j = 0; j < playniteNumRows; j++) {
      var playniteValue = playniteValueArray[j][0].toString();
      if (wishListValue.toUpperCase() === playniteValue.toUpperCase()) {
        // シート間で重複していたら色変え
        sheetSteamWishList.getRange(i + 2, titleColumn).setBackground("orange");
        // ストアページURLを追記(ウィッシュリストボタン押しに行く用)
        var appid = sheetSteamWishList.getRange(i + 2, appidColumn).getValue().toString();
        sheetSteamWishList.getRange(i + 2, storeURLColumn).setValue(`https://store.steampowered.com/app/${appid}`);
        targetCount++;
        continue;
      }
    }
  }
  Logger.log(`Check completed. targetCount = ${targetCount}`);

  // カラム名取得
  function getColumnName(sheet, columnName) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var columnIndex = headers.indexOf(columnName) + 1;
    return columnIndex;
  }
}
