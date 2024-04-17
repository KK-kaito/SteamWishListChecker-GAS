//  ゲームタイトル取得とシートへの反映
function getAppTitleMain() {
  // wishlistを取り込んだシートを取得
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('sheetID');
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheetSteamWishList = spreadsheet.getSheetByName('SteamWishList');
  const appidColumn = getColumnName(sheetSteamWishList, 'appid');
  const priorityColumn = getColumnName(sheetSteamWishList, 'priority');
  const addedColumn = getColumnName(sheetSteamWishList, 'added');
  const statusColumn = getColumnName(sheetSteamWishList, 'status');
  const titleColumn = getColumnName(sheetSteamWishList, 'タイトル');

  // 取得したwishlist件数を絞る
  const wishListDataArray = getWishListDataArray(sheetSteamWishList).slice(0, 100);

  wishListDataArray.forEach(wishListData => {
    wishListData.getAppTitle();
    // APIのlimitに引っかかるのがこわいのでとりあえず1秒間隔で
    Utilities.sleep(1000);
    wishListData.refreshSheet(sheetSteamWishList, statusColumn, titleColumn);
    Logger.log(`appid = ${wishListData.appid} , title = ${wishListData.title}`);
  });

  function getWishListDataArray(sheet) {
    // シートのデータを取得
    const range = sheet.getDataRange();
    const values = range.getValues();

    // データをWishListDataクラスのインスタンスに変換
    var wishListDataArray = [];
    for (var i = 1; i < values.length; i++) { // ヘッダーを除く
      var rowNumber = i + 1; // 行番号は1からカウント
      var rowData = values[i];
      var wishListData = new WishListData(rowData[appidColumn - 1], rowData[priorityColumn - 1], rowData[addedColumn - 1], rowData[statusColumn - 1]);
      // ステータスが未実行のもののみpush
      if (wishListData.isYetAPI()) {
        wishListData.setRowNumber(rowNumber);
        wishListDataArray.push(wishListData);
      }
    }

    return wishListDataArray;
  }

  // カラム名取得
  function getColumnName(sheet, columnName) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var columnIndex = headers.indexOf(columnName) + 1;
    return columnIndex;
  }
}

class WishListData {
  // 後からメンバーに"title"と"rowNumber"を追加
  constructor(appid, priority, added, status) {
    this.appid = appid;
    this.priority = priority;
    this.added = added;
    this.status = status;
  }

  isYetAPI() {
    return this.status === '未実行';
  }

  setRowNumber(rowNumber) {
    this.rowNumber = rowNumber;
  }

  getAppTitle() {
    // APIを叩きゲームタイトルを取得する
    try {
      const url = `http://store.steampowered.com/api/appdetails?appids=${this.appid}`;
      const res = UrlFetchApp.fetch(url);

      // JSON 形式に変換する
      const resJson = JSON.parse(res.getContentText());
      // 自身のappidをキーとしたオブジェクトで返ってくるため、指定して取り出す
      const appDetail = resJson[this.appid];

      if (appDetail.success && appDetail.data.name) {
        this.title = appDetail.data.name;
      } else {
        this.title = "";
      }

      this.status = '実行済';
    } catch (e) {
      Logger.log(`API Failed. appid = ${this.appid}`);
    }
  }

  refreshSheet(sheet, statusColumn, titleColumn) {
    //　SteamWishListシートに処理内容を反映させる
    sheet.getRange(this.rowNumber, statusColumn).setValue(this.status);
    sheet.getRange(this.rowNumber, titleColumn).setNumberFormat('@');
    sheet.getRange(this.rowNumber, titleColumn).setValue(this.title);
  }
}
