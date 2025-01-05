function deleteDataE2023() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  //シート名を確認
  var sheetName = 'Eデータ(日経)';
  var sheet = ss.getSheetByName(sheetName);


  //シートが正しく取得できているか確認
  if(!sheet){
    Logger.log('Eデータ(日経)が見つかりません:' + sheetName);
    SpreadsheetApp.getUi().alert('Eデータ(日経)が見つかりません:' + sheetName);
    return;
  }

  //データの削除を実行
  var LastRow = sheet.getLastRow();
  var LastColumn = sheet.getLastColumn();
  sheet.getRange(10,1,LastRow-5,LastColumn).clear();
  sheet.getRange("C6").clearContent();  // C4セルのテキストを削除
  SpreadsheetApp.getUi().alert('削除処理が正常に完了しました');
}
