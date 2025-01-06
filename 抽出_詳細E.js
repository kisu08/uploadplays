function searchDataEdetail2023() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 別のスプレッドシートのIDを指定
    var externalSpreadsheetId1 = '171cd4G1arGMUmHHkMeZdMj8S8Lo5mTbi4R63PgjspYA';
    var externalSpreadsheetId2 = '1VeArrE9c8LF7_4v21wk0RAw9fF4KbuRfuVtLKx9FOuY';
  
  // 別のスプレッドシートを開く
  var externalSpreadsheet1 = SpreadsheetApp.openById(externalSpreadsheetId1);
  var externalSpreadsheet2 = SpreadsheetApp.openById(externalSpreadsheetId2);
  
  // シート名を確認
  var sheetAName1 = 'E詳細';
  var sheetAName2 = '入力シート（詳細E）';
  var sheetBName = 'Eデータ(詳細)';
  
  var sheetA1 = externalSpreadsheet1.getSheetByName(sheetAName1);
  var sheetA2 = externalSpreadsheet2.getSheetByName(sheetAName2);
  var sheetB = ss.getSheetByName(sheetBName);

  // シートが正しく取得できているか確認
  if (!sheetA1) {
    Logger.log('E詳細が見つかりません: ' + sheetAName1);
    SpreadsheetApp.getUi().alert('E詳細が見つかりません: ' + sheetAName1);
    return;
  }
  if (!sheetA2) {
    Logger.log('入力シート（詳細E）が見つかりません: ' + sheetAName2);
    SpreadsheetApp.getUi().alert('入力シート（詳細E）が見つかりません: ' + sheetAName2);
    return;
  }
  if (!sheetB) {
    Logger.log('Eデータ(詳細)が見つかりません: ' + sheetBName);
    SpreadsheetApp.getUi().alert('Eデータ(詳細)が見つかりません: ' + sheetBName);
    return;
  }

  // 確認2023(日経E)のC4からH4までのセルの値を取得
  var codesToSearch = sheetB.getRange('C4:H4').getValues()[0];  // C4:H4の範囲を取得
  // 空欄を除外したコードのみを抽出
  codesToSearch = codesToSearch.filter(function(code) {
    return code !== '';  // 空欄を除外
  });

    // 現在のシート (Eデータ(詳細)) に既に存在する銘柄コードを取得
  var existingCodes = sheetB.getRange(9, 3, sheetB.getLastRow() - 8, 1).getValues() // 9行目以降のC列
    .flat()
    .filter(function(code) {
      return code !== ''; // 空欄を除外
    });

  // 抽出対象のコードから既存コードを除外
  codesToSearch = codesToSearch.filter(function(code) {
    return !existingCodes.includes(code); // 既に存在するコードは除外
  });

  // E詳細の範囲を取得
  var dataA1 = sheetA1.getDataRange().getValues();

  // E詳細のヘッダー行を取得
  var headersA1 = sheetA1.getRange(2, 1, 1, sheetA1.getLastColumn()).getValues()[0];

  // 入力シート（詳細E）の範囲を取得
  var dataA2 = sheetA2.getDataRange().getValues();

  // 入力シート（詳細E）のヘッダー行を取得
  var headersA2 = sheetA2.getRange(2, 1, 1, sheetA2.getLastColumn()).getValues()[0];

  // 銘柄コードの列インデックスを取得
  var codeColumnIndex1 = headersA1.indexOf('コード');
  var codeColumnIndex2 = headersA2.indexOf('コード');
  if (codeColumnIndex1 === -1 || codeColumnIndex2 === -1) {
    SpreadsheetApp.getUi().alert('コード列が見つかりません。');
    return;
  }
  // E詳細から抽出
  var matchingRows1 = [];
  for (var i = 1; i < dataA1.length; i++) {
    if (codesToSearch.includes(dataA1[i][codeColumnIndex1])) {
      matchingRows1.push(i);
    }
  }
  // 入力シート（詳細E）から抽出
  var matchingRows2 = [];
  for (var i = 1; i < dataA2.length; i++) {
    if (codesToSearch.includes(dataA2[i][codeColumnIndex2])) {
      matchingRows2.push(i);
    }
  }

  if (matchingRows1.length == 0 && matchingRows2.length == 0) {
    // 該当する銘柄コードが見つからない場合
    SpreadsheetApp.getUi().alert('該当する銘柄コードが見つかりませんでした。');
    return;
  }

  // 8行目のA列から始まり、空のセルに到達するまでの範囲を取得
  var headersB = [];
  var colIndex = 1;
  while (sheetB.getRange(8, colIndex).getValue() !== "") {
      headersB.push(sheetB.getRange(8, colIndex).getValue());
      colIndex++;
  }

  // Eデータ(詳細)の6行目の項目名を取得
  var headersB = sheetB.getRange(8, 1, 1, sheetB.getLastColumn()).getValues()[0];

  // 変更箇所: startRowBを最後のデータ行の次に設定
  var lastRowB = sheetB.getLastRow();  // シートの最後の行を取得
  var startRowB = lastRowB + 1;  // 最後の行の次からデータを挿入

  // Eデータ(詳細)にデータを反映させる開始行
  for (var r = 0; r < matchingRows1.length; r++) {
    var rowIndex = matchingRows1[r];
    var dataToReflect1 = dataA1[rowIndex];

    for (var k = 0; k < headersB.length; k++) {
        // Eデータ(詳細)の5行目のJ列以降とE詳細2行目の項目名を比較
        var headerIndexA1 = headersA1.indexOf(headersB[k]);
        if (headerIndexA1 !== -1) {
          var cell = sheetB.getRange(startRowB + r,k+1);
          cell.setNumberFormat('@'); //書式をテキストに設定
          cell.setValue(dataToReflect1[headerIndexA1]);
        }
      }
  }
  
  // Eデータ(詳細)にデータを反映させる開始行
  for (var r = 0; r < matchingRows2.length; r++) {
    var rowIndex = matchingRows2[r];
    var dataToReflect2 = dataA2[rowIndex];
    for (var k = 0; k < headersB.length; k++) {
        // Eデータ(詳細)の5行目のJ列以降と入力シート（詳細E）2行目の項目名を比較
        var headerIndexA2 = headersA2.indexOf(headersB[k]);
        if (headerIndexA2 !== -1) {
          var cell = sheetB.getRange(startRowB + r,k+1);
          cell.setNumberFormat('@'); //書式をテキストに設定
          cell.setValue(dataToReflect2[headerIndexA2]);
        }
    }
  }
  SpreadsheetApp.getUi().alert('抽出処理が正常に完了しました');
}