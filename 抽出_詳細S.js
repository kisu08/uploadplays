function searchDataSdetail2023() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 別のスプレッドシートのIDを指定
  var externalSpreadsheetId = '10kAc4hzeyxmbkIWOBXp3vdfrSVSl4XjRrYZdlNVPcEY';
  
  // 別のスプレッドシートを開く
  var externalSpreadsheet = SpreadsheetApp.openById(externalSpreadsheetId);
  
  // シート名を確認
  var sheetAName = 'S詳細';
  var sheetBName = 'Sデータ(詳細)';

  var sheetA = externalSpreadsheet.getSheetByName(sheetAName);
  var sheetB = ss.getSheetByName(sheetBName);

  // シートが正しく取得できているか確認
  if (!sheetA) {
    Logger.log('S詳細が見つかりません: ' + sheetAName);
    SpreadsheetApp.getUi().alert('S詳細が見つかりません: ' + sheetAName);
    return;
  }
  if (!sheetB) {
    Logger.log('Sデータ(詳細)が見つかりません: ' + sheetBName);
    SpreadsheetApp.getUi().alert('Sデータ(詳細)が見つかりません: ' + sheetBName);
    return;
  }

  // 確認2023(日経E)のC4からH4までのセルの値を取得
  var codesToSearch = sheetB.getRange('C4:H4').getValues()[0];  // C4:H4の範囲を取得
  // 空欄を除外したコードのみを抽出
  codesToSearch = codesToSearch.filter(function(code) {
    return code !== '';  // 空欄を除外
  });

    // 現在のシート (Sデータ(詳細)) に既に存在する銘柄コードを取得
  var existingCodes = sheetB.getRange(9, 3, sheetB.getLastRow() - 8, 1).getValues() // 9行目以降のC列
    .flat()
    .filter(function(code) {
      return code !== ''; // 空欄を除外
    });

  // 抽出対象のコードから既存コードを除外
  codesToSearch = codesToSearch.filter(function(code) {
    return !existingCodes.includes(code); // 既に存在するコードは除外
  });

  // S詳細の範囲を取得
  var dataA = sheetA.getDataRange().getValues();

  // S詳細のヘッダー行を取得
  var headersA = sheetA.getRange(2, 1, 1, sheetA.getLastColumn()).getValues()[0];

  // 銘柄コードの列インデックスを取得
  var codeColumnIndex = headersA.indexOf('コード');
  if (codeColumnIndex === -1) {
    SpreadsheetApp.getUi().alert('S詳細に銘柄コード列が見つかりません。');
    return;
  }

  // マッチした行を保持する配列
  var matchingRows = [];

  // すべてのコードに対してマッチする行を検索
  for (var i = 1; i < dataA.length; i++) {
    if (codesToSearch.includes(dataA[i][codeColumnIndex])) {
      matchingRows.push(i);  // マッチした行番号を保存
    }
  }

  if (matchingRows.length == 0) {
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

  // Sデータ(詳細)の6行目の項目名を取得
  var headersB = sheetB.getRange(8, 1, 1, sheetB.getLastColumn()).getValues()[0];

  // 変更箇所: startRowBを最後のデータ行の次に設定
  var lastRowB = sheetB.getLastRow();  // シートの最後の行を取得
  var startRowB = lastRowB + 1;  // 最後の行の次からデータを挿入

  for (var r = 0; r < matchingRows.length; r++) {
    var rowIndex = matchingRows[r];
    var dataToReflect = dataA[rowIndex];

    for (var j = 0; j < headersB.length; j++) {
      // Sデータ(詳細)の6行目のB列までとE詳細2行目の項目名を比較
      var headerIndexA = headersA.indexOf(headersB[j]);
      if (headerIndexA !== -1) {
        var cell = sheetB.getRange(startRowB + r, j+1);
        cell.setValue(dataToReflect[headerIndexA]);
      }
    }
  }
  SpreadsheetApp.getUi().alert('抽出処理が正常に完了しました');
}