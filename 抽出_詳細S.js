function searchDataSdetail2023() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 別のスプレッドシートのIDを指定
    var externalSpreadsheetId1 = '10kAc4hzeyxmbkIWOBXp3vdfrSVSl4XjRrYZdlNVPcEY';
    var externalSpreadsheetId2 = '1FXrNWdrhGiq6hB9KZO8pF8XHA5tgAbAoi4ni1EHcYDc';
  
  // 別のスプレッドシートを開く
  var externalSpreadsheet1 = SpreadsheetApp.openById(externalSpreadsheetId1);
  var externalSpreadsheet2 = SpreadsheetApp.openById(externalSpreadsheetId2);
  
  // シート名を確認
  var sheetAName1 = 'S詳細';
  var sheetAName2 = '入力シート（詳細S）';
  var sheetBName = 'Sデータ(詳細)';

  var sheetA1 = externalSpreadsheet1.getSheetByName(sheetAName1);
  var sheetA2 = externalSpreadsheet2.getSheetByName(sheetAName2);
  var sheetB = ss.getSheetByName(sheetBName);

  // シートが正しく取得できているか確認
  if (!sheetA1) {
    Logger.log('S詳細が見つかりません: ' + sheetAName1);
    SpreadsheetApp.getUi().alert('S詳細が見つかりません: ' + sheetAName1);
    return;
  }
  if (!sheetA2) {
    Logger.log('入力シート（詳細S）が見つかりません: ' + sheetAName2);
    SpreadsheetApp.getUi().alert('入力シート（詳細S）が見つかりません: ' + sheetAName2);
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
  var dataA1 = sheetA1.getDataRange().getValues();

  // S詳細のヘッダー行を取得
  var headersA1 = sheetA1.getRange(2, 1, 1, sheetA1.getLastColumn()).getValues()[0];

  // 入力シート（詳細S）の範囲を取得
  var dataA2 = sheetA2.getDataRange().getValues();

  // 入力シート（詳細S）のヘッダー行を取得
  var headersA2 = sheetA2.getRange(2, 1, 1, sheetA2.getLastColumn()).getValues()[0];

  // 銘柄コードの列インデックスを取得
  var codeColumnIndex1 = headersA1.indexOf('コード');
  var codeColumnIndex2 = headersA2.indexOf('コード');
  if (codeColumnIndex1 === -1 || codeColumnIndex2 === -1) {
    SpreadsheetApp.getUi().alert('コード列が見つかりません。');
    return;
  }

  // S詳細から抽出
  var matchingRows1 = [];
  for (var i = 1; i < dataA1.length; i++) {
    if (codesToSearch.includes(dataA1[i][codeColumnIndex1])) {
      matchingRows1.push(i);
    }
  }
  // 入力シート（詳細S）から抽出
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

  // 確認2023(詳細E)の6行目の項目名を取得
  var headersB = sheetB.getRange(8, 1, 1, sheetB.getLastColumn()).getValues()[0];

  // 現在のシート (Eデータ(QUICK)) の最後の行を取得
  var lastRowB = sheetB.getLastRow(); // シートのデータがある最後の行番号
  var startRowB1 = lastRowB + 1;      // E（日経＋QUICK合計）のデータ開始行

  // マッチした行のデータを反映
  for (var r = 0; r < matchingRows1.length; r++) {
    var rowIndex = matchingRows1[r]; // マッチした行番号を取得
    var dataToReflect1 = dataA1[rowIndex]; // 該当行のデータを取得
    Logger.log('マッチしたデータを反映: 行番号 ' + rowIndex);

    for (var j = 0; j < headersB.length; j++) {
        var headerIndexA1 = headersA1.indexOf(headersB[j]);
        if (headerIndexA1 !== -1) {
          var cell = sheetB.getRange(startRowB1 + r, j+1);
          cell.setNumberFormat('@'); //書式をテキストに設定
          cell.setValue(dataToReflect1[headerIndexA1]);
      }
    }
  }
  // 次の開始行を計算（1つ目のシートのデータ行数を加算）
  var startRowB2 = startRowB1 + matchingRows1.length;

  // マッチした行のデータを反映
  for (var r = 0; r < matchingRows2.length; r++) {
    var rowIndex = matchingRows2[r]; // マッチした行番号を取得
    var dataToReflect2 = dataA2[rowIndex]; // 該当行のデータを取得
    Logger.log('マッチしたデータを反映: 行番号 ' + rowIndex);

    for (var j = 0; j < headersB.length; j++) {
        var headerIndexA2 = headersA2.indexOf(headersB[j]);
        if (headerIndexA2 !== -1) {
          var cell = sheetB.getRange(startRowB2 + r, j+1);
          cell.setNumberFormat('@'); //書式をテキストに設定
          cell.setValue(dataToReflect2[headerIndexA2]);
      }
    }
  }
  SpreadsheetApp.getUi().alert('抽出処理が正常に完了しました');
}