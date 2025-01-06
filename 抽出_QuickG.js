function searchDataGQ2023() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 別のスプレッドシートのIDを指定
    var externalSpreadsheetId1 = '1_b5l6hPz50367J-LvWxu5nMgM2i2d-bkj59NDRSqNI0';
    var externalSpreadsheetId2 = '1usUvHJWjFdiFdRw0jBz2AqQXHU-FWUxumGrsqSBawKQ';
  
  // 別のスプレッドシートを開く
  var externalSpreadsheet1 = SpreadsheetApp.openById(externalSpreadsheetId1);
  var externalSpreadsheet2 = SpreadsheetApp.openById(externalSpreadsheetId2);
  
  // シート名を確認
  var sheetAName1 = 'G（日経＋QUICK合計）';
  var sheetAName2 = '入力シート（合計QUICKG）';
  var sheetBName = 'Gデータ(QUICK)';
  
  var sheetA1 = externalSpreadsheet1.getSheetByName(sheetAName1);
  var sheetA2 = externalSpreadsheet2.getSheetByName(sheetAName2);
  var sheetB = ss.getSheetByName(sheetBName);

  // シートが正しく取得できているか確認
  if (!sheetA1) {
    Logger.log('G（日経＋QUICK合計）が見つかりません: ' + sheetAName1);
    SpreadsheetApp.getUi().alert('G（日経＋QUICK合計）が見つかりません: ' + sheetAName1);
    return;
  }
  if (!sheetA2) {
    Logger.log('入力シート（合計QUICKG）が見つかりません: ' + sheetAName2);
    SpreadsheetApp.getUi().alert('入力シート（合計QUICKG）が見つかりません: ' + sheetAName2);
    return;
  }
  if (!sheetB) {
    Logger.log('Gデータ(QUICK)が見つかりません: ' + sheetBName);
    SpreadsheetApp.getUi().alert('Gデータ(QUICK)が見つかりません: ' + sheetBName);
    return;
  }

  // Gデータ(QUICK)のC4からH4までのセルの値を取得
  var codesToSearch = sheetB.getRange('C4:H4').getValues()[0];  // C4:H4の範囲を取得
  // 空欄を除外したコードのみを抽出
  codesToSearch = codesToSearch.filter(function(code) {
    return code !== '';  // 空欄を除外
  });

    // 現在のシート (Gデータ(QUICK)) に既に存在する銘柄コードを取得
  var existingCodes = sheetB.getRange(9, 3, sheetB.getLastRow() - 8, 1).getValues() // 9行目以降のC列
    .flat()
    .filter(function(code) {
      return code !== ''; // 空欄を除外
    });

  // 抽出対象のコードから既存コードを除外
  codesToSearch = codesToSearch.filter(function(code) {
    return !existingCodes.includes(code); // 既に存在するコードは除外
  });

  // G（日経＋QUICK合計）の範囲を取得
  var dataA1 = sheetA1.getDataRange().getValues();

  // G（日経＋QUICK合計）のヘッダー行を取得
  var headersA1 = sheetA1.getRange(3, 1, 1, sheetA1.getLastColumn()).getValues()[0];

  // 入力シート（合計QUICKG）の範囲を取得
  var dataA2 = sheetA2.getDataRange().getValues();

  // 入力シート（合計QUICKG）のヘッダー行を取得
  var headersA2 = sheetA2.getRange(3, 1, 1, sheetA2.getLastColumn()).getValues()[0];

  // 銘柄コードの列インデックスを取得
  var codeColumnIndex1 = headersA1.indexOf('コード');
  var codeColumnIndex2 = headersA2.indexOf('コード');
  if (codeColumnIndex1 === -1 || codeColumnIndex2 === -1) {
    SpreadsheetApp.getUi().alert('コード列が見つかりません。');
    return;
  }

  // G（日経＋QUICK合計）から抽出
  var matchingRows1 = [];
  for (var i = 1; i < dataA1.length; i++) {
    if (codesToSearch.includes(dataA1[i][codeColumnIndex1])) {
      matchingRows1.push(i);
    }
  }
  // 入力シート（合計QUICKG）から抽出
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

  // G（日経＋QUICK合計）の3行目の項目名を取得
  var headersA4 = sheetA1.getRange(3, 1, 1, sheetA1.getLastColumn()).getValues()[0];
  
  // G（日経＋QUICK合計）の2行目の項目名を取得
  var headersA3 = sheetA1.getRange(2, 1, 1, sheetA1.getLastColumn()).getValues()[0];

  //入力シート（合計QUICKG）の3行目の項目名を取得
  var headersA6 = sheetA2.getRange(3, 1, 1, sheetA2.getLastColumn()).getValues()[0];
  
  //入力シート（合計QUICKG）の2行目の項目名を取得
  var headersA5 = sheetA2.getRange(2, 1, 1, sheetA2.getLastColumn()).getValues()[0];

  // 現在のシート (Gデータ(QUICK)) の最後の行を取得
  var lastRowB = sheetB.getLastRow(); // シートのデータがある最後の行番号
  var startRowB1 = lastRowB + 1;      // G（日経＋QUICK合計）のデータ開始行
  
  // マッチした行のデータを反映
  for (var r = 0; r < matchingRows1.length; r++) {
    var rowIndex = matchingRows1[r]; // マッチした行番号を取得
    var dataToReflect1 = dataA1[rowIndex]; // 該当行のデータを取得
    Logger.log('マッチしたデータを反映: 行番号 ' + rowIndex);

    for (var h = 0; h < headersB.length; h++) {
      if (h < 8) { // H列まで
        // Gデータ(QUICK)の5行目のH列までとG（日経＋QUICK合計）3行目の項目名を比較
        var headerIndexA4 = headersA4.indexOf(headersB[h]);
        if (headerIndexA4 !== -1) {
          var cell = sheetB.getRange(startRowB1 + r, h+1);
          cell.setNumberFormat('@'); //書式をテキストに設定
          cell.setValue(dataToReflect1[headerIndexA4]);
        }
      } else {
        // Gデータ(QUICK)の5行目のH列以降とG（日経＋QUICK合計）2行目の項目名を比較
        var headerIndexA3 = headersA3.indexOf(headersB[h]);
        if (headerIndexA3 !== -1) {
          var cell = sheetB.getRange(startRowB1 + r,h+1);
          cell.setNumberFormat('@'); //書式をテキストに設定
          cell.setValue(dataToReflect1[headerIndexA3]);
        }
      }
    }
  }

  // 次の開始行を計算（1つ目のシートのデータ行数を加算）
  var startRowB2 = startRowB1 + matchingRows1.length;

  // マッチした行のデータを反映
  for (var r = 0; r < matchingRows2.length; r++) {
    var rowIndex = matchingRows2[r];
    var dataToReflect2 = dataA2[rowIndex];
    for (var h = 0; h < headersB.length; h++) {
      if (h < 8) { // H列まで
        // 確認2023(QuickE)の5行目のJ列までと入力シート（合計QUICKG）3行目の項目名を比較
        var headerIndexA6 = headersA6.indexOf(headersB[h]);
        if (headerIndexA6 !== -1) {
          var cell = sheetB.getRange(startRowB2 + r, h+1);
          cell.setNumberFormat('@'); //書式をテキストに設定
          cell.setValue(dataToReflect2[headerIndexA6]);
        }
      } else {
        // 確認2023(QuickE)の5行目のJ列以降と入力シート（合計QUICKG）2行目の項目名を比較
        var headerIndexA5 = headersA5.indexOf(headersB[h]);
        if (headerIndexA5 !== -1) {
          var cell = sheetB.getRange(startRowB2 + r,h+1);
          cell.setNumberFormat('@'); //書式をテキストに設定
          cell.setValue(dataToReflect2[headerIndexA5]);
        }
      }
    }
  }

  SpreadsheetApp.getUi().alert('抽出処理が正常に完了しました');
}