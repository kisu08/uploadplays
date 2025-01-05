function searchDataE2023() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 別のスプレッドシートのIDを指定
  var externalSpreadsheetId = '171cd4G1arGMUmHHkMeZdMj8S8Lo5mTbi4R63PgjspYA';

  // 別のスプレッドシートを開く
  var externalSpreadsheet = SpreadsheetApp.openById(externalSpreadsheetId);

  // シート名を確認
  var sheetAName = 'E（日経＋QUICK合計）';
  var sheetBName = 'Eデータ(日経)';

  var sheetA = externalSpreadsheet.getSheetByName(sheetAName);
  var sheetB = ss.getSheetByName(sheetBName);

  // シートが正しく取得できているか確認
  if (!sheetA) {
    Logger.log('E（日経＋QUICK合計）が見つかりません: ' + sheetAName);
    SpreadsheetApp.getUi().alert('E（日経＋QUICK合計）が見つかりません: ' + sheetAName);
    return;
  }
  if (!sheetB) {
    Logger.log('Eデータ(日経)が見つかりません: ' + sheetBName);
    SpreadsheetApp.getUi().alert('Eデータ(日経)が見つかりません: ' + sheetBName);
    return;
  }

  // 確認2023(日経E)のC4からH4までのセルの値を取得
  var codesToSearch = sheetB.getRange('C4:H4').getValues()[0];  // C4:H4の範囲を取得
  // 空欄を除外したコードのみを抽出
  codesToSearch = codesToSearch.filter(function(code) {
    return code !== '';  // 空欄を除外
  });

    // 現在のシート (Eデータ(日経)) に既に存在する銘柄コードを取得
  var existingCodes = sheetB.getRange(9, 3, sheetB.getLastRow() - 8, 1).getValues() // 9行目以降のC列
    .flat()
    .filter(function(code) {
      return code !== ''; // 空欄を除外
    });

  // 抽出対象のコードから既存コードを除外
  codesToSearch = codesToSearch.filter(function(code) {
    return !existingCodes.includes(code); // 既に存在するコードは除外
  });

  // E（日経＋QUICK合計）の範囲を取得
  var dataA = sheetA.getDataRange().getValues();

  // E（日経＋QUICK合計）のヘッダー行を取得
  var headersA = sheetA.getRange(3, 1, 1, sheetA.getLastColumn()).getValues()[0];

  // 銘柄コードの列インデックスを取得
  var codeColumnIndex = headersA.indexOf('コード');
  if (codeColumnIndex === -1) {
    SpreadsheetApp.getUi().alert('E（日経＋QUICK合計）に銘柄コード列が見つかりません。');
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

  // E（日経＋QUICK合計）の3行目の項目名を取得
  var headersA3 = sheetA.getRange(3, 1, 1, sheetA.getLastColumn()).getValues()[0];

  // E（日経＋QUICK合計）の2行目の項目名を取得
  var headersA2 = sheetA.getRange(2, 1, 1, sheetA.getLastColumn()).getValues()[0];

  // 変更箇所: startRowBを最後のデータ行の次に設定
  var lastRowB = sheetB.getLastRow();  // シートの最後の行を取得
  var startRowB = lastRowB + 1;  // 最後の行の次からデータを挿入

  // マッチした行のデータを反映
  for (var r = 0; r < matchingRows.length; r++) {
    var rowIndex = matchingRows[r];
    var dataToReflect = dataA[rowIndex];
    Logger.log('マッチしたデータを反映: 行番号 ' + rowIndex);

    for (var h = 0; h < headersB.length; h++) {
      if (h < 8) { // H列まで
        // 確認2023(日経E)の5行目のH列までとE（日経＋QUICK合計）3行目の項目名を比較
        var headerIndexA3 = headersA3.indexOf(headersB[h]);
        if (headerIndexA3 !== -1) {
          var cell = sheetB.getRange(startRowB + r, h + 1);
          cell.setNumberFormat('@'); //書式をテキストに設定
          cell.setValue(dataToReflect[headerIndexA3]);
          Logger.log('セル(' + (startRowB + r) + ', ' + (h + 1) + ') にデータを設定: ' + dataToReflect[headerIndexA3]);
        }
      } else {
        // 確認2023(日経E)の5行目のH列以降とE（日経＋QUICK合計）2行目の項目名を比較
        var headerIndexA2 = headersA2.indexOf(headersB[h]);
        if (headerIndexA2 !== -1) {
          var cell = sheetB.getRange(startRowB + r, h + 1);
          cell.setNumberFormat('@'); //書式をテキストに設定
          cell.setValue(dataToReflect[headerIndexA2]);
          Logger.log('セル(' + (startRowB + r) + ', ' + (h + 1) + ') にデータを設定: ' + dataToReflect[headerIndexA2]);
        }
      }
    }
  }

  SpreadsheetApp.getUi().alert('抽出処理が正常に完了しました');
}