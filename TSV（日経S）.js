function exportTsvFromSheetS() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // 9行目からデータを取得（表示されているデータを取得する）
  var dataRange = sheet.getRange(9, 1, sheet.getLastRow() - 8, sheet.getLastColumn()); // (行, 列, 行数, 列数)
  var data = dataRange.getDisplayValues();  // セルの表示されている値を取得

  // A列（1列目）のデータをゼロパディングせずそのまま取得
  // dataはすでにgetDisplayValues()で取得しているため、A列はそのまま「006」として表示されます

  // データをTSV形式に変換
  // データをTSV形式に変換
  var tsvContent = data.map(function(row) {
    // 各セルのデータを文字列として扱う
    return row.map(function(cell) {
      return cell;
    }).join("\t");  // 各行のデータをタブ区切りに変換
  }).join("\n");  // 各行を改行で区切る

  // 日付を取得して「yyyymmdd」形式にフォーマット
  var today = new Date();
  var formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyyMMdd');

  // ベースファイル名を作成
  var baseFileName = "s_nikkei_" + formattedDate;
  var fileName = baseFileName + ".tsv";  // 「e_nikkei_yyyymmdd.tsv」

  var blob = Utilities.newBlob(tsvContent, 'text/tab-separated-values', fileName);  // TSVファイル用のBlobを作成

  // ユーザーごとのフォルダIDを設定
  var userEmail = Session.getActiveUser().getEmail();  // 現在ログインしているユーザーのメールアドレスを取得
  var folderId;

  // メールアドレスごとにフォルダIDを設定
  if (userEmail === 'keisuke.mitsui@quick.jp') {
    folderId = '11MpAF3YQ4n7Fq3LbQ6v_G1t10_ibxGVW';
    } else if (userEmail === 'takae.arai21.s@quick.jp') {
    folderId = '1gQdrvk5GQADAnZR19JqtNYLjETGEB2xD'; 
    } else if (userEmail === 'hanako.sarashina.s@quick.jp') {
    folderId = '17MX3TZsxfxdAgqAP74ZIXAaSajdKGX_5';
    } else if (userEmail === 'mayumi.arai.s@quick.jp') {
    folderId = '1r0oyj4NPAIuLuNNSThpVPIzwHZCOJi1e'; 
    } else if (userEmail === 'eri.hamanaga.s@quick.jp') {
    folderId = '11IMQ8JHNWnb4SjGSbUVJVW8uGMXkOM-y'; 
    } else if (userEmail === 'yuriko.ishizuka.s@quick.jp') {
    folderId = '1drg4b74fDXu62oWAqlgCkR3KIzXXjMi8'; 
    } else if (userEmail === 'tomoya.tanabe@quick.jp') {
    folderId = '1xvtq2ZafASjhEAPvHWqwpnEds63PBb56'; 
    } else if (userEmail === 'momoko.awai@quick.jp') {
    folderId = '1NzjMcp1DB_0g8ysbSQa_PiOkjraGMrio'; 
    } else if (userEmail === 'hiroki.inoue.24@quick.jp') {
    folderId = '1x1s_claOgISmupMQeRWEnQ6fKWZ1D6em'; 
  }

  // フォルダを指定してファイルを保存
  if (folderId) {
    var folder = DriveApp.getFolderById(folderId);

    // 同名ファイルが存在するかを確認
    var files = folder.getFilesByName(fileName);
    var fileIndex = 1;

    // 同名ファイルが存在する場合、連番を付けたファイル名に変更
    while (files.hasNext()) {
      fileIndex++;
      fileName = baseFileName + "_" + fileIndex + ".tsv";
      files = folder.getFilesByName(fileName);
    }

    var blob = Utilities.newBlob(tsvContent, 'text/tab-separated-values', fileName);  // 更新されたファイル名でBlobを作成
    var newFile = folder.createFile(blob);  // ファイルを保存
    var fileUrl = newFile.getUrl();  // 作成されたファイルのURLを取得
  
    // C6にTSVファイルのURLを表示
    sheet.getRange("C6").setValue(fileUrl);  // C6セルにURLをセット
    
    // ファイル生成完了メッセージを表示
    SpreadsheetApp.getUi().alert("TSVデータが生成されました: " + fileName);
  }
}