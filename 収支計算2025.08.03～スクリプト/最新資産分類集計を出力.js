function 最新資産分類集計を出力() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 資産シート = ss.getSheetByName("資産");
  const マスターシート = ss.getSheetByName("資産マスター");
  const 出力シート名 = "資産分類集計";
  let 出力シート = ss.getSheetByName(出力シート名);

  if (!出力シート) {
    出力シート = ss.insertSheet(出力シート名);
  } else {
    出力シート.clearContents();
  }

  const 資産データ = 資産シート.getDataRange().getValues();
  const 資産マスター = マスターシート.getDataRange().getValues();

  const ヘッダー = 資産データ[0];
  const データ本体 = 資産データ.slice(1);

  // 最新のタイムスタンプの行を取得
  const latestRow = データ本体.reduce((latest, row) => {
    return (!latest || new Date(row[0]) > new Date(latest[0])) ? row : latest;
  }, null);
  const 最新日付 = Utilities.formatDate(new Date(latestRow[0]), Session.getScriptTimeZone(), "yyyy-MM-dd");

  // 資産名ごとの金額マップを作成
  const 資産Map = {};
  for (let i = 1; i < ヘッダー.length; i++) {
    const 名称 = ヘッダー[i];
    const 金額 = latestRow[i] || 0;
    資産Map[名称] = 金額;
  }

  // 資産マスターをもとに分類ごとに集計
  const 集計結果 = {};
  for (let i = 1; i < 資産マスター.length; i++) {
    const 名称 = 資産マスター[i][1];
    const 分類 = 資産マスター[i][2];
    const 金額 = 資産Map[名称] || 0;
    if (!集計結果[分類]) {
      集計結果[分類] = 0;
    }
    集計結果[分類] += 金額;
  }

  // 出力
  const 出力 = [["日付", "分類", "金額"]];
  for (let 分類 in 集計結果) {
    出力.push([最新日付, 分類, 集計結果[分類]]);
  }

  出力シート.getRange(1, 1, 出力.length, 3).setValues(出力);
}
