function 正味収支_積み上げグラフを作成() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("正味収支");
  if (!sheet) throw new Error("シート『正味収支』が見つかりません。");

  const data = sheet.getDataRange().getValues();
  const numRows = data.length;
  const headers = data[0];

  if (numRows < 2 || headers.length < 4) {
    throw new Error("十分なデータがありません。");
  }

  const monthRange = sheet.getRange(2, 1, numRows - 1); // A列：月
  const deviceCount = headers.length - 4; // B列〜(X列=合計)まで（最後に初期費用・差引後合計がある）
  const deviceRange = sheet.getRange(2, 2, numRows - 1, deviceCount); // 機種別収支

  const charts = sheet.getCharts();
  charts.forEach(chart => sheet.removeChart(chart));

  const chart = sheet.newChart()
    .asColumnChart()
    .addRange(monthRange)
    .addRange(deviceRange)
    .setStacked()
    .setOption("title", "月別 機種別 正味収支（積み上げ棒グラフ）")
    .setOption("vAxis", { title: "正味収支ポイント" })
    .setPosition(5, 28, 0, 0)
    .build();

  sheet.insertChart(chart);
}
