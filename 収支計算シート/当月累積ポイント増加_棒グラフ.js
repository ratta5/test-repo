function 当月累積ポイント増加_棒グラフ() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("TikTok集計ﾃﾞｰﾀ");

  if (!sheet) {
    throw new Error("シート『TikTok集計ﾃﾞｰﾀ』が見つかりません。");
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // ✅ 「当月累計」列の列番号を自動で探す
  const cumulativeColIndex = headers.findIndex(h => h === "当月累計");
  if (cumulativeColIndex === -1) {
    throw new Error("「当月累計」列が見つかりません。");
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    throw new Error("十分なデータがありません。");
  }

  const dateRange = sheet.getRange(2, 2, lastRow - 1);  // B列：日付（2列目固定）
  const cumulativeRange = sheet.getRange(2, cumulativeColIndex + 1, lastRow - 1);  // 0-index → 1-index

  // 最大値を取得してY軸に使う
  const cumulativeValues = cumulativeRange.getValues().flat().filter(v => typeof v === "number");
  const maxY = Math.max(...cumulativeValues);
  const yAxisMax = Math.ceil(maxY / 1000) * 1000;

  // 古いグラフを削除
  const charts = sheet.getCharts();
  charts.forEach(chart => {
    const title = chart.getOptions()?.title || "";
    if (title.includes("当月累積ポイント増加")) {
      sheet.removeChart(chart);
    }
  });

  // グラフ作成
  const chart = sheet.newChart()
    .asColumnChart()
    .addRange(dateRange)
    .addRange(cumulativeRange)
    .setOption("title", "当月累積ポイント増加（棒グラフ）")
    .setOption("vAxis", {
      title: "累積ポイント",
      viewWindow: {
        max: yAxisMax,
        min: 0
      }
    })
    .setPosition(23, 27, 0, 0)
    .build();

  sheet.insertChart(chart);
}
