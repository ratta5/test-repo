function ポイント増加量作成() {
  // ✅ フォームから送信された出金データを機種ごとに分解して同期
  出金データ同期();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("TikTokﾌｫｰﾑﾃﾞｰﾀ");
  const withdrawSheet = ss.getSheetByName("TikTok出金ﾃﾞｰﾀ");
  const outputSheet = ss.getSheetByName("TikTok集計ﾃﾞｰﾀ");

  if (!sourceSheet || !withdrawSheet || !outputSheet) {
    throw new Error("必要なシートが見つかりません。");
  }

  const sourceData = sourceSheet.getDataRange().getValues();
  const withdrawData = withdrawSheet.getDataRange().getValues();
  const headers = sourceData[0];
  const dateColIndex = 1;

  // ✅ 残高列の特定（空白や表記揺れにも対応）
  const balanceCols = [];
  for (let col = 2; col < headers.length; col++) {
    const header = headers[col] ? headers[col].toString().trim() : "";
    const match = header.match(/^残高\s*(.+)$/);  // 「残高」以降を抽出
    if (match) {
      const deviceId = match[1].trim();
      balanceCols.push({ col, deviceId });
    }
  }

  // 出金データをマップ化
  const withdrawMap = {};
  for (let i = 1; i < withdrawData.length; i++) {
    const row = withdrawData[i];
    const dateCell = row[1];
    const device = row[2];
    const amount = parseFloat(row[3]);
    if (!dateCell || !device || isNaN(amount)) continue;

    const date = Utilities.formatDate(new Date(dateCell), "Asia/Tokyo", "yyyy/MM/dd");
    if (!withdrawMap[date]) withdrawMap[date] = {};
    if (!withdrawMap[date][device]) withdrawMap[date][device] = 0;
    withdrawMap[date][device] += amount;
  }

  // 出力配列作成
  const output = [];
  const outputHeader = ["", "日付"];
  balanceCols.forEach(obj => outputHeader.push(obj.deviceId));
  outputHeader.push("増加合計");
  outputHeader.push("当月累計");
  output.push(outputHeader);

  const monthlySumMap = {};
  for (let row = 2; row < sourceData.length; row++) {
    const currentRow = sourceData[row];
    const prevRow = sourceData[row - 1];
    const rawDate = currentRow[dateColIndex];
    if (!rawDate) continue;

    const dateStr = Utilities.formatDate(new Date(rawDate), "Asia/Tokyo", "yyyy/MM/dd");
    const monthKey = dateStr.slice(0, 7);
    const rowOut = ["", dateStr];
    let dailyTotal = 0;

    for (let i = 0; i < balanceCols.length; i++) {
      const colIndex = balanceCols[i].col;
      const deviceId = balanceCols[i].deviceId;
      const current = currentRow[colIndex];
      const prev = prevRow[colIndex];
      const diff = (typeof current === "number" && typeof prev === "number") ? current - prev : 0;
      const withdraw = withdrawMap[dateStr]?.[deviceId] || 0;
      const totalGain = diff + withdraw;

      rowOut.push(totalGain);
      dailyTotal += totalGain;
    }

    rowOut.push(dailyTotal); // 増加合計
    monthlySumMap[monthKey] = (monthlySumMap[monthKey] || 0) + dailyTotal;
    rowOut.push(monthlySumMap[monthKey]); // 当月累計
    output.push(rowOut);
  }

  // 出力
  outputSheet.clearContents();
  outputSheet.getRange(1, 1, output.length, output[0].length).setValues(output);

  // グラフ削除
  const charts = outputSheet.getCharts();
  charts.forEach(chart => outputSheet.removeChart(chart));

  const lastRow = output.length;

  // ✅ グラフ①：積み上げ棒グラフ
  const chart1 = outputSheet.newChart()
    .asColumnChart()
    .addRange(outputSheet.getRange(1, 2, lastRow))  // B列：日付
    .addRange(outputSheet.getRange(1, 3, lastRow, balanceCols.length))  // 各機種列
    .setStacked()
    .setOption("title", "機種別 日別ポイント増加（積み上げ棒グラフ）")
    .setPosition(2, 27, 0, 0)
    .build();
  outputSheet.insertChart(chart1);
}


function 列を拡張する() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("TikTok集計ﾃﾞｰﾀ");

  if (!sheet) {
    throw new Error("シート『TikTok集計ﾃﾞｰﾃ』が見つかりません。");
  }

  const currentCols = sheet.getMaxColumns();
  const additionalCols = 10;

  if (currentCols < 36) {
    sheet.insertColumnsAfter(currentCols, additionalCols);
    Logger.log(`${additionalCols} 列追加しました。合計列数: ${sheet.getMaxColumns()}`);
  } else {
    Logger.log("既に十分な列があります。追加は不要です。");
  }
}

