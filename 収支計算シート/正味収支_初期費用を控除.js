function 正味収支_初期費用を控除() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const costSheet = ss.getSheetByName("在庫一覧_TikTok関連");
  const netSheet = ss.getSheetByName("正味収支");
  if (!costSheet || !netSheet) throw new Error("必要なシートが見つかりません。");

  const costData = costSheet.getDataRange().getValues();
  const monthlyCostMap = {};

  for (let i = 1; i < costData.length; i++) {
    const monthStr = costData[i][11]; // L列：年月（文字列）
    const cost = costData[i][12];     // M列：初期費用
    if (typeof monthStr !== "string" || typeof cost !== "number") continue;
    monthlyCostMap[monthStr] = (monthlyCostMap[monthStr] || 0) + cost;
  }

  const netData = netSheet.getDataRange().getValues();
  const headers = netData[0];
  const output = [headers.concat("初期費用", "差引後合計")];

  for (let i = 1; i < netData.length; i++) {
    const row = netData[i];
    const monthCell = row[0];
    let monthStr = "";

    if (monthCell instanceof Date) {
      monthStr = Utilities.formatDate(monthCell, "Asia/Tokyo", "yyyy/MM");
    } else if (typeof monthCell === "string") {
      monthStr = monthCell;
    } else {
      output.push(row.concat("", row[headers.length - 1]));
      continue;
    }

    const total = row[headers.length - 1];
    const cost = monthlyCostMap[monthStr] || 0;
    const finalNet = (typeof total === "number") ? total - cost : "";
    output.push(row.concat(cost, finalNet));
  }

  netSheet.clear();
  netSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}

