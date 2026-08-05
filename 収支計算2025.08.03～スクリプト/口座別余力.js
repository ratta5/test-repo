function 口座別余力() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const assetSheet = ss.getSheetByName("資産");
  const repaymentSheet = ss.getSheetByName("口座別カード返済集計");
  const outputSheetName = "口座別余力";

  // 出力シート準備
  let outputSheet = ss.getSheetByName(outputSheetName);
  if (!outputSheet) {
    outputSheet = ss.insertSheet(outputSheetName);
  } else {
    outputSheet.clearContents();
  }

  const assetData = assetSheet.getDataRange().getValues();
  const repaymentData = repaymentSheet.getDataRange().getValues();

  const assetHeaders = assetData[0];
  const latestAssetRow = assetData.slice(1).reduce((a, b) => {
    return new Date(a[0]) > new Date(b[0]) ? a : b;
  });

  const latestTimestamp = new Date(latestAssetRow[0]);
  const latestMonth = Utilities.formatDate(latestTimestamp, Session.getScriptTimeZone(), "yyyy-MM");
  const latestMonthDate = new Date(latestMonth + "-01");

  // 次月を求める
  const nextMonthDate = new Date(latestMonthDate);
  nextMonthDate.setMonth(nextMonthDate.getMonth() + 1);
  const nextMonthStr = Utilities.formatDate(nextMonthDate, Session.getScriptTimeZone(), "yyyy-MM");

  // 当月・次月返済マップを作成
  const repaymentHeaders = repaymentData[0];
  const repaymentCurrent = {};
  const repaymentNext = {};

  repaymentData.slice(1).forEach(row => {
    const rowDate = new Date(row[0]);
    const ym = rowDate.getFullYear() + "-" + String(rowDate.getMonth() + 1).padStart(2, "0");
    for (let i = 1; i < row.length; i++) {
      const key = repaymentHeaders[i];
      if (ym === latestMonth) {
        repaymentCurrent[key] = row[i];
      } else if (ym === nextMonthStr) {
        repaymentNext[key] = row[i];
      }
    }
  });

  const banks = ["TKS銀行", "RKT銀行", "SZK信金"];
  const points = ["MRPY", "GNKN", "PYPY", "RTPY_YSK", "RTPY_SZK"];
  const receivables = ["AMZ売掛"];
  const allKeys = assetHeaders.slice(1);
  const result = [["項目", "資産額", "返済額", "差額（" + latestMonth + "）", "次月返済額（" + nextMonthStr + "）"]];

  // 銀行口座
  for (let key of banks) {
    const idx = assetHeaders.indexOf(key);
    const asset = latestAssetRow[idx] || 0;
    const repay = repaymentCurrent[key] || 0;
    const nextRepay = repaymentNext[key] || 0;
    result.push([key, asset, repay, asset - repay, nextRepay]);
  }

  // その他現金
  const excluded = new Set(["タイムスタンプ", ...banks, ...points, ...receivables]);
  let otherCash = 0;
  for (let key of allKeys) {
    if (!excluded.has(key)) {
      const idx = assetHeaders.indexOf(key);
      otherCash += latestAssetRow[idx] || 0;
    }
  }
  result.push(["その他現金", otherCash, 0, otherCash, 0]);

  // ポイントと売掛金
  for (let key of [...points, ...receivables]) {
    const idx = assetHeaders.indexOf(key);
    const asset = latestAssetRow[idx] || 0;
    result.push([key, asset, 0, asset, 0]);
  }

  outputSheet.getRange(1, 1, result.length, result[0].length).setValues(result);
}
