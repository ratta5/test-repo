function 各機種当月累計_月別縦追記() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("TikTok集計ﾃﾞｰﾀ");
  const outputSheetName = "TikTok_機種別月累計";
  let outputSheet = ss.getSheetByName(outputSheetName);

  if (!sourceSheet) throw new Error("シート『TikTok集計ﾃﾞｰﾀ』が見つかりません。");

  if (!outputSheet) {
    outputSheet = ss.insertSheet(outputSheetName);
  }

  const data = sourceSheet.getDataRange().getValues();
  if (data.length < 3) throw new Error("集計データが不十分です。");

  const headers = data[0].slice(2, data[0].length - 2); // C列〜各機種名（増加合計と当月累計を除く）
  const deviceCount = headers.length;
  const monthMap = {};

  for (let i = 1; i < data.length; i++) {
    const dateCell = data[i][1];
    const date = new Date(dateCell);
    if (isNaN(date)) continue;

    const rowMonth = Utilities.formatDate(date, "Asia/Tokyo", "yyyy-MM");

    if (!monthMap[rowMonth]) {
      monthMap[rowMonth] = new Array(deviceCount).fill(0);
    }

    for (let j = 0; j < deviceCount; j++) {
      const val = data[i][j + 2];
      if (typeof val === 'number') {
        monthMap[rowMonth][j] += val;
      }
    }
  }

  // 出力整形
  const months = Object.keys(monthMap).sort();
  const output = [["月", ...headers, "合計"]]; // ← 「合計」列を見出しに追加

  months.forEach(month => {
    const values = monthMap[month];
    const sum = values.reduce((acc, val) => acc + (typeof val === 'number' ? val : 0), 0); // 合計値
    output.push([month, ...values, sum]);
  });

  outputSheet.clearContents();
  outputSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}
