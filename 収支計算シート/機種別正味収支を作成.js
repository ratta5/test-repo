// 各機種当月累計_月別縦追記 の末尾に追加
function 機種別正味収支を作成() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pointSheet = ss.getSheetByName("TikTok_機種別月累計");
  const planSheet = ss.getSheetByName("スマホ通信費");
  let resultSheet = ss.getSheetByName("正味収支");

  if (!pointSheet || !planSheet) {
    throw new Error("必要なシートが見つかりません。");
  }

  if (!resultSheet) {
    resultSheet = ss.insertSheet("正味収支");
  } else {
    resultSheet.clear();
  }

  const pointData = pointSheet.getDataRange().getValues();
  const headers = pointData[0].slice(1); // 機種名
  const months = pointData.slice(1).map(row => row[0]); // 月
  const values = pointData.slice(1).map(row => row.slice(1)); // ポイント部分

  const planData = planSheet.getRange(3, 2, planSheet.getLastRow() - 2, 2).getValues(); // B3:C
  const planMap = {};
  for (const [device, cost] of planData) {
    if (device && typeof cost === 'number') {
      planMap[device] = cost;
    }
  }

  const startMonthMap = {};
  headers.forEach((device, colIndex) => {
    for (let row = 0; row < months.length; row++) {
      const val = values[row][colIndex];
      if (typeof val === "number" && val > 0) {
        startMonthMap[device] = months[row];
        break;
      }
    }
  });

  const result = [];
  result.push(["月", ...headers, "合計"]);

  for (let row = 0; row < months.length; row++) {
    const currentMonth = months[row];
    const resultRow = [currentMonth];
    let monthlySum = 0;

    for (let col = 0; col < headers.length; col++) {
      const device = headers[col];
      const point = values[row][col] || 0;
      const shouldSubtract = startMonthMap[device] && currentMonth >= startMonthMap[device];
      const plan = shouldSubtract ? (planMap[device] || 0) : 0;
      const net = point - plan;
      resultRow.push(net);
      monthlySum += net;
    }

    resultRow.push(monthlySum);
    result.push(resultRow);
  }

  resultSheet.getRange(1, 1, result.length, result[0].length).setValues(result);
  正味収支_積み上げグラフを作成();
  正味収支_初期費用を控除();
}

