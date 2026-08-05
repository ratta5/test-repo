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

  // 機種ごとに最新の有効残高と入力日を追跡するオブジェクト
  const deviceStates = {};
  balanceCols.forEach(obj => {
    deviceStates[obj.deviceId] = {
      lastBalance: null,
      lastDate: null
    };
  });

  // 1行目のデータ(row = 1)が存在する場合、最初の基準残高として初期セット
  if (sourceData.length > 1) {
    const firstRow = sourceData[1];
    const firstRawDate = firstRow[dateColIndex];
    if (firstRawDate) {
      for (let i = 0; i < balanceCols.length; i++) {
        const colIndex = balanceCols[i].col;
        const deviceId = balanceCols[i].deviceId;
        const val = firstRow[colIndex];
        if (typeof val === "number" || (val !== "" && val !== null && !isNaN(Number(val)))) {
          deviceStates[deviceId].lastBalance = Number(val);
          deviceStates[deviceId].lastDate = firstRawDate;
        }
      }
    }
  }

  const monthlySumMap = {};
  for (let row = 2; row < sourceData.length; row++) {
    const currentRow = sourceData[row];
    const rawDate = currentRow[dateColIndex];
    if (!rawDate) continue;

    const dateStr = Utilities.formatDate(new Date(rawDate), "Asia/Tokyo", "yyyy/MM/dd");
    const monthKey = dateStr.slice(0, 7);
    const rowOut = ["", dateStr];
    let dailyTotal = 0;

    for (let i = 0; i < balanceCols.length; i++) {
      const colIndex = balanceCols[i].col;
      const deviceId = balanceCols[i].deviceId;
      const currentRaw = currentRow[colIndex];
      const state = deviceStates[deviceId];

      let totalGain = 0;
      const hasValue = (typeof currentRaw === "number") || (currentRaw !== "" && currentRaw !== null && !isNaN(Number(currentRaw)));

      if (hasValue) {
        const current = Number(currentRaw);
        if (state.lastBalance !== null && state.lastDate !== null) {
          const diff = current - state.lastBalance;
          // 前回入力日より後 〜 今回入力日までの期間出金合計
          const withdrawSum = getWithdrawSumInPeriod(withdrawData, deviceId, state.lastDate, rawDate);
          totalGain = diff + withdrawSum;
        } else {
          // 初回の有効残高記録
          const withdrawSum = withdrawMap[dateStr]?.[deviceId] || 0;
          totalGain = withdrawSum;
        }

        // 有効な残高・日付で最新状態を更新
        state.lastBalance = current;
        state.lastDate = rawDate;
      } else {
        // 残高が未入力（空欄）の日は今回の獲得ポイントを 0 とする（次回入力時にまとめて精算）
        totalGain = 0;
      }

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


/**
 * 日付オブジェクトを 00:00:00 に正規化して取得するヘルパー関数
 */
function parseDate(val) {
  if (!val) return null;
  let d = (val instanceof Date) ? val : new Date(val);
  if (isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/**
 * 前回入力日(startDateStr/Date)より後 〜 今回入力日(endDateStr/Date)までの期間で発生した出金額を合計
 */
function getWithdrawSumInPeriod(withdrawData, deviceId, startDateVal, endDateVal) {
  const startD = parseDate(startDateVal);
  const endD = parseDate(endDateVal);
  if (!startD || !endD) return 0;

  const startTime = startD.getTime();
  const endTime = endD.getTime();
  const targetDevice = String(deviceId).trim();

  let sum = 0;
  for (let i = 1; i < withdrawData.length; i++) {
    const row = withdrawData[i];
    const dateCell = row[1];
    const device = row[2];
    const amount = parseFloat(row[3]);

    if (!dateCell || !device || isNaN(amount)) continue;
    if (String(device).trim() !== targetDevice) continue;

    const d = parseDate(dateCell);
    if (!d) continue;

    const t = d.getTime();
    // 前回入力日より後 〜 今回入力日以前
    if (t > startTime && t <= endTime) {
      sum += amount;
    }
  }
  return sum;
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

