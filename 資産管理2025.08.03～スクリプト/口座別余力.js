function 口座別余力() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const assetSheet = ss.getSheetByName("資産");
  const repaymentSheet = ss.getSheetByName("口座別カード返済集計");
  const masterSheet = ss.getSheetByName("カードマスター");
  const paymentSheet = ss.getSheetByName("カード返済額_整形");
  const inventorySheet = ss.getSheetByName("棚卸_取り込み");
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
  const masterData = masterSheet ? masterSheet.getDataRange().getValues() : [];
  const paymentData = paymentSheet ? paymentSheet.getDataRange().getValues() : [];

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

  // 当月・次月返済マップを作成（口座別合計）
  const repaymentHeaders = repaymentData[0];
  const repaymentCurrent = {};
  const repaymentNext = {};

  repaymentData.slice(1).forEach(row => {
    const rowDate = new Date(row[0]);
    const ym = !isNaN(rowDate) 
      ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM") 
      : String(row[0]).substring(0, 7);
    for (let i = 1; i < row.length; i++) {
      const key = repaymentHeaders[i];
      if (ym === latestMonth) {
        repaymentCurrent[key] = row[i];
      } else if (ym === nextMonthStr) {
        repaymentNext[key] = row[i];
      }
    }
  });

  // カード名 -> 引落口座 のマッピングを作成
  const cardToBank = {};
  for (let i = 1; i < masterData.length; i++) {
    const card = masterData[i][1];
    const bank = masterData[i][2];
    if (card && bank) {
      cardToBank[card] = bank;
    }
  }

  // 金額パース補助関数
  function parseAmount(val) {
    if (typeof val === "number") return val;
    if (typeof val === "string") {
      const num = parseFloat(val.replace(/[¥,]/g, ""));
      return isNaN(num) ? 0 : num;
    }
    return 0;
  }

  // カード返済額_整形 から 当月・次月のカード別内訳データを取得
  const paymentHeaders = paymentData[0] || [];
  let currentPaymentRow = null;
  let nextPaymentRow = null;

  paymentData.slice(1).forEach(row => {
    const rowDate = new Date(row[0]);
    const ym = !isNaN(rowDate) 
      ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM") 
      : String(row[0]).substring(0, 7);
    if (ym === latestMonth) {
      currentPaymentRow = row;
    } else if (ym === nextMonthStr) {
      nextPaymentRow = row;
    }
  });

  // 口座ごとのカード内訳文字列を作成する関数
  function getCardBreakdown(bankName, row) {
    if (!row) return "";
    const details = [];
    paymentHeaders.forEach((header, colIdx) => {
      if (colIdx >= 2 && cardToBank[header] === bankName) {
        const amt = parseAmount(row[colIdx]);
        if (amt > 0) {
          details.push(header + ": ¥" + amt.toLocaleString());
        }
      }
    });
    return details.join("\n");
  }

  // 棚卸_取り込み シートから最新の棚卸額・棚卸日を取得
  let inventoryAmount = 0;
  let inventoryDateStr = "";
  if (inventorySheet) {
    const inventoryData = inventorySheet.getDataRange().getValues();
    if (inventoryData.length > 1) {
      const headers = inventoryData[0].map(h => String(h).trim());

      // ヘッダー位置の特定（"実在庫額総額" を優先指定）
      let amountColIdx = headers.findIndex(h => h === "実在庫額総額" || h.includes("実在庫額総額"));
      if (amountColIdx === -1) {
        amountColIdx = headers.findIndex(h => h.includes("実在庫額") || h.includes("在庫額"));
      }

      let dateColIdx = headers.findIndex(h => h === "棚卸日" || h.includes("棚卸日"));
      if (dateColIdx === -1) {
        dateColIdx = headers.findIndex(h => h.includes("日付"));
      }

      if (dateColIdx === -1) dateColIdx = 0;
      if (amountColIdx === -1) amountColIdx = headers.length > 1 ? 1 : 0;

      let latestRow = null;
      let maxTime = -Infinity;

      for (let i = 1; i < inventoryData.length; i++) {
        const row = inventoryData[i];
        if (!row || row.length === 0) continue;
        const dateVal = row[dateColIdx];
        const amtVal = row[amountColIdx];

        if ((dateVal !== "" && dateVal !== null && dateVal !== undefined) ||
            (amtVal !== "" && amtVal !== null && amtVal !== undefined)) {
          
          let t = -Infinity;
          if (dateVal instanceof Date) {
            t = dateVal.getTime();
          } else if (dateVal) {
            const d = new Date(dateVal);
            if (!isNaN(d.getTime())) t = d.getTime();
          }

          if (t > maxTime) {
            maxTime = t;
            latestRow = row;
          } else if (maxTime === -Infinity) {
            latestRow = row;
          }
        }
      }

      if (latestRow) {
        const rawDate = latestRow[dateColIdx];
        if (rawDate instanceof Date) {
          inventoryDateStr = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
        } else if (rawDate) {
          const d = new Date(rawDate);
          inventoryDateStr = !isNaN(d.getTime()) 
            ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd")
            : String(rawDate);
        }
        
        inventoryAmount = parseAmount(latestRow[amountColIdx]);
      }
    }
  }

  const banks = ["TKS銀行", "RKT銀行", "SZK信金"];
  const points = ["MRPY", "GNKN", "PYPY", "RTPY_YSK", "RTPY_SZK"];
  const receivables = ["AMZ売掛"];
  const allKeys = assetHeaders.slice(1);

  const result = [[
    "項目",
    "資産額",
    "カード内訳（" + latestMonth + "）",
    "返済額",
    "差額（" + latestMonth + "）",
    "カード内訳（" + nextMonthStr + "）",
    "次月返済額（" + nextMonthStr + "）",
    "差額（" + nextMonthStr + "）"
  ]];

  // 銀行口座
  for (let key of banks) {
    const idx = assetHeaders.indexOf(key);
    const asset = latestAssetRow[idx] || 0;
    const repay = repaymentCurrent[key] || 0;
    const nextRepay = repaymentNext[key] || 0;
    const diffCurrent = asset - repay;
    const diffNext = diffCurrent - nextRepay;

    const cardDetailsCurrent = getCardBreakdown(key, currentPaymentRow);
    const cardDetailsNext = getCardBreakdown(key, nextPaymentRow);

    result.push([key, asset, cardDetailsCurrent, repay, diffCurrent, cardDetailsNext, nextRepay, diffNext]);
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
  result.push(["その他現金", otherCash, "", 0, otherCash, "", 0, otherCash]);

  // ポイントと売掛金
  for (let key of [...points, ...receivables]) {
    const idx = assetHeaders.indexOf(key);
    const asset = latestAssetRow[idx] || 0;
    result.push([key, asset, "", 0, asset, "", 0, asset]);
  }

  // 棚卸資産
  const inventoryNote = inventoryDateStr ? "棚卸日: " + inventoryDateStr : "";
  result.push(["棚卸資産", inventoryAmount, inventoryNote, 0, inventoryAmount, "", 0, inventoryAmount]);

  // 総合計の算出
  let totalAsset = 0;
  let totalRepay = 0;
  let totalDiffCurrent = 0;
  let totalNextRepay = 0;
  let totalDiffNext = 0;

  for (let i = 1; i < result.length; i++) {
    totalAsset += parseAmount(result[i][1]);
    totalRepay += parseAmount(result[i][3]);
    totalDiffCurrent += parseAmount(result[i][4]);
    totalNextRepay += parseAmount(result[i][6]);
    totalDiffNext += parseAmount(result[i][7]);
  }

  result.push(["総合計", totalAsset, "", totalRepay, totalDiffCurrent, "", totalNextRepay, totalDiffNext]);

  // シートに出力
  outputSheet.getRange(1, 1, result.length, result[0].length).setValues(result);

  // 折返し表示
  outputSheet.getRange(1, 1, result.length, result[0].length).setWrap(true);

  // 通貨フォーマット設定
  if (result.length > 1) {
    outputSheet.getRange(2, 2, result.length - 1, 1).setNumberFormat('"¥"#,##0');
    outputSheet.getRange(2, 4, result.length - 1, 2).setNumberFormat('"¥"#,##0');
    outputSheet.getRange(2, 7, result.length - 1, 2).setNumberFormat('"¥"#,##0');
  }
}
