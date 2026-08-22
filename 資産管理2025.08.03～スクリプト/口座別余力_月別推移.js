function 口座別余力_月別推移() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const assetSheet = ss.getSheetByName("資産");
  const repaymentSheet = ss.getSheetByName("口座別カード返済集計");
  const masterSheet = ss.getSheetByName("カードマスター");
  const paymentSheet = ss.getSheetByName("カード返済額_整形");
  const estimatedStockSheet = ss.getSheetByName("推測在庫高_推移");
  const inventorySheet = ss.getSheetByName("棚卸_最新") || ss.getSheetByName("棚卸_取り込み") || ss.getSheetByName("棚卸");
  const outputSheetName = "口座別余力_月別推移";

  if (!assetSheet) return;

  // 出力シート準備
  let outputSheet = ss.getSheetByName(outputSheetName);
  if (!outputSheet) {
    outputSheet = ss.insertSheet(outputSheetName);
  } else {
    outputSheet.clearContents();
  }

  const assetData = assetSheet.getDataRange().getValues();
  if (assetData.length <= 1) return;

  const masterData = masterSheet ? masterSheet.getDataRange().getValues() : [];
  const paymentData = paymentSheet ? paymentSheet.getDataRange().getValues() : [];
  const repaymentData = repaymentSheet ? repaymentSheet.getDataRange().getValues() : [];

  // 金額パース補助関数
  function parseAmount(val) {
    if (typeof val === "number") return val;
    if (typeof val === "string") {
      const num = parseFloat(val.replace(/[¥,]/g, ""));
      return isNaN(num) ? 0 : num;
    }
    return 0;
  }

  // 1. 口座別返済額マップの作成 (ym -> { bankName: amount })
  const repaymentMap = {};
  if (repaymentData.length > 1) {
    const repayHeaders = repaymentData[0];
    repaymentData.slice(1).forEach(row => {
      const dateVal = row[0];
      const ym = dateVal instanceof Date 
        ? Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM")
        : String(dateVal).substring(0, 7);
      if (!repaymentMap[ym]) repaymentMap[ym] = {};
      for (let i = 1; i < row.length; i++) {
        const bank = repayHeaders[i];
        repaymentMap[ym][bank] = parseAmount(row[i]);
      }
    });
  } else if (paymentData.length > 1 && masterData.length > 1) {
    // 口座別カード返済集計が無い場合はカード返済額_整形から算出（フォールバック）
    const cardToBank = {};
    for (let i = 1; i < masterData.length; i++) {
      const card = masterData[i][1];
      const bank = masterData[i][2];
      if (card && bank) cardToBank[card] = bank;
    }
    const payHeaders = paymentData[0];
    paymentData.slice(1).forEach(row => {
      const dateVal = row[0];
      const ym = dateVal instanceof Date 
        ? Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM")
        : String(dateVal).substring(0, 7);
      if (!repaymentMap[ym]) repaymentMap[ym] = {};
      payHeaders.forEach((card, colIdx) => {
        if (colIdx >= 2 && cardToBank[card]) {
          const bank = cardToBank[card];
          const amt = parseAmount(row[colIdx]);
          repaymentMap[ym][bank] = (repaymentMap[ym][bank] || 0) + amt;
        }
      });
    });
  }

  // 2. 「推測在庫高_推移」シートから「月末推測在庫高」を取得
  const estimatedStockByMonth = {};
  let latestEstimatedStock = 0;
  if (estimatedStockSheet) {
    const stockData = estimatedStockSheet.getDataRange().getValues();
    if (stockData.length > 1) {
      const headers = stockData[0].map(h => String(h).trim());
      const idxYm = headers.findIndex(h => h.includes("年月") || h.includes("月"));
      const idxStock = headers.findIndex(h => h.includes("月末推測在庫高") || h.includes("推測在庫高") || h.includes("在庫高"));

      const colYm = idxYm !== -1 ? idxYm : 0;
      const colStock = idxStock !== -1 ? idxStock : (headers.length >= 4 ? 3 : headers.length - 1);

      for (let i = 1; i < stockData.length; i++) {
        const row = stockData[i];
        if (!row || row.length === 0) continue;
        const ymVal = row[colYm];
        const ym = ymVal instanceof Date
          ? Utilities.formatDate(ymVal, Session.getScriptTimeZone(), "yyyy-MM")
          : String(ymVal).substring(0, 7);
        const stockAmt = parseAmount(row[colStock]);
        if (ym) {
          estimatedStockByMonth[ym] = stockAmt;
          latestEstimatedStock = stockAmt;
        }
      }
    }
  }

  // フォールバック用の棚卸額取得
  const inventoryByMonth = {};
  let defaultInventory = 0;
  if (inventorySheet) {
    const invData = inventorySheet.getDataRange().getValues();
    if (invData.length > 1) {
      const headers = invData[0].map(h => String(h).trim());
      let amountColIdx = headers.findIndex(h => h.includes("実在庫額") || h.includes("在庫額"));
      let dateColIdx = headers.findIndex(h => h.includes("棚卸日") || h.includes("日付"));
      if (dateColIdx === -1) dateColIdx = 0;
      if (amountColIdx === -1) amountColIdx = headers.length > 1 ? 1 : 0;

      invData.slice(1).forEach(row => {
        const d = row[dateColIdx];
        const amt = parseAmount(row[amountColIdx]);
        if (amt > 0) {
          const ym = d instanceof Date 
            ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM")
            : (d ? String(d).substring(0, 7) : "");
          if (ym) {
            inventoryByMonth[ym] = amt;
          }
          defaultInventory = amt;
        }
      });
    }
  }

  // 3. 資産シートの各月の最終レコード（月末時点の残高）を抽出
  const assetHeaders = assetData[0];
  const assetRows = assetData.slice(1);

  const banks = ["TKS銀行", "RKT銀行", "SZK信金"];
  const points = ["MRPY", "GNKN", "PYPY", "RTPY_YSK", "RTPY_SZK"];
  const receivables = ["AMZ売掛"];
  const excludedKeys = new Set(["タイムスタンプ", ...banks, ...points, ...receivables]);

  // 月ごとに最新の行をまとめる
  const monthlyLatestAsset = {};
  let latestOverallDate = null;
  let latestOverallRow = null;

  assetRows.forEach(row => {
    const ts = new Date(row[0]);
    if (isNaN(ts.getTime())) return;
    const ym = Utilities.formatDate(ts, Session.getScriptTimeZone(), "yyyy-MM");

    if (!monthlyLatestAsset[ym] || ts > monthlyLatestAsset[ym].timestamp) {
      monthlyLatestAsset[ym] = {
        timestamp: ts,
        row: row
      };
    }

    if (!latestOverallDate || ts > latestOverallDate) {
      latestOverallDate = ts;
      latestOverallRow = row;
    }
  });

  const months = Object.keys(monthlyLatestAsset).sort();
  if (months.length === 0) return;

  const currentMonthStr = Utilities.formatDate(latestOverallDate, Session.getScriptTimeZone(), "yyyy-MM");

  // 次月を求める
  const latestMonthDate = new Date(currentMonthStr + "-01");
  const nextMonthDate = new Date(latestMonthDate);
  nextMonthDate.setMonth(nextMonthDate.getMonth() + 1);
  const nextMonthStr = Utilities.formatDate(nextMonthDate, Session.getScriptTimeZone(), "yyyy-MM");

  // 出力ヘッダー
  const output = [[
    "年月", "種別",
    "TKS銀行(資産)", "TKS銀行(返済)", "TKS銀行(余力)",
    "RKT銀行(資産)", "RKT銀行(返済)", "RKT銀行(余力)",
    "SZK信金(資産)", "SZK信金(返済)", "SZK信金(余力)",
    "その他現金", "ポイント・売掛金", "棚卸資産",
    "総資産", "総返済額", "総余力"
  ]];

  // 過去月〜当月のデータ行作成
  let lastKnownStock = latestEstimatedStock || defaultInventory;

  months.forEach(ym => {
    const item = monthlyLatestAsset[ym];
    const row = item.row;
    const isCurrent = (ym === currentMonthStr);
    const typeLabel = isCurrent ? "当月(最新)" : "確定(月末)";

    const repayThisMonth = repaymentMap[ym] || {};

    let totalBankAsset = 0;
    let totalRepay = 0;
    let totalBankDiff = 0;

    const rowData = [ym, typeLabel];

    // 各銀行
    banks.forEach(bank => {
      const idx = assetHeaders.indexOf(bank);
      const asset = idx !== -1 ? parseAmount(row[idx]) : 0;
      const repay = repayThisMonth[bank] || 0;
      const diff = asset - repay;

      rowData.push(asset, repay, diff);

      totalBankAsset += asset;
      totalRepay += repay;
      totalBankDiff += diff;
    });

    // その他現金
    let otherCash = 0;
    assetHeaders.forEach((header, idx) => {
      if (idx > 0 && !excludedKeys.has(header)) {
        otherCash += parseAmount(row[idx]);
      }
    });
    rowData.push(otherCash);

    // ポイント・売掛金
    let pointAndReceivables = 0;
    [...points, ...receivables].forEach(key => {
      const idx = assetHeaders.indexOf(key);
      if (idx !== -1) {
        pointAndReceivables += parseAmount(row[idx]);
      }
    });
    rowData.push(pointAndReceivables);

    // N列：棚卸資産（推測在庫高_推移シートの「月末推測在庫高」を適用）
    let stockAmt = 0;
    if (estimatedStockByMonth[ym] !== undefined) {
      stockAmt = estimatedStockByMonth[ym];
      lastKnownStock = stockAmt;
    } else if (inventoryByMonth[ym] !== undefined) {
      stockAmt = inventoryByMonth[ym];
      lastKnownStock = stockAmt;
    } else {
      stockAmt = lastKnownStock;
    }
    rowData.push(stockAmt);

    // 総計
    const totalAsset = totalBankAsset + otherCash + pointAndReceivables + stockAmt;
    const totalDiff = totalBankDiff + otherCash + pointAndReceivables + stockAmt;

    rowData.push(totalAsset, totalRepay, totalDiff);

    output.push(rowData);
  });

  // 次月（予想）の行を追加（次月返済データがある場合）
  if (repaymentMap[nextMonthStr] && latestOverallRow) {
    const nextRepayMap = repaymentMap[nextMonthStr];
    const prevMonthItem = monthlyLatestAsset[currentMonthStr];
    const prevRow = prevMonthItem ? prevMonthItem.row : latestOverallRow;

    let totalBankAsset = 0;
    let totalRepay = 0;
    let totalBankDiff = 0;

    const nextRowData = [nextMonthStr, "次月(予想)"];

    // 銀行
    banks.forEach(bank => {
      const idx = assetHeaders.indexOf(bank);
      const currentAsset = idx !== -1 ? parseAmount(prevRow[idx]) : 0;
      const currentRepay = (repaymentMap[currentMonthStr] && repaymentMap[currentMonthStr][bank]) || 0;
      const currentDiff = currentAsset - currentRepay;
      
      const nextRepay = nextRepayMap[bank] || 0;
      const nextDiff = currentDiff - nextRepay;

      nextRowData.push(currentDiff, nextRepay, nextDiff);

      totalBankAsset += currentDiff;
      totalRepay += nextRepay;
      totalBankDiff += nextDiff;
    });

    // その他現金
    let otherCash = 0;
    assetHeaders.forEach((header, idx) => {
      if (idx > 0 && !excludedKeys.has(header)) {
        otherCash += parseAmount(prevRow[idx]);
      }
    });
    nextRowData.push(otherCash);

    // ポイント・売掛金
    let pointAndReceivables = 0;
    [...points, ...receivables].forEach(key => {
      const idx = assetHeaders.indexOf(key);
      if (idx !== -1) {
        pointAndReceivables += parseAmount(prevRow[idx]);
      }
    });
    nextRowData.push(pointAndReceivables);

    // N列：棚卸資産（次月予想は直近推測在庫高を引き継ぐ）
    const nextStockAmt = estimatedStockByMonth[nextMonthStr] ?? estimatedStockByMonth[currentMonthStr] ?? lastKnownStock;
    nextRowData.push(nextStockAmt);

    // 総計
    const totalAsset = totalBankAsset + otherCash + pointAndReceivables + nextStockAmt;
    const totalDiff = totalBankDiff + otherCash + pointAndReceivables + nextStockAmt;

    nextRowData.push(totalAsset, totalRepay, totalDiff);

    output.push(nextRowData);
  }

  const numRows = output.length;
  const numCols = output[0].length;

  // シートに出力
  outputSheet.getRange(1, 1, numRows, numCols).setValues(output);

  // --- 書式とデザイン設定 ---
  // 全体：垂直方向中央揃え
  outputSheet.getRange(1, 1, numRows, numCols).setVerticalAlignment("middle");

  // ヘッダー行（1行目）
  const headerRange = outputSheet.getRange(1, 1, 1, numCols);
  headerRange
    .setBackground("#2C3E50")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  outputSheet.setRowHeight(1, 32);

  // データ行
  if (numRows > 1) {
    // 年月・種別（A, B列）: センタリング
    outputSheet.getRange(2, 1, numRows - 1, 2).setHorizontalAlignment("center");

    // 金額列（C列〜Q列）: 通貨フォーマット & 右揃え
    outputSheet.getRange(2, 3, numRows - 1, numCols - 2)
      .setNumberFormat('"¥"#,##0')
      .setHorizontalAlignment("right");

    // 各データ行の高さ設定
    for (let r = 2; r <= numRows; r++) {
      outputSheet.setRowHeight(r, 26);
    }

    // 総資産・総返済額・総余力（O, P, Q列）のハイライト背景色（淡いグレー）
    outputSheet.getRange(2, 15, numRows - 1, 3).setBackground("#F8F9FA");

    // 次月(予想)の行があれば薄いハイライト
    if (output[numRows - 1][1] === "次月(予想)") {
      outputSheet.getRange(numRows, 1, 1, numCols).setBackground("#EFF6FF");
    }
  }

  // --- タイトルが確実に表示される列幅調整 ---
  function calculateTitleWidth(text) {
    if (!text) return 60;
    const str = String(text);
    let w = 0;
    for (let i = 0; i < str.length; i++) {
      const code = str.charCodeAt(i);
      if ((code >= 0x3000 && code <= 0x9fff) || (code >= 0xff01 && code <= 0xff60)) {
        w += 14.5; // 全角
      } else {
        w += 8.5;  // 半角
      }
    }
    return Math.ceil(w + 24); // 左右余白
  }

  outputSheet.autoResizeColumns(1, numCols);
  for (let col = 1; col <= numCols; col++) {
    const autoW = outputSheet.getColumnWidth(col);
    const headerTitle = output[0][col - 1];
    const minTitleW = calculateTitleWidth(headerTitle);
    outputSheet.setColumnWidth(col, Math.max(autoW + 8, minTitleW));
  }
}
