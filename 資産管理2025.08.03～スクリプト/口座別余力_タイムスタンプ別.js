function 口座別余力_タイムスタンプ別() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const assetSheet = ss.getSheetByName("資産");
  const repaymentSheet = ss.getSheetByName("カード返済額");
  const masterSheet = ss.getSheetByName("カードマスター");
  const outputSheetName = "口座別余力_タイムスタンプ別";

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

  // カード名 -> 引落口座 のマッピングを作成
  const cardToBank = {};
  for (let i = 1; i < masterData.length; i++) {
    const card = masterData[i][1];
    const bank = masterData[i][2];
    if (card && bank) {
      cardToBank[card] = bank;
    }
  }

  // カード返済額シートのヘッダー解析
  const repaymentHeaders = repaymentData[0] || [];
  const repaymentRows = repaymentData.slice(1);

  const atsumCols = [];
  const nextCols = [];

  repaymentHeaders.forEach((h, idx) => {
    const headerStr = String(h);
    if (headerStr.includes("当月")) {
      const cardName = headerStr.replace("当月", "").trim();
      const bank = cardToBank[cardName] || "";
      atsumCols.push({ index: idx, cardName: cardName, bank: bank });
    } else if (headerStr.includes("次月")) {
      const cardName = headerStr.replace("次月", "").trim();
      const bank = cardToBank[cardName] || "";
      nextCols.push({ index: idx, cardName: cardName, bank: bank });
    }
  });

  // タイムスタンプごとのカード返済額データマップ（または配列）を構築
  const repaymentMapByTime = [];
  repaymentRows.forEach(row => {
    const ts = new Date(row[0]);
    if (isNaN(ts.getTime())) return;

    const currentByBank = {};
    const nextByBank = {};

    atsumCols.forEach(col => {
      const amt = parseAmount(row[col.index]);
      if (col.bank) {
        currentByBank[col.bank] = (currentByBank[col.bank] || 0) + amt;
      }
    });

    nextCols.forEach(col => {
      const amt = parseAmount(row[col.index]);
      if (col.bank) {
        nextByBank[col.bank] = (nextByBank[col.bank] || 0) + amt;
      }
    });

    repaymentMapByTime.push({
      time: ts.getTime(),
      date: ts,
      currentByBank: currentByBank,
      nextByBank: nextByBank
    });
  });

  // 最も近いタイムスタンプの返済額データを取得する関数
  function getRepaymentForTime(targetTime) {
    if (repaymentMapByTime.length === 0) {
      return { currentByBank: {}, nextByBank: {} };
    }

    let minDiff = Infinity;
    let closest = repaymentMapByTime[0];

    for (let item of repaymentMapByTime) {
      const diff = Math.abs(item.time - targetTime);
      if (diff < minDiff) {
        minDiff = diff;
        closest = item;
      }
    }
    return closest;
  }

  const assetHeaders = assetData[0];
  const assetRows = assetData.slice(1);

  const banks = ["TKS銀行", "RKT銀行", "SZK信金"];
  const points = ["MRPY", "GNKN", "PYPY", "RTPY_YSK", "RTPY_SZK"];
  const receivables = ["AMZ売掛"];
  const excludedKeys = new Set(["タイムスタンプ", ...banks, ...points, ...receivables]);

  // 出力ヘッダー作成
  const outputHeaders = [
    "タイムスタンプ",
    "TKS銀行(資産)", "TKS銀行(返済当月)", "TKS銀行(余力当月)", "TKS銀行(返済次月)", "TKS銀行(余力次月)",
    "RKT銀行(資産)", "RKT銀行(返済当月)", "RKT銀行(余力当月)", "RKT銀行(返済次月)", "RKT銀行(余力次月)",
    "SZK信金(資産)", "SZK信金(返済当月)", "SZK信金(余力当月)", "SZK信金(返済次月)", "SZK信金(余力次月)",
    "その他現金", "ポイント・売掛金",
    "総資産", "当月総返済額", "当月総余力", "次月総返済額", "次月総余力"
  ];

  const result = [outputHeaders];

  // タイムスタンプ昇順でソートして処理
  const sortedAssetRows = assetRows.filter(row => {
    const d = new Date(row[0]);
    return !isNaN(d.getTime());
  }).sort((a, b) => new Date(a[0]) - new Date(b[0]));

  sortedAssetRows.forEach(row => {
    const ts = new Date(row[0]);
    const tsFormatted = Utilities.formatDate(ts, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
    const targetTime = ts.getTime();

    const repaymentDataForTs = getRepaymentForTime(targetTime);
    const currentRepayMap = repaymentDataForTs.currentByBank;
    const nextRepayMap = repaymentDataForTs.nextByBank;

    let totalAsset = 0;
    let totalCurrentRepay = 0;
    let totalNextRepay = 0;

    const rowOutput = [tsFormatted];

    // 各銀行の計算
    banks.forEach(bank => {
      const idx = assetHeaders.indexOf(bank);
      const assetAmt = idx !== -1 ? parseAmount(row[idx]) : 0;
      const repayCurrent = currentRepayMap[bank] || 0;
      const diffCurrent = assetAmt - repayCurrent;
      const repayNext = nextRepayMap[bank] || 0;
      const diffNext = diffCurrent - repayNext;

      rowOutput.push(assetAmt, repayCurrent, diffCurrent, repayNext, diffNext);

      totalAsset += assetAmt;
      totalCurrentRepay += repayCurrent;
      totalNextRepay += repayNext;
    });

    // その他現金
    let otherCash = 0;
    assetHeaders.forEach((header, idx) => {
      if (idx > 0 && !excludedKeys.has(header)) {
        otherCash += parseAmount(row[idx]);
      }
    });
    rowOutput.push(otherCash);
    totalAsset += otherCash;

    // ポイント・売掛金
    let pointAndReceivables = 0;
    [...points, ...receivables].forEach(key => {
      const idx = assetHeaders.indexOf(key);
      if (idx !== -1) {
        pointAndReceivables += parseAmount(row[idx]);
      }
    });
    rowOutput.push(pointAndReceivables);
    totalAsset += pointAndReceivables;

    // 総計
    const totalDiffCurrent = totalAsset - totalCurrentRepay;
    const totalDiffNext = totalDiffCurrent - totalNextRepay;

    rowOutput.push(totalAsset, totalCurrentRepay, totalDiffCurrent, totalNextRepay, totalDiffNext);

    result.push(rowOutput);
  });

  // シートに出力
  outputSheet.getRange(1, 1, result.length, result[0].length).setValues(result);

  // 通貨フォーマット設定
  if (result.length > 1) {
    outputSheet.getRange(2, 2, result.length - 1, result[0].length - 1).setNumberFormat('"¥"#,##0');
  }

  // タイトルが確実に表示される列幅調整
  function calculateTitleWidth(text) {
    if (!text) return 60;
    const str = String(text);
    let w = 0;
    for (let i = 0; i < str.length; i++) {
      const code = str.charCodeAt(i);
      if ((code >= 0x3000 && code <= 0x9fff) || (code >= 0xff01 && code <= 0xff60)) {
        w += 14.5;
      } else {
        w += 8.5;
      }
    }
    return Math.ceil(w + 24);
  }

  outputSheet.autoResizeColumns(1, result[0].length);
  for (let col = 1; col <= result[0].length; col++) {
    const autoW = outputSheet.getColumnWidth(col);
    const headerTitle = result[0][col - 1];
    const minTitleW = calculateTitleWidth(headerTitle);
    outputSheet.setColumnWidth(col, Math.max(autoW + 8, minTitleW));
  }
}
