/**
 * 商品マスタ_統合の販売状況更新・出品可能一覧生成スクリプト
 * 
 * 1. 「統合_販売記録」（または「販売速報」）のデータをもとに「商品マスタ_統合」の販売ステータス等を更新
 * 2. 「販売済」の行を条件付き書式で自動グレーアウト
 * 3. 「出品可能_在庫一覧」（未販売・残在庫あり）シートを自動生成
 */
function 商品マスタ_販売状況更新() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. シートの取得 ---
  const masterSheetName = "商品マスタ_統合";
  const masterSheet = ss.getSheetByName(masterSheetName);
  if (!masterSheet) {
    throw new Error(`「${masterSheetName}」シートが見つかりません。`);
  }

  // 販売記録シートの検索（統合_販売記録、販売速報（フォーム回答）、販売記録 の順で探索）
  const salesSheetCandidates = [
    "統合_販売記録",
    "販売速報（フォーム回答）",
    "販売記録",
    "販売速報"
  ];
  let salesSheet = null;
  for (const name of salesSheetCandidates) {
    const sh = ss.getSheetByName(name);
    if (sh) {
      salesSheet = sh;
      break;
    }
  }

  if (!salesSheet) {
    throw new Error(
      "販売実績シートが見つかりません。対象候補: " + salesSheetCandidates.join(", ")
    );
  }

  // --- 2. 販売記録データの読み込み & 集計 ---
  const salesData = salesSheet.getDataRange().getValues();
  if (salesData.length < 2) {
    SpreadsheetApp.getActive().toast("販売実績データがありません。", "販売状況更新", 5);
    return;
  }
  const salesHeaders = salesData[0].map(h => String(h).trim());

  const idxSalesNo = findHeaderIndex_(salesHeaders, ["商品マスタNo", "商品マスタ No", "No", "No."]);
  const idxSalesDate = findHeaderIndex_(salesHeaders, ["販売日", "日付", "売上日", "date"]);
  const idxSalesQty = findHeaderIndex_(salesHeaders, ["販売数", "数量", "個数"]);
  const idxSalesPrice = findHeaderIndex_(salesHeaders, ["販売価格", "売価", "販売額", "売上金額", "金額"]);
  const idxSalesProfit = findHeaderIndex_(salesHeaders, ["実利益", "利益", "純利益", "粗利"]);

  if (idxSalesNo === -1) {
    throw new Error("販売実績シートに「商品マスタNo」列が見つかりません。");
  }

  // 商品マスタNoごとの販売実績集計
  // key: 商品マスタNo -> { totalQty, lastDate, lastPrice, totalProfit }
  const salesSummary = {};

  for (let i = 1; i < salesData.length; i++) {
    const row = salesData[i];
    const no = String(row[idxSalesNo] ?? "").trim();
    if (!no) continue;

    const rawDate = idxSalesDate !== -1 ? row[idxSalesDate] : null;
    const sDate = parseDate_(rawDate);
    const qty = idxSalesQty !== -1 ? Math.max(1, parseNumber_(row[idxSalesQty])) : 1;
    const price = idxSalesPrice !== -1 ? parseAmount_(row[idxSalesPrice]) : 0;
    const profit = idxSalesProfit !== -1 ? parseAmount_(row[idxSalesProfit]) : 0;

    if (!salesSummary[no]) {
      salesSummary[no] = {
        totalQty: 0,
        lastDate: null,
        lastPrice: 0,
        totalProfit: 0
      };
    }

    salesSummary[no].totalQty += qty;
    salesSummary[no].totalProfit += profit;

    // 最新の販売日と直近価格を記録
    if (sDate) {
      if (!salesSummary[no].lastDate || sDate >= salesSummary[no].lastDate) {
        salesSummary[no].lastDate = sDate;
        if (price > 0) salesSummary[no].lastPrice = price;
      }
    } else if (price > 0 && salesSummary[no].lastPrice === 0) {
      salesSummary[no].lastPrice = price;
    }
  }

  // --- 3. 商品マスタ_統合のヘッダー解析と必要列の確保 ---
  const masterLastRow = masterSheet.getLastRow();
  const masterLastCol = masterSheet.getLastColumn();
  if (masterLastRow < 1) {
    throw new Error("「商品マスタ_統合」シートにデータがありません。");
  }

  let masterHeaders = masterSheet.getRange(1, 1, 1, masterLastCol).getValues()[0].map(h => String(h).trim());

  // 既存の入力列インデックス取得
  const idxMasterNo = findHeaderIndex_(masterHeaders, ["商品マスタ No", "商品マスタNo", "No", "No."]);
  const idxPurchaseQty = findHeaderIndex_(masterHeaders, ["購入数", "数量", "仕入数", "個数"]);
  const idxPurchaseDate = findHeaderIndex_(masterHeaders, ["仕入日", "購入日", "日付"]);
  const idxUnitPrice = findHeaderIndex_(masterHeaders, ["実質仕入額(単品)", "仕入評価額", "単価", "実質仕入額"]);
  const idxProductName = findHeaderIndex_(masterHeaders, ["商品名", "品名"]);
  const idxPurchasePurpose = findHeaderIndex_(masterHeaders, ["購入目的", "目的"]);
  const idxMercariCode = findHeaderIndex_(masterHeaders, ["メルカリ管理コード", "管理コード", "コード"]);

  if (idxMasterNo === -1) {
    throw new Error("「商品マスタ_統合」シートに「商品マスタ No」列が見つかりません。");
  }

  // 追加/更新したい対象列
  const targetNewCols = [
    "販売ステータス",
    "残在庫数",
    "最終販売日",
    "販売価格",
    "実利益"
  ];

  // 不足している列を右端に追加
  let currentLastCol = masterLastCol;
  const colIndexMap = {};

  targetNewCols.forEach(colName => {
    let colIdx = masterHeaders.indexOf(colName);
    if (colIdx === -1) {
      currentLastCol++;
      masterSheet.getRange(1, currentLastCol).setValue(colName)
        .setFontWeight("bold")
        .setBackground("#d9ead3"); // 薄緑
      masterHeaders.push(colName);
      colIdx = masterHeaders.length - 1;
    }
    colIndexMap[colName] = colIdx;
  });

  // --- 4. 商品マスタ_統合データの更新 ---
  const numDataRows = masterLastRow - 1;
  if (numDataRows > 0) {
    const fullMasterData = masterSheet.getRange(2, 1, numDataRows, masterHeaders.length).getValues();

    // 書き込み用バッファを作成
    const statusColValues = [];
    const stockQtyColValues = [];
    const lastDateColValues = [];
    const salesPriceColValues = [];
    const profitColValues = [];

    for (let r = 0; r < numDataRows; r++) {
      const row = fullMasterData[r];
      const no = String(row[idxMasterNo] ?? "").trim();
      const pQty = idxPurchaseQty !== -1 ? Math.max(0, parseNumber_(row[idxPurchaseQty])) : 1;

      if (!no) {
        statusColValues.push([""]);
        stockQtyColValues.push([""]);
        lastDateColValues.push([""]);
        salesPriceColValues.push([""]);
        profitColValues.push([""]);
        continue;
      }

      const sale = salesSummary[no];
      const sQty = sale ? sale.totalQty : 0;
      const remainQty = Math.max(0, pQty - sQty);

      let status = "未販売";
      if (sQty > 0) {
        if (sQty >= pQty) {
          status = "販売済";
        } else {
          status = "一部販売済";
        }
      }

      const dateStr = sale && sale.lastDate ? Utilities.formatDate(sale.lastDate, Session.getScriptTimeZone(), "yyyy/MM/dd") : "";
      const priceVal = sale && sale.lastPrice ? sale.lastPrice : "";
      const profitVal = sale && sale.totalProfit ? sale.totalProfit : "";

      statusColValues.push([status]);
      stockQtyColValues.push([remainQty]);
      lastDateColValues.push([dateStr]);
      salesPriceColValues.push([priceVal]);
      profitColValues.push([profitVal]);
    }

    // 列ごとにまとめて書き込み
    masterSheet.getRange(2, colIndexMap["販売ステータス"] + 1, numDataRows, 1).setValues(statusColValues);
    masterSheet.getRange(2, colIndexMap["残在庫数"] + 1, numDataRows, 1).setValues(stockQtyColValues);
    masterSheet.getRange(2, colIndexMap["最終販売日"] + 1, numDataRows, 1).setValues(lastDateColValues);
    masterSheet.getRange(2, colIndexMap["販売価格"] + 1, numDataRows, 1).setValues(salesPriceColValues);
    masterSheet.getRange(2, colIndexMap["実利益"] + 1, numDataRows, 1).setValues(profitColValues);

    // 書式設定
    masterSheet.getRange(2, colIndexMap["残在庫数"] + 1, numDataRows, 1).setNumberFormat('#,##0');
    masterSheet.getRange(2, colIndexMap["販売価格"] + 1, numDataRows, 1).setNumberFormat('"¥"#,##0');
    masterSheet.getRange(2, colIndexMap["実利益"] + 1, numDataRows, 1).setNumberFormat('"¥"#,##0');
  }

  // --- 5. 条件付き書式の設定（販売済行のグレーアウト） ---
  setupMasterConditionalFormat_(masterSheet, colIndexMap["販売ステータス"] + 1, masterLastRow, masterHeaders.length);

  // --- 6. 「出品可能_在庫一覧」シートの自動生成 ---
  generateAvailableStockSheet_(ss, masterSheet, {
    idxMasterNo: idxMasterNo,
    idxProductName: idxProductName,
    idxMercariCode: idxMercariCode,
    idxPurchaseDate: idxPurchaseDate,
    idxUnitPrice: idxUnitPrice,
    idxPurchaseQty: idxPurchaseQty,
    idxPurchasePurpose: idxPurchasePurpose,
    idxStockQty: colIndexMap["残在庫数"],
    idxStatus: colIndexMap["販売ステータス"]
  });

  SpreadsheetApp.getActive().toast(
    "販売状況の反映、条件付き書式、出品可能一覧の作成が完了しました。",
    "商品マスタ更新",
    6
  );
}

/**
 * 商品マスタ_統合の「販売済」行をグレーアウトする条件付き書式を設定
 */
function setupMasterConditionalFormat_(sheet, statusColNumber, lastRow, lastCol) {
  if (lastRow < 2 || lastCol < 1) return;

  const statusColLetter = columnToLetter_(statusColNumber);
  const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);

  // 既存のルールを保持しつつ、当スクリプト用のルールを優先設定
  const existingRules = sheet.getConditionalFormatRules();
  const filteredRules = existingRules.filter(rule => {
    // 既存のステータス条件式がある場合は重複を避けるために除外
    const cond = rule.getBooleanCondition();
    if (!cond) return true;
    const formula = cond.getCriteriaValues()[0] || "";
    return !String(formula).includes(`$${statusColLetter}2="販売済"`);
  });

  // 「販売済」の行全体を薄いグレー背景 ＆ グレー文字にする
  const ruleSold = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${statusColLetter}2="販売済"`)
    .setBackground("#f3f3f3")
    .setFontColor("#888888")
    .setRanges([dataRange])
    .build();

  // 「一部販売済」の行全体を薄い黄色背景にする（任意）
  const rulePartial = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${statusColLetter}2="一部販売済"`)
    .setBackground("#fff2cc")
    .setRanges([dataRange])
    .build();

  filteredRules.unshift(rulePartial);
  filteredRules.unshift(ruleSold);
  sheet.setConditionalFormatRules(filteredRules);
}

/**
 * 「出品可能_在庫一覧」（未販売・残在庫あり）シートを生成・更新
 */
function generateAvailableStockSheet_(ss, masterSheet, indices) {
  const outputSheetName = "出品可能_在庫一覧";
  let outSheet = ss.getSheetByName(outputSheetName);

  if (outSheet) {
    outSheet.clear();
  } else {
    outSheet = ss.insertSheet(outputSheetName);
  }

  const masterData = masterSheet.getDataRange().getValues();
  if (masterData.length < 2) return;

  // 出品可能一覧のヘッダー定義
  const headers = [
    "商品マスタ No",
    "商品名",
    "メルカリ管理コード",
    "仕入日",
    "実質仕入額(単品)",
    "購入数",
    "残在庫数",
    "販売ステータス",
    "購入目的"
  ];

  const outRows = [headers];
  let totalStockCount = 0;
  let totalStockAmount = 0;

  for (let i = 1; i < masterData.length; i++) {
    const row = masterData[i];
    const no = String(row[indices.idxMasterNo] ?? "").trim();
    if (!no) continue;

    const stockQty = parseNumber_(row[indices.idxStockQty]);
    const status = String(row[indices.idxStatus] ?? "").trim();

    // 残在庫が0より大きい（未販売または一部販売済）のものだけを抽出
    if (stockQty > 0 && status !== "販売済") {
      const pDate = indices.idxPurchaseDate !== -1 ? parseDate_(row[indices.idxPurchaseDate]) : null;
      const unitPrice = indices.idxUnitPrice !== -1 ? parseAmount_(row[indices.idxUnitPrice]) : 0;
      const pQty = indices.idxPurchaseQty !== -1 ? parseNumber_(row[indices.idxPurchaseQty]) : 0;
      const name = indices.idxProductName !== -1 ? String(row[indices.idxProductName] ?? "") : "";
      const code = indices.idxMercariCode !== -1 ? String(row[indices.idxMercariCode] ?? "") : "";
      const purpose = indices.idxPurchasePurpose !== -1 ? String(row[indices.idxPurchasePurpose] ?? "") : "";

      outRows.push([
        no,
        name,
        code,
        pDate ? Utilities.formatDate(pDate, Session.getScriptTimeZone(), "yyyy/MM/dd") : "",
        unitPrice,
        pQty,
        stockQty,
        status,
        purpose
      ]);

      totalStockCount += stockQty;
      totalStockAmount += (stockQty * unitPrice);
    }
  }

  if (outRows.length === 1) {
    outRows.push(["現在、出品可能な在庫はありません。", "", "", "", "", "", "", "", ""]);
  }

  // シートへ一括書き込み
  outSheet.getRange(1, 1, outRows.length, headers.length).setValues(outRows);

  // --- 書式設定 ---
  // ヘッダー行
  outSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#d0e0e3"); // ライトシアン

  if (outRows.length > 1 && totalStockCount > 0) {
    const numData = outRows.length - 1;
    // 実質仕入額（E列=5）
    outSheet.getRange(2, 5, numData, 1).setNumberFormat('"¥"#,##0');
    // 購入数（F列=6）、残在庫数（G列=7）
    outSheet.getRange(2, 6, numData, 2).setNumberFormat('#,##0');
  }

  // 列幅調整
  outSheet.autoResizeColumns(1, headers.length);
  outSheet.setColumnWidth(2, 260); // 商品名列
  if (outRows.length > 1) {
    outSheet.getRange(2, 2, outRows.length - 1, 1).setWrap(true);
  }
}

/* ===== 共通ヘルパー関数 ===== */

/**
 * 候補リストから一致するヘッダーのインデックスを返す
 */
function findHeaderIndex_(headers, candidates) {
  for (const cand of candidates) {
    const idx = headers.indexOf(cand);
    if (idx !== -1) return idx;
  }
  return -1;
}

/**
 * 日付のパース
 */
function parseDate_(val) {
  const minValidDate = new Date("2020-01-01");
  if (val instanceof Date) {
    return val >= minValidDate ? val : null;
  }
  if (!val) return null;
  const d = new Date(val);
  if (isNaN(d.getTime())) return null;
  return d >= minValidDate ? d : null;
}

/**
 * 金額数値のパース
 */
function parseAmount_(val) {
  if (typeof val === "number") return val;
  if (typeof val === "string") {
    const num = parseFloat(val.replace(/[¥,]/g, ""));
    return isNaN(num) ? 0 : num;
  }
  return 0;
}

/**
 * 数量数値のパース
 */
function parseNumber_(val) {
  if (typeof val === "number") return val;
  if (typeof val === "string") {
    const num = parseInt(val.replace(/[,]/g, ""), 10);
    return isNaN(num) ? 0 : num;
  }
  return 0;
}

/**
 * 列番号 → 列記号変換 (例: 1 -> A, 23 -> W, 27 -> AA)
 */
function columnToLetter_(col) {
  let temp = "", n = col;
  while (n > 0) {
    let rem = (n - 1) % 26;
    temp = String.fromCharCode(65 + rem) + temp;
    n = Math.floor((n - 1) / 26);
  }
  return temp;
}
