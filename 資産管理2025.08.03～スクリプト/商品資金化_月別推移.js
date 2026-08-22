function 商品資金化_月別推移() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 対象シートの取得
  const stockHistorySheet = ss.getSheetByName("推測在庫高_推移");
  const salesCsvSheet = ss.getSheetByName("販売記録（CSV取り込み）");
  const matchSheet = ss.getSheetByName("販売記録CSV照合表");
  const quickSalesSheet = ss.getSheetByName("販売速報（フォーム回答）");
  const masterSheet = ss.getSheetByName("商品マスタ_統合");

  const outputSheetName = "商品資金化_月別推移";
  let outputSheet = ss.getSheetByName(outputSheetName);
  if (outputSheet) {
    outputSheet.clearContents();
  } else {
    outputSheet = ss.insertSheet(outputSheetName);
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

  function parseFeeAmount(val, price) {
    if (val === null || val === undefined || val === "") return 0;
    if (typeof val === "number") {
      if (val > 0 && val <= 1 && price > 0) return Math.round(price * val);
      return val;
    }
    if (typeof val === "string") {
      const cleanStr = val.trim();
      if (cleanStr.includes("%")) {
        const num = parseFloat(cleanStr.replace(/[%]/g, ""));
        return (!isNaN(num) && price > 0) ? Math.round(price * (num / 100)) : 0;
      }
      const num = parseFloat(cleanStr.replace(/[¥,]/g, ""));
      if (!isNaN(num)) {
        if (num > 0 && num <= 1 && price > 0) return Math.round(price * num);
        return num;
      }
    }
    return 0;
  }

  function parseNumber(val) {
    if (typeof val === "number") return val;
    if (typeof val === "string") {
      const num = parseInt(val.replace(/[,]/g, ""), 10);
      return isNaN(num) ? 0 : num;
    }
    return 0;
  }

  function findHeaderIndex(headers, candidates) {
    for (let i = 0; i < headers.length; i++) {
      const h = String(headers[i]).replace(/[\r\n\s　_（）\(\)]/g, "").toLowerCase();
      for (let j = 0; j < candidates.length; j++) {
        const cand = String(candidates[j]).replace(/[\r\n\s　_（）\(\)]/g, "").toLowerCase();
        if (h === cand || h.includes(cand)) {
          return i;
        }
      }
    }
    return -1;
  }

  function parseDate(val) {
    if (val instanceof Date) return val;
    if (!val) return null;
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }

  // --- 1. 「推測在庫高_推移」シートから月ごとの在庫・仕入・販売原価を取得 ---
  const monthlyStockData = {};
  const ymList = [];

  if (stockHistorySheet) {
    const stockData = stockHistorySheet.getDataRange().getValues();
    if (stockData.length > 1) {
      const headers = stockData[0].map(h => String(h).trim());
      const idxYm = findHeaderIndex(headers, ["年月", "月"]);
      const idxPurchase = findHeaderIndex(headers, ["当月仕入額", "仕入額", "仕入"]);
      const idxCost = findHeaderIndex(headers, ["当月販売原価", "販売原価", "原価"]);
      const idxEndStock = findHeaderIndex(headers, ["月末推測在庫高", "推測在庫高", "在庫高", "月末在庫"]);

      for (let i = 1; i < stockData.length; i++) {
        const row = stockData[i];
        if (!row || row.length === 0) continue;
        const ymVal = row[idxYm !== -1 ? idxYm : 0];
        const ym = ymVal instanceof Date
          ? Utilities.formatDate(ymVal, Session.getScriptTimeZone(), "yyyy-MM")
          : String(ymVal).substring(0, 7);
        if (!ym || ym === "合計" || ym.length < 7) continue;

        const purchase = idxPurchase !== -1 ? parseAmount(row[idxPurchase]) : 0;
        const cost = idxCost !== -1 ? parseAmount(row[idxCost]) : 0;
        const endStock = idxEndStock !== -1 ? parseAmount(row[idxEndStock]) : 0;

        monthlyStockData[ym] = {
          purchase: purchase,
          cost: cost,
          endStock: endStock,
          startStock: 0
        };
        ymList.push(ym);
      }
    }
  }

  // ymList をソートして月初在庫を計算
  ymList.sort();
  let prevEndStock = null;
  ymList.forEach(ym => {
    const item = monthlyStockData[ym];
    if (prevEndStock === null) {
      // 初月: 月末在庫 + 出庫原価 - 当月仕入 (または0)
      item.startStock = Math.max(0, item.endStock + item.cost - item.purchase);
    } else {
      item.startStock = prevEndStock;
    }
    prevEndStock = item.endStock;
  });

  // --- 2. 販売データから月ごとの「販売総額（売上）」と「手取り入金額（純売上）」を集計 ---
  const monthlySalesData = {};

  // 2-A. 販売記録（CSV取り込み）の読み込み
  const csvPriceMap = new Map();
  if (salesCsvSheet) {
    const csvData = salesCsvSheet.getDataRange().getValues();
    if (csvData.length > 1) {
      const srcHeaders = csvData[0].map(h => String(h).trim());
      const idxId = findHeaderIndex(srcHeaders, ["商品ID", "商品id", "id", "ID", "商品ＩＤ"]);
      const idxDate = findHeaderIndex(srcHeaders, ["購入日時", "販売日時", "取引日時", "日時", "日付"]);
      const idxPrice = findHeaderIndex(srcHeaders, ["価格", "商品価格", "販売価格", "売上金額", "金額"]);
      const idxFee = findHeaderIndex(srcHeaders, ["販売手数料金額", "販売手数料", "手数料額", "手数料"]);
      const idxShipping = findHeaderIndex(srcHeaders, ["配送料", "送料"]);
      const idxProfit = findHeaderIndex(srcHeaders, ["販売利益", "純利益", "入金額", "振込金額"]);

      for (let i = 1; i < csvData.length; i++) {
        const row = csvData[i];
        const id = idxId !== -1 ? String(row[idxId]).trim() : "";
        if (!id) continue;

        const dateVal = parseDate(idxDate !== -1 ? row[idxDate] : null);
        const price = idxPrice !== -1 ? parseAmount(row[idxPrice]) : 0;
        let fee = idxFee !== -1 ? parseFeeAmount(row[idxFee], price) : 0;
        if (fee === 0 && price > 0) {
          fee = Math.round(price * 0.1);
        }
        const shipping = idxShipping !== -1 ? parseAmount(row[idxShipping]) : 0;
        
        let netProfit = 0;
        if (idxProfit !== -1 && parseAmount(row[idxProfit]) > 0) {
          netProfit = parseAmount(row[idxProfit]);
        } else if (price > 0) {
          netProfit = Math.max(0, price - fee - shipping);
        }

        csvPriceMap.set(id, {
          price: price,
          net: netProfit,
          date: dateVal
        });
      }
    }
  }

  // 2-B. 「販売記録CSV照合表」から確定分の売上を集計
  const processedIds = new Set();
  const processedMasterNos = new Set();

  if (matchSheet) {
    const matchData = matchSheet.getDataRange().getValues();
    if (matchData.length > 1) {
      const headers = matchData[0].map(h => String(h).trim());
      const idxId = findHeaderIndex(headers, ["商品ID", "商品id", "id", "ID"]);
      const idxMasterNo = findHeaderIndex(headers, ["商品マスタ No", "商品マスタNo", "No"]);
      const idxDate = findHeaderIndex(headers, ["購入日時", "販売日時", "日付"]);

      for (let i = 1; i < matchData.length; i++) {
        const row = matchData[i];
        const id = idxId !== -1 ? String(row[idxId]).trim() : "";
        const masterNo = idxMasterNo !== -1 ? String(row[idxMasterNo]).trim() : "";
        const rawDate = idxDate !== -1 ? row[idxDate] : null;
        const sDate = parseDate(rawDate);

        if (sDate) {
          const ym = Utilities.formatDate(sDate, Session.getScriptTimeZone(), "yyyy-MM");
          if (!monthlySalesData[ym]) {
            monthlySalesData[ym] = { salesAmount: 0, netAmount: 0, count: 0 };
          }

          const csvInfo = csvPriceMap.get(id);
          const price = csvInfo ? csvInfo.price : 0;
          const net = csvInfo ? csvInfo.net : (price > 0 ? Math.round(price * 0.9) : 0);

          monthlySalesData[ym].salesAmount += price;
          monthlySalesData[ym].netAmount += net;
          monthlySalesData[ym].count += 1;

          if (id) processedIds.add(id);
          if (masterNo) processedMasterNos.add(masterNo);
        }
      }
    }
  }

  // 2-C. 「販売速報（フォーム回答）」から未確定分（速報のみ）の売上を補完
  if (quickSalesSheet) {
    const quickData = quickSalesSheet.getDataRange().getValues();
    if (quickData.length > 1) {
      const headers = quickData[0].map(h => String(h).trim());
      const idxNo = findHeaderIndex(headers, ["商品マスタNo", "商品マスタ No", "No"]);
      const idxDate = findHeaderIndex(headers, ["販売日", "日付"]);
      const idxPrice = findHeaderIndex(headers, ["販売価格", "売上", "価格", "金額"]);
      const idxFee = findHeaderIndex(headers, ["販売手数料金額", "販売手数料額", "手数料金額", "手数料額", "販売手数料", "手数料"]);
      const idxShipping = findHeaderIndex(headers, ["送料", "配送料", "送料額"]);

      for (let i = 1; i < quickData.length; i++) {
        const row = quickData[i];
        const no = idxNo !== -1 ? String(row[idxNo]).trim() : "";
        if (no && processedMasterNos.has(no)) continue;

        const sDate = parseDate(idxDate !== -1 ? row[idxDate] : null);
        if (sDate) {
          const ym = Utilities.formatDate(sDate, Session.getScriptTimeZone(), "yyyy-MM");
          if (!monthlySalesData[ym]) {
            monthlySalesData[ym] = { salesAmount: 0, netAmount: 0, count: 0 };
          }

          const price = idxPrice !== -1 ? parseAmount(row[idxPrice]) : 0;
          let fee = idxFee !== -1 ? parseFeeAmount(row[idxFee], price) : 0;
          if (fee === 0 && price > 0) {
            fee = Math.round(price * 0.1);
          }
          const shipping = idxShipping !== -1 ? parseAmount(row[idxShipping]) : 0;

          const net = Math.max(0, price - fee - shipping);

          monthlySalesData[ym].salesAmount += price;
          monthlySalesData[ym].netAmount += net;
          monthlySalesData[ym].count += 1;
        }
      }
    }
  }

  // --- 3. 出力データの構築 ---
  const allMonths = [...new Set([...ymList, ...Object.keys(monthlySalesData)])].sort();
  if (allMonths.length === 0) return;

  const outputHeaders = [
    "年月",
    "①月初在庫",
    "②当月仕入",
    "③取扱商品計",
    "④売れた原価",
    "⑤月末在庫",
    "⑥販売金額(売上)",
    "⑦手取り入金",
    "⑧回収差額(粗利)",
    "⑨資金化倍率",
    "⑩仕入回収率"
  ];

  const output = [outputHeaders];

  let totalPurchaseSum = 0;
  let totalCostSum = 0;
  let totalSalesSum = 0;
  let totalNetSum = 0;
  let totalGrossProfitSum = 0;

  allMonths.forEach(ym => {
    const stock = monthlyStockData[ym] || { startStock: 0, purchase: 0, cost: 0, endStock: 0 };
    const sales = monthlySalesData[ym] || { salesAmount: 0, netAmount: 0, count: 0 };

    const startStock = stock.startStock;
    const purchase = stock.purchase;
    const totalProduct = startStock + purchase;
    const cost = stock.cost;
    const endStock = stock.endStock > 0 ? stock.endStock : Math.max(0, totalProduct - cost);

    const salesAmt = sales.salesAmount;
    const netAmt = sales.netAmount > 0 ? sales.netAmount : salesAmt;
    const grossProfit = netAmt - cost;

    // 資金化倍率（入金 ÷ 売れた原価）
    const cashMultiplier = cost > 0 ? (netAmt / cost) : (netAmt > 0 ? 1 : 0);
    // 仕入回収率（入金 ÷ 当月仕入額）
    const purchaseRecoveryRate = purchase > 0 ? (netAmt / purchase) : (netAmt > 0 ? 1 : 0);

    totalPurchaseSum += purchase;
    totalCostSum += cost;
    totalSalesSum += salesAmt;
    totalNetSum += netAmt;
    totalGrossProfitSum += grossProfit;

    output.push([
      ym,
      startStock,
      purchase,
      totalProduct,
      cost,
      endStock,
      salesAmt,
      netAmt,
      grossProfit,
      cashMultiplier,
      purchaseRecoveryRate
    ]);
  });

  // 合計 / 平均行の追加
  const avgMultiplier = totalCostSum > 0 ? (totalNetSum / totalCostSum) : 0;
  const avgRecoveryRate = totalPurchaseSum > 0 ? (totalNetSum / totalPurchaseSum) : 0;

  output.push([
    "合計 / 平均",
    "-",
    totalPurchaseSum,
    "-",
    totalCostSum,
    monthlyStockData[allMonths[allMonths.length - 1]]?.endStock || 0,
    totalSalesSum,
    totalNetSum,
    totalGrossProfitSum,
    avgMultiplier,
    avgRecoveryRate
  ]);

  const numRows = output.length;
  const numCols = output[0].length;

  // シートに出力
  outputSheet.getRange(1, 1, numRows, numCols).setValues(output);

  // --- 書式とデザイン設定 ---
  outputSheet.getRange(1, 1, numRows, numCols).setVerticalAlignment("middle");

  // ヘッダー（1行目）
  const headerRange = outputSheet.getRange(1, 1, 1, numCols);
  headerRange
    .setBackground("#1E3A8A")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  outputSheet.setRowHeight(1, 32);

  // データ行
  if (numRows > 1) {
    // 年月（A列）: センタリング
    outputSheet.getRange(2, 1, numRows - 1, 1).setHorizontalAlignment("center");

    // 金額列（B列〜I列 / 2〜9列目）: 通貨フォーマット & 右揃え
    outputSheet.getRange(2, 2, numRows - 1, 8)
      .setNumberFormat('"¥"#,##0')
      .setHorizontalAlignment("right");

    // パーセント列（J列: 資金化倍率, K列: 仕入回収率）
    outputSheet.getRange(2, 10, numRows - 1, 2)
      .setNumberFormat('0.0%')
      .setHorizontalAlignment("right");

    // 行の高さ
    for (let r = 2; r <= numRows; r++) {
      outputSheet.setRowHeight(r, 26);
    }

    // 手取り入金（H列）と粗利（I列）を淡いグリーンで強調
    outputSheet.getRange(2, 8, numRows - 2, 2).setBackground("#F0FDF4");

    // 合計行（最下行）のデザイン
    const totalRowRange = outputSheet.getRange(numRows, 1, 1, numCols);
    totalRowRange
      .setFontWeight("bold")
      .setBackground("#F1F5F9");
    totalRowRange.getCell(1, 2).setHorizontalAlignment("center");
    totalRowRange.getCell(1, 4).setHorizontalAlignment("center");
  }

  // --- タイトルが確実に表示される列幅調整 ---
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

  outputSheet.autoResizeColumns(1, numCols);
  for (let col = 1; col <= numCols; col++) {
    const autoW = outputSheet.getColumnWidth(col);
    const headerTitle = output[0][col - 1];
    const minTitleW = calculateTitleWidth(headerTitle);
    outputSheet.setColumnWidth(col, Math.max(autoW + 8, minTitleW));
  }
}
