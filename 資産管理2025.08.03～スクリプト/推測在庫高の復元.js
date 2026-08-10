/**
 * 過去の推測在庫高を復元・推定するスクリプト
 * 
 * 最新の実地棚卸データを基準点とし、仕入（商品マスタ）と販売（販売速報）のデータから
 * 過去の各月末時点における「推測在庫数」を逆算し、仕入単価を掛けて「推測在庫高」を算出します。
 */
function 推測在庫高の復元() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 対象シートの取得
  const masterSheet = ss.getSheetByName("商品マスタ_統合");
  const inventorySheet = ss.getSheetByName("棚卸_取り込み");
  const salesSheet = ss.getSheetByName("販売速報（フォーム回答）");
  const matchSheet = ss.getSheetByName("販売記録CSV照合表");
  const outputSheetName = "推測在庫高_推移";

  if (!masterSheet) {
    throw new Error("「商品マスタ_統合」シートが見つかりません。");
  }
  if (!inventorySheet) {
    throw new Error("「棚卸_取り込み」シートが見つかりません。");
  }
  if (!salesSheet) {
    throw new Error("「販売速報（フォーム回答）」シートが見つかりません。");
  }

  // 出力シートの準備
  let outputSheet = ss.getSheetByName(outputSheetName);
  if (outputSheet) {
    outputSheet.clearContents();
  } else {
    outputSheet = ss.insertSheet(outputSheetName);
  }

  // --- 1. 商品マスタ_統合の読み込み ---
  const masterData = masterSheet.getDataRange().getValues();
  const masterHeaders = masterData[0].map(h => String(h).trim());
  
  const idxMasterNo = findHeaderIndex(masterHeaders, ["商品マスタ No", "商品マスタNo", "No", "No."]);
  const idxPurchaseDate = findHeaderIndex(masterHeaders, ["仕入日", "購入日"]);
  const idxUnitPrice = findHeaderIndex(masterHeaders, ["実質仕入額(単品)", "仕入評価額", "単価"]);
  const idxPurchaseQty = findHeaderIndex(masterHeaders, ["購入数", "数量"]);
  const idxProductName = findHeaderIndex(masterHeaders, ["商品名", "品名"]);
  const idxPurchasePurpose = findHeaderIndex(masterHeaders, ["購入目的", "目的"]);

  if (idxMasterNo === -1 || idxPurchaseDate === -1 || idxUnitPrice === -1 || idxPurchaseQty === -1) {
    throw new Error("「商品マスタ_統合」に必要な列が見つかりません。必須列: 商品マスタ No, 仕入日, 実質仕入額(単品), 購入数");
  }

  const productsMap = {}; // 商品マスタNo -> 商品情報
  let minDate = new Date(); // データ全体の最古日付（逆算の開始月判定用）

  for (let i = 1; i < masterData.length; i++) {
    const row = masterData[i];
    const no = String(row[idxMasterNo]).trim();
    if (!no) continue;

    const rawDate = row[idxPurchaseDate];
    const purchaseDate = parseDate(rawDate);
    
    // 仕入日が未記載のものは非表示（計算および出力から除外）にする
    if (!purchaseDate) continue;

    // 購入目的が物販でないものは非表示にする
    if (idxPurchasePurpose !== -1) {
      const purpose = String(row[idxPurchasePurpose]).trim();
      if (!purpose.includes("1") && !purpose.includes("物販")) continue;
    }

    const purchaseQty = parseNumber(row[idxPurchaseQty]);
    // 購入数がない（0以下）ものも非表示にする
    if (purchaseQty <= 0) continue;

    const rawPurpose = idxPurchasePurpose !== -1 ? String(row[idxPurchasePurpose]).trim() : "";
    const unitPrice = parseAmount(row[idxUnitPrice]);
    const name = idxProductName !== -1 ? String(row[idxProductName]).trim() : "";

    productsMap[no] = {
      no: no,
      purchaseDate: purchaseDate,
      unitPrice: unitPrice,
      purchaseQty: purchaseQty,
      purchasePurpose: rawPurpose,
      name: name,
      monthlyStock: {}
    };

    if (purchaseDate < minDate) {
      minDate = new Date(purchaseDate);
    }
  }

  // --- 2. 棚卸_取り込みの読み込み ---
  const inventoryData = inventorySheet.getDataRange().getValues();
  const inventoryHeaders = inventoryData[0].map(h => String(h).trim());

  const idxInvNo = findHeaderIndex(inventoryHeaders, ["商品マスタ No", "商品マスタNo"]);
  const idxInvDate = findHeaderIndex(inventoryHeaders, ["棚卸日", "日付"]);
  const idxInvQty = findHeaderIndex(inventoryHeaders, ["実在庫", "在庫数", "数量"]);

  if (idxInvNo === -1 || idxInvDate === -1 || idxInvQty === -1) {
    throw new Error("「棚卸_取り込み」に必要な列が見つかりません。必須列: 商品マスタ No, 棚卸日, 実在庫");
  }

  const inventoryMap = {}; // 商品マスタNo -> 最新の棚卸情報

  for (let i = 1; i < inventoryData.length; i++) {
    const row = inventoryData[i];
    const no = String(row[idxInvNo]).trim();
    if (!no) continue;

    const rawDate = row[idxInvDate];
    const invDate = parseDate(rawDate);
    const qty = parseNumber(row[idxInvQty]);

    if (!invDate) continue;

    // 同一商品で複数の棚卸データがある場合は、最新のものを採用
    if (!inventoryMap[no] || invDate > inventoryMap[no].date) {
      inventoryMap[no] = {
        date: invDate,
        qty: qty
      };
    }
  }

  // --- 3. 販売データの統合（Solution B: 確定CSV照合表 ＋ フォーム回答） ---
  const salesMap = {}; // 商品マスタNo -> 販売履歴リスト
  let maxCsvDate = null; // CSV照合表内の最新購入日時

  // 3-A. 「販売記録CSV照合表」（確定メルカリ売上）の読み込み（最優先）
  const usedMasterNoSet = new Set();

  if (matchSheet) {
    const matchData = matchSheet.getDataRange().getValues();
    if (matchData.length > 1) {
      const matchHeaders = matchData[0].map(h => String(h).trim());
      const idxMatchNo = findHeaderIndex(matchHeaders, ["商品マスタ No", "商品マスタNo", "No"]);
      const idxMatchDate = findHeaderIndex(matchHeaders, ["購入日時", "販売日時", "日付"]);

      if (idxMatchNo !== -1 && idxMatchDate !== -1) {
        for (let i = 1; i < matchData.length; i++) {
          const row = matchData[i];
          const no = String(row[idxMatchNo]).trim();
          const rawDate = row[idxMatchDate];
          const sDate = parseDate(rawDate);

          // 有効な商品マスタNoかつ日付が存在する場合
          if (no && isValidMasterNo(no) && sDate) {
            if (!salesMap[no]) {
              salesMap[no] = [];
            }
            salesMap[no].push({
              date: sDate,
              qty: 1 // CSV照合表は1取引1個
            });

            usedMasterNoSet.add(no);

            if (!maxCsvDate || sDate > maxCsvDate) {
              maxCsvDate = sDate;
            }
            if (sDate < minDate) {
              minDate = new Date(sDate);
            }
          }
        }
      }
    }
  }

  // 3-B. 「販売速報（フォーム回答）」の読み込み（確定CSVに未存在の商品マスタNoのみ補完）
  const salesData = salesSheet.getDataRange().getValues();
  const salesHeaders = salesData[0].map(h => String(h).trim());

  const idxSalesNo = findHeaderIndex(salesHeaders, ["商品マスタNo", "商品マスタ No", "No"]);
  const idxSalesDate = findHeaderIndex(salesHeaders, ["販売日", "日付"]);
  const idxSalesQty = findHeaderIndex(salesHeaders, ["販売数", "数量"]);

  if (idxSalesNo === -1 || idxSalesDate === -1 || idxSalesQty === -1) {
    throw new Error("「販売速報（フォーム回答）」に必要な列が見つかりません。必須列: 商品マスタNo, 販売日, 販売数");
  }

  for (let i = 1; i < salesData.length; i++) {
    const row = salesData[i];
    const no = String(row[idxSalesNo]).trim();

    // 商品マスタNoが無効、またはすでに確定CSV（または処理済みの速報）で採用済みの場合はスキップ
    if (!no || !isValidMasterNo(no) || usedMasterNoSet.has(no)) continue;

    const rawDate = row[idxSalesDate];
    const salesDate = parseDate(rawDate);
    const qty = parseNumber(row[idxSalesQty]);

    if (!salesDate) continue;

    if (!salesMap[no]) {
      salesMap[no] = [];
    }
    salesMap[no].push({
      date: salesDate,
      qty: qty
    });

    usedMasterNoSet.add(no);

    if (salesDate < minDate) {
      minDate = new Date(salesDate);
    }
  }

  // --- 4. 期間の設定 ---
  // 開始月はデータ上の最古の日付（仕入または販売）の月初めとする
  const startYear = minDate.getFullYear();
  const startMonth = minDate.getMonth(); // 0-11
  
  const today = new Date();
  const endYear = today.getFullYear();
  const endMonth = today.getMonth();

  // 年月のリスト（yyyy-MM）を作成
  const months = [];
  let currY = startYear;
  let currM = startMonth;

  while (currY < endYear || (currY === endYear && currM <= endMonth)) {
    months.push({
      year: currY,
      month: currM,
      label: currY + "-" + String(currM + 1).padStart(2, "0")
    });
    currM++;
    if (currM > 11) {
      currM = 0;
      currY++;
    }
  }

  // --- 5. 月別の逆算・集計処理 ---
  const resultRows = [];
  
  months.forEach(m => {
    // 該当月の月末時点の日時を取得（翌月の0日 ＝ 当月の末日）
    const monthEndDate = new Date(m.year, m.month + 1, 0, 23, 59, 59, 999);
    const monthStartDate = new Date(m.year, m.month, 1, 0, 0, 0, 0);

    let totalPurchaseAmount = 0; // 当月の仕入総額
    let totalSalesCost = 0;      // 当月の販売原価

    // 商品ごとに在庫数を算出して集計
    let totalInventoryValue = 0;

    Object.keys(productsMap).forEach(no => {
      const prod = productsMap[no];
      const inv = inventoryMap[no];
      const sales = salesMap[no] || [];

      // 1. 当月仕入額の加算
      if (prod.purchaseDate && prod.purchaseDate >= monthStartDate && prod.purchaseDate <= monthEndDate) {
        totalPurchaseAmount += prod.purchaseQty * prod.unitPrice;
      }

      // 2. 当月販売原価の加算
      sales.forEach(sale => {
        if (sale.date >= monthStartDate && sale.date <= monthEndDate) {
          totalSalesCost += sale.qty * prod.unitPrice;
        }
      });

      // 3. 月末時点の在庫数を計算
      let stockQty = 0;

      if (inv) {
        // 棚卸データがある場合
        const invDate = inv.date;
        const invQty = inv.qty;

        if (monthEndDate.getTime() < invDate.getTime()) {
          // 月末が棚卸日より「過去」の場合: 過去へ逆算
          // 在庫 = 棚卸実在庫 - (月末から棚卸日までの仕入) + (月末から棚卸日までの販売)
          let periodPurchase = 0;
          if (prod.purchaseDate && prod.purchaseDate > monthEndDate && prod.purchaseDate <= invDate) {
            periodPurchase = prod.purchaseQty;
          }

          let periodSales = 0;
          sales.forEach(sale => {
            if (sale.date > monthEndDate && sale.date <= invDate) {
              periodSales += sale.qty;
            }
          });

          stockQty = invQty - periodPurchase + periodSales;
        } else {
          // 月末が棚卸日より「未来または同時点」の場合: 未来へ順算
          // 在庫 = 棚卸実在庫 + (棚卸日から月末までの仕入) - (棚卸日から月末までの販売)
          let periodPurchase = 0;
          if (prod.purchaseDate && prod.purchaseDate > invDate && prod.purchaseDate <= monthEndDate) {
            periodPurchase = prod.purchaseQty;
          }

          let periodSales = 0;
          sales.forEach(sale => {
            if (sale.date > invDate && sale.date <= monthEndDate) {
              periodSales += sale.qty;
            }
          });

          stockQty = invQty + periodPurchase - periodSales;
        }
      } else {
        // 棚卸データがない場合: 仕入日と販売履歴から単純計算
        // 在庫 = (月末までに仕入れた数) - (月末までに販売した数)
        let periodPurchase = 0;
        if (prod.purchaseDate && prod.purchaseDate <= monthEndDate) {
          periodPurchase = prod.purchaseQty;
        }

        let periodSales = 0;
        sales.forEach(sale => {
          if (sale.date <= monthEndDate) {
            periodSales += sale.qty;
          }
        });

        stockQty = periodPurchase - periodSales;
      }

      // マイナス在庫はデータ不整合として0に補正
      stockQty = Math.max(0, stockQty);

      // 月別の在庫数を記録
      prod.monthlyStock[m.label] = stockQty;

      // 在庫金額の算出
      totalInventoryValue += stockQty * prod.unitPrice;
    });

    resultRows.push([
      m.label,
      totalPurchaseAmount,
      totalSalesCost,
      totalInventoryValue
    ]);
  });

  // --- 6. 結果 of シート出力 ---
  const outputData = [
    ["年月", "当月仕入額", "当月販売原価", "月末推測在庫高"]
  ];
  resultRows.forEach(row => outputData.push(row));

  outputSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);

  // 書式の設定
  outputSheet.getRange(1, 1, 1, outputData[0].length).setFontWeight("bold").setBackground("#f3f3f3");
  if (outputData.length > 1) {
    outputSheet.getRange(2, 2, outputData.length - 1, 3).setNumberFormat('"¥"#,##0');
  }
  outputSheet.autoResizeColumns(1, outputData[0].length);

  // --- 7. 商品別在庫数推移のシート出力 ---
  const productOutputSheetName = "推測在庫_商品別推移";
  let prodSheet = ss.getSheetByName(productOutputSheetName);
  if (prodSheet) {
    prodSheet.clearContents();
  } else {
    prodSheet = ss.insertSheet(productOutputSheetName);
  }

  // ヘッダー作成
  const prodHeaders = ["商品マスタ No", "商品名", "仕入日", "実質仕入額(単品)", "購入数", "購入目的"];
  months.forEach(m => {
    prodHeaders.push(m.label + " (数)");
    prodHeaders.push(m.label + " (金)");
  });

  const prodOutputData = [prodHeaders];

  // 各月の縦合計を計算するためのオブジェクトを初期化
  const monthlySums = {};
  months.forEach(m => {
    monthlySums[m.label + " (数)"] = 0;
    monthlySums[m.label + " (金)"] = 0;
  });
  let totalPurchaseQtySum = 0;

  // 商品データ行を一時的に格納する配列
  const tempProductRows = [];

  // 商品ごとに1行作成
  Object.keys(productsMap).sort().forEach(no => {
    const prod = productsMap[no];
    const row = [
      prod.no,
      prod.name,
      prod.purchaseDate ? Utilities.formatDate(prod.purchaseDate, Session.getScriptTimeZone(), "yyyy/MM/dd") : "",
      prod.unitPrice,
      prod.purchaseQty,
      prod.purchasePurpose
    ];
    totalPurchaseQtySum += prod.purchaseQty;

    // 各月の在庫数と在庫金額を追記
    months.forEach(m => {
      const qty = prod.monthlyStock[m.label] ?? 0;
      const amt = qty * prod.unitPrice;
      row.push(qty);
      row.push(amt);

      // 縦合計に加算
      monthlySums[m.label + " (数)"] += qty;
      monthlySums[m.label + " (金)"] += amt;
    });
    tempProductRows.push(row);
  });

  // 合計行を作成して追加（ヘッダー直後の2行目にする）
  const totalRow = [
    "合計",
    "",
    "",
    "",
    totalPurchaseQtySum,
    ""
  ];
  months.forEach(m => {
    totalRow.push(monthlySums[m.label + " (数)"]);
    totalRow.push(monthlySums[m.label + " (金)"]);
  });
  prodOutputData.push(totalRow);

  // 商品明細行を追加（3行目以降）
  tempProductRows.forEach(row => {
    prodOutputData.push(row);
  });

  prodSheet.getRange(1, 1, prodOutputData.length, prodOutputData[0].length).setValues(prodOutputData);

  // 書式の設定
  // ヘッダー行
  prodSheet.getRange(1, 1, 1, prodOutputData[0].length).setFontWeight("bold").setBackground("#e6f7ff");
  
  if (prodOutputData.length > 2) {
    // 単価（4列目）
    prodSheet.getRange(2, 4, prodOutputData.length - 1, 1).setNumberFormat('"¥"#,##0'); 
    // 購入数（5列目）
    prodSheet.getRange(2, 5, prodOutputData.length - 1, 1).setNumberFormat('#,##0'); 

    // 各月の在庫数・金額の列フォーマットを適用 (7列目から開始)
    const numCols = prodOutputData[0].length;
    for (let col = 7; col <= numCols; col++) {
      const formatRange = prodSheet.getRange(2, col, prodOutputData.length - 1, 1);
      if (col % 2 === 1) {
        // 奇数列 ＝ 個数
        formatRange.setNumberFormat('#,##0');
      } else {
        // 偶数列 ＝ 金額
        formatRange.setNumberFormat('"¥"#,##0');
      }
    }
  }

  // 合計行（2行目）のデザイン
  prodSheet.getRange(2, 1, 1, prodOutputData[0].length)
    .setFontWeight("bold")
    .setBackground("#f2f2f2");
  
  // 列幅の調整
  prodSheet.autoResizeColumns(1, prodOutputData[0].length);
  // 商品名（2列目/B列）の幅が広くなりすぎるのを防ぐため、固定幅にして折り返しを有効にする
  prodSheet.setColumnWidth(2, 250);
  if (prodOutputData.length > 1) {
    prodSheet.getRange(2, 2, prodOutputData.length - 1, 1).setWrap(true);
  }

  SpreadsheetApp.getUi().alert(
    "推測在庫高の復元が完了しました。\n" +
    "「" + outputSheetName + "」および「" + productOutputSheetName + "」シートに出力されました。"
  );
}

/**
 * 候補リストから一致するヘッダーのインデックスを返す
 */
function findHeaderIndex(headers, candidates) {
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i].toLowerCase();
    for (let j = 0; j < candidates.length; j++) {
      const cand = candidates[j].toLowerCase();
      if (h === cand || h.includes(cand)) {
        return i;
      }
    }
  }
  return -1;
}

/**
 * 日付のパース
 */
function parseDate(val) {
  const minValidDate = new Date("2024-01-01"); // 2024年より古い日付は無効（空行等による誤作動防止）
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
function parseAmount(val) {
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
function parseNumber(val) {
  if (typeof val === "number") return val;
  if (typeof val === "string") {
    const num = parseInt(val.replace(/[,]/g, ""), 10);
    return isNaN(num) ? 0 : num;
  }
  return 0;
}

/**
 * 商品マスタ No が有効な数値/IDであるか判定
 */
function isValidMasterNo(val) {
  if (val === null || val === undefined) return false;
  const str = String(val).trim();
  if (str === "" || str === "-" || str === "0") return false;
  if (str.includes("該当なし") || str.includes("未登録") || str.includes("不明")) return false;
  return /\d/.test(str);
}

