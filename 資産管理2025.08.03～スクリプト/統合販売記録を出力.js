/**
 * 「販売記録CSV照合表（確定CSV）」と「販売速報（フォーム回答）」から、
 * 重複排除ルール（Solution B）を適用した後の統合販売明細を抽出し、
 * 手数料・送料等を差し引いた「入金額(手取り)」、商品マスタからの「仕入額」、
 * そして「仕入控除後利益（粗利）」を含めて「統合_販売記録」シートに出力するスクリプト。
 */
function 統合販売記録を出力() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 対象シートの取得
  const matchSheet = ss.getSheetByName("販売記録CSV照合表");
  const salesSheet = ss.getSheetByName("販売速報（フォーム回答）");
  const masterSheet = ss.getSheetByName("商品マスタ_統合");
  const srcSheet = ss.getSheetByName("販売記録（CSV取り込み）");

  if (!salesSheet) {
    throw new Error("「販売速報（フォーム回答）」シートが見つかりません。");
  }

  // 出力シートの準備
  const outputSheetName = "統合_販売記録";
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

  // 手数料パース補助関数（パーセント文字列や小数にも対応）
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

  // --- 0. 販売記録（CSV取り込み）データのマップ化 (商品ID -> 入金額・手取り) ---
  const csvPriceMap = new Map();
  if (srcSheet) {
    const srcData = srcSheet.getDataRange().getValues();
    if (srcData.length > 1) {
      const srcHeaders = srcData[0].map(h => String(h).trim());
      const idxSrcId = findHeaderIndex(srcHeaders, ["商品ID", "商品id", "id", "ID", "商品ＩＤ", "注文番号"]);
      const idxPrice = findHeaderIndex(srcHeaders, ["価格", "商品価格", "販売価格", "売上金額", "金額"]);
      const idxFee = findHeaderIndex(srcHeaders, ["販売手数料金額", "販売手数料", "手数料額", "手数料"]);
      const idxShipping = findHeaderIndex(srcHeaders, ["配送料", "送料"]);
      const idxProfit = findHeaderIndex(srcHeaders, ["販売利益", "純利益", "入金額", "振込金額"]);

      for (let i = 1; i < srcData.length; i++) {
        const row = srcData[i];
        const id = idxSrcId !== -1 ? String(row[idxSrcId]).trim() : "";
        if (!id) continue;

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
          netProfit: netProfit
        });
      }
    }
  }

  // 商品マスタ情報のマップ化 (商品マスタNo -> { name, unitPrice })
  const masterInfoMap = new Map();
  if (masterSheet) {
    const masterData = masterSheet.getDataRange().getValues();
    if (masterData.length > 1) {
      const masterHeaders = masterData[0].map(h => String(h).trim());
      const idxMasterNo = findHeaderIndex(masterHeaders, ["商品マスタ No", "商品マスタNo", "No", "No."]);
      const idxMasterName = findHeaderIndex(masterHeaders, ["商品名", "品名"]);
      const idxMasterUnitPrice = findHeaderIndex(masterHeaders, ["実質仕入額(単品)", "仕入評価額", "仕入額", "仕入単価", "単価"]);

      if (idxMasterNo !== -1) {
        for (let i = 1; i < masterData.length; i++) {
          const no = String(masterData[i][idxMasterNo]).trim();
          const name = idxMasterName !== -1 ? String(masterData[i][idxMasterName]).trim() : "";
          const unitPrice = idxMasterUnitPrice !== -1 ? parseAmount(masterData[i][idxMasterUnitPrice]) : 0;
          if (no) {
            masterInfoMap.set(no, { name: name, unitPrice: unitPrice });
          }
        }
      }
    }
  }

  const integratedRows = [];
  let maxCsvDate = null; // CSV照合表内の最新購入日時

  // 使用済み商品マスタNoの記録セット (重複防止用)
  const usedMasterNoSet = new Set();

  // --- 1. 「販売記録CSV照合表」（確定メルカリ売上）の読み込み（最優先） ---
  if (matchSheet) {
    const matchData = matchSheet.getDataRange().getValues();
    if (matchData.length > 1) {
      const matchHeaders = matchData[0].map(h => String(h).trim());
      const idxMatchId = findHeaderIndex(matchHeaders, ["商品ID", "商品id", "id", "ID"]);
      const idxMatchNo = findHeaderIndex(matchHeaders, ["商品マスタ No", "商品マスタNo", "No"]);
      const idxMatchDate = findHeaderIndex(matchHeaders, ["購入日時", "販売日時", "日付"]);
      const idxMatchName = findHeaderIndex(matchHeaders, ["マスタ商品名", "メルカリ商品名", "商品名", "品名"]);
      const idxMatchBuyer = findHeaderIndex(matchHeaders, ["購入者", "バイヤー", "顧客名"]);

      if (idxMatchNo !== -1 && idxMatchDate !== -1) {
        for (let i = 1; i < matchData.length; i++) {
          const row = matchData[i];
          const id = idxMatchId !== -1 ? String(row[idxMatchId]).trim() : String(row[0]).trim();
          const no = String(row[idxMatchNo]).trim();
          const rawDate = row[idxMatchDate];
          const sDate = parseDate(rawDate);

          // 有効な商品マスタNoかつ日付が存在する場合
          if (no && isValidMasterNo(no) && sDate) {
            const masterInfo = masterInfoMap.get(no);
            let productName = idxMatchName !== -1 ? String(row[idxMatchName]).trim() : "";
            if (!productName || productName === "マスタ未登録") {
              productName = masterInfo ? masterInfo.name : "";
            }
            const buyer = idxMatchBuyer !== -1 ? String(row[idxMatchBuyer]).trim() : "";

            const csvInfo = csvPriceMap.get(id);
            const netProfit = csvInfo ? csvInfo.netProfit : 0;
            const unitPrice = masterInfo ? masterInfo.unitPrice : 0;
            const cost = unitPrice * 1; // 確定CSVは1個
            const profit = netProfit - cost; // 仕入控除後利益

            integratedRows.push({
              date: sDate,
              source: "確定CSV",
              masterNo: no,
              productName: productName,
              qty: 1,
              netAmount: netProfit,
              cost: cost,
              profit: profit,
              buyerInfo: buyer
            });

            // 採用済み商品マスタNoとして記録
            usedMasterNoSet.add(no);

            if (!maxCsvDate || sDate > maxCsvDate) {
              maxCsvDate = sDate;
            }
          }
        }
      }
    }
  }

  // --- 2. 「販売速報（フォーム回答）」の読み込み（確定CSVに未存在の商品マスタNoのみ補完） ---
  const salesData = salesSheet.getDataRange().getValues();
  if (salesData.length > 1) {
    const salesHeaders = salesData[0].map(h => String(h).trim());

    const idxSalesNo = findHeaderIndex(salesHeaders, ["商品マスタNo", "商品マスタ No", "No"]);
    const idxSalesDate = findHeaderIndex(salesHeaders, ["販売日", "日付"]);
    const idxSalesQty = findHeaderIndex(salesHeaders, ["販売数", "数量"]);
    const idxSalesPrice = findHeaderIndex(salesHeaders, ["販売価格", "売上", "価格", "金額"]);
    const idxSalesFee = findHeaderIndex(salesHeaders, ["販売手数料金額", "販売手数料額", "手数料金額", "手数料額", "販売手数料", "手数料"]);
    const idxSalesShipping = findHeaderIndex(salesHeaders, ["送料", "配送料", "送料額"]);
    const idxSalesBuyer = findHeaderIndex(salesHeaders, ["購入者情報（任意）", "購入者情報", "購入者", "バイヤー", "顧客名", "顧客", "備考", "理由", "用途", "チャネル"]);
    const idxSalesName = findHeaderIndex(salesHeaders, ["商品名", "品名"]);

    if (idxSalesNo !== -1 && idxSalesDate !== -1 && idxSalesQty !== -1) {
      for (let i = 1; i < salesData.length; i++) {
        const row = salesData[i];
        const no = String(row[idxSalesNo]).trim();

        // 商品マスタNoが無効、またはすでに確定CSV（または処理済みの速報）で採用済みの場合はスキップ
        if (!no || !isValidMasterNo(no) || usedMasterNoSet.has(no)) continue;

        const rawDate = row[idxSalesDate];
        const salesDate = parseDate(rawDate);
        const qty = parseNumber(row[idxSalesQty]);

        if (!salesDate) continue;

        const buyerInfo = idxSalesBuyer !== -1 ? String(row[idxSalesBuyer]).trim() : "";
        const isLatest = !maxCsvDate || salesDate > maxCsvDate;
        const sourceLabel = isLatest ? "販売速報(最新)" : "販売速報(補完)";

        const masterInfo = masterInfoMap.get(no);
        let productName = idxSalesName !== -1 ? String(row[idxSalesName]).trim() : "";
        if (!productName) {
          productName = masterInfo ? masterInfo.name : "";
        }

        const price = idxSalesPrice !== -1 ? parseAmount(row[idxSalesPrice]) : 0;
        
        // 手数料の確実な取得・計算（列から取得、なければ販売価格の10%）
        let fee = idxSalesFee !== -1 ? parseFeeAmount(row[idxSalesFee], price) : 0;
        if (fee === 0 && price > 0) {
          fee = Math.round(price * 0.1);
        }
        
        // 送料の取得
        const shipping = idxSalesShipping !== -1 ? parseAmount(row[idxSalesShipping]) : 0;

        // 入金額(手取り) ＝ 販売価格 － 販売手数料 － 送料
        const netProfit = Math.max(0, price - fee - shipping);

        const unitPrice = masterInfo ? masterInfo.unitPrice : 0;
        const cost = unitPrice * qty;
        const profit = netProfit - cost; // 仕入控除後利益

        integratedRows.push({
          date: salesDate,
          source: sourceLabel,
          masterNo: no,
          productName: productName,
          qty: qty,
          netAmount: netProfit,
          cost: cost,
          profit: profit,
          buyerInfo: buyerInfo
        });

        // 採用済み商品マスタNoとして記録（速報内での同No重複も防止）
        usedMasterNoSet.add(no);
      }
    }
  }

  // --- 3. 日付の降順（新しい順）でソート ---
  integratedRows.sort((a, b) => b.date.getTime() - a.date.getTime());

  // --- 4. シート出力データの作成 ---
  const headers = [
    "販売日時",
    "データソース",
    "商品マスタ No",
    "商品名",
    "販売数",
    "入金額(手取り)",
    "仕入額",
    "仕入控除後利益",
    "購入者 / 備考情報"
  ];
  const outputValues = [headers];

  integratedRows.forEach(item => {
    const dateStr = Utilities.formatDate(item.date, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
    outputValues.push([
      dateStr,
      item.source,
      item.masterNo,
      item.productName,
      item.qty,
      item.netAmount,
      item.cost,
      item.profit,
      item.buyerInfo
    ]);
  });

  // 書き込み
  outputSheet.getRange(1, 1, outputValues.length, headers.length).setValues(outputValues);

  // --- 5. デザイン・フォーマット設定 ---
  // ヘッダーデザイン
  outputSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#e6f2ff")
    .setHorizontalAlignment("center");

  outputSheet.setFrozenRows(1);

  if (outputValues.length > 1) {
    // 数量列（E列）の数値フォーマット
    outputSheet.getRange(2, 5, outputValues.length - 1, 1).setNumberFormat('#,##0').setHorizontalAlignment("right");
    // 金額列（F列: 入金額, G列: 仕入額, H列: 仕入控除後利益）の通貨フォーマット
    outputSheet.getRange(2, 6, outputValues.length - 1, 3).setNumberFormat('"¥"#,##0').setHorizontalAlignment("right");
    // 仕入控除後利益（H列）を薄い緑背景で強調
    outputSheet.getRange(2, 8, outputValues.length - 1, 1).setBackground("#F0FDF4");
  }

  // 列幅調整
  outputSheet.autoResizeColumns(1, headers.length);

  // 商品名（4列目/D列）が広くなりすぎるのを防ぐため、適正幅（250px）に設定
  outputSheet.setColumnWidth(4, 250);
  // 入金額(手取り)（6列目/F列）
  outputSheet.setColumnWidth(6, 115);
  // 仕入額（7列目/G列）
  outputSheet.setColumnWidth(7, 105);
  // 仕入控除後利益（8列目/H列）
  outputSheet.setColumnWidth(8, 125);
  // 購入者 / 備考情報（9列目/I列）も幅上限を200pxに調整
  if (outputSheet.getColumnWidth(9) > 220) {
    outputSheet.setColumnWidth(9, 200);
  }

  // 垂直方向中央揃え & 商品名・備考の折り返し設定
  outputSheet.getRange(1, 1, outputValues.length, headers.length).setVerticalAlignment("middle");
  if (outputValues.length > 1) {
    outputSheet.getRange(2, 4, outputValues.length - 1, 1).setWrap(true);
    outputSheet.getRange(2, 9, outputValues.length - 1, 1).setWrap(true);
  }

  const totalCount = outputValues.length - 1;
  SpreadsheetApp.getUi().alert(
    `重複排除後の統合販売記録 ${totalCount} 件を「${outputSheetName}」シートに出力しました。`
  );
}

/**
 * 候補リストから一致するヘッダーのインデックスを返す（改行・空白・記号を正規化して比較）
 */
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

/**
 * 日付のパース
 */
function parseDate(val) {
  const minValidDate = new Date("2024-01-01");
  if (val instanceof Date) {
    return val >= minValidDate ? val : null;
  }
  if (!val) return null;
  const d = new Date(val);
  if (isNaN(d.getTime())) return null;
  return d >= minValidDate ? d : null;
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
