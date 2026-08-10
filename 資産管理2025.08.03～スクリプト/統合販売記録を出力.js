/**
 * 「販売記録CSV照合表（確定CSV）」と「販売速報（フォーム回答）」から、
 * 重複排除ルール（Solution B）を適用した後の統合販売明細を抽出し、
 * 「統合_販売記録」シートに出力するスクリプト。
 */
function 統合販売記録を出力() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 対象シートの取得
  const matchSheet = ss.getSheetByName("販売記録CSV照合表");
  const salesSheet = ss.getSheetByName("販売速報（フォーム回答）");
  const masterSheet = ss.getSheetByName("商品マスタ_統合");

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

  // 商品マスタ名情報のマップ化 (商品マスタNo -> 商品名)
  const masterNameMap = new Map();
  if (masterSheet) {
    const masterData = masterSheet.getDataRange().getValues();
    if (masterData.length > 1) {
      const masterHeaders = masterData[0].map(h => String(h).trim());
      const idxMasterNo = findHeaderIndex(masterHeaders, ["商品マスタ No", "商品マスタNo", "No", "No."]);
      const idxMasterName = findHeaderIndex(masterHeaders, ["商品名", "品名"]);

      if (idxMasterNo !== -1 && idxMasterName !== -1) {
        for (let i = 1; i < masterData.length; i++) {
          const no = String(masterData[i][idxMasterNo]).trim();
          const name = String(masterData[i][idxMasterName]).trim();
          if (no && name) {
            masterNameMap.set(no, name);
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
      const idxMatchNo = findHeaderIndex(matchHeaders, ["商品マスタ No", "商品マスタNo", "No"]);
      const idxMatchDate = findHeaderIndex(matchHeaders, ["購入日時", "販売日時", "日付"]);
      const idxMatchName = findHeaderIndex(matchHeaders, ["マスタ商品名", "メルカリ商品名", "商品名", "品名"]);
      const idxMatchBuyer = findHeaderIndex(matchHeaders, ["購入者", "バイヤー", "顧客名"]);

      if (idxMatchNo !== -1 && idxMatchDate !== -1) {
        for (let i = 1; i < matchData.length; i++) {
          const row = matchData[i];
          const no = String(row[idxMatchNo]).trim();
          const rawDate = row[idxMatchDate];
          const sDate = parseDate(rawDate);

          // 有効な商品マスタNoかつ日付が存在する場合
          if (no && isValidMasterNo(no) && sDate) {
            let productName = idxMatchName !== -1 ? String(row[idxMatchName]).trim() : "";
            if (!productName || productName === "マスタ未登録") {
              productName = masterNameMap.get(no) || "";
            }
            const buyer = idxMatchBuyer !== -1 ? String(row[idxMatchBuyer]).trim() : "";

            integratedRows.push({
              date: sDate,
              source: "確定CSV",
              masterNo: no,
              productName: productName,
              qty: 1,
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

        let productName = idxSalesName !== -1 ? String(row[idxSalesName]).trim() : "";
        if (!productName) {
          productName = masterNameMap.get(no) || "";
        }

        integratedRows.push({
          date: salesDate,
          source: sourceLabel,
          masterNo: no,
          productName: productName,
          qty: qty,
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
  const headers = ["販売日時", "データソース", "商品マスタ No", "商品名", "販売数", "購入者 / 備考情報"];
  const outputValues = [headers];

  integratedRows.forEach(item => {
    const dateStr = Utilities.formatDate(item.date, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
    outputValues.push([
      dateStr,
      item.source,
      item.masterNo,
      item.productName,
      item.qty,
      item.buyerInfo
    ]);
  });

  // 書き込み
  outputSheet.getRange(1, 1, outputValues.length, headers.length).setValues(outputValues);

  // --- 5. デザイン・フォーマット設定 ---
  // ヘッダーデザイン
  outputSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#e6f2ff");

  outputSheet.setFrozenRows(1);

  if (outputValues.length > 1) {
    // 数量列（E列）の数値フォーマット
    outputSheet.getRange(2, 5, outputValues.length - 1, 1).setNumberFormat('#,##0');
  }

  // 列幅調整
  outputSheet.autoResizeColumns(1, headers.length);

  const totalCount = outputValues.length - 1;
  SpreadsheetApp.getUi().alert(
    `重複排除後の統合販売記録 ${totalCount} 件を「${outputSheetName}」シートに出力しました。`
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
