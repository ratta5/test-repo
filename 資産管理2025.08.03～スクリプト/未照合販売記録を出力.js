/**
 * 「販売記録CSV照合表」シートの中で、
 * 商品マスタ No（C列）が「空白」「該当なし」「マスタ未登録」または「数値が入っていない」データを抽出し、
 * 「未照合_販売記録」シートに出力するスクリプト。
 * 
 * 出力項目: 購入日時, 商品ID, 商品名, 商品URL, 購入者
 */
function 未照合販売記録を出力() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const destSheet = ss.getSheetByName("販売記録CSV照合表");
  const srcSheet = ss.getSheetByName("販売記録（CSV取り込み）");

  if (!destSheet) {
    throw new Error("「販売記録CSV照合表」シートが見つかりません。");
  }

  // 出力先シート「未照合_販売記録」の準備
  const outputSheetName = "未照合_販売記録";
  let outputSheet = ss.getSheetByName(outputSheetName);
  if (outputSheet) {
    outputSheet.clearContents();
  } else {
    outputSheet = ss.insertSheet(outputSheetName);
  }

  // 「販売記録（CSV取り込み）」データをマップ化 (商品URL等の補完用)
  const srcMap = new Map(); // id -> { date, url, name, buyer }
  let idxSrcId = -1, idxSrcDate = -1, idxSrcName = -1, idxSrcUrl = -1, idxSrcBuyer = -1;

  if (srcSheet) {
    const srcData = srcSheet.getDataRange().getValues();
    if (srcData.length > 0) {
      const srcHeaders = srcData[0].map(h => String(h).trim());
      idxSrcId = findHeaderIndex(srcHeaders, ["商品ID", "商品id", "id", "ID", "商品ＩＤ", "注文番号"]);
      idxSrcDate = findHeaderIndex(srcHeaders, ["購入日時", "販売日時", "取引日時", "日時", "取引完了日時", "日付"]);
      idxSrcName = findHeaderIndex(srcHeaders, ["商品名", "品名", "タイトル", "商品タイトル"]);
      idxSrcUrl = findHeaderIndex(srcHeaders, ["商品URL", "商品url", "URL", "url", "リンク", "商品ページURL", "商品ページ"]);
      idxSrcBuyer = findHeaderIndex(srcHeaders, ["購入者", "バイヤー", "顧客名", "顧客", "購入者名"]);

      if (idxSrcId !== -1) {
        for (let i = 1; i < srcData.length; i++) {
          const row = srcData[i];
          const id = String(row[idxSrcId]).trim();
          if (id) {
            const dateVal = idxSrcDate !== -1 ? row[idxSrcDate] : "";
            let dateStr = "";
            if (dateVal instanceof Date) {
              dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
            } else {
              dateStr = String(dateVal).trim();
            }
            const name = idxSrcName !== -1 ? String(row[idxSrcName]).trim() : "";
            const url = idxSrcUrl !== -1 ? String(row[idxSrcUrl]).trim() : "";
            const buyer = idxSrcBuyer !== -1 ? String(row[idxSrcBuyer]).trim() : "";

            srcMap.set(id, { date: dateStr, name: name, url: url, buyer: buyer });
          }
        }
      }
    }
  }

  // 「販売記録CSV照合表」から未照合（C列が空白、該当なし、数値なし）の行を抽出
  const destData = destSheet.getDataRange().getValues();
  if (destData.length <= 1) {
    SpreadsheetApp.getUi().alert("「販売記録CSV照合表」にデータが存在しません。");
    return;
  }

  const outputHeaders = ["購入日時", "商品ID", "商品名", "商品URL", "購入者"];
  const outputRows = [outputHeaders];

  // A列: 商品ID (0), B列: メルカリ商品名 (1), C列: 商品マスタ No (2), E列: 購入者 (4), F列: 購入日時 (5)
  for (let i = 1; i < destData.length; i++) {
    const row = destData[i];
    const id = String(row[0]).trim();
    const masterNo = row.length > 2 ? row[2] : "";

    if (id && !isValidMasterNo(masterNo)) {
      // C列が有効な商品マスタNo（数値含む）でない場合（空白、該当なし、数値なし等）
      const srcInfo = srcMap.get(id);

      const dateStr = srcInfo ? srcInfo.date : (row.length > 5 ? String(row[5]).trim() : "");
      const name = srcInfo ? srcInfo.name : (row.length > 1 ? String(row[1]).trim() : "");
      const url = srcInfo ? srcInfo.url : "";
      const buyer = srcInfo ? srcInfo.buyer : (row.length > 4 ? String(row[4]).trim() : "");

      outputRows.push([dateStr, id, name, url, buyer]);
    }
  }

  // シートへ書き込み・装飾
  outputSheet.getRange(1, 1, outputRows.length, outputRows[0].length).setValues(outputRows);

  // ヘッダー行のデザイン
  outputSheet.getRange(1, 1, 1, outputRows[0].length)
    .setFontWeight("bold")
    .setBackground("#f3f3f3");
  outputSheet.setFrozenRows(1);

  // 列幅の自動調整
  outputSheet.autoResizeColumns(1, outputRows[0].length);

  const unmappedCount = outputRows.length - 1;
  if (unmappedCount > 0) {
    SpreadsheetApp.getUi().alert(`未照合（商品マスタNoが空白・該当なし・数値なし）の販売記録 ${unmappedCount} 件を「${outputSheetName}」シートに出力しました。`);
  } else {
    SpreadsheetApp.getUi().alert(`すべての販売記録の照合（商品マスタNoに数値入力）が完了しています。未照合データはありません。`);
  }
}

/**
 * 商品マスタ No が有効な数値/IDであるか判定
 */
function isValidMasterNo(val) {
  if (val === null || val === undefined) return false;
  const str = String(val).trim();
  if (str === "" || str === "-" || str === "0") return false;
  if (str.includes("該当なし") || str.includes("未登録") || str.includes("不明")) return false;
  
  // 数値（数字 0-9）が含まれているか確認
  return /\d/.test(str);
}

/**
 * 候補リストから一致するヘッダーのインデックスを返す
 */
function findHeaderIndex(headers, candidates) {
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i].toLowerCase();
    for (let j = 0; j < candidates.length; j++) {
      if (h === candidates[j].toLowerCase() || h.includes(candidates[j].toLowerCase())) {
        return i;
      }
    }
  }
  return -1;
}
