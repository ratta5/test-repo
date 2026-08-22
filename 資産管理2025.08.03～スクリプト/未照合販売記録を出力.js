/**
 * 「販売記録CSV照合表」シートの中で、
 * 商品マスタ No（C列）が「空白」「該当なし」「マスタ未登録」または「数値が入っていない」データを抽出し、
 * 「未照合_販売記録」シートに出力するスクリプト。
 * 
 * 出力項目: 購入日時, 商品ID, 商品名, 商品URL, 購入者, 入力用_商品マスタNo
 */
function 未照合販売記録を出力() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const destSheet = ss.getSheetByName("販売記録CSV照合表");
  const srcSheet = ss.getSheetByName("販売記録（CSV取り込み）");

  if (!destSheet) {
    throw new Error("「販売記録CSV照合表」シートが見つかりません。");
  }

  const outputSheetName = "未照合_販売記録";
  let outputSheet = ss.getSheetByName(outputSheetName);

  // 再出力前に既存の「入力用_商品マスタNo（F列）」のメモ入力を一時保存して保持する
  const existingInputMap = new Map(); // id -> inputMemo
  if (outputSheet && outputSheet.getLastRow() > 1) {
    const existingData = outputSheet.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      const id = String(existingData[i][1]).trim(); // B列: 商品ID
      const memo = existingData[i].length > 5 ? String(existingData[i][5]).trim() : ""; // F列: メモ
      if (id && memo) {
        existingInputMap.set(id, memo);
      }
    }
  }

  // シート準備
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

  // 「販売記録CSV照合表」から未照合（商品マスタ No に数値が入っていない）の行を抽出
  const destData = destSheet.getDataRange().getValues();
  if (destData.length <= 1) {
    SpreadsheetApp.getUi().alert("「販売記録CSV照合表」にデータが存在しません。");
    return;
  }

  const destHeaders = destData[0].map(h => String(h).trim());
  const idxDestId = findHeaderIndex(destHeaders, ["商品ID", "商品id", "id", "ID", "商品ＩＤ"]);
  const idxDestName = findHeaderIndex(destHeaders, ["メルカリ商品名", "商品名", "品名", "タイトル"]);
  const idxDestMasterNo = findHeaderIndex(destHeaders, ["商品マスタ No", "商品マスタNo", "No", "No."]);
  const idxDestBuyer = findHeaderIndex(destHeaders, ["購入者", "バイヤー", "顧客名"]);
  const idxDestDate = findHeaderIndex(destHeaders, ["購入日時", "販売日時", "取引日時", "日時"]);

  const outputHeaders = ["購入日時", "商品ID", "商品名", "商品URL", "購入者", "入力用_商品マスタNo"];
  const outputRows = [outputHeaders];

  for (let i = 1; i < destData.length; i++) {
    const row = destData[i];
    const id = idxDestId !== -1 ? String(row[idxDestId]).trim() : String(row[0]).trim();
    const masterNo = idxDestMasterNo !== -1 ? row[idxDestMasterNo] : (row.length > 2 ? row[2] : "");

    if (id && !isValidMasterNo(masterNo)) {
      const srcInfo = srcMap.get(id);

      const dateStr = srcInfo ? srcInfo.date : (idxDestDate !== -1 ? String(row[idxDestDate]).trim() : (row.length > 5 ? String(row[5]).trim() : ""));
      const name = srcInfo ? srcInfo.name : (idxDestName !== -1 ? String(row[idxDestName]).trim() : (row.length > 1 ? String(row[1]).trim() : ""));
      const rawUrl = srcInfo ? srcInfo.url : "";
      const buyer = srcInfo ? srcInfo.buyer : (idxDestBuyer !== -1 ? String(row[idxDestBuyer]).trim() : (row.length > 4 ? String(row[4]).trim() : ""));

      // HYPERLINK数式の作成（URLが存在する場合）
      const urlFormula = rawUrl && rawUrl.startsWith("http")
        ? `=HYPERLINK("${rawUrl.replace(/"/g, '""')}", "メルカリで開く ↗")`
        : rawUrl;

      // 過去入力途中のメモがあれば保持
      const savedMemo = existingInputMap.get(id) || "";

      outputRows.push([dateStr, id, name, urlFormula, buyer, savedMemo]);
    }
  }

  // シートへ書き込み
  outputSheet.getRange(1, 1, outputRows.length, outputRows[0].length).setValues(outputRows);

  // ヘッダー行のデザイン
  outputSheet.getRange(1, 1, 1, outputRows[0].length)
    .setFontWeight("bold")
    .setBackground("#f3f3f3");
  outputSheet.setFrozenRows(1);

  // 入力用_商品マスタNo列（F列）の背景色を入力しやすい薄黄色に装飾
  if (outputRows.length > 1) {
    outputSheet.getRange(2, 6, outputRows.length - 1, 1)
      .setBackground("#fffbe6")
      .setFontWeight("bold");
  }

  // 列幅の自動調整
  outputSheet.autoResizeColumns(1, outputRows[0].length);

  const unmappedCount = outputRows.length - 1;
  if (unmappedCount > 0) {
    SpreadsheetApp.getUi().alert(
      `未照合の販売記録 ${unmappedCount} 件を「${outputSheetName}」シートに出力しました。\n\n` +
      `【使い方】\n` +
      `1. D列「商品URL」のリンクからメルカリ商品ページを開きます。\n` +
      `2. 確認した商品コードを F列「入力用_商品マスタNo」に入力します。\n` +
      `3. 上部メニュー「資産管理メニュー ＞ 2. メモしたマスタNoを照合表へ転記」を実行してください。`
    );
  } else {
    SpreadsheetApp.getUi().alert(`すべての販売記録の照合（商品マスタNoに数値入力）が完了しています。未照合データはありません。`);
  }
}

/**
 * 「未照合_販売記録」シートの F列（入力用_商品マスタNo）に入力された値を
 * 「販売記録CSV照合表」の「商品マスタ No」列へ一括転記するスクリプト
 */
function 未照合結果を照合表へ転記() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const unmappedSheet = ss.getSheetByName("未照合_販売記録");
  const matchSheet = ss.getSheetByName("販売記録CSV照合表");

  if (!unmappedSheet) {
    throw new Error("「未照合_販売記録」シートが見つかりません。先に「未照合の販売記録を出力」を実行してください。");
  }
  if (!matchSheet) {
    throw new Error("「販売記録CSV照合表」シートが見つかりません。");
  }

  const unmappedData = unmappedSheet.getDataRange().getValues();
  if (unmappedData.length <= 1) {
    SpreadsheetApp.getUi().alert("「未照合_販売記録」シートに転記対象のデータが存在しません。");
    return;
  }

  // 未照合シートのヘッダー解析
  const unmappedHeaders = unmappedData[0].map(h => String(h).trim());
  const idxUnmappedId = findHeaderIndex(unmappedHeaders, ["商品ID", "商品id", "id", "ID"]);
  const idxUnmappedMemo = findHeaderIndex(unmappedHeaders, ["入力用_商品マスタNo", "商品マスタNo", "メモ", "コード"]);

  const targetIdCol = idxUnmappedId !== -1 ? idxUnmappedId : 1; // デフォルトB列 (index 1)
  const targetMemoCol = idxUnmappedMemo !== -1 ? idxUnmappedMemo : 5; // デフォルトF列 (index 5)

  // 1. 入力済みの商品マスタNoをマップ化 (商品ID -> 入力マスタNo)
  const updateMap = new Map();
  for (let i = 1; i < unmappedData.length; i++) {
    const row = unmappedData[i];
    const id = String(row[targetIdCol]).trim();
    const memoVal = row.length > targetMemoCol ? String(row[targetMemoCol]).trim() : "";

    // 数値が含まれる有効なマスタNoが入力されているかチェック
    if (id && isValidMasterNo(memoVal)) {
      updateMap.set(id, memoVal);
    }
  }

  if (updateMap.size === 0) {
    SpreadsheetApp.getUi().alert("「入力用_商品マスタNo」列に数値（有効な商品マスタNo）が入力されていません。メモ入力後に再度実行してください。");
    return;
  }

  // 2. 「販売記録CSV照合表」のヘッダー取得・更新
  const matchData = matchSheet.getDataRange().getValues();
  const matchHeaders = matchData[0].map(h => String(h).trim());

  const idxMatchId = findHeaderIndex(matchHeaders, ["商品ID", "商品id", "id", "ID", "商品ＩＤ"]);
  const idxMatchMasterNo = findHeaderIndex(matchHeaders, ["商品マスタ No", "商品マスタNo", "No", "No."]);

  const destIdCol = idxMatchId !== -1 ? idxMatchId : 0;
  const destMasterNoCol = idxMatchMasterNo !== -1 ? idxMatchMasterNo : 2;

  let updatedCount = 0;
  for (let i = 1; i < matchData.length; i++) {
    const id = String(matchData[i][destIdCol]).trim();
    if (updateMap.has(id)) {
      const newMasterNo = updateMap.get(id);
      matchSheet.getRange(i + 1, destMasterNoCol + 1).setValue(newMasterNo);
      updatedCount++;
    }
  }

  // 3. 転記完了メッセージ & 未照合シートの再更新
  SpreadsheetApp.getUi().alert(`${updatedCount} 件の商品マスタ No を「販売記録CSV照合表」に転記しました。`);

  // 未照合シートを再抽出して最新化
  未照合販売記録を出力();
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

