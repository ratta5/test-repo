/**
 * 販売記録（CSV取り込み）から新しい商品IDを抽出し、
 * 販売記録CSV照合表に追加するスクリプト。
 */
function 販売記録CSV照合の更新() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 対象シートの取得
  const srcSheet = ss.getSheetByName("販売記録（CSV取り込み）");
  const masterSheet = ss.getSheetByName("商品マスタ_統合");

  if (!srcSheet) {
    throw new Error("「販売記録（CSV取り込み）」シートが見つかりません。");
  }
  if (!masterSheet) {
    throw new Error("「商品マスタ_統合」シートが見つかりません。");
  }

  // 出力先シートの取得・作成
  const destSheetName = "販売記録CSV照合表";
  let destSheet = ss.getSheetByName(destSheetName);
  let isNewDestSheet = false;

  const destHeaders = ["商品ID", "メルカリ商品名", "商品マスタ No", "マスタ商品名", "購入者", "購入日時"];

  if (!destSheet) {
    destSheet = ss.insertSheet(destSheetName);
    destSheet.getRange(1, 1, 1, destHeaders.length).setValues([destHeaders]);
    // ヘッダー行を太字にして固定
    destSheet.getRange(1, 1, 1, destHeaders.length).setFontWeight("bold");
    destSheet.setFrozenRows(1);
    isNewDestSheet = true;
  } else {
    // 既存シートのヘッダーが不足している場合は「購入者」「購入日時」を追記
    const currentHeaders = destSheet.getRange(1, 1, 1, destHeaders.length).getValues()[0];
    if (currentHeaders.length < 5 || String(currentHeaders[4]).trim() === "") {
      destSheet.getRange(1, 5).setValue("購入者").setFontWeight("bold");
    }
    if (currentHeaders.length < 6 || String(currentHeaders[5]).trim() === "") {
      destSheet.getRange(1, 6).setValue("購入日時").setFontWeight("bold");
    }
  }

  // 2. ヘッダーの列インデックス特定
  // 入力元シート
  const srcData = srcSheet.getDataRange().getValues();
  if (srcData.length <= 1) {
    SpreadsheetApp.getUi().alert("「販売記録（CSV取り込み）」にデータが存在しません。");
    return;
  }
  const srcHeaders = srcData[0].map(h => String(h).trim());
  const idxSrcId = findHeaderIndex(srcHeaders, ["商品ID", "商品id", "id", "ID", "商品ＩＤ"]);
  const idxSrcName = findHeaderIndex(srcHeaders, ["商品名", "品名", "タイトル", "商品タイトル"]);
  const idxSrcBuyer = findHeaderIndex(srcHeaders, ["購入者", "バイヤー", "顧客名", "顧客", "購入者名"]);
  const idxSrcDate = findHeaderIndex(srcHeaders, ["購入日時", "販売日時", "取引日時", "日時", "取引完了日時", "日付"]);

  if (idxSrcId === -1) {
    throw new Error("「販売記録（CSV取り込み）」に「商品ID」列が見つかりません。");
  }
  if (idxSrcName === -1) {
    throw new Error("「販売記録（CSV取り込み）」に「商品名」列が見つかりません。");
  }

  // 商品マスタ_統合シート (数式作成用)
  const masterHeaders = masterSheet.getDataRange().getValues()[0].map(h => String(h).trim());
  const idxMasterNo = findHeaderIndex(masterHeaders, ["商品マスタ No", "商品マスタNo", "No", "No."]);
  const idxMasterName = findHeaderIndex(masterHeaders, ["商品名", "品名"]);

  if (idxMasterNo === -1 || idxMasterName === -1) {
    throw new Error("「商品マスタ_統合」に必要な列（商品マスタ No, 商品名）が見つかりません。");
  }

  const masterNoColLetter = getColumnLetter(idxMasterNo);
  const masterNameColLetter = getColumnLetter(idxMasterName);

  // 3. 既存の登録済みIDの取得
  const destData = destSheet.getDataRange().getValues();
  const existingIds = new Set();
  // 1行目はヘッダーなので2行目（インデックス1）から読み込む
  for (let i = 1; i < destData.length; i++) {
    const id = String(destData[i][0]).trim();
    if (id) {
      existingIds.add(id);
    }
  }

  // 4. 新規IDの抽出
  const newRows = [];
  const newlyAddedIds = new Set(); // 今回の実行内で重複して追加するのを防ぐ

  for (let i = 1; i < srcData.length; i++) {
    const row = srcData[i];
    const id = String(row[idxSrcId]).trim();
    const name = String(row[idxSrcName]).trim();
    const buyer = idxSrcBuyer !== -1 ? String(row[idxSrcBuyer]).trim() : "";

    if (id && !existingIds.has(id) && !newlyAddedIds.has(id)) {
      newRows.push({ id: id, name: name, buyer: buyer });
      newlyAddedIds.add(id);
    }
  }

  // 5. データの追加
  let newlyAddedCount = newRows.length;
  if (newlyAddedCount > 0) {
    let startRow = destSheet.getLastRow() + 1;
    const valuesToAppend = [];
    for (let i = 0; i < newlyAddedCount; i++) {
      valuesToAppend.push([
        newRows[i].id,
        newRows[i].name,
        "" // 商品マスタ No (手動入力用で最初は空欄)
      ]);
    }
    // 値の書き込み (A列〜C列)
    destSheet.getRange(startRow, 1, valuesToAppend.length, 3).setValues(valuesToAppend);
  }

  // 6. 全データ行の数式（D列・E列・F列）を一括設定・更新
  const lastRow = destSheet.getLastRow();
  if (lastRow > 1) {
    const dataRangeSize = lastRow - 1; // ヘッダーを除くデータ行数
    const formulasD = [];
    const formulasE = [];
    const formulasF = [];

    const srcIdColLetter = getColumnLetter(idxSrcId);
    const srcBuyerColLetter = idxSrcBuyer !== -1 ? getColumnLetter(idxSrcBuyer) : null;
    const srcDateColLetter = idxSrcDate !== -1 ? getColumnLetter(idxSrcDate) : null;

    for (let i = 2; i <= lastRow; i++) {
      // D列: マスタ商品名 (商品マスタ Noから引く)
      formulasD.push([
        `=XLOOKUP(C${i}, '商品マスタ_統合'!$${masterNoColLetter}:$${masterNoColLetter}, '商品マスタ_統合'!$${masterNameColLetter}:$${masterNameColLetter}, "マスタ未登録", 0)`
      ]);

      // E列: 購入者 (商品IDからCSV取り込みシートを引く)
      if (srcBuyerColLetter) {
        formulasE.push([
          `=XLOOKUP(A${i}, '販売記録（CSV取り込み）'!$${srcIdColLetter}:$${srcIdColLetter}, '販売記録（CSV取り込み）'!$${srcBuyerColLetter}:$${srcBuyerColLetter}, "購入者不明", 0)`
        ]);
      } else {
        formulasE.push([""]);
      }

      // F列: 購入日時 (商品IDからCSV取り込みシートを引く)
      if (srcDateColLetter) {
        formulasF.push([
          `=XLOOKUP(A${i}, '販売記録（CSV取り込み）'!$${srcIdColLetter}:$${srcIdColLetter}, '販売記録（CSV取り込み）'!$${srcDateColLetter}:$${srcDateColLetter}, "日時不明", 0)`
        ]);
      } else {
        formulasF.push([""]);
      }
    }

    // 数式の書き込み (D列)
    destSheet.getRange(2, 4, dataRangeSize, 1).setFormulas(formulasD);
    // 数式の書き込み (E列)
    destSheet.getRange(2, 5, dataRangeSize, 1).setFormulas(formulasE);
    // 数式の書き込み (F列)
    destSheet.getRange(2, 6, dataRangeSize, 1).setFormulas(formulasF);
  }

  if (newlyAddedCount > 0) {
    SpreadsheetApp.getUi().alert(`新たに ${newlyAddedCount} 件の商品IDを「販売記録CSV照合表」に追加しました。購入者列等もルックアップ関数で更新しました。`);
  } else {
    SpreadsheetApp.getUi().alert("新しく登録が必要な商品IDはありませんでした。購入者列等の数式を再設定・更新しました。");
  }
}

/**
 * ヘッダー名リストから、候補となる文字列に一致するインデックスを返す
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

/**
 * 0始まりの列インデックスをアルファベットの列文字（A, B, C...）に変換する
 */
function getColumnLetter(colIndex) {
  let letter = "";
  let temp = colIndex;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}
