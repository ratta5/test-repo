/***** 設定 *****/
const DEST_SHEET_NAME   = '商品マスタ25.09-';
const DEST_HEADER_ROW   = 1;
const SRC_HEADER_ROW    = 10;
const SRC_DATA_STARTROW = 13;
const MONTH_SHEET_REGEX = /^\d{2}\.\d{1,2}月$/;
const DEST_WRITE_START_COL = 2;

/** メニュー */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('仕入・商品管理')
    .addItem('★シミュレーター全商品を履歴に一括保存', 'シミュレーター全商品を履歴に一括保存')
    .addItem('シミュレーターを一括クリア', 'シミュレーターを一括クリア')
    .addSeparator()
    .addItem('商品マスタ今すぐ更新', '商品マスタ更新')
    .addItem('仕入試算シートを新規作成/初期化', '仕入試算シート作成')
    .addToUi();
}

/** 集約本体 */
function 商品マスタ更新() {
  try {
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const shDest = ss.getSheetByName(DEST_SHEET_NAME);
    if (!shDest) throw new Error(`シート「${DEST_SHEET_NAME}」が見つかりません。`);

    const destHeader = getDestHeader_(shDest);
    clearBelowFromCol_(shDest, DEST_HEADER_ROW + 1, DEST_WRITE_START_COL, destHeader.length);

    const monthSheets = ss.getSheets().filter(sh => MONTH_SHEET_REGEX.test(sh.getName()));
    if (!monthSheets.length) {
      toast_('対象の月シートが見つかりませんでした。'); return;
    }

    let writeRow = DEST_HEADER_ROW + 1, copied = 0, blanked = 0, totalRows = 0;

    for (const shSrc of monthSheets) {
      const { rows, blank } = processSheet_(shSrc, destHeader, shDest, writeRow);
      if (!rows.length) continue;

      shDest.getRange(writeRow, DEST_WRITE_START_COL, rows.length, destHeader.length).setValues(rows);
      writeRow  += rows.length;
      copied    += rows.length;
      blanked   += blank;
      totalRows += rows.length;
      Logger.log(`[${shSrc.getName()}] ${rows.length}行（空欄化 ${blank}行）`);
    }

    toast_(
      `更新完了：${monthSheets.length}枚から ${totalRows}行を処理／空欄化 ${blanked}行。` +
      `「${DEST_SHEET_NAME}」B列以降に ${copied}行を反映（A列は未変更）。`
    );
  } catch (err) {
    Logger.log(`エラー: ${err?.stack || err}`);
    toast_(`エラー: ${err?.message || err}`);
    throw err;
  }
}

/* ===== ヘルパー ===== */

/** 出力先の見出し配列（B列以降・空列は除外）を返す */
function getDestHeader_(sh) {
  const all = readRow_(sh, DEST_HEADER_ROW).map(normalize_);
  const raw = all.slice(DEST_WRITE_START_COL - 1);
  const last = raw.reduce((i, v, idx) => v !== '' ? idx : i, -1);
  return raw.slice(0, last + 1);
}

/** 月シート1枚分を処理し、出力行配列と空欄化数を返す */
function processSheet_(shSrc, destHeader, shDest, writeRow) {
  const lastRow = shSrc.getLastRow();
  if (lastRow < SRC_DATA_STARTROW) return { rows: [], blank: 0 };

  const srcHeader    = readRow_(shSrc, SRC_HEADER_ROW).map(normalize_);
  const srcIndex     = buildIndexMap_(srcHeader);
  const idxPurchase  = srcIndex.get('購入有無');
  const numRows      = lastRow - SRC_DATA_STARTROW + 1;
  const srcValues    = shSrc.getRange(SRC_DATA_STARTROW, 1, numRows, shSrc.getLastColumn()).getValues();

  // 転記先「商品マスタ」シートのA列（商品マスタ No）の既存値を取得
  const destNoValues = (shDest && writeRow) ? shDest.getRange(writeRow, 1, numRows, 1).getValues() : [];

  // 「メルカリ管理コード」生成用の参照列インデックスを取得（月シート側）
  const idxNoInSrc     = findIndexByPatterns_(srcIndex, [/商品マスタ.*(No|№)/i, /商品マスタ/i, /^No\.?$/i, /^№$/i]);
  const idxBreakEven   = findIndexByPatterns_(srcIndex, [/損益分岐.*1[％%].*差引後/i, /損益分岐販売価格/i, /損益分岐/i]);
  const idxTargetPrice = findIndexByPatterns_(srcIndex, [/想定売価.*当初見込み/i, /想定売価/i]);

  let blank = 0;
  const rows = [];
  for (let r = 0; r < srcValues.length; r++) {
    const row = srcValues[r];
    if (isRowEmpty_(row)) continue;
    if (idxPurchase !== undefined && valueIsNoPurchase_(row[idxPurchase])) {
      rows.push(new Array(destHeader.length).fill(''));
      blank++;
    } else {
      // 商品マスタ No: 転記先A列の値を優先して取得
      const destNo = (destNoValues[r] && destNoValues[r][0] !== undefined) ? destNoValues[r][0] : '';
      const srcNo  = idxNoInSrc !== undefined ? row[idxNoInSrc] : '';
      const vNo    = (destNo !== '' && destNo !== null && destNo !== undefined) ? destNo : srcNo;

      rows.push(destHeader.map(col => {
        if (col === 'メルカリ管理コード') {
          let vBE = idxBreakEven !== undefined ? row[idxBreakEven] : '';
          let vTP = idxTargetPrice !== undefined ? row[idxTargetPrice] : '';

          // 10で割って四捨五入（整数化）
          if (vBE !== '' && vBE !== null && !isNaN(Number(vBE))) {
            vBE = Math.round(Number(vBE) / 10);
          }
          if (vTP !== '' && vTP !== null && !isNaN(Number(vTP))) {
            vTP = Math.round(Number(vTP) / 10);
          }

          const parts = [vNo, vBE, vTP].map(v => String(v ?? '').trim()).filter(v => v !== '');
          if (parts.length > 0) {
            return parts.join('-');
          }
          const i = srcIndex.get(col);
          return i !== undefined ? row[i] : '';
        }
        const i = srcIndex.get(col);
        return i !== undefined ? row[i] : '';
      }));
    }
  }
  return { rows, blank };
}

/** パターンリストにマッチする見出しのインデックスを検索 */
function findIndexByPatterns_(srcIndex, patterns) {
  for (const pattern of patterns) {
    for (const [key, idx] of srcIndex.entries()) {
      if (pattern.test(key)) return idx;
      // 記号やピリオドを除去した文字列でも検索
      const cleanKey = key.replace(/[.．_＿]/g, '');
      if (pattern.test(cleanKey)) return idx;
    }
  }
  return undefined;
}

/** 行読み取り（1始まり） */
function readRow_(sh, row) {
  return sh.getRange(row, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
}

/** startCol から numCols 幅のデータ部をクリア */
function clearBelowFromCol_(sh, startRow, startCol, numCols) {
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < startRow || lastCol < startCol || numCols <= 0) return;
  sh.getRange(startRow, startCol, lastRow - startRow + 1, numCols).clearContent();
}

/** 完全空行の判定 */
function isRowEmpty_(row) {
  return row.every(v => v === '' || v === null);
}

/** 見出し名を正規化（全角空白・前後空白・連続空白を除去） */
function normalize_(v) {
  return String(v ?? '').replace(/\u3000/g, ' ').trim().replace(/\s+/g, ' ');
}

/** 見出し名 → インデックスの Map を作成（空名除外・重複は先勝ち） */
function buildIndexMap_(arr) {
  return arr.reduce((m, name, i) => {
    if (name !== '' && !m.has(name)) m.set(name, i);
    return m;
  }, new Map());
}

/**
 * 「購入有無」が「2 無」または空欄かどうかを判定
 * 例: 2, '2', '2 無', '無', '2-無', '' (空欄) → true
 *     1, '1', '1 有', '有'                     → false
 */
function valueIsNoPurchase_(val) {
  if (val === null || val === '') return true;
  const s = normalize_(String(val));
  return /無/.test(s) || /^\s*2(\b|[^0-9])?/.test(s) || val === 2;
}

/** toast 通知（共通） */
function toast_(msg) {
  SpreadsheetApp.getActive().toast(msg, '商品集約', 8);
}