/**
 * 年月を "YYYY-MM" (または指定の区切り文字) の文字列に変換する
 */
function toYearMonthString_(v, separator = '-') {
  if (v === null || v === '' || typeof v === 'undefined') return '';
  const tz = 'Asia/Tokyo';
  
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return '';
    return Utilities.formatDate(v, tz, `yyyy${separator}MM`);
  }
  
  if (typeof v === 'number') {
    if (v > 100000) { // 例: 202508
      const s = String(v);
      return s.slice(0, 4) + separator + s.slice(4, 6);
    }
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, tz, `yyyy${separator}MM`);
    }
    return '';
  }
  
  if (typeof v === 'string') {
    const s = v.trim();
    if (!s) return '';
    const ascii = s.replace(/[！-～]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
                   .replace(/\s+/g, ' ');
    // YYYY/MM/DD, YYYY-MM-DD, YYYY年MM月DD日 など
    const m = ascii.match(/^(\d{4})[-\/\.年\s]?\s*(\d{1,2})(?:[-\/\.月\s]?\s*(\d{1,2})日?)?$/);
    if (m) {
      const y = parseInt(m[1], 10);
      const mo = parseInt(m[2], 10);
      if (y >= 1900 && mo >= 1 && mo <= 12) {
        return `${y}${separator}${('0' + mo).slice(-2)}`;
      }
    }
    const norm = ascii.replace(/[.\-年]/g, '/').replace(/月/g, '/').replace(/日/g, '');
    const d = new Date(norm);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, tz, `yyyy${separator}MM`);
    }
  }
  return '';
}

/**
 * 年月を YYYYMM (数値) に変換する
 */
function toYearMonthNumber_(v) {
  const ymStr = toYearMonthString_(v, '');
  return ymStr ? parseInt(ymStr, 10) : null;
}

/**
 * 通貨記号やカンマ等を除去して数値化する
 */
function toNumberSafe_(v) {
  if (v === null || v === '' || typeof v === 'undefined') return 0;
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  const ascii = String(v).replace(/[！-～]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));
  const cleaned = ascii.replace(/[^\d.\-]/g, '');
  const n = Number(cleaned);
  return isNaN(n) ? 0 : n;
}

/**
 * シートが見つからない場合にエラーを投げる
 */
function getSheetOrThrow_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`シート「${name}」が見つかりません。`);
  return sh;
}

/**
 * ヘッダー名リストから一致する列インデックスを探す (0始まり)
 */
function findHeaderIndex_(headers, candidates) {
  const canon = s => String(s).trim().toLowerCase();
  const canonHeaders = headers.map(canon);
  for (const cand of candidates) {
    const c = canon(cand);
    const i = canonHeaders.indexOf(c);
    if (i !== -1) return i;
  }
  for (let i = 0; i < canonHeaders.length; i++) {
    const h = canonHeaders[i];
    for (const cand of candidates) {
      const c = canon(cand);
      if (h.includes(c)) return i;
    }
  }
  return -1;
}

/**
 * 一致する列インデックスを1始まりで取得する。見つからない場合はエラー。
 */
function getHeaderIndexOrThrow_(sheet, candidates) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  const idx = findHeaderIndex_(headers, candidates);
  if (idx === -1) {
    throw new Error(`シート「${sheet.getName()}」でヘッダー列が見つかりません（候補: ${candidates.join(', ')}）`);
  }
  return idx + 1;
}

/**
 * 汎用的な明細データの月次集計およびサマリシート反映関数
 */
function 集計してサマリへ反映_(params) {
  const {
    元シート名,
    元日付ヘッダ候補,
    元金額ヘッダ候補,
    サマリ列ヘッダ,
    クリーンアップ要否 = false,
    金額書式 = '#,##0'
  } = params;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shSrc = getSheetOrThrow_(ss, 元シート名);
  const shDst = getSheetOrThrow_(ss, '月次収支サマリ');

  // --- 1) 元データの位置特定と集計 ---
  const lastRowSrc = shSrc.getLastRow();
  const monthSumMap = {};

  if (lastRowSrc >= 2) {
    const srcHeaders = shSrc.getRange(1, 1, 1, shSrc.getLastColumn()).getValues()[0];
    const colDateIdx = findHeaderIndex_(srcHeaders, 元日付ヘッダ候補);
    const colAmtIdx = findHeaderIndex_(srcHeaders, 元金額ヘッダ候補);

    if (colDateIdx === -1 || colAmtIdx === -1) {
      throw new Error(`シート「${元シート名}」で日付列または金額列が見つかりません。`);
    }

    const values = shSrc.getRange(2, 1, lastRowSrc - 1, shSrc.getLastColumn()).getValues();
    for (let i = 0; i < values.length; i++) {
      const dateVal = values[i][colDateIdx];
      const amtVal = values[i][colAmtIdx];
      const ym = toYearMonthString_(dateVal, '-'); // YYYY-MM で比較
      if (!ym) continue;
      
      const amt = toNumberSafe_(amtVal);
      monthSumMap[ym] = (monthSumMap[ym] || 0) + amt;
    }
  }

  // --- 2) 反映先（サマリ）の準備 ---
  const lastRowDst = shDst.getLastRow();
  if (lastRowDst < 2) return;

  const dstHeaders = shDst.getRange(1, 1, 1, shDst.getLastColumn()).getValues()[0];
  const colMonthIdx = findHeaderIndex_(dstHeaders, ['年月', '月', 'YM']);
  const colTargetIdx = findHeaderIndex_(dstHeaders, [サマリ列ヘッダ]);

  if (colMonthIdx === -1 || colTargetIdx === -1) {
    throw new Error(`サマリシートで年月列または「${サマリ列ヘッダ}」列が見つかりません。`);
  }

  const colMonth = colMonthIdx + 1;
  const colTarget = colTargetIdx + 1;

  // クリーンアップが必要な場合（強制更新など）
  if (クリーンアップ要否) {
    対象列の徹底クリーン_(shDst, colTarget, lastRowDst);
  }

  // --- 3) 年月列に合わせて書き込み用データを作成 ---
  const ymValues = shDst.getRange(2, colMonth, lastRowDst - 1, 1).getValues();
  const out = ymValues.map(([cellVal]) => {
    const ym = toYearMonthString_(cellVal, '-');
    return [ ym && monthSumMap[ym] ? monthSumMap[ym] : 0 ];
  });

  // --- 4) 書き込み ---
  const writeRange = shDst.getRange(2, colTarget, out.length, 1);
  writeRange.setValues(out);
  if (金額書式) {
    writeRange.setNumberFormat(金額書式);
  }
}

/**
 * 対象列を徹底クリーンする（強制更新用）
 */
function 対象列の徹底クリーン_(sh, col, lastRow) {
  const filter = sh.getFilter?.();
  if (filter) filter.remove();

  sh.getRange(1, col, sh.getMaxRows(), 1).breakApart();

  const protections = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];
  protections.forEach(p => {
    const r = p.getRange();
    if (r.getColumn() === col && r.getNumColumns() === 1) {
      p.remove();
    }
  });

  const rng = sh.getRange(2, col, Math.max(lastRow - 1, 0), 1);
  rng.clearContent();
  rng.clearDataValidations();
  rng.clearNote();
  rng.clearFormat();

  const h = sh.getRange(1, col).getValue();
  if (typeof h === 'string' && h.startsWith('=')) {
    const text = sh.getRange(1, col).getDisplayValue();
    sh.getRange(1, col).setValue(text);
  }
}
