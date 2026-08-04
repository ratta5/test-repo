/** 支出_見込み費用を月次収支サマリに反映 */
function 支出見込み費用を反映() {
  const TZ = 'Asia/Tokyo';
  const SUMMARY_SHEET = '月次収支サマリ';
  const SOURCE_SHEET  = '支出_見込み費用';
  const COL_SUMMARY_YM = 1; // A列: 年月
  const COL_SUMMARY_OUT = 5; // E列: 支出_見込み費用
  const COL_SRC_DATE = 1; // A列: 日付
  const COL_SRC_AMOUNT = 5; // E列: 金額

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shSummary = ss.getSheetByName(SUMMARY_SHEET);
  const shSrc = ss.getSheetByName(SOURCE_SHEET);
  if (!shSummary || !shSrc) {
    throw new Error(`シートが見つかりません: ${SUMMARY_SHEET} / ${SOURCE_SHEET}`);
  }

  // --- 1) サマリ側の年月キー一覧を作成 ---
  const lastRowSummary = shSummary.getLastRow();
  if (lastRowSummary < 2) return;
  const ymVals = shSummary.getRange(2, COL_SUMMARY_YM, lastRowSummary - 1, 1).getValues();
  const summaryKeys = ymVals.map(v => 月キーに変換_(v[0], TZ));
  const positions = {};
  summaryKeys.forEach((k, i) => { if (k) positions[k] = i; });

  // --- 2) 元データを集計 ---
  const srcData = shSrc.getDataRange().getValues();
  if (srcData.length <= 1) return;
  const body = srcData.slice(1);

  const monthly = {};
  body.forEach(row => {
    const d = 日付に変換_(row[COL_SRC_DATE - 1], TZ);
    const amt = 数値に変換_(row[COL_SRC_AMOUNT - 1]);
    if (!d || amt == null) return;
    const key = Utilities.formatDate(d, TZ, 'yyyy/MM');
    monthly[key] = (monthly[key] || 0) + amt;
  });

  // --- 3) サマリE列へ書き戻し ---
  const out = summaryKeys.map(k => [ (k && monthly[k]) ? monthly[k] : 0 ]);
  shSummary.getRange(2, COL_SUMMARY_OUT, out.length, 1).setValues(out);
}

/** "yyyy/MM" に変換 */
function 月キーに変換_(value, tz) {
  const d = 日付に変換_(value, tz);
  return d ? Utilities.formatDate(d, tz, 'yyyy/MM') : null;
}

/** 入力をDate化 */
function 日付に変換_(value, tz) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) return value;
  if (typeof value === 'string') {
    const s = value.trim().replace(/\./g, '/').replace(/-/g, '/');
    const m = s.match(/^(\d{4})[\/](\d{1,2})[\/](\d{1,2})/);
    if (m) {
      const y = +m[1], mo = +m[2]-1, d = +m[3];
      return new Date(y, mo, d);
    }
    const parsed = new Date(s);
    return isNaN(parsed) ? null : parsed;
  }
  if (typeof value === 'number') {
    const ms = Math.round((value - 25569) * 86400 * 1000);
    return new Date(ms);
  }
  return null;
}

/** 金額を数値に変換 */
function 数値に変換_(v) {
  if (v === null || v === '') return null;
  if (typeof v === 'number') return v;
  if (typeof v === 'string') {
    const n = Number(v.replace(/[,\s]/g, ''));
    return isNaN(n) ? null : n;
  }
  return null;
}
