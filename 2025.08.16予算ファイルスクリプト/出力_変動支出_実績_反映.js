/***** 集計_変動(A=年月, B=支出) → 月次収支サマリ(M=変動支出（実績）) 反映 *****/
function 出力_変動支出_実績_反映() {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const SH_SRC  = '集計_変動';         // 元：A=年月, B=支出
  const SH_DEST = '月次収支サマリ';     // 先：A=年月, M=変動支出（実績）

  const COL_DEST_MONTH = 1;   // 先A列（年月）
  const COL_DEST_OUT   = 13;  // 先M列（変動支出（実績））
  const HEADER_OUT     = '変動支出（実績）';
  const TZ = Session.getScriptTimeZone() || 'Asia/Tokyo';

  const shSrc  = SS.getSheetByName(SH_SRC);
  const shDest = SS.getSheetByName(SH_DEST);
  if (!shSrc || !shDest) throw new Error('シート名を確認：集計_変動／月次収支サマリ');

  // --- 元データ取得 ---
  const lastRowSrc = shSrc.getLastRow();
  // 見出し行のみなら終了
  if (lastRowSrc < 2) {
    shDest.getRange(1, COL_DEST_OUT).setValue(HEADER_OUT);
    return;
  }
  const ymSrc  = shSrc.getRange(2, 1, lastRowSrc - 1, 1).getValues().flat(); // A:年月
  const expSrc = shSrc.getRange(2, 2, lastRowSrc - 1, 1).getValues().flat(); // B:支出

  // 月→合計のマップを作成（重複月は合算／文字列金額にも対応）
  const expMap = {};
  for (let i = 0; i < ymSrc.length; i++) {
    const ymLabel = normalizeYearMonthLabel_(ymSrc[i], TZ); // "YYYY-MM"に寄せる
    if (!ymLabel) continue;
    const n = toNumberSafe_(expSrc[i]);
    if (n == null) continue;
    expMap[ymLabel] = (expMap[ymLabel] || 0) + n;
  }

  // --- 反映先の年月に合わせてM列を作成 ---
  const lastRowDest = shDest.getLastRow();
  if (lastRowDest < 2) {
    shDest.getRange(1, COL_DEST_OUT).setValue(HEADER_OUT);
    return;
  }
  const ymDestVals = shDest.getRange(2, COL_DEST_MONTH, lastRowDest - 1, 1).getValues();
  const out = new Array(ymDestVals.length);

  for (let r = 0; r < ymDestVals.length; r++) {
    const ym = normalizeYearMonthLabel_(ymDestVals[r][0], TZ); // 先A列も正規化して突合
    out[r] = [ (ym && expMap[ym] != null) ? expMap[ym] : 0 ];
  }

  // 見出し・書式・書き込み
  shDest.getRange(1, COL_DEST_OUT).setValue(HEADER_OUT);
  const rngOut = shDest.getRange(2, COL_DEST_OUT, out.length, 1);
  rngOut.clearContent();
  rngOut.setValues(out);
  rngOut.setNumberFormat('#,##0'); // 任意：円書式にしたい場合は '[$¥-ja-JP]#,##0'
}

/***** 値を "YYYY-MM" ラベルに正規化（文字列/日付/シリアルに対応）*****/
function normalizeYearMonthLabel_(v, tz) {
  if (v == null || v === '') return null;

  // すでに "YYYY-MM" 文字列
  if (typeof v === 'string') {
    const s = v.trim();
    if (/^\d{4}-\d{2}$/.test(s)) return s;

    // 全角→半角、"2025/09/01" 等をDate化
    const norm = s
      .replace(/[！-～]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
      .replace(/[.\-年]/g, '/').replace(/月/g, '/').replace(/日/g, '').trim();
    const d = new Date(norm);
    if (!isNaN(d)) return Utilities.formatDate(d, tz, 'yyyy-MM');
    return null;
  }

  // 日付型
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, tz, 'yyyy-MM');
  }

  // 数値（シリアル日付）
  if (typeof v === 'number') {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000)); // Excel/Sheetsシリアル
    if (!isNaN(d)) return Utilities.formatDate(d, tz, 'yyyy-MM');
  }

  return null;
}

/***** 通貨記号・全角・カンマ等を除去して数値化 *****/
function toNumberSafe_(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;

  const s = String(v)
    .replace(/[！-～]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)) // 全角→半角
    .replace(/[^\d.\-]/g, '')  // 通貨記号・カンマ・空白除去
    .trim();
  if (!s) return null;

  const n = Number(s);
  return isFinite(n) ? n : null;
}
