/***** エントリーポイント *****/
function 出力_変動収入_物販_反映() {
  const TZ = 'Asia/Tokyo';
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const SH_SRC = '販売速報（フォーム回答）';   // 元データ：C=販売日, R=実利益
  const SH_SUM = '月次収支サマリ';            // 反映先：A=年月, L=変動収入（物販）
  const COL_MONTH_SUMMARY = 1;                // A列
  const COL_TARGET_SUMMARY = 12;              // L列（1始まりの列番号で12）
  const HEADER_TARGET = '変動収入（物販）';

  const shSrc = SS.getSheetByName(SH_SRC);
  const shSum = SS.getSheetByName(SH_SUM);
  if (!shSrc || !shSum) throw new Error('シート名の確認：販売速報（フォーム回答）／月次収支サマリ');

  /***********************
   * 1) 元データを月次集計
   ***********************/
  const lastRowSrc = shSrc.getLastRow();
  if (lastRowSrc < 2) return; // データなし

  // C列=販売日, R列=実利益 をまとめて取得（2行目から最終行まで）
  const rngSrc = shSrc.getRange(2, 3, lastRowSrc - 2 + 1, 16); // C列からR列まで(16列分)
  const values = rngSrc.getValues(); // [ [C,..,R], ... ]

  // 月別合計マップ { 'YYYY-MM': number }
  const monthSumMap = {};

  for (let i = 0; i < values.length; i++) {
    const dateRaw = values[i][0];      // C列（販売日）
    const profitRaw = values[i][16-1]; // R列（C起点で16列目）

    const ym = toYearMonth_(dateRaw, TZ);
    if (!ym) continue;

    const profit = toNumber_(profitRaw);
    if (profit === null) continue;

    monthSumMap[ym] = (monthSumMap[ym] || 0) + profit;
  }

  /****************************************
   * 2) サマリA列（年月）に合わせてL列を更新
   ****************************************/
  const lastRowSum = shSum.getLastRow();
  if (lastRowSum < 2) {
    // 見出しだけ作って終了
    shSum.getRange(1, COL_TARGET_SUMMARY).setValue(HEADER_TARGET);
    return;
  }

  // A列（年月）を取得（2行目～最終）
  const ymRange = shSum.getRange(2, COL_MONTH_SUMMARY, lastRowSum - 1, 1);
  const ymValues = ymRange.getValues(); // [[A2],[A3],...]

  // 出力配列を作成
  const out = new Array(ymValues.length).fill(0).map(_ => [0]);

  for (let r = 0; r < ymValues.length; r++) {
    const ymCell = ymValues[r][0];
    const ym = normalizeYearMonthLabel_(ymCell, TZ); // "YYYY-MM" へ正規化
    out[r][0] = ym && monthSumMap[ym] != null ? monthSumMap[ym] : 0;
  }

  // 見出し＆数値書式
  shSum.getRange(1, COL_TARGET_SUMMARY).setValue(HEADER_TARGET);
  shSum.getRange(2, COL_TARGET_SUMMARY, out.length, 1).setValues(out);
  shSum.getRange(2, COL_TARGET_SUMMARY, Math.max(out.length, 1), 1).setNumberFormat('#,##0'); // 任意：カンマ書式
}

/********** ユーティリティ **********/
// 任意の値を "YYYY-MM" に変換（失敗時は null）
function toYearMonth_(v, tz) {
  if (!v) return null;

  // 日付型
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, tz, 'yyyy-MM');
  }

  // 文字列→日付パースを試みる
  if (typeof v === 'string') {
    const s = v.trim();
    if (!s) return null;

    // 全角→半角
    const ascii = s.replace(/[！-～]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
                   .replace(/\s+/g, ' ');

    // よくある区切りをスラッシュに寄せる（YYYY-MM-DD / YYYY.MM.DD など）
    const norm = ascii.replace(/[.\-年]/g, '/').replace(/月/g, '/').replace(/日/g, '');

    const d = new Date(norm);
    if (!isNaN(d)) return Utilities.formatDate(d, tz, 'yyyy-MM');
  }

  // 数値（ほぼ来ない想定だが一応）：シリアル日付を日付に
  if (typeof v === 'number') {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000)); // Excel/Sheetsシリアル → JS Date
    if (!isNaN(d)) return Utilities.formatDate(d, tz, 'yyyy-MM');
  }

  return null;
}

// 金額を数値へ（"￥12,345" → 12345）。変換不可は null。
function toNumber_(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return v;

  if (typeof v === 'string') {
    // 全角→半角、通貨記号・カンマ・空白など除去
    const ascii = v.replace(/[！-～]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));
    const cleaned = ascii.replace(/[^\d.\-]/g, '');
    if (!cleaned) return null;
    const n = Number(cleaned);
    return isFinite(n) ? n : null;
  }
  return null;
}

// サマリA列の表示（"2025-01" 文字列 or 日付型）を "YYYY-MM" へ統一
function normalizeYearMonthLabel_(v, tz) {
  // 既に "YYYY-MM" 形式の文字列ならそのまま
  if (typeof v === 'string' && /^\d{4}-\d{2}$/.test(v.trim())) return v.trim();

  // 日付型なら "YYYY-MM"
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, tz, 'yyyy-MM');
  }

  // それ以外は toYearMonth_ に任せる
  return toYearMonth_(v, tz);
}
