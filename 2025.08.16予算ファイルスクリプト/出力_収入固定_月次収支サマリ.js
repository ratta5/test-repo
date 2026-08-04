/***** エントリーポイント *****/
function 出力_収入固定_月次収支サマリ() {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const SH_INCOME = '収入_固定';
  const SH_SUMMARY = '月次収支サマリ';
  const SUMMARY_MONTH_HEADER = '年月';       // A列想定
  const SUMMARY_TARGET_HEADER = '収入_固定';  // B列

  const shIncome = SS.getSheetByName(SH_INCOME);
  const shSummary = SS.getSheetByName(SH_SUMMARY);

  const incomeMap = getHeaderMap_(shIncome);
  const colStart = colOrThrow_(incomeMap, '開始年月');
  const colEnd   = colOrThrow_(incomeMap, '終了年月');
  const colAmt   = colOrThrow_(incomeMap, '金額');

  const summaryMap = getHeaderMap_(shSummary);
  const colMonth  = colOrThrow_(summaryMap, SUMMARY_MONTH_HEADER);
  const colTarget = colOrThrow_(summaryMap, SUMMARY_TARGET_HEADER);

  // 月リスト（A2:最終行）
  const lastRow = Math.max(shSummary.getLastRow(), 2);
  const monthVals = shSummary.getRange(2, colMonth, lastRow - 1, 1).getValues();
  const months = monthVals.map(r => normalizeToMonthDate_(r[0]))
                          .map(d => d ? monthKey_(d) : null);

  if (months.length === 0) return;

  // 合計配列
  const monthTotals = new Array(months.length).fill(0);

  // 収入_固定を集計
  const rows = shIncome.getDataRange().getValues().slice(1);
  for (const row of rows) {
    const start = row[colStart - 1];
    const end   = row[colEnd   - 1];
    const amt   = toNumberOrZero_(row[colAmt   - 1]);
    if (!amt) continue;

    const startKey = start ? monthKey_(normalizeToMonthDate_(start)) : null;
    const endKey   = end   ? monthKey_(normalizeToMonthDate_(end))   : null;

    months.forEach((mk, i) => {
      if (!mk) return; // 年月が空/無効はスキップ
      const okStart = (startKey === null) || (mk >= startKey);
      const okEnd   = (endKey   === null) || (mk <= endKey);
      if (okStart && okEnd) monthTotals[i] += amt;
    });
  }

  // --- B列をシート末尾まで丸ごと初期化（B1は残す）---
  const maxRows = shSummary.getMaxRows();
  if (maxRows > 1) {
    const wholeB = shSummary.getRange(2, colTarget, maxRows - 1, 1);
    wholeB.clearContent();
    wholeB.clearDataValidations();
    wholeB.clearFormat();
  }
  SpreadsheetApp.flush();

  // 書き込み（無効な年月行は空欄のまま）
  const outRange = shSummary.getRange(2, colTarget, months.length, 1);
  const out = monthTotals.map((v, i) => [months[i] ? Number(v) || 0 : '']);
  outRange.setValues(out);

  // 必要なら通貨表示（まずは数値として正常表示するか確認したいのでコメント可）
  outRange.setNumberFormat('¥#,##0;[Red]-¥#,##0;0');
}

/***** ヘルパー群 *****/
function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => map[String(h).trim()] = i + 1);
  return map;
}
function colOrThrow_(map, name) {
  const c = map[name];
  if (!c) throw new Error('見出し「' + name + '」が見つかりません');
  return c;
}
function normalizeToMonthDate_(v) {
  if (!v && v !== 0) return null;
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), 1);
  const s = String(v).trim();
  const m = s.match(/^(\d{4})[\/\-\.年]?(\d{1,2})/);
  if (m) return new Date(parseInt(m[1],10), parseInt(m[2],10)-1, 1);
  const n = Number(s);
  if (!isNaN(n)) {
    const d = new Date(Math.round((n - 25569) * 86400 * 1000));
    return new Date(d.getFullYear(), d.getMonth(), 1);
  }
  return null;
}
function monthKey_(d) {
  return d.getFullYear()*100 + (d.getMonth()+1);
}
function toNumberOrZero_(v) {
  if (typeof v === 'number') return v || 0;
  // 通貨記号/スペース/全角マイナス/長音等を除去、括弧マイナスにも対応
  const s = String(v).trim();
  const neg = /^\(.*\)$/.test(s);
  const cleaned = s.replace(/[\(\)]/g, '')
                   .replace(/[¥,\s\u3000\u2212\uFF0D]/g, '');
  const n = Number(cleaned);
  if (isNaN(n)) return 0;
  return neg ? -n : n;
}
