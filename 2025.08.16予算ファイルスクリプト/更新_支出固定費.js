/**
 * 支出_固定費を更新
 */
function 更新_支出固定費() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shFixed = getSheetOrThrow_(ss, '支出_固定費');
  const shSummary = getSheetOrThrow_(ss, '月次収支サマリ');

  const colStart = getHeaderIndexOrThrow_(shFixed, ['開始年月']);
  const colEnd = getHeaderIndexOrThrow_(shFixed, ['終了年月']);
  const colAmt = getHeaderIndexOrThrow_(shFixed, ['月換算額']);
  const colMonth = getHeaderIndexOrThrow_(shSummary, ['年月']);
  const colTarget = getHeaderIndexOrThrow_(shSummary, ['支出_固定費']);

  const lastRow = Math.max(shSummary.getLastRow(), 2);
  const months = shSummary.getRange(2, colMonth, lastRow - 1, 1).getValues().map(r => toYearMonthNumber_(r[0]));

  const fixedRows = shFixed.getDataRange().getValues().slice(1);
  const monthTotals = months.map(mKey => {
    if (!mKey) return [''];
    const sum = fixedRows.reduce((acc, row) => {
      const amt = toNumberSafe_(row[colAmt - 1]);
      if (!amt) return acc;

      const startKey = toYearMonthNumber_(row[colStart - 1]);
      const endKey = toYearMonthNumber_(row[colEnd - 1]);
      const okStart = !startKey || mKey >= startKey;
      const okEnd = !endKey || mKey <= endKey;

      return (okStart && okEnd) ? acc + amt : acc;
    }, 0);
    return [sum];
  });

  const targetRange = shSummary.getRange(2, colTarget, monthTotals.length, 1);
  targetRange.clearContent();
  targetRange.setValues(monthTotals);
  targetRange.setNumberFormat('¥#,##0;[Red]-¥#,##0;0');
}
