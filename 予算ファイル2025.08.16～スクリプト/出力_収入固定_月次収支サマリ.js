/**
 * 収入固定_月次収支サマリ出力
 */
function 出力_収入固定_月次収支サマリ() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shIncome = getSheetOrThrow_(ss, '収入_固定');
  const shSummary = getSheetOrThrow_(ss, '月次収支サマリ');

  const colStart = getHeaderIndexOrThrow_(shIncome, ['開始年月']);
  const colEnd = getHeaderIndexOrThrow_(shIncome, ['終了年月']);
  const colAmt = getHeaderIndexOrThrow_(shIncome, ['金額']);
  const colMonth = getHeaderIndexOrThrow_(shSummary, ['年月']);
  const colTarget = getHeaderIndexOrThrow_(shSummary, ['収入_固定']);

  const lastRow = Math.max(shSummary.getLastRow(), 2);
  const months = shSummary.getRange(2, colMonth, lastRow - 1, 1).getValues().map(r => toYearMonthNumber_(r[0]));

  const incomeRows = shIncome.getDataRange().getValues().slice(1);
  const monthTotals = months.map(mKey => {
    if (!mKey) return [''];
    const sum = incomeRows.reduce((acc, row) => {
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
