/***** 収入固定_月次収支サマリ出力（スリム版） *****/
function 出力_収入固定_月次収支サマリ() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shIncome = ss.getSheetByName('収入_固定');
  const shSummary = ss.getSheetByName('月次収支サマリ');

  // 列インデックスの取得 (0始まり)
  const getColIdx = (sh, name) => {
    const idx = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].indexOf(name);
    if (idx === -1) throw new Error(`見出し「${name}」が見つかりません`);
    return idx;
  };

  const colStart = getColIdx(shIncome, '開始年月');
  const colEnd = getColIdx(shIncome, '終了年月');
  const colAmt = getColIdx(shIncome, '金額');
  const colMonth = getColIdx(shSummary, '年月') + 1;  // Range用に1始まり
  const colTarget = getColIdx(shSummary, '収入_固定') + 1;

  // サマリ側の「年月」リストを取得して YYYYMM (例: 202508) の数値キーに変換
  const lastRow = Math.max(shSummary.getLastRow(), 2);
  const months = shSummary.getRange(2, colMonth, lastRow - 1, 1).getValues().map(r => parseMonthKey_(r[0]));

  if (months.length === 0) return;

  // 収入_固定データの集計
  const incomeRows = shIncome.getDataRange().getValues().slice(1);
  const monthTotals = months.map(mKey => {
    if (!mKey) return ['']; // 年月が空の行は空欄

    const sum = incomeRows.reduce((acc, row) => {
      const amt = Number(String(row[colAmt]).replace(/[^\d.-]/g, '')) || 0; // ¥やカンマを除去して数値化
      if (!amt) return acc;

      const startKey = parseMonthKey_(row[colStart]);
      const endKey = parseMonthKey_(row[colEnd]);
      const okStart = !startKey || mKey >= startKey;
      const okEnd = !endKey || mKey <= endKey;

      return (okStart && okEnd) ? acc + amt : acc;
    }, 0);

    return [sum];
  });

  // クリア & 出力
  const maxRows = shSummary.getMaxRows();
  if (maxRows > 1) {
    shSummary.getRange(2, colTarget, maxRows - 1, 1).clearContent();
  }

  const outRange = shSummary.getRange(2, colTarget, monthTotals.length, 1);
  outRange.setValues(monthTotals);
  outRange.setNumberFormat('¥#,##0;[Red]-¥#,##0;0');
}

/**
 * 年月を YYYYMM (数値) に変換するヘルパー
 */
function parseMonthKey_(val) {
  if (!val && val !== 0) return null;
  if (val instanceof Date) return val.getFullYear() * 100 + (val.getMonth() + 1);
  const m = String(val).match(/^(\d{4})[\/\-\.年]?(\d{1,2})/);
  return m ? parseInt(m[1], 10) * 100 + parseInt(m[2], 10) : null;
}
