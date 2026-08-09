function 年間正味収支を正味収支シートに表示() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("正味収支");
  if (!sheet) throw new Error("シート『正味収支』が見つかりません。");

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error("データが不足しています。");

  const header = data[0];
  const monthIndex = 0; // A列：月（Date型）
  const totalIndex = header.length - 1; // 最後の列（合計）

  const yearMap = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const monthCell = row[monthIndex];
    const total = row[totalIndex];

    if (!(monthCell instanceof Date) || typeof total !== "number") continue;

    const year = monthCell.getFullYear().toString(); // "2025"
    yearMap[year] = (yearMap[year] || 0) + total;
  }

  const sortedYears = Object.keys(yearMap).sort();

  // 出力位置：AA列以降（Z列 = 26列 → AA = 27列）
  const startCol = header.length + 2;
  const startRow = 1;

  const output = [["年", "年別正味収支合計"]];
  sortedYears.forEach(year => {
    output.push([year, yearMap[year]]);
  });

  // 書き込み前に範囲クリア（必要なら上書き）
  const targetRange = sheet.getRange(startRow, startCol, output.length, output[0].length);
  targetRange.clearContent();
  targetRange.setValues(output);
}
