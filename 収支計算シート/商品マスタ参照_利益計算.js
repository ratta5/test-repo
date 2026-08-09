function onFormSubmit(e) {
  const sheetName = "販売記録"; // フォーム連携シート名
  const masterSheetName = "資産帳"; // 商品マスタシート名
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();

  // C列：商品マスタNo
  const productNoCell = `C${lastRow}`;

  // K列〜Q列に数式を挿入
  const formulas = {
    K: `=IFERROR(VLOOKUP(${productNoCell}, '${masterSheetName}'!$B$6:$Z, 3, FALSE), "")`,  // 品名
    L: `=IFERROR(VLOOKUP(${productNoCell}, '${masterSheetName}'!$B$6:$Z, 7, FALSE), "")`,  // 仕入額（P引後）
    M: `=IFERROR(D${lastRow} * VALUE(SUBSTITUTE(I${lastRow}, "%", "")) / 100, "")`,         // 販売手数料金額（"10%" 文字列対応）
    N: `=IFERROR(D${lastRow} - M${lastRow} - L${lastRow}, "")`,                            // 粗利
    O: `=IFERROR(N${lastRow} / D${lastRow}, "")`,                                           // 粗利益率
    P: `=IFERROR(N${lastRow} - F${lastRow} + H${lastRow}, "")`,                            // 実利益
    Q: `=IFERROR(P${lastRow} / D${lastRow}, "")`                                            // 実利益率
  };

  // 式をそれぞれの列に挿入
  for (let col in formulas) {
    const colNum = col.charCodeAt(0) - 64; // "A" → 1, "B" → 2 ...
    sheet.getRange(lastRow, colNum).setFormula(formulas[col]);
  }
}
function testInsertFormulas() {
  const sheetName = "販売記録"; // 実際のシート名に置き換えてください
  const masterSheetName = "資産帳";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();

  const productNoCell = `C${lastRow}`;
  const formulas = {
    K: `=IFERROR(VLOOKUP(${productNoCell}, '${masterSheetName}'!$B$6:$Z, 3, FALSE), "")`,
    L: `=IFERROR(VLOOKUP(${productNoCell}, '${masterSheetName}'!$B$6:$Z, 7, FALSE), "")`,
    M: `=IFERROR(D${lastRow} * VALUE(SUBSTITUTE(I${lastRow}, "%", "")) / 100, "")`,
    N: `=IFERROR(D${lastRow} - M${lastRow} - L${lastRow}, "")`,
    O: `=IFERROR(N${lastRow} / D${lastRow}, "")`,
    P: `=IFERROR(N${lastRow} - F${lastRow} + H${lastRow}, "")`,
    Q: `=IFERROR(P${lastRow} / D${lastRow}, "")`
  };

  for (let col in formulas) {
    const colNum = col.charCodeAt(0) - 64;
    sheet.getRange(lastRow, colNum).setFormula(formulas[col]);
  }
}

