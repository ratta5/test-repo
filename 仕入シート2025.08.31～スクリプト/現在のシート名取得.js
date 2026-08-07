/**
 * このセルが属するシート名を返す
 * 使い方例: =SHEETNAME(A1) あるいは =SHEETNAME(ROW())
 */
function SHEETNAME(dummy) {
  return SpreadsheetApp.getActiveRange().getSheet().getName();
}
