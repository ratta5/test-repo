function SHEETNAME() {
  var name = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  return [[name.toString()]];
}
