/***** エントリーポイント：最新日付シートを「棚卸_最新」にリンク参照で反映 *****/
function 棚卸_最新_リンク参照() {
  // ==== 固有設定（関数内に閉じ込め） ====
  const FILE_NAME = '棚卸ファイル2025.10.05～';
  const SHEET_PREFIX = '棚卸_';
  const LATEST_SHEET_NAME = '棚卸_最新';
  const HEADER_ROW = 1;

  // ==== スプレッドシート取得 ====
  const it = DriveApp.getFilesByName(FILE_NAME);
  if (!it.hasNext()) throw new Error(`ドライブ内に「${FILE_NAME}」が見つかりません。`);
  const ss = SpreadsheetApp.openById(it.next().getId());

  const dst = ss.getSheetByName(LATEST_SHEET_NAME);
  if (!dst) throw new Error(`「${LATEST_SHEET_NAME}」が見つかりません。先に作成してください。`);

  // ==== 最新の棚卸シートを探す ====
  const re = new RegExp('^' + SHEET_PREFIX + '(\\d{4})-(\\d{2})-(\\d{2})$');
  let latest = null;
  let latestTime = -1;
  ss.getSheets().forEach(sh => {
    const m = sh.getName().match(re);
    if (m) {
      const t = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])).getTime();
      if (t > latestTime) { latestTime = t; latest = sh; }
    }
  });
  if (!latest) throw new Error('「棚卸_YYYY-MM-DD」形式の元シートが見つかりません。');

  // ==== リンク数式を張る ====
  const lastRow = Math.max(latest.getLastRow(), HEADER_ROW + 1);
  const lastCol = latest.getLastColumn();
  const endColLetter = columnToLetter_(lastCol);
  const srcName = latest.getName();
  const startAddr = `A${HEADER_ROW + 1}`;
  const endAddr = `${endColLetter}${lastRow}`;
  const formula = `=IFERROR(INDIRECT("'${srcName}'!${startAddr}:${endAddr}"),)`; 

  // データ部をクリアして新しい数式をセット
  const lastRowDst = dst.getLastRow();
  if (lastRowDst > HEADER_ROW) {
    dst.getRange(HEADER_ROW + 1, 1, lastRowDst - HEADER_ROW, dst.getLastColumn()).clearContent();
  }
  dst.getRange(HEADER_ROW + 1, 1).setFormula(formula);

  SpreadsheetApp.getActive().toast(`最新シート「${srcName}」を参照しました。`, '棚卸', 5);
}

/***** 列番号→列記号変換 *****/
function columnToLetter_(col) {
  let temp = '', n = col;
  while (n > 0) {
    let rem = (n - 1) % 26;
    temp = String.fromCharCode(65 + rem) + temp;
    n = Math.floor((n - 1) / 26);
  }
  return temp;
}
