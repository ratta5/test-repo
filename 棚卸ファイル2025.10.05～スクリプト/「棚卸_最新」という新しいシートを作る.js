//このスクリプトが「やっていること」
//棚卸ファイル（例：棚卸ファイル2025.10.05～）を開く
//既存の 「棚卸_最新」 シートを削除
//ファイル内で最も新しい日付のシート（例：棚卸_2025-10-05）を探す
//そのシートをコピーして「棚卸_最新」という新しいシートを作る
//中のデータ（ヘッダー以外）はすべて消して、書式だけ残す

/***** 設定 *****/
const TZ_TANA = 'Asia/Tokyo';  // ← 名称変更（他と被らないように）
const FILE_NAME_TANA = '棚卸ファイル2025.10.05～';
const SHEET_PREFIX_TANA = '棚卸_';
const LATEST_SHEET_NAME_TANA = '棚卸_最新';
const HEADER_ROW_TANA = 1;

/***** エントリーポイント：棚卸_最新 シートを作成 *****/
function 棚卸_最新シート作成() {
  const ss = openSpreadsheetByName_TANA(FILE_NAME_TANA);

  // 既に「棚卸_最新」が存在する場合は削除
  const existing = ss.getSheetByName(LATEST_SHEET_NAME_TANA);
  if (existing) ss.deleteSheet(existing);

  // 最新の棚卸シートを探す
  const tmpl = findLatestInventorySheet_TANA(ss);
  if (!tmpl) throw new Error('「棚卸_YYYY-MM-DD」形式のテンプレシートが見つかりません。');

  // コピーして「棚卸_最新」にする
  const newSh = tmpl.copyTo(ss).setName(LATEST_SHEET_NAME_TANA);
  ss.setActiveSheet(newSh);

  // データ行をクリア
  clearDataRowsKeepFormats_TANA(newSh, HEADER_ROW_TANA);

  SpreadsheetApp.getActive().toast(`「${LATEST_SHEET_NAME_TANA}」を作成しました。`, '棚卸', 5);
}

/***** ヘルパー関数群（棚卸専用） *****/
function openSpreadsheetByName_TANA(name) {
  const it = DriveApp.getFilesByName(name);
  if (!it.hasNext()) throw new Error(`ドライブ内に「${name}」が見つかりません。`);
  return SpreadsheetApp.openById(it.next().getId());
}

function findLatestInventorySheet_TANA(ss) {
  const re = new RegExp('^' + SHEET_PREFIX_TANA + '(\\d{4})-(\\d{2})-(\\d{2})$');
  let latest = null;
  let latestTime = -1;

  ss.getSheets().forEach(sh => {
    const m = sh.getName().match(re);
    if (m) {
      const t = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])).getTime();
      if (t > latestTime) {
        latestTime = t;
        latest = sh;
      }
    }
  });
  return latest;
}

function clearDataRowsKeepFormats_TANA(sh, headerRow) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= headerRow) return;
  sh.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).clearContent();
}
