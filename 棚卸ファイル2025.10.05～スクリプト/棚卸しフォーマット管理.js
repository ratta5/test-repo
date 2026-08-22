/***** 棚卸しフォーマット管理スクリプト（安全版：式の自動注入なし） *****/
/***** 設定 *****/
const TZ = 'Asia/Tokyo';
const LATEST_TEMPLATE_NAME = '棚卸_最新フォーマット';
const HISTORY_SHEET_NAME   = 'フォーマット履歴';

/***** メニュー *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('棚卸し運用')
    .addItem('① 今日の棚卸しを作成（最新フォーマットから）', 'createTodayInventorySheet')
    .addItem('② 現在のシートを「最新フォーマット」に昇格', 'promoteActiveSheetToLatestFormat')
    .addSeparator()
    .addItem('フォーマット履歴を開く/用意する', 'openOrCreateHistory')
    // 既存運用に合わせたメニュー（関数未定義なら押すまで影響なし）
    .addItem('棚卸_最新_リンク参照', '棚卸_最新_リンク参照')
    .addToUi();
}

/***** 共通ユーティリティ *****/
function getSheetOrThrow_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`シートが見つかりません: ${name}`);
  return sh;
}
function ensureHistorySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(HISTORY_SHEET_NAME);
    sh.getRange(1,1,1,6).setValues([['Ver','元になったシート','変更日','変更者','主な変更点','備考']]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function getNextVersionLabel_() {
  const sh = ensureHistorySheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 'Ver1';
  const prev = sh.getRange(lastRow,1).getDisplayValue(); // 例: "Ver7"
  const n = Number((prev || '').replace(/[^0-9]/g,''));
  return `Ver${(isNaN(n)?0:n)+1}`;
}
function todayStr_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd');
}
function todaySheetSuffix_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

/***** ① 最新フォーマットから、今日の棚卸しを作成（式は入れない） *****/
function createTodayInventorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tmpl = getSheetOrThrow_(ss, LATEST_TEMPLATE_NAME);
  const newName = `棚卸_${todaySheetSuffix_()}`;

  if (ss.getSheetByName(newName)) {
    SpreadsheetApp.getUi().alert(`すでに存在します：${newName}`);
    return;
  }

  // テンプレをコピー（テンプレ内の見出し・体裁のみを継承。式はテンプレ側に無い前提）
  const newSheet = tmpl.copyTo(ss).setName(newName);
  ss.setActiveSheet(newSheet);

  // ★式の自動注入はしない★（棚卸後に直近シートから手動でコピペ）
  SpreadsheetApp.getUi().alert(`今日の棚卸しシートを作成しました：${newName}\n※ 数式は直近シートから手動でコピペしてください。`);
}

/***** ② 現在のシートを「棚卸_最新フォーマット」に昇格（コピー＆履歴のみ） *****/
function promoteActiveSheetToLatestFormat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const active = ss.getActiveSheet();
  const sourceName = active.getName();

  // 現在のシートをコピーして一時名で作成
  let latest = ss.getSheetByName(LATEST_TEMPLATE_NAME);
  const newLatest = active.copyTo(ss).setName(LATEST_TEMPLATE_NAME + '_tmp_replace');

  // 入力欄の初期化（任意）：数量が D 列想定なら「値だけ」消す（数式があっても触らない）
  // テンプレは“見出しだけ”の設計なので、ここは実質何もしなくてもOK。
  const lastRow = Math.max(newLatest.getLastRow(), 2);
  if (lastRow >= 2) {
    const qtyCol = 4; // D列=4（必要なら列名検出ロジックに差し替え可能）
    const rng = newLatest.getRange(2, qtyCol, lastRow - 1, 1);
    const formulas = rng.getFormulas();
    const values   = rng.getValues();
    for (let r = 0; r < formulas.length; r++) {
      if (!formulas[r][0]) values[r][0] = '';  // 数式なしセルのみ空に
    }
    rng.setValues(values);
  }

  // ★D2/E2/I2/J2 などへの setFormula は行わない★

  // 既存の最新フォーマットを削除し、差し替え
  if (latest) ss.deleteSheet(latest);
  newLatest.setName(LATEST_TEMPLATE_NAME);

  // 履歴追記
  const ver = getNextVersionLabel_();
  const shH = ensureHistorySheet_();
  const row = shH.getLastRow() + 1;
  const user = (Session.getActiveUser && Session.getActiveUser().getEmail()) || '';
  shH.getRange(row,1,1,6).setValues([[ver, sourceName, todayStr_(), user, 'コピーしてテンプレ差し替え（式注入なし）', '自動追記']]);

  SpreadsheetApp.getUi().alert(`フォーマットに昇格しました：${LATEST_TEMPLATE_NAME}\n履歴へ ${ver} を追記しました。\n※ 数式は直近シートから手動でコピペしてください。`);
}

/***** 補助：履歴シートを開く/作る *****/
function openOrCreateHistory() {
  const sh = ensureHistorySheet_();
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sh);
  SpreadsheetApp.getUi().alert('「フォーマット履歴」を開きました。');
}
