/***** エントリーポイント *****/
function 更新_支出固定費() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shSummary = ss.getSheetByName('月次収支サマリ');
  const shFixed   = ss.getSheetByName('支出_固定費');
  if (!shSummary || !shFixed) {
    throw new Error('必要なシートが見つかりません（"月次収支サマリ", "支出_固定費"）。');
  }

  // --- ヘッダ位置特定（日本語ヘッダ名で検索） ---
  const headSummary = getHeaderMap_(shSummary);
  const headFixed   = getHeaderMap_(shFixed);

  const COL_YM_SUMMARY     = must(headSummary, '年月');
  const COL_OUT_FIXED_SUM  = must(headSummary, '支出_固定費');

  const COL_MONTHLY_EQ     = must(headFixed,   '月換算額'); // 合計に使う
  const COL_START          = must(headFixed,   '開始年月'); // 月判定に使用
  const COL_END            = must(headFixed,   '終了年月'); // 月判定に使用

  // --- 支出_固定費 シートの全行を取得（2行目以降） ---
  const lastRowFixed = shFixed.getLastRow();
  const lastColFixed = shFixed.getLastColumn();
  const fixedValues = lastRowFixed >= 2
    ? shFixed.getRange(2, 1, lastRowFixed - 1, lastColFixed).getValues()
    : [];

  // 固定費行をオブジェクト配列に整形
  const fixedRows = fixedValues.map(r => ({
    monthlyEq : toNumberSafe_(r[COL_MONTHLY_EQ - 1]),
    start     : toMonthStart_(r[COL_START - 1]), // Date|null
    end       : toMonthStart_(r[COL_END   - 1])  // Date|null
  }));

  // --- 月次収支サマリ側の処理 ---
  const lastRowSummary = shSummary.getLastRow();
  if (lastRowSummary < 2) return;

  // 対象列の現値を取得（上書き用）
  const ymVals     = shSummary.getRange(2, COL_YM_SUMMARY, lastRowSummary - 1, 1).getValues().map(v => v[0]);
  const writeRange = shSummary.getRange(2, COL_OUT_FIXED_SUM, lastRowSummary - 1, 1);
  const out = [];

  for (let i = 0; i < ymVals.length; i++) {
    const ym = toMonthStart_(ymVals[i]); // その行の対象月（各日付を月初に正規化）
    if (!ym) { out.push([0]); continue; }

    let total = 0;
    for (const row of fixedRows) {
      const active =
        (row.start === null || !row.start || ym.getTime() >= row.start.getTime()) &&
        (row.end   === null || !row.end   || ym.getTime() <= row.end.getTime());
      if (active) total += row.monthlyEq;
    }
    out.push([total]);
  }

  writeRange.setValues(out);
}

/***** メニュー追加（任意） *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('収支ユーティリティ')
    .addItem('支出_固定費を更新', '更新_支出固定費')
    .addToUi();
}

/***** ヘルパー群 *****/
// ヘッダ名→列番号のマップを返す（1始まり）
function getHeaderMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  header.forEach((name, idx) => {
    if (name !== '' && name !== null) map[String(name).trim()] = idx + 1;
  });
  return map;
}

function must(map, key) {
  const c = map[key];
  if (!c) throw new Error('ヘッダ "' + key + '" が見つかりません。');
  return c;
}

// 値を月初日(Date)に正規化。空欄は null。
function toMonthStart_(v) {
  if (v === null || v === '' || typeof v === 'undefined') return null;

  // Date型（シリアルを含む）はその月の1日に
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return new Date(v.getFullYear(), v.getMonth(), 1);
  }

  // 文字列パターン: "YYYY-MM" / "YYYY/MM" / "YYYY-MM-DD" / "YYYY/MM/DD"
  if (typeof v === 'string') {
    const s = v.trim();
    let m = s.match(/^(\d{4})[-\/](\d{1,2})(?:[-\/]\d{1,2})?$/);
    if (m) {
      const y = parseInt(m[1], 10);
      const mo = parseInt(m[2], 10) - 1; // 0始まり
      if (!isNaN(y) && !isNaN(mo)) return new Date(y, mo, 1);
    }
    // "YYYYMM" のような連結形式
    m = s.match(/^(\d{4})(\d{2})$/);
    if (m) {
      const y = parseInt(m[1], 10);
      const mo = parseInt(m[2], 10) - 1;
      if (!isNaN(y) && !isNaN(mo)) return new Date(y, mo, 1);
    }
  }

  return null; // 解釈できない場合は無視
}

// 数値に安全変換（空欄や非数は0）
function toNumberSafe_(v) {
  if (v === null || v === '' || typeof v === 'undefined') return 0;
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  // "12,345" のような文字列も許容
  const n = Number(String(v).replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}
