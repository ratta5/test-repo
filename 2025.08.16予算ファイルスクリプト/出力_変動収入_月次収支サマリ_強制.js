/***** エントリーポイント：強制クリーン→再出力 *****/
function 出力_変動収入_月次収支サマリ_強制() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const DST_NAME = '月次収支サマリ';
  const DST_MONTH_HEADER = '年月';
  const DST_TARGET_HEADER = '変動収入（実績）';

  const SRC_NAME = '集計_変動';
  const SRC_MONTH_HEADER = '年月';
  const SRC_AMOUNT_HEADER = '収入';

  const shDst = getSheetOrThrow_(ss, DST_NAME);
  const shSrc = getSheetOrThrow_(ss, SRC_NAME);

  // ---- 出力先の列位置を特定 ----
  const header = toStringRow_(shDst.getRange(1,1,1,shDst.getLastColumn()).getValues()[0]);
  const colMonth  = findHeaderIndex_(header, [DST_MONTH_HEADER]) + 1;   // 1-based
  const colTarget = findHeaderIndex_(header, [DST_TARGET_HEADER]) + 1;  // 1-based
  const lastRow = shDst.getLastRow();
  if (lastRow < 2) throw new Error(`「${DST_NAME}」に月次の行がありません。`);

  // ---- J列（対象列）を強制クリーン ----
  hardCleanOneColumn_(shDst, colTarget, lastRow);

  // ---- ソース読み取り＆集計 ----
  const src = shSrc.getDataRange().getValues();
  if (src.length < 2) throw new Error(`「${SRC_NAME}」にデータ行がありません。`);
  const srcHeader = toStringRow_(src[0]);
  const cMonth = findHeaderIndex_(srcHeader, [SRC_MONTH_HEADER]);
  const cAmt   = findHeaderIndex_(srcHeader, [SRC_AMOUNT_HEADER]);

  const monthToAmt = new Map();
  for (let r=1; r<src.length; r++){
    const key = normalizeMonthKey_(src[r][cMonth]);
    if (!key) continue;
    const amt = toNumber_(src[r][cAmt]);
    monthToAmt.set(key, (monthToAmt.get(key)||0) + amt);
  }

  // ---- 月キーを参照して出力値を作成 ----
  const months = shDst.getRange(2, colMonth, lastRow-1, 1).getValues();
  const out = months.map(([v]) => {
    const key = normalizeMonthKey_(v);
    return [ key ? (monthToAmt.get(key) ?? 0) : 0 ];
  });

  // ---- 書き込み（値のみ） ----
  shDst.getRange(2, colTarget, out.length, 1).setValues(out)
       .setNumberFormat('#,##0;[Red]-#,##0;0');
  SpreadsheetApp.flush();
}

/***** 対象列を徹底クリーンする *****/
function hardCleanOneColumn_(sh, col, lastRow) {
  // 1) フィルタがあれば一旦外す（後で必要なら戻してください）
  const filter = sh.getFilter?.();
  if (filter) filter.remove();

  // 2) 結合解除（列全体）
  sh.getRange(1, col, sh.getMaxRows(), 1).breakApart();

  // 3) 保護（レンジ保護）を外す
  const protections = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];
  protections.forEach(p => {
    const r = p.getRange();
    if (r.getColumn() === col && r.getNumColumns() === 1) {
      // 対象列に完全一致の保護のみ解除（部分一致が多い場合は intersects 判定に変更）
      p.remove();
    }
  });

  // 4) データ検証・値・メモ・書式をクリア（見出し行は残す）
  const rng = sh.getRange(2, col, Math.max(lastRow-1, 0), 1);
  rng.clearContent();
  rng.clearDataValidations();
  rng.clearNote();
  rng.clearFormat(); // 表示形式・条件付き書式のローカル書式をクリア

  // 5) 見出しセル（1行目）が式なら値に固定（配列数式が居座っているケース対策）
  const h = sh.getRange(1, col).getValue();
  if (typeof h === 'string' && h.startsWith('=')) {
    const text = sh.getRange(1, col).getDisplayValue(); // 表示テキストを固定
    sh.getRange(1, col).setValue(text);
  }
}

/***** ユーティリティ *****/
function getSheetOrThrow_(ss, name){
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`シート「${name}」が見つかりません。`);
  return sh;
}
function toStringRow_(row){ return row.map(v => (v==null ? '' : String(v).trim())); }
function findHeaderIndex_(headerRow, cands){
  for (const c of cands){
    const i = headerRow.indexOf(String(c).trim());
    if (i !== -1) return i;
  }
  throw new Error(`ヘッダー「${cands.join(' / ')}」が見つかりません。`);
}
function toNumber_(v){
  if (v==null || v==='') return 0;
  if (typeof v === 'number') return v;
  const s = String(v).replace(/－/g,'-').replace(/[^\d.\-]/g,'');
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}
function normalizeMonthKey_(v){
  if (v==null || v==='') return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)){
    const y=v.getFullYear(), m=v.getMonth()+1;
    return `${y}-${('0'+m).slice(-2)}`;
  }
  const s = String(v).trim();
  const m = s.match(/(\d{4})\D{1,3}?(\d{1,2})/);
  if (m){
    const y=+m[1], mm=+m[2];
    if (y>=1900 && mm>=1 && mm<=12) return `${y}-${('0'+mm).slice(-2)}`;
  }
  const m2 = s.match(/^(\d{4})-(\d{2})$/);
  return m2 ? s : '';
}
