/***** エントリーポイント *****/
function 収入_変動を月次収支サマリへ反映() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SRC_NAME = '収入_変動';
  const DST_NAME = '月次収支サマリ';
  const DST_CATEGORY_HEADER = '収入_変動';

  const shSrc = getSheetOrThrow_(ss, SRC_NAME);
  const shDst = getSheetOrThrow_(ss, DST_NAME);

  // ---- 収入_変動を読み取り & 月(YYYY/MM)ごとに集計 ----
  const src = shSrc.getDataRange().getValues();
  if (src.length < 2) {
    throw new Error('「収入_変動」シートにデータ行がありません。');
  }
  const srcHeader = src[0].map(String);

  const dateCol = findHeaderIndex_(srcHeader, ['日付','日時','date','日付け']);
  const amtCol  = findHeaderIndex_(srcHeader, ['金額','金額(円)','amount','値']);

  if (dateCol === -1 || amtCol === -1) {
    throw new Error('「収入_変動」シートで「日付」または「金額」列が見つかりません。見出しをご確認ください。');
  }

  /** @type {Object.<string, number>} */
  const monthSum = {};
  for (let i = 1; i < src.length; i++) {
    const row = src[i];
    const d = row[dateCol];
    const amount = toNumberOrZero_(row[amtCol]);
    if (!amount) continue;
    const ym = toYYYYMM_(d);
    if (!ym) continue;
    monthSum[ym] = (monthSum[ym] || 0) + amount;
  }

  // ---- 月次収支サマリへ書き込み ----
  const dst = shDst.getDataRange().getValues();
  if (dst.length < 2) {
    throw new Error('「月次収支サマリ」シートに見出しまたはデータ行がありません。');
  }
  const dstHeader = dst[0].map(String);

  // 年月列（基本は先頭列を想定、見出し名でも検出）
  let ymCol = 0; // 既定: 1列目
  const ymHeaderIdx = findHeaderIndex_(dstHeader, ['年月','月','YM','YearMonth']);
  if (ymHeaderIdx !== -1) ymCol = ymHeaderIdx;

  // 収入_変動列を検出
  const catCol = findHeaderIndex_(dstHeader, [DST_CATEGORY_HEADER]);
  if (catCol === -1) {
    throw new Error('「月次収支サマリ」シートに「収入_変動」列が見つかりません。見出しをご確認ください。');
  }

  // 書き込み用バッファ（2行目以降）
  const out = [];
  for (let r = 1; r < dst.length; r++) {
    const ymCell = dst[r][ymCol];
    const ym = normalizeYYYYMMCell_(ymCell);
    const sum = ym ? (monthSum[ym] || 0) : 0;
    // 既存の行配列をコピーして目的列だけ更新
    const rowCopy = dst[r].slice();
    rowCopy[catCol] = sum;
    out.push(rowCopy);
  }
    // ---- 書き戻し（対象列のみ）+ 表示形式リセット ----
  if (dst.length > 1) {
    const colValues = [];
    for (let r = 1; r < dst.length; r++) {
      const ymCell = dst[r][ymCol];
      const ym = normalizeYYYYMMCell_(ymCell);
      const sum = ym ? (monthSum[ym] || 0) : 0;
      colValues.push([sum]);
    }

    const rng = shDst.getRange(2, catCol + 1, colValues.length, 1);
    rng.setValues(colValues);

    // 表示形式を通貨（円）に設定
    rng.setNumberFormat('¥#,##0;[Red]-¥#,##0');
  }

}

/***** ユーティリティ *****/
function getSheetOrThrow_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`シート「${name}」が見つかりません。`);
  return sh;
}

function findHeaderIndex_(headers, candidates) {
  const canon = s => String(s).trim().toLowerCase();
  const set = new Set(headers.map(canon));
  for (const cand of candidates) {
    const i = headers.findIndex(h => canon(h) === canon(cand));
    if (i !== -1) return i;
  }
  // ゆるめ：含むパターンでも探す
  for (let i = 0; i < headers.length; i++) {
    const h = canon(headers[i]);
    for (const cand of candidates) {
      const c = canon(cand);
      if (h.includes(c)) return i;
    }
  }
  return -1;
}

function toYYYYMM_(value) {
  // Date型・シリアル・文字列(YYYY/MM, YYYY-MM, YYYY年MM月 など)に対応
  let d;
  if (value instanceof Date) {
    d = value;
  } else if (typeof value === 'number') {
    d = new Date(Math.round((value - 25569) * 86400 * 1000)); // Excel/Sheetsシリアル対応
  } else if (typeof value === 'string') {
    const s = value.trim();
    // 例: 2025/08, 2025-8, 2025年8月
    const m = s.match(/(\d{4})[\/\-年\. ]\s*(\d{1,2})/);
    if (m) {
      const y = Number(m[1]);
      const mo = Number(m[2]) - 1;
      d = new Date(y, mo, 1);
    } else {
      // 日付文字列（2025/8/3 等）
      const t = new Date(s);
      if (!isNaN(t)) d = t;
    }
  }
  if (!d || isNaN(d)) return '';
  const y = d.getFullYear();
  const m = d.getMonth() + 1;
  return `${y}/${('0' + m).slice(-2)}`;
}

function normalizeYYYYMMCell_(cell) {
  // サマリ側の年月セルを YYYY/MM に正規化
  if (cell instanceof Date || typeof cell === 'number') {
    return toYYYYMM_(cell);
  }
  if (typeof cell === 'string') {
    const s = cell.trim();
    // 既に YYYY/MM ならそのまま
    if (/^\d{4}\/\d{2}$/.test(s)) return s;
    return toYYYYMM_(s);
  }
  return '';
}

function toNumberOrZero_(v) {
  if (typeof v === 'number') return v || 0;
  if (typeof v === 'string') {
    const n = Number(v.replace(/[,¥\s]/g, ''));
    return isNaN(n) ? 0 : n;
  }
  return 0;
}
