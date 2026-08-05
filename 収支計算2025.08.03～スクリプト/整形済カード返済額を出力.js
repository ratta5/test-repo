function 整形済カード返済額を出力() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("カード返済額");
  const outputSheetName = "カード返済額_整形";
  let outputSheet = ss.getSheetByName(outputSheetName);

  if (!outputSheet) {
    outputSheet = ss.insertSheet(outputSheetName);
  } else {
    outputSheet.clearContents();
  }

  const data = sourceSheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const resultMap = new Map();

  // 「当月」「次月」列を抽出（index, カード名）
  const atsumList = headers.map((h, i) => ({ index: i, name: h.replace('当月', '').trim() }))
    .filter(h => headers[h.index].includes('当月'));

  const nextList = headers.map((h, i) => ({ index: i, name: h.replace('次月', '').trim() }))
    .filter(h => headers[h.index].includes('次月'));

  // 各行ごとに「当月」と「次月」処理
  rows.forEach(row => {
    const timestamp = new Date(row[0]);
    if (isNaN(timestamp)) return;

    const ymThis = `${timestamp.getFullYear()}-${String(timestamp.getMonth() + 1).padStart(2, '0')}`;
    const nextDate = new Date(timestamp);
    nextDate.setMonth(nextDate.getMonth() + 1);
    const ymNext = `${nextDate.getFullYear()}-${String(nextDate.getMonth() + 1).padStart(2, '0')}`;

    // 当月処理
    atsumList.forEach(h => {
      const val = row[h.index];
      if (!resultMap.has(ymThis)) resultMap.set(ymThis, {});
      const target = resultMap.get(ymThis);
      if (!(h.name in target) || new Date(target[h.name].timestamp) < timestamp) {
        target[h.name] = { value: isNaN(val) || val === '' ? 0 : Number(val), timestamp };
      }
    });

    // 次月処理
    nextList.forEach(h => {
      const val = row[h.index];
      if (!resultMap.has(ymNext)) resultMap.set(ymNext, {});
      const target = resultMap.get(ymNext);
      if (!(h.name in target) || new Date(target[h.name].timestamp) < timestamp) {
        target[h.name] = { value: isNaN(val) || val === '' ? 0 : Number(val), timestamp };
      }
    });
  });

  // 出力整形
  const allCardNames = [...new Set([...atsumList, ...nextList].map(h => h.name))];
  const sortedKeys = [...resultMap.keys()].sort();
  const output = [['年月', '合計', ...allCardNames]];

  sortedKeys.forEach(ym => {
    const monthData = resultMap.get(ym);
    const values = allCardNames.map(name => (monthData[name]?.value ?? 0));
    const total = values.reduce((a, b) => a + b, 0);
    const formattedValues = values.map(v => `¥${v.toLocaleString()}`);
    output.push([ym, `¥${total.toLocaleString()}`, ...formattedValues]);
  });

  outputSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}
