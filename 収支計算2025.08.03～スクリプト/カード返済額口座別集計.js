function カード返済額口座別集計() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paymentSheet = ss.getSheetByName("カード返済額_整形");
  const masterSheet = ss.getSheetByName("カードマスター");
  const outputSheetName = "口座別カード返済集計";
  let outputSheet = ss.getSheetByName(outputSheetName);

  // 出力シートが存在すればクリア、なければ作成
  if (outputSheet) {
    outputSheet.clearContents();
  } else {
    outputSheet = ss.insertSheet(outputSheetName);
  }

  const paymentData = paymentSheet.getDataRange().getValues();
  const masterData = masterSheet.getDataRange().getValues();

  const headers = paymentData[0];
  const dataRows = paymentData.slice(1);

  // カード名と引落口座のマッピング
  const cardToBank = {};
  for (let i = 1; i < masterData.length; i++) {
    const card = masterData[i][1];
    const bank = masterData[i][2];
    if (card && bank) {
      cardToBank[card] = bank;
    }
  }

  // 口座ごとに集計対象カードを分類
  const bankToCards = {};
  headers.forEach((header, colIndex) => {
    if (colIndex >= 2 && cardToBank[header]) {
      const bank = cardToBank[header];
      if (!bankToCards[bank]) bankToCards[bank] = [];
      bankToCards[bank].push(colIndex);
    }
  });

  // 出力用データ構築
  const output = [];
  const outputHeader = ["年月", ...Object.keys(bankToCards)];
  output.push(outputHeader);

  dataRows.forEach(row => {
    const date = Utilities.formatDate(new Date(row[0]), "Asia/Tokyo", "yyyy-MM");
    const line = [date];
    for (const bank of Object.keys(bankToCards)) {
      const sum = bankToCards[bank].reduce((acc, colIndex) => {
        const val = row[colIndex];
        if (typeof val === "number") {
          return acc + val;
        } else if (typeof val === "string") {
          const num = parseFloat(val.replace(/[¥,]/g, ""));
          return acc + (isNaN(num) ? 0 : num);
        } else {
          return acc;
        }
      }, 0);
      line.push(sum);
    }
    output.push(line);
  });

  // 出力
  outputSheet.getRange(1, 1, output.length, output[0].length).setValues(output);

  // 通貨形式に設定
  const lastCol = output[0].length;
  if (output.length > 1) {
    outputSheet.getRange(2, 2, output.length - 1, lastCol - 1)
      .setNumberFormat('"¥"#,##0');
  }
}
