/**
 * Googleフォームから取得した複数選択（カンマ区切り）の出金データを、
 * 1機種1行ずつの形式に分解して「TikTok出金ﾃﾞｰﾀ」シートに同期します。
 */
function 出金データ同期() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("TikTok出金ﾃﾞｰﾀ_フォーム回答");
  const cleanSheet = ss.getSheetByName("TikTok出金ﾃﾞｰﾀ");
  
  if (!rawSheet) {
    throw new Error("シート『TikTok出金ﾃﾞｰﾀ_フォーム回答』が見つかりません。現在のフォーム回答シートの名前を『TikTok出金ﾃﾞｰﾀ_フォーム回答』に変更してください。");
  }
  if (!cleanSheet) {
    throw new Error("シート『TikTok出金ﾃﾞｰﾀ』が見つかりません。新しく『TikTok出金ﾃﾞｰﾀ』という名前のシートを作成してください。");
  }
  
  const rawData = rawSheet.getDataRange().getValues();
  if (rawData.length < 2) {
    cleanSheet.clear();
    cleanSheet.appendRow(["タイムスタンプ", "日付", "機種", "金額", "出金年月"]); // ヘッダー
    return;
  }
  
  const headers = ["タイムスタンプ", "日付", "機種", "金額", "出金年月"];
  const output = [headers]; // ヘッダーを維持
  
  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    const timestamp = row[0];
    const date = row[1];
    const devicesRaw = String(row[2]);
    const amount = row[3];
    
    if (!devicesRaw) continue;
    
    // 出金年月を取得（日付からyyyy/MM形式を作成。日付が無効の場合はタイムスタンプからフォールバック）
    const yearMonth = formatYearMonth(date, timestamp);
    
    // カンマ（半角・全角）、読点、改行などで分割してトリミング
    const devices = devicesRaw.split(/[,，、\n]+/).map(d => d.trim()).filter(Boolean);
    
    devices.forEach(device => {
      // 1機種ごとに独立した行を作成（金額はそのまま適用）
      output.push([timestamp, date, device, amount, yearMonth]);
    });
  }
  
  cleanSheet.clear();
  cleanSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}

/**
 * 日付（またはタイムスタンプ）から「yyyy/MM」形式の文字列を生成します。
 */
function formatYearMonth(dateVal, timestampVal) {
  let d = null;
  if (dateVal) {
    if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
      d = dateVal;
    } else {
      const str = String(dateVal).replace(/\./g, '/').trim();
      const parsed = new Date(str);
      if (!isNaN(parsed.getTime())) {
        d = parsed;
      }
    }
  }
  if (!d && timestampVal) {
    if (timestampVal instanceof Date && !isNaN(timestampVal.getTime())) {
      d = timestampVal;
    } else {
      const str = String(timestampVal).replace(/\./g, '/').trim();
      const parsed = new Date(str);
      if (!isNaN(parsed.getTime())) {
        d = parsed;
      }
    }
  }
  return d ? Utilities.formatDate(d, "Asia/Tokyo", "yyyy/MM") : "";
}
