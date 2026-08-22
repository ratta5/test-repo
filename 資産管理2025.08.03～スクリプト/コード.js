/**
 * スプレッドシート起動時に「資産管理メニュー」を追加するカスタムメニュー関数
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("資産管理メニュー")
    .addItem("1. 未照合の販売記録を出力", "未照合販売記録を出力")
    .addItem("2. メモしたマスタNoを照合表へ転記", "未照合結果を照合表へ転記")
    .addItem("3. 重複排除後の統合販売記録を出力", "統合販売記録を出力")
    .addSeparator()
    .addItem("全処理を実行", "全処理を実行")
    .addToUi();
}

