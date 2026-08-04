/***** エントリーポイント：7関数を順次実行 *****/
function すべて実行_月次収支サマリ更新() {
  const 開始時刻 = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();

  // 30秒以内にロック取得（同時実行防止）
  if (!lock.tryLock(30 * 1000)) {
    SpreadsheetApp.getActive().toast(
      '別の処理が動作中のため実行できませんでした。少し待って再度お試しください。',
      '月次収支サマリ',
      5
    );
    throw new Error('実行ロックを取得できませんでした。');
  }

  ss.toast('更新を開始します…', '月次収支サマリ（まとめ実行）', 5);

  /***** 実行ステップ定義 *****/
  const ステップ = [
    { 表示名: '収入_固定 → 月次収支サマリ', fn: 出力_収入固定_月次収支サマリ },
    { 表示名: '収入_変動 → 月次収支サマリ', fn: 収入_変動を月次収支サマリへ反映 },
    { 表示名: '支出_固定費 更新',           fn: 更新_支出固定費 },
    { 表示名: '支出_見込み費用 反映',        fn: 支出見込み費用を反映 },
    { 表示名: '変動収入（物販）反映',        fn: 出力_変動収入_物販_反映 },
    { 表示名: '変動支出（実績）反映',        fn: 出力_変動支出_実績_反映 },
    { 表示名: '変動収入（月次収支サマリ強制）', fn: 出力_変動収入_月次収支サマリ_強制 },
  ];

  const 結果一覧 = [];

  try {
    // 各ステップを順番に安全実行
    ステップ.forEach(step => {
      const r = 安全実行_(step.表示名, step.fn);
      結果一覧.push(r);
      ss.toast(`${step.表示名}: ${r.ok ? '完了' : '失敗'}（${r.ms}ms）`, '進捗', 5);
    });
  } finally {
    // ロックを必ず解放
    lock.releaseLock();
  }

  /***** 実行結果まとめ *****/
  const 成功数 = 結果一覧.filter(r => r.ok).length;
  const 失敗数 = 結果一覧.length - 成功数;
  const 合計時間ms = new Date() - 開始時刻;

  Logger.log('―― 月次収支サマリ（まとめ実行）結果 ――');
  結果一覧.forEach(r => 
    Logger.log(`${r.表示名}: ${r.ok ? 'OK' : 'NG'} (${r.ms}ms) ${r.ok ? '' : '｜原因: ' + r.errMsg}`)
  );
  Logger.log(`合計: 成功 ${成功数} / 失敗 ${失敗数} ｜ 所要時間: ${合計時間ms}ms`);

  const 最終メッセージ =
    `完了: 成功 ${成功数} / 失敗 ${失敗数}（${Math.round(合計時間ms / 1000)}秒）` +
    (失敗数 > 0 ? '｜詳細は実行ログ（View → Logs）を参照' : '');
  ss.toast(最終メッセージ, '月次収支サマリ（まとめ実行）', 8);
}

/***** 個別関数を安全に実行（計測＆エラーハンドリング） *****/
function 安全実行_(表示名, fn) {
  const t0 = new Date();
  try {
    fn(); // 関数を実行
    const ms = new Date() - t0;
    return { 表示名, ok: true, ms };
  } catch (e) {
    const ms = new Date() - t0;
    Logger.log(`✖ ${表示名} でエラー: ${e && e.stack ? e.stack : e}`);
    return { 表示名, ok: false, ms, errMsg: String(e) };
  }
}

/***** スプレッドシートのメニューに追加 *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('月次収支サマリ')
    .addItem('すべて実行', 'すべて実行_月次収支サマリ更新')
    .addToUi();
}
