/**
 * ===================================================================
 * 仕入試算＆検討履歴データベース 自動管理スクリプト
 * 
 * 【仕入送料 ＆ ポイント利用列 追加版】
 * ・「仕入送料（購入時送料）」と「ポイント利用」の列を新設
 * ・「支払税込額 ＝ 税込定価 － クーポン ＋ 仕入送料」を正確に算出
 * ・楽天ポイント獲得計算（税抜基準・クーポン引後）と実質原価を完全連動
 * ===================================================================
 */

const SHEET_NAME_CALC = '仕入試算';

/**
 * シート初期構築
 */
function 仕入試算シート作成() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME_CALC);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_CALC);
  } else {
    sheet.clear();
    sheet.clearFormats();
    sheet.clearConditionalFormatRules();
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
  }

  // -----------------------------------------------------------------
  // 1. マスタデータ作成 (AH列〜AI列)
  // -----------------------------------------------------------------
  sheet.getRange('AH1:AI1').setValues([['発送方法マスタ', '送料(円)']])
    .setBackground('#4a86e8').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  const masterData = [
    ['ネコポス', 210],
    ['ゆうパケットポスト', 215],
    ['宅急便コンパクト', 520],
    ['ゆうパケットプラス', 520],
    ['宅急便 60サイズ', 750],
    ['宅急便 80サイズ', 850],
    ['宅急便 100サイズ', 1050],
    ['普通郵便 / 定形外', 140],
    ['その他 (手入力)', 0]
  ];
  sheet.getRange('AH2:AI10').setValues(masterData);
  sheet.getRange('AI2:AI10').setNumberFormat('#,##0');
  sheet.getRange('AH1:AI10').setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);

  // -----------------------------------------------------------------
  // 2. 上部：イベント設定 ＆ 日付別臨時イベント(9日分) ＆ 全体サマリー (1〜6行目)
  // -----------------------------------------------------------------
  // タイトル帯
  sheet.getRange('A1:D1').merge().setValue('【 楽天・基本＆買い回り設定 】')
    .setBackground('#cc0000').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('E1:M1').merge().setValue('【 日付別・臨時イベント計算テーブル（9日分） 】── 楽天仕入限定')
    .setBackground('#e69138').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('N1:S1').merge().setValue('【 今回のマラソン仕入集計・上限チェック（仕入のみ） 】')
    .setBackground('#4a148c').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  // 左ブロック：基本＆買い回り設定
  const baseDate = new Date();
  const endDate = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate() + 8);

  sheet.getRange('A2:D6').setValues([
    ['全体SPU(モバイル等)', 0.060, '目標利益率(○)', 0.15],
    ['仕入専用SPU(でんき等)', 0.005, '許容利益率(△)', 0.05],
    ['マラソン開始日', baseDate, 'マラソン終了日', endDate],
    ['仕入店舗数(自動)', '=COUNTA(FILTER(D10:D19, A10:A19="仕入", B10:B19="楽天", D10:D19<>""))', '私物/家計店舗(自動)', '=COUNTA(FILTER(D10:D19, (A10:A19="私物")+(A10:A19="家計"), B10:B19="楽天", D10:D19<>""))'],
    ['買い回り総店舗数', '=MIN(B5 + D5, 10)', '買い回り付与率', '=MIN(MAX(B6-1, 0), 9)/100']
  ]);

  // 中央ブロック：日付別・臨時イベント計算テーブル（9日分）
  sheet.getRange('E2:M2').setValues([[
    '=$B$4', '=$E$2+1', '=$E$2+2', '=$E$2+3', '=$E$2+4',
    '=$E$2+5', '=$E$2+6', '=$E$2+7', '=$E$2+8'
  ]]);

  sheet.getRange('E3:M3').setValues([[0.01, 0.02, 0.00, 0.01, 0.01, 0.00, 0.00, 0.00, 0.00]]); // 勝ったら倍/5の日
  sheet.getRange('E4:M4').setValues([[0.00, 0.03, 0.00, 0.01, 0.00, 0.00, 0.00, 0.00, 0.00]]); // 39ショップ/変倍等
  sheet.getRange('E5:M5').setValues([[0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00]]); // その他自由枠
  sheet.getRange('E6:M6').setValues([[
    '=SUM(E3:E5)', '=SUM(F3:F5)', '=SUM(G3:G5)', '=SUM(H3:H5)', '=SUM(I3:I5)',
    '=SUM(J3:J5)', '=SUM(K3:K5)', '=SUM(L3:L5)', '=SUM(M3:M5)'
  ]]);

  // 右ブロック：マラソン全体サマリー（仕入区分のみを集計）
  sheet.getRange('N2:S4').setValues([
    ['総仕入支払額', '=SUMIF(A10:A19, "仕入", K10:K19)', '総獲得ポイント', '=SUMIF(A10:A19, "仕入", S10:S19)', '実質仕入原価計', '=SUMIF(A10:A19, "仕入", T10:T19)'],
    ['総売上想定額', '=SUMIF(A10:A19, "仕入", U10:U19)', '総手残り利益額', '=SUMIF(A10:A19, "仕入", Z10:Z19)', '全体利益率', '=IF(O3>0, Q3/O3, "")'],
    ['マラソン上限枠', 5000, 'マラソン上限残枠', '=MAX(O4 - IFERROR(SUMPRODUCT((A10:A19="仕入") * (B10:B19="楽天") * R10:R19 * $D$6), 0), 0)', '上限判定', '=IF(Q4=0, "⚠️ 上限到達", "OK")']
  ]);

  // 上部書式設定
  sheet.getRange('B2:B3').setNumberFormat('0.0%');
  sheet.getRange('D2:D3').setNumberFormat('0.0%');
  sheet.getRange('B4:D4').setNumberFormat('yyyy/MM/dd').setBackground('#fff2cc').setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange('B5').setNumberFormat('#,##0" 店"');
  sheet.getRange('D5').setNumberFormat('#,##0" 店"');
  sheet.getRange('B6').setNumberFormat('#,##0" 店"').setFontWeight('bold');
  sheet.getRange('D6').setNumberFormat('+0.0%').setFontWeight('bold');

  // 日付別テーブル書式
  sheet.getRange('E2:M2').setNumberFormat('m/d(aaa)').setBackground('#cfe2f3').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('E3:M5').setNumberFormat('0.0%');
  sheet.getRange('E6:M6').setNumberFormat('+0.0%').setBackground('#fce5cd').setFontWeight('bold').setHorizontalAlignment('center');

  // サマリー書式
  sheet.getRange('O2:O4').setNumberFormat('#,##0" 円"');
  sheet.getRange('Q2:Q3').setNumberFormat('#,##0" pt"');
  sheet.getRange('S2').setNumberFormat('#,##0" 円"');
  sheet.getRange('S3').setNumberFormat('0.0%').setFontWeight('bold');
  sheet.getRange('Q4').setNumberFormat('#,##0" pt"').setFontWeight('bold');
  sheet.getRange('S4').setHorizontalAlignment('center').setFontWeight('bold');

  sheet.getRange('A2:A6').setBackground('#f3f3f3').setFontWeight('bold');
  sheet.getRange('C2:C6').setBackground('#f3f3f3').setFontWeight('bold');
  sheet.getRange('N2:N4').setBackground('#f3f3f3').setFontWeight('bold');
  sheet.getRange('P2:P4').setBackground('#f3f3f3').setFontWeight('bold');
  sheet.getRange('R2:R4').setBackground('#f3f3f3').setFontWeight('bold');

  sheet.getRange('A1:D6').setBorder(true, true, true, true, true, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('E1:M6').setBorder(true, true, true, true, true, true, '#e69138', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('N1:S4').setBorder(true, true, true, true, true, true, '#4a148c', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // -----------------------------------------------------------------
  // 3. 複数店舗シミュレーター（8〜19行目：10店舗分）
  // -----------------------------------------------------------------
  sheet.getRange('A8:AC8').merge().setValue('【 お買い物マラソン・複数店舗一括シミュレーター（仕入送料＆ポイント利用対応） 】')
    .setBackground('#0b5394').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  const simHeaders = [
    '購入区分', '仕入先区分', '検討日', '店舗/ショップ名', 'JANコード', '商品名 / 型番',
    '税込定価', 'クーポン', '仕入送料', 'ポイント利用', '支払税込額', '個別P%', '全体SPU%', '仕入専用SPU%', 'マラソンP%', '当日臨時P(自動)', '適用還元率',
    '税抜仕入額', '獲得P', '実質仕入原価', 'メルカリ売価', '発送方法',
    '販売送料', '梱包費', '販売手数料', '予想利益額', '予想利益率', '損益分岐売価', '仕入判定'
  ];
  sheet.getRange(9, 1, 1, simHeaders.length).setValues([simHeaders])
    .setBackground('#134f5c').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  // サンプルデータ（全29列）
  const sampleSimRows = [
    ['仕入', '楽天', '=E2+1', '楽天ビック', '4549995000001', 'サンプル商品A (送料無料・P利用500)', 12000, 1000, 0, 500, '', 0.09, '', '', '', '', '', '', '', '', 13500, '宅急便コンパクト', '', 30, '', '', '', '', ''],
    ['仕入', '楽天', '=E2+3', 'エディオン 楽天市場店', '4549995000002', 'サンプル商品B (仕入送料550円あり)', 8000, 0, 550, 0, '', 0.05, '', '', '', '', '', '', '', '', 9800, 'ネコポス', '', 30, '', '', '', '', ''],
    ['私物', '楽天', '=E2', '楽天ブックス', '', '私物の本 (店舗数稼ぎ)', 1500, 0, 0, 0, '', 0.00, '', '', '', '', '', '', '', '', '', '', '', 0, '', '', '', '', ''],
    ['家計', '楽天', '=E2', 'コンタクトレンズ通販', '', '日用品消耗品 (店舗数稼ぎ)', 3000, 200, 0, 0, '', 0.00, '', '', '', '', '', '', '', '', '', '', '', 0, '', '', '', '', ''],
    ['仕入', '店舗', '=E2', 'ヤマダ電機 新宿店', '4549995000010', '店舗仕入サンプル', 5000, 0, 0, 0, '', 0.01, '', '', '', '', '', '', '', '', 8000, 'ネコポス', '', 30, '', '', '', '', ''],
    ['仕入', '楽天', '=E2', '', '', '', '', '', 0, 0, '', 0, '', '', '', '', '', '', '', '', '', 'ネコポス', '', 30, '', '', '', '', ''],
    ['仕入', '楽天', '=E2', '', '', '', '', '', 0, 0, '', 0, '', '', '', '', '', '', '', '', '', 'ネコポス', '', 30, '', '', '', '', ''],
    ['仕入', '楽天', '=E2', '', '', '', '', '', 0, 0, '', 0, '', '', '', '', '', '', '', '', '', 'ネコポス', '', 30, '', '', '', '', ''],
    ['仕入', '楽天', '=E2', '', '', '', '', '', 0, 0, '', 0, '', '', '', '', '', '', '', '', '', 'ネコポス', '', 30, '', '', '', '', ''],
    ['仕入', '楽天', '=E2', '', '', '', '', '', 0, 0, '', 0, '', '', '', '', '', '', '', '', '', 'ネコポス', '', 30, '', '', '', '', '']
  ];
  sheet.getRange(10, 1, 10, simHeaders.length).setValues(sampleSimRows);

  // 数式を10行目〜19行目にセット
  for (let r = 10; r <= 19; r++) {
    // K列: 支払税込額（税込定価 G － クーポン H ＋ 仕入送料 I）
    sheet.getRange(`K${r}`).setFormula(`=IF(G${r}="","", G${r} - N(H${r}) + N(I${r}))`);
    // M列: 全体SPU%（楽天なら上部B2の6.0%、店舗は0%）
    sheet.getRange(`M${r}`).setFormula(`=IF(B${r}="楽天", $B$2, 0)`);
    // N列: 仕入専用SPU%（★「仕入」かつ「楽天」のみ上部B3の0.5%、私物・家計・店舗は0%）
    sheet.getRange(`N${r}`).setFormula(`=IF(AND(A${r}="仕入", B${r}="楽天"), $B$3, 0)`);
    // O列: マラソンP%（★「仕入」かつ「楽天」のみ付与率を適用、私物・家計・店舗は0%）
    sheet.getRange(`O${r}`).setFormula(`=IF(AND(A${r}="仕入", B${r}="楽天"), $D$6, 0)`);
    // P列: 当日臨時P（★楽天仕入のみ上部テーブルからLookup、その他は0%）
    sheet.getRange(`P${r}`).setFormula(`=IF(B${r}="楽天", IFERROR(HLOOKUP(C${r}, $E$2:$M$6, 5, FALSE), 0), 0)`);
    // Q列: 適用還元率（個別P L ＋ 全体SPU M ＋ 仕入専用SPU N ＋ マラソンP O ＋ 当日臨時P P）
    sheet.getRange(`Q${r}`).setFormula(`=IF(B${r}="楽天", L${r} + M${r} + N${r} + O${r} + P${r}, L${r})`);
    // R列: 税抜仕入額（定価 G － クーポン H を1.1で割る ※送料除く商品代金ベース）
    sheet.getRange(`R${r}`).setFormula(`=IF(G${r}="","", ROUNDDOWN((G${r} - N(H${r})) / 1.1))`);
    // S列: 獲得P
    sheet.getRange(`S${r}`).setFormula(`=IF(R${r}="","", ROUNDDOWN(R${r} * Q${r}))`);
    // T列: 実質仕入原価（支払税込額 K － 獲得P S）
    sheet.getRange(`T${r}`).setFormula(`=IF(K${r}="","", K${r} - S${r})`);
    // W列: 販売送料（仕入のみ計算）
    sheet.getRange(`W${r}`).setFormula(`=IF(A${r}="仕入", IFERROR(VLOOKUP(V${r}, $AH$2:$AI$10, 2, FALSE), 0), "-")`);
    // Y列: 販売手数料（仕入のみ計算）
    sheet.getRange(`Y${r}`).setFormula(`=IF(AND(A${r}="仕入", U${r}<>""), ROUNDDOWN(U${r} * 0.1), "-")`);
    // Z列: 予想利益額（仕入のみ計算: 売価 U － 手数料 Y － 販売送料 W － 梱包費 X － 実質原価 T）
    sheet.getRange(`Z${r}`).setFormula(`=IF(AND(A${r}="仕入", U${r}<>""), U${r} - Y${r} - W${r} - X${r} - T${r}, "-")`);
    // AA列: 予想利益率（仕入のみ計算）
    sheet.getRange(`AA${r}`).setFormula(`=IF(AND(A${r}="仕入", U${r}<>""), Z${r} / U${r}, "-")`);
    // AB列: 損益分岐売価（仕入のみ計算）
    sheet.getRange(`AB${r}`).setFormula(`=IF(AND(A${r}="仕入", T${r}<>""), ROUNDUP((T${r} + W${r} + X${r}) / 0.9), "-")`);
    // AC列: 仕入判定
    sheet.getRange(`AC${r}`).setFormula(`=IF(A${r}="仕入", IF(AA${r}="","", IF(AA${r}>=$D$2, "○ 仕入推奨", IF(AA${r}>=0.05, "△ 要検討", "× 見送り"))), "─ (私物/家計)")`);
  }

  // シミュレーター書式設定
  sheet.getRange('A10:B19').setHorizontalAlignment('center');
  sheet.getRange('A10:A19').setFontWeight('bold');
  sheet.getRange('C10:C19').setNumberFormat('yyyy/MM/dd').setHorizontalAlignment('center');
  sheet.getRange('E10:E19').setNumberFormat('@');
  sheet.getRange('G10:K19').setNumberFormat('#,##0');
  sheet.getRange('L10:Q19').setNumberFormat('0.0%');
  sheet.getRange('M10:M19').setBackground('#cfe2f3').setFontWeight('bold'); // 全体SPU%
  sheet.getRange('N10:N19').setBackground('#d0e0e3').setFontWeight('bold'); // 仕入専用SPU%
  sheet.getRange('O10:O19').setBackground('#d9ead3').setFontWeight('bold'); // マラソンP%
  sheet.getRange('P10:P19').setBackground('#fce5cd').setFontWeight('bold'); // 当日臨時P%
  sheet.getRange('R10:U19').setNumberFormat('#,##0');
  sheet.getRange('W10:Z19').setNumberFormat('#,##0');
  sheet.getRange('AA10:AA19').setNumberFormat('0.0%').setFontWeight('bold');
  sheet.getRange('AB10:AB19').setNumberFormat('#,##0');
  sheet.getRange('AC10:AC19').setHorizontalAlignment('center');

  sheet.getRange('K10:K19').setBackground('#fff2cc');
  sheet.getRange('Q10:T19').setBackground('#e8f0fe');
  sheet.getRange('Z10:AC19').setBackground('#f3f3f3');
  sheet.getRange('A9:AC19').setBorder(true, true, true, true, true, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);

  // -----------------------------------------------------------------
  // 4. ガイドバー (21行目)
  // -----------------------------------------------------------------
  sheet.getRange('A21:AC21').merge().setValue('👉 検討が終わったら、メニュー「仕入・商品管理」→「★シミュレーター全商品を履歴に一括保存」を押すと下に値として蓄積されます')
    .setBackground('#fce5cd').setFontColor('#b45f06').setFontWeight('bold').setHorizontalAlignment('center');

  // -----------------------------------------------------------------
  // 5. 検討履歴データベース（23行目〜）
  // -----------------------------------------------------------------
  sheet.getRange('A23:AE23').merge().setValue('【 検討履歴データベース（何ヶ月分でも値として蓄積） 】')
    .setBackground('#1c4587').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  const historyHeaders = [
    'No.', '購入区分', '仕入先区分', '検討日', 'マラソン期間', '店舗/ショップ名', 'JANコード',
    '商品名 / 型番', '税込定価', 'クーポン', '仕入送料', 'ポイント利用', '支払税込額',
    '個別P%', '全体SPU%', '仕入専用SPU%', 'マラソンP%', '当日臨時P%', '適用還元率',
    '還元内訳メモ', '獲得P', '実質仕入原価', 'メルカリ売価', '発送方法', '販売送料', '梱包費', '販売手数料', '予想利益額', '予想利益率', '損益分岐売価', '仕入判定'
  ];
  sheet.getRange(24, 1, 1, historyHeaders.length).setValues([historyHeaders])
    .setBackground('#3c78d8').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  sheet.getRange(24, 1, 1, historyHeaders.length).setBorder(true, true, true, true, true, true, '#1155cc', SpreadsheetApp.BorderStyle.SOLID);

  // -----------------------------------------------------------------
  // 6. 入力規則（プルダウン）
  // -----------------------------------------------------------------
  sheet.getRange('A10:A19').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['仕入', '私物', '家計'], true).build()
  );

  sheet.getRange('B10:B19').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['楽天', '店舗', 'その他ネット'], true).build()
  );

  sheet.getRange('V10:V19').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(sheet.getRange('AH2:AH10'), true).build()
  );

  // -----------------------------------------------------------------
  // 7. 条件付き書式（仕入判定の色分け & 購入区分の色分け）
  // -----------------------------------------------------------------
  const ruleGood = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('○ 仕入推奨')
    .setBackground('#d9ead3').setFontColor('#274e13').setBold(true)
    .setRanges([sheet.getRange('AC10:AC19'), sheet.getRange('AE25:AE1000')])
    .build();

  const ruleWarning = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('△ 要検討')
    .setBackground('#fff2cc').setFontColor('#7f6000').setBold(true)
    .setRanges([sheet.getRange('AC10:AC19'), sheet.getRange('AE25:AE1000')])
    .build();

  const ruleBad = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('× 見送り')
    .setBackground('#f4cccc').setFontColor('#990000').setBold(true)
    .setRanges([sheet.getRange('AC10:AC19'), sheet.getRange('AE25:AE1000')])
    .build();

  const rulePrivate = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('─ (私物/家計)')
    .setBackground('#eeeeee').setFontColor('#888888')
    .setRanges([sheet.getRange('AC10:AC19'), sheet.getRange('AE25:AE1000')])
    .build();

  sheet.setConditionalFormatRules([ruleGood, ruleWarning, ruleBad, rulePrivate]);

  // -----------------------------------------------------------------
  // 8. 列幅の最適化
  // -----------------------------------------------------------------
  sheet.setColumnWidth(1, 75);  // A: 購入区分
  sheet.setColumnWidth(2, 85);  // B: 仕入先区分
  sheet.setColumnWidth(3, 95);  // C: 検討日
  sheet.setColumnWidth(4, 140); // D: 店舗名
  sheet.setColumnWidth(5, 120); // E: JAN
  sheet.setColumnWidth(6, 180); // F: 商品名
  sheet.setColumnWidth(7, 95);  // G: 税込定価
  sheet.setColumnWidth(8, 85);  // H: クーポン
  sheet.setColumnWidth(9, 75);  // I: 仕入送料(新設)
  sheet.setColumnWidth(10, 85); // J: ポイント利用(新設)
  sheet.setColumnWidth(11, 95); // K: 支払税込
  sheet.setColumnWidth(12, 70); // L: 個別P
  sheet.setColumnWidth(13, 80); // M: 全体SPU
  sheet.setColumnWidth(14, 85); // N: 仕入専用SPU
  sheet.setColumnWidth(15, 85); // O: マラソンP
  sheet.setColumnWidth(16, 90); // P: 当日臨時P
  sheet.setColumnWidth(17, 85); // Q: 還元率
  sheet.setColumnWidth(18, 85); // R: 税抜
  sheet.setColumnWidth(19, 75); // S: 獲得P
  sheet.setColumnWidth(20, 95); // T: 実質原価
  sheet.setColumnWidth(21, 100);// U: 売価
  sheet.setColumnWidth(22, 130);// V: 発送方法
  sheet.setColumnWidth(23, 65); // W: 販売送料
  sheet.setColumnWidth(24, 65); // X: 梱包費
  sheet.setColumnWidth(25, 80); // Y: 手数料
  sheet.setColumnWidth(26, 95); // Z: 利益額
  sheet.setColumnWidth(27, 75); // AA: 利益率
  sheet.setColumnWidth(28, 95); // AB: 損益分岐
  sheet.setColumnWidth(29, 100);// AC: 判定
  sheet.setColumnWidth(30, 110);// AD: 期間(履歴用)
  sheet.setColumnWidth(31, 100);// AE: 判定(履歴用)

  SpreadsheetApp.getActive().toast('仕入送料＆ポイント利用列を追加したシートを作成しました！', '完了', 5);
}

/**
 * ===================================================================
 * ★ シミュレーターに入力された全商品を履歴データベースに一括保存
 * ===================================================================
 */
function シミュレーター全商品を履歴に一括保存() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_CALC);
  if (!sheet) return;

  // 10行目〜19行目のデータを一括取得 (A〜AC列: 29列)
  const simRange = sheet.getRange('A10:AC19').getValues();
  const rowsToSave = [];

  const startD = sheet.getRange('B4').getValue();
  const endD = sheet.getRange('D4').getValue();
  let periodText = '';
  if (startD instanceof Date && endD instanceof Date) {
    periodText = `${Utilities.formatDate(startD, 'Asia/Tokyo', 'MM/dd')}〜${Utilities.formatDate(endD, 'Asia/Tokyo', 'MM/dd')}`;
  } else {
    periodText = `${startD}〜${endD}`;
  }

  for (let i = 0; i < simRange.length; i++) {
    const row = simRange[i];
    const buyCategory    = row[0];  // A: 購入区分
    const typeVal        = row[1];  // B: 仕入先区分
    const dateVal        = row[2];  // C: 検討日
    const shopVal        = row[3];  // D: 店舗名
    const janVal         = row[4];  // E: JAN
    const nameVal        = row[5];  // F: 商品名
    const priceVal       = row[6];  // G: 税込定価
    const couponVal      = row[7] || 0; // H: クーポン
    const inShippingVal  = row[8] || 0; // I: 仕入送料
    const pointPayVal    = row[9] || 0; // J: ポイント利用
    const payPriceVal    = row[10]; // K: 支払税込
    const indPointVal    = row[11] || 0; // L: 個別P
    const allSpuVal      = row[12] || 0; // M: 全体SPU
    const buySpuVal      = row[13] || 0; // N: 仕入専用SPU
    const maraPointVal   = row[14] || 0; // O: マラソンP
    const extraPointVal  = row[15] || 0; // P: 当日臨時P
    const rateVal        = row[16]; // Q: 適用還元率
    const getPointVal    = row[18]; // S: 獲得P
    const realCostVal    = row[19]; // T: 実質原価
    const targetSellVal  = row[20]; // U: 売価
    const shippingVal    = row[21]; // V: 発送方法
    const shipCostVal    = row[22]; // W: 販売送料
    const packingVal     = row[23] || 0; // X: 梱包費
    const feeVal         = row[24]; // Y: 手数料
    const profitVal      = row[25]; // Z: 予想利益額
    const profitRateVal  = row[26]; // AA: 予想利益率
    const breakEvenVal   = row[27]; // AB: 損益分岐
    const judgeVal       = row[28]; // AC: 判定

    // 定価が空欄の行はスキップ
    if (!priceVal) continue;

    let periodForSave = '';
    let detailNote = '';

    if (typeVal === '楽天') {
      periodForSave = periodText;
      const ind = (Number(indPointVal) * 100).toFixed(1);
      const allSpu = (Number(allSpuVal) * 100).toFixed(1);
      const buySpu = (Number(buySpuVal) * 100).toFixed(1);
      const mara = (Number(maraPointVal) * 100).toFixed(1);
      const extra = (Number(extraPointVal) * 100).toFixed(1);
      detailNote = `全体SPU${allSpu}%/仕入SPU${buySpu}%/買回り+${mara}%/当日臨時+${extra}%/個別+${ind}% [${buyCategory}]`;
    } else {
      periodForSave = '─';
      detailNote = `個別+${(Number(indPointVal) * 100).toFixed(1)}% [${buyCategory}]`;
    }

    const dateFormatted = dateVal instanceof Date ? Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy/MM/dd') : dateVal;

    rowsToSave.push([
      0, // No.（後で採番）
      buyCategory,
      typeVal,
      dateFormatted,
      periodForSave,
      shopVal,
      String(janVal || ''),
      nameVal,
      priceVal,
      couponVal,
      inShippingVal, // 仕入送料
      pointPayVal,   // ポイント利用
      payPriceVal,
      indPointVal,
      allSpuVal,
      buySpuVal,
      maraPointVal,
      extraPointVal,
      rateVal,
      detailNote,
      getPointVal,
      realCostVal,
      targetSellVal,
      shippingVal,
      shipCostVal,
      packingVal,
      feeVal,
      profitVal,
      profitRateVal,
      breakEvenVal,
      judgeVal
    ]);
  }

  if (rowsToSave.length === 0) {
    SpreadsheetApp.getUi().alert('シミュレーター（10〜19行目）に商品が入力されていません。');
    return;
  }

  // 履歴テーブルの書き込み開始行を特定（25行目以降）
  const lastRow = Math.max(sheet.getLastRow(), 24);
  let nextRow = 25;

  const checkColA = sheet.getRange(25, 1, Math.max(lastRow - 24, 1), 1).getValues();
  for (let i = 0; i < checkColA.length; i++) {
    if (checkColA[i][0] === '' || checkColA[i][0] === null) {
      nextRow = 25 + i;
      break;
    }
    if (i === checkColA.length - 1) {
      nextRow = 25 + i + 1;
    }
  }

  // No. を採番
  let currentNo = nextRow - 24;
  for (let i = 0; i < rowsToSave.length; i++) {
    rowsToSave[i][0] = currentNo++;
  }

  // 一括書き込み
  sheet.getRange(nextRow, 1, rowsToSave.length, rowsToSave[0].length).setValues(rowsToSave);

  // 書式一括適用
  const writeRange = sheet.getRange(nextRow, 1, rowsToSave.length, rowsToSave[0].length);
  sheet.getRange(nextRow, 1).setHorizontalAlignment('center'); // No.
  sheet.getRange(nextRow, 2).setHorizontalAlignment('center').setFontWeight('bold'); // 購入区分
  sheet.getRange(nextRow, 3).setHorizontalAlignment('center'); // 仕入先区分
  sheet.getRange(nextRow, 4).setNumberFormat('yyyy/MM/dd').setHorizontalAlignment('center'); // 検討日
  sheet.getRange(nextRow, 5).setHorizontalAlignment('center'); // マラソン期間
  sheet.getRange(nextRow, 7).setNumberFormat('@'); // JAN
  sheet.getRange(nextRow, 9, rowsToSave.length, 5).setNumberFormat('#,##0'); // 定価・クーポン・仕入送料・P利用・支払
  sheet.getRange(nextRow, 14, rowsToSave.length, 6).setNumberFormat('0.0%'); // 個別P・全体SPU・仕入SPU・マラソンP・当日臨時P・還元率
  sheet.getRange(nextRow, 21, rowsToSave.length, 3).setNumberFormat('#,##0'); // 獲得P・実質原価・売価
  sheet.getRange(nextRow, 25, rowsToSave.length, 4).setNumberFormat('#,##0'); // 販売送料・梱包・手数料・利益
  sheet.getRange(nextRow, 29, rowsToSave.length, 1).setNumberFormat('0.0%').setFontWeight('bold'); // 利益率
  sheet.getRange(nextRow, 30, rowsToSave.length, 1).setNumberFormat('#,##0'); // 損益分岐
  sheet.getRange(nextRow, 31, rowsToSave.length, 1).setHorizontalAlignment('center'); // 判定

  writeRange.setBorder(true, true, true, true, true, true, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID);

  SpreadsheetApp.getActive().toast(`${rowsToSave.length}件の商品を検討履歴に一括保存しました！`, '保存完了', 5);
}

/**
 * シミュレーター（10〜19行目）の商品情報のみを一括クリア
 */
function シミュレーターを一括クリア() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_CALC);
  if (!sheet) return;

  sheet.getRange('D10:H19').clearContent(); // 店舗名〜クーポン
  sheet.getRange('I10:J19').setValue(0); // 仕入送料 & ポイント利用
  sheet.getRange('L10:L19').setValue(0); // 個別P
  sheet.getRange('U10:U19').clearContent(); // 売価

  SpreadsheetApp.getActive().toast('シミュレーターをクリアしました。次のマラソンを試算できます。', 'リセット', 3);
}
