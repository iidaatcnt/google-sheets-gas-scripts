function setupConstructionManagementSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 顧客マスタシートの作成と設定
  let customerMasterSheet = ss.getSheetByName("顧客マスタシート");
  if (!customerMasterSheet) {
    customerMasterSheet = ss.insertSheet("顧客マスタシート", 0);
  }
  customerMasterSheet.clear(); // 既存の内容をクリア
  const customerMasterHeaders = [
    "顧客ID", "顧客名", "顧客区分", "現場名", "郵便番号", "都道府県",
    "市区町村", "番地以降", "現場連絡先", "担当者名", "メールアドレス", "備考"
  ];
  customerMasterSheet.getRange(1, 1, 1, customerMasterHeaders.length).setValues([customerMasterHeaders]).setFontWeight("bold");
  customerMasterSheet.setFrozenRows(1); // 1行目を固定

  // 顧客区分プルダウン
  const customerTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['法人', '個人'])
    .setAllowInvalid(false)
    .setHelpText('法人か個人かを選択してください。')
    .build();
  customerMasterSheet.getRange("C2:C").setDataValidation(customerTypeRule);

  // 都道府県プルダウン (主要な都道府県をリストアップ)
  const prefectures = [
    '北海道', '青森県', '岩手県', '宮城県', '秋田県', '山形県', '福島県',
    '茨城県', '栃木県', '群馬県', '埼玉県', '千葉県', '東京都', '神奈川県',
    '新潟県', '富山県', '石川県', '福井県', '山梨県', '長野県', '岐阜県',
    '静岡県', '愛知県', '三重県', '滋賀県', '京都府', '大阪府', '兵庫県',
    '奈良県', '和歌山県', '鳥取県', '島根県', '岡山県', '広島県', '山口県',
    '徳島県', '香川県', '愛媛県', '高知県', '福岡県', '佐賀県', '長崎県',
    '熊本県', '大分県', '宮崎県', '鹿児島県', '沖縄県'
  ];
  const prefectureRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(prefectures)
    .setAllowInvalid(false)
    .setHelpText('都道府県を選択してください。')
    .build();
  customerMasterSheet.getRange("F2:F").setDataValidation(prefectureRule);

  // 顧客IDの備考
  customerMasterSheet.getRange("A1").setNote("各顧客を一意に識別するID。手動で連番を振るか、GASで自動採番を設定できます。");


  // 2. 施工履歴シートの作成と設定
  let constructionHistorySheet = ss.getSheetByName("施工履歴シート");
  if (!constructionHistorySheet) {
    constructionHistorySheet = ss.insertSheet("施工履歴シート", 1);
  }
  constructionHistorySheet.clear(); // 既存の内容をクリア
  const constructionHistoryHeaders = [
    "履歴ID", "顧客ID", "現場名", "施工日", "担当者", "作業区分",
    "施工内容詳細", "所要時間（分）", "特記事項", "金額"
  ];
  constructionHistorySheet.getRange(1, 1, 1, constructionHistoryHeaders.length).setValues([constructionHistoryHeaders]).setFontWeight("bold");
  constructionHistorySheet.setFrozenRows(1); // 1行目を固定

  // 作業区分プルダウン
  const workTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['クリニック清掃', 'エアコンクリーニング', 'ハウスクリーニング', '定期清掃', 'その他'])
    .setAllowInvalid(false)
    .setHelpText('清掃の種類を選択してください。')
    .build();
  constructionHistorySheet.getRange("F2:F").setDataValidation(workTypeRule);

  // 履歴IDの備考
  constructionHistorySheet.getRange("A1").setNote("各施工履歴を一意に識別するID。手動で連番を振るか、GASで自動採番を設定できます。");
  // 顧客IDと現場名の備考
  constructionHistorySheet.getRange("B1").setNote("顧客マスタシートの顧客IDと現場名を参照してください。データ検証で顧客マスタから選択できるように設定すると便利です。");
  constructionHistorySheet.getRange("C1").setNote("顧客マスタシートの顧客IDと現場名を参照してください。データ検証で顧客マスタから選択できるように設定すると便利です。");
  // 施工日の備考 (日付形式の推奨)
  constructionHistorySheet.getRange("D1").setNote("日付形式で入力してください。");
  constructionHistorySheet.getRange("D2:D").setNumberFormat("yyyy/mm/dd");


  // 3. クレーム・注意点シートの作成と設定
  let claimNotesSheet = ss.getSheetByName("クレーム・注意点シート");
  if (!claimNotesSheet) {
    claimNotesSheet = ss.insertSheet("クレーム・注意点シート", 2);
  }
  claimNotesSheet.clear(); // 既存の内容をクリア
  const claimNotesHeaders = [
    "項目ID", "顧客ID", "現場名", "種別", "発生・登録日", "内容",
    "対応状況", "対応者", "対応履歴・詳細"
  ];
  claimNotesSheet.getRange(1, 1, 1, claimNotesHeaders.length).setValues([claimNotesHeaders]).setFontWeight("bold");
  claimNotesSheet.setFrozenRows(1); // 1行目を固定

  // 種別プルダウン
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['クレーム', '注意点'])
    .setAllowInvalid(false)
    .setHelpText('クレームか注意点かを選択してください。')
    .build();
  claimNotesSheet.getRange("D2:D").setDataValidation(typeRule);

  // 対応状況プルダウン
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['解決済み', '対応中', '再訪問予定', '未対応'])
    .setAllowInvalid(false)
    .setHelpText('対応状況を選択してください。')
    .build();
  claimNotesSheet.getRange("G2:G").setDataValidation(statusRule);

  // 項目IDの備考
  claimNotesSheet.getRange("A1").setNote("各項目を一意に識別するID。手動で連番を振るか、GASで自動採番を設定できます。");
  // 顧客IDと現場名の備考
  claimNotesSheet.getRange("B1").setNote("顧客マスタシートの顧客IDと現場名を参照してください。データ検証で顧客マスタから選択できるように設定すると便利です。");
  claimNotesSheet.getRange("C1").setNote("顧客マスタシートの顧客IDと現場名を参照してください。データ検証で顧客マスタから選択できるように設定すると便利です。");
  // 発生・登録日の備考 (日付形式の推奨)
  claimNotesSheet.getRange("E1").setNote("日付形式で入力してください。");
  claimNotesSheet.getRange("E2:E").setNumberFormat("yyyy/mm/dd");

  SpreadsheetApp.getUi().alert('スプレッドシートのセットアップが完了しました！');
}