function insertSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 顧客マスタシートのサンプルデータ
  const customerMasterSheet = ss.getSheetByName("顧客マスタシート");
  if (customerMasterSheet) {
    const customerData = [
      ["C001", "鈴木クリニック", "法人", "本院", "100-0001", "東京都", "千代田区", "丸の内1-1-1", "03-1111-2222", "山田太郎", "yamada@suzuki-clinic.com", "駐車場あり"],
      ["C002", "田中様", "個人", "自宅", "150-0002", "東京都", "渋谷区", "道玄坂2-2-2", "090-3333-4444", "田中一郎", "tanaka@example.com", "犬を飼っている"],
      ["C003", "株式会社ABC商事", "法人", "東京支社", "530-0003", "大阪府", "北区", "梅田3-3-3 ABCビル10F", "06-5555-6666", "佐藤花子", "sato@abc-corp.co.jp", "毎週水曜午前中希望"]
    ];
    customerMasterSheet.getRange(customerMasterSheet.getLastRow() + 1, 1, customerData.length, customerData[0].length).setValues(customerData);
  } else {
    SpreadsheetApp.getUi().alert("エラー: '顧客マスタシート' が見つかりません。");
    return;
  }

  // 施工履歴シートのサンプルデータ
  const constructionHistorySheet = ss.getSheetByName("施工履歴シート");
  if (constructionHistorySheet) {
    const historyData = [
      ["H001", "C001", "本院", "2025/07/01", "佐藤", "クリニック清掃", "診察室・待合室の床洗浄", 180, "〇〇先生からお礼の言葉", 30000],
      ["H002", "C002", "自宅", "2025/07/05", "鈴木", "ハウスクリーニング", "リビング・寝室の清掃", 120, "窓の汚れが目立つ", 20000],
      ["H003", "C001", "本院", "2025/07/10", "佐藤", "エアコンクリーニング", "待合室エアコン2台洗浄", 90, "室外機も洗浄済み", 15000],
      ["H004", "C003", "東京支社", "2025/07/15", "田中", "定期清掃", "オフィスフロア清掃", 240, "次回は給湯室も確認", 40000]
    ];
    constructionHistorySheet.getRange(constructionHistorySheet.getLastRow() + 1, 1, historyData.length, historyData[0].length).setValues(historyData);
    // 日付形式を適用
    constructionHistorySheet.getRange(constructionHistorySheet.getLastRow() - historyData.length + 1, 4, historyData.length, 1).setNumberFormat("yyyy/mm/dd");
  } else {
    SpreadsheetApp.getUi().alert("エラー: '施工履歴シート' が見つかりません。");
    return;
  }

  // クレーム・注意点シートのサンプルデータ
  const claimNotesSheet = ss.getSheetByName("クレーム・注意点シート");
  if (claimNotesSheet) {
    const claimData = [
      ["K001", "C001", "本院", "注意点", "2025/06/28", "〇〇室は立ち入り禁止", "未対応", "", ""],
      ["K002", "C002", "自宅", "クレーム", "2025/07/06", "窓拭きに拭きムラがあった", "対応中", "りょうた", "〇月〇日に謝罪訪問、再清掃実施予定"],
      ["K003", "C003", "東京支社", "注意点", "2025/07/12", "〇〇部長室は特に丁寧に", "未対応", "", ""]
    ];
    claimNotesSheet.getRange(claimNotesSheet.getLastRow() + 1, 1, claimData.length, claimData[0].length).setValues(claimData);
    // 日付形式を適用
    claimNotesSheet.getRange(claimNotesSheet.getLastRow() - claimData.length + 1, 5, claimData.length, 1).setNumberFormat("yyyy/mm/dd");
  } else {
    SpreadsheetApp.getUi().alert("エラー: 'クレーム・注意点シート' が見つかりません。");
    return;
  }

  SpreadsheetApp.getUi().alert('サンプルデータの挿入が完了しました！');
}