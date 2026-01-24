/**
 * 初期セットアップ
 * 必要なシートが存在しない場合に作成し、ヘッダーを設定します。
 */
function initialSetup() {
  const config = getConfig();
  const ss = SpreadsheetApp.openById(config.SPREADSHEET_ID);
  
  const sheetDefinitions = [
    { name: config.SHEET_NAMES.MASTER_EQUIPMENT, headers: ['店舗コード', '店舗名', '区別', 'ブラシ種類', '本体設置日', 'レール交換実施日', '布ブラシ最終交換日', '布ブラシ最終交換台数'] },
    { name: config.SHEET_NAMES.MASTER_STORE, headers: ['店舗コード', '店舗名', '担当者名', 'メールアドレス'] },
    { name: config.SHEET_NAMES.MONTHLY_DATA, headers: ['年月', '店舗コード', '区別', '日次台数', '累計台数'] },
    { name: config.SHEET_NAMES.SCHEDULE, headers: ['ID', '店舗コード', '区別', '部品名', '予定日', 'ステータス', 'カレンダーID', '発注先'] },
    { name: config.SHEET_NAMES.HISTORY, headers: ['店舗コード', '区別', '部品名', '交換日', '交換時累計台数'] },
    { name: config.SHEET_NAMES.STATUS_SUMMARY, headers: ['店舗コード', '店舗名', '区別', '累計台数', '本体設置日', 'レールステータス', 'ブラシステータス', '本体ステータス', '布交換対象', 'railMonths'] }
  ];

  for (const def of sheetDefinitions) {
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
      sheet.getRange(1, 1, 1, def.headers.length)
           .setValues([def.headers])
           .setFontWeight('bold')
           .setBackground('#cfe2f3');
    }
  }

  Logger.log('✅ 初期設定完了。各シートにデータを準備してください。');
}