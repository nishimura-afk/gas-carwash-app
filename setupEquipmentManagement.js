/**
 * 設備マスタに設備構成列を追加し、データを一括登録するスクリプト
 * 既存のA〜I列はそのまま維持し、J列以降に新規追加
 */

function addEquipmentColumns() {
  const SPREADSHEET_ID = '1VMZ2RpYmCh0BZRwY4DPzriAafFOo-U9PqvOUZu67cf4';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('設備マスタ');
  
  if (!sheet) {
    throw new Error('設備マスタシートが見つかりません');
  }
  
  // 1. ヘッダー行（1行目）にJ〜M列を追加
  const headers = ['総設置枠数', 'クリーナー台数', 'SB台数', 'マット設置日'];
  sheet.getRange(1, 10, 1, 4).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#cfe2f3');
  
  Logger.log('✓ ヘッダー追加完了');
  
  // 2. 各店舗の設備構成データ
  const equipmentData = {
    'S001': { slots: 1, cleaner: 1, sb: 0, mattDate: '2023/04/01' },
    'S002': { slots: 1, cleaner: 1, sb: 0, mattDate: '2023/12/01' },
    'S003': { slots: 4, cleaner: 4, sb: 0, mattDate: '2023/06/01' },
    'S004': { slots: 4, cleaner: 4, sb: 0, mattDate: '2023/09/01' },
    'S005': { slots: 2, cleaner: 2, sb: 0, mattDate: '2023/12/01' },
    'S006': { slots: 5, cleaner: 4, sb: 1, mattDate: '2024/12/01' },
    'S007': { slots: 0, cleaner: 0, sb: 0, mattDate: null },
    'S008': { slots: 5, cleaner: 4, sb: 1, mattDate: '2024/12/01' },
    'S009': { slots: 5, cleaner: 4, sb: 1, mattDate: '2025/03/01' },
    'S010': { slots: 3, cleaner: 3, sb: 0, mattDate: '2024/05/01' },
    'S011': { slots: 4, cleaner: 3, sb: 1, mattDate: '2025/05/01' },
    'S012': { slots: 4, cleaner: 4, sb: 0, mattDate: '2024/02/01' },
    'S013': { slots: 3, cleaner: 3, sb: 0, mattDate: '2023/11/01' },
    'S014': { slots: 5, cleaner: 4, sb: 1, mattDate: '2022/08/01' },
    'S015': { slots: 4, cleaner: 4, sb: 0, mattDate: '2022/12/01' },
    'S016': { slots: 2, cleaner: 2, sb: 0, mattDate: '2023/06/01' },
    'S017': { slots: 2, cleaner: 2, sb: 0, mattDate: '2024/10/01' },
    'S018': { slots: 3, cleaner: 3, sb: 0, mattDate: '2024/04/01' },
    'S019': { slots: 3, cleaner: 3, sb: 0, mattDate: '2024/06/01' },
    'S020': { slots: 5, cleaner: 5, sb: 0, mattDate: '2024/05/01' },
    'S021': { slots: 5, cleaner: 4, sb: 1, mattDate: '2025/11/29' },
    'S022': { slots: 6, cleaner: 6, sb: 0, mattDate: '2023/01/01' },
    'S023': { slots: 3, cleaner: 2, sb: 1, mattDate: '2024/03/01' },
    'S024': { slots: 4, cleaner: 3, sb: 1, mattDate: '2025/05/01' },
    'S025': { slots: 5, cleaner: 4, sb: 1, mattDate: '2025/11/01' },
    'S026': { slots: 5, cleaner: 4, sb: 1, mattDate: '2026/02/01' }
  };
  
  // 3. 全データを取得
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log('データ行がありません');
    return;
  }
  
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 9); // A2:I列まで取得
  const data = dataRange.getValues();
  
  // 4. 店舗コードごとに最初の行にのみデータを設定
  const processedShops = new Set();
  const newData = [];
  
  for (let i = 0; i < data.length; i++) {
    const shopCode = data[i][0]; // A列: 店舗コード
    
    if (!processedShops.has(shopCode) && equipmentData[shopCode]) {
      // この店舗の最初の行 → データを設定
      const eq = equipmentData[shopCode];
      newData.push([
        eq.slots,
        eq.cleaner,
        eq.sb,
        eq.mattDate || ''
      ]);
      processedShops.add(shopCode);
    } else {
      // 2行目以降 → 空白
      newData.push(['', '', '', '']);
    }
  }
  
  // 5. J2:M列にデータを書き込み
  if (newData.length > 0) {
    sheet.getRange(2, 10, newData.length, 4).setValues(newData);
    Logger.log(`✓ ${processedShops.size}店舗のデータを登録しました`);
  }
  
  // 6. 日付列（M列）の書式設定
  sheet.getRange(2, 13, lastRow - 1, 1).setNumberFormat('yyyy/MM/dd');
  
  Logger.log('✅ 設備構成データの登録が完了しました！');
  Logger.log('登録店舗数: ' + processedShops.size);
}

/**
 * 設備構成変更履歴シートを作成
 */
function createEquipmentHistorySheet() {
  const SPREADSHEET_ID = '1VMZ2RpYmCh0BZRwY4DPzriAafFOo-U9PqvOUZu67cf4';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 既存シートがあるか確認
  let sheet = ss.getSheetByName('設備構成変更履歴');
  
  if (sheet) {
    Logger.log('設備構成変更履歴シートは既に存在します');
    return;
  }
  
  // 新規作成
  sheet = ss.insertSheet('設備構成変更履歴');
  
  const headers = [
    '変更日',
    '店舗コード',
    '店舗名',
    '変更前_総設置枠数',
    '変更後_総設置枠数',
    '変更前_クリーナー台数',
    '変更後_クリーナー台数',
    '変更前_SB台数',
    '変更後_SB台数',
    '変更理由',
    '本体交換と同時'
  ];
  
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#cfe2f3');
  
  // 列幅を調整
  sheet.setColumnWidth(1, 100);  // 変更日
  sheet.setColumnWidth(2, 80);   // 店舗コード
  sheet.setColumnWidth(3, 120);  // 店舗名
  sheet.setColumnWidth(10, 200); // 変更理由
  
  Logger.log('✅ 設備構成変更履歴シートを作成しました');
}

/**
 * 両方を実行
 */
function setupEquipmentManagement() {
  try {
    Logger.log('========================================');
    Logger.log('設備構成管理のセットアップを開始');
    Logger.log('========================================');
    
    addEquipmentColumns();
    createEquipmentHistorySheet();
    
    Logger.log('========================================');
    Logger.log('✅ すべての設定が完了しました！');
    Logger.log('========================================');
    
  } catch (e) {
    Logger.log('❌ エラーが発生しました: ' + e.message);
    throw e;
  }
}