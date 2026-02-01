/**
 * 設備構成データのデバッグスクリプト（修正版）
 * GASエディタで実行して、ログを確認してください
 */

function debugEquipmentData() {
  const SPREADSHEET_ID = '1VMZ2RpYmCh0BZRwY4DPzriAafFOo-U9PqvOUZu67cf4';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('設備マスタ');
  
  if (!sheet) {
    Logger.log('❌ 設備マスタシートが見つかりません');
    return;
  }
  
  Logger.log('=== ヘッダー行の確認 ===');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  headers.forEach((header, index) => {
    const col = String.fromCharCode(65 + index);
    Logger.log(`${col}列 (index ${index}): ${header}`);
  });
  
  Logger.log('\n=== S001 糸我のデータ（2行目） ===');
  const firstRow = sheet.getRange(2, 1, 1, 13).getValues()[0];
  firstRow.forEach((value, index) => {
    const col = String.fromCharCode(65 + index);
    Logger.log(`${col}列 (index ${index}): "${value}" (型: ${typeof value})`);
  });
  
  Logger.log('\n=== J〜M列の生データ（インデックス9〜12） ===');
  Logger.log(`J列（総枠）index 9: "${firstRow[9]}" (型: ${typeof firstRow[9]})`);
  Logger.log(`K列（クリーナー）index 10: "${firstRow[10]}" (型: ${typeof firstRow[10]})`);
  Logger.log(`L列（SB）index 11: "${firstRow[11]}" (型: ${typeof firstRow[11]})`);
  Logger.log(`M列（マット日）index 12: "${firstRow[12]}" (型: ${typeof firstRow[12]})`);
  
  Logger.log('\n=== getAllShopEquipment()の実行 ===');
  try {
    const equipmentMap = getAllShopEquipment();
    Logger.log('✓ getAllShopEquipment()実行成功');
    Logger.log(`取得店舗数: ${Object.keys(equipmentMap).length}`);
    Logger.log('\nS001 糸我:', JSON.stringify(equipmentMap['S001'], null, 2));
    Logger.log('\nS014 徳島石井:', JSON.stringify(equipmentMap['S014'], null, 2));
    Logger.log('\nS021 牛久:', JSON.stringify(equipmentMap['S021'], null, 2));
  } catch (e) {
    Logger.log('❌ getAllShopEquipment()でエラー:', e.message);
  }
  
  Logger.log('\n=== 手動でS001のデータを取得 ===');
  const testData = sheet.getRange(2, 1, 1, 13).getValues()[0];
  const shopCode = testData[0];
  const shopName = testData[1];
  const totalSlots = testData[9];
  const cleanerCount = testData[10];
  const sbCount = testData[11];
  const mattDate = testData[12];
  
  Logger.log('店舗コード:', shopCode);
  Logger.log('店舗名:', shopName);
  Logger.log('総設置枠数:', totalSlots, '(空か？', totalSlots === '', ')');
  Logger.log('クリーナー台数:', cleanerCount, '(空か？', cleanerCount === '', ')');
  Logger.log('SB台数:', sbCount, '(空か？', sbCount === '', ')');
  Logger.log('マット設置日:', mattDate);
  
  Logger.log('\n=== parseInt()の結果 ===');
  Logger.log('parseInt(totalSlots):', parseInt(totalSlots));
  Logger.log('parseInt(cleanerCount):', parseInt(cleanerCount));
  Logger.log('parseInt(sbCount):', parseInt(sbCount));
}

// GASエディタで実行
function testEquipmentAfterFix() {
  const equipment = getAllShopEquipment();
  Logger.log('取得店舗数:', Object.keys(equipment).length);
  Logger.log('S001:', JSON.stringify(equipment['S001'], null, 2));
  Logger.log('S014:', JSON.stringify(equipment['S014'], null, 2));
  Logger.log('S021:', JSON.stringify(equipment['S021'], null, 2));
}

/**
 * 設備構成取得の最終確認テスト
 */
function testEquipmentDataFinal() {
  Logger.log('=== getAllShopEquipment() テスト ===');
  
  const equipment = getAllShopEquipment();
  const shopCodes = Object.keys(equipment);
  
  Logger.log(`✅ 取得店舗数: ${shopCodes.length}店舗`);
  Logger.log('店舗コード一覧:', shopCodes.join(', '));
  
  // S001のデータを詳細表示
  Logger.log('\n=== S001 糸我 ===');
  if (equipment['S001']) {
    const s001 = equipment['S001'];
    Logger.log(`店舗コード: ${s001.shopCode}`);
    Logger.log(`店舗名: ${s001.shopName}`);
    Logger.log(`総設置枠数: ${s001.totalSlots}`);
    Logger.log(`クリーナー台数: ${s001.cleanerCount}`);
    Logger.log(`SB台数: ${s001.sbCount}`);
    Logger.log(`マット設置日: ${s001.mattDate}`);
  } else {
    Logger.log('❌ S001が見つかりません');
  }
  
  // S014のデータを詳細表示
  Logger.log('\n=== S014 徳島石井 ===');
  if (equipment['S014']) {
    const s014 = equipment['S014'];
    Logger.log(`店舗コード: ${s014.shopCode}`);
    Logger.log(`店舗名: ${s014.shopName}`);
    Logger.log(`総設置枠数: ${s014.totalSlots}`);
    Logger.log(`クリーナー台数: ${s014.cleanerCount}`);
    Logger.log(`SB台数: ${s014.sbCount}`);
    Logger.log(`マット設置日: ${s014.mattDate}`);
  } else {
    Logger.log('❌ S014が見つかりません');
  }
  
  // S021のデータを詳細表示
  Logger.log('\n=== S021 牛久 ===');
  if (equipment['S021']) {
    const s021 = equipment['S021'];
    Logger.log(`店舗コード: ${s021.shopCode}`);
    Logger.log(`店舗名: ${s021.shopName}`);
    Logger.log(`総設置枠数: ${s021.totalSlots}`);
    Logger.log(`クリーナー台数: ${s021.cleanerCount}`);
    Logger.log(`SB台数: ${s021.sbCount}`);
    Logger.log(`マット設置日: ${s021.mattDate}`);
  } else {
    Logger.log('❌ S021が見つかりません');
  }
  
  // refreshStatusSummaryのテスト
  Logger.log('\n=== refreshStatusSummary() 実行 ===');
  try {
    refreshStatusSummary();
    Logger.log('✅ ステータス集計シート更新完了');
    
    // 更新後のデータを確認
    Logger.log('\n=== ステータス集計シートの確認 ===');
    const cached = getCarWashListCached();
    Logger.log(`取得行数: ${cached.length}行`);
    
    if (cached.length > 0) {
      const firstRow = cached[0];
      Logger.log('\n1行目のデータ:');
      Logger.log(`店舗コード: ${firstRow['店舗コード']}`);
      Logger.log(`店舗名: ${firstRow['店舗名']}`);
      Logger.log(`区別: ${firstRow['区別']}`);
      Logger.log(`cleanerCount: ${firstRow['cleanerCount']}`);
      Logger.log(`sbCount: ${firstRow['sbCount']}`);
    }
    
  } catch (e) {
    Logger.log('❌ エラー:', e.message);
  }
}