/**
 * データ取得・計算ロジック・キャッシュ生成
 * 【高速化版】データの事前グルーピングにより計算量を劇的に削減
 */
function getEquipmentMasterData() {
  const sheet = getSheet(getConfig().SHEET_NAMES.MASTER_EQUIPMENT);
  // データがある範囲だけを取得（無駄な空行を読まない）
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  const values = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  const headers = values[0];
  return values.slice(1).map(row => {
    const obj = {}; headers.forEach((h, i) => obj[h] = row[i]);
    if (row.length > 8) {
      obj['次回付帯作業メモ'] = row[8];
    } else {
      obj['次回付帯作業メモ'] = "";
    }
    return obj;
  }).filter(o => o['店舗コード']);
}

function getMonthlyData() {
  const sheet = getSheet(getConfig().SHEET_NAMES.MONTHLY_DATA);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  const values = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  const headers = values[0];
  return values.slice(1).map(row => {
    const obj = {}; headers.forEach((h, i) => obj[h] = row[i]); return obj;
  }).filter(o => o['店舗コード']);
}

function calculateExchangeStatus(masterItem, currentAccumulatedCount, avgMonthlyCount = 0) {
  const config = getConfig();
  const criteria = config.EXCHANGE_CRITERIA;
  const status = config.STATUS;
  const today = new Date();

  const lastBrushCount = parseFloat(masterItem['布ブラシ最終交換台数']) || 0;
  const countSinceBrushExchange = Math.max(0, currentAccumulatedCount - lastBrushCount);
  const countSinceBodyExchange = currentAccumulatedCount;

  const getMonthsDiff = (date) => {
    if (!date || !(date instanceof Date) || isNaN(date.getTime())) return 0;
    return (today.getFullYear() - date.getFullYear()) * 12 + (today.getMonth() - date.getMonth());
  };

  const railMonths = getMonthsDiff(masterItem['レール交換実施日']);
  const brushMonths = getMonthsDiff(masterItem['布ブラシ最終交換日']);

  // --- レール判定 ---
  let railStatus = status.NORMAL;
  const hasRailExchanged = masterItem['レール交換実施日'] && !isNaN(new Date(masterItem['レール交換実施日']).getTime());

  // 未交換の場合のみ判定
  if (!hasRailExchanged) {
    if (currentAccumulatedCount >= criteria.RAIL_WHEEL.THRESHOLD) {
      railStatus = status.NOTICE;
    }
  }

  // --- ブラシ判定 ---
  let brushStatus = status.NORMAL;
  if (masterItem['ブラシ種類'] === '布') {
    if (countSinceBrushExchange >= criteria.CLOTH_BRUSH.SECOND || brushMonths >= criteria.CLOTH_BRUSH.MONTHS_WARNING * 2) brushStatus = status.EXCHANGE_2;
    else if (countSinceBrushExchange >= criteria.CLOTH_BRUSH.FIRST || brushMonths >= criteria.CLOTH_BRUSH.MONTHS_WARNING) brushStatus = status.EXCHANGE;
  } else { brushStatus = '対象外'; }
  if (currentAccumulatedCount < 1000 && masterItem['ブラシ種類'] === '布') brushStatus = status.NORMAL;

  // --- 本体判定（新ロジック） ---
  let bodyStatus = status.NORMAL;
  const currentMonth = today.getMonth() + 1; // 1-12
  
  if (currentMonth === 1) {
    // 1月の場合：15ヶ月後の予測で判定
    const forecastCount = countSinceBodyExchange + (avgMonthlyCount * criteria.MACHINE_BODY.FORECAST_MONTHS);
    Logger.log(`[本体判定-1月] 累計: ${countSinceBodyExchange}, 月平均: ${avgMonthlyCount}, 予測: ${forecastCount}, 閾値: ${criteria.MACHINE_BODY.THRESHOLD}`);
    if (forecastCount >= criteria.MACHINE_BODY.THRESHOLD) {
      bodyStatus = status.PREPARE;
      Logger.log(`[本体判定-1月] アラート発生: ${forecastCount} >= ${criteria.MACHINE_BODY.THRESHOLD}`);
    }
  } else {
    // 2月以降：累計台数のみで判定
    if (countSinceBodyExchange >= criteria.MACHINE_BODY.THRESHOLD) {
      bodyStatus = status.PREPARE;
    }
  }

  return { railStatus, brushStatus, bodyStatus, railMonths };
}

function refreshStatusSummary() {
  const masterData = getEquipmentMasterData();
  const monthlyData = getMonthlyData();
  const config = getConfig();

  // ★高速化ポイント: 月次データを「店舗コード_区別」をキーにした辞書(Map)に事前変換
  // これにより、ループの中で毎回全データを検索する必要がなくなる（計算量が激減）
  const monthlyDataMap = new Map();
  
  monthlyData.forEach(r => {
    // 累計台数が有効なもののみ対象
    if (r['累計台数'] !== "" && r['累計台数'] != null) {
      const key = `${r['店舗コード']}_${r['区別']}`;
      if (!monthlyDataMap.has(key)) {
        monthlyDataMap.set(key, []);
      }
      monthlyDataMap.get(key).push(r);
    }
  });

  const allStatus = masterData.map(item => {
    if (!item || !item['店舗コード']) return null;
    
    const toDate = (d) => (d instanceof Date && !isNaN(d)) ? d : null;
    const bodyInstallDate = toDate(item['本体設置日']);
    
    // ★高速化: 辞書から一発で対象データ群を取得
    const key = `${item['店舗コード']}_${item['区別']}`;
    const targetRecords = monthlyDataMap.get(key) || [];
    
    // 最新レコード特定（日付＞数値）
    const latestRecord = targetRecords.sort((a, b) => {
      const dateA = new Date(a['年月']);
      const dateB = new Date(b['年月']);
      const diffTime = dateB - dateA;
      if (diffTime !== 0) return diffTime; 
      return (parseFloat(b['累計台数']) || 0) - (parseFloat(a['累計台数']) || 0);
    })[0];
    
    let count = latestRecord ? (parseFloat(latestRecord['累計台数']) || 0) : 0;

    if (bodyInstallDate && latestRecord) {
      const recordDate = new Date(latestRecord['年月']);
      const installMonth = new Date(bodyInstallDate.getFullYear(), bodyInstallDate.getMonth(), 1);
      const recordMonth = new Date(recordDate.getFullYear(), recordDate.getMonth(), 1);
      if (installMonth > recordMonth) count = 0; 
    }

    // 月平均台数の計算（月次データの累計台数の増分の平均）
    let avgCount = 0;
    if (targetRecords.length > 1) {
      // 時系列でソート
      const sortedRecords = targetRecords.sort((a, b) => {
        const dateA = new Date(a['年月']);
        const dateB = new Date(b['年月']);
        return dateA - dateB;
      });
      
      // 各月の累計台数の増分を計算
      const monthlyIncrements = [];
      for (let i = 1; i < sortedRecords.length; i++) {
        const prevCount = parseFloat(sortedRecords[i - 1]['累計台数']) || 0;
        const currCount = parseFloat(sortedRecords[i]['累計台数']) || 0;
        const increment = currCount - prevCount;
        if (increment > 0) {
          monthlyIncrements.push(increment);
        }
      }
      
      // 増分の平均を計算（直近12ヶ月分を使用）
      if (monthlyIncrements.length > 0) {
        const recentIncrements = monthlyIncrements.slice(-12); // 直近12ヶ月
        const sum = recentIncrements.reduce((acc, val) => acc + val, 0);
        avgCount = Math.round(sum / recentIncrements.length);
      }
    } else if (targetRecords.length === 1) {
      // データが1件のみの場合は日次台数を使用（後方互換性）
      avgCount = Math.round(parseFloat(targetRecords[0]['日次台数']) || 0);
    }
    
    const validItem = { 
      ...item, 
      '本体設置日': bodyInstallDate, 
      'レール交換実施日': toDate(item['レール交換実施日']), 
      '布ブラシ最終交換日': toDate(item['布ブラシ最終交換日']) 
    };

    const res = calculateExchangeStatus(validItem, count, avgCount);

    let isSubsidy = config.SUBSIDY_TARGET_SHOPS.some(s => item['店舗名'].includes(s));
    const memo = item['次回付帯作業メモ'] || "";

    return [
      item['店舗コード'], item['店舗名'], item['区別'], count, avgCount,
      validItem['本体設置日'], res.railStatus, res.brushStatus, res.bodyStatus, 
      item['ブラシ種類'] === '布' ? '布' : 'スポンジ', res.railMonths,
      isSubsidy,
      memo
    ];
  }).filter(i => i !== null);

  const sheet = getSheet(config.SHEET_NAMES.STATUS_SUMMARY);
  sheet.clearContents();
  
  const headers = ['店舗コード', '店舗名', '区別', '累計台数', '月平均台数', '本体設置日', 'レールステータス', 'ブラシステータス', '本体ステータス', '布交換対象', 'railMonths', 'isSubsidy', 'nextWorkMemo'];
  
  if (allStatus.length > 0) {
    sheet.getRange(1, 1, allStatus.length + 1, headers.length).setValues([headers, ...allStatus]);
    sheet.getRange(2, 4, allStatus.length, 2).setNumberFormat('#,##0');
    sheet.getRange(2, 6, allStatus.length, 1).setNumberFormat('yyyy/MM/dd');
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  Logger.log('✅ キャッシュ(ステータス集計)更新完了');
}

function getCarWashListCached() {
  const sheet = getSheet(getConfig().SHEET_NAMES.STATUS_SUMMARY);
  // 全範囲ではなくデータがある部分だけを取得
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  const values = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  const headers = values[0];
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { 
      obj[h] = (row[i] instanceof Date) ? Utilities.formatDate(row[i], Session.getScriptTimeZone(), 'yyyy/MM/dd') : row[i]; 
    });
    return obj;
  });
}