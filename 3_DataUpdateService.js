/**
 * データ更新・登録系
 * 【高速化版】セット操作の集約と、検索ロジックのMap化
 */

function recordExchangeComplete(shopCode, machineIdentifier, exchangeType, exchangeDate, accumulatedCount) {
  const config = getConfig();
  const masterSheet = getSheet(config.SHEET_NAMES.MASTER_EQUIPMENT);
  const historySheet = getSheet(config.SHEET_NAMES.HISTORY);
  
  const masterData = masterSheet.getDataRange().getValues();
  const headers = masterData[0];
  
  const rowIndex = masterData.findIndex(row => row[0] == shopCode && row[2] == machineIdentifier);
  if (rowIndex === -1) throw new Error('マスタに対象の洗車機が見つかりません');
  
  const rowRange = masterSheet.getRange(rowIndex + 1, 1, 1, headers.length);
  const rowValues = rowRange.getValues()[0]; // 現在の値を取得
  
  const railDateIdx = headers.indexOf('レール交換実施日');
  const brushDateIdx = headers.indexOf('布ブラシ最終交換日');
  const brushCountIdx = headers.indexOf('布ブラシ最終交換台数');
  const bodyDateIdx = headers.indexOf('本体設置日');
  
  // ★高速化: setValuesで一括更新するために配列を操作
  if (exchangeType === 'レール車輪') { 
    rowValues[railDateIdx] = exchangeDate;
  }
  else if (exchangeType === '布ブラシ') { 
    rowValues[brushDateIdx] = exchangeDate;
    rowValues[brushCountIdx] = accumulatedCount;
  }
  else if (exchangeType === '本体') { 
    rowValues[bodyDateIdx] = exchangeDate;
    rowValues[railDateIdx] = exchangeDate;
    rowValues[brushDateIdx] = exchangeDate;
    rowValues[brushCountIdx] = 0;
  }
  
  // メモ欄（9列目=インデックス8）をクリア
  if (rowValues.length > 8) rowValues[8] = "";

  // ★高速化: 1回の通信で書き込む
  rowRange.setValues([rowValues]);
  
  historySheet.appendRow([shopCode, machineIdentifier, exchangeType, exchangeDate, accumulatedCount]);
  
  SpreadsheetApp.flush();
  refreshStatusSummary();
  return { success: true };
}

function recordMonthlyUpdate(updateRecords) {
  const config = getConfig();
  const monthlySheet = getSheet(config.SHEET_NAMES.MONTHLY_DATA);
  const masterData = getEquipmentMasterData();
  
  const installDateMap = new Map();
  masterData.forEach(item => {
    const key = `${item['店舗コード']}_${item['区別']}`;
    if (item['本体設置日'] && item['本体設置日'] instanceof Date) {
      installDateMap.set(key, item['本体設置日']);
    }
  });

  const range = monthlySheet.getDataRange();
  const allData = range.getValues();
  const headers = allData[0];
  const idxDate = headers.indexOf('年月');
  const idxShop = headers.indexOf('店舗コード');
  const idxMachine = headers.indexOf('区別');
  const idxDaily = headers.indexOf('日次台数');
  const idxAccum = headers.indexOf('累計台数');

  const parseDate = (d) => new Date(d);
  
  // ★高速化: 全データを「店舗_区別」でグループ化してMapにする
  // これにより findPrevAccum が爆速になる（O(N) -> O(1)）
  const machineDataMap = new Map();
  for (let i = 1; i < allData.length; i++) {
    const key = `${allData[i][idxShop]}_${allData[i][idxMachine]}`;
    if (!machineDataMap.has(key)) {
      machineDataMap.set(key, []);
    }
    // 行インデックスとデータを保存
    machineDataMap.get(key).push({
      rowIndex: i, 
      date: parseDate(allData[i][idxDate]),
      accum: parseFloat(allData[i][idxAccum]) || 0
    });
  }

  const updatesMap = new Map();
  updateRecords.forEach(r => {
    const d = new Date(r.dateMonth);
    const key = `${d.getFullYear()}-${d.getMonth()}_${r.shopCode}_${r.machineIdentifier}`;
    updatesMap.set(key, r);
  });

  // 高速化された検索関数
  const findPrevAccumFast = (targetShop, targetMachine, targetDate) => {
    const key = `${targetShop}_${targetMachine}`;
    const records = machineDataMap.get(key);
    if (!records) return 0;

    let maxDate = null;
    let prevAccum = 0;

    // 自分のマシンのレコードだけを走査（件数が少ないので速い）
    for (const rec of records) {
      if (rec.date < targetDate) {
        if (maxDate === null || rec.date > maxDate) {
          maxDate = rec.date;
          prevAccum = rec.accum;
        }
      }
    }

    const installDate = installDateMap.get(key);
    if (installDate) {
      const installMonthStart = new Date(installDate.getFullYear(), installDate.getMonth(), 1);
      if ((maxDate === null || installMonthStart > maxDate) && installMonthStart <= targetDate) {
        return 0;
      }
    }
    return prevAccum;
  };

  // 1. 既存行の更新 (メモリ上の配列を更新)
  let isDataChanged = false;
  for (let i = 1; i < allData.length; i++) {
    const rowDate = parseDate(allData[i][idxDate]);
    const shop = allData[i][idxShop];
    const machine = allData[i][idxMachine];
    const key = `${rowDate.getFullYear()}-${rowDate.getMonth()}_${shop}_${machine}`;
    
    if (updatesMap.has(key)) {
      const updateData = updatesMap.get(key);
      const newDaily = Number(updateData.dailyCount);
      const prevAccum = findPrevAccumFast(shop, machine, rowDate);
      const newAccum = prevAccum + newDaily;

      // 配列を直接更新
      allData[i][idxDaily] = newDaily;
      allData[i][idxAccum] = newAccum;
      
      updatesMap.delete(key);
      isDataChanged = true;
    }
  }

  // 変更があれば一括更新
  if (isDataChanged) {
    range.setValues(allData);
  }

  // 2. 新規行の追加
  const newRows = [];
  updatesMap.forEach((r) => {
    const targetDate = new Date(r.dateMonth);
    const prevAccum = findPrevAccumFast(r.shopCode, r.machineIdentifier, targetDate);
    const newAccum = prevAccum + Number(r.dailyCount);

    newRows.push([r.dateMonth, r.shopCode, r.machineIdentifier, r.dailyCount, newAccum]);
  });

  if (newRows.length > 0) {
    const lastRow = monthlySheet.getLastRow();
    monthlySheet.getRange(lastRow + 1, 1, newRows.length, 5).setValues(newRows);
  }

  SpreadsheetApp.flush();
  refreshStatusSummary(); 
  return { success: true, updatedCount: updateRecords.length };
}