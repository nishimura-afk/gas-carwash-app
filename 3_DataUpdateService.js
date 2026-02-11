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
    monthlySheet.getRange(lastRow + 1, 1, lastRow + newRows.length, 5).setValues(newRows);
  }

  SpreadsheetApp.flush();
  refreshStatusSummary(); 
  return { success: true, updatedCount: updateRecords.length };
}

/**
 * 指定年月・店舗・区別の月次データ1件を取得（修正モード用）
 * @return {Object|null} { dailyCount, accumCount } または null
 */
function getMonthlyDataForEdit(yearMonth, shopCode, machineIdentifier) {
  const config = getConfig();
  const monthlySheet = getSheet(config.SHEET_NAMES.MONTHLY_DATA);
  const allData = monthlySheet.getDataRange().getValues();
  const headers = allData[0];
  const idxDate = headers.indexOf('年月');
  const idxShop = headers.indexOf('店舗コード');
  const idxMachine = headers.indexOf('区別');
  const idxDaily = headers.indexOf('日次台数');
  const idxAccum = headers.indexOf('累計台数');
  if (idxDate < 0 || idxShop < 0 || idxMachine < 0 || idxDaily < 0 || idxAccum < 0) return null;

  const targetDate = new Date(yearMonth + '-01');
  const targetYm = targetDate.getFullYear() + '-' + ('0' + (targetDate.getMonth() + 1)).slice(-2);

  for (let i = 1; i < allData.length; i++) {
    const rowDate = allData[i][idxDate];
    const rowYm = (rowDate instanceof Date)
      ? (rowDate.getFullYear() + '-' + ('0' + (rowDate.getMonth() + 1)).slice(-2))
      : String(rowDate).slice(0, 7);
    if (rowYm === targetYm && allData[i][idxShop] === shopCode && allData[i][idxMachine] === machineIdentifier) {
      return {
        dailyCount: parseFloat(allData[i][idxDaily]) || 0,
        accumCount: parseFloat(allData[i][idxAccum]) || 0,
        yearMonth: targetYm
      };
    }
  }
  return null;
}

/**
 * 月次データの個別修正（指定年月・店舗・区別の日次台数を上書きし、累計を再計算して保存）
 * その月の1行のみ更新する。以降の月は「累計台数を再計算」で一括更新すること。
 */
function saveMonthlyDataCorrection(yearMonth, shopCode, machineIdentifier, dailyCount) {
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
  const targetDate = new Date(yearMonth + '-01');
  const key = `${shopCode}_${machineIdentifier}`;
  const installDate = installDateMap.get(key);
  const installMonthStart = installDate ? new Date(installDate.getFullYear(), installDate.getMonth(), 1) : null;

  let prevAccum = 0;
  let maxPrevDate = null;
  let targetRowIndex = -1;
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][idxShop] !== shopCode || allData[i][idxMachine] !== machineIdentifier) continue;
    const rowDate = parseDate(allData[i][idxDate]);
    if (rowDate.getFullYear() === targetDate.getFullYear() && rowDate.getMonth() === targetDate.getMonth()) {
      targetRowIndex = i;
      continue;
    }
    if (rowDate >= targetDate) continue;
    if (installMonthStart && rowDate < installMonthStart) continue;
    const acc = parseFloat(allData[i][idxAccum]) || 0;
    if (maxPrevDate === null || rowDate > maxPrevDate) {
      maxPrevDate = rowDate;
      prevAccum = acc;
    }
  }
  if (targetRowIndex < 0) return { success: false, error: '該当する月次データが見つかりません' };

  const newDaily = Number(dailyCount);
  const newAccum = (installMonthStart && targetDate < installMonthStart) ? 0 : (prevAccum + newDaily);
  allData[targetRowIndex][idxDaily] = newDaily;
  allData[targetRowIndex][idxAccum] = newAccum;
  range.setValues(allData);
  SpreadsheetApp.flush();
  refreshStatusSummary();
  return { success: true };
}

/**
 * 累計台数の一括再計算
 * 月次データを年月順にソートし、店舗コード＋区別ごとに古い月から累計を再計算する。
 * 本体設置日より前のデータは累計を0にする。最後に refreshStatusSummary() でキャッシュ更新。
 * @return {{ updatedRows: number }}
 */
function recalculateAllAccumulatedCounts() {
  const config = getConfig();
  const monthlySheet = getSheet(config.SHEET_NAMES.MONTHLY_DATA);
  const masterData = getEquipmentMasterData();
  const installDateMap = new Map();
  masterData.forEach(item => {
    const k = `${item['店舗コード']}_${item['区別']}`;
    if (item['本体設置日'] && item['本体設置日'] instanceof Date) {
      installDateMap.set(k, item['本体設置日']);
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
  if ([idxDate, idxShop, idxMachine, idxDaily, idxAccum].some(i => i < 0)) {
    throw new Error('月次データシートの列が見つかりません');
  }

  const parseDate = (d) => new Date(d);
  const groupKey = (row) => `${row[idxShop]}_${row[idxMachine]}`;

  const rowsByGroup = new Map();
  for (let i = 1; i < allData.length; i++) {
    const key = groupKey(allData[i]);
    if (!rowsByGroup.has(key)) rowsByGroup.set(key, []);
    rowsByGroup.get(key).push({ rowIndex: i, date: parseDate(allData[i][idxDate]), daily: parseFloat(allData[i][idxDaily]) || 0 });
  }

  rowsByGroup.forEach((rows, key) => {
    rows.sort((a, b) => a.date.getTime() - b.date.getTime());
    const installDate = installDateMap.get(key);
    const installMonthStart = installDate ? new Date(installDate.getFullYear(), installDate.getMonth(), 1) : null;
    let prevAccum = 0;
    rows.forEach(r => {
      let newAccum;
      if (installMonthStart && r.date < installMonthStart) {
        newAccum = 0;
      } else {
        newAccum = prevAccum + r.daily;
      }
      prevAccum = newAccum;
      allData[r.rowIndex][idxAccum] = newAccum;
    });
  });

  range.setValues(allData);
  SpreadsheetApp.flush();
  refreshStatusSummary();
  return { updatedRows: allData.length - 1 };
}