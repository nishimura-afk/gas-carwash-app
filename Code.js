/**
 * Webアプリのエントリーポイントとルーター
 * 案件管理、補助金アラート、スプラッシュブロー対応を含むメインロジック
 */
function doGet() {
  const t = HtmlService.createTemplateFromFile('index');
  t.include = function(f) { return HtmlService.createHtmlOutputFromFile(f).getContent(); };
  return t.evaluate()
    .setTitle('洗車機管理システム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 送信元（From）を指定してGmail下書きを作成する。
 * Gmail API が有効な場合は From に config.ADMIN_EMAIL を使用し、無効または失敗時は GmailApp にフォールバックする。
 */
function createDraftWithFrom(to, subject, body) {
  const config = getConfig();
  const fromEmail = config.ADMIN_EMAIL || Session.getActiveUser().getEmail();
  try {
    if (typeof Gmail !== 'undefined' && Gmail.Users && Gmail.Users.Drafts) {
      const mime = 'From: ' + fromEmail + '\r\nTo: ' + to + '\r\nSubject: ' + subject + '\r\n\r\n' + body;
      const raw = Utilities.base64EncodeWebSafe(Utilities.newBlob(mime, 'UTF-8').getBytes());
      Gmail.Users.Drafts.create({ userId: 'me', resource: { message: { raw: raw } } });
      return;
    }
  } catch (e) {
    console.warn('Gmail API で下書き作成失敗、GmailApp にフォールバック:', e);
  }
  GmailApp.createDraft(to, subject, body);
}

function getCarWashList() { return getCarWashListCached(); }

function checkSubsidyAlert(shopName, installDateStr) {
  const config = getConfig();
  const isTargetShop = config.SUBSIDY_TARGET_SHOPS.some(target => shopName.includes(target));
  if (!isTargetShop) return null;

  if (!installDateStr) return null;
  const installDate = new Date(installDateStr);
  const today = new Date();
  
  const limitDate = new Date(installDate);
  limitDate.setFullYear(limitDate.getFullYear() + config.SUBSIDY_YEARS_LIMIT);

  if (today < limitDate) {
    const limitDateStr = Utilities.formatDate(limitDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    const installStr = Utilities.formatDate(installDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    return {
      isAlert: true,
      message: `【補助金注意】設置(${installStr})から${config.SUBSIDY_YEARS_LIMIT}年未満です。返金義務が発生する可能性があります。（解禁日: ${limitDateStr}）`
    };
  }
  return null;
}

function getRecentActionList(scheduleData, config) {
  const actions = [];
  const today = new Date();
  const recentLimit = new Date();
  recentLimit.setDate(today.getDate() - 30); 

  if (scheduleData.length > 1) {
    scheduleData.slice(1).forEach(row => {
      const status = row[5];
      const date = new Date(row[4]);
      const isCompletedRecently = (status === config.PROJECT_STATUS.COMPLETED && date >= recentLimit);
      const isActive = (status !== config.PROJECT_STATUS.COMPLETED && status !== config.PROJECT_STATUS.CANCELLED);
      if (isActive || isCompletedRecently) {
        actions.push(`${row[1]}_${row[2]}_${row[3]}`);
      }
    });
  }
  return actions;
}

function getDashboardData() {
  const data = getCarWashListCached();
  const config = getConfig();
  const sheet = getSheet(config.SHEET_NAMES.SCHEDULE);
  const scheduleData = sheet.getDataRange().getValues();
  const ignoreActions = getRecentActionList(scheduleData, config);

  const notices = [];
  let normal = 0;
  
  const shopsWithBodyExchange = new Set();
  // 店舗単位で本体交換が案件化されているかチェック
  const shopsWithBodyExchangeRegistered = new Set();
  ignoreActions.forEach(action => {
    if (action.endsWith('_本体')) {
      const parts = action.split('_');
      if (parts.length >= 2) {
        const shopCode = parts[0];
        const machineId = parts[1];
        // 「全機」として登録されている場合、その店舗を除外
        if (machineId === '全機') {
          shopsWithBodyExchangeRegistered.add(shopCode);
        }
      }
    }
  });
  
  data.forEach(m => {
    const shopCode = m['店舗コード'];
    // 既に案件化されている店舗は除外
    if (shopsWithBodyExchangeRegistered.has(shopCode)) {
      return;
    }
    const key = `${shopCode}_${m['区別']}_本体`;
    if (m['本体ステータス'] === config.STATUS.PREPARE && !ignoreActions.includes(key)) {
      shopsWithBodyExchange.add(shopCode);
    }
  });
  
  // 店舗単位でグループ化（本体交換の場合は店舗ごとに1件のみ表示）
  const shopGroups = {};
  
  data.forEach(m => {
    const shopCode = m['店舗コード'];
    const keyPrefix = `${m['店舗コード']}_${m['区別']}_`;
    const isBodyExchangeInProgress = ignoreActions.includes(keyPrefix + '本体');

    // 既に案件化されている店舗は除外
    if (shopsWithBodyExchangeRegistered.has(shopCode)) {
      normal++;
      return;
    }
    
    if (shopsWithBodyExchange.has(shopCode)) {
      // 店舗内の1台でも本体が「交換準備」なら、その店舗を本体交換対象として1件のみ表示
      if (!shopGroups[shopCode]) {
        // 店舗内の最初の機械のデータを使用（補助金チェック用）
        const firstMachine = data.find(d => d['店舗コード'] === shopCode);
        const maskedItem = { ...firstMachine };
        maskedItem['レールステータス'] = config.STATUS.NORMAL; 
        maskedItem['ブラシステータス'] = '対象外';
        maskedItem['本体ステータス'] = config.STATUS.PREPARE;
        maskedItem['区別'] = '全機'; // 全機を表す
        const subsidyCheck = checkSubsidyAlert(firstMachine['店舗名'], firstMachine['本体設置日']);
        if (subsidyCheck) maskedItem['subsidyAlert'] = subsidyCheck.message;
        shopGroups[shopCode] = maskedItem;
      }
      // 本体交換対象店舗の機械はnormalにカウントしない
    } else {
      let isRailAlert = m['レールステータス'] === config.STATUS.NOTICE && !ignoreActions.includes(keyPrefix + 'レール車輪');
      if (isBodyExchangeInProgress) isRailAlert = false;
      let isBrushAlert = (m['ブラシステータス'] && m['ブラシステータス'].includes('通知')) && !ignoreActions.includes(keyPrefix + '布ブラシ');
      if (isBodyExchangeInProgress) isBrushAlert = false;

      if (isRailAlert || isBrushAlert) {
        notices.push(m);
      } else { 
        normal++; 
      }
    }
  });
  
  // 店舗単位でグループ化された本体交換対象を追加
  Object.values(shopGroups).forEach(item => {
    notices.push(item);
  });
  const projects = getOngoingProjects();
  return { noticeCount: notices.length, normalCount: normal, noticeList: notices, projects: projects };
}

function getOngoingProjects() {
  const config = getConfig();
  const sheet = getSheet(config.SHEET_NAMES.SCHEDULE);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const carWashList = getCarWashListCached();
  const shopNameMap = {};
  carWashList.forEach(row => { shopNameMap[row['店舗コード']] = row['店舗名']; });
  const targetStatuses = [config.PROJECT_STATUS.ESTIMATE_REQ, config.PROJECT_STATUS.ESTIMATE_RCV, config.PROJECT_STATUS.ORDERED];
  return data.slice(1).map((r, i) => ({
    id: r[0],
    shopCode: r[1],
    shopName: shopNameMap[r[1]] || r[1],
    machineId: r[2],
    part: r[3],
    status: r[5],
    rowNumber: i + 2
  })).filter(p => targetStatuses.includes(p.status));
}

function getAllActiveProjects() {
  const config = getConfig();
  const sheet = getSheet(config.SHEET_NAMES.SCHEDULE);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const carWashList = getCarWashListCached();
  const shopNameMap = {};
  carWashList.forEach(row => { shopNameMap[row['店舗コード']] = row['店舗名']; });
  return data.slice(1).map((r, i) => ({
    id: r[0],
    shopCode: r[1],
    shopName: shopNameMap[r[1]] || r[1],
    machineId: r[2],
    part: r[3],
    date: (r[4] instanceof Date) ? Utilities.formatDate(r[4], Session.getScriptTimeZone(), 'yyyy-MM-dd') : r[4],
    status: r[5],
    rowNumber: i + 2
  })).filter(p => p.status !== config.PROJECT_STATUS.COMPLETED && p.status !== config.PROJECT_STATUS.CANCELLED);
}

function updateProjectStatus(id, newStatus) {
  const config = getConfig();
  const sheet = getSheet(config.SHEET_NAMES.SCHEDULE);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      const row = i + 1;
      const currentStatus = data[i][5];
      const eventId = data[i][6];
      // ステータスを「未完了」以外に戻す場合、登録済みのカレンダーがあれば削除し予定日・カレンダーIDをクリア
      if (currentStatus === config.PROJECT_STATUS.SCHEDULED && eventId && newStatus !== config.PROJECT_STATUS.SCHEDULED) {
        try {
          const calendar = config.CALENDAR_ID === 'primary' ? CalendarApp.getDefaultCalendar() : CalendarApp.getCalendarById(config.CALENDAR_ID);
          if (calendar) {
            const event = calendar.getEventById(eventId);
            if (event) event.deleteEvent();
          }
        } catch (e) { console.warn(e); }
        sheet.getRange(row, 5).clearContent();
        sheet.getRange(row, 7).clearContent();
      }
      sheet.getRange(row, 6).setValue(newStatus);
      return { success: true };
    }
  }
  throw new Error('案件が見つかりません');
}

function cancelProject(id) {
  const config = getConfig();
  const sheet = getSheet(config.SHEET_NAMES.SCHEDULE);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      const row = i + 1;
      const eventId = data[i][6];
      if (eventId) {
        try {
          const calendar = config.CALENDAR_ID === 'primary' ? CalendarApp.getDefaultCalendar() : CalendarApp.getCalendarById(config.CALENDAR_ID);
          if (calendar) {
            const event = calendar.getEventById(eventId);
            if (event) event.deleteEvent();
          }
        } catch (e) { console.warn(e); }
        sheet.getRange(row, 5).clearContent();
        sheet.getRange(row, 7).clearContent();
      }
      sheet.getRange(row, 6).setValue(config.PROJECT_STATUS.CANCELLED);
      return { success: true };
    }
  }
  throw new Error('案件が見つかりません');
}

function saveNextWorkMemo(shopCode, machineIdentifier, memo) {
  const config = getConfig();
  const masterSheet = getSheet(config.SHEET_NAMES.MASTER_EQUIPMENT);
  const data = masterSheet.getDataRange().getValues();
  const rowIndex = data.findIndex(row => row[0] == shopCode && row[2] == machineIdentifier);
  if (rowIndex === -1) throw new Error('マスタが見つかりません');
  masterSheet.getRange(rowIndex + 1, 9).setValue(memo);
  refreshStatusSummary();
  return { success: true };
}

function getExchangeTargetsForUI() {
  const config = getConfig();
  const carWashList = getCarWashListCached();
  const scheduleSheet = getSheet(config.SHEET_NAMES.SCHEDULE);
  const scheduleData = scheduleSheet.getDataRange().getValues();
  const ignoreActions = getRecentActionList(scheduleData, config);
  const shopsWithBodyExchange = new Set();
  
  carWashList.forEach(m => {
    const key = `${m['店舗コード']}_${m['区別']}_本体`;
    if (m['本体ステータス'] === config.STATUS.PREPARE && !ignoreActions.includes(key)) {
      shopsWithBodyExchange.add(m['店舗コード']);
    }
  });

  return carWashList.map(m => {
    const parts = [];
    const keyPrefix = `${m['店舗コード']}_${m['区別']}_`;
    const shopCode = m['店舗コード'];
    
    let subsidyInfo = null;
    if (m['本体ステータス'] === config.STATUS.PREPARE) {
       const check = checkSubsidyAlert(m['店舗名'], m['本体設置日']);
       if (check) subsidyInfo = check.message;
    }

    const nextWorkMemo = m['nextWorkMemo'] || "";
    const isBodyExchangeInProgress = ignoreActions.includes(keyPrefix + '本体');

    if (shopsWithBodyExchange.has(shopCode)) {
      // 店舗内の1台でも本体が「交換準備」なら、その店舗の全機を本体交換対象とする
      if (!isBodyExchangeInProgress) {
        parts.push('本体');
      }
    } else {
      if (m['レールステータス'] === config.STATUS.NOTICE && !ignoreActions.includes(keyPrefix + 'レール車輪')) {
        if (!isBodyExchangeInProgress) parts.push('レール車輪');
      }
      if (m['ブラシステータス'] && m['ブラシステータス'].includes('通知') && !ignoreActions.includes(keyPrefix + '布ブラシ')) {
        if (!isBodyExchangeInProgress) parts.push('布ブラシ');
      }
      if (m['本体ステータス'] === config.STATUS.PREPARE && !ignoreActions.includes(keyPrefix + '本体')) {
        parts.push('本体');
      }
    }

    if (parts.length === 0) return null;
    return { 
      shopCode: m['店舗コード'], 
      shopName: m['店舗名'], 
      machineIdentifier: m['区別'], 
      exchangeTargets: parts.join('/'), 
      allExchangeParts: parts,
      subsidyAlert: subsidyInfo,
      nextWorkMemo: nextWorkMemo
    };
  }).filter(item => item !== null);
}

function registerProject(s, m, t, status) {
  const sheet = getSheet(getConfig().SHEET_NAMES.SCHEDULE);
  const uniqueId = 'ID-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
  sheet.appendRow([uniqueId, s, m, t, '', status, '', '']);
  return { success: true, id: uniqueId };
}

/**
 * 選択された機械を案件リストに登録（見積依頼中）
 * @param {Array} selectedData - [{ shopCode, shopName, machineId, parts: ['レール車輪', '布ブラシ'], nextWorkMemo }]
 * @returns {Object} { registeredCount: number }
 */
function registerSelectedProjects(selectedData) {
  try {
    if (!selectedData || !Array.isArray(selectedData) || selectedData.length === 0) {
      throw new Error('選択データが無効です');
    }
    
    const config = getConfig();
    const sheet = getSheet(config.SHEET_NAMES.SCHEDULE);
    const scheduleData = sheet.getDataRange().getValues();
    
    // 既存のアクティブな案件を取得（重複登録防止）
    const activeProjects = new Set();
    if (scheduleData.length > 1) {
      scheduleData.slice(1).forEach(row => {
        if (row[5] !== config.PROJECT_STATUS.COMPLETED && row[5] !== config.PROJECT_STATUS.CANCELLED) {
          activeProjects.add(`${row[1]}_${row[2]}_${row[3]}`);
        }
      });
    }
    
    let registeredCount = 0;
    
    selectedData.forEach(item => {
      if (!item.shopCode || !item.machineId || !item.parts || !Array.isArray(item.parts)) {
        console.warn('無効なデータ項目をスキップ:', item);
        return;
      }
      
      item.parts.forEach(part => {
        if (!part) return;
        
        const key = `${item.shopCode}_${item.machineId}_${part}`;
        
        // 重複チェック
        if (!activeProjects.has(key)) {
          try {
            registerProject(item.shopCode, item.machineId, part, config.PROJECT_STATUS.ESTIMATE_REQ);
            activeProjects.add(key); // 同じバッチ内での重複も防止
            registeredCount++;
          } catch (e) {
            console.error('案件登録エラー:', e, item, part);
            throw new Error(`案件登録に失敗しました: ${item.shopCode}_${item.machineId}_${part} - ${e.message}`);
          }
        }
      });
    });
    
    return { registeredCount: registeredCount };
  } catch (e) {
    console.error('registerSelectedProjects エラー:', e);
    throw e;
  }
}

/**
 * 選択された機械のメール下書き作成 + 案件登録
 * @param {Array} selectedData - [{ shopCode, shopName, machineId, parts: ['レール車輪', '布ブラシ'], nextWorkMemo }]
 * @returns {Object} { registeredCount: number, draftCount: number }
 */
function createDraftsForSelected(selectedData) {
  try {
    if (!selectedData || !Array.isArray(selectedData) || selectedData.length === 0) {
      throw new Error('選択データが無効です');
    }
    
    const config = getConfig();
    const sheet = getSheet(config.SHEET_NAMES.SCHEDULE);
    const scheduleData = sheet.getDataRange().getValues();
    
    // 既存のアクティブな案件を取得（重複登録防止）
    const activeProjects = new Set();
    if (scheduleData.length > 1) {
      scheduleData.slice(1).forEach(row => {
        if (row[5] !== config.PROJECT_STATUS.COMPLETED && row[5] !== config.PROJECT_STATUS.CANCELLED) {
          activeProjects.add(`${row[1]}_${row[2]}_${row[3]}`);
        }
      });
    }
    
    // 宛先定義
    const RECIPIENT_DAIFUKU = "nishimura@selfix.jp"; // 本来は木村様宛
    const RECIPIENT_BEAUTY = "nishimura@selfix.jp";  // 本来は松永様宛
    
    // 業者別の依頼事項リスト
    const daifukuItems = [];
    const beautyItems = [];
    
    let registeredCount = 0;
    
    // 店舗ごとにグループ化
    const shopGroups = {};
    selectedData.forEach(item => {
      if (!item.shopCode || !item.shopName || !item.machineId || !item.parts || !Array.isArray(item.parts)) {
        console.warn('無効なデータ項目をスキップ:', item);
        return;
      }
    const shopCode = item.shopCode;
    if (!shopGroups[shopCode]) {
      shopGroups[shopCode] = {
        name: item.shopName,
        machines: [],
        subsidyInfo: null,
        nextWorkMemo: item.nextWorkMemo || ""
      };
      
      // 補助金チェック（洗車機リストから取得）
      const carWashList = getCarWashListCached();
      const machineData = carWashList.find(m => m['店舗コード'] === shopCode);
      if (machineData && machineData['本体ステータス'] === config.STATUS.PREPARE) {
        const check = checkSubsidyAlert(machineData['店舗名'], machineData['本体設置日']);
        if (check) shopGroups[shopCode].subsidyInfo = check.message;
      }
    }
    
      // 各機械の部品を処理
      item.parts.forEach(part => {
        if (!part) return;
        
        const key = `${item.shopCode}_${item.machineId}_${part}`;
        
        // 重複チェック
        if (!activeProjects.has(key)) {
          try {
            // 案件登録
            registerProject(item.shopCode, item.machineId, part, config.PROJECT_STATUS.ESTIMATE_REQ);
            activeProjects.add(key);
            registeredCount++;
          } catch (e) {
            console.error('案件登録エラー:', e, item, part);
            throw new Error(`案件登録に失敗しました: ${item.shopCode}_${item.machineId}_${part} - ${e.message}`);
          }
        
        // メール項目の追加
        let displayPartName = part;
        if (displayPartName === 'レール車輪') {
          displayPartName = 'レール車輪等';
        }
        
        if (part === '本体') {
          // 本体交換 → ダイフク案件
          daifukuItems.push({
            shopName: item.shopName,
            machineId: '全機',
            content: '洗車機・クリーナー・マット洗い機 本体入れ替え',
            memo: shopGroups[shopCode].subsidyInfo || "",
            userMemo: item.nextWorkMemo || ""
          });
          
          // スプラッシュブロー（特定の店舗以外）→ ビユーテー案件
          const excludedShops = ['東和歌山', '糸我', '貴志川'];
          if (!excludedShops.includes(item.shopName)) {
            const isSbAlreadyActive = Array.from(activeProjects).some(p => p.includes(shopCode) && p.includes('スプラッシュブロー'));
            if (!isSbAlreadyActive) {
              beautyItems.push({
                shopName: item.shopName,
                machineId: 'SB1',
                content: 'スプラッシュブローSB1 入れ替え',
                memo: '',
                userMemo: ''
              });
              registerProject(shopCode, 'SB1', 'スプラッシュブロー', config.PROJECT_STATUS.ESTIMATE_REQ);
              registeredCount++;
            }
          }
        } else if (part === 'スプラッシュブロー') {
          // スプラッシュブロー → ビユーテー案件
          beautyItems.push({
            shopName: item.shopName,
            machineId: item.machineId,
            content: 'スプラッシュブローSB1 入れ替え',
            memo: '',
            userMemo: item.nextWorkMemo || ""
          });
        } else {
          // 部品交換 → ダイフク案件
          daifukuItems.push({
            shopName: item.shopName,
            machineId: item.machineId,
            content: `${displayPartName} 交換`,
            memo: "",
            userMemo: item.nextWorkMemo || ""
          });
        }
      }
    });
  });
  
    // メール下書き作成（ダイフク：木村様）
    let draftCount = 0;
    if (daifukuItems.length > 0) {
      try {
        const subject = `【見積依頼】洗車機関連案件のお見積り依頼（${daifukuItems.length}件）`;
        const body = generateConsolidatedTemplate(daifukuItems, '木村様');
        createDraftWithFrom(RECIPIENT_DAIFUKU, subject, body);
        draftCount++;
      } catch (e) {
        console.error('ダイフクメール作成エラー:', e);
        throw new Error(`メール下書き作成に失敗しました（ダイフク）: ${e.message}`);
      }
    }
    
    // メール下書き作成（ビユーテー：松永様）
    if (beautyItems.length > 0) {
      try {
        const subject = `【見積依頼】スプラッシュブロー関連案件のお見積り依頼（${beautyItems.length}件）`;
        const body = generateConsolidatedTemplate(beautyItems, 'ビユーテー\n松永様');
        createDraftWithFrom(RECIPIENT_BEAUTY, subject, body);
        draftCount++;
      } catch (e) {
        console.error('ビユーテーメール作成エラー:', e);
        throw new Error(`メール下書き作成に失敗しました（ビユーテー）: ${e.message}`);
      }
    }
    
    return { registeredCount: registeredCount, draftCount: draftCount };
  } catch (e) {
    console.error('createDraftsForSelected エラー:', e);
    throw e;
  }
}

function createScheduleAndRecord(s, m, t, d, n, existingId = null) { 
  const config = getConfig();
  const r = createExchangeEvent(s, m, t, d, n); 
  const sheet = getSheet(config.SHEET_NAMES.SCHEDULE);
  
  if (existingId) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(existingId)) {
        const row = i + 1;
        sheet.getRange(row, 5).setValue(d);
        sheet.getRange(row, 6).setValue(config.PROJECT_STATUS.SCHEDULED);
        sheet.getRange(row, 7).setValue(r.eventId);
        sheet.getRange(row, 8).setValue(n);
        return r;
      }
    }
  }

  const uniqueId = 'ID-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
  sheet.appendRow([uniqueId, s, m, t, d, config.PROJECT_STATUS.SCHEDULED, r.eventId, n]);
  return r; 
}

function getScheduledTasksForCompletion() {
  const config = getConfig();
  const vals = getSheet(config.SHEET_NAMES.SCHEDULE).getDataRange().getValues();
  if (vals.length <= 1) return [];
  const carWashList = getCarWashListCached();
  const shopNameMap = {};
  carWashList.forEach(row => { shopNameMap[row['店舗コード']] = row['店舗名']; });
  return vals.slice(1).map((r, i) => {
    return { 
      rowNumber: i + 2, 
      'id': r[0],
      '店舗コード': r[1], 
      '店舗名': shopNameMap[r[1]] || '名称不明',
      '区別': r[2], 
      '部品名': r[3], 
      '予定日': (r[4] instanceof Date) ? Utilities.formatDate(r[4], Session.getScriptTimeZone(), 'yyyy/MM/dd') : r[4], 
      'ステータス': r[5],
      'eventId': r[6]
    };
  }).filter(t => t['ステータス'] === config.PROJECT_STATUS.SCHEDULED);
}

function completeExchange(uniqueId, date, count) {
  const config = getConfig();
  const sheet = getSheet(config.SHEET_NAMES.SCHEDULE);
  const data = sheet.getDataRange().getValues();
  let targetRowIndex = -1;
  let targetData = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(uniqueId)) {
      targetRowIndex = i + 1;
      targetData = data[i];
      break;
    }
  }
  if (targetRowIndex === -1) throw new Error('対象のスケジュールが見つかりません。');
  recordExchangeComplete(targetData[1], targetData[2], targetData[3], date, count);
  markEventAsCompleted(targetData[6], date);
  sheet.getRange(targetRowIndex, 6).setValue(config.PROJECT_STATUS.COMPLETED);
  return { message: '完了報告成功' };
}

function cancelSchedule(rowNumber, eventId) {
  const config = getConfig();
  const sheet = getSheet(config.SHEET_NAMES.SCHEDULE);
  if (eventId) {
    try {
      const calendar = config.CALENDAR_ID === 'primary' ? CalendarApp.getDefaultCalendar() : CalendarApp.getCalendarById(config.CALENDAR_ID);
      if (calendar) {
        const event = calendar.getEventById(eventId);
        if (event) event.deleteEvent();
      }
    } catch (e) { console.warn(e); }
  }
  sheet.getRange(rowNumber, 6).setValue(config.PROJECT_STATUS.CANCELLED);
  return { success: true };
}

/**
 * 【改修】アラート一括メール作成
 * ダイフク（木村様）とビユーテー（松永様）に分けてメールを作成します。
 */
function createDraftsForAlerts() {
  const config = getConfig();
  const list = getCarWashListCached();
  
  // 作成ログ保存用
  const createdLog = [];
  
  // ★宛先定義
  const RECIPIENT_DAIFUKU = "nishimura@selfix.jp"; // 本来は木村様宛
  const RECIPIENT_BEAUTY = "nishimura@selfix.jp";  // 本来は松永様宛
  
  // 現在進行中のプロジェクトを取得（重複登録防止）
  const sheet = getSheet(config.SHEET_NAMES.SCHEDULE);
  const scheduleData = sheet.getDataRange().getValues();
  const activeProjects = [];
  if (scheduleData.length > 1) {
    scheduleData.slice(1).forEach(row => {
      if (row[5] !== config.PROJECT_STATUS.COMPLETED && row[5] !== config.PROJECT_STATUS.CANCELLED) {
        activeProjects.push(`${row[1]}_${row[2]}_${row[3]}`);
      }
    });
  }

  // 業者別の依頼事項リスト
  const daifukuItems = [];
  const beautyItems = [];

  // 店舗ごとにグループ化してチェック
  const shopGroups = {};
  list.forEach(m => {
    const shopCode = m['店舗コード'];
    if (!shopGroups[shopCode]) {
      shopGroups[shopCode] = { 
        name: m['店舗名'], 
        hasBodyExchange: false, 
        partExchanges: [], 
        subsidyInfo: null, 
        nextWorkMemo: m['nextWorkMemo'] || "" 
      };
    }
    
    // 補助金チェック
    if (m['本体ステータス'] === config.STATUS.PREPARE) {
      const check = checkSubsidyAlert(m['店舗名'], m['本体設置日']);
      if (check) shopGroups[shopCode].subsidyInfo = check.message;
    }

    const keyPrefix = `${m['店舗コード']}_${m['区別']}_`;
    const isBodyExchangeInProgress = activeProjects.includes(keyPrefix + '本体');

    // 本体交換判定
    if (m['本体ステータス'] === config.STATUS.PREPARE && !isBodyExchangeInProgress) {
      shopGroups[shopCode].hasBodyExchange = true;
    }

    // 部品交換判定（本体交換がない場合のみ）
    if (m['レールステータス'] === config.STATUS.NOTICE && !activeProjects.includes(keyPrefix + 'レール車輪')) {
      if (!isBodyExchangeInProgress) shopGroups[shopCode].partExchanges.push({ machineId: m['区別'], part: 'レール車輪' });
    }
    if (m['ブラシステータス'] && m['ブラシステータス'].includes('通知') && !activeProjects.includes(keyPrefix + '布ブラシ')) {
      if (!isBodyExchangeInProgress) shopGroups[shopCode].partExchanges.push({ machineId: m['区別'], part: '布ブラシ' });
    }
  });

  // リクエストリストの作成
  Object.keys(shopGroups).forEach(shopCode => {
    const group = shopGroups[shopCode];
    
    // 本体交換がある場合
    if (group.hasBodyExchange) {
      // ダイフク案件（本体）
      daifukuItems.push({
        shopName: group.name,
        machineId: '全機',
        content: '洗車機・クリーナー・マット洗い機 本体入れ替え',
        memo: group.subsidyInfo || "",
        userMemo: group.nextWorkMemo
      });
      registerProject(shopCode, '全機', '本体', config.PROJECT_STATUS.ESTIMATE_REQ);
      createdLog.push({ shopName: group.name, machineId: '全機', part: '本体入れ替え' });

      // スプラッシュブロー（特定の店舗以外）→ ビユーテー案件
      const excludedShops = ['東和歌山', '糸我', '貴志川'];
      if (!excludedShops.includes(group.name)) {
        const isSbAlreadyActive = activeProjects.some(p => p.includes(shopCode) && p.includes('スプラッシュブロー'));
        if (!isSbAlreadyActive) {
          beautyItems.push({
            shopName: group.name,
            machineId: 'SB1',
            content: 'スプラッシュブローSB1 入れ替え',
            memo: '',
            userMemo: ''
          });
          registerProject(shopCode, 'SB1', 'スプラッシュブロー', config.PROJECT_STATUS.ESTIMATE_REQ);
          createdLog.push({ shopName: group.name, machineId: 'SB1', part: 'スプラッシュブロー' });
        }
      }
    } else {
      // 部品交換のみの場合（現状は全てダイフクと仮定）
      group.partExchanges.forEach(item => {
        let displayPartName = item.part;
        if (displayPartName === 'レール車輪') {
          displayPartName = 'レール車輪等';
        }

        daifukuItems.push({
          shopName: group.name,
          machineId: item.machineId,
          content: `${displayPartName} 交換`,
          memo: "",
          userMemo: group.nextWorkMemo
        });
        registerProject(shopCode, item.machineId, item.part, config.PROJECT_STATUS.ESTIMATE_REQ);
        createdLog.push({ shopName: group.name, machineId: item.machineId, part: item.part });
      });
    }
  });

  // メール下書き作成（ダイフク：木村様）
  if (daifukuItems.length > 0) {
    const subject = `【見積依頼】洗車機関連案件のお見積り依頼（${daifukuItems.length}件）`;
    const body = generateConsolidatedTemplate(daifukuItems, '木村様');
    createDraftWithFrom(RECIPIENT_DAIFUKU, subject, body);
  }

  // メール下書き作成（ビユーテー：松永様）
  if (beautyItems.length > 0) {
    const subject = `【見積依頼】スプラッシュブロー関連案件のお見積り依頼（${beautyItems.length}件）`;
    const body = generateConsolidatedTemplate(beautyItems, 'ビユーテー\n松永様');
    createDraftWithFrom(RECIPIENT_BEAUTY, subject, body);
  }

  return createdLog;
}

/**
 * 統合メールテンプレート（汎用）
 */
function generateConsolidatedTemplate(items, recipientName) {
  let itemText = "";
  
  items.forEach(item => {
    let alertText = "";
    if (item.memo) alertText = `\n  ※注意: ${item.memo}`;
    if (item.userMemo) alertText += `\n  ※付帯作業メモ: ${item.userMemo}`;
    
    itemText += `--------------------------------------------------\n`;
    itemText += `■店舗名：セルフィックス${item.shopName}SS\n`;
    itemText += `■対象機：${item.machineId}\n`;
    itemText += `■依頼内容：${item.content}${alertText}\n`;
    itemText += `--------------------------------------------------\n\n`;
  });

  return `${recipientName}\n\n` +
         `いつも大変お世話になっております。\n` +
         `日商有田株式会社の西村です。\n\n` +
         `表題の件、以下の対応が必要となりました。\n` +
         `下記内容にてお見積りを作成いただけますでしょうか。\n\n` +
         `【ご依頼案件一覧】\n\n` +
         itemText +
         `本件に関するお見積書、実施スケジュール、および作業完了報告書につきましては、\n` +
         `メールでのご送付をお願いします。\n\n` +
         `よろしくお願いします。\n\n` +
         `--------------------------------------------------\n` +
         `日商有田株式会社\n` +
         `西村\n` +
         `--------------------------------------------------`;
}

// -------------------------------------------------------------
// 手動作成用テンプレート関数
// -------------------------------------------------------------

function generateBodyReplacementTemplate(shopName) {
  const itemText = 
         `【対象店舗】\nセルフィックス${shopName}SS\n\n` +
         `【依頼内容】\n洗車機・クリーナー・マット洗い機入れ替え\n\n`;

  return `木村様\n\n` +
         `いつも大変お世話になっております。\n` +
         `日商有田株式会社の西村です。\n\n` +
         `表題の件、以下の対応が必要となりました。\n` +
         `下記内容にてお見積りを作成いただけますでしょうか。\n\n` +
         itemText +
         `本件に関するお見積書、実施スケジュール、および作業完了報告書につきましては、\n` +
         `メールでのご送付をお願いします。\n\n` +
         `よろしくお願いします。\n\n` +
         `--------------------------------------------------\n` +
         `日商有田株式会社\n` +
         `西村\n` +
         `--------------------------------------------------`;
}

function generateSplashBlowTemplate(shopName) {
  // ★宛先変更：ビユーテー松永様
  const itemText = 
         `【対象店舗】\nセルフィックス${shopName}SS\n\n` +
         `【依頼内容】\nスプラッシュブローＳＢ１　1台\n\n`;

  return `ビユーテー\n松永様\n\n` +
         `いつも大変お世話になっております。\n` +
         `日商有田株式会社の西村です。\n\n` +
         `表題の件、以下の対応が必要となりました。\n` +
         `下記内容にてお見積りを作成いただけますでしょうか。\n\n` +
         itemText +
         `本件に関するお見積書および作業完了報告書につきましては、\n` +
         `メールでのご送付をお願いします。\n\n` +
         `よろしくお願いします。\n\n` +
         `--------------------------------------------------\n` +
         `日商有田株式会社\n` +
         `西村\n` +
         `--------------------------------------------------`;
}

function generatePartsExchangeTemplate(shopName, machineId, partName) {
  // レール車輪等の名称調整
  let displayPartName = partName;
  if (displayPartName.includes('レール車輪') && !displayPartName.includes('等')) {
      displayPartName = displayPartName.replace('レール車輪', 'レール車輪等');
  }

  const itemText = 
         `【対象店舗】\nセルフィックス${shopName}SS\n\n` +
         `【対象機】\n${machineId}\n\n` +
         `【交換必要部品】\n${displayPartName}\n\n`;

  return `木村様\n\n` +
         `いつも大変お世話になっております。\n` +
         `日商有田株式会社の西村です。\n\n` +
         `表題の件、以下の対応が必要となりました。\n` +
         `下記内容にてお見積りを作成いただけますでしょうか。\n\n` +
         itemText +
         `本件に関するお見積書、実施スケジュール、および作業完了報告書につきましては、\n` +
         `メールでのご送付をお願いします。\n\n` +
         `よろしくお願いします。\n\n` +
         `--------------------------------------------------\n` +
         `日商有田株式会社\n` +
         `西村\n` +
         `--------------------------------------------------`;
}

function generateOrderTemplate(shopName, machineId, partName) {
  // 発注時も同様に調整
  let displayPartName = partName;
  if (displayPartName.includes('レール車輪') && !displayPartName.includes('等')) {
      displayPartName = displayPartName.replace('レール車輪', 'レール車輪等');
  }

  const itemText = 
         `【店舗名】セルフィックス${shopName}SS\n` +
         `【対象機】${machineId}\n` +
         `【発注内容】${displayPartName}\n\n`;

  return `木村様\n\n` +
         `いつも大変お世話になっております。\n` +
         `日商有田株式会社の西村です。\n\n` +
         `表題の件、以下の内容にて正式に発注いたします。\n\n` +
         itemText +
         `いただいたお見積り内容にて進めていただけますでしょうか。\n` +
         `本件に関するお見積書、実施スケジュール、および作業完了報告書につきましては、\n` +
         `メールでのご送付をお願いします。\n\n` +
         `よろしくお願いします。\n\n` +
         `--------------------------------------------------\n` +
         `日商有田株式会社\n` +
         `西村\n` +
         `--------------------------------------------------`;
}

function generateQuoteRequest(shopName, machineId, partName) {
  // 手動作成時用ルーター
  if (partName.includes('本体')) return generateBodyReplacementTemplate(shopName);
  if (partName.includes('スプラッシュブロー')) return generateSplashBlowTemplate(shopName);
  return generatePartsExchangeTemplate(shopName, machineId, partName);
}

function generateOrder(shopName, machineId, partName) {
  return generateOrderTemplate(shopName, machineId, partName);
}