/**
 * 設定ファイル
 * スプレッドシートIDやシート名、交換基準などを一元管理します。
 */
function getConfig() {
  // ★重要: 移行後のスプレッドシートID
  const SPREADSHEET_ID = '1VMZ2RpYmCh0BZRwY4DPzriAafFOo-U9PqvOUZu67cf4'; 
  
  const SHEET_NAMES = {
    MASTER_EQUIPMENT: '設備マスタ',
    MASTER_STORE: '店舗マスタ',
    MONTHLY_DATA: '月次データ',
    SCHEDULE: '交換スケジュール',
    HISTORY: '交換履歴',
    STATUS_SUMMARY: 'ステータス集計'
  };

  // 交換基準
  const EXCHANGE_CRITERIA = {
    RAIL_WHEEL: { THRESHOLD: 55000, MONTHS_WARNING: 18 },
    CLOTH_BRUSH: { FIRST: 35000, SECOND: 70000, MONTHS_WARNING: 18 },
    MACHINE_BODY: { THRESHOLD: 100000, FORECAST_MONTHS: 15 }
  };

  // ★重要: 補助金対象店舗リスト（部分一致で判定します）
  // 今後、対象店舗が増えた場合は、ここに店舗名を追加してください。
  // 記述例: ['徳島石井', '小松島', '紀三井寺', '坂出', '新しい店舗名']
  const SUBSIDY_TARGET_SHOPS = ['徳島石井', '小松島', '紀三井寺', '坂出'];

  // ★追加: 補助金縛りの年数（変更があればここを修正）
  const SUBSIDY_YEARS_LIMIT = 8;

  const STATUS = {
    NORMAL: '正常',
    NOTICE: '通知',
    PREPARE: '交換準備',
    EXCHANGE: '1回目通知',
    EXCHANGE_2: '2回目通知'
  };

  const PROJECT_STATUS = {
    ESTIMATE_REQ: '見積依頼中',
    ESTIMATE_RCV: '見積受領',
    ORDERED: '発注済',
    SCHEDULED: '未完了',
    COMPLETED: '完了',
    CANCELLED: '取り消し'
  };

  const CALENDAR_ID = 'primary'; 
  const ADMIN_EMAIL = 'nishimura@selfix.jp';

  return { 
    SPREADSHEET_ID, SHEET_NAMES, EXCHANGE_CRITERIA, 
    STATUS, PROJECT_STATUS, CALENDAR_ID, ADMIN_EMAIL,
    SUBSIDY_TARGET_SHOPS, SUBSIDY_YEARS_LIMIT 
  };
}

function getSheet(sheetName) {
  const config = getConfig();
  const ss = SpreadsheetApp.openById(config.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`エラー: シート "${sheetName}" が見つかりません。`);
  return sheet;
}