/**
 * カレンダー連携
 */
const EXCHANGE_CALENDAR_ID = getConfig().CALENDAR_ID;

function createExchangeEvent(shopCode, machineIdentifier, exchangeType, dateString, notes) {
  try {
    const list = getCarWashListCached();
    const target = list.find(d => d['店舗コード'] == shopCode && d['区別'] == machineIdentifier);
    const shopName = target ? target['店舗名'] : '店舗情報なし';
    
    const eventDate = new Date(dateString); 
    const calendar = CalendarApp.getCalendarById(EXCHANGE_CALENDAR_ID) || CalendarApp.getDefaultCalendar();
    
    const title = `【洗車機交換予定】${shopName} ${machineIdentifier} - ${exchangeType}`;
    const description = `店舗:${shopName}\n洗車機:${machineIdentifier}\n部品:${exchangeType}\n備考:${notes}`;
    
    const event = calendar.createAllDayEvent(title, eventDate, { description: description });
    
    return { message: `カレンダー登録完了`, eventId: event.getId() };
  } catch (e) { 
    console.error(e);
    throw new Error('カレンダー登録失敗: ' + e.message); 
  }
}

function markEventAsCompleted(eventId, completionDate) {
  try {
    const calendar = CalendarApp.getCalendarById(EXCHANGE_CALENDAR_ID) || CalendarApp.getDefaultCalendar();
    const event = calendar.getEventById(eventId);
    if(event) {
      event.setTitle(`【完了】${event.getTitle().replace('【洗車機交換予定】', '')}`);
      // 色変更はカレンダーの権限によるが、タイトル変更で完了を示す
    }
    return { message: '完了マーク済み' };
  } catch (e) { 
    // イベントが見つからない場合等は無視して進める
    console.warn('カレンダーイベント更新失敗: ' + e.message);
    return { message: 'カレンダー更新スキップ' }; 
  }
}