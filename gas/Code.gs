/**
 * 増員 CONTROL HUB 2.0 - Google Apps Script
 * フォーム → カレンダー → Web App API連携
 */

// ============================================
// 設定
// ============================================
const CONFIG = {
  CALENDAR_ID: 'e3042faeef11a9b013da1351e2251f47b3666e0e88c439ae15afc43962c16ffc@group.calendar.google.com',
  
  COLORS: {
    MULTIPLE: CalendarApp.EventColor.RED,
    STAGE: CalendarApp.EventColor.GREEN,
    SOUND: CalendarApp.EventColor.BLUE,
    LIGHTING: CalendarApp.EventColor.YELLOW
  },
  
  EMOJI: {
    MULTIPLE: '🔴',
    STAGE: '🟢',
    SOUND: '🔵',
    LIGHTING: '🟡'
  }
};

// ============================================
// スタッフ名簿設定
// ============================================
const STAFF_CONFIG = {
  SPREADSHEET_ID: '1sV3fQyTMB1jThs1s2VUqT4ZyM2ISUTaaqPgXO4Z-93A',
  SHEET_NAME: '必要事項の回答',
  NAME_COLUMN: 6
};

// ============================================
// スプレッドシート列インデックス
// ============================================
const COL = {
  TIMESTAMP: 0,
  DATE_1: 1,
  EVENT_NAME: 2,
  TIME_SLOT_1: 3,
  HALL: 4,
  STAGE_1: 5,
  SOUND_1: 6,
  LIGHTING_1: 7,
  CONTENT: 8,
  HAS_DAY2: 9,
  DATE_2: 10,
  TIME_SLOT_2: 11,
  STAGE_2: 12,
  SOUND_2: 13,
  LIGHTING_2: 14,
  HAS_DAY3: 15,
  CONSECUTIVE_2: 16,
  DATE_3: 17,
  TIME_SLOT_3: 18,
  STAGE_3: 19,
  SOUND_3: 20,
  LIGHTING_3: 21,
  CONSECUTIVE_3: 22,
  JUNIOR_OK: 23
};

// ============================================
// フォーム送信時のトリガー
// ============================================
function onFormSubmit(e) {
  try {
    const responses = e.values;
    const eventData = parseFormResponse(responses);
    
    if (eventData.days.length === 0) {
      Logger.log('有効な日付がありません');
      return;
    }
    
    if (eventData.days.length > 1 && areDatesConsecutive(eventData.days.map(d => d.date))) {
      createMultiDayEvent(eventData);
    } else {
      eventData.days.forEach((day, index) => {
        createSingleDayEvent(eventData, day, index + 1);
      });
    }
    
    Logger.log('カレンダーイベント作成完了');
    
    // 該当セクションのスタッフにLINE通知を送信
    try {
      notifyNewRecruitment(eventData);
    } catch (notifyError) {
      Logger.log('新規募集通知エラー: ' + notifyError.toString());
    }
  } catch (error) {
    Logger.log('エラー: ' + error.toString());
  }
}

// ============================================
// フォーム回答をパース
// ============================================
function parseFormResponse(responses) {
  const days = [];
  
  const date1 = parseDate(responses[COL.DATE_1]);
  if (date1) {
    days.push({
      date: date1,
      timeSlot: responses[COL.TIME_SLOT_1] || '',
      stage: parseSectionCount(responses[COL.STAGE_1]),
      sound: parseSectionCount(responses[COL.SOUND_1]),
      lighting: parseSectionCount(responses[COL.LIGHTING_1])
    });
  }
  
  const hasDay2 = responses[COL.HAS_DAY2] && responses[COL.HAS_DAY2] !== 'いいえ';
  if (hasDay2) {
    const date2 = parseDate(responses[COL.DATE_2]);
    if (date2) {
      days.push({
        date: date2,
        timeSlot: responses[COL.TIME_SLOT_2] || '',
        stage: parseSectionCount(responses[COL.STAGE_2]),
        sound: parseSectionCount(responses[COL.SOUND_2]),
        lighting: parseSectionCount(responses[COL.LIGHTING_2])
      });
    }
  }
  
  const hasDay3 = responses[COL.HAS_DAY3] && responses[COL.HAS_DAY3] !== 'いいえ';
  if (hasDay3) {
    const date3 = parseDate(responses[COL.DATE_3]);
    if (date3) {
      days.push({
        date: date3,
        timeSlot: responses[COL.TIME_SLOT_3] || '',
        stage: parseSectionCount(responses[COL.STAGE_3]),
        sound: parseSectionCount(responses[COL.SOUND_3]),
        lighting: parseSectionCount(responses[COL.LIGHTING_3])
      });
    }
  }
  
  let consecutivePreference = 'なし';
  if (days.length === 3 && responses[COL.CONSECUTIVE_3]) {
    consecutivePreference = responses[COL.CONSECUTIVE_3];
  } else if (days.length === 2 && responses[COL.CONSECUTIVE_2]) {
    consecutivePreference = responses[COL.CONSECUTIVE_2];
  }
  
  return {
    eventName: responses[COL.EVENT_NAME] || '',
    hall: responses[COL.HALL] || '',
    content: responses[COL.CONTENT] || '',
    juniorOk: responses[COL.JUNIOR_OK] || '',
    consecutivePreference: consecutivePreference,
    days: days
  };
}

function parseDate(dateStr) {
  if (!dateStr) return null;
  const date = new Date(dateStr);
  return isNaN(date.getTime()) ? null : date;
}

function parseSectionCount(value) {
  if (!value || value === 'なし' || value === '') return 0;
  const num = parseInt(value, 10);
  return isNaN(num) ? 0 : num;
}

function areDatesConsecutive(dates) {
  if (dates.length <= 1) return true;
  const sorted = dates.slice().sort((a, b) => a.getTime() - b.getTime());
  for (let i = 1; i < sorted.length; i++) {
    const diff = (sorted[i].getTime() - sorted[i-1].getTime()) / (1000 * 60 * 60 * 24);
    if (Math.round(diff) !== 1) return false;
  }
  return true;
}

function getSectionColorAndEmoji(days) {
  let hasStage = false, hasSound = false, hasLighting = false;
  days.forEach(day => {
    if (day.stage > 0) hasStage = true;
    if (day.sound > 0) hasSound = true;
    if (day.lighting > 0) hasLighting = true;
  });
  const sectionCount = [hasStage, hasSound, hasLighting].filter(Boolean).length;
  if (sectionCount >= 2) return { color: CONFIG.COLORS.MULTIPLE, emoji: CONFIG.EMOJI.MULTIPLE };
  if (hasStage) return { color: CONFIG.COLORS.STAGE, emoji: CONFIG.EMOJI.STAGE };
  if (hasSound) return { color: CONFIG.COLORS.SOUND, emoji: CONFIG.EMOJI.SOUND };
  if (hasLighting) return { color: CONFIG.COLORS.LIGHTING, emoji: CONFIG.EMOJI.LIGHTING };
  return { color: CONFIG.COLORS.STAGE, emoji: CONFIG.EMOJI.STAGE };
}

function buildDescription(eventData, days) {
  const lines = [];
  if (eventData.juniorOk === 'はい') lines.push('🔰若手可🔰');
  
  days.forEach((day, index) => {
    if (days.length > 1) {
      if (index > 0) lines.push('');
      lines.push(`◆${index + 1}日目　${formatDate(day.date)}`);
    }
    lines.push('【利用区分】');
    lines.push(`　${day.timeSlot || '未設定'}`);
    lines.push('');
    lines.push('【募集セクション】');
    if (day.stage > 0) lines.push(`　舞台: ${day.stage}人`);
    if (day.sound > 0) lines.push(`　音響: ${day.sound}人`);
    if (day.lighting > 0) lines.push(`　照明: ${day.lighting}人`);
  });
  
  if (eventData.eventName) {
    lines.push('');
    lines.push('【催事名】');
    lines.push(`　${eventData.eventName}`);
  }
  
  if (eventData.content) {
    lines.push('');
    lines.push('【増員内容】');
    lines.push(`　${eventData.content}`);
  }
  
  if (days.length > 1 && eventData.consecutivePreference) {
    lines.push('');
    lines.push('💡連日通し希望💡');
    let displayValue = eventData.consecutivePreference;
    if (displayValue === 'はい') displayValue = 'あり';
    if (displayValue === 'いいえ') displayValue = 'なし';
    lines.push(`　${displayValue}`);
  }
  
  return lines.join('\n');
}

function formatDate(date) {
  const month = date.getMonth() + 1;
  const day = date.getDate();
  const weekdays = ['日', '月', '火', '水', '木', '金', '土'];
  return `${month}/${day}(${weekdays[date.getDay()]})`;
}

function createMultiDayEvent(eventData) {
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!calendar) { Logger.log('カレンダーが見つかりません'); return; }
  
  const { color, emoji } = getSectionColorAndEmoji(eventData.days);
  const juniorMark = (eventData.juniorOk === 'はい') ? '🔰' : '';
  const title = `${emoji}${juniorMark}【増員】${eventData.hall}`;
  const description = buildDescription(eventData, eventData.days);
  
  const sortedDays = eventData.days.slice().sort((a, b) => a.date.getTime() - b.date.getTime());
  const startDate = sortedDays[0].date;
  const endDate = new Date(sortedDays[sortedDays.length - 1].date);
  endDate.setDate(endDate.getDate() + 1);
  
  const event = calendar.createAllDayEvent(title, startDate, endDate, { description });
  event.setColor(color);
  Logger.log(`複数日イベント作成: ${title}`);
}

function createSingleDayEvent(eventData, day, dayNumber) {
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!calendar) { Logger.log('カレンダーが見つかりません'); return; }
  
  const { color, emoji } = getSectionColorAndEmoji([day]);
  const juniorMark = (eventData.juniorOk === 'はい') ? '🔰' : '';
  const dayLabel = eventData.days.length > 1 ? ` (${dayNumber}/${eventData.days.length})` : '';
  const title = `${emoji}${juniorMark}【増員】${eventData.hall}${dayLabel}`;
  const description = buildDescription(eventData, [day]);
  
  const event = calendar.createAllDayEvent(title, day.date, { description });
  event.setColor(color);
  Logger.log(`単日イベント作成: ${title}`);
}

// ============================================
// Web App API: doGet
// ============================================
function doGet(e) {
  try {
    const params = e.parameter || {};
    
    if (params.page === 'liff') {
      return HtmlService.createTemplateFromFile('Liff').evaluate()
        .setTitle('LINE連携 - 増員 CONTROL HUB')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
    }
    
    if (params.action === 'verifyStaff') return createResponse(verifyStaff(params.name), e);
    if (params.action === 'saveLineUid') return createResponse(handleSaveLineUid(params), e);
    if (params.action === 'verifyLineUid') return createResponse(verifyLineUid(params.lineUid), e);
    if (params.action === 'registerLineStaff') return createResponse(registerLineStaff(params), e);
    if (params.action === 'linkLineToEmail') return createResponse(linkLineToEmail(params), e);
    if (params.action === 'submitApplication') return createResponse(handleSubmitApplication(params), e);
    if (params.action === 'getStaffProfile') return createResponse(getStaffProfile(params.email), e);
    if (params.action === 'getStaffInfo') return createResponse(getStaffInfoFromMainSheet(params.email), e);
    if (params.action === 'toggleRecruitment') return createResponse(toggleRecruitmentStatus(params.eventKey, params.hall, params.status, params.email, params.targetDay, params.targetSection), e);
    if (params.action === 'getRecruitmentStatuses') return createResponse(getRecruitmentStatuses(), e);
    if (params.action === 'getApplicantCounts') return createResponse(getApplicantCounts(), e);
    if (params.action === 'getHallList') return createResponse(getHallListForWeb(), e);
    if (params.action === 'getHallDetail') return createResponse(getHallDetailForWeb(params.hall), e);
    
    // カレンダーイベント取得
    const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    if (!calendar) return createResponse({ success: false, error: 'カレンダーが見つかりません' }, e);
    
    const startDate = params.start ? new Date(params.start) : new Date();
    const endDate = params.end ? new Date(params.end) : new Date(startDate.getTime() + 730 * 24 * 60 * 60 * 1000);
    const events = calendar.getEvents(startDate, endDate);
    
    // 申込者カウントを取得
    var applicantCountsResult = getApplicantCounts();
    var applicantCounts = (applicantCountsResult.success && applicantCountsResult.counts) ? applicantCountsResult.counts : {};
    
    const eventList = events.map(event => {
      const start = event.getAllDayStartDate();
      const end = event.getAllDayEndDate();
      const title = event.getTitle();
      const description = event.getDescription() || '';
      
      const hallMatch = title.match(/【増員】(.+?)(?:\s|\(|$)/);
      const hall = hallMatch ? hallMatch[1] : '';
      const parsedDescription = parseDescription(description);
      
      let color = '#4ade80';
      if (title.includes('🔴')) color = '#f87171';
      else if (title.includes('🟢')) color = '#4ade80';
      else if (title.includes('🔵')) color = '#60a5fa';
      else if (title.includes('🟡')) color = '#facc15';
      
      const isMultiDay = (end.getTime() - start.getTime()) > (24 * 60 * 60 * 1000);
      
      // このイベントのイベントキーを生成して申込数を取得
      // getApplicantCountsと同じキー形式: "ホール名_開始日"
      var eventKey = hall + '_' + formatISODate(start);
      var eventApplicantCounts = applicantCounts[eventKey] || { stage: 0, sound: 0, lighting: 0 };
      
      return {
        id: event.getId(),
        title: title,
        start: formatISODate(start),
        end: formatISODate(new Date(end.getTime() - 24 * 60 * 60 * 1000)),
        color: color,
        description: description,
        extendedProps: {
          hall: hall, isMultiDay: isMultiDay,
          sections: parsedDescription.sections, juniorOk: parsedDescription.juniorOk,
          consecutivePreference: parsedDescription.consecutivePreference,
          timeSlot: parsedDescription.timeSlot, content: parsedDescription.content,
          eventName: parsedDescription.eventName,
          applicantCounts: eventApplicantCounts
        }
      };
    });
    
    return createResponse({ success: true, events: eventList, totalCount: eventList.length }, e);
  } catch (error) {
    return createResponse({ success: false, error: error.toString() }, e);
  }
}

function parseDescription(description) {
  const result = { juniorOk: false, consecutivePreference: '', timeSlot: '', content: '', eventName: '', sections: { stage: 0, sound: 0, lighting: 0 } };
  if (!description) return result;
  
  const lines = description.split('\n');
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (line.includes('🔰若手可🔰')) result.juniorOk = true;
    else if (line.includes('連日通し希望')) {
      const parts = line.split('：');
      if (parts.length > 1) result.consecutivePreference = parts[1].trim();
    } else if (line.includes('【利用区分】') && i + 1 < lines.length) {
      const nextLine = lines[i + 1];
      if (nextLine && nextLine.startsWith('　')) result.timeSlot = nextLine.trim();
    } else if (line.includes('【増員内容】') && i + 1 < lines.length) {
      const nextLine = lines[i + 1];
      if (nextLine && nextLine.startsWith('　')) result.content = nextLine.trim();
    } else if (line.includes('【催事名】') && i + 1 < lines.length) {
      const nextLine = lines[i + 1];
      if (nextLine && nextLine.startsWith('　')) result.eventName = nextLine.trim();
    } else if (line.includes('舞台:') || line.includes('音響:') || line.includes('照明:')) {
      const stageMatch = line.match(/舞台:\s*(\d+)/);
      const soundMatch = line.match(/音響:\s*(\d+)/);
      const lightingMatch = line.match(/照明:\s*(\d+)/);
      if (stageMatch) result.sections.stage = parseInt(stageMatch[1], 10);
      if (soundMatch) result.sections.sound = parseInt(soundMatch[1], 10);
      if (lightingMatch) result.sections.lighting = parseInt(lightingMatch[1], 10);
    }
  }
  return result;
}

function formatISODate(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function createResponse(data, e) {
  const jsonStr = JSON.stringify(data);
  const params = e ? e.parameter : {};
  if (params && params.callback) {
    return ContentService.createTextOutput(params.callback + '(' + jsonStr + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(jsonStr).setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// スタッフ名簿照合
// ============================================
function normalizeName(name) {
  if (!name) return '';
  return name.trim().replace(/　/g, '').replace(/\s+/g, '').toLowerCase().normalize('NFKC');
}

function verifyStaff(identifier) {
  try {
    const normalizedInput = normalizeName(identifier);
    if (!normalizedInput) return { success: false, error: '入力されていません' };
    
    const ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(STAFF_CONFIG.SHEET_NAME) || ss.getSheetByName('フォームの回答 1') || ss.getSheetByName('必要事項の回答');
    if (!sheet) return { success: false, error: 'シートが見つかりません' };
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, verified: false };
    
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    for (const row of data) {
      if (normalizeName(row[5]) === normalizedInput) {
        return { success: true, verified: true, staffName: row[1], staffHall: row[3] || '', staffSection: row[4] || '' };
      }
    }
    return { success: true, verified: false };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// ============================================
// LINE Messaging API 設定
// ============================================
const LINE_CONFIG = {
  CHANNEL_ACCESS_TOKEN: 'mTRGTBaO0OUNO3Wv01lPa+QDMzX96HJSh93z9KwezgYo2PB8fqjJ8quUXRt2pTB5ts3lUsB3feOznP9cZR1K0tEgXE6gYE4LW733tsf1Dv3p2gyBYWO+Yw2JOAhASlEL8jTu6cB87QiUtABUsVi6+gdB04t89/1O/w1cDnyilFU=',
  OFFICIAL_ACCOUNT_ID: '825gnfcx',
  ADMIN_SHEET_NAME: 'LINE友だち'
};

// ============================================
// LINE友だちUID一覧を取得してスプレッドシートに保存
// ============================================
function fetchLineFollowers() {
  const url = 'https://api.line.me/v2/bot/followers/ids';
  const options = { method: 'get', headers: { 'Authorization': 'Bearer ' + LINE_CONFIG.CHANNEL_ACCESS_TOKEN }, muteHttpExceptions: true };
  let allUserIds = [];
  let nextToken = null;
  
  do {
    const requestUrl = nextToken ? url + '?start=' + nextToken : url;
    const response = UrlFetchApp.fetch(requestUrl, options);
    const result = JSON.parse(response.getContentText());
    if (result.message) { Logger.log('APIエラー: ' + result.message); return []; }
    if (result.userIds) allUserIds = allUserIds.concat(result.userIds);
    nextToken = result.next;
  } while (nextToken);
  
  saveFollowersToSheet(allUserIds);
  return allUserIds;
}

function saveFollowersToSheet(userIds) {
  const ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(LINE_CONFIG.ADMIN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LINE_CONFIG.ADMIN_SHEET_NAME);
    sheet.getRange(1, 1, 1, 4).setValues([['LINE UID', '名前', '種別', '備考']]);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }
  const existingUids = sheet.getDataRange().getValues().slice(1).map(row => row[0]);
  const newUids = userIds.filter(uid => !existingUids.includes(uid));
  if (newUids.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, newUids.length, 4).setValues(newUids.map(uid => [uid, '', '', '']));
  }
  return newUids.length;
}

// ============================================
// LINE通知関連
// ============================================
function notifyAdmins(staffName, staffLineUid, eventInfo) {
  const ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(LINE_CONFIG.ADMIN_SHEET_NAME);
  if (!sheet) return;
  
  const adminUids = sheet.getDataRange().getValues().slice(1).filter(row => row[2] === '管理者').map(row => row[0]);
  if (adminUids.length === 0) return;
  
  const message = createNotificationMessage(staffName, staffLineUid, eventInfo);
  sendMulticast(adminUids, message);
}

function createNotificationMessage(staffName, staffLineUid, eventInfo) {
  const chatUrl = staffLineUid ? 'https://chat.line.biz/Ue1234/chat/' + staffLineUid : 'https://manager.line.biz/';
  return {
    type: 'flex', altText: '📋 新規申し込み: ' + staffName + 'さん',
    contents: {
      type: 'bubble',
      header: { type: 'box', layout: 'vertical', contents: [{ type: 'text', text: '📋 新規申し込み', weight: 'bold', size: 'lg', color: '#1DB446' }] },
      body: { type: 'box', layout: 'vertical', contents: [
        { type: 'text', text: '👤 ' + staffName + 'さん', weight: 'bold', size: 'md' },
        { type: 'text', text: '📅 ' + eventInfo.date, margin: 'md' },
        { type: 'text', text: '🏢 ' + eventInfo.hall, margin: 'sm' },
        { type: 'text', text: '🎭 ' + eventInfo.section, margin: 'sm' }
      ]},
      footer: { type: 'box', layout: 'vertical', contents: [{ type: 'button', action: { type: 'uri', label: staffName + 'さんとチャット', uri: chatUrl }, style: 'primary', color: '#1DB446' }] }
    }
  };
}

function sendMulticast(userIds, message) {
  const options = {
    method: 'post',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_CONFIG.CHANNEL_ACCESS_TOKEN },
    payload: JSON.stringify({ to: userIds, messages: [message] }),
    muteHttpExceptions: true
  };
  return UrlFetchApp.fetch('https://api.line.me/v2/bot/message/multicast', options).getResponseCode() === 200;
}

// ============================================
// LINE UID保存ハンドラ（LIFFから呼び出し）
// ============================================
function handleSaveLineUid(params) {
  try {
    const lineUid = params.lineUid;
    const displayName = params.displayName || '';
    const staffName = params.staffName || '';
    if (!lineUid) return { success: false, error: 'LINE UIDが必要です' };
    
    const ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(LINE_CONFIG.ADMIN_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(LINE_CONFIG.ADMIN_SHEET_NAME);
      sheet.getRange(1, 1, 1, 9).setValues([['LINE UID', 'LINE表示名', 'スタッフ名', 'メール', '登録日時', 'セクション', 'ホール', '種別', '希望エリア']]);
      sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
    } else {
      // 既存シートにヘッダーが足りなければ追加
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (headers.length < 9 || !headers[5]) {
        if (headers.length < 6) sheet.getRange(1, 6).setValue('セクション');
        if (headers.length < 7) sheet.getRange(1, 7).setValue('ホール');
        if (headers.length < 8) sheet.getRange(1, 8).setValue('種別');
        if (headers.length < 9) sheet.getRange(1, 9).setValue('希望エリア');
        // 旧D列「種別」をH列に移動、旧E列「登録日時」をE列のまま
        // 旧D列を「メール」に変更
        sheet.getRange(1, 4).setValue('メール');
      }
    }
    
    // スタッフ管理シートから表示名で照合してメール・セクション・ホールを取得
    var staffInfo = lookupStaffByName(displayName);
    var email = staffInfo.email || '';
    var section = staffInfo.section || '';
    var hall = staffInfo.hall || '';
    var preferredArea = staffInfo.preferredArea || '';
    var matchedName = staffInfo.staffName || staffName || '';
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === lineUid) {
        // 既存UIDを更新
        sheet.getRange(i + 1, 2).setValue(displayName);
        if (matchedName) sheet.getRange(i + 1, 3).setValue(matchedName);
        if (email) sheet.getRange(i + 1, 4).setValue(email);
        if (section) sheet.getRange(i + 1, 6).setValue(section);
        if (hall) sheet.getRange(i + 1, 7).setValue(hall);
        if (preferredArea) sheet.getRange(i + 1, 9).setValue(preferredArea);
        return { success: true, message: '更新しました', isNew: false };
      }
    }
    
    sheet.appendRow([lineUid, displayName, matchedName, email, new Date(), section, hall, '', preferredArea]);
    return { success: true, message: '登録しました', isNew: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// スタッフ管理シートから名前で照合
function lookupStaffByName(displayName) {
  try {
    if (!displayName) return {};
    var normalizedInput = normalizeName(displayName);
    
    var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(STAFF_CONFIG.SHEET_NAME);
    if (!sheet) return {};
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};
    
    var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    for (var i = 0; i < data.length; i++) {
      // B列(名前)で照合
      if (normalizeName(data[i][1]) === normalizedInput) {
        return {
          staffName: data[i][1] || '',
          hall: data[i][3] || '',      // D列: ホール
          section: data[i][4] || '',   // E列: セクション
          email: data[i][5] || '',     // F列: メール
          preferredArea: data[i][6] || ''  // G列: 希望エリア
        };
      }
    }
    return {};
  } catch (error) {
    Logger.log('lookupStaffByName エラー: ' + error.toString());
    return {};
  }
}

function verifyLineUid(lineUid) {
  try {
    if (!lineUid) return { success: false, verified: false, error: 'LINE UIDが必要です' };
    const ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(LINE_CONFIG.ADMIN_SHEET_NAME);
    if (!sheet) return { success: true, verified: false };
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === lineUid) {
        const staffName = data[i][2];
        return staffName ? { success: true, verified: true, staffName } : { success: true, verified: false, message: 'スタッフ名が未登録です' };
      }
    }
    return { success: true, verified: false, message: 'LINE UID未登録' };
  } catch (error) {
    return { success: false, verified: false, error: error.toString() };
  }
}

function registerLineStaff(params) {
  try {
    const name = params.name;
    const lineUid = params.lineUid;
    const lineDisplayName = params.lineDisplayName || '';
    if (!name || !lineUid) return { success: false, verified: false, error: '名前とLINE UIDが必要です' };
    
    const staffResult = verifyStaff(name);
    if (!staffResult.success || !staffResult.verified) return { success: true, verified: false, message: 'スタッフ名簿に登録されていません' };
    
    const actualStaffName = staffResult.staffName || name;
    const ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(LINE_CONFIG.ADMIN_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(LINE_CONFIG.ADMIN_SHEET_NAME);
      sheet.getRange(1, 1, 1, 5).setValues([['LINE UID', 'LINE表示名', 'スタッフ名', '種別', '登録日時']]);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    }
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === lineUid) {
        sheet.getRange(i + 1, 2).setValue(lineDisplayName);
        sheet.getRange(i + 1, 3).setValue(actualStaffName);
        return { success: true, verified: true, staffName: actualStaffName };
      }
    }
    
    sheet.appendRow([lineUid, lineDisplayName, actualStaffName, '', new Date()]);
    return { success: true, verified: true, staffName: actualStaffName };
  } catch (error) {
    return { success: false, verified: false, error: error.toString() };
  }
}

// ============================================
// 申し込み処理（LINE通知付き）
// ============================================
function handleSubmitApplication(params) {
  try {
    // デバッグ: 受け取った全パラメータをログ出力
    Logger.log('handleSubmitApplication params: ' + JSON.stringify(params));
    const email = params.email;
    const staffName = params.staffName;
    const staffHall = params.staffHall || '';
    const staffSection = params.staffSection || '';
    const eventTitle = params.eventTitle || '';
    const hall = params.hall || '';
    const section = params.section || '';
    const date = params.date || '';
    const selectedDates = params.selectedDates || date;
    const sendLineNotification = params.sendLineNotification === 'true';
    
    const ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName('申し込み履歴');
    if (!sheet) {
      sheet = ss.insertSheet('申し込み履歴');
      sheet.getRange(1, 1, 1, 10).setValues([['申込日時', 'メールアドレス', 'スタッフ名', '所属ホール', '所属セクション', 'イベント', '会場ホール', 'セクション', '希望日程', 'ステータス']]);
      sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
    } else {
      // 既存シートにステータス列がなければ追加
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (headers.length < 10 || headers[9] !== 'ステータス') {
        sheet.getRange(1, 10).setValue('ステータス');
        sheet.getRange(1, 10).setFontWeight('bold');
      }
    }
    
    sheet.appendRow([new Date(), email, staffName, staffHall, staffSection, eventTitle, hall, section, selectedDates, '確認中']);
    updateStaffProfile(email, staffHall, staffSection);
    
    // FlexメッセージはLINEメッセージ受信時にdoPostで返信するため、ここでは送信しない
    
    // 管理者通知はLINEメッセージ受信時にdoPostでFlex送信するため、ここでは送信しない
    
    return { success: true, message: '申し込みを受け付けました' };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function updateStaffProfile(email, staffHall, staffSection) {
  try {
    if (!email) return;
    const ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName('スタッフプロフィール');
    if (!sheet) {
      sheet = ss.insertSheet('スタッフプロフィール');
      sheet.getRange(1, 1, 1, 4).setValues([['メールアドレス', '所属ホール', 'セクション', '更新日時']]);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    }
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        sheet.getRange(i + 1, 2).setValue(staffHall);
        sheet.getRange(i + 1, 3).setValue(staffSection);
        sheet.getRange(i + 1, 4).setValue(new Date());
        return;
      }
    }
    sheet.appendRow([email, staffHall, staffSection, new Date()]);
  } catch (error) {
    Logger.log('プロフィール更新エラー: ' + error.toString());
  }
}

function getStaffInfoFromMainSheet(email) {
  try {
    if (!email) return { success: false, error: 'メールアドレスが必要です' };
    var normalizedEmail = normalizeName(email);
    var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(STAFF_CONFIG.SHEET_NAME) || ss.getSheetByName('フォームの回答 1') || ss.getSheetByName('必要事項の回答');
    if (!sheet) return { success: false, error: 'シートが見つかりません' };
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, found: false };
    var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    
    for (var i = 0; i < data.length; i++) {
      if (normalizeName(data[i][5]) === normalizedEmail) {
        return {
          success: true,
          found: true,
          staffName: data[i][1] || '',
          staffHall: data[i][3] || '',
          staffSection: data[i][4] || '',
          preferredArea: data[i][6] || '',
          isAdmin: data[i][8] === true || data[i][8] === 'TRUE',
          isSystemAdmin: data[i][9] === true || data[i][9] === 'TRUE'
        };
      }
    }
    return { success: true, found: false };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// ============================================
// 募集完了ステータス管理
// ============================================
function getOrCreateManagementSheet() {
  var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('イベント管理');
  if (!sheet) {
    sheet = ss.insertSheet('イベント管理');
    sheet.getRange('A1:D1').setValues([['イベントキー', 'ホール名', '完了', '更新日時']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function toggleRecruitmentStatus(eventKey, hall, status, email, targetDay, targetSection) {
  try {
    if (!eventKey || !hall || !email) return { success: false, error: 'パラメータ不足' };
    var staffInfo = getStaffInfoFromMainSheet(email);
    if (!staffInfo.found || !staffInfo.isAdmin) return { success: false, error: '管理者権限がありません' };
    if (staffInfo.staffHall !== hall) return { success: false, error: '自分のホールのみ操作可能' };
    
    // サフィックス付きキーを生成
    var actualKey = eventKey;
    if (targetDay) actualKey += '__day:' + targetDay;
    if (targetSection) actualKey += '__sec:' + targetSection;
    
    // セクション一括操作の判定（targetSectionあり、targetDayなし）
    var isSectionBulk = targetSection && !targetDay;
    
    var sheet = getOrCreateManagementSheet();
    var lastRow = sheet.getLastRow();
    var isComplete = (status === 'true' || status === true);
    var updated = false;
    var childUpdated = false;
    
    if (lastRow >= 2) {
      var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
      for (var i = 0; i < data.length; i++) {
        var rowKey = data[i][0];
        var isMatch = false;
        
        // 通常のマッチング（完全一致 or 前方一致）
        if (rowKey === actualKey || rowKey.indexOf(actualKey + '__') === 0) {
          isMatch = true;
        }
        
        // セクション一括操作: __day:X__sec:Y のエントリもマッチ
        // 例: actualKey = eventKey__sec:stage の時
        //     rowKey = eventKey__day:2026-06-10__sec:stage もマッチさせる
        if (isSectionBulk && !isMatch) {
          if (rowKey.indexOf(eventKey + '__day:') === 0 && rowKey.indexOf('__sec:' + targetSection) > 0) {
            isMatch = true;
          }
        }
        
        if (isMatch) {
          // isComplete == false（再開）の場合は派生キーもまとめて再開させる
          // isComplete == true（終了）の場合は指定されたキーのみ終了させる
          if (!isComplete || rowKey === actualKey) {
            sheet.getRange(i + 2, 3).setValue(isComplete);
            sheet.getRange(i + 2, 4).setValue(new Date());
            if (rowKey === actualKey) {
              updated = true;
            } else {
              childUpdated = true;
            }
          }
        }
      }
    }
    // 自分自身のキーが見つからなかった場合のみ新規行を追加
    // ただし、再開操作で子キーのみ更新した場合は不要な行を追加しない
    if (!updated && !(childUpdated && !isComplete)) {
      sheet.appendRow([actualKey, hall, isComplete, new Date()]);
    }
    return { success: true, eventKey: actualKey, closed: isComplete, targetDay: targetDay || '', targetSection: targetSection || '' };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function getRecruitmentStatuses() {
  try {
    var sheet = getOrCreateManagementSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, statuses: {} };
    var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    var statuses = {};
    for (var i = 0; i < data.length; i++) {
      // サフィックス付きキー（__day:, __sec:）も含めて全て返却
      if (data[i][2] === true || data[i][2] === 'TRUE') statuses[data[i][0]] = true;
    }
    return { success: true, statuses: statuses };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function getStaffProfile(email) {
  try {
    if (!email) return { success: false, error: 'メールアドレスが必要です' };
    const ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('スタッフプロフィール');
    if (!sheet) return { success: true, found: false };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) return { success: true, found: true, staffHall: data[i][1], staffSection: data[i][2] };
    }
    return { success: true, found: false };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// ============================================
// メールアドレスとLINE UIDを紐付け（LIFF連携用）
// ============================================
function linkLineToEmail(params) {
  try {
    const email = params.email;
    const lineUid = params.lineUid;
    const lineDisplayName = params.lineDisplayName || '';
    if (!email || !lineUid) return { success: false, linked: false, error: 'メールアドレスとLINE UIDが必要です' };
    
    const staffResult = verifyStaff(email);
    if (!staffResult.success || !staffResult.verified) return { success: true, linked: false, message: 'このメールアドレスはスタッフ名簿に登録されていません' };
    
    const staffName = staffResult.staffName || email;
    const ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(LINE_CONFIG.ADMIN_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(LINE_CONFIG.ADMIN_SHEET_NAME);
      sheet.getRange(1, 1, 1, 6).setValues([['LINE UID', 'LINE表示名', 'スタッフ名', 'メールアドレス', '種別', '登録日時']]);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    }
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === lineUid) {
        sheet.getRange(i + 1, 2).setValue(lineDisplayName);
        sheet.getRange(i + 1, 3).setValue(staffName);
        sheet.getRange(i + 1, 4).setValue(email);
        return { success: true, linked: true, staffName };
      }
    }
    
    sheet.appendRow([lineUid, lineDisplayName, staffName, email, '', new Date()]);
    return { success: true, linked: true, staffName };
  } catch (error) {
    return { success: false, linked: false, error: error.toString() };
  }
}

function getLineUidByEmail(email) {
  try {
    const ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(LINE_CONFIG.ADMIN_SHEET_NAME);
    if (!sheet) return null;
    var lastCol = Math.max(sheet.getLastColumn(), 8);
    const data = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
    for (let i = 1; i < data.length; i++) {
      // D列(メール)で検索
      if (email && data[i][3] && data[i][3].toString().trim() === email.trim()) return data[i][0];
    }
    return null;
  } catch (error) {
    return null;
  }
}

// スタッフ名またはLINE表示名でUIDを検索
function getLineUidByName(staffName) {
  try {
    if (!staffName) return null;
    var normalizedInput = normalizeName(staffName);
    const ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(LINE_CONFIG.ADMIN_SHEET_NAME);
    if (!sheet) return null;
    var lastCol = Math.max(sheet.getLastColumn(), 8);
    const data = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
    for (let i = 1; i < data.length; i++) {
      // C列(スタッフ名)またはB列(LINE表示名)で検索
      if (normalizeName(data[i][2]) === normalizedInput || normalizeName(data[i][1]) === normalizedInput) {
        return data[i][0];
      }
    }
    return null;
  } catch (error) {
    return null;
  }
}

// 申し込み完了Flexメッセージ送信（スタッフ名で直接UID検索）
function sendApplicationCompleteFlex(staffName, eventInfo) {
  try {
    if (!eventInfo) eventInfo = {};
    
    // LINE友だちシートから直接UIDを検索（B列=表示名, C列=スタッフ名）
    var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(LINE_CONFIG.ADMIN_SHEET_NAME);
    if (!sheet) {
      Logger.log('LINE友だちシートがありません');
      return false;
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('LINE友だちシートにデータがありません');
      return false;
    }
    
    var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); // A-C列のみ
    var lineUid = null;
    var searchName = staffName ? normalizeName(staffName) : '';
    
    // 全行を検索（B列=LINE表示名, C列=スタッフ名で照合）
    for (var i = 0; i < data.length; i++) {
      var uid = (data[i][0] || '').toString().trim();
      var displayName = normalizeName(data[i][1] || '');
      var registeredName = normalizeName(data[i][2] || '');
      
      if (!uid) continue;
      
      if (searchName && (displayName === searchName || registeredName === searchName)) {
        lineUid = uid;
        Logger.log('UID見つかりました: ' + uid + ' (名前: ' + staffName + ')');
        break;
      }
    }
    
    // 名前で見つからなければ、1件だけの場合はそのUIDを使う
    if (!lineUid && data.length === 1 && data[0][0]) {
      lineUid = data[0][0].toString().trim();
      Logger.log('LINE友だちが1件のみ、そのUIDを使用: ' + lineUid);
    }
    
    if (!lineUid) {
      Logger.log('UIDが見つかりません。検索名: ' + staffName);
      // LINE友だちシートの全データをログ出力
      for (var j = 0; j < data.length; j++) {
        Logger.log('  行' + (j+2) + ': UID=' + data[j][0] + ', 表示名=' + data[j][1] + ', スタッフ名=' + data[j][2]);
      }
      return false;
    }
    
    // 日付フォーマット
    var dateText = (eventInfo.date || '').split(',').map(function(d) {
      d = d.trim();
      var match = d.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
      return match ? parseInt(match[2]) + '月' + parseInt(match[3]) + '日' : d;
    }).join('\n');
    
    // Flexメッセージ
    var message = buildFlexCard(dateText, eventInfo);
    
    var response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_CONFIG.CHANNEL_ACCESS_TOKEN },
      payload: JSON.stringify({ to: lineUid, messages: [message] }),
      muteHttpExceptions: true
    });
    
    var code = response.getResponseCode();
    Logger.log('LINE API応答: ' + code + ' ' + response.getContentText());
    return code === 200;
  } catch (error) {
    Logger.log('sendApplicationCompleteFlex エラー: ' + error.toString());
    return false;
  }
}

// Flexカードを生成（申し込み完了通知）
function buildFlexCard(dateText, eventInfo) {
  return {
    type: 'flex', altText: '✅ 申し込み完了',
    contents: {
      type: 'bubble',
      styles: {
        header: { backgroundColor: '#1a6b4a' },
        body: { backgroundColor: '#0a1628' },
        footer: { backgroundColor: '#0a1628' }
      },
      header: {
        type: 'box', layout: 'vertical',
        contents: [
          {
            type: 'box', layout: 'vertical',
            contents: [{ type: 'text', text: '✅ 申し込み完了', weight: 'bold', size: 'xl', color: '#ffffff', align: 'center' }],
            backgroundColor: '#2d9a6a', cornerRadius: 'md', paddingAll: '16px'
          }
        ],
        backgroundColor: '#1a6b4a', paddingAll: '16px'
      },
      body: {
        type: 'box', layout: 'vertical',
        contents: [
          { type: 'text', text: '申し込み内容', weight: 'bold', size: 'md', color: '#60a5fa', align: 'center', margin: 'sm' },
          {
            type: 'box', layout: 'vertical',
            contents: [
              { type: 'box', layout: 'horizontal', contents: [
                { type: 'text', text: '日付', size: 'sm', color: '#7a8a9a', flex: 1, align: 'start' },
                { type: 'text', text: dateText || '-', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
              ], margin: 'lg' },
              { type: 'box', layout: 'horizontal', contents: [
                { type: 'text', text: '催事名', size: 'sm', color: '#7a8a9a', flex: 1, align: 'start' },
                { type: 'text', text: eventInfo.eventTitle || '-', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
              ], margin: 'md' },
              { type: 'box', layout: 'horizontal', contents: [
                { type: 'text', text: 'ホール', size: 'sm', color: '#7a8a9a', flex: 1, align: 'start' },
                { type: 'text', text: eventInfo.hall || '-', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
              ], margin: 'md' },
              { type: 'box', layout: 'horizontal', contents: [
                { type: 'text', text: '申込者', size: 'sm', color: '#7a8a9a', flex: 1, align: 'start' },
                { type: 'text', text: eventInfo.staffName || '-', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
              ], margin: 'md' },
              { type: 'box', layout: 'horizontal', contents: [
                { type: 'text', text: 'セクション', size: 'sm', color: '#7a8a9a', flex: 1, align: 'start' },
                { type: 'text', text: eventInfo.section || '-', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start' }
              ], margin: 'md' }
            ],
            backgroundColor: '#142238', cornerRadius: 'md', paddingAll: '14px', margin: 'md'
          }
        ],
        backgroundColor: '#0a1628', paddingAll: '18px'
      },
      footer: {
        type: 'box', layout: 'vertical',
        contents: [
          { type: 'separator', color: '#1e3a5f' },
          { type: 'text', text: 'お申し込みありがとうございます☺️', size: 'xs', color: '#7a8a9a', wrap: true, align: 'center', margin: 'lg' },
          { type: 'text', text: '現時点では増員は確定しておりません⚠️', size: 'xs', color: '#7a8a9a', wrap: true, align: 'center', margin: 'sm' },
          { type: 'text', text: '担当者からのご連絡をもって', size: 'xs', color: '#7a8a9a', wrap: true, align: 'center', margin: 'sm' },
          { type: 'text', text: '正式決定となります！！', size: 'xs', color: '#7a8a9a', wrap: true, align: 'center', margin: 'sm' },
          { type: 'text', text: 'ご連絡まで、今しばらくお待ちください📱', size: 'xs', color: '#7a8a9a', wrap: true, align: 'center', margin: 'sm' }
        ],
        backgroundColor: '#0a1628', paddingAll: '15px'
      }
    }
  };
}

// ============================================
// LINE Webhook受信（doPost）
// ============================================
function doPost(e) {
  try {
    if (!e || !e.postData) {
      Logger.log('doPost: リクエストデータなし（手動実行？）');
      return ContentService.createTextOutput(JSON.stringify({status: 'no data'})).setMimeType(ContentService.MimeType.JSON);
    }
    const data = JSON.parse(e.postData.contents);
    const events = data.events || [];
    
    for (const event of events) {
      // ============================================
      // postbackイベント処理（ホール案内Bot）
      // ============================================
      if (event.type === 'postback') {
        handleHallBotPostback(event);
        continue;
      }
      
      if (event.type !== 'message' || event.message.type !== 'text') continue;
      const messageText = event.message.text.trim();
      
      // ============================================
      // ホール案内Bot（テキスト「ホール案内」で起動）
      // ============================================
      if (messageText === 'ホール案内') {
        handleHallBotText(event);
        continue;
      }
      
      // ============================================
      // 既存の申し込み処理
      // ============================================
      if (messageText.includes('申し込みます')) {
        // メッセージから申し込み内容をパース
        var appInfo = parseApplicationMessage(messageText);
        var dateText = appInfo.date || '未定';
        var senderUid = event.source ? event.source.userId : '';
        
        // カウント+1（重複チェック付き）
        var confirmParams = {
          staffName: appInfo.name || '-',
          hall: appInfo.venue || appInfo.hall || '-',
          section: appInfo.section || '-',
          eventTitle: appInfo.eventTitle || '-',
          date: dateText
        };
        var confirmResult = handleConfirmApplication(confirmParams, senderUid);
        
        if (confirmResult.duplicate) {
          // 重複: 完了Flexは送るが、カウントは加算されない旨を追記
          Logger.log('重複申し込み検出（カウント加算なし）: ' + appInfo.name);
        } else {
          Logger.log('申し込みカウント+1: ' + appInfo.name + ' / ' + appInfo.section);
        }
        
        // 応募者へFlexメッセージで返信（重複でも完了通知は返す）
        var flexMessage = buildFlexCard(dateText, {
          hall: appInfo.venue || appInfo.hall || '-',
          staffName: appInfo.name || '-',
          section: appInfo.section || '-',
          eventTitle: appInfo.eventTitle || '-'
        });
        replyFlexMessage(event.replyToken, flexMessage);
        Logger.log('Flex返信完了: ' + appInfo.name);
        
        // 管理者へFlex通知をPush送信（重複時はスキップ）
        if (!confirmResult.duplicate) {
          try {
            var lineDisplayName = '';
            if (senderUid) {
              try {
                var profileRes = UrlFetchApp.fetch('https://api.line.me/v2/bot/profile/' + senderUid, {
                  headers: { 'Authorization': 'Bearer ' + LINE_CONFIG.CHANNEL_ACCESS_TOKEN },
                  muteHttpExceptions: true
                });
                if (profileRes.getResponseCode() === 200) {
                  lineDisplayName = JSON.parse(profileRes.getContentText()).displayName || '';
                }
              } catch (profileErr) {}
            }
            
            notifyAdminsByHall({
              date: dateText,
              hall: appInfo.venue || '-',
              section: appInfo.section || '-',
              staffName: appInfo.name || '-',
              staffHall: appInfo.hall || '-',
              lineDisplayName: lineDisplayName,
              eventTitle: appInfo.eventTitle || appInfo.venue || '-',
              senderUid: senderUid
            });
            Logger.log('管理者Flex通知完了');
          } catch (adminErr) {
            Logger.log('管理者通知エラー: ' + adminErr.toString());
          }
        }
      }
    }
  } catch (error) {
    Logger.log('Webhook処理エラー: ' + error.toString());
  }
  return ContentService.createTextOutput(JSON.stringify({status: 'ok'})).setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// 申込者カウント集計（確定済みのみ）
// ============================================
function getApplicantCounts() {
  try {
    var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('申し込み履歴');
    if (!sheet) return { success: true, counts: {} };
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, counts: {} };
    
    var lastCol = Math.max(sheet.getLastColumn(), 10);
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    // counts = { "イベントキー": { stage: N, sound: N, lighting: N } }
    var counts = {};
    
    for (var i = 0; i < data.length; i++) {
      var status = (data[i][9] || '').toString().trim();
      if (status !== '確定') continue; // 確定済みのみカウント
      
      var eventTitle = (data[i][5] || '').toString().trim(); // F列: イベント
      var hall = (data[i][6] || '').toString().trim();        // G列: 会場ホール
      var section = (data[i][7] || '').toString().trim();     // H列: セクション
      var dates = (data[i][8] || '').toString().trim();       // I列: 希望日程
      
      // イベントキーの生成: カレンダーイベントのタイトルと開始日で一致させる
      // カレンダー側のキーは "タイトル_開始日(YYYY-MM-DD)" 形式
      // 申し込み履歴では催事名とホールからキーを推測
      // 日付ごとにカウント（複数日選択の場合は各日にカウント）
      var dateList = dates.split(',').map(function(d) { return d.trim(); });
      
      // セクションをマッピング
      var sectionKey = '';
      if (section.indexOf('舞台') >= 0) sectionKey = 'stage';
      else if (section.indexOf('音響') >= 0) sectionKey = 'sound';
      else if (section.indexOf('照明') >= 0) sectionKey = 'lighting';
      
      if (!sectionKey) continue;
      
      // 全イベントキーに対してカウント
      // hall情報からカレンダーのタイトルパターンを推測
      for (var j = 0; j < dateList.length; j++) {
        var dateStr = dateList[j];
        // YYYY-MM-DD形式に統一する
        if (dateStr.match(/^\d{4}-\d{2}-\d{2}$/)) {
          // そのまま使用
        } else {
          // 「Wed May 20 2026...」のような形式の救済処置
          var parsedDate = new Date(dateStr);
          if (!isNaN(parsedDate.getTime())) {
            var y = parsedDate.getFullYear();
            var m = ('0' + (parsedDate.getMonth() + 1)).slice(-2);
            var d = ('0' + parsedDate.getDate()).slice(-2);
            dateStr = y + '-' + m + '-' + d;
          } else {
            continue; // 不正な日付はスキップ
          }
        }
        
        // カレンダーのイベントキーと申し込みのマッチング
        // イベントキーパターン: "*【増員】{hall}*_{startDate}"
        // 部分一致で探すため、hall + dateStr をキーに
        var matchKey = hall + '_' + dateStr;
        
        if (!counts[matchKey]) {
          counts[matchKey] = { stage: 0, sound: 0, lighting: 0 };
        }
        counts[matchKey][sectionKey]++;
      }
    }
    
    return { success: true, counts: counts };
  } catch (error) {
    Logger.log('getApplicantCounts エラー: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// ============================================
// 申し込み確定処理 + 重複チェック
// LINE「申し込みます」テキスト受信時に呼ばれる
// ============================================
function handleConfirmApplication(params, senderUid) {
  try {
    var staffName = params.staffName || '';
    var hall = params.hall || '';
    var section = params.section || '';
    var eventTitle = params.eventTitle || '';
    var dateText = params.date || '';
    
    var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('申し込み履歴');
    if (!sheet) {
      sheet = ss.insertSheet('申し込み履歴');
      sheet.getRange(1, 1, 1, 10).setValues([['申込日時', 'メールアドレス', 'スタッフ名', '所属ホール', '所属セクション', 'イベント', '会場ホール', 'セクション', '希望日程', 'ステータス']]);
      sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
    }
    
    var lastRow = sheet.getLastRow();
    
    // 1. まず既存の「確定」行がないか確認し、重複を防ぐ
    if (lastRow >= 2) {
      var lastCol = Math.max(sheet.getLastColumn(), 10);
      var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      
      for (var i = data.length - 1; i >= 0; i--) {
        var existingName = (data[i][2] || '').toString().trim();
        var existingEvent = (data[i][5] || '').toString().trim();
        var existingHall = (data[i][6] || '').toString().trim();
        var existingSection = (data[i][7] || '').toString().trim();
        var existingStatus = (data[i][9] || '').toString().trim();
        
        // 催事名または日付で重複判定
        var existingDates = (data[i][8] || '').toString().trim();
        var dateMatches = false;
        if (eventTitle) {
            dateMatches = (existingEvent === eventTitle);
        } else {
            dateMatches = (existingDates.indexOf(dateText) !== -1 || dateText.indexOf(existingDates) !== -1);
        }

        if (existingStatus === '確定' &&
            normalizeName(existingName) === normalizeName(staffName) &&
            existingHall === hall &&
            existingSection === section &&
            dateMatches) {
          Logger.log('重複確定検出: ' + staffName + ' / ' + eventTitle + ' / ' + hall);
          return { success: true, duplicate: true };
        }
      }
    }
    
    // 3. どちらにも該当しない場合は新規で確定行を追加（※例外的な手動送信など）
    sheet.appendRow([new Date(), '', staffName, '', section, eventTitle, hall, section, dateText, '確定']);
    Logger.log('確定行を新規追加: ' + staffName + ' / ' + eventTitle);
    
    return { success: true, duplicate: false };
  } catch (error) {
    Logger.log('handleConfirmApplication エラー: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// 名前の正規化（スペースの全角半角・前後空白の差を吸収）
function normalizeName(name) {
  if (!name) return '';
  return name.toString().trim().replace(/\s+/g, ' ').replace(/　/g, ' ');
}

// 申し込みLINEメッセージをパース
function parseApplicationMessage(text) {
  var info = { hall: '', name: '', section: '', date: '', venue: '', eventTitle: '' };
  
  var hallMatch = text.match(/【所属ホール】\s*\n?\s*(.+)/);
  if (hallMatch) info.hall = hallMatch[1].trim();
  
  var nameMatch = text.match(/【お名前】\s*\n?\s*(.+)/);
  if (nameMatch) info.name = nameMatch[1].trim();
  
  var sectionMatch = text.match(/【セクション】\s*\n?\s*(.+)/);
  if (sectionMatch) info.section = sectionMatch[1].trim();
  
  var eventTitleMatch = text.match(/【催事名】\s*\n?\s*(.+)/);
  if (eventTitleMatch) info.eventTitle = eventTitleMatch[1].trim();
  
  // 複数日対応: 【増員日】の後の「・○月○日」や「・YYYY-MM-DD」を全行取得
  var dateMatch = text.match(/【増員日】\s*\n([\s\S]*?)(?=\n\s*【|$)/);
  if (dateMatch) {
    var dateLines = dateMatch[1].split('\n')
      .map(function(line) { return line.replace(/[・\s　]/g, '').trim(); })
      .filter(function(line) { return line.length > 0; });
    info.date = dateLines.join(',');
  }
  
  var venueMatch = text.match(/【募集事業所】\s*\n?\s*(.+)/);
  if (venueMatch) info.venue = venueMatch[1].trim();
  
  return info;
}

// 申し込み完了Flexメッセージを生成
function createApplicationCompleteFlexMessage(appInfo) {
  return {
    type: 'flex',
    altText: '✅ 申し込み完了',
    contents: {
      type: 'bubble',
      styles: {
        header: { backgroundColor: '#1a6b4a' },
        body: { backgroundColor: '#0a1628' },
        footer: { backgroundColor: '#0a1628' }
      },
      header: {
        type: 'box', layout: 'vertical',
        contents: [
          {
            type: 'box', layout: 'vertical',
            contents: [{ type: 'text', text: '✅ 申し込み完了', weight: 'bold', size: 'xl', color: '#ffffff', align: 'center' }],
            backgroundColor: '#2d9a6a', cornerRadius: 'md', paddingAll: '16px'
          }
        ],
        backgroundColor: '#1a6b4a', paddingAll: '16px'
      },
      body: {
        type: 'box', layout: 'vertical',
        contents: [
          { type: 'text', text: '申し込み内容', weight: 'bold', size: 'md', color: '#60a5fa', align: 'center', margin: 'sm' },
          {
            type: 'box', layout: 'vertical',
            contents: [
              { type: 'box', layout: 'horizontal', contents: [
                { type: 'text', text: '日付', size: 'sm', color: '#7a8a9a', flex: 1, align: 'start' },
                { type: 'text', text: appInfo.date || '未定', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
              ], margin: 'lg' },
              { type: 'box', layout: 'horizontal', contents: [
                { type: 'text', text: '催事名', size: 'sm', color: '#7a8a9a', flex: 1, align: 'start' },
                { type: 'text', text: appInfo.eventTitle || '-', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
              ], margin: 'md' },
              { type: 'box', layout: 'horizontal', contents: [
                { type: 'text', text: 'ホール', size: 'sm', color: '#7a8a9a', flex: 1, align: 'start' },
                { type: 'text', text: appInfo.venue || '未定', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
              ], margin: 'md' },
              { type: 'box', layout: 'horizontal', contents: [
                { type: 'text', text: '申込者', size: 'sm', color: '#7a8a9a', flex: 1, align: 'start' },
                { type: 'text', text: appInfo.name || '不明', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
              ], margin: 'md' },
              { type: 'box', layout: 'horizontal', contents: [
                { type: 'text', text: 'セクション', size: 'sm', color: '#7a8a9a', flex: 1, align: 'start' },
                { type: 'text', text: appInfo.section || '不明', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start' }
              ], margin: 'md' }
            ],
            backgroundColor: '#142238', cornerRadius: 'md', paddingAll: '14px', margin: 'md'
          }
        ],
        backgroundColor: '#0a1628', paddingAll: '18px'
      },
      footer: {
        type: 'box', layout: 'vertical',
        contents: [
          { type: 'separator', color: '#1e3a5f' },
          { type: 'text', text: 'お申し込みありがとうございます😊', size: 'xs', color: '#7a8a9a', wrap: true, align: 'center', margin: 'lg' },
          { type: 'text', text: '現時点では増員は確定しておりません。', size: 'xs', color: '#7a8a9a', wrap: true, align: 'center', margin: 'sm' },
          { type: 'text', text: '担当者からのご連絡をもって正式決定となります！', size: 'xs', color: '#7a8a9a', wrap: true, align: 'center', margin: 'sm' },
          { type: 'text', text: 'ご連絡まで、今しばらくお待ちください📱', size: 'xs', color: '#7a8a9a', wrap: true, align: 'center', margin: 'sm' }
        ],
        backgroundColor: '#0a1628', paddingAll: '15px'
      }
    }
  };
}

function replyFlexMessage(replyToken, flexMessage) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_CONFIG.CHANNEL_ACCESS_TOKEN },
    payload: JSON.stringify({ replyToken: replyToken, messages: [flexMessage] }),
    muteHttpExceptions: true
  });
}

// ============================================
// 担当ホールの管理者にLINE通知を送信
// ============================================
function notifyAdminsByHall(eventInfo) {
  try {
    const targetHall = eventInfo.hall;
    if (!targetHall) return;
    
    const ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(STAFF_CONFIG.SHEET_NAME);
    if (!sheet) return;
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    
    const adminEmails = [], adminNames = [];
    for (let i = 0; i < data.length; i++) {
      if ((data[i][8] === true || data[i][8] === 'TRUE') && data[i][3] === targetHall && data[i][5]) {
        adminEmails.push(data[i][5]);
        adminNames.push(data[i][1] || data[i][5]);
      }
    }
    if (adminEmails.length === 0) return;
    
    for (let i = 0; i < adminEmails.length; i++) {
      const uid = getLineUidByEmail(adminEmails[i]);
      if (!uid) continue;
      
      const message = createAdminNotificationMessage(adminNames[i], eventInfo);
      UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
        method: 'post',
        headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_CONFIG.CHANNEL_ACCESS_TOKEN },
        payload: JSON.stringify({ to: uid, messages: [message] }),
        muteHttpExceptions: true
      });
    }
  } catch (error) {
    Logger.log('管理者通知エラー: ' + error.toString());
  }
}

function createAdminNotificationMessage(adminName, eventInfo) {
  // チャットボタン用URL（申込者のUIDがある場合のみリンク生成）
  var chatUrl = '';
  if (eventInfo.senderUid) {
    chatUrl = 'https://chat.line.biz/' + LINE_CONFIG.OFFICIAL_ACCOUNT_ID + '/chat/' + eventInfo.senderUid;
  }
  
  // ボディコンテンツ
  var bodyContents = [
    { type: 'text', text: '増員募集中の催事に申し込みがありました！', weight: 'bold', size: 'xs', align: 'center', color: '#ffffff', margin: 'sm' },
    { type: 'text', text: '催事情報', weight: 'bold', size: 'md', color: '#f0a050', align: 'center', margin: 'lg' },
    { type: 'box', layout: 'vertical', contents: [
      { type: 'box', layout: 'horizontal', contents: [
        { type: 'text', text: '募集日', size: 'sm', color: '#8a7a6a', flex: 1, align: 'start' },
        { type: 'text', text: eventInfo.date || '-', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
      ], margin: 'lg' },
      { type: 'box', layout: 'horizontal', contents: [
        { type: 'text', text: '催事名', size: 'sm', color: '#8a7a6a', flex: 1, align: 'start' },
        { type: 'text', text: eventInfo.eventTitle || '-', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
      ], margin: 'md' },
      { type: 'box', layout: 'horizontal', contents: [
        { type: 'text', text: 'セクション', size: 'sm', color: '#8a7a6a', flex: 1, align: 'start' },
        { type: 'text', text: eventInfo.section || '-', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start' }
      ], margin: 'md' }
    ], backgroundColor: '#2a1e10', cornerRadius: 'md', paddingAll: '14px', margin: 'sm' },
    { type: 'text', text: '申込者情報', weight: 'bold', size: 'md', color: '#f0a050', align: 'center', margin: 'lg' },
    { type: 'box', layout: 'vertical', contents: [
      { type: 'box', layout: 'horizontal', contents: [
        { type: 'text', text: '応募', size: 'sm', color: '#8a7a6a', flex: 1, align: 'start' },
        { type: 'text', text: eventInfo.date || '-', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
      ], margin: 'lg' },
      { type: 'box', layout: 'horizontal', contents: [
        { type: 'text', text: 'お名前', size: 'sm', color: '#8a7a6a', flex: 1, align: 'start' },
        { type: 'text', text: eventInfo.staffName || '-', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
      ], margin: 'md' },
      { type: 'box', layout: 'horizontal', contents: [
        { type: 'text', text: 'LINE名', size: 'sm', color: '#8a7a6a', flex: 1, align: 'start' },
        { type: 'text', text: eventInfo.lineDisplayName || '-', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
      ], margin: 'md' },
      { type: 'box', layout: 'horizontal', contents: [
        { type: 'text', text: '所属ホール', size: 'sm', color: '#8a7a6a', flex: 1, align: 'start' },
        { type: 'text', text: eventInfo.staffHall || '-', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
      ], margin: 'md' },
      { type: 'box', layout: 'horizontal', contents: [
        { type: 'text', text: 'セクション', size: 'sm', color: '#8a7a6a', flex: 1, align: 'start' },
        { type: 'text', text: eventInfo.section || '-', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start' }
      ], margin: 'md' }
    ], backgroundColor: '#2a1e10', cornerRadius: 'md', paddingAll: '14px', margin: 'sm' }
  ];
  
  // チャットボタン（UIDがある場合のみ追加。文字切れ防止のためBoxでボタン風に作成）
  if (chatUrl) {
    bodyContents.push({
      type: 'box',
      layout: 'vertical',
      backgroundColor: '#e67e22',
      cornerRadius: 'md',
      paddingAll: '12px',
      margin: 'xl',
      action: { type: 'uri', label: 'チャットを開く', uri: chatUrl },
      contents: [
        { 
          type: 'text', 
          text: '💬 申込者との公式LINEチャットへ', 
          color: '#ffffff', 
          weight: 'bold', 
          align: 'center', 
          wrap: true,
          size: 'sm'
        }
      ]
    });
  }
  
  return {
    type: 'flex',
    altText: '📢 管理者専用通知 - 増員応募あり',
    contents: {
      type: 'bubble',
      styles: {
        header: { backgroundColor: '#7a4510' },
        body: { backgroundColor: '#1a1008' },
        footer: { backgroundColor: '#1a1008' }
      },
      header: {
        type: 'box', layout: 'vertical',
        contents: [
          {
            type: 'box', layout: 'vertical',
            contents: [{ type: 'text', text: '📢 管理者専用通知 📢', weight: 'bold', size: 'lg', color: '#ffffff', align: 'center' }],
            backgroundColor: '#e67e22', cornerRadius: 'md', paddingAll: '16px'
          }
        ],
        backgroundColor: '#7a4510', paddingAll: '16px'
      },
      body: {
        type: 'box', layout: 'vertical',
        contents: bodyContents,
        backgroundColor: '#1a1008', paddingAll: '18px'
      },
      footer: {
        type: 'box', layout: 'vertical',
        contents: [
          { type: 'separator', color: '#3a2a10' },
          { type: 'text', text: 'ボタンをタップし、公式LINEのチャットへ', size: 'xs', color: '#8a7a6a', wrap: true, align: 'center', margin: 'lg' },
          { type: 'text', text: '移動します。そちらから申込者の方へ、', size: 'xs', color: '#8a7a6a', wrap: true, align: 'center', margin: 'sm' },
          { type: 'text', text: 'ご連絡をお願いします！', size: 'xs', color: '#8a7a6a', wrap: true, align: 'center', margin: 'sm' }
        ],
        backgroundColor: '#1a1008', paddingAll: '15px'
      }
    }
  };
}

// ============================================
// エリア判定機能（住所ベース自動判定）
// ============================================

/**
 * 住所から都道府県を抽出してエリアに変換
 * 東京都→東京, 神奈川県→神奈川, 千葉県→千葉, 埼玉県→埼玉
 * それ以外 → 東京（全員通知扱い）
 * @param {string} address - ホール住所
 * @return {string} エリア名
 */
function addressToArea(address) {
  if (!address) return '東京';
  address = address.toString().trim();
  // まず住所全体から都道府県を検索（先頭以外にあってもOK）
  if (address.indexOf('東京都') !== -1) return '東京';
  if (address.indexOf('神奈川県') !== -1) return '神奈川';
  if (address.indexOf('千葉県') !== -1) return '千葉';
  if (address.indexOf('埼玉県') !== -1) return '埼玉';
  // 都道府県が含まれない場合、主要市区町村名から判定
  var kanagawaCities = ['横浜市','川崎市','相模原市','横須賀市','藤沢市','平塚市','茅ヶ崎市','大和市','海老名市','座間市','秦野市','伊勢原市','小田原市','厚木市','鹾倉市','港北区','みなとみらい'];
  var chibaCities = ['千葉市','船橋市','柏市','松戸市','市川市','浦安市','成田市','木更津市','習志野市','流山市','舵ヶ谷市','幕張メッセ'];
  var saitamaCities = ['さいたま市','川口市','川越市','所沢市','越谷市','春日部市','草加市','大宮区','浦和区','南区'];
  for (var i = 0; i < kanagawaCities.length; i++) {
    if (address.indexOf(kanagawaCities[i]) !== -1) return '神奈川';
  }
  for (var j = 0; j < chibaCities.length; j++) {
    if (address.indexOf(chibaCities[j]) !== -1) return '千葉';
  }
  for (var k = 0; k < saitamaCities.length; k++) {
    if (address.indexOf(saitamaCities[k]) !== -1) return '埼玉';
  }
  // 東京の区名チェック（「港区」「渋谷区」など、23区のいずれか）
  var tokyoWards = ['千代田区','中央区','港区','新宿区','文京区','台東区','墨田区','江東区','品川区','目黒区','大田区','世田谷区','渋谷区','中野区','杉並区','豊島区','北区','荒川区','板橋区','練馬区','足立区','葛飾区','江戸川区'];
  for (var w = 0; w < tokyoWards.length; w++) {
    if (address.indexOf(tokyoWards[w]) !== -1) return '東京';
  }
  return '東京'; // 判定できない場合 → 全員通知扱い
}

/**
 * 『ホール周辺情報の回答』タブの住所からホール名→エリアのマッピングを自動生成
 * B列=ホール名, C列=ホール住所 から都道府県を抽出してエリアに変換
 * @return {Object} { ホール名: エリア } のマップ
 */
function getAreaMap() {
  var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('ホール周辺情報の回答');
  if (!sheet) {
    Logger.log('ホール周辺情報の回答シートが見つかりません');
    return {};
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  var data = sheet.getRange(2, 2, lastRow - 1, 2).getValues(); // B列=ホール名, C列=住所
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var hallName = (data[i][0] || '').toString().trim();
    var address = (data[i][1] || '').toString().trim();
    if (hallName) {
      map[hallName] = addressToArea(address);
    }
  }
  Logger.log('エリアマップ取得（住所ベース）: ' + Object.keys(map).length + '件');
  return map;
}

/**
 * ホールエリアとユーザー希望エリアから通知対象かを判定
 * @param {string} hallArea - 募集ホールのエリア（東京/神奈川/埼玉/千葉）
 * @param {string} userPreferredArea - ユーザーの希望エリア
 * @return {boolean} 通知対象ならtrue
 */
function shouldNotifyByArea(hallArea, userPreferredArea) {
  if (!userPreferredArea) return true;
  if (hallArea === '東京') return true;
  if (userPreferredArea === '全エリア') return true;
  return userPreferredArea === hallArea;
}

// ============================================
// ホール登録フォーム連携
// ============================================

/** 増員募集フォームID */
var RECRUITMENT_FORM_ID = '1YCmH7vW3P2CEb9CXXFZH-axuV9u5guYrwCp6rOyBdX4';

/**
 * 「ホール周辺情報の回答」タブのホール名一覧を、増員募集フォームのプルダウンに自動反映
 * GASエディタから手動実行、またはホール登録フォーム送信時に自動実行
 */
function updateFormDropdown() {
  try {
    var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('ホール周辺情報の回答');
    if (!sheet) {
      Logger.log('updateFormDropdown: ホール周辺情報の回答シートが見つかりません');
      return;
    }
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('updateFormDropdown: データがありません');
      return;
    }
    // B列（ホール名）を取得（重複除去）
    var data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    var hallNames = [];
    for (var i = 0; i < data.length; i++) {
      var name = (data[i][0] || '').toString().trim();
      if (name && hallNames.indexOf(name) === -1) {
        hallNames.push(name);
      }
    }
    if (hallNames.length === 0) {
      Logger.log('updateFormDropdown: 反映するホール名がありません');
      return;
    }
    // 増員募集フォームのプルダウンを更新
    var form = FormApp.openById(RECRUITMENT_FORM_ID);
    var formItems = form.getItems();
    var questionTitle = '増員募集のホールをお答えください';
    var targetItem = null;
    for (var j = 0; j < formItems.length; j++) {
      if (formItems[j].getTitle() === questionTitle) {
        if (formItems[j].getType() === FormApp.ItemType.LIST) {
          targetItem = formItems[j].asListItem();
        } else if (formItems[j].getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
          targetItem = formItems[j].asMultipleChoiceItem();
        }
        break;
      }
    }
    if (targetItem) {
      targetItem.setChoiceValues(hallNames);
      Logger.log('増員募集フォームのホール選択肢を更新: ' + hallNames.join(', '));
    } else {
      Logger.log('updateFormDropdown: 対象の質問が見つかりません。タイトル: ' + questionTitle);
    }
  } catch (error) {
    Logger.log('updateFormDropdown エラー: ' + error.toString());
  }
}

/** 必要事項入力フォームID */
var STAFF_FORM_ID = '1q1-TbIRehgna1Efkt8EAe6UjqS258ht8UtAf3RUfPDg';

/**
 * 「ホール名＆入館」タブのA列から、必要事項入力フォームの
 * 「所属ホールを教えてください。」プルダウンを自動更新
 * GASエディタから手動実行、またはホール登録フォーム送信時に自動実行
 */
function updateStaffFormDropdown() {
  try {
    var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('ホール名＆入館');
    if (!sheet) {
      Logger.log('updateStaffFormDropdown: ホール名＆入館シートが見つかりません');
      return;
    }
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('updateStaffFormDropdown: データがありません');
      return;
    }
    // A列（ホール名）を取得（重複除去）
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var hallNames = [];
    for (var i = 0; i < data.length; i++) {
      var name = (data[i][0] || '').toString().trim();
      if (name && hallNames.indexOf(name) === -1) {
        hallNames.push(name);
      }
    }
    if (hallNames.length === 0) {
      Logger.log('updateStaffFormDropdown: 反映するホール名がありません');
      return;
    }
    // 必要事項入力フォームのプルダウンを更新
    var form = FormApp.openById(STAFF_FORM_ID);
    var formItems = form.getItems();
    var questionTitle = '所属ホールを教えてください。';
    var targetItem = null;
    for (var j = 0; j < formItems.length; j++) {
      if (formItems[j].getTitle() === questionTitle) {
        if (formItems[j].getType() === FormApp.ItemType.LIST) {
          targetItem = formItems[j].asListItem();
        } else if (formItems[j].getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
          targetItem = formItems[j].asMultipleChoiceItem();
        }
        break;
      }
    }
    if (targetItem) {
      targetItem.setChoiceValues(hallNames);
      Logger.log('必要事項入力フォームのホール選択肢を更新: ' + hallNames.join(', '));
    } else {
      Logger.log('updateStaffFormDropdown: 対象の質問が見つかりません。タイトル: ' + questionTitle);
    }
  } catch (error) {
    Logger.log('updateStaffFormDropdown エラー: ' + error.toString());
  }
}

/** ホール情報収集フォームID */
var HALL_INFO_FORM_ID = '1DRBTl8ChMd2aEdsGAStGfFPwVSJ-RK4f1FWqzq12xAA';

/**
 * 「ホール名＆入館」タブのA列から、ホール情報収集フォームの
 * 「ホール名」プルダウンを自動更新
 */
function updateHallInfoFormDropdown() {
  try {
    var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('ホール名＆入館');
    if (!sheet) {
      Logger.log('updateHallInfoFormDropdown: ホール名＆入館シートが見つかりません');
      return;
    }
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('updateHallInfoFormDropdown: データがありません');
      return;
    }
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var hallNames = [];
    for (var i = 0; i < data.length; i++) {
      var name = (data[i][0] || '').toString().trim();
      if (name && hallNames.indexOf(name) === -1) {
        hallNames.push(name);
      }
    }
    if (hallNames.length === 0) {
      Logger.log('updateHallInfoFormDropdown: 反映するホール名がありません');
      return;
    }
    var form = FormApp.openById(HALL_INFO_FORM_ID);
    var formItems = form.getItems();
    var questionTitle = 'ホール名';
    var targetItem = null;
    for (var j = 0; j < formItems.length; j++) {
      if (formItems[j].getTitle() === questionTitle) {
        if (formItems[j].getType() === FormApp.ItemType.LIST) {
          targetItem = formItems[j].asListItem();
        } else if (formItems[j].getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
          targetItem = formItems[j].asMultipleChoiceItem();
        }
        break;
      }
    }
    if (targetItem) {
      targetItem.setChoiceValues(hallNames);
      Logger.log('ホール情報収集フォームのホール名選択肢を更新: ' + hallNames.join(', '));
    } else {
      Logger.log('updateHallInfoFormDropdown: 対象の質問が見つかりません。タイトル: ' + questionTitle);
    }
  } catch (error) {
    Logger.log('updateHallInfoFormDropdown エラー: ' + error.toString());
  }
}

/**
 * スプレッドシート編集時のトリガーハンドラ
 * 「ホール名＆入館」タブのA列が編集されたら、フォームのプルダウンを自動更新
 */
function onSheetEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    var sheetName = sheet.getName().trim();
    // 「ホール名＆入館」タブのA列が編集された場合のみ実行
    if (sheetName === 'ホール名＆入館' && e.range.getColumn() === 1) {
      Logger.log('ホール名＆入館 A列が編集されました。フォーム選択肢を更新します...');
      updateStaffFormDropdown();
      updateFormDropdown();
      updateHallInfoFormDropdown();
    }
  } catch (error) {
    Logger.log('onSheetEdit エラー: ' + error.toString());
  }
}

/**
 * スプレッドシート編集トリガーを設定
 * GASエディタから手動で1回だけ実行してください
 */
function setupSheetEditTrigger() {
  // 既存のonSheetEditトリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onSheetEdit') {
      ScriptApp.deleteTrigger(triggers[i]);
      Logger.log('既存のonSheetEditトリガーを削除');
    }
  }
  // ①スプレッドシートの編集トリガーを設定
  ScriptApp.newTrigger('onSheetEdit')
    .forSpreadsheet(STAFF_CONFIG.SPREADSHEET_ID)
    .onEdit()
    .create();
  Logger.log('onSheetEditトリガーを設定しました（①スプレッドシート編集時に発火）');
}

/**
 * 管理者権限を自動同期
 * 「ホール周辺情報の回答」F列（管理者指名）と「フォームの回答 1」B列（お名前）を照合
 * 一致 → I列=TRUE、不一致 → I列=FALSE（完全自動管理）
 */
function syncAdminPermissions() {
  try {
    var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    // 1. 「ホール周辺情報の回答」F列から全管理者名を取得
    var hallSheet = ss.getSheetByName('ホール周辺情報の回答');
    if (!hallSheet) {
      Logger.log('syncAdmin: ホール周辺情報の回答シートが見つかりません');
      return;
    }
    var hallLastRow = hallSheet.getLastRow();
    var adminNamesNormalized = [];
    if (hallLastRow >= 2) {
      var hallData = hallSheet.getRange(2, 6, hallLastRow - 1, 1).getValues(); // F列
      for (var i = 0; i < hallData.length; i++) {
        var name = (hallData[i][0] || '').toString().trim();
        if (name) {
          var normalized = normalizeName(name);
          if (adminNamesNormalized.indexOf(normalized) === -1) {
            adminNamesNormalized.push(normalized);
          }
        }
      }
    }
    // 2. 「フォームの回答 1」B列と照合し、I列を更新
    var staffSheet = ss.getSheetByName(STAFF_CONFIG.SHEET_NAME);
    if (!staffSheet) {
      Logger.log('syncAdmin: スタッフシートが見つかりません');
      return;
    }
    var staffLastRow = staffSheet.getLastRow();
    if (staffLastRow < 2) return;
    var staffData = staffSheet.getRange(2, 2, staffLastRow - 1, 1).getValues(); // B列（お名前）
    var updatedCount = 0;
    for (var j = 0; j < staffData.length; j++) {
      var staffName = normalizeName(staffData[j][0] || '');
      if (!staffName) continue;
      var shouldBeAdmin = adminNamesNormalized.indexOf(staffName) !== -1;
      staffSheet.getRange(j + 2, 9).setValue(shouldBeAdmin); // I列（管理者権限）
      if (shouldBeAdmin) updatedCount++;
    }
    Logger.log('管理者権限を同期: 管理者 ' + updatedCount + '名 / 全スタッフ ' + staffData.length + '名');
  } catch (error) {
    Logger.log('syncAdminPermissions エラー: ' + error.toString());
  }
}

/**
 * ホール登録フォーム送信時のトリガーハンドラ
 * 同じホール名の旧エントリを上書き（削除）し、
 * 募集フォームのプルダウン更新 + 管理者権限の同期を実行
 */
function onHallFormSubmit(e) {
  try {
    // フォーム送信先のシートを確認
    var isHallForm = true;
    if (e && e.range) {
      var triggerSheet = e.range.getSheet();
      var sheetName = triggerSheet.getName().trim();
      if (sheetName !== 'ホール周辺情報の回答') {
        isHallForm = false;
        // ホール登録以外のフォーム（必要事項の回答等）でも管理者権限は同期する
        syncAdminPermissions();
        return;
      }
    }
    var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('ホール周辺情報の回答');
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    // 最新行（今回送信されたデータ）のホール名を取得
    var newHallName = (sheet.getRange(lastRow, 2).getValue() || '').toString().trim();
    if (!newHallName) return;
    // 同じホール名の旧エントリを削除（上書き動作）
    if (lastRow > 2) {
      var data = sheet.getRange(2, 2, lastRow - 2, 1).getValues(); // 最新行を除く
      for (var i = data.length - 1; i >= 0; i--) { // 下から削除（行番号ズレ防止）
        var existingHall = (data[i][0] || '').toString().trim();
        if (existingHall === newHallName) {
          sheet.deleteRow(i + 2);
          Logger.log('旧エントリを上書き削除: ' + newHallName + ' (行' + (i + 2) + ')');
        }
      }
    }
    // 募集フォームのプルダウンを更新
    updateFormDropdown();
    // 必要事項入力フォームのプルダウンを更新
    updateStaffFormDropdown();
    // ホール情報収集フォームのプルダウンを更新
    updateHallInfoFormDropdown();
    // 管理者権限を同期
    syncAdminPermissions();
    
    // システム管理者へホール情報登録/更新の通知を送信
    try {
      var hallSheet = ss.getSheetByName('ホール周辺情報の回答');
      var updatedLastRow = hallSheet ? hallSheet.getLastRow() : 0;
      var hallResponses = {};
      if (hallSheet && updatedLastRow >= 2) {
        var rowData = hallSheet.getRange(updatedLastRow, 1, 1, 9).getValues()[0];
        hallResponses = {
          hallName: (rowData[1] || '').toString().trim(),
          address: (rowData[2] || '').toString().trim(),
          phone: (rowData[3] || '').toString().trim(),
          email: (rowData[4] || '').toString().trim(),
          adminName: (rowData[5] || '').toString().trim(),
          nearby1: (rowData[6] || '').toString().trim(),
          nearby2: (rowData[7] || '').toString().trim(),
          smoking: (rowData[8] || '').toString().trim()
        };
      }
      notifySystemAdminsHallInfoUpdate(newHallName, hallResponses);
    } catch (notifyErr) {
      Logger.log('システム管理者通知エラー: ' + notifyErr.toString());
    }
    
    Logger.log('ホール登録フォーム処理完了: ' + newHallName);
  } catch (error) {
    Logger.log('onHallFormSubmit エラー: ' + error.toString());
  }
}

/**
 * ホール登録フォームのonFormSubmitトリガーを設定
 * GASエディタから手動で1回だけ実行してください
 */
function setupHallFormTrigger() {
  // 既存のonHallFormSubmitトリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onHallFormSubmit') {
      ScriptApp.deleteTrigger(triggers[i]);
      Logger.log('既存のonHallFormSubmitトリガーを削除');
    }
  }
  // ①スプレッドシートのフォーム送信トリガーを設定
  ScriptApp.newTrigger('onHallFormSubmit')
    .forSpreadsheet(STAFF_CONFIG.SPREADSHEET_ID)
    .onFormSubmit()
    .create();
  Logger.log('onHallFormSubmitトリガーを設定しました（①スプレッドシートのフォーム送信時に発火）');
}

/**
 * エリア判定テスト用（GASエディタから手動実行）
 */
function testGetAreaMap() {
  var map = getAreaMap();
  for (var hall in map) {
    Logger.log(hall + ' → ' + map[hall]);
  }
}

// ============================================
// Web版ホール情報API
// ============================================

/**
 * Web版ホール一覧API
 * @return {object} { success: true, halls: [...] }
 */
function getHallListForWeb() {
  try {
    var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('ホール周辺情報の回答');
    if (!sheet) return { success: true, halls: [] };
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, halls: [] };
    
    var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    var hallMap = {};
    
    // 同じホール名は最新(最後)のデータを使用
    for (var i = 0; i < data.length; i++) {
      var name = (data[i][1] || '').toString().trim();
      if (name) {
        hallMap[name] = {
          hallName: name,
          address: (data[i][2] || '').toString().trim(),
          phone: (data[i][3] || '').toString().trim(),
          email: (data[i][4] || '').toString().trim(),
          adminName: (data[i][5] || '').toString().trim(),
          nearby1: (data[i][6] || '').toString().trim(),
          nearby2: (data[i][7] || '').toString().trim(),
          smoking: (data[i][8] || '').toString().trim()
        };
      }
    }
    
    // 入館画像URLも取得
    var entrySheet = ss.getSheetByName('ホール名＆入館');
    if (entrySheet) {
      var entryLastRow = entrySheet.getLastRow();
      if (entryLastRow >= 2) {
        var entryData = entrySheet.getRange(2, 1, entryLastRow - 1, 2).getValues();
        for (var j = 0; j < entryData.length; j++) {
          var eName = (entryData[j][0] || '').toString().trim();
          var eUrl = (entryData[j][1] || '').toString().trim();
          if (eName && hallMap[eName] && eUrl) {
            hallMap[eName].entryImageUrl = convertDriveUrl(eUrl);
          }
        }
      }
    }
    
    var halls = Object.keys(hallMap).map(function(k) { return hallMap[k]; });
    return { success: true, halls: halls };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Web版ホール詳細API
 * @param {string} hallName
 * @return {object} { success: true, hall: { ... } }
 */
function getHallDetailForWeb(hallName) {
  try {
    if (!hallName) return { success: false, error: 'ホール名が必要です' };
    
    var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('ホール周辺情報の回答');
    if (!sheet) return { success: true, hall: null };
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, hall: null };
    
    var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    var result = null;
    for (var i = 0; i < data.length; i++) {
      var name = (data[i][1] || '').toString().trim();
      if (name === hallName) {
        result = {
          hallName: name,
          address: (data[i][2] || '').toString().trim(),
          phone: (data[i][3] || '').toString().trim(),
          email: (data[i][4] || '').toString().trim(),
          adminName: (data[i][5] || '').toString().trim(),
          nearby1: (data[i][6] || '').toString().trim(),
          nearby2: (data[i][7] || '').toString().trim(),
          smoking: (data[i][8] || '').toString().trim()
        };
      }
    }
    
    // 入館画像URL
    if (result) {
      var entrySheet = ss.getSheetByName('ホール名＆入館');
      if (entrySheet) {
        var entryLastRow = entrySheet.getLastRow();
        if (entryLastRow >= 2) {
          var entryData = entrySheet.getRange(2, 1, entryLastRow - 1, 2).getValues();
          for (var j = 0; j < entryData.length; j++) {
            if ((entryData[j][0] || '').toString().trim() === hallName) {
              result.entryImageUrl = convertDriveUrl((entryData[j][1] || '').toString().trim());
              break;
            }
          }
        }
      }
    }
    
    return { success: true, hall: result };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}


// ============================================
// 新規募集通知（セクション別＋エリア判定）
// ============================================
function notifyNewRecruitment(eventData) {
  try {
    // 募集セクションを特定
    const recruitSections = [];
    eventData.days.forEach(function(day) {
      if (day.stage > 0) recruitSections.push('舞台');
      if (day.sound > 0) recruitSections.push('音響');
      if (day.lighting > 0) recruitSections.push('照明');
    });
    
    // 重複を除去
    const uniqueSections = recruitSections.filter(function(v, i, a) { return a.indexOf(v) === i; });
    
    if (uniqueSections.length === 0) {
      Logger.log('通知対象セクションなし');
      return;
    }
    
    Logger.log('募集セクション: ' + uniqueSections.join(', '));
    
    // エリアマスタを取得し、募集ホールのエリアを特定
    var areaMap = getAreaMap();
    var hallArea = areaMap[eventData.hall] || '';
    Logger.log('募集ホール: ' + eventData.hall + ' → エリア: ' + (hallArea || '未設定'));
    
    // LINE友だちシートからUID+セクション+希望エリアを取得
    var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(LINE_CONFIG.ADMIN_SHEET_NAME);
    if (!sheet) {
      Logger.log('LINE友だちシートが見つかりません');
      return;
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('LINE友だちシートにデータがありません');
      return;
    }
    
    // A列=UID, B列=表示名, C列=スタッフ名, F列=セクション(6列目), I列=希望エリア(9列目)
    var lastCol = Math.max(sheet.getLastColumn(), 9);
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    var notifiedCount = 0;
    
    for (var i = 0; i < data.length; i++) {
      var uid = (data[i][0] || '').toString().trim();        // A列: LINE UID
      var staffName = (data[i][2] || data[i][1] || '').toString().trim(); // C列 or B列
      var staffSection = (data[i][5] || '').toString().trim(); // F列(6列目): セクション
      var userArea = (data[i][8] || '').toString().trim();   // I列(9列目): 希望エリア
      
      if (!uid) continue;
      
      // セクション情報がない場合はスキップ
      var shouldNotify = false;
      if (!staffSection) {
        // セクション未設定 → スキップ
        Logger.log('セクション未設定のためスキップ: ' + staffName);
        continue;
      } else {
        // セクションが一致するか確認
        for (var j = 0; j < uniqueSections.length; j++) {
          if (staffSection.indexOf(uniqueSections[j]) >= 0) {
            shouldNotify = true;
            break;
          }
        }
      }
      
      if (!shouldNotify) continue;
      
      // エリア判定（エリアマスタにホールが登録されていない場合はスキップしない）
      if (hallArea) {
        if (!shouldNotifyByArea(hallArea, userArea)) {
          Logger.log('エリア不一致のためスキップ: ' + staffName + ' (希望: ' + userArea + ', 募集: ' + hallArea + ')');
          continue;
        }
      }
      
      // 通知メッセージを送信
      var message = createNewRecruitmentMessage(eventData, uniqueSections, staffName);
      
      try {
        UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
          method: 'post',
          headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_CONFIG.CHANNEL_ACCESS_TOKEN },
          payload: JSON.stringify({ to: uid, messages: [message] }),
          muteHttpExceptions: true
        });
        notifiedCount++;
        Logger.log('通知送信: ' + staffName + ' (セクション: ' + staffSection + ', エリア: ' + userArea + ')');
      } catch (pushError) {
        Logger.log('Push送信エラー(' + staffName + '): ' + pushError.toString());
      }
    }
    
    Logger.log('新規募集通知完了: ' + notifiedCount + '名に送信');
  } catch (error) {
    Logger.log('notifyNewRecruitment エラー: ' + error.toString());
  }
}

function createNewRecruitmentMessage(eventData, sections, staffName) {
  // Web App URL（カレンダーへのリンク）パラメーターで特定の日付・事業所を開く
  var firstDate = eventData.days[0].date;
  var formattedDate = firstDate.getFullYear() + '-' + ('0' + (firstDate.getMonth()+1)).slice(-2) + '-' + ('0' + firstDate.getDate()).slice(-2);
  var calendarUrl = 'https://liff.line.me/2009354296-Ae3FpVy2?date=' + formattedDate + '&hall=' + encodeURIComponent(eventData.hall || '');
  // 日付情報を整形
  var dateTexts = eventData.days.map(function(day) {
    var d = day.date;
    var weekdays = ['日', '月', '火', '水', '木', '金', '土'];
    return (d.getMonth() + 1) + '月' + d.getDate() + '日（' + weekdays[d.getDay()] + '）';
  });
  
  // 区分情報を取得
  var timeSlots = [];
  eventData.days.forEach(function(day) {
    if (day.timeSlot && timeSlots.indexOf(day.timeSlot) < 0) {
      timeSlots.push(day.timeSlot);
    }
  });
  
  // セクション別人数を集計（重複除去）
  var stageTotal = 0, soundTotal = 0, lightingTotal = 0;
  eventData.days.forEach(function(day) {
    if (day.stage > 0) stageTotal = Math.max(stageTotal, day.stage);
    if (day.sound > 0) soundTotal = Math.max(soundTotal, day.sound);
    if (day.lighting > 0) lightingTotal = Math.max(lightingTotal, day.lighting);
  });
  
  // セクション色設定
  var sectionColors = {
    stage:    { main: '#10b981', dark: '#0a2618', header: '#0a5c40', label: '舞台', bodyBg: '#061a12', cardBg: '#0c3024', infoTitle: '#6ee7b7', labelCol: '#5eaa8d' },
    sound:    { main: '#0ea5e9', dark: '#0a1a28', header: '#075985', label: '音響', bodyBg: '#061422', cardBg: '#0c2440', infoTitle: '#7dd3fc', labelCol: '#6b9cc0' },
    lighting: { main: '#eab308', dark: '#1a1808', header: '#854d0e', label: '照明', bodyBg: '#1a1508', cardBg: '#2a2210', infoTitle: '#fde68a', labelCol: '#a89860' }
  };
  
  // ヘッダー色とラベルを決定
  var activeSections = [];
  if (stageTotal > 0) activeSections.push('stage');
  if (soundTotal > 0) activeSections.push('sound');
  if (lightingTotal > 0) activeSections.push('lighting');
  
  var headerColor, headerDark, sectionLabel, bodyBg, cardBg, infoTitleColor, labelColor;
  if (activeSections.length > 1) {
    headerColor = '#ef4444';
    headerDark = '#7f1d1d';
    bodyBg = '#180808';
    cardBg = '#2a1010';
    infoTitleColor = '#fca5a5';
    labelColor = '#b06060';
    var labels = activeSections.map(function(s) { return sectionColors[s].label; });
    sectionLabel = labels.join('・') + '増員';
  } else if (activeSections.length === 1) {
    var sec = sectionColors[activeSections[0]];
    headerColor = sec.main;
    headerDark = sec.header;
    bodyBg = sec.bodyBg;
    cardBg = sec.cardBg;
    infoTitleColor = sec.infoTitle;
    labelColor = sec.labelCol;
    sectionLabel = sec.label + '増員';
  } else {
    headerColor = '#ef4444';
    headerDark = '#7f1d1d';
    bodyBg = '#180808';
    cardBg = '#2a1010';
    infoTitleColor = '#fca5a5';
    labelColor = '#b06060';
    sectionLabel = '増員';
  }
  
  // ボタン色（ヘッダーと同じ色）
  var buttonColor = headerColor;
  
  // 日程/区分を結合: 各日付にtimeSlotを添付
  var dateWithSlotTexts = eventData.days.map(function(day) {
    var d = day.date;
    var weekdays = ['日', '月', '火', '水', '木', '金', '土'];
    var dateStr = (d.getMonth() + 1) + '月' + d.getDate() + '日(' + weekdays[d.getDay()] + ')';
    var slot = day.timeSlot || '';
    return slot ? dateStr + ' ' + slot : dateStr;
  });
  var timeSlotsText = timeSlots.length > 0 ? timeSlots.join('・') : '';
  var dateSlotDisplay = dateWithSlotTexts.join('\n');
  
  // 催事情報カードの内容
  var eventInfoContents = [
    { type: 'box', layout: 'horizontal', contents: [
      { type: 'text', text: '催事名', size: 'sm', color: labelColor, flex: 1, align: 'start' },
      { type: 'text', text: eventData.eventName || '-', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
    ], margin: 'lg' },
    { type: 'box', layout: 'horizontal', contents: [
      { type: 'text', text: '日程/区分', size: 'sm', color: labelColor, flex: 1, align: 'start' },
      { type: 'text', text: dateSlotDisplay, size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
    ], margin: 'md' }
  ];
  
  // 若手OK表示
  if (eventData.juniorOk === 'はい') {
    eventInfoContents.push({ type: 'box', layout: 'horizontal', contents: [
      { type: 'text', text: '若手', size: 'sm', color: labelColor, flex: 1, align: 'start' },
      { type: 'text', text: '🔰 OK', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start' }
    ], margin: 'md' });
  }
  
  // 複数日募集の場合、通し希望を追加
  var pref = eventData.consecutivePreference;
  if (eventData.days && eventData.days.length > 1 && pref) {
    var prefText = pref;
    if (pref === 'はい' || pref === 'あり' || pref === '希望する') {
      prefText = '希望あり';
    } else if (pref === 'なし' || pref === 'いいえ') {
      prefText = '希望なし';
    } else if (pref === '可能であれば') {
      prefText = '可能であれば';
    }
    
    eventInfoContents.push({ type: 'box', layout: 'horizontal', contents: [
      { type: 'text', text: '通し希望', size: 'sm', color: labelColor, flex: 1, align: 'start' },
      { type: 'text', text: prefText, size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
    ], margin: 'md' });
  }
  
  // セクション別カード生成
  var sectionCards = [];
  if (stageTotal > 0) {
    sectionCards.push({
      type: 'box', layout: 'horizontal',
      contents: [
        { type: 'box', layout: 'vertical', contents: [{ type: 'filler' }], width: '4px', backgroundColor: '#10b981' },
        { type: 'box', layout: 'horizontal', contents: [
          { type: 'text', text: '舞台', size: 'sm', color: '#ffffff', weight: 'bold', flex: 1 },
          { type: 'text', text: stageTotal + '名', size: 'sm', color: '#ffffff', weight: 'bold', flex: 0, align: 'end' }
        ], paddingAll: '12px', flex: 1 }
      ],
      backgroundColor: '#0a2618', cornerRadius: 'md', margin: 'sm'
    });
  }
  if (soundTotal > 0) {
    sectionCards.push({
      type: 'box', layout: 'horizontal',
      contents: [
        { type: 'box', layout: 'vertical', contents: [{ type: 'filler' }], width: '4px', backgroundColor: '#0ea5e9' },
        { type: 'box', layout: 'horizontal', contents: [
          { type: 'text', text: '音響', size: 'sm', color: '#ffffff', weight: 'bold', flex: 1 },
          { type: 'text', text: soundTotal + '名', size: 'sm', color: '#ffffff', weight: 'bold', flex: 0, align: 'end' }
        ], paddingAll: '12px', flex: 1 }
      ],
      backgroundColor: '#0a1a28', cornerRadius: 'md', margin: 'sm'
    });
  }
  if (lightingTotal > 0) {
    sectionCards.push({
      type: 'box', layout: 'horizontal',
      contents: [
        { type: 'box', layout: 'vertical', contents: [{ type: 'filler' }], width: '4px', backgroundColor: '#eab308' },
        { type: 'box', layout: 'horizontal', contents: [
          { type: 'text', text: '照明', size: 'sm', color: '#ffffff', weight: 'bold', flex: 1 },
          { type: 'text', text: lightingTotal + '名', size: 'sm', color: '#ffffff', weight: 'bold', flex: 0, align: 'end' }
        ], paddingAll: '12px', flex: 1 }
      ],
      backgroundColor: '#1a1808', cornerRadius: 'md', margin: 'sm'
    });
  }
  
  // ボディ全体を組み立て
  var bodyContents = [
    { type: 'text', text: sectionLabel, size: 'xs', color: headerColor, align: 'center', margin: 'none' },
    { type: 'text', text: eventData.hall || '未定', size: 'xl', color: '#ffffff', weight: 'bold', align: 'center', margin: 'sm' },
    { type: 'text', text: '催事情報', weight: 'bold', size: 'md', color: infoTitleColor, align: 'center', margin: 'lg' },
    { type: 'box', layout: 'vertical', contents: eventInfoContents, backgroundColor: cardBg, cornerRadius: 'md', paddingAll: '14px', margin: 'sm' },
    { type: 'text', text: '募集セクション', weight: 'bold', size: 'md', color: infoTitleColor, align: 'center', margin: 'lg' }
  ];
  
  // セクションカードを追加
  sectionCards.forEach(function(card) { bodyContents.push(card); });
  
  // カレンダーボタンを追加
  bodyContents.push({
    type: 'button',
    action: { type: 'uri', label: '🔍 詳細確認＆応募はこちらから', uri: calendarUrl },
    style: 'primary',
    color: buttonColor,
    margin: 'xl',
    height: 'md'
  });
  
  return {
    type: 'flex',
    altText: '🔔 新規増員募集 - ' + (eventData.hall || '') + ' ' + sectionLabel,
    contents: {
      type: 'bubble',
      styles: {
        header: { backgroundColor: headerDark },
        body: { backgroundColor: bodyBg },
        footer: { backgroundColor: bodyBg }
      },
      header: {
        type: 'box', layout: 'vertical',
        contents: [
          {
            type: 'box', layout: 'vertical',
            contents: [{ type: 'text', text: '🔔 新規増員募集 🔔', weight: 'bold', size: 'lg', color: '#ffffff', align: 'center' }],
            backgroundColor: headerColor, cornerRadius: 'md', paddingAll: '16px'
          }
        ],
        backgroundColor: headerDark, paddingAll: '16px'
      },
      body: {
        type: 'box', layout: 'vertical',
        contents: bodyContents,
        backgroundColor: bodyBg, paddingAll: '18px'
      },
      footer: {
        type: 'box', layout: 'vertical',
        contents: [
          { type: 'filler' }
        ],
        backgroundColor: bodyBg, paddingAll: '0px', height: '4px'
      }
    }
  };
}

// ============================================
// LINE友だちシートのスタッフ情報を一括補完（手動実行用）
// ============================================
function refreshLineFriendsInfo() {
  var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(LINE_CONFIG.ADMIN_SHEET_NAME);
  if (!sheet) { Logger.log('LINE友だちシートがありません'); return; }
  
  // ヘッダー確認・追加
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.length < 6) sheet.getRange(1, 6).setValue('セクション');
  if (headers.length < 7) sheet.getRange(1, 7).setValue('ホール');
  if (headers.length < 8) sheet.getRange(1, 8).setValue('種別');
  if (headers.length < 9) sheet.getRange(1, 9).setValue('希望エリア');
  sheet.getRange(1, 4).setValue('メール');
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('データがありません'); return; }
  
  var data = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 8)).getValues();
  var updated = 0;
  
  for (var i = 0; i < data.length; i++) {
    var displayName = (data[i][1] || '').toString().trim(); // B列: LINE表示名
    if (!displayName) continue;
    
    var staffInfo = lookupStaffByName(displayName);
    if (!staffInfo.staffName) continue;
    
    var row = i + 2;
    if (staffInfo.staffName) sheet.getRange(row, 3).setValue(staffInfo.staffName);
    if (staffInfo.email) sheet.getRange(row, 4).setValue(staffInfo.email);
    if (staffInfo.section) sheet.getRange(row, 6).setValue(staffInfo.section);
    if (staffInfo.hall) sheet.getRange(row, 7).setValue(staffInfo.hall);
    if (staffInfo.preferredArea) sheet.getRange(row, 9).setValue(staffInfo.preferredArea);
    updated++;
    Logger.log('更新: ' + displayName + ' → セクション:' + staffInfo.section + ' ホール:' + staffInfo.hall + ' エリア:' + (staffInfo.preferredArea || ''));
  }
  
  Logger.log('完了: ' + updated + '名のスタッフ情報を補完しました');
}

// ============================================
// スタッフ登録フォーム送信時の自動UID紐付け
// ※トリガー設定: スタッフ管理スプシ → フォーム送信時
// ============================================
function onStaffFormSubmit(e) {
  try {
    var responses = e.values;
    if (!responses || responses.length < 3) return;
    
    // フォーム回答から情報取得
    var staffName = (responses[1] || '').trim();        // B列: お名前(フルネーム)
    var lineDisplayName = (responses[2] || '').trim();  // C列: LINE表示名
    var staffHall = (responses[3] || '').trim();        // D列: 所属ホール
    var staffSection = (responses[4] || '').trim();     // E列: セクション
    var staffEmail = (responses[5] || '').trim();       // F列: メール
    var preferredArea = (responses[6] || '').trim();    // G列: 希望エリア
    
    if (!lineDisplayName) {
      Logger.log('LINE表示名が空のためスキップ');
      return;
    }
    
    Logger.log('スタッフ登録: ' + staffName + ' (LINE: ' + lineDisplayName + ')');
    
    // LINE友だちシートでLINE表示名を検索
    var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(LINE_CONFIG.ADMIN_SHEET_NAME);
    if (!sheet) {
      Logger.log('LINE友だちシートがありません');
      return;
    }
    
    // ヘッダー確認
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.length < 6) sheet.getRange(1, 6).setValue('セクション');
    if (headers.length < 7) sheet.getRange(1, 7).setValue('ホール');
    if (headers.length < 8) sheet.getRange(1, 8).setValue('種別');
    if (headers.length < 9) sheet.getRange(1, 9).setValue('希望エリア');
    sheet.getRange(1, 4).setValue('メール');
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('LINE友だちシートにデータがありません');
      return;
    }
    
    var data = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 8)).getValues();
    var normalizedLineDisplay = normalizeName(lineDisplayName);
    var matched = false;
    
    for (var i = 0; i < data.length; i++) {
      var sheetDisplayName = (data[i][1] || '').toString().trim(); // B列: LINE表示名
      
      if (normalizeName(sheetDisplayName) === normalizedLineDisplay) {
        var row = i + 2;
        sheet.getRange(row, 3).setValue(staffName);      // C列: スタッフ名
        sheet.getRange(row, 4).setValue(staffEmail);      // D列: メール
        sheet.getRange(row, 6).setValue(staffSection);    // F列: セクション
        sheet.getRange(row, 7).setValue(staffHall);       // G列: ホール
        if (preferredArea) sheet.getRange(row, 9).setValue(preferredArea); // I列: 希望エリア
        matched = true;
        Logger.log('LINE友だち紐付け完了: ' + lineDisplayName + ' → ' + staffName + ' (セクション: ' + staffSection + ', エリア: ' + preferredArea + ')');
        break;
      }
    }
    
    if (!matched) {
      Logger.log('LINE友だちシートに該当なし: ' + lineDisplayName);
    }
  } catch (error) {
    Logger.log('onStaffFormSubmit エラー: ' + error.toString());
  }
}

// ============================================
// スプレッドシート削除時のカレンダーイベント連動削除
// ※トリガー設定: 時間主導型（5分おき等）で実行
// ============================================
function syncDeleteCalendarEvents() {
  try {
    var formSs = SpreadsheetApp.openById('1B_oyMPAaEq8rx3A8LcJUQsnNAhr7QnG-zvafg9peBNE');
    var formSheet = formSs.getSheetByName('フォームの回答 1');
    if (!formSheet) {
      Logger.log('フォーム回答シートが見つかりません');
      return;
    }
    
    var calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    if (!calendar) {
      Logger.log('カレンダーが見つかりません');
      return;
    }
    
    // スプレッドシートから全イベントデータを取得
    var lastRow = formSheet.getLastRow();
    var sheetEvents = {};
    
    if (lastRow >= 2) {
      var data = formSheet.getRange(2, 1, lastRow - 1, 24).getValues();
      for (var i = 0; i < data.length; i++) {
        // 各日付+ホール+催事名をキーとして記録
        var hall = (data[i][COL.HALL] || '').toString().trim();
        var eventName = (data[i][COL.EVENT_NAME] || '').toString().trim();
        
        // 日付1
        var date1 = parseDate(data[i][COL.DATE_1]);
        if (date1 && hall) {
          var key1 = makeEventKey(date1, hall);
          sheetEvents[key1] = true;
        }
        
        // 日付2
        var hasDay2 = data[i][COL.HAS_DAY2] && data[i][COL.HAS_DAY2] !== 'いいえ';
        if (hasDay2) {
          var date2 = parseDate(data[i][COL.DATE_2]);
          if (date2 && hall) {
            var key2 = makeEventKey(date2, hall);
            sheetEvents[key2] = true;
          }
        }
        
        // 日付3
        var hasDay3 = data[i][COL.HAS_DAY3] && data[i][COL.HAS_DAY3] !== 'いいえ';
        if (hasDay3) {
          var date3 = parseDate(data[i][COL.DATE_3]);
          if (date3 && hall) {
            var key3 = makeEventKey(date3, hall);
            sheetEvents[key3] = true;
          }
        }
      }
    }
    
    Logger.log('スプレッドシートのイベント数: ' + Object.keys(sheetEvents).length);
    
    // カレンダーから今後730日間のイベントを取得
    var now = new Date();
    var future = new Date(now.getTime() + 730 * 24 * 60 * 60 * 1000);
    var calEvents = calendar.getEvents(now, future);
    
    var deletedCount = 0;
    for (var j = 0; j < calEvents.length; j++) {
      var calEvent = calEvents[j];
      var title = calEvent.getTitle();
      
      // 増員イベントのみ対象
      if (!title.includes('【増員】')) continue;
      
      // タイトルからホール名を抽出
      var hallMatch = title.match(/【増員】(.+?)(?:\s*\(|$)/);
      var calHall = hallMatch ? hallMatch[1].trim() : '';
      
      var calDate = calEvent.getAllDayStartDate();
      if (!calDate || !calHall) continue;
      
      var calKey = makeEventKey(calDate, calHall);
      
      // スプレッドシートに存在しない場合 → 削除
      if (!sheetEvents[calKey]) {
        Logger.log('カレンダーイベント削除: ' + title + ' (' + calKey + ')');
        calEvent.deleteEvent();
        deletedCount++;
      }
    }
    
    Logger.log('カレンダー同期完了: ' + deletedCount + '件削除');
  } catch (error) {
    Logger.log('syncDeleteCalendarEvents エラー: ' + error.toString());
  }
}

function makeEventKey(date, hall) {
  var d = new Date(date);
  return d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate() + '_' + hall;
}

// ============================================
// 手動テスト用: GASエディタから実行してFlexメッセージを送信テスト
// ============================================
function testFlexMessage() {
  var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(LINE_CONFIG.ADMIN_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('LINE友だちシートにデータがありません');
    return;
  }
  
  var uid = sheet.getRange(2, 1).getValue().toString().trim();
  var displayName = sheet.getRange(2, 2).getValue().toString().trim();
  Logger.log('テスト送信先: UID=' + uid + ', 名前=' + displayName);
  
  if (!uid) {
    Logger.log('UIDが空です');
    return;
  }
  
  var message = buildFlexCard('3月28日', {
    hall: 'テストホール',
    staffName: displayName || 'テスト太郎',
    section: '音響'
  });
  
  var response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_CONFIG.CHANNEL_ACCESS_TOKEN },
    payload: JSON.stringify({ to: uid, messages: [message] }),
    muteHttpExceptions: true
  });
  
  Logger.log('レスポンスコード: ' + response.getResponseCode());
  Logger.log('レスポンス内容: ' + response.getContentText());
}

// ============================================
// 募集リマインド通知
// 募集日の1ヶ月前になっても募集終了していないイベントを再通知
// ※トリガー設定: 時間主導型（毎日1回等）で実行
// ============================================
function _OLD_sendRecruitmentReminders() { return; // DEPRECATED - checkAndSendReminders() に置き換え済み
  try {
    var calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    if (!calendar) {
      Logger.log('カレンダーが見つかりません');
      return;
    }
    
    // 募集完了ステータスを取得
    var closedStatuses = {};
    try {
      var mgmtSheet = getOrCreateManagementSheet();
      var mgmtLastRow = mgmtSheet.getLastRow();
      if (mgmtLastRow >= 2) {
        var mgmtData = mgmtSheet.getRange(2, 1, mgmtLastRow - 1, 3).getValues();
        for (var m = 0; m < mgmtData.length; m++) {
          if (mgmtData[m][2] === true || mgmtData[m][2] === 'TRUE') {
            closedStatuses[mgmtData[m][0]] = true;
          }
        }
      }
    } catch (e) {
      Logger.log('ステータス取得エラー: ' + e.toString());
    }
    
    // 今日から25日～35日後のイベントを取得（約1ヶ月前）
    var now = new Date();
    var start = new Date(now.getTime() + 25 * 24 * 60 * 60 * 1000);
    var end = new Date(now.getTime() + 35 * 24 * 60 * 60 * 1000);
    var calEvents = calendar.getEvents(start, end);
    
    // LINE友だちシートからUIDとセクションを取得
    var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    var friendSheet = ss.getSheetByName(LINE_CONFIG.ADMIN_SHEET_NAME);
    if (!friendSheet || friendSheet.getLastRow() < 2) {
      Logger.log('LINE友だちシートにデータがありません');
      return;
    }
    var lastCol = Math.max(friendSheet.getLastColumn(), 9);
    var friendData = friendSheet.getRange(2, 1, friendSheet.getLastRow() - 1, lastCol).getValues();
    
    // エリアマスタを取得
    var areaMap = getAreaMap();
    
    var reminderCount = 0;
    
    for (var j = 0; j < calEvents.length; j++) {
      var calEvent = calEvents[j];
      var title = calEvent.getTitle();
      
      // 増員イベントのみ対象
      if (!title.includes('【増員】')) continue;
      
      // イベントキーを生成して募集完了かチェック
      var eventDate = calEvent.getAllDayStartDate();
      var hallMatch = title.match(/【増員】(.+?)(?:\s*\(|$)/);
      var eventHall = hallMatch ? hallMatch[1].trim() : '';
      
      // イベントキーで募集完了チェック
      var eventKey = Utilities.formatDate(eventDate, 'Asia/Tokyo', 'yyyy-MM-dd') + '_' + eventHall;
      if (closedStatuses[eventKey]) {
        Logger.log('募集完了済みスキップ: ' + title);
        continue;
      }
      
      // イベントの説明からセクションと人数を抽出
      var description = calEvent.getDescription() || '';
      var sectionDetails = [];
      var recruitSections = [];
      
      var stageMatch = description.match(/舞台[:|：]\s*(\d+)人/);
      var soundMatch = description.match(/音響[:|：]\s*(\d+)人/);
      var lightMatch = description.match(/照明[:|：]\s*(\d+)人/);
      
      if (stageMatch) { sectionDetails.push('・舞台：' + stageMatch[1] + '人'); recruitSections.push('舞台'); }
      if (soundMatch) { sectionDetails.push('・音響：' + soundMatch[1] + '人'); recruitSections.push('音響'); }
      if (lightMatch) { sectionDetails.push('・照明：' + lightMatch[1] + '人'); recruitSections.push('照明'); }
      
      if (sectionDetails.length === 0) {
        Logger.log('セクション情報なしスキップ: ' + title);
        continue;
      }
      
      // 日付フォーマット
      var dateText = (eventDate.getMonth() + 1) + '月' + eventDate.getDate() + '日';
      
      // 時間帯情報
      var timeSlotMatch = description.match(/【利用区分】\s*\n\s*(.+)/);
      var timeSlot = timeSlotMatch ? timeSlotMatch[1].trim() : '';
      
      // 若手OKチェック
      var isJuniorOk = description.includes('若手可') || title.includes('🔰');
      
      // リマインドメッセージを生成
      var text = '⚠️ 増員募集再通知⚠️\n\n';
      text += '📍 ' + eventHall + '\n';
      text += '📅 ' + dateText + '\n';
      if (timeSlot) text += '🕔 ' + timeSlot + '\n';
      text += '\n🎭 募集セクション\n';
      text += sectionDetails.join('\n') + '\n';
      if (isJuniorOk) text += '\n🔰若手OK🔰\n';
      text += '\n🔍 詳細 確認＆🏃\u200d♂️ 応募方法\n';
      text += '　メニューの【SIGMA】をタップし\n';
      text += '増員カレンダーからお願いします。';
      
      var message = { type: 'text', text: text };
      
      // 該当セクションのスタッフに通知
      for (var k = 0; k < friendData.length; k++) {
        var friendUid = (friendData[k][0] || '').toString().trim();
        var friendSection = (friendData[k][5] || '').toString().trim(); // F列: セクション
        
        if (!friendUid || !friendSection) continue;
        
        var shouldNotify = false;
        for (var s = 0; s < recruitSections.length; s++) {
          if (friendSection.indexOf(recruitSections[s]) >= 0) {
            shouldNotify = true;
            break;
          }
        }
        
        if (!shouldNotify) continue;
        
        // エリア判定
        var eventArea = areaMap[eventHall] || '';
        if (eventArea) {
          var friendArea = (friendData[k][8] || '').toString().trim(); // I列: 希望エリア
          if (!shouldNotifyByArea(eventArea, friendArea)) continue;
        }
        
        try {
          UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
            method: 'post',
            headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_CONFIG.CHANNEL_ACCESS_TOKEN },
            payload: JSON.stringify({ to: friendUid, messages: [message] }),
            muteHttpExceptions: true
          });
          reminderCount++;
        } catch (pushErr) {
          Logger.log('リマインド送信エラー: ' + pushErr.toString());
        }
      }
      
      Logger.log('リマインド送信: ' + title);
    }
    
    Logger.log('リマインド完了: ' + reminderCount + '件送信');
  } catch (error) {
    Logger.log('sendRecruitmentReminders エラー: ' + error.toString());
  }
}

// ============================================
// [デバッグ用] カレンダーイベントのdescriptionを確認
// GASエディタからこの関数を手動実行し、ログで【催事名】の有無を確認
// ============================================
function debugCheckDescriptions() {
  var calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  var now = new Date();
  var future = new Date(now.getTime() + 60 * 24 * 60 * 60 * 1000);
  var events = calendar.getEvents(now, future);
  
  events.forEach(function(event) {
    var title = event.getTitle();
    var desc = event.getDescription() || '';
    var has催事名 = desc.indexOf('【催事名】') >= 0;
    var parsed = parseDescription(desc);
    
    Logger.log('---');
    Logger.log('タイトル: ' + title);
    Logger.log('【催事名】含む: ' + has催事名);
    Logger.log('parsed.eventName: ' + parsed.eventName);
    Logger.log('parsed.content: ' + parsed.content);
    Logger.log('description全文:\n' + desc);
  });
}

// ============================================
// 増員募集リマインド通知（1回限定・フォーム送信から7日経過）
// ※セクション判定・エリア判定を既存通知と同様に適用
// ============================================

// 再通知済みフラグ列インデックス（0-indexed、フォーム回答シートの空き列Y列）
var COL_REMINDER_FLAG = 24;

/**
 * フォーム送信から7日以上経過し、未終了かつ未通知の募集に再通知を送る
 * トリガー: 時間主導型（毎日午前10時）で実行
 * setupDailyReminderTrigger() をGASエディタから1回実行してトリガー登録
 */
function checkAndSendReminders() {
  try {
    // フォーム回答スプレッドシートを開く
    var formSs = SpreadsheetApp.openById('1B_oyMPAaEq8rx3A8LcJUQsnNAhr7QnG-zvafg9peBNE');
    var formSheet = formSs.getSheetByName('増員募集の回答') || formSs.getSheetByName('フォームの回答 1');
    if (!formSheet) {
      Logger.log('リマインド: フォーム回答シートが見つかりません');
      return;
    }
    
    var lastRow = formSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('リマインド: フォーム回答データがありません');
      return;
    }
    
    // ヘッダーに「再通知済み」列がなければ追加
    var headerColCount = formSheet.getLastColumn();
    if (headerColCount < COL_REMINDER_FLAG + 1) {
      formSheet.getRange(1, COL_REMINDER_FLAG + 1).setValue('再通知済み');
      formSheet.getRange(1, COL_REMINDER_FLAG + 1).setFontWeight('bold');
    }
    
    // 全データを取得（再通知フラグ列含む）
    var lastCol = Math.max(formSheet.getLastColumn(), COL_REMINDER_FLAG + 1);
    var data = formSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    // 募集完了ステータスを取得
    var closedStatuses = {};
    try {
      var mgmtSheet = getOrCreateManagementSheet();
      var mgmtLastRow = mgmtSheet.getLastRow();
      if (mgmtLastRow >= 2) {
        var mgmtData = mgmtSheet.getRange(2, 1, mgmtLastRow - 1, 3).getValues();
        for (var m = 0; m < mgmtData.length; m++) {
          if (mgmtData[m][2] === true || mgmtData[m][2] === 'TRUE') {
            closedStatuses[mgmtData[m][0]] = true;
          }
        }
      }
    } catch (e) {
      Logger.log('リマインド: ステータス取得エラー: ' + e.toString());
    }
    
    // エリアマスタ取得
    var areaMap = getAreaMap();
    
    // LINE友だちシート取得
    var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    var friendSheet = ss.getSheetByName(LINE_CONFIG.ADMIN_SHEET_NAME);
    if (!friendSheet || friendSheet.getLastRow() < 2) {
      Logger.log('リマインド: LINE友だちシートにデータがありません');
      return;
    }
    var friendLastCol = Math.max(friendSheet.getLastColumn(), 9);
    var friendData = friendSheet.getRange(2, 1, friendSheet.getLastRow() - 1, friendLastCol).getValues();
    
    var now = new Date();
    var sevenDaysMs = 7 * 24 * 60 * 60 * 1000;
    var reminderTotal = 0;
    var rowsProcessed = 0;
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      
      // 再通知済みフラグチェック（何か値が入っていたらスキップ）
      var reminderFlag = (row[COL_REMINDER_FLAG] || '').toString().trim();
      if (reminderFlag) continue;
      
      // タイムスタンプから7日経過チェック
      var timestamp = row[COL.TIMESTAMP];
      if (!timestamp) continue;
      var submitDate = new Date(timestamp);
      if (isNaN(submitDate.getTime())) continue;
      
      var elapsed = now.getTime() - submitDate.getTime();
      if (elapsed < sevenDaysMs) continue;
      
      // イベントデータをパース（既存のparseFormResponseを再利用）
      var eventData = parseFormResponse(row);
      if (!eventData || eventData.days.length === 0) continue;
      
      // 各日付について募集完了チェック
      var allClosed = true;
      for (var d = 0; d < eventData.days.length; d++) {
        var dayDate = eventData.days[d].date;
        var eventKey = Utilities.formatDate(dayDate, 'Asia/Tokyo', 'yyyy-MM-dd') + '_' + eventData.hall;
        if (!closedStatuses[eventKey]) {
          allClosed = false;
          break;
        }
      }
      if (allClosed) {
        formSheet.getRange(i + 2, COL_REMINDER_FLAG + 1).setValue('完了済');
        Logger.log('リマインド: 全日程募集完了済みスキップ: ' + eventData.hall);
        continue;
      }
      
      // 募集セクションを特定（重複除去）
      var recruitSections = [];
      eventData.days.forEach(function(day) {
        if (day.stage > 0 && recruitSections.indexOf('舞台') < 0) recruitSections.push('舞台');
        if (day.sound > 0 && recruitSections.indexOf('音響') < 0) recruitSections.push('音響');
        if (day.lighting > 0 && recruitSections.indexOf('照明') < 0) recruitSections.push('照明');
      });
      if (recruitSections.length === 0) continue;
      
      // エリア判定
      var hallArea = areaMap[eventData.hall] || '';
      Logger.log('リマインド対象: ' + eventData.hall + ' (セクション: ' + recruitSections.join(',') + ', エリア: ' + (hallArea || '未設定') + ')');
      
      // 再通知メッセージを生成（パープルテーマ）
      var message = createReminderMessage(eventData, recruitSections);
      
      // セクション＋エリアフィルタリングして通知
      var notifiedCount = 0;
      for (var k = 0; k < friendData.length; k++) {
        var friendUid = (friendData[k][0] || '').toString().trim();
        var friendSection = (friendData[k][5] || '').toString().trim();
        var friendArea = (friendData[k][8] || '').toString().trim();
        
        if (!friendUid || !friendSection) continue;
        
        // セクション一致チェック
        var sectionMatch = false;
        for (var s = 0; s < recruitSections.length; s++) {
          if (friendSection.indexOf(recruitSections[s]) >= 0) {
            sectionMatch = true;
            break;
          }
        }
        if (!sectionMatch) continue;
        
        // エリア判定
        if (hallArea && !shouldNotifyByArea(hallArea, friendArea)) continue;
        
        try {
          UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
            method: 'post',
            headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_CONFIG.CHANNEL_ACCESS_TOKEN },
            payload: JSON.stringify({ to: friendUid, messages: [message] }),
            muteHttpExceptions: true
          });
          notifiedCount++;
        } catch (pushErr) {
          Logger.log('リマインド送信エラー: ' + pushErr.toString());
        }
      }
      
      // 再通知済みフラグを書き込み（「済」= 通知済み）
      formSheet.getRange(i + 2, COL_REMINDER_FLAG + 1).setValue('済');
      reminderTotal += notifiedCount;
      rowsProcessed++;
      Logger.log('リマインド送信: ' + eventData.hall + ' → ' + notifiedCount + '名');
    }
    
    Logger.log('リマインド処理完了: ' + rowsProcessed + '件処理, 合計' + reminderTotal + '名に送信');
  } catch (error) {
    Logger.log('checkAndSendReminders エラー: ' + error.toString());
  }
}

/**
 * 再通知用Flexメッセージを生成（パープルテーマ）
 * 既存のcreateNewRecruitmentMessageと同じ構造で、ヘッダーとフッターをカスタマイズ
 */
function createReminderMessage(eventData, sections) {
  // パープルテーマカラー（再通知専用）
  var reminderColor = '#8b5cf6';
  var reminderDark = '#3b1764';
  var reminderBodyBg = '#110a20';
  var reminderCardBg = '#1e1538';
  var reminderInfoTitle = '#c4b5fd';
  var reminderLabel = '#8878a8';
  var reminderFooterText = '#ddd6fe';
  
  // Web App URL（カレンダーへのリンク）
  var firstDate = eventData.days[0].date;
  var formattedDate = firstDate.getFullYear() + '-' + ('0' + (firstDate.getMonth()+1)).slice(-2) + '-' + ('0' + firstDate.getDate()).slice(-2);
  var calendarUrl = 'https://liff.line.me/2009354296-Ae3FpVy2?date=' + formattedDate + '&hall=' + encodeURIComponent(eventData.hall || '');
  
  // 日程/区分テキスト
  var dateWithSlotTexts = eventData.days.map(function(day) {
    var d = day.date;
    var weekdays = ['日', '月', '火', '水', '木', '金', '土'];
    var dateStr = (d.getMonth() + 1) + '月' + d.getDate() + '日(' + weekdays[d.getDay()] + ')';
    var slot = day.timeSlot || '';
    return slot ? dateStr + ' ' + slot : dateStr;
  });
  var dateSlotDisplay = dateWithSlotTexts.join('\n');
  
  // セクション別人数集計
  var stageTotal = 0, soundTotal = 0, lightingTotal = 0;
  eventData.days.forEach(function(day) {
    if (day.stage > 0) stageTotal = Math.max(stageTotal, day.stage);
    if (day.sound > 0) soundTotal = Math.max(soundTotal, day.sound);
    if (day.lighting > 0) lightingTotal = Math.max(lightingTotal, day.lighting);
  });
  
  // セクションラベル
  var activeSections = [];
  if (stageTotal > 0) activeSections.push('舞台');
  if (soundTotal > 0) activeSections.push('音響');
  if (lightingTotal > 0) activeSections.push('照明');
  var sectionLabel = activeSections.join('・') + '増員';
  
  // 催事情報カード
  var eventInfoContents = [
    { type: 'box', layout: 'horizontal', contents: [
      { type: 'text', text: '催事名', size: 'sm', color: reminderLabel, flex: 1, align: 'start' },
      { type: 'text', text: eventData.eventName || '-', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
    ], margin: 'lg' },
    { type: 'box', layout: 'horizontal', contents: [
      { type: 'text', text: '日程/区分', size: 'sm', color: reminderLabel, flex: 1, align: 'start' },
      { type: 'text', text: dateSlotDisplay, size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
    ], margin: 'md' }
  ];
  
  if (eventData.juniorOk === 'はい') {
    eventInfoContents.push({ type: 'box', layout: 'horizontal', contents: [
      { type: 'text', text: '若手', size: 'sm', color: reminderLabel, flex: 1, align: 'start' },
      { type: 'text', text: '🔰 OK', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start' }
    ], margin: 'md' });
  }
  
  var pref = eventData.consecutivePreference;
  if (eventData.days && eventData.days.length > 1 && pref) {
    var prefText = pref;
    if (pref === 'はい' || pref === 'あり' || pref === '希望する') prefText = '希望あり';
    else if (pref === 'なし' || pref === 'いいえ') prefText = '希望なし';
    else if (pref === '可能であれば') prefText = '可能であれば';
    eventInfoContents.push({ type: 'box', layout: 'horizontal', contents: [
      { type: 'text', text: '通し希望', size: 'sm', color: reminderLabel, flex: 1, align: 'start' },
      { type: 'text', text: prefText, size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
    ], margin: 'md' });
  }
  
  // セクションカード
  var sectionCards = [];
  if (stageTotal > 0) {
    sectionCards.push({
      type: 'box', layout: 'horizontal',
      contents: [
        { type: 'box', layout: 'vertical', contents: [{ type: 'filler' }], width: '4px', backgroundColor: '#10b981' },
        { type: 'box', layout: 'horizontal', contents: [
          { type: 'text', text: '舞台', size: 'sm', color: '#ffffff', weight: 'bold', flex: 1 },
          { type: 'text', text: stageTotal + '名', size: 'sm', color: '#ffffff', weight: 'bold', flex: 0, align: 'end' }
        ], paddingAll: '12px', flex: 1 }
      ],
      backgroundColor: '#0a2618', cornerRadius: 'md', margin: 'sm'
    });
  }
  if (soundTotal > 0) {
    sectionCards.push({
      type: 'box', layout: 'horizontal',
      contents: [
        { type: 'box', layout: 'vertical', contents: [{ type: 'filler' }], width: '4px', backgroundColor: '#0ea5e9' },
        { type: 'box', layout: 'horizontal', contents: [
          { type: 'text', text: '音響', size: 'sm', color: '#ffffff', weight: 'bold', flex: 1 },
          { type: 'text', text: soundTotal + '名', size: 'sm', color: '#ffffff', weight: 'bold', flex: 0, align: 'end' }
        ], paddingAll: '12px', flex: 1 }
      ],
      backgroundColor: '#0a1a28', cornerRadius: 'md', margin: 'sm'
    });
  }
  if (lightingTotal > 0) {
    sectionCards.push({
      type: 'box', layout: 'horizontal',
      contents: [
        { type: 'box', layout: 'vertical', contents: [{ type: 'filler' }], width: '4px', backgroundColor: '#eab308' },
        { type: 'box', layout: 'horizontal', contents: [
          { type: 'text', text: '照明', size: 'sm', color: '#ffffff', weight: 'bold', flex: 1 },
          { type: 'text', text: lightingTotal + '名', size: 'sm', color: '#ffffff', weight: 'bold', flex: 0, align: 'end' }
        ], paddingAll: '12px', flex: 1 }
      ],
      backgroundColor: '#1a1808', cornerRadius: 'md', margin: 'sm'
    });
  }
  
  // ボディ組み立て
  var bodyContents = [
    { type: 'text', text: sectionLabel, size: 'xs', color: '#a78bfa', align: 'center', margin: 'none' },
    { type: 'text', text: eventData.hall || '未定', size: 'xl', color: '#ffffff', weight: 'bold', align: 'center', margin: 'sm' },
    { type: 'text', text: '催事情報', weight: 'bold', size: 'md', color: reminderInfoTitle, align: 'center', margin: 'lg' },
    { type: 'box', layout: 'vertical', contents: eventInfoContents, backgroundColor: reminderCardBg, cornerRadius: 'md', paddingAll: '14px', margin: 'sm' },
    { type: 'text', text: '募集セクション', weight: 'bold', size: 'md', color: reminderInfoTitle, align: 'center', margin: 'lg' }
  ];
  sectionCards.forEach(function(card) { bodyContents.push(card); });
  bodyContents.push({
    type: 'button',
    action: { type: 'uri', label: '🔍 詳細確認＆応募はこちらから', uri: calendarUrl },
    style: 'primary',
    color: reminderColor,
    margin: 'xl',
    height: 'md'
  });
  
  return {
    type: 'flex',
    altText: '🚨 増員募集再通知 - ' + (eventData.hall || '') + ' ' + sectionLabel,
    contents: {
      type: 'bubble',
      styles: {
        header: { backgroundColor: reminderDark },
        body: { backgroundColor: reminderBodyBg },
        footer: { backgroundColor: reminderDark }
      },
      header: {
        type: 'box', layout: 'vertical',
        contents: [
          {
            type: 'box', layout: 'vertical',
            contents: [{ type: 'text', text: '🚨 増員募集再通知 🚨', weight: 'bold', size: 'lg', color: '#ffffff', align: 'center' }],
            backgroundColor: reminderColor, cornerRadius: 'md', paddingAll: '16px'
          }
        ],
        backgroundColor: reminderDark, paddingAll: '16px'
      },
      body: {
        type: 'box', layout: 'vertical',
        contents: bodyContents,
        backgroundColor: reminderBodyBg, paddingAll: '18px'
      },
      footer: {
        type: 'box', layout: 'vertical',
        contents: [
          { type: 'separator', color: '#5b3a94' },
          { type: 'text', text: '増員がまだ見つかってません。', size: 'xs', color: reminderFooterText, wrap: true, align: 'center', margin: 'lg' },
          { type: 'text', text: 'ご協力よろしくお願いいたします。🙏', size: 'xs', color: reminderFooterText, wrap: true, align: 'center', margin: 'sm' }
        ],
        backgroundColor: reminderDark, paddingAll: '15px'
      }
    }
  };
}

/**
 * リマインド用の日次トリガーを設定（毎日午前10時）
 * GASエディタから手動で1回だけ実行してください
 * ※旧sendRecruitmentRemindersのトリガーも自動削除します
 */
function setupDailyReminderTrigger() {
  // 既存の関連トリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var handler = triggers[i].getHandlerFunction();
    if (handler === 'checkAndSendReminders' || handler === 'sendRecruitmentReminders' || handler === '_OLD_sendRecruitmentReminders') {
      ScriptApp.deleteTrigger(triggers[i]);
      Logger.log('トリガー削除: ' + handler);
    }
  }
  
  // 新規トリガー: 毎日午前10時
  ScriptApp.newTrigger('checkAndSendReminders')
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .create();
  
  Logger.log('リマインドトリガーを設定しました: 毎日午前10時にcheckAndSendRemindersを実行');
}

// ============================================
// システム管理者通知（ホール情報登録/更新時）
// J列（10列目）チェック付きユーザーへLINE Push通知
// ============================================

/**
 * 必要事項の回答シートのJ列がTRUEのシステム管理者のLINE UIDを取得
 * @return {Array} [{ uid, staffName, email }] の配列
 */
function getSystemAdminLineUids() {
  var result = [];
  try {
    var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
    var staffSheet = ss.getSheetByName(STAFF_CONFIG.SHEET_NAME);
    if (!staffSheet) {
      Logger.log('getSystemAdminLineUids: スタッフシートが見つかりません');
      return result;
    }
    
    var lastRow = staffSheet.getLastRow();
    if (lastRow < 2) return result;
    
    // A〜J列（10列）を取得
    var data = staffSheet.getRange(2, 1, lastRow - 1, 10).getValues();
    
    for (var i = 0; i < data.length; i++) {
      var isSystemAdmin = data[i][9]; // J列（10列目、インデックス9）
      if (isSystemAdmin !== true && isSystemAdmin !== 'TRUE') continue;
      
      var staffName = (data[i][1] || '').toString().trim();  // B列: 名前
      var staffEmail = (data[i][5] || '').toString().trim();  // F列: メール
      
      if (!staffName && !staffEmail) continue;
      
      // LINE UIDを検索（メール → 名前 の優先順）
      var uid = null;
      if (staffEmail) {
        uid = getLineUidByEmail(staffEmail);
      }
      if (!uid && staffName) {
        uid = getLineUidByName(staffName);
      }
      
      if (uid) {
        result.push({ uid: uid, staffName: staffName, email: staffEmail });
      } else {
        Logger.log('システム管理者のLINE UID未登録: ' + staffName + ' (' + staffEmail + ')');
      }
    }
    
    Logger.log('システム管理者（UID取得済み）: ' + result.length + '名');
  } catch (error) {
    Logger.log('getSystemAdminLineUids エラー: ' + error.toString());
  }
  return result;
}

/**
 * システム管理者全員にホール情報登録/更新の通知をPush送信
 * @param {string} hallName - 登録/更新されたホール名
 * @param {object} responses - フォーム回答内容 { hallName, address, phone, email, adminName, nearby1, nearby2, smoking }
 */
function notifySystemAdminsHallInfoUpdate(hallName, responses) {
  try {
    var admins = getSystemAdminLineUids();
    if (admins.length === 0) {
      Logger.log('通知対象のシステム管理者がいません');
      return;
    }
    
    var message = buildHallInfoUpdateFlexMessage(hallName, responses);
    var sentCount = 0;
    
    for (var i = 0; i < admins.length; i++) {
      try {
        var response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
          method: 'post',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + LINE_CONFIG.CHANNEL_ACCESS_TOKEN
          },
          payload: JSON.stringify({
            to: admins[i].uid,
            messages: [message]
          }),
          muteHttpExceptions: true
        });
        
        var code = response.getResponseCode();
        if (code === 200) {
          sentCount++;
          Logger.log('システム管理者通知送信: ' + admins[i].staffName);
        } else {
          Logger.log('システム管理者通知失敗: ' + admins[i].staffName + ' / HTTP ' + code + ' / ' + response.getContentText());
        }
      } catch (pushError) {
        Logger.log('Push送信エラー(' + admins[i].staffName + '): ' + pushError.toString());
      }
    }
    
    Logger.log('ホール情報更新通知完了: ' + sentCount + '/' + admins.length + '名に送信');
  } catch (error) {
    Logger.log('notifySystemAdminsHallInfoUpdate エラー: ' + error.toString());
  }
}

/**
 * ホール情報登録/更新通知用のFlexメッセージを生成
 * @param {string} hallName - ホール名
 * @param {object} responses - フォーム回答内容
 * @return {object} LINE Flex Message
 */
function buildHallInfoUpdateFlexMessage(hallName, responses) {
  var resp = responses || {};
  var now = new Date();
  var timestamp = (now.getMonth() + 1) + '/' + now.getDate() + ' ' +
    ('0' + now.getHours()).slice(-2) + ':' + ('0' + now.getMinutes()).slice(-2);
  
  // 登録内容のサマリーを作成
  var detailItems = [];
  
  // ホール名
  detailItems.push({
    type: 'box', layout: 'horizontal',
    contents: [
      { type: 'text', text: 'ホール名', size: 'sm', color: '#7a8a9a', flex: 1, align: 'start' },
      { type: 'text', text: hallName || '-', size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
    ],
    margin: 'lg'
  });
  
  // 住所
  if (resp.address) {
    detailItems.push({
      type: 'box', layout: 'horizontal',
      contents: [
        { type: 'text', text: '住所', size: 'sm', color: '#7a8a9a', flex: 1, align: 'start' },
        { type: 'text', text: resp.address, size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
      ],
      margin: 'md'
    });
  }
  
  // 電話番号
  if (resp.phone) {
    detailItems.push({
      type: 'box', layout: 'horizontal',
      contents: [
        { type: 'text', text: '電話番号', size: 'sm', color: '#7a8a9a', flex: 1, align: 'start' },
        { type: 'text', text: resp.phone, size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
      ],
      margin: 'md'
    });
  }
  
  // 担当者名
  if (resp.adminName) {
    detailItems.push({
      type: 'box', layout: 'horizontal',
      contents: [
        { type: 'text', text: '担当者', size: 'sm', color: '#7a8a9a', flex: 1, align: 'start' },
        { type: 'text', text: resp.adminName, size: 'sm', color: '#ffffff', flex: 2, weight: 'bold', align: 'start', wrap: true }
      ],
      margin: 'md'
    });
  }
  
  // 周辺情報①
  if (resp.nearby1) {
    var nearby1Text = resp.nearby1.length > 40 ? resp.nearby1.substring(0, 40) + '...' : resp.nearby1;
    detailItems.push({
      type: 'box', layout: 'horizontal',
      contents: [
        { type: 'text', text: '周辺情報', size: 'sm', color: '#7a8a9a', flex: 1, align: 'start' },
        { type: 'text', text: nearby1Text, size: 'sm', color: '#ffffff', flex: 2, align: 'start', wrap: true }
      ],
      margin: 'md'
    });
  }
  
  // 喫煙所
  if (resp.smoking) {
    var smokingText = resp.smoking.length > 40 ? resp.smoking.substring(0, 40) + '...' : resp.smoking;
    detailItems.push({
      type: 'box', layout: 'horizontal',
      contents: [
        { type: 'text', text: '喫煙所', size: 'sm', color: '#7a8a9a', flex: 1, align: 'start' },
        { type: 'text', text: smokingText, size: 'sm', color: '#ffffff', flex: 2, align: 'start', wrap: true }
      ],
      margin: 'md'
    });
  }
  
  return {
    type: 'flex',
    altText: '🏛 ホール情報更新通知 - ' + (hallName || ''),
    contents: {
      type: 'bubble',
      styles: {
        header: { backgroundColor: '#1a3a5c' },
        body: { backgroundColor: '#0a1628' },
        footer: { backgroundColor: '#0a1628' }
      },
      header: {
        type: 'box', layout: 'vertical',
        contents: [
          {
            type: 'box', layout: 'vertical',
            contents: [
              { type: 'text', text: '🏛 ホール情報更新通知', weight: 'bold', size: 'lg', color: '#ffffff', align: 'center' }
            ],
            backgroundColor: '#2e86c1', cornerRadius: 'md', paddingAll: '16px'
          }
        ],
        backgroundColor: '#1a3a5c', paddingAll: '16px'
      },
      body: {
        type: 'box', layout: 'vertical',
        contents: [
          { type: 'text', text: 'ホール情報が登録/更新されました', weight: 'bold', size: 'sm', color: '#7dd3fc', align: 'center', margin: 'sm' },
          {
            type: 'box', layout: 'vertical',
            contents: detailItems,
            backgroundColor: '#142238', cornerRadius: 'md', paddingAll: '14px', margin: 'md'
          }
        ],
        backgroundColor: '#0a1628', paddingAll: '18px'
      },
      footer: {
        type: 'box', layout: 'vertical',
        contents: [
          { type: 'separator', color: '#1e3a5f' },
          { type: 'text', text: '🔧 システム管理者通知', size: 'xs', color: '#7a8a9a', align: 'center', margin: 'lg' },
          { type: 'text', text: timestamp, size: 'xs', color: '#5a6a7a', align: 'center', margin: 'sm' }
        ],
        backgroundColor: '#0a1628', paddingAll: '15px'
      }
    }
  };
}

