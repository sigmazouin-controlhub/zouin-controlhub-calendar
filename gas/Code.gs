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
  SHEET_NAME: 'フォームの回答 1',
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
    if (params.action === 'toggleRecruitment') return createResponse(toggleRecruitmentStatus(params.eventKey, params.hall, params.status, params.email), e);
    if (params.action === 'getRecruitmentStatuses') return createResponse(getRecruitmentStatuses(), e);
    
    // カレンダーイベント取得
    const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    if (!calendar) return createResponse({ success: false, error: 'カレンダーが見つかりません' }, e);
    
    const startDate = params.start ? new Date(params.start) : new Date();
    const endDate = params.end ? new Date(params.end) : new Date(startDate.getTime() + 90 * 24 * 60 * 60 * 1000);
    const events = calendar.getEvents(startDate, endDate);
    
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
          timeSlot: parsedDescription.timeSlot, content: parsedDescription.content
        }
      };
    });
    
    return createResponse({ success: true, events: eventList, totalCount: eventList.length }, e);
  } catch (error) {
    return createResponse({ success: false, error: error.toString() }, e);
  }
}

function parseDescription(description) {
  const result = { juniorOk: false, consecutivePreference: '', timeSlot: '', content: '', sections: { stage: 0, sound: 0, lighting: 0 } };
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
    const sheet = ss.getSheetByName(STAFF_CONFIG.SHEET_NAME);
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
      sheet.getRange(1, 1, 1, 5).setValues([['LINE UID', 'LINE表示名', 'スタッフ名', '種別', '登録日時']]);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    }
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === lineUid) {
        sheet.getRange(i + 1, 2).setValue(displayName);
        if (staffName) sheet.getRange(i + 1, 3).setValue(staffName);
        return { success: true, message: '更新しました', isNew: false };
      }
    }
    
    sheet.appendRow([lineUid, displayName, staffName, '', new Date()]);
    return { success: true, message: '登録しました', isNew: true };
  } catch (error) {
    return { success: false, error: error.toString() };
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
      sheet.getRange(1, 1, 1, 9).setValues([['申込日時', 'メールアドレス', 'スタッフ名', '所属ホール', '所属セクション', 'イベント', '会場ホール', 'セクション', '希望日程']]);
      sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
    }
    
    sheet.appendRow([new Date(), email, staffName, staffHall, staffSection, eventTitle, hall, section, selectedDates]);
    updateStaffProfile(email, staffHall, staffSection);
    
    if (sendLineNotification && email) {
      sendLineNotificationToStaff(email, { date: selectedDates, hall, section, staffName });
    }
    
    try {
      notifyAdminsByHall({ date: selectedDates, hall, section: staffSection, staffName, eventTitle });
    } catch (adminError) {
      Logger.log('管理者通知エラー: ' + adminError.toString());
    }
    
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
    var sheet = ss.getSheetByName(STAFF_CONFIG.SHEET_NAME);
    if (!sheet) return { success: false, error: 'シートが見つかりません' };
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, found: false };
    var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    
    for (var i = 0; i < data.length; i++) {
      if (normalizeName(data[i][5]) === normalizedEmail) {
        return { success: true, found: true, staffName: data[i][1] || '', staffHall: data[i][3] || '', staffSection: data[i][4] || '', isAdmin: data[i][8] === true || data[i][8] === 'TRUE' };
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

function toggleRecruitmentStatus(eventKey, hall, status, email) {
  try {
    if (!eventKey || !hall || !email) return { success: false, error: 'パラメータ不足' };
    var staffInfo = getStaffInfoFromMainSheet(email);
    if (!staffInfo.found || !staffInfo.isAdmin) return { success: false, error: '管理者権限がありません' };
    if (staffInfo.staffHall !== hall) return { success: false, error: '自分のホールのみ操作可能' };
    
    var sheet = getOrCreateManagementSheet();
    var lastRow = sheet.getLastRow();
    var isComplete = (status === 'true' || status === true);
    
    if (lastRow >= 2) {
      var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
      for (var i = 0; i < data.length; i++) {
        if (data[i][0] === eventKey) {
          sheet.getRange(i + 2, 3).setValue(isComplete);
          sheet.getRange(i + 2, 4).setValue(new Date());
          return { success: true, eventKey: eventKey, closed: isComplete };
        }
      }
    }
    sheet.appendRow([eventKey, hall, isComplete, new Date()]);
    return { success: true, eventKey: eventKey, closed: isComplete };
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
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][3] === email) return data[i][0];
    }
    return null;
  } catch (error) {
    return null;
  }
}

function sendLineNotificationToStaff(email, eventInfo) {
  try {
    const lineUid = getLineUidByEmail(email);
    if (!lineUid) return false;
    
    const message = {
      type: 'flex', altText: '申し込み完了のお知らせ',
      contents: {
        type: 'bubble',
        header: { type: 'box', layout: 'vertical', contents: [{ type: 'text', text: '✅ 申し込み完了', weight: 'bold', size: 'lg', color: '#ffffff' }], backgroundColor: '#4ade80', paddingAll: '15px' },
        body: { type: 'box', layout: 'vertical', contents: [
          { type: 'box', layout: 'horizontal', contents: [{ type: 'text', text: '日付', size: 'sm', color: '#666666', flex: 2 }, { type: 'text', text: eventInfo.date || '-', size: 'sm', color: '#111111', flex: 5, wrap: true }], margin: 'md' },
          { type: 'box', layout: 'horizontal', contents: [{ type: 'text', text: 'ホール', size: 'sm', color: '#666666', flex: 2 }, { type: 'text', text: eventInfo.hall || '-', size: 'sm', color: '#111111', flex: 5, wrap: true }], margin: 'md' },
          { type: 'box', layout: 'horizontal', contents: [{ type: 'text', text: 'セクション', size: 'sm', color: '#666666', flex: 2 }, { type: 'text', text: eventInfo.section || '-', size: 'sm', color: '#111111', flex: 5, wrap: true }], margin: 'md' },
          { type: 'box', layout: 'horizontal', contents: [{ type: 'text', text: '申込者', size: 'sm', color: '#666666', flex: 2 }, { type: 'text', text: eventInfo.staffName || '-', size: 'sm', color: '#111111', flex: 5, wrap: true }], margin: 'md' },
          { type: 'separator', margin: 'lg' },
          { type: 'text', text: '担当者からの連絡までしばらくお待ちください。', size: 'sm', color: '#666666', margin: 'lg', wrap: true }
        ], paddingAll: '15px' }
      }
    };
    
    const options = {
      method: 'post',
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_CONFIG.CHANNEL_ACCESS_TOKEN },
      payload: JSON.stringify({ to: lineUid, messages: [message] }),
      muteHttpExceptions: true
    };
    return UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', options).getResponseCode() === 200;
  } catch (error) {
    return false;
  }
}

// ============================================
// LINE Webhook受信（doPost）
// ============================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const events = data.events || [];
    
    for (const event of events) {
      if (event.type !== 'message' || event.message.type !== 'text') continue;
      const messageText = event.message.text;
      
      if (messageText.includes('申し込みます')) {
        const nameMatch = messageText.match(/お名前[］\]]\s*[：:]\s*(.+)/);
        const staffName = nameMatch ? nameMatch[1].trim() : 'スタッフ';
        replyMessage(event.replyToken, `${staffName}さん\n増員応募ありがとうございます！\n\n確認出来次第、担当者からご連絡しますので、しばらくお待ちください。\n\nまた、数日経ってもご連絡が無い場合は、お手数ですが増員先の事業所まで直接ご連絡ください。`);
      }
    }
  } catch (error) {
    Logger.log('Webhook処理エラー: ' + error.toString());
  }
  return ContentService.createTextOutput(JSON.stringify({status: 'ok'})).setMimeType(ContentService.MimeType.JSON);
}

function replyMessage(replyToken, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_CONFIG.CHANNEL_ACCESS_TOKEN },
    payload: JSON.stringify({ replyToken, messages: [{ type: 'text', text }] }),
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
  const dates = (eventInfo.date || '').split(',').map(function(d) {
    d = d.trim();
    const match = d.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    return match ? parseInt(match[2]) + '月' + parseInt(match[3]) + '日' : d;
  });
  return {
    type: 'text',
    text: adminName + 'さん\n' + dates.join('・') + '（' + (eventInfo.eventTitle || '') + '）\n' + (eventInfo.section || '') + '　' + (eventInfo.staffName || '') + 'さんから\n増員の申し込みがありました。\n公式LINEアプリのチャットからご対応をお願いします！'
  };
}
