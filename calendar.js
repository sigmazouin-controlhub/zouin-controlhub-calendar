/**
 * 増員 CONTROL HUB 2.0 - カレンダーロジック
 */

// ============================================
// GAS API設定
// ============================================
const API_CONFIG = {
    // GAS WebアプリURL
    GAS_URL: 'https://script.google.com/macros/s/AKfycbxbV2AK-9edfJDLYd5roY5Lj3mcWDfLuVXyrLbkbmj-f0ghreoTpBLLHeze72BAMz6h/exec',

    // テストモード（falseでAPIからデータを取得）
    USE_SAMPLE_DATA: false
};

// 区分パターン
const timeSlots = {
    morning: '午前',
    afternoon: '午後',
    evening: '夜間',
    allday: '全日',
    morningAfternoon: '午前午後',
    afternoonEvening: '午後夜間'
};

// イベントデータ（APIから取得、または初期化時に空オブジェクト）
let eventsData = {};

// サンプルイベントデータ（テストモード用 - 現在は無効）
const sampleEvents = {};

// セクション情報
const sectionInfo = {
    stage: { name: '舞台', icon: '🎭', color: 'stage' },
    sound: { name: '音響', icon: '🎵', color: 'sound' },
    lighting: { name: '照明', icon: '💡', color: 'lighting' },
    multiple: { name: '複数', icon: '🎪', color: 'multiple' }
};

// カレンダー状態
let currentDate = new Date();
let selectedDate = null;
let selectedEvents = [];
let isLoading = false;

// DOM要素
const calendarGrid = document.getElementById('calendarGrid');
const currentMonthEl = document.getElementById('currentMonth');
const prevMonthBtn = document.getElementById('prevMonth');
const nextMonthBtn = document.getElementById('nextMonth');
const todayBtn = document.getElementById('todayBtn');

/**
 * カレンダーを描画
 */
function renderCalendar() {
    const year = currentDate.getFullYear();
    const month = currentDate.getMonth();

    // 月表示を更新
    currentMonthEl.textContent = `${year}年${month + 1}月`;

    // カレンダーグリッドをクリア
    calendarGrid.innerHTML = '';

    // 月の最初と最後の日
    const firstDay = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0);

    // 月の最初の曜日（0=日曜）
    const startDayOfWeek = firstDay.getDay();

    // 今日の日付
    const today = new Date();
    const todayStr = formatDate(today);

    // 前月の日を追加
    const prevMonthLastDay = new Date(year, month, 0).getDate();
    for (let i = startDayOfWeek - 1; i >= 0; i--) {
        const day = prevMonthLastDay - i;
        const date = new Date(year, month - 1, day);
        const dayEl = createDayElement(date, true);
        calendarGrid.appendChild(dayEl);
    }

    // 今月の日を追加
    for (let day = 1; day <= lastDay.getDate(); day++) {
        const date = new Date(year, month, day);
        const dateStr = formatDate(date);
        const isToday = dateStr === todayStr;
        const dayEl = createDayElement(date, false, isToday);
        calendarGrid.appendChild(dayEl);
    }

    // 次月の日を追加（6週間分になるように）
    const totalDays = startDayOfWeek + lastDay.getDate();
    const remainingDays = 42 - totalDays; // 6週間 = 42日
    for (let day = 1; day <= remainingDays; day++) {
        const date = new Date(year, month + 1, day);
        const dayEl = createDayElement(date, true);
        calendarGrid.appendChild(dayEl);
    }
}

/**
 * 日付セルを作成
 */
function createDayElement(date, isOtherMonth, isToday = false) {
    const dayEl = document.createElement('div');
    dayEl.className = 'calendar-day';

    if (isOtherMonth) {
        dayEl.classList.add('other-month');
    }

    if (isToday) {
        dayEl.classList.add('today');
    }

    // 曜日クラス
    const dayOfWeek = date.getDay();
    if (dayOfWeek === 0) dayEl.classList.add('sunday');
    if (dayOfWeek === 6) dayEl.classList.add('saturday');

    // 日付番号
    const dayNumber = document.createElement('div');
    dayNumber.className = 'day-number';
    dayNumber.textContent = date.getDate();
    dayEl.appendChild(dayNumber);

    // イベントバッジを追加
    const dateStr = formatDate(date);
    const events = eventsData[dateStr];

    if (events && events.length > 0) {
        const badgesContainer = document.createElement('div');
        badgesContainer.className = 'section-badges';

        // 連日イベントと単日イベントを分離
        const multiDayEvents = events.filter(e => e.startDate !== e.endDate);
        const singleDayEvents = events.filter(e => e.startDate === e.endDate);

        // イベントが1件のみの場合は通常表示
        if (events.length === 1) {
            const badge = createBadge(events[0], date);
            badgesContainer.appendChild(badge);
        } else {
            // 複数イベントの場合: 連日イベントを優先表示
            multiDayEvents.forEach(event => {
                const badge = createBadge(event, date);
                badgesContainer.appendChild(badge);
            });

            // 単日イベントの表示ロジック
            if (singleDayEvents.length > 0) {
                if (multiDayEvents.length > 0) {
                    // 連日イベントがある場合: 単日イベントは全て「+N件」ドロワーへ
                    const moreBadge = document.createElement('div');
                    moreBadge.className = 'section-badge more-badge';
                    moreBadge.textContent = `+${singleDayEvents.length}件`;
                    moreBadge.addEventListener('click', (e) => {
                        e.stopPropagation();
                        showEventList(date, singleDayEvents);
                    });
                    badgesContainer.appendChild(moreBadge);
                } else {
                    // 連日イベントがない場合: 1件目は通常表示、2件目以降は「+N件」ドロワーへ
                    const firstBadge = createBadge(singleDayEvents[0], date);
                    badgesContainer.appendChild(firstBadge);

                    const remainingEvents = singleDayEvents.slice(1);
                    if (remainingEvents.length > 0) {
                        const moreBadge = document.createElement('div');
                        moreBadge.className = 'section-badge more-badge';
                        moreBadge.textContent = `+${remainingEvents.length}件`;
                        moreBadge.addEventListener('click', (e) => {
                            e.stopPropagation();
                            showEventList(date, remainingEvents);
                        });
                        badgesContainer.appendChild(moreBadge);
                    }
                }
            }
        }

        dayEl.appendChild(badgesContainer);
    }

    return dayEl;
}

/**
 * セクションバッジを作成（クリック可能・連日対応）
 */
function createBadge(event, date) {
    const badge = document.createElement('div');
    badge.className = `section-badge ${event.section}`;

    const section = sectionInfo[event.section];
    const hallName = event.hall || '';

    // 連日イベントまたはグループ化イベントの判定
    const dateStr = formatDate(date);
    const isContiguousMultiDay = event.startDate !== event.endDate;
    const isGroupedEvent = event.groupId !== null && event.groupId !== undefined;
    const isMultiDay = isContiguousMultiDay || isGroupedEvent;

    if (isMultiDay) {
        // 連日バッジ: テキストを日にち跨いで分散表示
        badge.classList.add('multi-day');
        const juniorMark = event.juniorOk ? '🔰' : '';

        if (isContiguousMultiDay) {
            // 従来の連続日ロジック
            if (dateStr === event.startDate) {
                badge.classList.add('multi-start');
                badge.innerHTML = `${juniorMark}${section.name}`;
            } else if (dateStr === event.endDate) {
                badge.classList.add('multi-end');
                badge.innerHTML = hallName;
            } else {
                badge.classList.add('multi-middle');
                badge.innerHTML = 'セクション';
            }
        } else if (isGroupedEvent) {
            // グループ化イベント（飛び飛び）は2行表示
            badge.classList.add('multi-start');
            badge.innerHTML = `${juniorMark}${section.name}<br>${hallName}`;
        }
    } else {
        // 単日バッジ: 2行表示
        const juniorMark = event.juniorOk ? '🔰' : '';
        badge.innerHTML = `${juniorMark}${section.name}<br>${hallName}`;
    }

    // バッジをクリックしたら個別イベントを開く
    badge.addEventListener('click', (e) => {
        e.stopPropagation();
        openDrawer(date, [event]);
    });

    return badge;
}

/**
 * 日付をフォーマット（YYYY-MM-DD）
 */
function formatDate(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

/**
 * 日付を日本語フォーマット
 */
function formatDateJP(date) {
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const weekdays = ['日', '月', '火', '水', '木', '金', '土'];
    const weekday = weekdays[date.getDay()];
    return `${month}月${day}日（${weekday}）`;
}

/**
 * 前月へ
 */
function goToPrevMonth() {
    currentDate.setMonth(currentDate.getMonth() - 1);
    renderCalendar();
}

/**
 * 次月へ
 */
function goToNextMonth() {
    currentDate.setMonth(currentDate.getMonth() + 1);
    renderCalendar();
}

// イベントリスナー
prevMonthBtn.addEventListener('click', goToPrevMonth);
nextMonthBtn.addEventListener('click', goToNextMonth);
todayBtn.addEventListener('click', () => {
    currentDate = new Date();
    renderCalendar();
});

// スワイプ対応
let touchStartX = 0;
let touchEndX = 0;

document.addEventListener('touchstart', (e) => {
    touchStartX = e.changedTouches[0].screenX;
}, { passive: true });

document.addEventListener('touchend', (e) => {
    touchEndX = e.changedTouches[0].screenX;
    handleSwipe();
}, { passive: true });

function handleSwipe() {
    const threshold = 50;
    const diff = touchStartX - touchEndX;

    if (Math.abs(diff) > threshold) {
        if (diff > 0) {
            // 左スワイプ -> 次月
            goToNextMonth();
        } else {
            // 右スワイプ -> 前月
            goToPrevMonth();
        }
    }
}

/**
 * イベント一覧を表示（+N件をクリックした時）
 */
function showEventList(date, events) {
    // イベント選択モーダルを表示
    if (window.showEventModal) {
        window.showEventModal(date, events);
    } else {
        // フォールバック: 最初のイベントでドロワーを開く
        openDrawer(date, events);
    }
}

// ============================================
// GAS API からイベントを取得（fetch優先、失敗時はJSONP）
// ============================================
async function fetchEventsFromAPI() {
    if (API_CONFIG.USE_SAMPLE_DATA || API_CONFIG.GAS_URL === 'YOUR_GAS_WEB_APP_URL_HERE') {
        console.log('テストモード: サンプルデータを使用');
        eventsData = sampleEvents;
        return;
    }

    isLoading = true;
    console.log('GAS APIからイベントを取得中...');

    // まずfetchを試す
    try {
        const response = await fetch(API_CONFIG.GAS_URL, {
            method: 'GET',
            redirect: 'follow'
        });
        const data = await response.json();

        if (data.success && data.events) {
            eventsData = transformApiEvents(data.events);
            console.log('イベント取得成功:', data.totalCount, '件');
            isLoading = false;
            return;
        } else {
            throw new Error(data.error || 'APIエラー');
        }
    } catch (fetchError) {
        console.log('fetch失敗、JSONPを試行:', fetchError.message);
    }

    // fetchが失敗したらJSONPを試す
    return new Promise((resolve) => {
        const callbackName = 'gasCallback_' + Date.now();

        // タイムアウト設定（10秒）
        const timeout = setTimeout(() => {
            console.log('JSONP タイムアウト');
            eventsData = {};
            isLoading = false;
            delete window[callbackName];
            resolve();
        }, 10000);

        window[callbackName] = function (data) {
            clearTimeout(timeout);
            try {
                if (data.success && data.events) {
                    eventsData = transformApiEvents(data.events);
                    console.log('イベント取得成功（JSONP）:', data.totalCount, '件');
                } else {
                    console.error('API エラー:', data.error);
                    eventsData = {};
                }
            } catch (error) {
                console.error('データ変換エラー:', error);
                eventsData = {};
            } finally {
                isLoading = false;
                delete window[callbackName];
                if (script.parentNode) {
                    document.head.removeChild(script);
                }
                resolve();
            }
        };

        const script = document.createElement('script');
        script.src = API_CONFIG.GAS_URL + '?callback=' + callbackName;
        script.onerror = function () {
            clearTimeout(timeout);
            console.error('API 接続エラー（JSONP）');
            eventsData = {};
            isLoading = false;
            delete window[callbackName];
            resolve();
        };
        document.head.appendChild(script);
    });
}

// ============================================
// GAS APIレスポンスをWebアプリ用に変換
// ============================================
function transformApiEvents(apiEvents) {
    const transformed = {};

    // Step 1: グループ化 - タイトルに (1/N) パターンがあるイベントを検出
    const eventGroups = {}; // groupKey -> [events]

    apiEvents.forEach(event => {
        // タイトルから (N/M) パターンを検出
        const groupMatch = event.title.match(/(.+?)\s*\((\d+)\/(\d+)\)$/);
        if (groupMatch) {
            const baseTitle = groupMatch[1];
            const groupKey = baseTitle; // ベースタイトルをグループキーに

            if (!eventGroups[groupKey]) {
                eventGroups[groupKey] = [];
            }
            eventGroups[groupKey].push(event);
        }
    });

    apiEvents.forEach(event => {
        // descriptionからセクション情報をパース
        const description = event.description || '';
        const parsedSections = parseDescriptionSections(description);

        // セクションを判定
        let section = 'multiple';
        const sectionCount = [parsedSections.stage > 0, parsedSections.sound > 0, parsedSections.lighting > 0].filter(Boolean).length;

        if (sectionCount === 0) {
            section = 'stage';
        } else if (sectionCount === 1) {
            if (parsedSections.stage > 0) section = 'stage';
            else if (parsedSections.sound > 0) section = 'sound';
            else if (parsedSections.lighting > 0) section = 'lighting';
        }

        const totalCapacity = parsedSections.stage + parsedSections.sound + parsedSections.lighting;

        const timeSlotMap = {
            '全日': 'allday',
            '午前': 'morning',
            '午後': 'afternoon',
            '夜間': 'evening',
            '午前午後': 'morningAfternoon',
            '午後夜間': 'afternoonEvening'
        };

        const timeSlotRegex = /【利用区分】\n[　\s]*(.+)/g;
        let tsMatch;
        const timeSlotRawTexts = [];
        while ((tsMatch = timeSlotRegex.exec(description)) !== null) {
            timeSlotRawTexts.push(tsMatch[1].trim());
        }
        
        const timeSlotText = timeSlotRawTexts.length > 0 ? timeSlotRawTexts[0] : '';

        // グループ情報を追加
        const groupMatch = event.title.match(/(.+?)\s*\((\d+)\/(\d+)\)$/);
        let groupId = null;
        let relatedDates = [];

        if (groupMatch) {
            const baseTitle = groupMatch[1];
            groupId = baseTitle;

            // 同じグループの全イベントの日付を収集
            if (eventGroups[groupId]) {
                relatedDates = eventGroups[groupId].map(e => e.start).sort();
            }
        }

        const transformedEvent = {
            id: event.id,
            eventName: event.extendedProps?.eventName || '',
            title: event.title,
            section: section,
            capacity: totalCapacity,
            applied: 0,
            timeSlot: timeSlotMap[timeSlotText] || 'allday',
            startDate: event.start,
            endDate: event.end,
            hall: event.extendedProps?.hall || '',
            juniorOk: event.extendedProps?.juniorOk || false,
            description: event.description,
            extendedProps: event.extendedProps,
            color: event.color,
            parsedSections: parsedSections,
            timeSlotRawTexts: timeSlotRawTexts,
            groupId: groupId,              // グループID
            relatedDates: relatedDates,    // 関連日付リスト
            applicantCounts: event.extendedProps?.applicantCounts || { stage: 0, sound: 0, lighting: 0 }
        };

        // 開始日から終了日まで各日に追加
        const start = new Date(event.start);
        const end = new Date(event.end);

        for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
            const dateStr = formatDate(d);
            if (!transformed[dateStr]) {
                transformed[dateStr] = [];
            }
            if (!transformed[dateStr].find(e => e.id === event.id)) {
                transformed[dateStr].push({ ...transformedEvent });
            }
        }
    });

    return transformed;
}

// ============================================
// descriptionから募集セクション情報をパース
// ============================================
function parseDescriptionSections(description) {
    const result = { stage: 0, sound: 0, lighting: 0 };

    if (!description) return result;

    // 募集セクション部分を探す
    const stageMatch = description.match(/舞台[：:]\s*(\d+)人/);
    const soundMatch = description.match(/音響[：:]\s*(\d+)人/);
    const lightingMatch = description.match(/照明[：:]\s*(\d+)人/);

    if (stageMatch) result.stage = parseInt(stageMatch[1], 10);
    if (soundMatch) result.sound = parseInt(soundMatch[1], 10);
    if (lightingMatch) result.lighting = parseInt(lightingMatch[1], 10);

    return result;
}

// ============================================
// 初期化
// ============================================
async function initializeApp() {
    const urlParams = new URLSearchParams(window.location.search);
    let targetDate = urlParams.get('date');
    let targetHall = urlParams.get('hall');

    // LIFF環境ではパラメータが liff.state に格納される場合がある
    const liffState = urlParams.get('liff.state');
    if (liffState) {
        const stateParams = new URLSearchParams(liffState.startsWith('?') ? liffState : '?' + liffState);
        if (!targetDate) targetDate = stateParams.get('date');
        if (!targetHall) targetHall = stateParams.get('hall');
    }

    if (targetDate) {
        const parsedDate = new Date(targetDate);
        if (!isNaN(parsedDate.getTime())) {
            currentDate = parsedDate;
        }
    }

    await fetchEventsFromAPI();
    renderCalendar();

    if (targetDate && targetHall && eventsData[targetDate]) {
        const dayEvents = eventsData[targetDate];
        const targetEvents = dayEvents.filter(e => e.hall === targetHall);

        if (targetEvents.length > 0) {
            setTimeout(() => {
                const parts = targetDate.split('-');
                const exactDate = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
                
                if (typeof openDrawer === 'function') {
                    openDrawer(exactDate, [targetEvents[0]]);
                } else if (typeof window.showEventModal === 'function') {
                    window.showEventModal(exactDate, targetEvents);
                }
            }, 300);
        }
    }
}

document.addEventListener('DOMContentLoaded', () => {
    initializeApp();
});

// グローバルに公開（drawer.jsから使用）
window.calendarApp = {
    selectedDate,
    selectedEvents,
    sectionInfo,
    timeSlots,
    formatDateJP,
    API_CONFIG: API_CONFIG,
    refreshCalendar: async () => {
        await fetchEventsFromAPI();
        renderCalendar();
    },
    refreshEvents: async () => {
        await fetchEventsFromAPI();
        renderCalendar();
    }
};
