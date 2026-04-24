/**
 * 増員 CONTROL HUB 2.0 - ドロワーUI制御
 */

// GAS Web App URL
const DRAWER_GAS_URL = 'https://script.google.com/macros/s/AKfycbxbV2AK-9edfJDLYd5roY5Lj3mcWDfLuVXyrLbkbmj-f0ghreoTpBLLHeze72BAMz6h/exec';

// DOM要素
const drawer = document.getElementById('drawer');
const drawerOverlay = document.getElementById('drawerOverlay');
const drawerClose = document.getElementById('drawerClose');
const drawerInner = document.getElementById('drawerInner');
const drawerSectionLabel = document.getElementById('drawerSectionLabel');
const drawerHallName = document.getElementById('drawerHallName');
const drawerEventName = document.getElementById('drawerEventName');
const drawerDate = document.getElementById('drawerDate');
const drawerTime = document.getElementById('drawerTime');
const drawerStatus = document.getElementById('drawerStatus');
const drawerDescription = document.getElementById('drawerDescription');
const drawerNotice = document.getElementById('drawerNotice');
const noticeTitle = document.getElementById('noticeTitle');
const noticeApplicantList = document.getElementById('noticeApplicantList');
const noticeFooter = document.getElementById('noticeFooter');
const applyForm = document.getElementById('applyForm');
const confirmDialog = document.getElementById('confirmDialog');
const confirmClose = document.getElementById('confirmClose');

// イベントモーダル
const eventModalOverlay = document.getElementById('eventModalOverlay');
const eventModal = document.getElementById('eventModal');
const eventModalTitle = document.getElementById('eventModalTitle');
const eventModalList = document.getElementById('eventModalList');
const eventModalClose = document.getElementById('eventModalClose');

// 現在の選択状態
let currentSelectedDate = null;
let currentSelectedEvent = null;
let selectedDates = []; // 選択された日付（連日イベント用）
let recruitmentStatuses = {}; // 募集完了ステータスキャッシュ

// 日付選択関連のDOM要素
const dateSelectionGroup = document.getElementById('dateSelectionGroup');


/**
 * ドロワーを開く（イベントを選択して表示）
 */
function openDrawer(date, events) {
    currentSelectedDate = date;

    // イベントがない場合
    if (!events || events.length === 0) {
        return;
    }

    // 最初のイベントを表示
    const event = events[0];
    currentSelectedEvent = event;

    // セクション情報を取得
    const { sectionInfo, timeSlots, formatDateJP } = window.calendarApp;
    const section = sectionInfo[event.section];

    // ドロワーのセクションクラスを設定
    drawer.className = 'drawer active section-' + event.section;

    // ログイン名を応募フォームに自動セット（B列の名前を使用）
    const savedName = window.loggedInStaffName || localStorage.getItem('zouin_staff_display_name');
    if (savedName) {
        document.getElementById('userName').value = savedName;
    }

    // ヘッダー部分（セクション＋ホール名）
    if (event.section === 'multiple') {
        // 複数セクションの場合: descriptionから選ばれているセクションを抽出
        const selectedSections = [];
        if (event.parsedSections?.stage > 0 || event.description?.includes('舞台:')) {
            selectedSections.push('舞台');
        }
        if (event.parsedSections?.sound > 0 || event.description?.includes('音響:')) {
            selectedSections.push('音響');
        }
        if (event.parsedSections?.lighting > 0 || event.description?.includes('照明:')) {
            selectedSections.push('照明');
        }

        if (selectedSections.length > 0) {
            drawerSectionLabel.textContent = selectedSections.join('・') + ' 増員';
        } else {
            drawerSectionLabel.textContent = '増員';
        }
    } else {
        drawerSectionLabel.textContent = section.name + '増員';
    }
    drawerHallName.textContent = event.hall || '会場未定';

    // 催事名（descriptionの【催事名】から抽出、またはAPIのeventNameを使用）
    let eventNameText = '';
    const eventNameMatch = event.description?.match(/【催事名】\n?　?(.+)/);
    if (eventNameMatch) {
        eventNameText = eventNameMatch[1].trim();
    } else if (event.eventName) {
        eventNameText = event.eventName;
    } else {
        eventNameText = event.title;
    }
    drawerEventName.textContent = eventNameText;

    // 日付と区分の表示（各日付の横に区分を表示）
    const startDate = new Date(event.startDate);
    const endDate = new Date(event.endDate);

    const rawTexts = event.timeSlotRawTexts || [];
    const getSlot = (index) => rawTexts[index] || rawTexts[rawTexts.length - 1] || timeSlots[event.timeSlot] || '全日';

    // 日ごとの募集人数を取得
    const dailyCapacities = extractDailyCapacities(event.description, event.section);
    let defaultCapacityStr = `${event.capacity}名`;
    if (event.section === 'multiple' && event.parsedSections) {
        const sectionLabels = [];
        const abbrevMap = { stage: '舞', sound: '音', lighting: '照' };
        for (const [key, count] of Object.entries(event.parsedSections)) {
            if (count > 0) {
                sectionLabels.push(`${abbrevMap[key] || key}/${count}`);
            }
        }
        if (sectionLabels.length > 0) defaultCapacityStr = sectionLabels.join(' ');
    }
    const getCapacity = (index) => dailyCapacities[index] || defaultCapacityStr;

    let gridHtml = '<div style="display: grid; grid-template-columns: auto auto 1fr; row-gap: 8px; column-gap: 12px; font-size: 1rem; align-items: baseline;">';
    // テーブルヘッダー
    gridHtml += `<div style="color: rgba(255,255,255,0.7); font-size: 0.8rem; border-bottom: 1px solid rgba(255,255,255,0.2); padding-bottom: 4px;">日付</div>`;
    gridHtml += `<div style="color: rgba(255,255,255,0.7); font-size: 0.8rem; border-bottom: 1px solid rgba(255,255,255,0.2); padding-bottom: 4px;">区分</div>`;
    gridHtml += `<div style="color: rgba(255,255,255,0.7); font-size: 0.8rem; border-bottom: 1px solid rgba(255,255,255,0.2); padding-bottom: 4px;">募集人数</div>`;

    let parsedPref = '';
    if (event.description) {
        const prefMatch = event.description.match(/💡連日通し希望💡\n　([^\n]+)/);
        if (prefMatch) {
            parsedPref = prefMatch[1].trim();
        }
    }
    const pref = event.consecutivePreference || event.extendedProps?.consecutivePreference || parsedPref || '';
    const prefHtml = pref ? `&nbsp;&nbsp;💡連日通し希望💡 ${pref}` : '';

    if (event.groupId && event.relatedDates && event.relatedDates.length > 1) {
        event.relatedDates.forEach((dateStr, i) => {
            const d = new Date(dateStr);
            gridHtml += `<div style="white-space: nowrap; font-weight: bold; color: white;">${formatDateJP(d)}</div>`;
            gridHtml += `<div style="white-space: nowrap; font-weight: bold; color: white;">${getSlot(i)}</div>`;
            gridHtml += `<div style="white-space: nowrap; font-weight: bold; color: white;">${getCapacity(i)}</div>`;
        });
        gridHtml += `<div style="grid-column: 1 / -1; margin-top:2px; font-size:0.85em; opacity:0.8;">（${event.relatedDates.length}日間）${prefHtml}</div>`;
    } else if (event.startDate !== event.endDate) {
        const dayDiff = Math.floor((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
        for (let i = 0; i < dayDiff; i++) {
            const d = new Date(startDate);
            d.setDate(d.getDate() + i);
            gridHtml += `<div style="white-space: nowrap; font-weight: bold; color: white;">${formatDateJP(d)}</div>`;
            gridHtml += `<div style="white-space: nowrap; font-weight: bold; color: white;">${getSlot(i)}</div>`;
            gridHtml += `<div style="white-space: nowrap; font-weight: bold; color: white;">${getCapacity(i)}</div>`;
        }
        gridHtml += `<div style="grid-column: 1 / -1; margin-top:2px; font-size:0.85em; opacity:0.8;">（${dayDiff}日間）${prefHtml}</div>`;
    } else {
        gridHtml += `<div style="white-space: nowrap; font-weight: bold; color: white;">${formatDateJP(startDate)}</div>`;
        gridHtml += `<div style="white-space: nowrap; font-weight: bold; color: white;">${getSlot(0)}</div>`;
        gridHtml += `<div style="white-space: nowrap; font-weight: bold; color: white;">${getCapacity(0)}</div>`;
        if (pref) {
            gridHtml += `<div style="grid-column: 1 / -1; margin-top:2px; font-size:0.85em; opacity:0.8;">${prefHtml}</div>`;
        }
    }
    gridHtml += '</div>';
    drawerDate.innerHTML = gridHtml;

    // 説明文（GASからのdescriptionがある場合はそれを使用）
    if (event.description) {
        let displayDesc = event.description;
        // 【催事名】ブロックを削除（最後に空行が必要な場合もケア）
        displayDesc = displayDesc.replace(/\n?【催事名】\n　[^\n]+(\n|$)/g, '$1');
        // 💡連日通し希望💡ブロックを削除
        displayDesc = displayDesc.replace(/\n?💡連日通し希望💡\n　[^\n]+(\n|$)/g, '$1');
        
        // GASからのフォーマット済み説明文を表示（改行を<br>に変換）
        drawerDescription.innerHTML = displayDesc.replace(/\n/g, '<br>');
    } else {
        drawerDescription.textContent = `${event.title}のスタッフを募集しています。経験者歓迎。`;
    }

    // セクション別申込者カウントを表示
    const applicantCounts = event.extendedProps?.applicantCounts || event.applicantCounts || { stage: 0, sound: 0, lighting: 0 };
    const parsedSections = event.parsedSections || event.extendedProps?.sections || { stage: 0, sound: 0, lighting: 0 };
    
    // 募集セクションに対してカウントをリスト表示
    let listHtml = '';
    const sectionNameMap = { stage: '舞台', sound: '音響', lighting: '照明' };
    let hasAnySections = false;
    
    for (const [key, recruitCount] of Object.entries(parsedSections)) {
        if (recruitCount > 0) {
            hasAnySections = true;
            const count = applicantCounts[key] || 0;
            listHtml += `<div style="padding: 2px 0;">・${sectionNameMap[key] || key}：${count}名申し込み中</div>`;
        }
    }
    
    if (!hasAnySections) {
        listHtml = '<div style="padding: 2px 0; opacity: 0.7;">セクション情報なし</div>';
    }
    
    noticeTitle.textContent = '現在の申し込み状況';
    noticeApplicantList.innerHTML = listHtml;
    noticeFooter.textContent = '';
    noticeFooter.style.display = 'none';

    // 日付選択の表示/非表示を設定
    setupDateSelection(event);

    // イベントキーを生成
    const eventKey = `${event.title}_${event.startDate}`;
    currentSelectedEvent.eventKey = eventKey;

    // 募集完了状態の確認
    const isClosed = recruitmentStatuses[eventKey] === true;
    const formSection = document.querySelector('.drawer-form-section');
    const noticeEl = document.getElementById('drawerNotice');

    // 既存の管理者ボタンとオーバーレイをクリア
    const existingAdminBtn = document.getElementById('adminRecruitBtn');
    if (existingAdminBtn) existingAdminBtn.remove();
    const existingOverlay = document.getElementById('closedOverlay');
    if (existingOverlay) existingOverlay.remove();

    if (isClosed) {
        // 募集完了（B案）: フォームを非表示、インフォエリアに終了メッセージを表示
        if (formSection) formSection.style.display = 'none';
        if (noticeEl) noticeEl.style.display = '';
        
        // インフォエリアを「募集終了」テキストに上書き
        noticeTitle.textContent = '本案件の募集は終了しました';
        noticeApplicantList.innerHTML = '';
        noticeFooter.style.display = 'none';
        document.querySelector('#drawerNotice .notice-icon').textContent = '⛔';
    } else {
        // 募集中: フォームを表示
        if (formSection) formSection.style.display = '';
        if (noticeEl) noticeEl.style.display = '';
        document.querySelector('#drawerNotice .notice-icon').textContent = 'ℹ️';
    }

    // 管理者ボタン表示（自分のホールのみ）
    const isAdmin = window.isAdmin || localStorage.getItem('zouin_is_admin') === 'true';
    const staffHall = localStorage.getItem('zouin_staff_hall');
    const eventHall = event.hall || '';

    console.log('管理者チェック:', { isAdmin, staffHall, eventHall, isClosed, match: eventHall === staffHall, eventKey });

    if (isAdmin && staffHall && eventHall === staffHall) {
        const adminBtn = document.createElement('button');
        adminBtn.id = 'adminRecruitBtn';
        adminBtn.type = 'button';
        const btnColor = isClosed ? '#34c759' : '#ff3b30';
        const btnText = isClosed ? '▶ 募集を再開する' : '◼ 募集を終了する';
        adminBtn.style.cssText = `display:block;width:100%;padding:14px;margin:12px 0 0;border:none;background:${btnColor};color:#fff;font-weight:700;font-size:15px;border-radius:12px;cursor:pointer;letter-spacing:0.5px;box-shadow:0 2px 8px ${btnColor}40;`;
        adminBtn.textContent = btnText;
        adminBtn.onclick = () => toggleRecruitment(eventKey, eventHall, !isClosed);

        const drawerBody = document.querySelector('.drawer-inner');
        drawerBody.appendChild(adminBtn);
    }

    // オーバーレイを表示
    drawerOverlay.classList.add('active');
    document.body.style.overflow = 'hidden';
}

/**
 * 日付選択UIを設定
 */
function setupDateSelection(event) {
    const { formatDateJP } = window.calendarApp;
    const dateTabs = document.getElementById('dateTabs');

    // 日付選択をリセット
    selectedDates = [];
    dateTabs.innerHTML = '';

    // relatedDates（グループ化された非連日イベント）があるか確認
    const relatedDates = event.relatedDates || event.extendedProps?.relatedDates || [];
    const hasRelatedDates = Array.isArray(relatedDates) && relatedDates.length > 0;

    // 連日イベントかつ通し希望が「はい」「あり」以外の場合に表示
    const isMultiDay = event.isMultiDay || event.extendedProps?.isMultiDay ||
        (event.startDate !== event.endDate) || hasRelatedDates;
    const pref = event.consecutivePreference || event.extendedProps?.consecutivePreference || '';

    // 「はい」「あり」の場合は個別選択不可（ただしrelatedDatesがある場合は選択可）
    const showDateSelection = isMultiDay && (pref !== 'はい' && pref !== 'あり' || hasRelatedDates);

    console.log('日付選択判定:', { isMultiDay, pref, showDateSelection, hasRelatedDates, event });

    if (!showDateSelection) {
        dateSelectionGroup.classList.add('hidden');
        return;
    }

    // 日付選択を表示
    dateSelectionGroup.classList.remove('hidden');

    // イベントオブジェクトの構造を確認
    console.log('イベントオブジェクト:', JSON.stringify(event, null, 2));

    let dates = [];

    if (hasRelatedDates) {
        // relatedDatesから日付を生成
        relatedDates.forEach(d => {
            const parsed = new Date(d);
            if (!isNaN(parsed.getTime())) {
                dates.push(parsed);
            }
        });
        console.log('relatedDatesから日付生成:', dates.length);
    } else {
        // startDateとendDateから日付を生成
        const startStr = event.startDate || event.start;
        const endStr = event.endDate || event.end || startStr;
        console.log('start/end:', startStr, endStr);

        const startDate = new Date(startStr);
        const endDate = new Date(endStr);

        if (isNaN(startDate.getTime())) {
            console.error('無効な開始日:', startStr);
            dateSelectionGroup.classList.add('hidden');
            return;
        }
        if (isNaN(endDate.getTime())) {
            console.error('無効な終了日:', endStr);
            dateSelectionGroup.classList.add('hidden');
            return;
        }

        // 日付リストを生成
        let currentDate = new Date(startDate);
        let loopCount = 0;
        while (currentDate <= endDate && loopCount < 10) {
            dates.push(new Date(currentDate));
            currentDate.setDate(currentDate.getDate() + 1);
            loopCount++;
        }
    }

    console.log('生成された日付:', dates.length, '件');

    // 日付がない場合はエラー表示
    if (dates.length === 0) {
        console.error('日付リストが空です');
        dateSelectionGroup.classList.add('hidden');
        return;
    }

    // 個別日のタブを生成
    dates.forEach((d, index) => {
        const tab = createDateTab(d, `${index + 1}日目のみ`);
        dateTabs.appendChild(tab);
    });

    // 両日/全日タブを追加（2日以上の場合）
    if (dates.length >= 2) {
        const allDaysLabel = dates.length === 2 ? '両日' : `全${dates.length}日`;
        const allTab = document.createElement('div');
        allTab.className = 'date-tab date-tab-all';
        allTab.dataset.value = 'all';
        allTab.dataset.dates = dates.map(d => d.toISOString().split('T')[0]).join(',');
        allTab.innerHTML = `<span class="date-tab-all-label">${allDaysLabel}</span>`;
        allTab.addEventListener('click', () => handleTabClick(allTab, dates));
        dateTabs.appendChild(allTab);
    }
}

/**
 * 日付タブ要素を作成
 */
function createDateTab(date, label) {
    const tab = document.createElement('div');
    tab.className = 'date-tab';
    const dateStr = date.toISOString().split('T')[0];
    tab.dataset.value = dateStr;

    // 曜日と日付を抽出
    const dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'][date.getDay()];
    const month = date.getMonth() + 1;
    const day = date.getDate();

    tab.innerHTML = `
        <span class="date-tab-label">${label}</span>
        <span class="date-tab-date">${month}/${day}(${dayOfWeek})</span>
    `;
    tab.addEventListener('click', () => handleTabClick(tab, null));
    return tab;
}

/**
 * タブクリック時の処理
 */
function handleTabClick(tab, allDates) {
    const dateTabs = document.getElementById('dateTabs');

    if (tab.dataset.value === 'all') {
        // 全日タブがクリックされた場合
        const isActive = tab.classList.contains('active');

        // 全てのタブの状態を更新
        dateTabs.querySelectorAll('.date-tab').forEach(t => {
            if (isActive) {
                t.classList.remove('active');
            } else {
                t.classList.add('active');
            }
        });
    } else {
        // 個別タブがクリックされた場合
        tab.classList.toggle('active');

        // 全日タブの状態を更新
        const allTab = dateTabs.querySelector('[data-value="all"]');
        if (allTab) {
            const individualTabs = dateTabs.querySelectorAll('.date-tab:not([data-value="all"])');
            const activeTabs = dateTabs.querySelectorAll('.date-tab.active:not([data-value="all"])');
            if (activeTabs.length === individualTabs.length) {
                allTab.classList.add('active');
            } else {
                allTab.classList.remove('active');
            }
        }
    }

    // 選択状態を更新
    updateSelectedDates();
}

/**
 * 選択された日付を更新
 */
function updateSelectedDates() {
    const dateTabs = document.getElementById('dateTabs');
    selectedDates = [];

    dateTabs.querySelectorAll('.date-tab.active:not([data-value="all"])').forEach(tab => {
        selectedDates.push(tab.dataset.value);
    });

    console.log('選択された日付:', selectedDates);
}

/**
 * カスタム確認ダイアログを表示（Promise版）
 */
function showCustomConfirm(message) {
    return new Promise((resolve) => {
        // 既存ダイアログがあれば削除
        const existing = document.getElementById('customConfirmOverlay');
        if (existing) existing.remove();

        const overlay = document.createElement('div');
        overlay.id = 'customConfirmOverlay';
        overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.6);display:flex;align-items:center;justify-content:center;z-index:99999;padding:20px;';

        const dialog = document.createElement('div');
        dialog.style.cssText = 'background:rgba(30,35,55,0.92);backdrop-filter:blur(20px);-webkit-backdrop-filter:blur(20px);border-radius:20px;padding:32px 28px 24px;max-width:340px;width:100%;border:1px solid rgba(255,255,255,0.12);box-shadow:0 20px 60px rgba(0,0,0,0.5);text-align:center;';

        dialog.innerHTML = `
            <div style="font-size:48px;margin-bottom:16px;">⚠️</div>
            <div style="color:#fff;font-size:1.2rem;font-weight:700;line-height:1.6;margin-bottom:28px;">${message}</div>
            <div style="display:flex;gap:12px;">
                <button id="customConfirmCancel" style="flex:1;padding:14px;border:1px solid rgba(255,255,255,0.2);border-radius:12px;background:rgba(255,255,255,0.08);color:#fff;font-size:1rem;font-weight:600;cursor:pointer;">キャンセル</button>
                <button id="customConfirmOk" style="flex:1;padding:14px;border:none;border-radius:12px;background:linear-gradient(135deg,#f87171,#ef4444);color:#fff;font-size:1rem;font-weight:600;cursor:pointer;box-shadow:0 4px 15px rgba(239,68,68,0.3);">OK</button>
            </div>
        `;

        overlay.appendChild(dialog);
        document.body.appendChild(overlay);

        document.getElementById('customConfirmOk').addEventListener('click', () => {
            overlay.remove();
            resolve(true);
        });
        document.getElementById('customConfirmCancel').addEventListener('click', () => {
            overlay.remove();
            resolve(false);
        });
        overlay.addEventListener('click', (e) => {
            if (e.target === overlay) {
                overlay.remove();
                resolve(false);
            }
        });
    });
}

/**
 * 募集完了ステータスを切り替え（管理者用）
 */
async function toggleRecruitment(eventKey, hall, newStatus) {
    const email = localStorage.getItem('zouin_staff_name') || '';
    const confirmMsg = newStatus ? 'この募集を完了にしますか？' : 'この募集を再開しますか？';
    const confirmed = await showCustomConfirm(confirmMsg);
    if (!confirmed) return;

    const btn = document.getElementById('adminRecruitBtn');
    if (btn) {
        btn.textContent = '処理中...';
        btn.disabled = true;
    }

    try {
        const GAS_URL = window.calendarApp?.API_CONFIG?.GAS_URL || DRAWER_GAS_URL;

        const result = await new Promise((resolve) => {
            const callbackName = 'recruitCallback_' + Date.now();
            window[callbackName] = (data) => {
                delete window[callbackName];
                resolve(data);
            };
            const params = new URLSearchParams({
                action: 'toggleRecruitment',
                eventKey: eventKey,
                hall: hall,
                status: newStatus.toString(),
                email: email,
                callback: callbackName
            });
            const script = document.createElement('script');
            script.src = GAS_URL + '?' + params.toString();
            script.onerror = () => resolve({ success: false, error: '通信エラー' });
            document.body.appendChild(script);
        });

        if (result.success) {
            recruitmentStatuses[eventKey] = result.closed;
            // ドロワーを再描画
            if (currentSelectedEvent) {
                closeDrawer();
                setTimeout(() => {
                    openDrawer(currentSelectedDate, [currentSelectedEvent]);
                }, 100);
            }
            // カレンダーも更新
            if (window.calendarApp?.refreshCalendar) {
                window.calendarApp.refreshCalendar();
            }
        } else {
            alert('エラー: ' + (result.error || '不明なエラー'));
        }
    } catch (error) {
        console.error('募集完了切替エラー:', error);
        alert('エラーが発生しました。');
    }
}

/**
 * 募集完了ステータスを読み込み（起動時に呼ばれる）
 */
async function loadRecruitmentStatuses() {
    try {
        const GAS_URL = window.calendarApp?.API_CONFIG?.GAS_URL || DRAWER_GAS_URL;

        const result = await new Promise((resolve) => {
            const callbackName = 'statusCallback_' + Date.now();
            window[callbackName] = (data) => {
                delete window[callbackName];
                resolve(data);
            };
            const params = new URLSearchParams({
                action: 'getRecruitmentStatuses',
                callback: callbackName
            });
            const script = document.createElement('script');
            script.src = GAS_URL + '?' + params.toString();
            script.onerror = () => resolve({ success: false });
            document.body.appendChild(script);
        });

        if (result.success && result.statuses) {
            recruitmentStatuses = result.statuses;
            console.log('募集完了ステータス:', recruitmentStatuses);
        }
    } catch (error) {
        console.error('募集完了ステータス取得エラー:', error);
    }
}

// 起動時に募集完了ステータスを読み込み
document.addEventListener('DOMContentLoaded', () => {
    loadRecruitmentStatuses();
});


/**
 * ドロワーを閉じる
 */
function closeDrawer() {
    drawer.classList.remove('active');
    drawerOverlay.classList.remove('active');
    document.body.style.overflow = '';

    // フォームをリセット
    applyForm.reset();

    // リセット後にログイン情報を復元
    const savedName = window.loggedInStaffName || localStorage.getItem('zouin_staff_display_name');
    const savedHall = localStorage.getItem('zouin_staff_hall');
    const savedSection = localStorage.getItem('zouin_staff_section');
    if (savedName) document.getElementById('userName').value = savedName;
    if (savedHall) document.getElementById('staffHall').value = savedHall;
    if (savedSection) document.getElementById('staffSection').value = savedSection;
}

/**
 * 確認ダイアログを表示
 */
function showConfirmDialog() {
    confirmDialog.classList.remove('hidden');
}

/**
 * 確認ダイアログを閉じる
 */
function hideConfirmDialog() {
    confirmDialog.classList.add('hidden');
    closeDrawer();
}

/**
 * フォーム送信処理
 */
async function handleFormSubmit(e) {
    e.preventDefault();

    // 申し込み前に権限チェック
    if (window.checkPermission) {
        const hasPermission = await window.checkPermission();
        if (!hasPermission) {
            return;
        }
    }

    // 日付選択の検証（表示されている場合）
    if (!dateSelectionGroup.classList.contains('hidden') && selectedDates.length === 0) {
        alert('応募する日程を選択してください');
        return;
    }

    const submitBtn = applyForm.querySelector('button[type="submit"]');
    const originalBtnText = submitBtn.textContent;
    submitBtn.textContent = '送信中...';
    submitBtn.disabled = true;

    try {
        const email = localStorage.getItem('zouin_staff_name') || '';
        const staffName = localStorage.getItem('zouin_staff_display_name') || document.getElementById('userName').value;
        const staffHall = document.getElementById('staffHall').value;
        const staffSection = document.getElementById('staffSection').value;
        // 催事名を取得（descriptionの【催事名】から抽出、またはAPIのeventNameを使用）
        let eventTitle = '';
        const evtNameMatch = currentSelectedEvent?.description?.match(/【催事名】\n?　?(.+)/);
        if (evtNameMatch) {
            eventTitle = evtNameMatch[1].trim();
        } else if (currentSelectedEvent?.eventName) {
            eventTitle = currentSelectedEvent.eventName;
        } else if (currentSelectedEvent?.extendedProps?.eventName) {
            eventTitle = currentSelectedEvent.extendedProps.eventName;
        } else {
            eventTitle = currentSelectedEvent?.title || '';
        }
        const hall = currentSelectedEvent?.extendedProps?.hall || '';
        const dates = selectedDates.length > 0 ? selectedDates.join(', ') : currentSelectedDate;

        // LINE公式アカウントにメッセージを送信するURLを生成（先に作成）
        const lineMessage = buildLineMessage(staffHall, staffName, staffSection, dates, hall, eventTitle);
        const lineUrl = 'https://line.me/R/oaMessage/@825gnfcx/?' + encodeURIComponent(lineMessage);

        // ※Webからの「確認中」ステータス送信を完全に廃止し、直接LINEへ誘導するのみに変更

        // 確認ダイアログ表示後にLINEを開く
        showConfirmDialogWithLine(lineUrl);
    } catch (error) {
        console.error('送信エラー:', error);
        alert('送信に失敗しました。もう一度お試しください。');
    } finally {
        submitBtn.textContent = originalBtnText;
        submitBtn.disabled = false;
    }
}

/**
 * 日付を YYYY-MM-DD 形式に変換（カウント処理で正確にパースできるようにするため）
 */
function formatDateShort(dateStr) {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

/**
 * LINEプリセットメッセージを生成
 */
function buildLineMessage(staffHall, staffName, staffSection, dates, eventHall, eventTitle) {
    // 日付を○月○日形式に変換し、複数日は改行で表示
    const datesStr = String(dates || '');
    const dateList = datesStr.split(',').map(d => formatDateShort(d.trim()));
    const formattedDates = dateList.length > 1
        ? '\n' + dateList.map(d => `　・${d}`).join('\n')
        : dateList[0];

    return `下記の内容で申し込みます。

■申し込みスタッフ情報■
【所属ホール】
${staffHall}
【お名前】
${staffName}
【セクション】
${staffSection}

■申し込み催事■
【催事名】
${eventTitle || ''}
【増員日】
${formattedDates}
【募集事業所】
${eventHall}`;
}

/**
 * LINE送信付き確認ダイアログを表示
 */
function showConfirmDialogWithLine(lineUrl) {
    const dialog = document.getElementById('confirmDialog');
    const messageEl = dialog.querySelector('.confirm-message');

    // メッセージを更新
    messageEl.innerHTML = `
        <strong>CONTROL HUBに申し込み内容を送信してください。</strong><br><br>
        下のボタンを押すとLINEが開きます。<br>
        メッセージが入力された状態で開くので、<br>
        <strong>「送信」を押すだけ</strong>でOKです。
    `;

    // LINEボタンを追加
    let lineBtn = dialog.querySelector('.line-send-btn');
    if (!lineBtn) {
        lineBtn = document.createElement('a');
        lineBtn.className = 'line-send-btn';
        lineBtn.style.cssText = 'display:block;background:#06C755;color:#fff;text-align:center;padding:14px;border-radius:8px;font-weight:bold;font-size:16px;text-decoration:none;margin:12px 0 8px;';
        const closeBtn = dialog.querySelector('.confirm-btn');
        closeBtn.parentNode.insertBefore(lineBtn, closeBtn);
    }
    lineBtn.href = lineUrl;
    lineBtn.target = '_blank';
    lineBtn.textContent = '📱 LINEで送信する';

    dialog.classList.remove('hidden');
}

/**
 * GASに申し込みデータを送信（JSONP方式）
 */
function submitToGAS(data) {
    return new Promise((resolve) => {
        const GAS_URL = window.calendarApp?.API_CONFIG?.GAS_URL || DRAWER_GAS_URL;

        if (!GAS_URL) {
            console.error('GAS URLが設定されていません');
            resolve({ success: false, error: 'GAS URLが設定されていません' });
            return;
        }

        const callbackName = 'submitCallback_' + Date.now();

        const params = new URLSearchParams(data);
        params.append('callback', callbackName);

        window[callbackName] = function (response) {
            delete window[callbackName];
            if (script.parentNode) document.body.removeChild(script);
            resolve(response);
        };

        const script = document.createElement('script');
        script.src = `${GAS_URL}?${params.toString()}`;
        script.onerror = function (e) {
            console.error('JSONP script error:', e, 'URL:', script.src.substring(0, 100));
            delete window[callbackName];
            if (script.parentNode) document.body.removeChild(script);
            resolve({ success: false, error: 'ネットワークエラー: GASへの接続に失敗しました' });
        };

        setTimeout(() => {
            if (window[callbackName]) {
                console.error('JSONP タイムアウト');
                delete window[callbackName];
                if (script.parentNode) document.body.removeChild(script);
                resolve({ success: false, error: 'タイムアウト（30秒）' });
            }
        }, 30000);

        document.body.appendChild(script);
    });
}

// イベントリスナー
drawerOverlay.addEventListener('click', closeDrawer);
drawerClose.addEventListener('click', closeDrawer);
applyForm.addEventListener('submit', handleFormSubmit);
confirmClose.addEventListener('click', hideConfirmDialog);

/**
 * 説明文から日ごとの募集人数を抽出するヘルパー関数
 */
function extractDailyCapacities(description, eventSection) {
    if (!description) return [];
    
    // ◆n日目で分割
    const blocks = description.split(/◆\d+日目/);
    if (blocks.length <= 1) {
        // 単日イベントの場合
        return [extractCapacityFromBlock(description, eventSection)];
    }
    
    // 最初の要素（◆1日目より前のテキスト）を削除
    blocks.shift();
    return blocks.map(block => extractCapacityFromBlock(block, eventSection));
}

function extractCapacityFromBlock(block, eventSection) {
    let count = 0;
    let details = [];
    
    if (eventSection === 'stage' || eventSection === 'multiple') {
        const m = block.match(/舞台:\s*(\d+)人/);
        if (m) { count += parseInt(m[1], 10); details.push(`舞${m[1]}`); }
    }
    if (eventSection === 'sound' || eventSection === 'multiple') {
        const m = block.match(/音響:\s*(\d+)人/);
        if (m) { count += parseInt(m[1], 10); details.push(`音${m[1]}`); }
    }
    if (eventSection === 'lighting' || eventSection === 'multiple') {
        const m = block.match(/照明:\s*(\d+)人/);
        if (m) { count += parseInt(m[1], 10); details.push(`照${m[1]}`); }
    }
    
    if (eventSection === 'multiple' && details.length > 0) {
        return details.join(' ');
    }
    return count > 0 ? `${count}名` : '0名';
}

// ドラッグでドロワーを閉じる
let isDragging = false;
let startY = 0;
let currentY = 0;

drawerInner.addEventListener('touchstart', (e) => {
    // スクロール位置が一番上の時のみドラッグを有効にする
    if (drawerInner.scrollTop === 0) {
        isDragging = true;
        startY = e.touches[0].clientY;
    }
}, { passive: true });

drawerInner.addEventListener('touchmove', (e) => {
    if (!isDragging) return;

    currentY = e.touches[0].clientY;
    const diff = currentY - startY;

    if (diff > 0) {
        drawer.style.transform = `translateY(${diff}px)`;
    }
}, { passive: true });

drawerInner.addEventListener('touchend', () => {
    if (!isDragging) return;

    isDragging = false;
    const diff = currentY - startY;

    if (diff > 100) {
        closeDrawer();
    }

    drawer.style.transform = '';
}, { passive: true });

/**
 * イベント選択モーダルを開く
 */
function showEventModal(date, events) {
    const { sectionInfo, formatDateJP } = window.calendarApp;

    // タイトルを設定
    eventModalTitle.textContent = `${formatDateJP(date)}のイベント`;

    // イベントリストを生成
    eventModalList.innerHTML = '';
    events.forEach(event => {
        const section = sectionInfo[event.section];
        const item = document.createElement('div');
        item.className = `event-modal-item section-${event.section}`;
        item.innerHTML = `
            <span class="event-modal-section">${section.name}</span>
            <span class="event-modal-name">${event.hall || '会場未定'}</span>
            <span class="event-modal-status">${event.applied}/${event.capacity}</span>
        `;
        item.addEventListener('click', () => {
            closeEventModal();
            openDrawer(date, [event]);
        });
        eventModalList.appendChild(item);
    });

    // モーダルを表示
    eventModalOverlay.classList.add('active');
    eventModal.classList.add('active');
}

/**
 * イベント選択モーダルを閉じる
 */
function closeEventModal() {
    eventModalOverlay.classList.remove('active');
    eventModal.classList.remove('active');
}

// イベントモーダルのイベントリスナー
eventModalOverlay.addEventListener('click', closeEventModal);
eventModalClose.addEventListener('click', closeEventModal);

// calendar.jsから呼び出せるようにグローバルに公開
window.openDrawer = openDrawer;
window.showEventModal = showEventModal;
