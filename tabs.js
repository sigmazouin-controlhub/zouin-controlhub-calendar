/**
 * 増員 CONTROL HUB 2.0 - タブナビゲーション & 各タブコンテンツ
 * tabs.js
 */

(function () {
    'use strict';

    // ============================================
    // GAS API URL（calendar.js と同じ）
    // ============================================
    const GAS_URL = 'https://script.google.com/macros/s/AKfycbwqeiRNp7oueT0NZqxNrn5n4OONm6Y-NUaNhXEIAS5QrR8boqQzG7fdc0VQ_n_sPZ6_/exec';

    // 増員完了報告のリンクURL（ここを変更してください）
    const REPORT_URL = '';

    // ============================================
    // タブ切り替え
    // ============================================
    const tabButtons = document.querySelectorAll('.tab-btn');
    const tabContents = document.querySelectorAll('.tab-content');
    let currentTab = 'tabCalendar';
    let hallDataCache = null;

    // ============================================
    // キープ（お気に入り）機能ヘルパー
    // ============================================
    const KEEP_STORAGE_KEY = 'zouin_keep_events';

    /** イベントから一意なIDを生成 */
    function generateEventId(ev) {
        const title = ev.title || ev.hall || '';
        const start = ev.start || ev.startDate || '';
        return btoa(unescape(encodeURIComponent(title + '_' + start))).replace(/[=+/]/g, '');
    }

    /** キープ中のイベントID一覧を取得 */
    function getKeepEvents() {
        try {
            return JSON.parse(localStorage.getItem(KEEP_STORAGE_KEY)) || [];
        } catch (e) {
            return [];
        }
    }

    /** イベントのキープ状態をトグル */
    function toggleKeepEvent(eventId) {
        const keeps = getKeepEvents();
        const index = keeps.indexOf(eventId);
        if (index >= 0) {
            keeps.splice(index, 1);
        } else {
            keeps.push(eventId);
        }
        localStorage.setItem(KEEP_STORAGE_KEY, JSON.stringify(keeps));
        return index < 0; // trueなら追加、falseなら削除
    }

    /** イベントがキープ中かどうか */
    function isKept(eventId) {
        return getKeepEvents().includes(eventId);
    }

    // 現在のフィルタ状態
    let recruitmentFilter = 'all'; // 'all' | 'kept'

    function switchTab(tabId) {
        // コンテンツの切り替え
        tabContents.forEach(content => {
            content.classList.toggle('active', content.id === tabId);
        });
        // ボタンのアクティブ状態
        tabButtons.forEach(btn => {
            btn.classList.toggle('active', btn.dataset.tab === tabId);
        });
        currentTab = tabId;

        // タブ固有の初期化
        if (tabId === 'tabHallInfo' && !hallDataCache) {
            loadHallList();
        }
        if (tabId === 'tabMyPage') {
            renderMyPage();
        }
        if (tabId === 'tabRecruitment') {
            renderRecruitmentList();
        }
    }

    tabButtons.forEach(btn => {
        btn.addEventListener('click', () => switchTab(btn.dataset.tab));
    });

    // ============================================
    // ヘッダーのユーザー情報表示
    // ============================================
    function updateHeaderUserInfo() {
        const nameEl = document.getElementById('headerUserName');
        const roleEl = document.getElementById('headerUserRole');
        if (!nameEl) return;

        const name = window.loggedInStaffName || localStorage.getItem('zouin_staff_display_name') || '';
        const isSystemAdmin = window.isSystemAdmin || localStorage.getItem('zouin_is_system_admin') === 'true';
        const isAdmin = window.isAdmin || localStorage.getItem('zouin_is_admin') === 'true';

        nameEl.textContent = name;
        if (roleEl) {
            if (isSystemAdmin) {
                roleEl.textContent = 'システム管理者';
                roleEl.className = 'header-user-role system-admin';
            } else if (isAdmin) {
                roleEl.textContent = '管理者';
                roleEl.className = 'header-user-role admin';
            } else {
                roleEl.textContent = 'スタッフ';
                roleEl.className = 'header-user-role';
            }
        }
    }

    // showApp後に呼ばれるようにする
    const _origShowApp = window._tabsShowAppHook;
    window._tabsUpdateHeader = updateHeaderUserInfo;

    // 定期的にチェック（ログイン後の反映用）
    let headerCheckCount = 0;
    const headerCheckInterval = setInterval(() => {
        if (window.loggedInStaffName || headerCheckCount > 20) {
            updateHeaderUserInfo();
            clearInterval(headerCheckInterval);
        }
        headerCheckCount++;
    }, 500);

    // ============================================
    // 今後の増員募集セクション（カレンダータブ下部）
    // 【要件1】カードタップ→ドロワー直接起動
    // 【要件3】キープボタン追加
    // ============================================
    function renderUpcomingEvents() {
        const container = document.getElementById('upcomingEventsContainer');
        const sectionTitle = document.getElementById('upcomingSectionTitle');
        if (!container) return;

        const userSection = localStorage.getItem('zouin_staff_section') || '';
        const sectionLabel = userSection || 'あなた';

        // セクションカラー
        let sectionColor = '#60a5fa';
        let sectionDot = 'sound';
        if (userSection.includes('舞台')) { sectionColor = '#10b981'; sectionDot = 'stage'; }
        else if (userSection.includes('音響')) { sectionColor = '#0ea5e9'; sectionDot = 'sound'; }
        else if (userSection.includes('照明')) { sectionColor = '#eab308'; sectionDot = 'lighting'; }

        if (sectionTitle) {
            sectionTitle.innerHTML = `<span class="legend-dot ${sectionDot}" style="margin-right:6px;"></span>
                <strong>${sectionLabel}</strong> の今後の増員募集`;
        }

        // calendar.js のイベントデータを利用
        const events = window.calendarEvents || [];
        const now = new Date();
        now.setHours(0, 0, 0, 0);

        // 未来のイベントをフィルタ & セクションマッチ
        const upcoming = events.filter(ev => {
            const start = new Date(ev.start);
            if (start < now) return false;
            if (!ev.sections) return true;
            if (!userSection) return true;
            // セクションマッチ
            if (userSection.includes('舞台') && ev.sections.stage > 0) return true;
            if (userSection.includes('音響') && ev.sections.sound > 0) return true;
            if (userSection.includes('照明') && ev.sections.lighting > 0) return true;
            return false;
        }).sort((a, b) => new Date(a.start) - new Date(b.start)).slice(0, 4);

        if (upcoming.length === 0) {
            container.innerHTML = '<div class="upcoming-empty">現在、募集中のイベントはありません</div>';
            return;
        }

        container.innerHTML = upcoming.map(ev => {
            const start = new Date(ev.start);
            const weekdays = ['日', '月', '火', '水', '木', '金', '土'];
            const month = start.getMonth() + 1;
            const day = start.getDate();
            const dayOfWeek = weekdays[start.getDay()];

            // 残日数
            const diff = Math.ceil((start - now) / (1000 * 60 * 60 * 24));
            const daysLeft = diff > 0 ? `あと${diff}日` : '本日';

            // ホール名取得
            const hall = ev.hall || ev.title || '';
            const eventName = ev.eventName || '';

            // 区分
            const timeSlot = ev.timeSlot || '';

            // 募集セクション
            let recruitInfo = '';
            if (ev.sections) {
                const parts = [];
                if (ev.sections.stage > 0) parts.push(`舞台 ${ev.sections.stage}名`);
                if (ev.sections.sound > 0) parts.push(`音響 ${ev.sections.sound}名`);
                if (ev.sections.lighting > 0) parts.push(`照明 ${ev.sections.lighting}名`);
                recruitInfo = parts.join('、');
            }

            // キープ状態
            const eventId = generateEventId(ev);
            const kept = isKept(eventId);

            return `
                <div class="upcoming-card" data-event-id="${eventId}">
                    <button class="keep-btn ${kept ? 'active' : ''}" data-event-id="${eventId}" onclick="event.stopPropagation(); window._tabsToggleKeep(this, '${eventId}');" aria-label="キープ">
                        ${kept ? '★' : '☆'}
                    </button>
                    <div class="upcoming-card-date" onclick="window._tabsOpenDrawer('${ev.start}', '${eventId}')">
                        <span class="upcoming-date-month">${month}/${day}</span>
                        <span class="upcoming-date-day">（${dayOfWeek}）</span>
                    </div>
                    <div class="upcoming-card-info" onclick="window._tabsOpenDrawer('${ev.start}', '${eventId}')">
                        <div class="upcoming-card-hall">
                            <span class="legend-dot ${sectionDot}" style="margin-right:6px;"></span>
                            ${hall}
                        </div>
                        ${eventName ? `<div class="upcoming-card-event">${eventName}</div>` : ''}
                        <div class="upcoming-card-meta">
                            ${timeSlot ? `<span>⏰ ${timeSlot}</span>` : ''}
                            ${recruitInfo ? `<span>👥 募集：${recruitInfo}</span>` : ''}
                        </div>
                    </div>
                    <div class="upcoming-card-badge" onclick="window._tabsOpenDrawer('${ev.start}', '${eventId}')">
                        <span class="badge-recruiting">募集中</span>
                        <span class="badge-days">${daysLeft}</span>
                    </div>
                </div>
            `;
        }).join('');
    }

    // calendar.js がデータ読み込み完了時に呼ぶ
    window._tabsRenderUpcoming = renderUpcomingEvents;

    // ============================================
    // カードからドロワーを直接起動するグローバル関数
    // ============================================
    /** イベントIDからイベントオブジェクトを検索 */
    function findEventById(eventId) {
        const events = window.calendarEvents || [];
        return events.find(ev => generateEventId(ev) === eventId) || null;
    }

    /** カードタップ → ドロワー起動 */
    window._tabsOpenDrawer = function (startDateStr, eventId) {
        const ev = findEventById(eventId);
        if (!ev) return;
        const date = new Date(startDateStr);
        if (window.openDrawer) {
            window.openDrawer(date, [ev]);
        }
    };

    /** キープボタンのトグル */
    window._tabsToggleKeep = function (btnEl, eventId) {
        const nowKept = toggleKeepEvent(eventId);
        btnEl.classList.toggle('active', nowKept);
        btnEl.textContent = nowKept ? '★' : '☆';

        // 増員募集タブがキープ中フィルタなら再描画
        if (currentTab === 'tabRecruitment' && recruitmentFilter === 'kept') {
            renderRecruitmentList();
        }
    };

    // ============================================
    // ホール情報タブ
    // ============================================
    function fetchFromGAS(params) {
        return new Promise((resolve) => {
            const callbackName = 'tabsCallback_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6);
            const urlParams = new URLSearchParams(params);
            urlParams.append('callback', callbackName);

            window[callbackName] = function (data) {
                delete window[callbackName];
                if (script.parentNode) document.body.removeChild(script);
                resolve(data);
            };

            const script = document.createElement('script');
            script.src = `${GAS_URL}?${urlParams.toString()}`;
            script.onerror = function () {
                delete window[callbackName];
                if (script.parentNode) document.body.removeChild(script);
                resolve({ success: false, error: 'ネットワークエラー' });
            };

            setTimeout(() => {
                if (window[callbackName]) {
                    delete window[callbackName];
                    if (script.parentNode) document.body.removeChild(script);
                    resolve({ success: false, error: 'タイムアウト' });
                }
            }, 15000);

            document.body.appendChild(script);
        });
    }

    async function loadHallList() {
        const container = document.getElementById('hallInfoContent');
        if (!container) return;

        container.innerHTML = '<div class="tab-loading"><div class="tab-spinner"></div><p>ホール情報を読み込み中...</p></div>';

        try {
            const result = await fetchFromGAS({ action: 'getHallList' });
            if (result.success && result.halls && result.halls.length > 0) {
                hallDataCache = result.halls;
                renderHallList(result.halls);
            } else {
                container.innerHTML = '<div class="tab-empty">登録されているホールがありません</div>';
            }
        } catch (error) {
            container.innerHTML = '<div class="tab-empty">ホール情報の取得に失敗しました</div>';
        }
    }

    // ============================================
    // 【要件4】ホール情報 - エリア別グループ化 + アコーディオン + 検索
    // ============================================

    /** 住所から都道府県を抽出 */
    function extractPrefecture(address) {
        if (!address) return 'その他';
        // 北海道、東京都、京都府、大阪府、〇〇県
        const match = address.match(/^(北海道|東京都|京都府|大阪府|.{2,3}県)/);
        return match ? match[1] : 'その他';
    }

    /** 都道府県名から「○○エリア」ラベルに変換 */
    function toAreaLabel(pref) {
        // 「都」「府」「県」を除去してエリア化
        return pref.replace(/(都|府|県)$/, '') + 'エリア';
    }

    function renderHallList(halls) {
        const container = document.getElementById('hallInfoContent');
        if (!container) return;

        // 都道府県でグループ化
        const groupMap = {};
        halls.forEach(hall => {
            const pref = extractPrefecture(hall.address);
            if (!groupMap[pref]) groupMap[pref] = [];
            groupMap[pref].push(hall);
        });

        // 都道府県の並び順（おおよそ北→南）
        const prefOrder = [
            '北海道', '青森県', '岩手県', '宮城県', '秋田県', '山形県', '福島県',
            '茨城県', '栃木県', '群馬県', '埼玉県', '千葉県', '東京都', '神奈川県',
            '新潟県', '富山県', '石川県', '福井県', '山梨県', '長野県', '岐阜県',
            '静岡県', '愛知県', '三重県', '滋賀県', '京都府', '大阪府', '兵庫県',
            '奈良県', '和歌山県', '鳥取県', '島根県', '岡山県', '広島県', '山口県',
            '徳島県', '香川県', '愛媛県', '高知県', '福岡県', '佐賀県', '長崎県',
            '熊本県', '大分県', '宮崎県', '鹿児島県', '沖縄県', 'その他'
        ];

        const sortedPrefs = Object.keys(groupMap).sort((a, b) => {
            const ia = prefOrder.indexOf(a);
            const ib = prefOrder.indexOf(b);
            return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib);
        });

        const colors = ['#1a5276', '#1a6b4a', '#7d3c98', '#c0392b', '#2471a3', '#148f77', '#884ea0', '#cb4335'];

        // 検索バー + エリアグループをレンダリング
        let html = `
            <div class="hall-list-header">
                <h3>📋 ホールを選択してください</h3>
                <p class="hall-list-count">${halls.length}件のホールが登録されています</p>
            </div>
            <div class="hall-search-bar">
                <span class="hall-search-icon">🔍</span>
                <input type="text" id="hallSearchInput" placeholder="ホール名で検索..." oninput="window._tabsFilterHalls(this.value)">
            </div>
            <div id="hallAreaGroups">
        `;

        let globalIndex = 0;
        sortedPrefs.forEach((pref, prefIdx) => {
            const areaHalls = groupMap[pref];
            const areaLabel = toAreaLabel(pref);

            html += `
                <div class="hall-area-group" data-pref="${pref}">
                    <div class="hall-area-header open" onclick="window._tabsToggleArea(this)">
                        <div class="hall-area-title">
                            🏛 ${areaLabel}
                            <span class="hall-area-count">${areaHalls.length}件</span>
                        </div>
                        <span class="hall-area-chevron">▼</span>
                    </div>
                    <div class="hall-area-body open">
                        <div class="hall-area-body-inner">
                            ${areaHalls.map((hall) => {
                                const color = colors[globalIndex % colors.length];
                                globalIndex++;
                                return `
                                    <div class="hall-card" data-hall-name="${hall.hallName}" onclick="window._tabsShowHallDetail('${hall.hallName.replace(/'/g, "\\'")}')" style="--hall-color: ${color}">
                                        <div class="hall-card-icon" style="background: ${color}">🏛</div>
                                        <div class="hall-card-name">${hall.hallName}</div>
                                        ${hall.address ? `<div class="hall-card-address">${hall.address.substring(0, 20)}${hall.address.length > 20 ? '...' : ''}</div>` : ''}
                                    </div>
                                `;
                            }).join('')}
                        </div>
                    </div>
                </div>
            `;
        });

        html += '</div>';
        container.innerHTML = html;
    }

    /** アコーディオンの開閉トグル */
    window._tabsToggleArea = function (headerEl) {
        headerEl.classList.toggle('open');
        const body = headerEl.nextElementSibling;
        if (body) body.classList.toggle('open');
    };

    /** ホール名検索フィルタ */
    window._tabsFilterHalls = function (query) {
        const normalizedQuery = query.trim().toLowerCase();
        const groups = document.querySelectorAll('.hall-area-group');
        let totalVisible = 0;

        groups.forEach(group => {
            const cards = group.querySelectorAll('.hall-card');
            let visibleInGroup = 0;

            cards.forEach(card => {
                const hallName = (card.getAttribute('data-hall-name') || '').toLowerCase();
                const show = !normalizedQuery || hallName.includes(normalizedQuery);
                card.style.display = show ? '' : 'none';
                if (show) visibleInGroup++;
            });

            // グループ全体の表示/非表示
            group.style.display = visibleInGroup > 0 ? '' : 'none';

            // 検索中はすべてのアコーディオンを開く
            if (normalizedQuery && visibleInGroup > 0) {
                const header = group.querySelector('.hall-area-header');
                const body = group.querySelector('.hall-area-body');
                if (header) header.classList.add('open');
                if (body) body.classList.add('open');
            }

            totalVisible += visibleInGroup;
        });

        // 検索結果なしメッセージ
        let noResultMsg = document.getElementById('hallSearchNoResult');
        if (totalVisible === 0 && normalizedQuery) {
            if (!noResultMsg) {
                noResultMsg = document.createElement('div');
                noResultMsg.id = 'hallSearchNoResult';
                noResultMsg.className = 'hall-search-no-result';
                document.getElementById('hallAreaGroups').appendChild(noResultMsg);
            }
            noResultMsg.textContent = `「${query}」に一致するホールは見つかりませんでした`;
            noResultMsg.style.display = '';
        } else if (noResultMsg) {
            noResultMsg.style.display = 'none';
        }
    };

    window._tabsShowHallDetail = function (hallName) {
        const hall = hallDataCache ? hallDataCache.find(h => h.hallName === hallName) : null;
        if (!hall) return;
        renderHallDetail(hall);
    };

    function renderHallDetail(hall) {
        const container = document.getElementById('hallInfoContent');
        if (!container) return;

        // 情報カテゴリーカード
        const categories = [
            { type: 'entry', emoji: '🏢', label: '入館方法', color: '#2e86c1', desc: '入館方法の画像を表示' },
            { type: 'nearby', emoji: '🏪', label: '周辺情報', color: '#27ae60', desc: '周辺施設の情報を表示' },
            { type: 'contact', emoji: '📞', label: '連絡先', color: '#8e44ad', desc: '事務所の連絡先を表示' },
            { type: 'smoking', emoji: '🚬', label: '喫煙所', color: '#e67e22', desc: '喫煙所の情報を表示' }
        ];

        container.innerHTML = `
            <button class="hall-back-btn" onclick="window._tabsBackToHallList()">
                ← ホール一覧に戻る
            </button>
            <div class="hall-detail-header">
                <h3>🏛 ${hall.hallName}</h3>
                <p>知りたい情報を選択してください</p>
            </div>
            <div class="hall-category-grid">
                ${categories.map(cat => `
                    <div class="hall-category-card" onclick="window._tabsShowCategory('${hall.hallName.replace(/'/g, "\\'")}', '${cat.type}')" style="--cat-color: ${cat.color}">
                        <div class="hall-category-icon" style="background: ${cat.color}">${cat.emoji}</div>
                        <div class="hall-category-label">${cat.label}</div>
                        <div class="hall-category-desc">${cat.desc}</div>
                    </div>
                `).join('')}
            </div>
        `;
    }

    window._tabsBackToHallList = function () {
        if (hallDataCache) {
            renderHallList(hallDataCache);
        } else {
            loadHallList();
        }
    };

    window._tabsShowCategory = function (hallName, type) {
        const hall = hallDataCache ? hallDataCache.find(h => h.hallName === hallName) : null;
        if (!hall) return;

        const container = document.getElementById('hallInfoContent');
        if (!container) return;

        let contentHtml = '';

        switch (type) {
            case 'entry':
                if (hall.entryImageUrl) {
                    contentHtml = `
                        <div class="hall-info-section">
                            <h4>🏢 ${hall.hallName} の入館方法</h4>
                            <div class="hall-entry-image">
                                <img src="${hall.entryImageUrl}" alt="入館方法" onerror="this.parentElement.innerHTML='<p class=\\'hall-no-data\\'>画像の読み込みに失敗しました</p>'">
                            </div>
                        </div>`;
                } else {
                    contentHtml = `<div class="hall-info-section"><h4>🏢 入館方法</h4><p class="hall-no-data">入館方法の画像はまだ登録されていません</p></div>`;
                }
                break;

            case 'nearby':
                contentHtml = `<div class="hall-info-section"><h4>🏪 ${hall.hallName} の周辺情報</h4>`;
                if (hall.nearby1 || hall.nearby2) {
                    if (hall.nearby1) contentHtml += `<div class="hall-info-block"><span class="hall-info-label">📍 周辺情報①</span><p>${hall.nearby1}</p></div>`;
                    if (hall.nearby2) contentHtml += `<div class="hall-info-block"><span class="hall-info-label">📍 周辺情報②</span><p>${hall.nearby2}</p></div>`;
                    if (hall.address) contentHtml += `<a href="https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(hall.address)}" target="_blank" class="hall-maps-btn">🗺 Google Mapsで見る</a>`;
                } else {
                    contentHtml += `<p class="hall-no-data">周辺情報はまだ登録されていません</p>`;
                }
                contentHtml += `</div>`;
                break;

            case 'contact':
                contentHtml = `<div class="hall-info-section"><h4>📞 ${hall.hallName} の連絡先</h4>`;
                if (hall.address || hall.phone || hall.email) {
                    if (hall.address) contentHtml += `<div class="hall-info-block"><span class="hall-info-label">📍 住所</span><p>${hall.address}</p></div>`;
                    if (hall.phone) contentHtml += `<div class="hall-info-block"><span class="hall-info-label">📞 舞台事務所直通</span><p>${hall.phone}</p></div>`;
                    if (hall.email) contentHtml += `<div class="hall-info-block"><span class="hall-info-label">✉️ メールアドレス</span><p>${hall.email}</p></div>`;
                    const btns = [];
                    if (hall.phone) {
                        const tel = hall.phone.replace(/[^0-9]/g, '');
                        if (tel) btns.push(`<a href="tel:${tel}" class="hall-action-btn phone">📞 電話をかける</a>`);
                    }
                    if (hall.address) btns.push(`<a href="https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(hall.address)}" target="_blank" class="hall-action-btn maps">🗺 Google Mapsで見る</a>`);
                    if (btns.length) contentHtml += `<div class="hall-action-btns">${btns.join('')}</div>`;
                } else {
                    contentHtml += `<p class="hall-no-data">連絡先情報はまだ登録されていません</p>`;
                }
                contentHtml += `</div>`;
                break;

            case 'smoking':
                contentHtml = `<div class="hall-info-section"><h4>🚬 ${hall.hallName} の喫煙所</h4>`;
                if (hall.smoking) {
                    contentHtml += `<div class="hall-info-block"><span class="hall-info-label">喫煙所について</span><p>${hall.smoking}</p></div>`;
                } else {
                    contentHtml += `<p class="hall-no-data">喫煙所の情報はまだ登録されていません</p>`;
                }
                contentHtml += `</div>`;
                break;
        }

        container.innerHTML = `
            <button class="hall-back-btn" onclick="window._tabsShowHallDetail('${hallName.replace(/'/g, "\\'")}')">
                ← ${hallName} に戻る
            </button>
            ${contentHtml}
        `;
    };

    // ============================================
    // 増員募集タブ（カレンダーのイベント一覧表示）
    // 【要件1】カードタップ→ドロワー直接起動
    // 【要件3】キープボタン + フィルタUI
    // ============================================
    function renderRecruitmentList() {
        const container = document.getElementById('recruitmentContent');
        if (!container) return;

        const events = window.calendarEvents || [];
        const now = new Date();
        now.setHours(0, 0, 0, 0);

        // 未来のイベントのみ
        const futureEvents = events.filter(ev => {
            const start = new Date(ev.start);
            return start >= now;
        }).sort((a, b) => new Date(a.start) - new Date(b.start));

        // キープ中のイベントをカウント
        const keeps = getKeepEvents();
        const keptCount = futureEvents.filter(ev => keeps.includes(generateEventId(ev))).length;

        // フィルタ適用
        let displayEvents = futureEvents;
        if (recruitmentFilter === 'kept') {
            displayEvents = futureEvents.filter(ev => keeps.includes(generateEventId(ev)));
        }

        const weekdays = ['日', '月', '火', '水', '木', '金', '土'];

        // フィルタセグメント + リスト
        let headerHtml = `
            <div class="recruitment-header">
                <h3>📋 増員募集一覧</h3>
                <p class="recruitment-count">${futureEvents.length}件の募集中</p>
            </div>
            <div class="filter-segment">
                <button class="filter-segment-btn ${recruitmentFilter === 'all' ? 'active' : ''}" onclick="window._tabsSetFilter('all')">
                    すべて <span class="filter-count">${futureEvents.length}</span>
                </button>
                <button class="filter-segment-btn ${recruitmentFilter === 'kept' ? 'active' : ''}" onclick="window._tabsSetFilter('kept')">
                    ★ キープ中 <span class="filter-count">${keptCount}</span>
                </button>
            </div>
        `;

        if (displayEvents.length === 0) {
            const emptyMsg = recruitmentFilter === 'kept'
                ? 'キープ中の募集はありません。<br>各募集カードの ☆ をタップしてキープできます。'
                : '現在、募集中のイベントはありません';
            container.innerHTML = headerHtml + `<div class="tab-empty" style="min-height:120px;">${emptyMsg}</div>`;
            return;
        }

        container.innerHTML = headerHtml + `
            <div class="recruitment-list">
                ${displayEvents.map(ev => {
                    const start = new Date(ev.start);
                    const month = start.getMonth() + 1;
                    const day = start.getDate();
                    const dayOfWeek = weekdays[start.getDay()];
                    const diff = Math.ceil((start - now) / (1000 * 60 * 60 * 24));
                    const daysLeft = diff > 0 ? `あと${diff}日` : '本日';
                    const hall = ev.hall || ev.title || '';
                    const eventName = ev.eventName || '';
                    const timeSlot = ev.timeSlot || '';

                    // セクションドットの色
                    let dotClass = 'multiple';
                    if (ev.sections) {
                        const active = [];
                        if (ev.sections.stage > 0) active.push('stage');
                        if (ev.sections.sound > 0) active.push('sound');
                        if (ev.sections.lighting > 0) active.push('lighting');
                        if (active.length === 1) dotClass = active[0];
                    }

                    let recruitInfo = '';
                    if (ev.sections) {
                        const parts = [];
                        if (ev.sections.stage > 0) parts.push(`舞台 ${ev.sections.stage}名`);
                        if (ev.sections.sound > 0) parts.push(`音響 ${ev.sections.sound}名`);
                        if (ev.sections.lighting > 0) parts.push(`照明 ${ev.sections.lighting}名`);
                        recruitInfo = parts.join('、');
                    }

                    // キープ状態
                    const eventId = generateEventId(ev);
                    const kept = isKept(eventId);

                    return `
                        <div class="upcoming-card" data-event-id="${eventId}">
                            <button class="keep-btn ${kept ? 'active' : ''}" data-event-id="${eventId}" onclick="event.stopPropagation(); window._tabsToggleKeep(this, '${eventId}');" aria-label="キープ">
                                ${kept ? '★' : '☆'}
                            </button>
                            <div class="upcoming-card-date" onclick="window._tabsOpenDrawer('${ev.start}', '${eventId}')">
                                <span class="upcoming-date-month">${month}/${day}</span>
                                <span class="upcoming-date-day">（${dayOfWeek}）</span>
                            </div>
                            <div class="upcoming-card-info" onclick="window._tabsOpenDrawer('${ev.start}', '${eventId}')">
                                <div class="upcoming-card-hall">
                                    <span class="legend-dot ${dotClass}" style="margin-right:6px;"></span>
                                    ${hall}
                                </div>
                                ${eventName ? `<div class="upcoming-card-event">${eventName}</div>` : ''}
                                <div class="upcoming-card-meta">
                                    ${timeSlot ? `<span>⏰ ${timeSlot}</span>` : ''}
                                    ${recruitInfo ? `<span>👥 募集：${recruitInfo}</span>` : ''}
                                </div>
                            </div>
                            <div class="upcoming-card-badge" onclick="window._tabsOpenDrawer('${ev.start}', '${eventId}')">
                                <span class="badge-recruiting">募集中</span>
                                <span class="badge-days">${daysLeft}</span>
                            </div>
                        </div>
                    `;
                }).join('')}
            </div>
        `;
    }

    /** フィルタ切り替え */
    window._tabsSetFilter = function (filter) {
        recruitmentFilter = filter;
        renderRecruitmentList();
    };

    // ============================================
    // マイページタブ
    // 【要件2】申込履歴セクション追加
    // ============================================
    function renderMyPage() {
        const container = document.getElementById('myPageContent');
        if (!container) return;

        const name = window.loggedInStaffName || localStorage.getItem('zouin_staff_display_name') || '---';
        const section = localStorage.getItem('zouin_staff_section') || '---';
        const area = localStorage.getItem('zouin_staff_preferred_area') || '---';
        const hall = localStorage.getItem('zouin_staff_hall') || '---';
        const isAdmin = window.isAdmin || localStorage.getItem('zouin_is_admin') === 'true';

        // 申込履歴のレンダリング（ダミーデータ or 将来API対応）
        const applicationHistoryHtml = renderApplicationHistory();

        container.innerHTML = `
            <div class="mypage-header">
                <div class="mypage-avatar">
                    <span>${name.charAt(0) || '?'}</span>
                </div>
                <h3 class="mypage-name">${name}</h3>
                ${isAdmin ? '<span class="mypage-admin-badge">管理者権限あり</span>' : ''}
            </div>
            <div class="mypage-info-list">
                <div class="mypage-info-item">
                    <span class="mypage-info-label">👤 お名前</span>
                    <span class="mypage-info-value">${name}</span>
                </div>
                <div class="mypage-info-item">
                    <span class="mypage-info-label">🏛 所属ホール</span>
                    <span class="mypage-info-value">${hall}</span>
                </div>
                <div class="mypage-info-item">
                    <span class="mypage-info-label">🔧 セクション</span>
                    <span class="mypage-info-value">${section}</span>
                </div>
                <div class="mypage-info-item">
                    <span class="mypage-info-label">📍 希望エリア</span>
                    <span class="mypage-info-value">${area}</span>
                </div>
            </div>
            ${applicationHistoryHtml}
            <button class="mypage-logout-btn" onclick="if(confirm('ログアウトしますか？')){localStorage.clear();location.reload();}">
                ログアウト
            </button>
        `;
    }

    /**
     * 申込履歴セクションを描画
     * 現在はダミーデータで表示。将来APIから取得する場合は
     * renderApplicationHistory(apiData) のように引数を渡す設計。
     */
    function renderApplicationHistory(data) {
        // ダミーデータ（将来的にはGAS APIから取得）
        const dummyData = [
            { date: '2026-08-10', hall: '東京国際フォーラム', event: '夏のオーケストラコンサート', status: 'confirmed' },
            { date: '2026-08-15', hall: 'NHKホール', event: '歌謡祭リハーサル', status: 'pending' },
            { date: '2026-07-20', hall: 'サントリーホール', event: 'ジャズフェスティバル', status: 'cancelled' },
        ];

        const items = data || dummyData;
        const weekdays = ['日', '月', '火', '水', '木', '金', '土'];

        const statusLabels = {
            pending: '確認中',
            confirmed: '確定',
            cancelled: 'キャンセル'
        };

        if (!items || items.length === 0) {
            return `
                <h4 class="mypage-section-title">📋 申込履歴</h4>
                <div class="application-history-empty">
                    まだ申し込みの履歴はありません
                </div>
            `;
        }

        return `
            <h4 class="mypage-section-title">📋 申込履歴</h4>
            <div class="application-history-list">
                ${items.map(item => {
                    const d = new Date(item.date);
                    const month = d.getMonth() + 1;
                    const day = d.getDate();
                    const wd = weekdays[d.getDay()];
                    const statusClass = `status-${item.status}`;
                    const statusLabel = statusLabels[item.status] || item.status;

                    return `
                        <div class="application-history-item">
                            <div class="application-history-date">
                                <span class="date-num">${month}/${day}</span>
                                <span class="date-wd">(${wd})</span>
                            </div>
                            <div class="application-history-info">
                                <div class="application-history-hall">${item.hall}</div>
                                <div class="application-history-event">${item.event}</div>
                            </div>
                            <span class="application-status-badge ${statusClass}">${statusLabel}</span>
                        </div>
                    `;
                }).join('')}
            </div>
        `;
    }

})();
