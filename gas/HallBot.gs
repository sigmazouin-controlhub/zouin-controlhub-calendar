/**
 * 増員 CONTROL HUB 2.0 - ホール案内Bot (HallBot.gs)
 * リッチメニューからカルーセルUIでホール情報を段階的に案内
 */

// ============================================
// メインハンドラ
// ============================================

/**
 * ホール案内Botのメインルーター
 * doPostから呼ばれる。テキスト「ホール案内」で起動
 * @param {object} event - LINE Webhookイベント
 */
function handleHallBotText(event) {
  try {
    sendHallListCarousel(event.replyToken);
  } catch (error) {
    Logger.log('handleHallBotText エラー: ' + error.toString());
    replyTextMessage(event.replyToken, 'エラーが発生しました。もう一度お試しください。');
  }
}

/**
 * ホール案内Botのpostbackハンドラ
 * @param {object} event - LINE Webhookイベント（postback）
 */
function handleHallBotPostback(event) {
  try {
    var data = parsePostbackData(event.postback.data);
    var action = data.action || '';
    
    if (action === 'selectHall') {
      // ホール選択 → 情報種類カルーセル
      sendInfoTypeCarousel(event.replyToken, data.hall);
    } else if (action === 'hallInfo') {
      // 情報種類選択 → 詳細表示
      sendHallDetail(event.replyToken, data.hall, data.type);
    }
  } catch (error) {
    Logger.log('handleHallBotPostback エラー: ' + error.toString());
    replyTextMessage(event.replyToken, 'エラーが発生しました。もう一度お試しください。');
  }
}

// ============================================
// postbackデータパーサー
// ============================================

/**
 * postbackデータ文字列をオブジェクトに変換
 * @param {string} dataStr - "action=selectHall&hall=ホール名" 形式
 * @return {object} パース済みオブジェクト
 */
function parsePostbackData(dataStr) {
  var result = {};
  if (!dataStr) return result;
  var pairs = dataStr.split('&');
  for (var i = 0; i < pairs.length; i++) {
    var kv = pairs[i].split('=');
    if (kv.length === 2) {
      result[decodeURIComponent(kv[0])] = decodeURIComponent(kv[1]);
    }
  }
  return result;
}

// ============================================
// Step 2: ホール一覧カルーセル
// ============================================

/**
 * ホール一覧をカルーセルで送信
 * 「ホール名＆入館」タブA列からホール名を取得
 * @param {string} replyToken
 */
function sendHallListCarousel(replyToken) {
  var hallNames = getHallList();
  if (hallNames.length === 0) {
    replyTextMessage(replyToken, '登録されているホールがありません。');
    return;
  }
  
  // カルーセルのバブル（最大12件）
  var bubbles = [];
  var limit = Math.min(hallNames.length, 12);
  
  // テーマカラー
  var colors = ['#1a5276', '#1a6b4a', '#7d3c98', '#c0392b', '#2471a3', '#148f77', '#884ea0', '#cb4335', '#2e86c1', '#17a589', '#6c3483', '#e74c3c'];
  
  for (var i = 0; i < limit; i++) {
    var hallName = hallNames[i];
    var color = colors[i % colors.length];
    
    bubbles.push({
      type: 'bubble',
      size: 'kilo',
      header: {
        type: 'box',
        layout: 'vertical',
        contents: [
          {
            type: 'box',
            layout: 'vertical',
            contents: [
              {
                type: 'text',
                text: '🏛',
                size: 'xxl',
                align: 'center'
              }
            ],
            paddingTop: '15px',
            paddingBottom: '10px'
          }
        ],
        backgroundColor: color,
        paddingAll: '0px'
      },
      body: {
        type: 'box',
        layout: 'vertical',
        contents: [
          {
            type: 'text',
            text: hallName,
            weight: 'bold',
            size: 'md',
            align: 'center',
            wrap: true,
            maxLines: 2
          }
        ],
        spacing: 'sm',
        paddingAll: '15px',
        justifyContent: 'center',
        alignItems: 'center'
      },
      footer: {
        type: 'box',
        layout: 'vertical',
        contents: [
          {
            type: 'button',
            action: {
              type: 'postback',
              label: '選択する',
              data: 'action=selectHall&hall=' + encodeURIComponent(hallName),
              displayText: hallName + ' を選択'
            },
            style: 'primary',
            color: color,
            height: 'sm'
          }
        ],
        paddingAll: '10px'
      }
    });
  }
  
  var flexMessage = {
    type: 'flex',
    altText: 'ホール案内 - ホールを選択してください',
    contents: {
      type: 'carousel',
      contents: bubbles
    }
  };
  
  replyMessage(replyToken, [
    { type: 'text', text: '📋 知りたいホールの情報を選択してください\n（横にスクロールできます）' },
    flexMessage
  ]);
}

// ============================================
// Step 3: 情報種類カルーセル
// ============================================

/**
 * 情報種類のカルーセルを送信
 * @param {string} replyToken
 * @param {string} hallName
 */
function sendInfoTypeCarousel(replyToken, hallName) {
  var infoTypes = [
    { type: 'entry',   emoji: '🏢', label: '入館方法',  color: '#2e86c1', desc: '入館方法の画像を表示' },
    { type: 'nearby',  emoji: '🏪', label: '周辺情報',  color: '#27ae60', desc: '周辺施設の情報を表示' },
    { type: 'contact', emoji: '📞', label: '連絡先',    color: '#8e44ad', desc: '事務所の連絡先を表示' },
    { type: 'smoking', emoji: '🚬', label: '喫煙所',    color: '#e67e22', desc: '喫煙所の情報を表示' }
  ];
  
  var bubbles = [];
  for (var i = 0; i < infoTypes.length; i++) {
    var info = infoTypes[i];
    bubbles.push({
      type: 'bubble',
      size: 'kilo',
      header: {
        type: 'box',
        layout: 'vertical',
        contents: [
          {
            type: 'text',
            text: info.emoji,
            size: 'xxl',
            align: 'center'
          }
        ],
        backgroundColor: info.color,
        paddingTop: '20px',
        paddingBottom: '15px'
      },
      body: {
        type: 'box',
        layout: 'vertical',
        contents: [
          {
            type: 'text',
            text: info.label,
            weight: 'bold',
            size: 'lg',
            align: 'center'
          },
          {
            type: 'text',
            text: info.desc,
            size: 'xs',
            color: '#888888',
            align: 'center',
            wrap: true,
            margin: 'sm'
          }
        ],
        spacing: 'sm',
        paddingAll: '15px'
      },
      footer: {
        type: 'box',
        layout: 'vertical',
        contents: [
          {
            type: 'button',
            action: {
              type: 'postback',
              label: '選択する',
              data: 'action=hallInfo&hall=' + encodeURIComponent(hallName) + '&type=' + info.type,
              displayText: hallName + ' の' + info.label
            },
            style: 'primary',
            color: info.color,
            height: 'sm'
          }
        ],
        paddingAll: '10px'
      }
    });
  }
  
  var flexMessage = {
    type: 'flex',
    altText: hallName + ' - 知りたい情報を選択',
    contents: {
      type: 'carousel',
      contents: bubbles
    }
  };
  
  replyMessage(replyToken, [
    { type: 'text', text: '🏛 ' + hallName + '\n知りたい情報を選択してください' },
    flexMessage
  ]);
}

// ============================================
// Step 4-5: 詳細情報表示
// ============================================

/**
 * ホール詳細情報を送信
 * @param {string} replyToken
 * @param {string} hallName
 * @param {string} infoType - 'entry' | 'nearby' | 'contact' | 'smoking'
 */
function sendHallDetail(replyToken, hallName, infoType) {
  switch (infoType) {
    case 'entry':
      sendEntryInfo(replyToken, hallName);
      break;
    case 'nearby':
      sendNearbyInfo(replyToken, hallName);
      break;
    case 'contact':
      sendContactInfo(replyToken, hallName);
      break;
    case 'smoking':
      sendSmokingInfo(replyToken, hallName);
      break;
    default:
      replyTextMessage(replyToken, '不明な情報タイプです。');
  }
}

/**
 * 入館方法を送信（画像）
 */
function sendEntryInfo(replyToken, hallName) {
  var entryInfo = getHallEntryInfo(hallName);
  if (!entryInfo || !entryInfo.imageUrl) {
    replyTextMessage(replyToken, '🏢 ' + hallName + ' の入館方法\n\n入館方法の画像はまだ登録されていません。');
    return;
  }
  
  var imageUrl = convertDriveUrl(entryInfo.imageUrl);
  
  replyMessage(replyToken, [
    { type: 'text', text: '🏢 ' + hallName + ' の入館方法' },
    {
      type: 'image',
      originalContentUrl: imageUrl,
      previewImageUrl: imageUrl
    }
  ]);
}

/**
 * 周辺情報を送信（テキスト Flexメッセージ）
 */
function sendNearbyInfo(replyToken, hallName) {
  var detail = getHallDetailInfo(hallName);
  if (!detail) {
    replyTextMessage(replyToken, hallName + ' の情報が見つかりません。');
    return;
  }
  
  if (!detail.nearby1 && !detail.nearby2) {
    replyTextMessage(replyToken, '🏪 ' + hallName + ' の周辺情報\n\n周辺情報はまだ登録されていません。');
    return;
  }
  
  var contentItems = [];
  
  // 周辺情報①
  if (detail.nearby1) {
    contentItems.push(
      {
        type: 'box',
        layout: 'vertical',
        contents: [
          { type: 'text', text: '📍 周辺情報①', weight: 'bold', size: 'sm', color: '#27ae60' },
          { type: 'text', text: detail.nearby1, size: 'sm', wrap: true, margin: 'sm', color: '#333333' }
        ],
        margin: 'lg'
      }
    );
  }
  
  // 周辺情報②
  if (detail.nearby2) {
    contentItems.push(
      {
        type: 'box',
        layout: 'vertical',
        contents: [
          { type: 'text', text: '📍 周辺情報②', weight: 'bold', size: 'sm', color: '#27ae60' },
          { type: 'text', text: detail.nearby2, size: 'sm', wrap: true, margin: 'sm', color: '#333333' }
        ],
        margin: 'lg'
      }
    );
  }
  
  var flexBubble = {
    type: 'bubble',
    header: {
      type: 'box',
      layout: 'vertical',
      contents: [
        { type: 'text', text: '🏪 ' + hallName, weight: 'bold', size: 'lg', color: '#ffffff', wrap: true }
      ],
      backgroundColor: '#27ae60',
      paddingAll: '15px'
    },
    body: {
      type: 'box',
      layout: 'vertical',
      contents: contentItems,
      paddingAll: '15px'
    }
  };
  
  // 住所があればGoogle Mapsボタンを追加
  if (detail.address) {
    flexBubble.footer = {
      type: 'box',
      layout: 'vertical',
      contents: [
        {
          type: 'button',
          action: {
            type: 'uri',
            label: '🗺 Google Mapsで見る',
            uri: 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(detail.address)
          },
          style: 'secondary',
          height: 'sm'
        }
      ],
      paddingAll: '10px'
    };
  }
  
  replyMessage(replyToken, [{
    type: 'flex',
    altText: hallName + ' の周辺情報',
    contents: flexBubble
  }]);
}

/**
 * 連絡先を送信（Flexメッセージ）
 */
function sendContactInfo(replyToken, hallName) {
  var detail = getHallDetailInfo(hallName);
  if (!detail) {
    replyTextMessage(replyToken, hallName + ' の情報が見つかりません。');
    return;
  }
  
  var contentItems = [];
  
  // 住所
  if (detail.address) {
    contentItems.push(
      {
        type: 'box',
        layout: 'vertical',
        contents: [
          { type: 'text', text: '📍 住所', weight: 'bold', size: 'sm', color: '#1a5276' },
          { type: 'text', text: detail.address, size: 'sm', wrap: true, margin: 'sm', color: '#333333' }
        ],
        margin: 'lg'
      }
    );
  }
  
  // 電話番号
  if (detail.phone) {
    contentItems.push(
      {
        type: 'box',
        layout: 'vertical',
        contents: [
          { type: 'text', text: '📞 舞台事務所直通', weight: 'bold', size: 'sm', color: '#1a5276' },
          { type: 'text', text: detail.phone, size: 'sm', margin: 'sm', color: '#333333' }
        ],
        margin: 'lg'
      }
    );
  }
  
  // メールアドレス
  if (detail.email) {
    contentItems.push(
      {
        type: 'box',
        layout: 'vertical',
        contents: [
          { type: 'text', text: '✉️ メールアドレス', weight: 'bold', size: 'sm', color: '#1a5276' },
          { type: 'text', text: detail.email, size: 'sm', margin: 'sm', wrap: true, color: '#333333' }
        ],
        margin: 'lg'
      }
    );
  }
  
  if (contentItems.length === 0) {
    replyTextMessage(replyToken, '📞 ' + hallName + ' の連絡先\n\n連絡先情報はまだ登録されていません。');
    return;
  }
  
  var footerContents = [];
  // 電話ボタン
  if (detail.phone) {
    // 電話番号からハイフン等を除去
    var telNumber = detail.phone.replace(/[^0-9]/g, '');
    if (telNumber) {
      footerContents.push({
        type: 'button',
        action: {
          type: 'uri',
          label: '📞 電話をかける',
          uri: 'tel:' + telNumber
        },
        style: 'primary',
        color: '#2e86c1',
        height: 'sm'
      });
    }
  }
  // Google Maps ボタン
  if (detail.address) {
    footerContents.push({
      type: 'button',
      action: {
        type: 'uri',
        label: '🗺 Google Mapsで見る',
        uri: 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(detail.address)
      },
      style: 'secondary',
      height: 'sm'
    });
  }
  
  var flexBubble = {
    type: 'bubble',
    header: {
      type: 'box',
      layout: 'vertical',
      contents: [
        { type: 'text', text: '📞 ' + hallName, weight: 'bold', size: 'lg', color: '#ffffff', wrap: true }
      ],
      backgroundColor: '#1a5276',
      paddingAll: '15px'
    },
    body: {
      type: 'box',
      layout: 'vertical',
      contents: contentItems,
      paddingAll: '15px'
    }
  };
  
  if (footerContents.length > 0) {
    flexBubble.footer = {
      type: 'box',
      layout: 'vertical',
      contents: footerContents,
      spacing: 'sm',
      paddingAll: '10px'
    };
  }
  
  replyMessage(replyToken, [{
    type: 'flex',
    altText: hallName + ' の連絡先',
    contents: flexBubble
  }]);
}

/**
 * 喫煙所情報を送信
 */
function sendSmokingInfo(replyToken, hallName) {
  var detail = getHallDetailInfo(hallName);
  if (!detail) {
    replyTextMessage(replyToken, hallName + ' の情報が見つかりません。');
    return;
  }
  
  if (!detail.smoking) {
    replyTextMessage(replyToken, '🚬 ' + hallName + ' の喫煙所\n\n喫煙所の情報はまだ登録されていません。');
    return;
  }
  
  var flexBubble = {
    type: 'bubble',
    header: {
      type: 'box',
      layout: 'vertical',
      contents: [
        { type: 'text', text: '🚬 ' + hallName, weight: 'bold', size: 'lg', color: '#ffffff', wrap: true }
      ],
      backgroundColor: '#e67e22',
      paddingAll: '15px'
    },
    body: {
      type: 'box',
      layout: 'vertical',
      contents: [
        {
          type: 'text',
          text: '喫煙所について',
          weight: 'bold',
          size: 'md',
          color: '#e67e22'
        },
        {
          type: 'text',
          text: detail.smoking,
          size: 'sm',
          wrap: true,
          margin: 'md',
          color: '#333333'
        }
      ],
      paddingAll: '15px'
    }
  };
  
  replyMessage(replyToken, [{
    type: 'flex',
    altText: hallName + ' の喫煙所情報',
    contents: flexBubble
  }]);
}

// ============================================
// データ取得ヘルパー
// ============================================

/**
 * ホール名一覧を取得（ホール周辺情報の回答 B列）
 * 周辺情報が登録されているホールのみ表示
 * @return {string[]} ホール名の配列（重複除去済み）
 */
function getHallList() {
  var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('ホール周辺情報の回答');
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  var data = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // B列（ホール名）
  var hallNames = [];
  for (var i = 0; i < data.length; i++) {
    var name = (data[i][0] || '').toString().trim();
    if (name && hallNames.indexOf(name) === -1) {
      hallNames.push(name);
    }
  }
  return hallNames;
}

/**
 * ホール入館情報を取得（ホール名＆入館）
 * @param {string} hallName
 * @return {object|null} { hallName, imageUrl }
 */
function getHallEntryInfo(hallName) {
  var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('ホール名＆入館');
  if (!sheet) return null;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // A-B列
  for (var i = 0; i < data.length; i++) {
    var name = (data[i][0] || '').toString().trim();
    if (name === hallName) {
      return {
        hallName: name,
        imageUrl: (data[i][1] || '').toString().trim()
      };
    }
  }
  return null;
}

/**
 * ホール詳細情報を取得（ホール周辺情報の回答）
 * @param {string} hallName
 * @return {object|null} { hallName, address, phone, email, nearby1, nearby2, smoking }
 */
function getHallDetailInfo(hallName) {
  var ss = SpreadsheetApp.openById(STAFF_CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('ホール周辺情報の回答');
  if (!sheet) return null;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  
  // A〜I列を取得（9列）
  var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  
  // 同じホール名が複数ある場合は最新（最後）を使用
  var result = null;
  for (var i = 0; i < data.length; i++) {
    var name = (data[i][1] || '').toString().trim(); // B列=ホール名
    if (name === hallName) {
      result = {
        hallName: name,
        address: (data[i][2] || '').toString().trim(),   // C列
        phone: (data[i][3] || '').toString().trim(),     // D列
        email: (data[i][4] || '').toString().trim(),     // E列
        nearby1: (data[i][6] || '').toString().trim(),   // G列
        nearby2: (data[i][7] || '').toString().trim(),   // H列
        smoking: (data[i][8] || '').toString().trim()    // I列
      };
    }
  }
  return result;
}

// ============================================
// Google Drive URL変換
// ============================================

/**
 * Google Drive共有URLを直接アクセス可能なURLに変換
 * LINE APIで画像送信するにはダイレクトリンクが必要
 * @param {string} url - Drive共有URLまたは直接URL
 * @return {string} 変換後のURL
 */
function convertDriveUrl(url) {
  if (!url) return '';
  
  // 既にダイレクトリンクならそのまま返す
  if (url.indexOf('lh3.googleusercontent.com') !== -1) return url;
  if (url.indexOf('drive.google.com/uc') !== -1) return url;
  
  // Google Drive共有リンクからファイルIDを抽出
  var fileId = '';
  
  // パターン1: https://drive.google.com/file/d/FILE_ID/view...
  var match1 = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (match1) {
    fileId = match1[1];
  }
  
  // パターン2: https://drive.google.com/open?id=FILE_ID
  if (!fileId) {
    var match2 = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (match2) {
      fileId = match2[1];
    }
  }
  
  // パターン3: ファイルIDのみ
  if (!fileId && /^[a-zA-Z0-9_-]{10,}$/.test(url.trim())) {
    fileId = url.trim();
  }
  
  if (fileId) {
    // lh3形式が最も安定（リダイレクトなし）
    return 'https://lh3.googleusercontent.com/d/' + fileId;
  }
  
  // 変換できない場合はそのまま返す
  return url;
}

// ============================================
// LINE メッセージ送信ヘルパー
// ============================================

/**
 * テキストメッセージを返信
 * @param {string} replyToken
 * @param {string} text
 */
function replyTextMessage(replyToken, text) {
  replyMessage(replyToken, [{ type: 'text', text: text }]);
}

/**
 * 複数メッセージを返信（汎用）
 * @param {string} replyToken
 * @param {Array} messages - メッセージオブジェクトの配列（最大5件）
 */
function replyMessage(replyToken, messages) {
  var url = 'https://api.line.me/v2/bot/message/reply';
  var options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_CONFIG.CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: messages.slice(0, 5) // 最大5メッセージ
    }),
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    Logger.log('replyMessage エラー: ' + response.getContentText());
  }
}

// ============================================
// テスト関数
// ============================================

/**
 * ホール一覧カルーセルのJSON確認用
 */
function testHallListCarousel() {
  var hallNames = getHallList();
  Logger.log('ホール一覧: ' + JSON.stringify(hallNames));
  Logger.log('件数: ' + hallNames.length);
}

/**
 * ホール詳細情報の確認用
 */
function testHallDetail() {
  var hallNames = getHallList();
  if (hallNames.length === 0) {
    Logger.log('ホールが登録されていません');
    return;
  }
  var testHall = hallNames[0];
  Logger.log('テスト対象: ' + testHall);
  
  var entryInfo = getHallEntryInfo(testHall);
  Logger.log('入館情報: ' + JSON.stringify(entryInfo));
  
  var detailInfo = getHallDetailInfo(testHall);
  Logger.log('詳細情報: ' + JSON.stringify(detailInfo));
}

/**
 * Drive URL変換テスト
 */
function testDriveUrlConvert() {
  var testUrls = [
    'https://drive.google.com/file/d/1abcDEFghiJKLmno/view?usp=sharing',
    'https://drive.google.com/open?id=1abcDEFghiJKLmno',
    '1abcDEFghiJKLmno',
    'https://lh3.googleusercontent.com/d/1abcDEFghiJKLmno'
  ];
  for (var i = 0; i < testUrls.length; i++) {
    Logger.log(testUrls[i] + ' → ' + convertDriveUrl(testUrls[i]));
  }
}
