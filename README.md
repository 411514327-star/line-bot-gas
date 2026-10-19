# // LINE bot on Google Apps Script：Gemini 判斷「待辦」或「提問」，並記錄到試算表
function doGet() {
  return ContentService.createTextOutput('ok');
}

function getProp(name) {
  return PropertiesService.getScriptProperties().getProperty(name);
}

// 回傳 {type, reply}；失敗回傳 null
function callGemini(userText) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent';
  const instruction =
    '你是 LINE 上的助理。判斷使用者這則訊息：' +
    '如果是要記下來的待辦、行程或提醒，type 填 "todo"，reply 寫一句簡短確認（不要寫列號）；' +
    '否則 type 填 "chat"，reply 用繁體中文簡短回答，300 字以內。' +
    '只輸出 JSON，格式：{"type":"todo或chat","reply":"..."}';
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': getProp('GEMINI_API_KEY') },
    payload: JSON.stringify({
      systemInstruction: { parts: [{ text: instruction }] },
      contents: [{ role: 'user', parts: [{ text: userText }] }],
      generationConfig: { responseMimeType: 'application/json' }
    }),
    muteHttpExceptions: true
  };

  for (let i = 0; i < 3; i++) {
    const res = UrlFetchApp.fetch(url, options);
    const code = res.getResponseCode();
    if (code === 200) {
      const data = JSON.parse(res.getContentText());
      const c = data.candidates && data.candidates[0];
      const parts = c && c.content && c.content.parts;
      if (!parts) return null;
      const raw = parts.map(function (p) { return p.text || ''; }).join('');
      try {
        return JSON.parse(raw);
      } catch (err) {
        return { type: 'chat', reply: raw };
      }
    }
    console.error(code + ' ' + res.getContentText());
    if (code !== 503 && code !== 429) return null;
    Utilities.sleep(1500);
  }
  return null;
}

// 處理一則文字訊息：呼叫 Gemini、寫試算表、回傳要回給使用者的字
function handleText(text) {
  const r = callGemini(text);
  if (!r) return '⚠️ 我現在答不出來，請稍後再試';

  const isTodo = r.type === 'todo';
  let reply = String(r.reply || '').slice(0, 4000);

  try {
    const sheet = SpreadsheetApp.openById(getProp('SHEET_ID')).getSheets()[0];
    const row = sheet.getLastRow() + 1;
    if (isTodo) reply = '✅ 記好了（試算表第 ' + row + ' 列）：' + reply;
    const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([now, isTodo ? '待辦' : '聊天', text, reply]);
  } catch (err) {
    console.error('寫試算表失敗：' + err);
    if (isTodo) reply = '⚠️ 待辦沒記進試算表：' + reply;
  }
  return reply;
}

function doPost(e) {
  const token = getProp('LINE_CHANNEL_ACCESS_TOKEN');
  const body = JSON.parse(e.postData.contents);

  (body.events || []).forEach(function (event) {
    if (event.type !== 'message' || !event.replyToken) return;

    const text = event.message.type === 'text'
      ? handleText(event.message.text)
      : '目前只支援文字訊息';

    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({
        replyToken: event.replyToken,
        messages: [{ type: 'text', text: text }]
      }),
      muteHttpExceptions: true
    });
  });

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 在編輯器手動測試用
function testHandle() {
  console.log(handleText('明天下午三點要開會'));
}
