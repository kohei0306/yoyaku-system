// ===== カレンダーIDを設定 =====
const CALENDAR_ID = '3b5869895ddbcdb12fb3ea0fd0e416f259c2964a991c1b7eb329bd203a63409e@group.calendar.google.com';

// ===== フォーム送信時の処理 =====
function onFormSubmit(e) {
  const values = e.values;
  const timestamp = values[0];
  const name      = values[1];
  const phone     = values[2];
  const email     = values[3];
  const date      = values[4];
  const time      = values[5];
  const memo      = values[6] || 'なし';

  // ===== 重複チェック =====
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  let isDuplicate = false;

  for (let i = 2; i < lastRow; i++) {
    const existingDate = sheet.getRange(i, 5).getValue();
    const existingTime = sheet.getRange(i, 6).getValue();

    const existingDateStr = existingDate instanceof Date
      ? Utilities.formatDate(existingDate, 'Asia/Tokyo', 'yyyy/MM/dd')
      : existingDate.toString();

    if (existingDateStr === date && existingTime === time) {
      isDuplicate = true;
      break;
    }
  }

  // ===== 重複あり =====
  if (isDuplicate) {
    const dupMessage =
      '⚠️ 予約が重複しています！\n\n' +
      '👤 お名前：' + name + '\n' +
      '📞 電話番号：' + phone + '\n' +
      '📆 希望日：' + date + '\n' +
      '🕐 希望時間：' + time + '\n\n' +
      '❌ この時間帯はすでに予約が入っています。\n' +
      'お客様に別の時間帯をご案内ください。';
    sendLineMessagePush(dupMessage);
    return; // ← ここで処理を止める！
  }

  // ===== カレンダーに予約を追加 =====
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const dateParts = date.split('/');
  const hour = parseInt(time.replace('時〜', ''));
  const startTime = new Date(dateParts[0], dateParts[1] - 1, dateParts[2], hour, 0, 0);
  const endTime = new Date(dateParts[0], dateParts[1] - 1, dateParts[2], hour + 1, 0, 0);

  calendar.createEvent(
    '予約：' + name + ' 様',
    startTime,
    endTime,
    { description: '📞 ' + phone + '\n📧 ' + email + '\n📝 ' + memo }
  );

  // ===== お客さんへ自動返信メール =====
  const subject = '【予約完了】お問い合わせありがとうございます';
  const body =
    name + ' 様\n\n' +
    'この度はお問い合わせいただきありがとうございます。\n' +
    '以下の内容で予約を受け付けました。\n\n' +
    '─────────────────\n' +
    '希望日：' + date + '\n' +
    '希望時間：' + time + '\n' +
    'ご要望：' + memo + '\n' +
    '─────────────────\n\n' +
    '担当者より改めてご連絡いたします。\n' +
    'しばらくお待ちください。\n\n' +
    '─────────────────\n' +
    'アサヒ屋\n' +
    'Email：asahiya.kk@gmail.com\n' +
    '─────────────────';

  GmailApp.sendEmail(
    email,
    subject,
    body,
    { name: 'アサヒ屋' }
  );

  // ===== LINE通知 =====
  const message =
    '📅 新しい予約が入りました！\n\n' +
    '👤 お名前：' + name + '\n' +
    '📞 電話番号：' + phone + '\n' +
    '📧 メール：' + email + '\n' +
    '📆 希望日：' + date + '\n' +
    '🕐 希望時間：' + time + '\n' +
    '📝 ご要望：' + memo + '\n\n' +
    '⏰ 受付時刻：' + timestamp + '\n' +
    '📆 カレンダーに追加しました！';

  sendLineMessagePush(message);
}

// ===== LINE送信関数 =====
function sendLineMessagePush(message) {
  const props = PropertiesService.getScriptProperties();
  const token  = props.getProperty('LINE_TOKEN');
  const userId = props.getProperty('LINE_USER_ID');

  const url = 'https://api.line.me/v2/bot/message/push';
  const payload = { to: userId, messages: [{ type: 'text', text: message }] };
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  UrlFetchApp.fetch(url, options);
}

// ===== テスト用 =====
function testReservationNotify() {
  sendLineMessagePush(
    '📅 新しい予約が入りました！\n\n' +
    '👤 お名前：テスト太郎\n' +
    '📞 電話番号：090-0000-0000\n' +
    '📧 メール：test@example.com\n' +
    '📆 希望日：2026/04/15\n' +
    '🕐 希望時間：14時〜\n' +
    '📝 ご要望：なし\n\n' +
    '⏰ 受付時刻：2026/04/09 10:30:00\n' +
    '📆 カレンダーに追加しました！'
  );
}

// ===== Webアプリ用：予約データ取得 =====
function getBookings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const bookings = [];

  if (lastRow < 2) return bookings;

  for (let i = 2; i <= lastRow; i++) {
    const dateVal = sheet.getRange(i, 5).getValue();
    const dateStr = dateVal instanceof Date
      ? Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy/MM/dd')
      : dateVal.toString();

    bookings.push({
      date:  dateStr,
      name:  sheet.getRange(i, 2).getValue(),
      phone: sheet.getRange(i, 3).getValue(),
      email: sheet.getRange(i, 4).getValue(),
      time:  sheet.getRange(i, 6).getValue(),
      memo:  sheet.getRange(i, 7).getValue()
    });
  }

  return bookings;
}

// ===== Webアプリのエントリーポイント =====
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('予約状況カレンダー')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}