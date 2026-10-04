const SHEET_NAME = 'Registrations';
const HEADERS = [
  'Дата и время регистрации',
  'Имя и фамилия',
  'Возраст',
  'Email',
  'Школа / учебное учреждение',
  'Номер телефона',
  'Название команды'
];

function doPost(event) {
  const params = event && event.parameter ? event.parameter : {};
  let stage = 'validate request';
  const requestId = /^[a-zA-Z0-9-]{16,80}$/.test(String(params.requestId || ''))
    ? String(params.requestId)
    : '';

  try {
    if (!requestId || params.website) throw new Error('Invalid submission');

    const name = readText(params.fullName, 120);
    const age = Number(params.age);
    const email = readText(params.email, 254);
    const school = readText(params.school, 160);
    const phone = readText(params.phone, 16);
    const team = readText(params.teamName, 120);
    if (!name || !school || !team || !Number.isInteger(age) || age < 10 || age > 100) {
      throw new Error('Invalid submission');
    }
    if (!/^[A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?(?:\.[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?)+$/i.test(email)) {
      throw new Error('Invalid submission');
    }
    if (!/^\+996 \d{3} \d{3} \d{3}$/.test(phone)) throw new Error('Invalid submission');

    const cache = CacheService.getScriptCache();
    const cacheKey = 'saved_' + requestId;
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      if (!cache.get(cacheKey)) {
        stage = 'read spreadsheet configuration';
        const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
        if (!spreadsheetId) throw new Error('Missing configuration');
        stage = 'open spreadsheet';
        const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        stage = 'open or create registrations sheet';
        const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
        if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
        stage = 'save registration row';
        sheet.appendRow([
          new Date(),
          safeCell(name),
          age,
          safeCell(email),
          safeCell(school),
          safeCell(phone),
          safeCell(team)
        ]);
        SpreadsheetApp.flush();
        cache.put(cacheKey, '1', 21600);
      }
    } finally {
      lock.releaseLock();
    }
    return responsePage(requestId, 'success');
  } catch (error) {
    // Только технический этап и тип ошибки: не записывать данные заявки или ID таблицы.
    console.error('Registration failed at stage: ' + stage + '; error type: ' + (error && error.name ? error.name : 'Error'));
    return responsePage(requestId, 'failure');
  }
}

function readText(value, maxLength) {
  const text = String(value || '').replace(/[\u0000-\u001f\u007f]/g, '').trim();
  if (text.length > maxLength) throw new Error('Invalid submission');
  return text;
}

function safeCell(value) {
  const text = String(value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function responsePage(requestId, status) {
  const message = JSON.stringify({
    type: 'umut-registration-result',
    requestId: requestId,
    status: status
  });
  const html = '<!doctype html><html><head><meta charset="utf-8"></head><body>' +
    '<script>window.top.postMessage(' + message + ', "*");</script>' +
    '</body></html>';
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
