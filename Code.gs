function doGet(e) {
  const data = {
    stamps: getSheetData('stamps'),
    qrCodes: getSheetData('qr_codes'),
    participants: getSheetData('participants')
  };
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = (() => {
    try { return JSON.parse(e.postData.contents || '{}'); } catch { return {}; }
  })();
  appendHistory(body);
  return ContentService
    .createTextOutput(JSON.stringify({ result: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheetData(sheetName) {
  const ss = SpreadsheetApp.openById('143tDLYxxs29hir9Me2UGXyL70yIlR5WFcTZr9yHhtmY');
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values.shift().map(h => String(h).trim());
  return values.map(row => {
    const item = {};
    row.forEach((cell, i) => {
      const key = headers[i] || `col_${i}`;
      item[key] = cell;
    });
    return item;
  });
}

function appendHistory(body) {
  const ss = SpreadsheetApp.openById('143tDLYxxs29hir9Me2UGXyL70yIlR5WFcTZr9yHhtmY');
  let sheet = ss.getSheetByName('history');
  if (!sheet) {
    // history シートが無ければ作る（初回のみ）
    sheet = ss.insertSheet('history');
    sheet.appendRow(['timestamp','participantId','bookCode','qrId','stampId','note']);
  }
  sheet.appendRow([
    new Date(),
    body.participantId || '',
    body.bookCode || '',
    body.qrId || '',
    body.stampId || '',
    body.note || ''
  ]);
}