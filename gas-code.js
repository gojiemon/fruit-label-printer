// ウッドベリーズ ラベルデータ共有 - Google Apps Script
// このコードをGoogle Sheetの拡張機能 → Apps Script に貼り付けてください

const SHEET_NAME = 'data';

function doGet(e) {
  try {
    const sheet = getOrCreateSheet();
    const cell = sheet.getRange('A1').getValue();
    const data = cell ? JSON.parse(cell) : { customLabels: [], overrides: {} };
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const sheet = getOrCreateSheet();
    const data = JSON.parse(e.postData.contents);
    sheet.getRange('A1').setValue(JSON.stringify(data));
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange('A1').setValue(JSON.stringify({ customLabels: [], overrides: {} }));
  }
  return sheet;
}
