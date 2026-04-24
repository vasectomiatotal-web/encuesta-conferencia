// GOOGLE APPS SCRIPT - Encuesta Conferencia (Optimizado)
const SPREADSHEET_NAME = 'EncuestaConferencia';
const SHEET_NAME = 'Conteo';

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'getStats') return getStats();
  if (action === 'registrar') return registrar(e.parameter.grupo);
  if (action === 'reset') return resetConteo();
  return ContentService.createTextOutput(JSON.stringify({error: 'Invalid action'})).setMimeType(ContentService.MimeType.JSON);
}

function getStats() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const counts = [
      sheet.getRange('B2').getValue() || 0,
      sheet.getRange('B3').getValue() || 0,
      sheet.getRange('B4').getValue() || 0
    ];
    return ContentService.createTextOutput(JSON.stringify({counts: counts})).setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({counts: [0,0,0], error: e.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function registrar(grupo) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const fila = parseInt(grupo);
    const celda = sheet.getRange(fila, 2);
    const actual = celda.getValue() || 0;
    celda.setValue(actual + 1);
    return ContentService.createTextOutput(JSON.stringify({success: true, count: actual + 1})).setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({success: false, error: e.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function resetConteo() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    sheet.getRange('B2:B4').setValues([[0], [0], [0]]);
    return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({success: false, error: e.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function getSpreadsheet() {
  const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (ssId) {
    try { return SpreadsheetApp.openById(ssId); } catch (e) {}
  }
  
  const files = DriveApp.getFilesByName(SPREADSHEET_NAME);
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType() === 'application/vnd.google-apps.spreadsheet') {
      const ss = SpreadsheetApp.open(f);
      PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
      return ss;
    }
  }
  
  const ss = SpreadsheetApp.create(SPREADSHEET_NAME);
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  const sheet = ss.getSheetByName(SHEET_NAME);
  sheet.getRange('A1:C1').setValues([['Grupo', 'Conteo', 'Ultima']]);
  sheet.getRange('A2:A4').setValues([['Grupo 1'], ['Grupo 2'], ['Grupo 3']]);
  sheet.getRange('B2:B4').setValues([[0], [0], [0]]);
  return ss;
}