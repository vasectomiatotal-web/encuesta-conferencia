// ============================================
// GOOGLE APPS SCRIPT - Encuesta Conferencia
// ============================================
// Instrucciones:
// 1. Ve a https://script.google.com/
// 2. Crea un nuevo proyecto
// 3. Pega este código
// 4. Ejecuta setup() una vez
// 5. Despliega como Web App
// 6. Copia el URL del Web App
// ============================================

// CONFIGURACIÓN - Cambia estos valores
const SPREADSHEET_NAME = 'EncuestaConferencia';
const SHEET_NAME = 'Conteo';

// No necesitas cambiar nada más
function doGet(e) {
  const action = e.parameter.action;
  
  if (action === 'getStats') {
    return getStats();
  } else if (action === 'registrar') {
    const grupo = e.parameter.grupo;
    return registrar(grupo);
  } else if (action === 'reset') {
    return resetConteo();
  }
  
  return ContentService.createTextOutput(JSON.stringify({error: 'Acción no válida'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function getStats() {
  try {
    const sheet = getSheet();
    const counts = [0, 0, 0];
    
    // Leer conteos (fila 1: conteo grupo 1, fila 2: grupo 2, fila 3: grupo 3)
    for (let i = 0; i < 3; i++) {
      counts[i] = sheet.getRange(i + 1, 2).getValue() || 0;
    }
    
    return ContentService.createTextOutput(JSON.stringify({counts: counts}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({counts: [0,0,0], error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function registrar(grupo) {
  try {
    const sheet = getSheet();
    const fila = parseInt(grupo);
    
    // Incrementar conteo
    const celda = sheet.getRange(fila, 2);
    const actual = celda.getValue() || 0;
    celda.setValue(actual + 1);
    
    // Registrar timestamp
    const timeCell = sheet.getRange(fila, 3);
    timeCell.setValue(new Date());
    
    return ContentService.createTextOutput(JSON.stringify({success: true, count: actual + 1}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({success: false, error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function resetConteo() {
  try {
    const sheet = getSheet();
    sheet.getRange('B2').setValue(0);
    sheet.getRange('B3').setValue(0);
    sheet.getRange('B4').setValue(0);
    sheet.getRange('C2').setValue('');
    sheet.getRange('C3').setValue('');
    sheet.getRange('C4').setValue('');
    
    return ContentService.createTextOutput(JSON.stringify({success: true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({success: false, error: e.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getSheet() {
  // Intentar abrir por ID guardado o buscar por nombre
  const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  let ss = null;
  
  if (ssId) {
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch (e) {
      // No se encontró
    }
  }
  
  // Buscar si existe uno con ese nombre
  if (!ss) {
    const files = DriveApp.getFilesByName(SPREADSHEET_NAME);
    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType() === 'application/vnd.google-apps.spreadsheet') {
        ss = SpreadsheetApp.open(file);
        break;
      }
    }
  }
  
  // Crear nuevo si no existe
  if (!ss) {
    ss = SpreadsheetApp.create(SPREADSHEET_NAME);
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  }
  
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Inicializar encabezados
    sheet.getRange('A1').setValue('Grupo');
    sheet.getRange('A2').setValue('Grupo 1');
    sheet.getRange('A3').setValue('Grupo 2');
    sheet.getRange('A4').setValue('Grupo 3');
    sheet.getRange('B1').setValue('Conteo');
    sheet.getRange('B2').setValue(0);
    sheet.getRange('B3').setValue(0);
    sheet.getRange('B4').setValue(0);
    sheet.getRange('C1').setValue('Última actualización');
  }
  
  return sheet;
}

// ============================================
function setup() {
  const sheet = getSheet();
  const ss = sheet.getParent();
  Logger.log('Setup completado!');
  Logger.log('Spreadsheet URL: ' + ss.getUrl());
}