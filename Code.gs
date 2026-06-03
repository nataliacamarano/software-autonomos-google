// ============================================================
// Code.gs — Inicialización del sistema y punto de entrada web
// ============================================================

/**
 * SETUP — Ejecutar UNA SOLA VEZ desde el editor de Apps Script
 * para crear todas las hojas, configuración y carpetas de Drive.
 */
function setup() {
  try {
    Logger.log('=== INICIO SETUP ===');

    createAllSheets_();
    Logger.log('✓ Hojas creadas');

    initConfig();
    Logger.log('✓ CONFIG inicializada');

    createDriveFolders_();
    Logger.log('✓ Carpetas Drive creadas');

    applySheetFormatting_();
    Logger.log('✓ Formato aplicado');

    Logger.log('=== SETUP COMPLETADO ===');
    Logger.log('Siguiente paso: abre la hoja CONFIG y completa los datos de tu empresa.');
  } catch (e) {
    logEvent('ERROR', 'setup', e.message, e.stack);
    Logger.log('❌ ERROR en setup: ' + e.message);
    throw e;
  }
}

// ---- Creación de hojas ----

function createAllSheets_() {
  const ss = getTargetSS_();

  Object.entries(SHEET_HEADERS).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  });

  // Eliminar hojas vacías creadas por defecto por Google Sheets
  ['Sheet1', 'Hoja 1', 'Hoja1'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s && ss.getSheets().length > 1) {
      try { ss.deleteSheet(s); } catch (_) {}
    }
  });
}

// ---- Estructura de carpetas en Drive ----

function createDriveFolders_() {
  const year = getCurrentYear();
  const root = DriveApp.getRootFolder();

  const main = getOrCreateDriveFolder_(root, 'Álvaro — Gestión');

  const facturas = getOrCreateDriveFolder_(main, 'Facturas');
  getOrCreateDriveFolder_(facturas, year);

  const trabajos = getOrCreateDriveFolder_(main, 'Trabajos');
  getOrCreateDriveFolder_(trabajos, year);

  setConfig('carpeta_facturas_id', facturas.getId());
  setConfig('carpeta_trabajos_id', trabajos.getId());

  Logger.log(`Drive Facturas: ${facturas.getId()} | Trabajos: ${trabajos.getId()}`);
}

function getOrCreateDriveFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// ---- Formato de hojas ----

function applySheetFormatting_() {
  const ss = getTargetSS_();

  // Cabeceras
  Object.entries(SHEET_HEADERS).forEach(([name, headers]) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    applyHeaderStyle_(sheet, headers.length);
  });

  // Anchos de columna específicos por hoja
  const colWidths = {
    CLIENTES:    { 1: 100, 3: 150, 4: 150, 5: 180, 6: 200, 8: 250 },
    OPERARIOS:   { 1: 110, 3: 150, 4: 150, 6: 130, 7: 150 },
    TRABAJOS:    { 1: 110, 4: 160, 6: 160, 7: 250, 9: 200, 12: 110 },
    FACTURAS:    { 1: 140, 6: 160, 8: 220, 11: 250, 19: 300 },
    PENDIENTES:  { 3: 160, 4: 160, 5: 220 },
    COMPLETADOS: { 3: 160, 4: 160, 5: 220 },
    LOGS:        { 1: 160, 4: 280, 5: 300 }
  };

  Object.entries(colWidths).forEach(([sheetName, widths]) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    Object.entries(widths).forEach(([col, width]) => {
      sheet.setColumnWidth(Number(col), width);
    });
  });

  SpreadsheetApp.flush();
}

// ---- Web App (Fase 5 — placeholder) ----

function doGet(e) {
  return HtmlService
    .createHtmlOutput(
      '<html><body style="font-family:sans-serif;padding:40px;color:#444">' +
      '<h2>⚙️ Panel en construcción</h2>' +
      '<p>El panel web completo se implementa en la Fase 5.</p>' +
      '</body></html>'
    )
    .setTitle('Sistema de Gestión — Álvaro');
}
