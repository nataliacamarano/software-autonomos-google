// ============================================================
// Utils.gs — Funciones de utilidad compartidas
// ============================================================

/**
 * Devuelve el spreadsheet objetivo.
 * Lee TARGET_SS_ID de Script Properties (establecido por freshInstall).
 * Si no existe la propiedad, usa getActiveSpreadsheet como fallback.
 */
function getTargetSS_() {
  var id = PropertiesService.getScriptProperties().getProperty('TARGET_SS_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) {}
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Devuelve una hoja del spreadsheet objetivo por nombre.
 * Lanza error si no existe.
 */
function getSheet(name) {
  var ss    = getTargetSS_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Hoja no encontrada: "' + name + '"');
  return sheet;
}

/**
 * Devuelve una hoja o la crea si no existe.
 */
function getOrCreateSheet(name) {
  var ss = getTargetSS_();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/**
 * Genera el siguiente ID secuencial con prefijo para una hoja.
 * Ejemplo: nextId('CLIENTES', 'CLI') → 'CLI-001'
 */
function nextId(sheetName, prefix) {
  try {
    const sheet   = getSheet(sheetName);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return prefix + '-001';
    const values  = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const numbers = values
      .filter(v => String(v).startsWith(prefix + '-'))
      .map(v => parseInt(String(v).split('-').pop(), 10))
      .filter(n => !isNaN(n));
    const next = numbers.length ? Math.max(...numbers) + 1 : 1;
    return prefix + '-' + String(next).padStart(3, '0');
  } catch (_) {
    return prefix + '-001';
  }
}

/**
 * Formatea un número como moneda EUR.
 */
function formatCurrency(amount) {
  return Number(amount || 0).toLocaleString('es-ES', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }) + ' €';
}

/**
 * Formatea una fecha como DD/MM/YYYY.
 */
function formatDate(date) {
  if (!date) return '';
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

/**
 * Formatea una fecha y hora como DD/MM/YYYY HH:mm.
 */
function formatDateTime(date) {
  if (!date) return '';
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

/**
 * Devuelve el año actual como string.
 */
function getCurrentYear() {
  return String(new Date().getFullYear());
}

/**
 * Convierte un array de arrays en array de objetos.
 * La primera fila se usa como claves. Omite filas vacías.
 */
function rowsToObjects(rows) {
  if (!rows || rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1)
    .filter(row => row.some(cell => cell !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

/**
 * Aplica el estilo de cabecera estándar a la primera fila.
 */
function applyHeaderStyle_(sheet, numCols) {
  sheet.getRange(1, 1, 1, numCols)
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10);
  sheet.setFrozenRows(1);
}

/**
 * Busca un ID en la columna 1 de una hoja y devuelve su número de fila.
 * Devuelve null si no lo encuentra.
 */
function findRowById_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 2;
  }
  return null;
}

/**
 * Registra un evento en la hoja LOGS.
 */
function logEvent(level, funcion, mensaje, detalle) {
  try {
    const sheet = getOrCreateSheet('LOGS');
    sheet.appendRow([new Date(), level, funcion, mensaje, String(detalle || '')]);
  } catch (e) {
    Logger.log('[LOG ERROR] ' + e.message);
  }
}
