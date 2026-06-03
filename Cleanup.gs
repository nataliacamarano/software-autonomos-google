// ============================================================
// Cleanup.gs — Limpieza y reparación del sistema
// ============================================================

const SYSTEM_SHEET_NAMES = new Set([
  'CONFIG', 'CLIENTES', 'OPERARIOS', 'TRABAJOS', 'FACTURAS',
  'PENDIENTES', 'COMPLETADOS', 'LOGS', 'RESP_CLIENTES', 'RESP_OPERARIOS'
]);

// ── Reparación de vínculos formulario-hoja ────────────────────

/**
 * Repara el vínculo entre formularios y hojas de respuestas.
 *
 * Secuencia:
 *  1. Detecta y mueve respuestas de clientes que estén en RESP_OPERARIOS.
 *  2. Desvincula el formulario Clientes de su destino actual y lo revincula
 *     creando una RESP_CLIENTES limpia.
 *  3. Ídem para el formulario Operarios → RESP_OPERARIOS.
 *
 * Ejecutar desde el editor cuando los formularios escriben en la hoja incorrecta.
 */
function fixFormBindings() {
  Logger.log('=== INICIO REPARACIÓN DE VÍNCULOS ===');

  // Paso 1: mover respuestas de clientes mal ubicadas en RESP_OPERARIOS
  moveClienteResponsesMisplaced_();

  // Paso 2: revincular formulario Clientes → RESP_CLIENTES
  const clienteId = getConfig('form_clientes_id');
  if (clienteId) {
    Logger.log('--- Revinculando formulario Clientes ---');
    rebindFormSafely_(clienteId, 'RESP_CLIENTES');
  } else {
    Logger.log('⚠️ form_clientes_id no encontrado en CONFIG. Ejecuta setupForms() primero.');
  }

  // Paso 3: revincular formulario Operarios → RESP_OPERARIOS
  const operarioId = getConfig('form_operarios_id');
  if (operarioId) {
    Logger.log('--- Revinculando formulario Operarios ---');
    rebindFormSafely_(operarioId, 'RESP_OPERARIOS');
  } else {
    Logger.log('⚠️ form_operarios_id no encontrado en CONFIG. Ejecuta setupForms() primero.');
  }

  Logger.log('=== REPARACIÓN COMPLETADA ===');
  Logger.log('Abre cada formulario en Google Forms y verifica que "Respuestas" apunta a la hoja correcta.');
}

// ── Limpieza de hojas antiguas ────────────────────────────────

/**
 * Archiva hojas antiguas (AppSheet / formularios anteriores) añadiendo
 * el prefijo ARCHIVO_. No borra datos. Registra todo en LOGS.
 */
function cleanupLegacyFormsAndSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let protegidas = 0, renombradas = 0, archivadas = 0;

  logEvent('INFO', 'cleanupLegacyFormsAndSheets', 'Inicio', sheets.length + ' hojas encontradas');
  Logger.log('=== INICIO LIMPIEZA ===');
  Logger.log('Total hojas: ' + sheets.length);

  sheets.forEach(function(sheet) {
    const name = sheet.getName();

    if (SYSTEM_SHEET_NAMES.has(name)) {
      Logger.log('✓ PROTEGIDA: "' + name + '"');
      protegidas++;
      return;
    }

    if (name.startsWith('ARCHIVO_')) {
      Logger.log('ℹ️ Ya archivada, sin cambios: "' + name + '"');
      return;
    }

    if (!ss.getSheetByName('RESP_CLIENTES') && esHojaRespuestasClientes_(sheet)) {
      sheet.setName('RESP_CLIENTES');
      logEvent('INFO', 'cleanupLegacyFormsAndSheets', 'Renombrada a RESP_CLIENTES', 'era: "' + name + '"');
      Logger.log('✅ RENOMBRADA → RESP_CLIENTES: "' + name + '"');
      renombradas++;
      return;
    }

    const nombreArchivo = buildArchivoName_(ss, name);
    sheet.setName(nombreArchivo);
    logEvent('INFO', 'cleanupLegacyFormsAndSheets', 'Archivada: "' + name + '"', '→ "' + nombreArchivo + '"');
    Logger.log('📦 ARCHIVADA: "' + name + '" → "' + nombreArchivo + '"');
    archivadas++;
  });

  const resumen = 'Protegidas: ' + protegidas + ' | Renombradas: ' + renombradas + ' | Archivadas: ' + archivadas;
  Logger.log('=== LIMPIEZA COMPLETADA — ' + resumen + ' ===');
  logEvent('INFO', 'cleanupLegacyFormsAndSheets', 'Completada', resumen);
}

// ── Utilidades privadas ───────────────────────────────────────

/**
 * Mueve respuestas de clientes que están en RESP_OPERARIOS a RESP_CLIENTES.
 * Detecta el error comparando cabeceras con campos del formulario de clientes.
 * Archiva RESP_OPERARIOS en vez de borrarla.
 */
function moveClienteResponsesMisplaced_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const operSheet = ss.getSheetByName('RESP_OPERARIOS');

  if (!operSheet) {
    Logger.log('ℹ️ RESP_OPERARIOS no existe. Nada que mover.');
    return;
  }

  const lastCol = operSheet.getLastColumn();
  const lastRow = operSheet.getLastRow();

  if (lastRow < 1 || lastCol < 1) {
    Logger.log('ℹ️ RESP_OPERARIOS vacía. Nada que mover.');
    return;
  }

  // Comprobar si sus cabeceras son de cliente
  if (!esHojaRespuestasClientes_(operSheet)) {
    Logger.log('ℹ️ RESP_OPERARIOS tiene cabeceras de operario correctas. No se mueve nada.');
    return;
  }

  Logger.log('⚠️ RESP_OPERARIOS contiene respuestas de CLIENTES. Iniciando traslado...');
  Logger.log('Filas a mover: ' + (lastRow - 1));

  const data = operSheet.getDataRange().getValues();
  const numCols = data[0].length;

  let clientSheet = ss.getSheetByName('RESP_CLIENTES');
  if (!clientSheet) {
    clientSheet = ss.insertSheet('RESP_CLIENTES');
    Logger.log('✓ Hoja RESP_CLIENTES creada');
  }

  if (clientSheet.getLastRow() === 0) {
    clientSheet.getRange(1, 1, data.length, numCols).setValues(data);
  } else {
    const dataRows = data.slice(1);
    if (dataRows.length > 0) {
      clientSheet.getRange(clientSheet.getLastRow() + 1, 1, dataRows.length, numCols).setValues(dataRows);
    }
  }

  const moved = lastRow - 1;
  Logger.log('✅ ' + moved + ' respuesta(s) trasladada(s) a RESP_CLIENTES');
  logEvent('INFO', 'moveClienteResponsesMisplaced_', moved + ' filas trasladadas a RESP_CLIENTES', '');

  // Archivar RESP_OPERARIOS sin borrarla
  operSheet.setName('ARCHIVO_RESP_OPERARIOS_datos_cliente');
  Logger.log('📦 RESP_OPERARIOS archivada como "ARCHIVO_RESP_OPERARIOS_datos_cliente"');
}

/**
 * Desvincula un formulario de su destino actual y lo revincula al spreadsheet,
 * garantizando que la nueva hoja de respuestas se llame exactamente targetSheetName.
 */
function rebindFormSafely_(formId, targetSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let form;
  try {
    form = FormApp.openById(formId);
  } catch (e) {
    Logger.log('❌ No se pudo abrir el formulario: ' + formId + ' — ' + e.message);
    logEvent('ERROR', 'rebindFormSafely_', 'Formulario no accesible: ' + formId, e.message);
    return false;
  }

  // Archivar la hoja destino actual si existe (puede tener datos incorrectos)
  const existingTarget = ss.getSheetByName(targetSheetName);
  if (existingTarget) {
    const archiveName = buildArchivoName_(ss, targetSheetName);
    existingTarget.setName(archiveName);
    Logger.log('📦 Hoja existente archivada: "' + targetSheetName + '" → "' + archiveName + '"');
  }

  // Desvincular destino actual
  if (form.getDestinationId()) {
    form.removeDestination();
    Utilities.sleep(1500);
    Logger.log('✓ Destino anterior desvinculado del formulario');
  }

  // Snapshot antes de crear el nuevo vínculo
  const snapshot = ss.getSheets().map(function(s) { return s.getName(); });

  // Vincular de nuevo al spreadsheet → Google crea nueva hoja de respuestas
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // Detectar y renombrar la nueva hoja por diferencia con el snapshot
  const ok = renameNewSheet_(ss, snapshot, targetSheetName);

  if (ok) {
    Logger.log('✅ Formulario vinculado correctamente → "' + targetSheetName + '"');
    logEvent('INFO', 'rebindFormSafely_', 'Revinculado: ' + formId, '→ ' + targetSheetName);
  } else {
    Logger.log('❌ No se encontró la nueva hoja de respuestas para "' + targetSheetName + '"');
    logEvent('ERROR', 'rebindFormSafely_', 'No se encontró nueva hoja', formId + ' → ' + targetSheetName);
  }
  return ok;
}

/**
 * Devuelve true si la hoja tiene cabeceras del formulario de Clientes
 * (al menos 3 de: nombre, apellidos, empresa, email, teléfono, dirección).
 */
function esHojaRespuestasClientes_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (sheet.getLastRow() === 0 || lastCol === 0) return false;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h).trim().toLowerCase(); });

  const camposClave = ['nombre', 'apellidos', 'empresa', 'email', 'teléfono', 'dirección'];
  const coincidencias = camposClave.filter(function(c) {
    return headers.some(function(h) { return h.includes(c); });
  });

  Logger.log('  "' + sheet.getName() + '": ' + coincidencias.length + '/6 campos cliente — [' + headers.join(', ') + ']');
  return coincidencias.length >= 3;
}

/**
 * Construye un nombre único con prefijo ARCHIVO_ (máx. 100 caracteres).
 * Añade sufijo _2, _3… en caso de colisión.
 */
function buildArchivoName_(ss, originalName) {
  const MAX = 100;
  const base = ('ARCHIVO_' + originalName).substring(0, MAX);
  if (!ss.getSheetByName(base)) return base;
  for (let i = 2; i <= 99; i++) {
    const candidate = ('ARCHIVO_' + originalName).substring(0, MAX - 3) + '_' + i;
    if (!ss.getSheetByName(candidate)) return candidate;
  }
  return base.substring(0, MAX - 5) + '_' + String(Date.now()).slice(-4);
}

// ── Normalización de nombres de hojas de respuestas ───────────

/**
 * Renombra las hojas de respuestas activas a sus nombres canónicos.
 *
 * Situación actual conocida:
 *   "ARCHIVO_Respuestas de formulario 6"  →  RESP_CLIENTES
 *   "RESP_CLIENTES" (nombre incorrecto)   →  RESP_OPERARIOS
 *
 * Usa nombre temporal para resolver el conflicto de nombres sin
 * crear hojas, borrar datos ni tocar formularios.
 *
 * NO crea hojas.
 * NO borra hojas.
 * NO mueve datos.
 * NO toca formularios ni vínculos (sin setDestination/removeDestination).
 */
function normalizeActiveResponseSheetNames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Nombres implicados
  var SRC_CLIENTES  = 'ARCHIVO_Respuestas de formulario 6'; // hoja real de Clientes
  var WRONG_NAME    = 'RESP_CLIENTES';  // nombre incorrecto — apunta a Operarios
  var FINAL_CLI     = 'RESP_CLIENTES';
  var FINAL_OPE     = 'RESP_OPERARIOS';
  var TEMP          = 'TEMP_RESP_SWAP'; // nombre intermedio para evitar conflicto

  Logger.log('================================================');
  Logger.log('normalizeActiveResponseSheetNames()');
  Logger.log('Solo renombra hojas. Sin datos, sin formularios.');
  Logger.log('================================================');

  // ── Lectura del estado actual ─────────────────────────────
  var sheetSrc  = ss.getSheetByName(SRC_CLIENTES);
  var sheetWrong = ss.getSheetByName(WRONG_NAME);
  var sheetFinalOpe = ss.getSheetByName(FINAL_OPE);
  var sheetTemp = ss.getSheetByName(TEMP);

  Logger.log('Estado inicial:');
  Logger.log('  "' + SRC_CLIENTES + '": ' + (sheetSrc       ? 'EXISTE (GID=' + sheetSrc.getSheetId()       + ', filas=' + Math.max(0, sheetSrc.getLastRow()-1)   + ')' : 'NO EXISTE'));
  Logger.log('  "' + WRONG_NAME   + '": ' + (sheetWrong     ? 'EXISTE (GID=' + sheetWrong.getSheetId()     + ', filas=' + Math.max(0, sheetWrong.getLastRow()-1)  + ')' : 'NO EXISTE'));
  Logger.log('  "' + FINAL_OPE   + '": ' + (sheetFinalOpe  ? 'EXISTE (GID=' + sheetFinalOpe.getSheetId()  + ')' : 'NO EXISTE'));
  Logger.log('  "' + TEMP        + '": ' + (sheetTemp       ? 'EXISTE — conflicto de nombre temporal' : 'libre'));

  // ── Validaciones previas ──────────────────────────────────
  if (!sheetSrc) {
    Logger.log('\nERROR: "' + SRC_CLIENTES + '" no existe. Abortando sin cambios.');
    logEvent('ERROR', 'normalizeActiveResponseSheetNames', 'Hoja origen no encontrada', SRC_CLIENTES);
    return;
  }

  if (sheetFinalOpe) {
    Logger.log('\nERROR: "' + FINAL_OPE + '" ya existe (GID=' + sheetFinalOpe.getSheetId() + ').');
    Logger.log('       Para evitar sobreescribir, abortando sin cambios.');
    Logger.log('       Revisa el estado del spreadsheet y confirma manualmente.');
    logEvent('ERROR', 'normalizeActiveResponseSheetNames', FINAL_OPE + ' ya existe', 'GID=' + sheetFinalOpe.getSheetId());
    return;
  }

  if (sheetTemp) {
    Logger.log('\nERROR: nombre temporal "' + TEMP + '" ya en uso.');
    Logger.log('       Renombra o elimina esa hoja manualmente y reintenta.');
    logEvent('ERROR', 'normalizeActiveResponseSheetNames', 'Nombre temporal ocupado', TEMP);
    return;
  }

  // ── Paso 1: liberar RESP_CLIENTES renombrándola a temporal ─
  if (sheetWrong) {
    sheetWrong.setName(TEMP);
    Logger.log('\nPaso 1: "' + WRONG_NAME + '" → "' + TEMP + '" (temporal)');
    logEvent('INFO', 'normalizeActiveResponseSheetNames', WRONG_NAME + ' → ' + TEMP, '');
  } else {
    Logger.log('\nPaso 1: "' + WRONG_NAME + '" no existe, sin accion.');
  }

  // ── Paso 2: renombrar fuente a RESP_CLIENTES ─────────────
  sheetSrc.setName(FINAL_CLI);
  Logger.log('Paso 2: "' + SRC_CLIENTES + '" → "' + FINAL_CLI + '"');
  logEvent('INFO', 'normalizeActiveResponseSheetNames', SRC_CLIENTES + ' → ' + FINAL_CLI, '');

  // ── Paso 3: renombrar temporal a RESP_OPERARIOS ───────────
  if (sheetWrong) {
    var sheetTempNow = ss.getSheetByName(TEMP);
    if (sheetTempNow) {
      sheetTempNow.setName(FINAL_OPE);
      Logger.log('Paso 3: "' + TEMP + '" → "' + FINAL_OPE + '"');
      logEvent('INFO', 'normalizeActiveResponseSheetNames', TEMP + ' → ' + FINAL_OPE, '');
    } else {
      Logger.log('Paso 3: hoja temporal no encontrada tras paso 1 (inesperado).');
      logEvent('WARN', 'normalizeActiveResponseSheetNames', 'Temporal no encontrada en paso 3', TEMP);
    }
  } else {
    Logger.log('Paso 3: sin temporal que renombrar.');
  }

  // ── Resumen ───────────────────────────────────────────────
  var rCli = ss.getSheetByName(FINAL_CLI);
  var rOpe = ss.getSheetByName(FINAL_OPE);
  Logger.log('\n================================================');
  Logger.log('RESULTADO FINAL:');
  Logger.log('  "' + FINAL_CLI + '": ' + (rCli ? 'OK GID=' + rCli.getSheetId() + '  filas=' + Math.max(0, rCli.getLastRow()-1) : 'NO EXISTE (error)'));
  Logger.log('  "' + FINAL_OPE + '": ' + (rOpe ? 'OK GID=' + rOpe.getSheetId() + '  filas=' + Math.max(0, rOpe.getLastRow()-1) : 'NO EXISTE (sin hoja de Operarios previa)'));
  Logger.log('Siguiente paso: ejecuta findTestResponses() para verificar.');
  Logger.log('================================================');
}

// ── Eliminación controlada de hojas archivadas ────────────────

/**
 * Elimina únicamente las 4 hojas archivadas antiguas listadas en TO_DELETE.
 *
 * Antes de borrar:
 *  1. Verifica que RESP_CLIENTES y RESP_OPERARIOS existen (aborta si no).
 *  2. Verifica que contienen los datos de prueba (avisa si no, pero sigue).
 *  3. Muestra cada hoja que se va a eliminar con su GID y filas.
 *  4. Elimina solo esas hojas.
 *  5. Lista todas las hojas restantes al finalizar.
 *
 * NO crea hojas. NO renombra hojas. NO modifica formularios. NO mueve datos.
 */
function deleteArchivedLegacySheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Hojas que NUNCA se eliminan — doble candado
  var PROTECTED = [
    'CONFIG','CLIENTES','OPERARIOS','TRABAJOS','FACTURAS',
    'PENDIENTES','COMPLETADOS','LOGS',
    'RESP_CLIENTES','RESP_OPERARIOS',
    'ARCHIVO_Trabajo de formulario 1','ARCHIVO_Facturacion'
  ];

  // Exactamente estas 4 hojas — ni una más
  var TO_DELETE = [
    'ARCHIVO_RESP_CLIENTES',
    'ARCHIVO_RESP_OPERARIOS_datos_cliente',
    'ARCHIVO_Alta empleados de formulario 3',
    'ARCHIVO_Alta clientes de formulario 2'
  ];

  var TEST_C = 'CLIENTE_TEST_LIMPIEZA';
  var TEST_O = 'OPERARIO_TEST_LIMPIEZA';

  Logger.log('================================================');
  Logger.log('deleteArchivedLegacySheets()');
  Logger.log('Elimina unicamente las 4 hojas archivadas listadas.');
  Logger.log('================================================');

  // ── 1. Verificar que RESP_CLIENTES y RESP_OPERARIOS existen ──
  Logger.log('\n[1] Verificacion de hojas activas:');
  var respCli = ss.getSheetByName('RESP_CLIENTES');
  var respOpe = ss.getSheetByName('RESP_OPERARIOS');

  Logger.log('  RESP_CLIENTES:  ' + (respCli
    ? 'EXISTE  GID=' + respCli.getSheetId() + '  filas=' + Math.max(0, respCli.getLastRow() - 1)
    : 'NO EXISTE'));
  Logger.log('  RESP_OPERARIOS: ' + (respOpe
    ? 'EXISTE  GID=' + respOpe.getSheetId() + '  filas=' + Math.max(0, respOpe.getLastRow() - 1)
    : 'NO EXISTE'));

  if (!respCli || !respOpe) {
    Logger.log('\nABORTANDO: RESP_CLIENTES o RESP_OPERARIOS no existen.');
    Logger.log('No se elimina ninguna hoja. Corrige el estado primero.');
    logEvent('ERROR', 'deleteArchivedLegacySheets',
             'Abortado — falta hoja activa', (!respCli ? 'RESP_CLIENTES' : 'RESP_OPERARIOS'));
    return;
  }

  // ── 2. Verificar datos de prueba ──────────────────────────
  Logger.log('\n[2] Verificacion de datos de prueba:');
  var hasTestC = sheetContainsText_(respCli, TEST_C);
  var hasTestO = sheetContainsText_(respOpe, TEST_O);

  Logger.log('  "' + TEST_C + '" en RESP_CLIENTES:  ' + (hasTestC ? 'ENCONTRADO' : 'NO ENCONTRADO'));
  Logger.log('  "' + TEST_O + '" en RESP_OPERARIOS: ' + (hasTestO ? 'ENCONTRADO' : 'NO ENCONTRADO'));

  if (!hasTestC || !hasTestO) {
    Logger.log('  AVISO: datos de prueba no encontrados en una o ambas hojas.');
    Logger.log('  Las hojas activas pueden estar correctas aunque el test no se haya enviado.');
    Logger.log('  Continuando con la eliminacion (hojas confirmadas por el usuario).');
  } else {
    Logger.log('  OK: datos de prueba verificados en ambas hojas activas.');
  }

  // ── 3. Mostrar hojas candidatas ───────────────────────────
  Logger.log('\n[3] Hojas candidatas a eliminar:');
  var sheetsToDelete = [];

  for (var i = 0; i < TO_DELETE.length; i++) {
    var candidateName = TO_DELETE[i];
    var candidateSheet = ss.getSheetByName(candidateName);
    if (candidateSheet) {
      Logger.log('  [EXISTE]    "' + candidateName + '"' +
                 '  GID=' + candidateSheet.getSheetId() +
                 '  filas=' + Math.max(0, candidateSheet.getLastRow() - 1));
      sheetsToDelete.push(candidateSheet);
    } else {
      Logger.log('  [NO EXISTE] "' + candidateName + '" — omitida');
    }
  }

  if (sheetsToDelete.length === 0) {
    Logger.log('\nNinguna hoja candidata existe. Nada que eliminar.');
    Logger.log('================================================');
    return;
  }

  // ── 4. Eliminar ───────────────────────────────────────────
  Logger.log('\n[4] Eliminando ' + sheetsToDelete.length + ' hoja(s)...');
  var deleted = [];
  var errors  = [];

  for (var i = 0; i < sheetsToDelete.length; i++) {
    var sheet = sheetsToDelete[i];
    var name  = sheet.getName();

    // Doble candado: nunca tocar hojas protegidas
    if (PROTECTED.indexOf(name) !== -1) {
      Logger.log('  SALTADA (protegida): "' + name + '"');
      continue;
    }

    try {
      ss.deleteSheet(sheet);
      Logger.log('  ELIMINADA: "' + name + '"');
      logEvent('INFO', 'deleteArchivedLegacySheets', 'Eliminada: ' + name, '');
      deleted.push(name);
    } catch (e) {
      Logger.log('  ERROR eliminando "' + name + '": ' + e.message);
      logEvent('ERROR', 'deleteArchivedLegacySheets', 'Error: ' + name, e.message);
      errors.push(name + ' (' + e.message + ')');
    }
  }

  // ── 5. Hojas restantes ────────────────────────────────────
  Logger.log('\n[5] HOJAS RESTANTES EN EL SPREADSHEET:');
  var remaining = ss.getSheets();
  for (var i = 0; i < remaining.length; i++) {
    var s = remaining[i];
    Logger.log('  "' + s.getName() + '"' +
               '  GID=' + s.getSheetId() +
               '  filas=' + Math.max(0, s.getLastRow() - 1));
  }

  // ── Resumen final ─────────────────────────────────────────
  Logger.log('\n================================================');
  Logger.log('RESUMEN:');
  Logger.log('  Eliminadas (' + deleted.length + '): ' +
             (deleted.length > 0 ? deleted.join(', ') : 'ninguna'));
  if (errors.length > 0) {
    Logger.log('  Errores    (' + errors.length + '): ' + errors.join(' | '));
  }
  Logger.log('  Hojas restantes: ' + remaining.length);
  Logger.log('================================================');
}

// ── Eliminación de hoja con vínculo residual ─────────────────

/**
 * Identifica el formulario vinculado a "ARCHIVO_Alta empleados de formulario 3",
 * confirma que NO es uno de los formularios activos (CONFIG),
 * desvincula el formulario legado y elimina la hoja.
 *
 * Hojas protegidas: RESP_CLIENTES, RESP_OPERARIOS, CONFIG, CLIENTES,
 * OPERARIOS, TRABAJOS, FACTURAS, PENDIENTES, COMPLETADOS, LOGS.
 *
 * NO modifica ninguna otra hoja ni formulario.
 */
function unlinkAndDeleteLegacyEmployeeSheet() {
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var TARGET     = 'ARCHIVO_Alta empleados de formulario 3';

  var PROTECTED = [
    'CONFIG','CLIENTES','OPERARIOS','TRABAJOS','FACTURAS',
    'PENDIENTES','COMPLETADOS','LOGS',
    'RESP_CLIENTES','RESP_OPERARIOS'
  ];

  Logger.log('================================================');
  Logger.log('unlinkAndDeleteLegacyEmployeeSheet()');
  Logger.log('Objetivo: desvincular formulario residual y eliminar');
  Logger.log('"' + TARGET + '"');
  Logger.log('================================================');

  // ── 1. Verificar que la hoja existe ──────────────────────────
  var targetSheet = ss.getSheetByName(TARGET);
  Logger.log('\n[1] Hoja objetivo:');
  if (!targetSheet) {
    Logger.log('  "' + TARGET + '": NO EXISTE. Nada que hacer.');
    return;
  }
  Logger.log('  Nombre: "' + targetSheet.getName() + '"');
  Logger.log('  GID:    ' + targetSheet.getSheetId());
  Logger.log('  Filas:  ' + Math.max(0, targetSheet.getLastRow() - 1));

  // ── 2. Formularios activos en CONFIG ──────────────────────────
  var clienteId  = getConfig('form_clientes_id')  || '';
  var operarioId = getConfig('form_operarios_id') || '';
  var activeIds  = [];
  if (clienteId)  activeIds.push(clienteId);
  if (operarioId) activeIds.push(operarioId);

  Logger.log('\n[2] Formularios activos en CONFIG:');
  Logger.log('  form_clientes_id:  ' + (clienteId  || '(no configurado)'));
  Logger.log('  form_operarios_id: ' + (operarioId || '(no configurado)'));

  // ── 3. Buscar formularios de Drive vinculados a este spreadsheet
  Logger.log('\n[3] Buscando formularios vinculados a este spreadsheet...');
  var ssId       = ss.getId();
  var legacyList = [];
  var checked    = 0;

  var files = DriveApp.getFilesByType(MimeType.GOOGLE_FORMS);
  while (files.hasNext()) {
    var file = files.next();
    checked++;
    try {
      var form    = FormApp.openById(file.getId());
      var destId  = form.getDestinationId();
      if (destId !== ssId) continue;

      var isActive = activeIds.indexOf(file.getId()) !== -1;
      Logger.log('\n  Formulario: "' + form.getTitle() + '"');
      Logger.log('    ID:     ' + file.getId());
      Logger.log('    Estado: ' + (isActive ? 'ACTIVO — NO tocar' : 'LEGADO — candidato'));
      if (!isActive) legacyList.push({ form: form, id: file.getId(), title: form.getTitle() });
    } catch (e) {
      // Formulario no accesible — ignorar
    }
  }
  Logger.log('\n  Formularios revisados: '   + checked);
  Logger.log('  Formularios legados: '       + legacyList.length);

  // ── 4. Desvincular formularios legados ────────────────────────
  Logger.log('\n[4] Desvinculando formularios legados:');
  if (legacyList.length === 0) {
    Logger.log('  Ningún formulario legado accesible encontrado.');
    Logger.log('  El vínculo puede ser interno de Google (proceder igualmente a eliminar).');
  }
  for (var i = 0; i < legacyList.length; i++) {
    var lf = legacyList[i];
    try {
      lf.form.removeDestination();
      Logger.log('  DESVINCULADO: "' + lf.title + '"  ID=' + lf.id);
      logEvent('INFO', 'unlinkAndDeleteLegacyEmployeeSheet',
               'Formulario desvinculado: ' + lf.title, lf.id);
    } catch (e) {
      Logger.log('  ERROR desvinculando "' + lf.title + '": ' + e.message);
      logEvent('ERROR', 'unlinkAndDeleteLegacyEmployeeSheet',
               'Error desvincular: ' + lf.title, e.message);
    }
  }

  // ── 5. Eliminar la hoja ───────────────────────────────────────
  Logger.log('\n[5] Eliminando "' + TARGET + '":');

  if (PROTECTED.indexOf(TARGET) !== -1) {
    Logger.log('  ERROR: hoja en lista protegida. Abortando.');
    return;
  }

  var sheetToDelete = ss.getSheetByName(TARGET);
  if (!sheetToDelete) {
    Logger.log('  Hoja ya no existe tras desvincular. Nada que eliminar.');
  } else {
    try {
      ss.deleteSheet(sheetToDelete);
      Logger.log('  ELIMINADA: "' + TARGET + '"');
      logEvent('INFO', 'unlinkAndDeleteLegacyEmployeeSheet', 'Eliminada: ' + TARGET, '');
    } catch (e) {
      Logger.log('  ERROR al eliminar hoja: ' + e.message);
      logEvent('ERROR', 'unlinkAndDeleteLegacyEmployeeSheet', 'Error eliminar hoja', e.message);
    }
  }

  Logger.log('\n================================================');
  Logger.log('COMPLETADO.');
  Logger.log('================================================');
}

// ── Utilidad privada compartida ───────────────────────────────

/**
 * Devuelve true si alguna celda de la hoja contiene el texto buscado
 * (búsqueda insensible a mayúsculas).
 */
function sheetContainsText_(sheet, searchText) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return false;
  var data = sheet.getRange(1, 1, lastRow, Math.min(lastCol, 20)).getValues();
  var lower = searchText.toLowerCase();
  for (var r = 0; r < data.length; r++) {
    if (data[r].join('|').toLowerCase().indexOf(lower) !== -1) return true;
  }
  return false;
}
