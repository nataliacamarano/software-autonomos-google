// ============================================================
// DriveMigrate185.gs — FASE 18.5: Migración Drive M1–M4
// Mueve archivos a la estructura ALVARO_EMPRESA y actualiza CONFIG.
// Solo ejecutar tras confirmar el diagnóstico de DriveScan185.gs.
// ============================================================

var DM185_LOG_SHEET = 'DRIVE_MIGRATE_LOG';

// ─── Punto de entrada público ────────────────────────────────

function migrarDriveM1M2M3M4() {
  _dm185_run_();
}

// ─── Runner principal ────────────────────────────────────────

function _dm185_run_() {
  var ss = getTargetSS_();
  var ui = SpreadsheetApp.getUi();

  // ── Leer IDs destino de DRIVE_CONFIG ─────────────────────
  var facTargetId = String(_dmGetConfig_(ss, 'FACTURAS_FOLDER_ID') || '').trim();
  var gasTargetId = String(_dmGetConfig_(ss, 'GASTOS_FOLDER_ID')  || '').trim();

  if (!facTargetId || !gasTargetId) {
    ui.alert('❌ Faltan IDs en DRIVE_CONFIG\n\nAsegúrate de haber ejecutado "DRIVE — Crear estructura segura" primero.');
    return;
  }

  // ── Leer IDs origen de CONFIG ─────────────────────────────
  var facSourceId = String(getConfig('carpeta_facturas_id') || '').trim();
  var gasSourceId = String(getConfig('gastos_folder_id')   || '').trim();

  // ── Obtener nombres para el resumen de confirmación ──────
  var facSourceNm  = facSourceId ? _dm185_folderName_(facSourceId)  : '(no configurado)';
  var facTargetNm  = _dm185_folderName_(facTargetId);
  var gasSourceNm  = gasSourceId ? _dm185_folderName_(gasSourceId)  : '(no configurado)';
  var gasTargetNm  = _dm185_folderName_(gasTargetId);

  // ── Contar archivos antes de confirmar ───────────────────
  var facCount = facSourceId && facSourceId !== facTargetId ? _dm185_countFiles_(facSourceId) : 0;
  var gasCount = gasSourceId && gasSourceId !== gasTargetId ? _dm185_countFiles_(gasSourceId) : 0;

  // ── Diálogo de confirmación ───────────────────────────────
  var lineas = [
    'M1 — Mover ' + facCount + ' PDFs:',
    '  Origen:  ' + facSourceNm,
    '  Destino: ' + facTargetNm + ' (ALVARO_EMPRESA/2026/FACTURAS/)',
    '',
    'M2 — Mover ' + gasCount + ' tickets:',
    '  Origen:  ' + gasSourceNm,
    '  Destino: ' + gasTargetNm + ' (ALVARO_EMPRESA/2026/GASTOS/)',
    '',
    'M3 — Actualizar carpeta_facturas_id en CONFIG',
    'M4 — Actualizar gastos_folder_id en CONFIG',
    '',
    'Los IDs de Drive son permanentes — las URLs en',
    'FACTURAS y GASTOS seguirán funcionando tras mover.',
    '',
    '¿Confirmar migración?'
  ];

  var resp = ui.alert(
    '⚠ MIGRACIÓN DRIVE M1–M4',
    lineas.join('\n'),
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) {
    ui.alert('Migración cancelada. No se ha modificado nada.');
    return;
  }

  // ── Obtener carpetas destino ──────────────────────────────
  var facTarget, gasTarget;
  try { facTarget = DriveApp.getFolderById(facTargetId); }
  catch(e) { ui.alert('❌ FACTURAS_FOLDER no accesible:\n' + e.message); return; }
  try { gasTarget = DriveApp.getFolderById(gasTargetId); }
  catch(e) { ui.alert('❌ GASTOS_FOLDER no accesible:\n' + e.message);   return; }

  var log  = [];
  var errs = 0;
  var ts   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  log.push(['Timestamp', 'Paso', 'Estado', 'Archivo', 'Detalle']);
  log.push([ts, 'INICIO', 'INFO', '', 'Migración M1–M4 iniciada']);

  // ── M1: Mover PDFs de facturas ───────────────────────────
  var movedFac = 0;
  if (!facSourceId) {
    log.push([ts, 'M1', '⚠ SKIP', '', 'carpeta_facturas_id no configurado']);
  } else if (facSourceId === facTargetId) {
    log.push([ts, 'M1', '✅ YA OK', '', 'Origen y destino son la misma carpeta — nada que mover']);
  } else {
    try {
      var facSource = DriveApp.getFolderById(facSourceId);
      var files     = facSource.getFiles();
      while (files.hasNext()) {
        var file = files.next();
        var fname = file.getName();
        try {
          file.moveTo(facTarget);
          log.push([ts, 'M1', '✅ OK', fname, facTarget.getName()]);
          movedFac++;
        } catch(e) {
          log.push([ts, 'M1', '❌ ERROR', fname, e.message]);
          errs++;
        }
      }
      log.push([ts, 'M1', 'RESUMEN', '', movedFac + ' archivos movidos, ' + errs + ' errores']);
    } catch(e) {
      log.push([ts, 'M1', '❌ ERROR CARPETA', facSourceId, e.message]);
      errs++;
    }
  }

  // ── M2: Mover tickets de gastos ──────────────────────────
  var movedGas = 0;
  if (!gasSourceId) {
    log.push([ts, 'M2', '⚠ SKIP', '', 'gastos_folder_id no configurado']);
  } else if (gasSourceId === gasTargetId) {
    log.push([ts, 'M2', '✅ YA OK', '', 'Origen y destino son la misma carpeta — nada que mover']);
  } else {
    try {
      var gasSource = DriveApp.getFolderById(gasSourceId);
      var files2    = gasSource.getFiles();
      while (files2.hasNext()) {
        var file2  = files2.next();
        var fname2 = file2.getName();
        try {
          file2.moveTo(gasTarget);
          log.push([ts, 'M2', '✅ OK', fname2, gasTarget.getName()]);
          movedGas++;
        } catch(e) {
          log.push([ts, 'M2', '❌ ERROR', fname2, e.message]);
          errs++;
        }
      }
      log.push([ts, 'M2', 'RESUMEN', '', movedGas + ' archivos movidos, ' + errs + ' errores']);
    } catch(e) {
      log.push([ts, 'M2', '❌ ERROR CARPETA', gasSourceId, e.message]);
      errs++;
    }
  }

  // ── M3: Actualizar carpeta_facturas_id ───────────────────
  try {
    setConfig('carpeta_facturas_id', facTargetId);
    log.push([ts, 'M3', '✅ OK', 'carpeta_facturas_id', facTargetId + ' (' + facTargetNm + ')']);
  } catch(e) {
    log.push([ts, 'M3', '❌ ERROR', 'carpeta_facturas_id', e.message]);
    errs++;
  }

  // ── M4: Actualizar gastos_folder_id ──────────────────────
  try {
    setConfig('gastos_folder_id', gasTargetId);
    log.push([ts, 'M4', '✅ OK', 'gastos_folder_id', gasTargetId + ' (' + gasTargetNm + ')']);
  } catch(e) {
    log.push([ts, 'M4', '❌ ERROR', 'gastos_folder_id', e.message]);
    errs++;
  }

  log.push([ts, 'FIN', errs === 0 ? '✅ SIN ERRORES' : '⚠ ' + errs + ' errores', '',
    'FAC movidos: ' + movedFac + ' | GAS movidos: ' + movedGas]);

  // ── Escribir log ──────────────────────────────────────────
  var logSh = ss.getSheetByName(DM185_LOG_SHEET);
  if (!logSh) logSh = ss.insertSheet(DM185_LOG_SHEET);
  else logSh.clearContents();

  logSh.getRange(1, 1, log.length, 5).setValues(log);
  logSh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  try { logSh.autoResizeColumns(1, 5); } catch(e) {}

  // ── Alerta final ──────────────────────────────────────────
  var msg = errs === 0
    ? '✅ Migración completada sin errores\n\n' +
      'M1: ' + movedFac + ' PDFs movidos a FACTURAS\n' +
      'M2: ' + movedGas + ' tickets movidos a GASTOS\n' +
      'M3: carpeta_facturas_id actualizado\n' +
      'M4: gastos_folder_id actualizado\n\n' +
      'Revisa la hoja ' + DM185_LOG_SHEET + ' para el detalle.'
    : '⚠ Migración completada con ' + errs + ' errores.\n\n' +
      'M1: ' + movedFac + ' PDFs movidos\n' +
      'M2: ' + movedGas + ' tickets movidos\n\n' +
      'Revisa la hoja ' + DM185_LOG_SHEET + ' para los errores.';

  ui.alert(msg);
  logEvent('INFO', 'migrarDriveM1M2M3M4',
    'M1:' + movedFac + ' M2:' + movedGas + ' errores:' + errs, '');
}

// ─── Helpers ─────────────────────────────────────────────────

function _dm185_folderName_(folderId) {
  if (!folderId) return '';
  try { return DriveApp.getFolderById(String(folderId).trim()).getName(); }
  catch(e) { return '? (' + folderId + ')'; }
}

function _dm185_countFiles_(folderId) {
  if (!folderId) return 0;
  try {
    var folder = DriveApp.getFolderById(String(folderId).trim());
    var files  = folder.getFiles();
    var count  = 0;
    while (files.hasNext()) { files.next(); count++; }
    return count;
  } catch(e) { return -1; }
}
