// ============================================================
// BackupAutomatico.gs — Copia de seguridad automática de TODA la hoja de Google
// ============================================================
//
// Cada 12 horas Google hace una copia completa del archivo de la hoja (todas las
// pestañas, también las de notas, empleados, máquinas, chat…) y la guarda en una
// carpeta de Drive. Se conservan las últimas 28 copias (14 días): las más antiguas
// pasan a la papelera de Drive (no se borran para siempre).
//
// Solo se tocan archivos de esa carpeta cuyo nombre empieza por BKA_.PREFIJO.
// Si una copia falla, se anota en LOGS y se avisa por email a la cuenta del programa.
//
// Se instala desde el menú "⚙️ Sistema": "💾 Backup automático — Instalar (cada 12 h)".
// (No confundir con Backup.gs, que copia solo 9 pestañas dentro de la misma hoja y es manual.)
// ============================================================

var BKA_ = {
  HORAS:       12,
  MAX_COPIAS:  28,
  PREFIJO:     'Backup automático - ',
  HANDLER:     'ejecutarBackupAutomatico',
  PROP_CARPETA: 'backup_auto_carpeta_id'
};

/** Carpeta de Drive donde se guardan las copias (se crea la primera vez y se recuerda su ID). */
function _bka_carpeta_(ss) {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(BKA_.PROP_CARPETA);
  if (id) {
    try {
      var f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (e) { /* la carpeta ya no existe: se crea otra */ }
  }
  var nueva = DriveApp.createFolder('Backups automáticos - ' + ss.getName());
  props.setProperty(BKA_.PROP_CARPETA, nueva.getId());
  return nueva;
}

/** Hace una copia completa ahora y limpia las antiguas. Devuelve { nombre, url, carpeta, borradas }. */
function _bka_copiaAhora_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Ya hay otra copia en marcha. Inténtalo en unos minutos.');
  try {
    var ss = getTargetSS_();
    var carpeta = _bka_carpeta_(ss);
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm');
    var nombre = BKA_.PREFIJO + ss.getName() + ' ' + stamp;
    var copia = DriveApp.getFileById(ss.getId()).makeCopy(nombre, carpeta);
    var borradas = _bka_limpiarAntiguas_(carpeta);
    logEvent('INFO', 'backupAutomatico', 'Copia creada: ' + nombre, 'Antiguas a la papelera: ' + borradas);
    return { nombre: nombre, url: copia.getUrl(), carpeta: carpeta.getUrl(), borradas: borradas };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/** Deja solo las últimas BKA_.MAX_COPIAS copias en la carpeta; el resto va a la papelera. */
function _bka_limpiarAntiguas_(carpeta) {
  var lista = [], it = carpeta.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    if (f.getName().indexOf(BKA_.PREFIJO) === 0) lista.push({ f: f, t: f.getDateCreated().getTime() });
  }
  lista.sort(function(a, b) { return b.t - a.t; });            // las más nuevas primero
  var n = 0;
  for (var i = BKA_.MAX_COPIAS; i < lista.length; i++) { lista[i].f.setTrashed(true); n++; }
  return n;
}

/** Lo ejecuta Google solo, cada 12 horas. */
function ejecutarBackupAutomatico() {
  try {
    _bka_copiaAhora_();
  } catch (e) {
    logEvent('ERROR', 'ejecutarBackupAutomatico', e.message, e.stack);
    try {
      GmailApp.sendEmail(Session.getEffectiveUser().getEmail(),
        '⚠️ Falló la copia de seguridad automática',
        'No se pudo hacer la copia de seguridad de la hoja.\n\nMotivo: ' + e.message +
        '\n\nSe volverá a intentar en la próxima ronda. Más detalle en la pestaña LOGS.');
    } catch (_) {}
  }
}

// ---- Opciones del menú ----

function backupAutomaticoInstalar() {
  var ui = SpreadsheetApp.getUi();
  var r;
  try { r = _bka_copiaAhora_(); }   // primera copia ya: así se comprueba que todo funciona antes de dejarlo programado
  catch (e) { ui.alert('❌ No se pudo hacer la primera copia, así que no he programado nada.\n\n' + e.message); return; }
  ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() === BKA_.HANDLER) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger(BKA_.HANDLER).timeBased().everyHours(BKA_.HORAS).create();
  logEvent('INFO', 'backupAutomaticoInstalar', 'Copia automática cada ' + BKA_.HORAS + ' h instalada', '');
  ui.alert('✅ Backup automático instalado',
    'Primera copia hecha: ' + r.nombre + '\n\nA partir de ahora se hará una copia completa cada ' + BKA_.HORAS +
    ' horas y se guardan las últimas ' + BKA_.MAX_COPIAS + ' (las más antiguas van a la papelera).\n\nCarpeta en Drive:\n' + r.carpeta,
    ui.ButtonSet.OK);
}

function backupAutomaticoAhora() {
  var ui = SpreadsheetApp.getUi();
  try {
    var r = _bka_copiaAhora_();
    ui.alert('✅ Copia hecha', r.nombre + '\n\nCarpeta en Drive:\n' + r.carpeta, ui.ButtonSet.OK);
  } catch (e) {
    logEvent('ERROR', 'backupAutomaticoAhora', e.message, e.stack);
    ui.alert('❌ No se pudo hacer la copia.\n\n' + e.message);
  }
}

function backupAutomaticoDesinstalar() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() === BKA_.HANDLER) { ScriptApp.deleteTrigger(t); n++; } });
  logEvent('INFO', 'backupAutomaticoDesinstalar', 'Triggers eliminados: ' + n, '');
  SpreadsheetApp.getUi().alert(n > 0 ? '✅ Backup automático desactivado. Las copias ya hechas se quedan en Drive.' : 'ℹ️ No había backup automático activo.');
}
