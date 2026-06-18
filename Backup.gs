// ============================================================
// Backup.gs — FASE A: Sistema de backup automático
// Crea copias completas de todas las hojas de datos.
// NO modifica ni elimina datos originales.
// ============================================================

var BACKUP_HOJAS = [
  'CLIENTES', 'TRABAJOS', 'FACTURAS', 'PROFORMAS',
  'PRESUPUESTOS', 'GASTOS', 'INFORMES',
  'RECORDATORIOS_LOG', 'ALERTAS_LOG'
];

// ---- Punto de entrada público ----

/**
 * Crea backup completo de todas las hojas de datos.
 * Retorna { ok, timestamp, prefijo, resumen[] } o lanza excepción si falla.
 */
function ejecutarBackupCompleto() {
  var ui = SpreadsheetApp.getUi();
  var ss = getTargetSS_();
  var tz = Session.getScriptTimeZone();
  var ahora = new Date();
  var timestamp = Utilities.formatDate(ahora, tz, 'yyyyMMdd_HHmm');
  var prefijo   = 'BKP_' + timestamp + '_';

  ui.alert(
    '💾 BACKUP EN CURSO',
    'Creando copia de seguridad...\nPrefijo: ' + prefijo + '\nNo cierres la hoja.',
    ui.ButtonSet.OK
  );

  try {
    var resultado = _bkp_crearBackup_(ss, prefijo, tz, ahora);
    if (!resultado.ok) {
      throw new Error('El backup falló en: ' + resultado.error);
    }
    _bkp_generarInforme_(ss, resultado, timestamp, tz);

    var sh = ss.getSheetByName('BACKUP_REALIZADO');
    if (sh) ss.setActiveSheet(sh);

    ui.alert(
      '✅ BACKUP COMPLETADO',
      'Prefijo de hojas: ' + prefijo + '\n'
      + 'Hojas copiadas: ' + resultado.resumen.length + '\n'
      + 'Total registros respaldados: ' + resultado.totalRegistros + '\n\n'
      + 'Ver hoja: BACKUP_REALIZADO',
      ui.ButtonSet.OK
    );

    return resultado;

  } catch (e) {
    logEvent('ERROR', 'ejecutarBackupCompleto', e.message, e.stack);
    ui.alert('❌ ERROR CRÍTICO EN BACKUP:\n' + e.message + '\n\nNO SE EJECUTARÁ NINGUNA LIMPIEZA.');
    throw e;
  }
}

// ---- Motor de backup ----

function _bkp_crearBackup_(ss, prefijo, tz, ahora) {
  var resultado = {
    ok: true,
    prefijo: prefijo,
    fecha: Utilities.formatDate(ahora, tz, 'dd/MM/yyyy'),
    hora: Utilities.formatDate(ahora, tz, 'HH:mm:ss'),
    resumen: [],
    totalRegistros: 0,
    error: null
  };

  for (var i = 0; i < BACKUP_HOJAS.length; i++) {
    var nombre = BACKUP_HOJAS[i];
    try {
      var shOrigen = ss.getSheetByName(nombre);
      if (!shOrigen) {
        resultado.resumen.push({
          hoja: nombre, estado: 'NO EXISTE', registros: 0, nombreBkp: '—'
        });
        continue;
      }

      var nRegistros = Math.max(0, shOrigen.getLastRow() - 1);
      var abreviatura = _bkp_abreviar_(nombre);
      var nombreBkp   = prefijo + abreviatura;

      // Si ya existe una hoja con ese nombre, eliminarla primero
      var shExistente = ss.getSheetByName(nombreBkp);
      if (shExistente) ss.deleteSheet(shExistente);

      // Copiar la hoja y renombrarla
      var shCopia = shOrigen.copyTo(ss);
      shCopia.setName(nombreBkp);

      // Mover la copia al final de las hojas
      ss.moveActiveSheet(ss.getNumSheets());

      resultado.resumen.push({
        hoja: nombre, estado: 'OK', registros: nRegistros, nombreBkp: nombreBkp
      });
      resultado.totalRegistros += nRegistros;

    } catch (e) {
      resultado.ok    = false;
      resultado.error = nombre + ': ' + e.message;
      resultado.resumen.push({
        hoja: nombre, estado: 'ERROR: ' + e.message, registros: 0, nombreBkp: '—'
      });
      logEvent('ERROR', '_bkp_crearBackup_', 'Fallo en hoja ' + nombre, e.message);
      // Si una hoja falla, abortamos todo para no continuar con datos incompletos
      return resultado;
    }
  }

  logEvent('INFO', '_bkp_crearBackup_',
    'Backup OK — ' + resultado.resumen.length + ' hojas — ' + resultado.totalRegistros + ' registros',
    'Prefijo: ' + prefijo);

  return resultado;
}

function _bkp_abreviar_(nombre) {
  var mapa = {
    'CLIENTES':         'CLI',
    'TRABAJOS':         'TRA',
    'FACTURAS':         'FAC',
    'PROFORMAS':        'PRO',
    'PRESUPUESTOS':     'PRE',
    'GASTOS':           'GAS',
    'INFORMES':         'INF',
    'RECORDATORIOS_LOG':'REC',
    'ALERTAS_LOG':      'ALT'
  };
  return mapa[nombre] || nombre.substring(0, 4);
}

// ---- Informe de backup ----

function _bkp_generarInforme_(ss, resultado, timestamp, tz) {
  var sh = ss.getSheetByName('BACKUP_REALIZADO');
  if (sh) { sh.clear(); sh.clearFormats(); }
  else    { sh = ss.insertSheet('BACKUP_REALIZADO'); }

  sh.setColumnWidth(1, 180); sh.setColumnWidth(2, 100);
  sh.setColumnWidth(3, 130); sh.setColumnWidth(4, 120);
  sh.setColumnWidth(5, 120);

  var r = 1;

  // Título
  sh.getRange(r, 1, 1, 5).merge()
    .setValue('💾 BACKUP REALIZADO — ' + resultado.fecha + ' ' + resultado.hora)
    .setBackground('#137333').setFontColor('#ffffff')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(r, 45); r++;

  sh.getRange(r, 1, 1, 5).merge()
    .setValue('Prefijo de hojas de backup: ' + resultado.prefijo
      + '   |   Estado: ' + (resultado.ok ? '✅ COMPLETADO' : '❌ ERROR'))
    .setBackground('#e6f4ea').setFontColor('#137333').setHorizontalAlignment('center');
  r += 2;

  // Tabla resumen
  sh.getRange(r, 1, 1, 5).setValues([['Hoja original', 'Estado', 'Registros respaldados', 'Nombre hoja backup', '']])
    .setBackground('#202124').setFontColor('#ffffff').setFontWeight('bold');
  r++;

  resultado.resumen.forEach(function(item, i) {
    var bg = item.estado === 'OK' ? (i % 2 === 0 ? '#f0fff4' : '#e8f5e9') : '#ffebee';
    sh.getRange(r, 1, 1, 5).setBackground(bg);
    sh.getRange(r, 1).setValue(item.hoja);
    sh.getRange(r, 2).setValue(item.estado).setFontColor(item.estado === 'OK' ? '#137333' : '#c5221f');
    sh.getRange(r, 3).setValue(item.registros).setHorizontalAlignment('center');
    sh.getRange(r, 4).setValue(item.nombreBkp);
    r++;
  });

  // Total
  sh.getRange(r, 1, 1, 5).setBackground('#e8f0fe');
  sh.getRange(r, 1).setValue('TOTAL').setFontWeight('bold');
  sh.getRange(r, 3).setValue(resultado.totalRegistros).setFontWeight('bold').setHorizontalAlignment('center');
  r += 2;

  // Instrucciones de restauración
  sh.getRange(r, 1, 1, 5).merge()
    .setValue('📋 CÓMO RESTAURAR UN BACKUP')
    .setBackground('#fff3cd').setFontColor('#856404').setFontWeight('bold');
  r++;

  var instrucciones = [
    '1. Localiza la hoja de backup correspondiente (prefijo: ' + resultado.prefijo + ')',
    '2. Haz clic derecho en la pestaña de la hoja de backup → "Copiar a..."',
    '3. O selecciona todos los datos del backup y pégalos en la hoja original',
    '4. ⚠️ NO eliminar las hojas de backup hasta verificar que la restauración es correcta'
  ];
  instrucciones.forEach(function(txt, i) {
    var bg = i % 2 === 0 ? '#fffde7' : '#fff9c4';
    sh.getRange(r, 1, 1, 5).merge().setValue(txt).setBackground(bg).setFontSize(10).setWrap(true);
    r++;
  });

  SpreadsheetApp.flush();
}

// ---- Verificar que existe backup reciente ----

/**
 * Comprueba si existe un backup reciente (último prefijo guardado).
 * Retorna el prefijo o null.
 */
function _bkp_verificarExisteBackup_(ss) {
  var hojas = ss.getSheets().map(function(s) { return s.getName(); });
  var bkpHojas = hojas.filter(function(n) { return /^BKP_\d{8}_\d{4}_/.test(n); });
  if (bkpHojas.length === 0) return null;

  // Extraer el timestamp más reciente
  var timestamps = {};
  bkpHojas.forEach(function(n) {
    var m = n.match(/^(BKP_\d{8}_\d{4}_)/);
    if (m) timestamps[m[1]] = true;
  });
  var prefijos = Object.keys(timestamps).sort();
  return prefijos[prefijos.length - 1]; // más reciente
}
