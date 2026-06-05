// ============================================================
// DriveManager.gs — FASE 18.2A
// Gestión segura de la estructura global de Google Drive
// REGLAS: solo crear/encontrar carpetas. NUNCA borrar, mover
//         ni modificar archivos o carpetas existentes.
// ============================================================

// ── Constantes de estructura ──────────────────────────────────

var DM_ROOT_NAME = 'ALVARO_EMPRESA';
var DM_CFG_SHEET = 'DRIVE_CONFIG';

var DM_YEARS = [2026, 2027, 2028];

// Categorías dentro de cada año
var DM_YEAR_CATS = [
  'FACTURAS', 'GASTOS', 'FOTOS_TRABAJOS', 'PRESUPUESTOS',
  'PROFORMAS', 'INFORMES', 'IVA', 'CONTABILIDAD', 'BACKUP'
];

// Solo estas categorías tienen subcarpetas mensuales
var DM_MONTH_CATS = ['FACTURAS', 'GASTOS', 'FOTOS_TRABAJOS', 'CONTABILIDAD'];

var DM_MONTHS = [
  '01_Enero',    '02_Febrero',  '03_Marzo',    '04_Abril',
  '05_Mayo',     '06_Junio',    '07_Julio',    '08_Agosto',
  '09_Septiembre','10_Octubre','11_Noviembre','12_Diciembre'
];

// Subcarpetas de SISTEMA
var DM_SYSTEM_SUBS = ['CONFIG', 'LOGS', 'EXPORTACIONES', 'TEMP', 'PLANTILLAS'];

// Subcarpetas de cada cliente
var DM_CLIENT_SUBS = ['FACTURAS', 'PRESUPUESTOS', 'PROFORMAS', 'FOTOS', 'INFORMES', 'HISTORIAL'];

// Claves de config para categorías del año en curso (2026)
var DM_CAT_KEYS_2026 = {
  FACTURAS:       'FACTURAS_FOLDER_ID',
  GASTOS:         'GASTOS_FOLDER_ID',
  FOTOS_TRABAJOS: 'FOTOS_TRABAJOS_FOLDER_ID',
  PRESUPUESTOS:   'PRESUPUESTOS_FOLDER_ID',
  PROFORMAS:      'PROFORMAS_FOLDER_ID',
  INFORMES:       'INFORMES_FOLDER_ID',
  IVA:            'IVA_FOLDER_ID',
  CONTABILIDAD:   'CONTABILIDAD_FOLDER_ID',
  BACKUP:         'BACKUP_FOLDER_ID'
};

// ============================================================
// FUNCIONES PÚBLICAS
// ============================================================

function crearEstructuraDriveAlvaro() {
  _crearEstructuraDriveAlvaro_();
}

function diagnosticarDriveAlvaro() {
  _diagnosticarDriveAlvaro_();
}

function actualizarDriveConfig() {
  _actualizarDriveConfig_();
}

// ============================================================
// CREACIÓN DE ESTRUCTURA
// ============================================================

function _crearEstructuraDriveAlvaro_() {
  var ss   = getTargetSS_();
  var log  = [];
  var stat = { created: 0, found: 0, err: 0 };
  var cfg  = {};  // clave → valor para guardar en DRIVE_CONFIG

  function toast(msg) {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, '📁 Drive', 6);
  }

  // Helper seguro: busca la carpeta por nombre bajo parent; la crea si no existe.
  // NUNCA borra ni mueve nada.
  function ensure(parent, name, cfgKey) {
    if (!parent) { stat.err++; log.push('  ❌ Sin parent para: ' + name); return null; }
    try {
      var it = parent.getFoldersByName(name);
      var f;
      if (it.hasNext()) {
        f = it.next();
        stat.found++;
        log.push('  ✓ ' + name);
      } else {
        f = parent.createFolder(name);
        stat.created++;
        log.push('  + ' + name + ' (creada)');
      }
      if (cfgKey) cfg[cfgKey] = f.getId();
      return f;
    } catch (e) {
      stat.err++;
      log.push('  ❌ ' + name + ': ' + e.message);
      return null;
    }
  }

  try {
    // ── 1. Carpeta raíz ───────────────────────────────────────
    toast('1/5 — Carpeta raíz...');
    log.push('RAÍZ:');
    var driveRoot = DriveApp.getRootFolder();
    var alvaro    = ensure(driveRoot, DM_ROOT_NAME, 'ROOT_FOLDER_ID');
    cfg['ROOT_FOLDER_NAME'] = DM_ROOT_NAME;
    if (!alvaro) throw new Error('No se pudo crear/encontrar la carpeta raíz ' + DM_ROOT_NAME);

    // ── 2. SISTEMA ────────────────────────────────────────────
    toast('2/5 — SISTEMA...');
    log.push('\nSISTEMA:');
    var sistema = ensure(alvaro, 'SISTEMA', 'SISTEMA_FOLDER_ID');
    if (sistema) {
      DM_SYSTEM_SUBS.forEach(function(sub) { ensure(sistema, sub, null); });
    }

    // ── 3. CLIENTES (carpeta raíz, sin contenido aún) ─────────
    log.push('\nCLIENTES:');
    var clientesFolder = ensure(alvaro, 'CLIENTES', 'CLIENTES_FOLDER_ID');

    // ── 4. Estructura por años ────────────────────────────────
    toast('3/5 — Años y categorías...');
    DM_YEARS.forEach(function(year) {
      log.push('\nAÑO ' + year + ':');
      var yf = ensure(alvaro, String(year), 'YEAR_' + year + '_FOLDER_ID');
      if (!yf) return;

      DM_YEAR_CATS.forEach(function(cat) {
        var cfgKey = (year === 2026) ? (DM_CAT_KEYS_2026[cat] || null) : null;
        var cf     = ensure(yf, cat, cfgKey);
        if (!cf) return;
        if (DM_MONTH_CATS.indexOf(cat) >= 0) {
          DM_MONTHS.forEach(function(m) { ensure(cf, m, null); });
        }
      });
    });

    // ── 5. Carpetas de clientes ───────────────────────────────
    toast('4/5 — Carpetas de clientes...');
    if (clientesFolder) {
      _dmEnsureClientFolders_(ss, clientesFolder, stat, log, ensure);
    }

    // ── 6. Guardar config en DRIVE_CONFIG ─────────────────────
    toast('5/5 — Guardando configuración...');
    _dmFlushConfig_(ss, cfg);

    // Resultado
    var resumen =
      '✅ Estructura Drive creada.\n\n' +
      '  Carpetas encontradas: ' + stat.found + '\n' +
      '  Carpetas creadas:     ' + stat.created + '\n' +
      (stat.err > 0 ? '  ❌ Errores:           ' + stat.err + '\n' : '') +
      '\nRaíz: ' + DM_ROOT_NAME + ' (' + (cfg['ROOT_FOLDER_ID'] || '—') + ')' +
      '\nNada fue movido ni borrado.';

    Logger.log('=== DRIVE CREACIÓN ===\n' + log.join('\n') + '\n\n' + resumen);
    logEvent('INFO', 'crearEstructuraDriveAlvaro',
      'created=' + stat.created + ' found=' + stat.found, DM_ROOT_NAME);
    SpreadsheetApp.getUi().alert('📁 Drive — Estructura global\n\n' + resumen);

  } catch (e) {
    Logger.log('ERROR crearEstructuraDriveAlvaro: ' + e.message);
    logEvent('ERROR', 'crearEstructuraDriveAlvaro', e.message, '');
    SpreadsheetApp.getUi().alert('❌ Error:\n\n' + e.message);
  }
}

// ============================================================
// CARPETAS POR CLIENTE
// ============================================================

function _dmEnsureClientFolders_(ss, clientesFolder, stat, log, ensure) {
  var clSh = ss.getSheetByName('CLIENTES');
  if (!clSh || clSh.getLastRow() < 2) {
    log.push('  ⚠ Hoja CLIENTES vacía o no existe — omitiendo carpetas de clientes');
    return;
  }

  var lastCol = clSh.getLastColumn();
  var hdrs    = clSh.getRange(1, 1, 1, lastCol).getValues()[0];

  // Detectar columna nombre (COL_CLI.NOMBRE = 3, idx 2)
  var nomIdx = 2;   // Nombre
  var apIdx  = 3;   // Apellidos
  var empIdx = 4;   // Empresa (idx de COL_CLI.EMPRESA - 1)

  var data = clSh.getRange(2, 1, clSh.getLastRow() - 1, lastCol).getValues();
  var seen = {};
  var count = 0;

  data.forEach(function(r) {
    var id     = String(r[0] || '').trim();
    if (!id) return;

    var nombre    = String(r[nomIdx] || '').trim();
    var apellidos = String(r[apIdx]  || '').trim();
    var empresa   = String(r[empIdx] || '').trim();

    var label = empresa
      ? empresa
      : (nombre + (apellidos ? ' ' + apellidos : '')).trim();

    label = _dmSanitizeFolderName_(label);
    if (!label || seen[label]) return;
    seen[label] = true;

    var cf = ensure(clientesFolder, label, null);
    if (!cf) return;
    count++;
    DM_CLIENT_SUBS.forEach(function(sub) { ensure(cf, sub, null); });
  });

  log.push('  → ' + count + ' cliente(s) procesado(s)');
}

// ============================================================
// ACTUALIZAR CONFIG
// ============================================================

/**
 * Lee las IDs ya guardadas en PropertiesService (si hay) y las vuelca
 * a DRIVE_CONFIG actualizando las entradas existentes.
 * También verifica si las carpetas de la estructura existen.
 */
function _actualizarDriveConfig_() {
  var ss  = getTargetSS_();
  var cfg = {};
  var log = [];

  // Verificar raíz
  var rootId = _dmGetConfig_(ss, 'ROOT_FOLDER_ID');
  if (rootId) {
    try {
      var f = DriveApp.getFolderById(rootId);
      cfg['ROOT_FOLDER_ID']   = rootId;
      cfg['ROOT_FOLDER_NAME'] = f.getName();
      log.push('✅ Raíz OK: ' + f.getName());

      // Verificar carpetas principales
      var keys = ['CLIENTES', 'SISTEMA'];
      keys.forEach(function(k) {
        var it = f.getFoldersByName(k);
        if (it.hasNext()) {
          cfg[k + '_FOLDER_ID'] = it.next().getId();
          log.push('✅ ' + k);
        } else {
          log.push('⚠ ' + k + ' no encontrada');
        }
      });

      // Años
      DM_YEARS.forEach(function(year) {
        var yit = f.getFoldersByName(String(year));
        if (yit.hasNext()) {
          var yf = yit.next();
          cfg['YEAR_' + year + '_FOLDER_ID'] = yf.getId();
          log.push('✅ Año ' + year);
          if (year === 2026) {
            Object.keys(DM_CAT_KEYS_2026).forEach(function(cat) {
              var cit = yf.getFoldersByName(cat);
              if (cit.hasNext()) cfg[DM_CAT_KEYS_2026[cat]] = cit.next().getId();
            });
          }
        } else {
          log.push('⚠ Año ' + year + ' no encontrado');
        }
      });
    } catch (e) {
      log.push('❌ ID raíz inválida: ' + e.message);
    }
  } else {
    log.push('⚠ ROOT_FOLDER_ID no configurada — ejecuta crearEstructuraDriveAlvaro primero');
  }

  _dmFlushConfig_(ss, cfg);

  Logger.log(log.join('\n'));
  logEvent('INFO', 'actualizarDriveConfig', log.length + ' entradas procesadas', '');
  SpreadsheetApp.getUi().alert(
    '📁 Drive Config actualizada\n\n' + log.join('\n')
  );
}

// ============================================================
// DIAGNÓSTICO
// ============================================================

function _diagnosticarDriveAlvaro_() {
  var ss   = getTargetSS_();
  var a    = [];
  var warn = [];
  var totalVerified = 0;

  a.push('=== DIAGNÓSTICO DRIVE ÁLVARO ===');
  a.push('Fecha: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'));

  // ── 1. Carpeta raíz ───────────────────────────────────────────
  a.push('\n📁 CARPETA RAÍZ:');
  var alvaro  = null;
  var rootId  = _dmGetConfig_(ss, 'ROOT_FOLDER_ID');

  if (rootId) {
    try {
      alvaro = DriveApp.getFolderById(rootId);
      a.push('  ✅ ' + alvaro.getName() + ' — ID: ' + rootId);
      totalVerified++;
    } catch (e) {
      warn.push('ID de raíz en config no válida (' + rootId + ')');
      rootId = null;
    }
  }
  if (!alvaro) {
    var rit = DriveApp.getRootFolder().getFoldersByName(DM_ROOT_NAME);
    if (rit.hasNext()) {
      alvaro = rit.next();
      a.push('  ✅ ' + DM_ROOT_NAME + ' encontrada por nombre — ID: ' + alvaro.getId());
      warn.push('ROOT_FOLDER_ID no estaba en config — ejecuta actualizarDriveConfig');
      totalVerified++;
    } else {
      a.push('  ⚠ ' + DM_ROOT_NAME + ' no existe aún');
      warn.push('Estructura no creada — ejecuta crearEstructuraDriveAlvaro');
    }
  }

  // ── 2. Años ───────────────────────────────────────────────────
  if (alvaro) {
    a.push('\n📅 AÑOS:');
    DM_YEARS.forEach(function(year) {
      var yit = alvaro.getFoldersByName(String(year));
      if (!yit.hasNext()) {
        a.push('  ⚠ ' + year + ' — no creado');
        warn.push('Año ' + year + ' no existe');
        return;
      }
      var yf = yit.next();
      totalVerified++;
      var missing = [];
      DM_YEAR_CATS.forEach(function(cat) {
        totalVerified++;
        var cit = yf.getFoldersByName(cat);
        if (!cit.hasNext()) missing.push(cat);
        else if (DM_MONTH_CATS.indexOf(cat) >= 0) {
          var cf = cit.next();
          var missingMonths = [];
          DM_MONTHS.forEach(function(m) {
            totalVerified++;
            if (!cf.getFoldersByName(m).hasNext()) missingMonths.push(m);
          });
          if (missingMonths.length > 0) {
            warn.push(year + '/' + cat + ': faltan meses: ' + missingMonths.join(', '));
          }
        }
      });
      if (missing.length > 0) {
        a.push('  ⚠ ' + year + ' — faltan: ' + missing.join(', '));
        warn.push(year + ': categorías faltantes: ' + missing.join(', '));
      } else {
        a.push('  ✅ ' + year + ' — todas las categorías OK');
      }
    });

    // ── 3. CLIENTES ───────────────────────────────────────────
    a.push('\n👤 CLIENTES:');
    var clIt = alvaro.getFoldersByName('CLIENTES');
    if (!clIt.hasNext()) {
      a.push('  ⚠ Carpeta CLIENTES no encontrada');
      warn.push('Carpeta CLIENTES no creada');
    } else {
      var clRoot     = clIt.next();
      var clTotal    = 0, clConCarpeta = 0, clSinCarpeta = [];
      var ss2        = getTargetSS_();
      var clSh       = ss2.getSheetByName('CLIENTES');
      if (clSh && clSh.getLastRow() >= 2) {
        var clData = clSh.getRange(2, 1, clSh.getLastRow() - 1, 5).getValues();
        clData.forEach(function(r) {
          var id     = String(r[0] || '').trim();
          if (!id) return;
          clTotal++;
          var nombre = String(r[2] || '').trim();
          var apell  = String(r[3] || '').trim();
          var emp    = String(r[4] || '').trim();
          var label  = _dmSanitizeFolderName_(emp || (nombre + (apell ? ' '+apell : '')));
          if (!label) return;
          if (clRoot.getFoldersByName(label).hasNext()) {
            clConCarpeta++;
            totalVerified++;
          } else {
            clSinCarpeta.push(label);
          }
        });
        a.push('  ✅ ' + clConCarpeta + '/' + clTotal + ' clientes con carpeta Drive');
        if (clSinCarpeta.length > 0) {
          a.push('  ⚠ Sin carpeta: ' + clSinCarpeta.slice(0, 5).join(', ') +
                 (clSinCarpeta.length > 5 ? ' (+' + (clSinCarpeta.length-5) + ' más)' : ''));
          warn.push(clSinCarpeta.length + ' clientes sin carpeta Drive');
        }
      } else {
        a.push('  ⚪ Hoja CLIENTES vacía — no se verifican carpetas');
      }
    }

    // ── 4. SISTEMA ────────────────────────────────────────────
    a.push('\n⚙️ SISTEMA:');
    var sitIt = alvaro.getFoldersByName('SISTEMA');
    if (!sitIt.hasNext()) {
      a.push('  ⚠ Carpeta SISTEMA no encontrada');
    } else {
      var sif = sitIt.next();
      var missSub = DM_SYSTEM_SUBS.filter(function(s) {
        return !sif.getFoldersByName(s).hasNext();
      });
      if (missSub.length === 0) {
        a.push('  ✅ Todas las subcarpetas de SISTEMA OK');
      } else {
        a.push('  ⚠ Faltan en SISTEMA: ' + missSub.join(', '));
        warn.push('SISTEMA: faltan subcarpetas: ' + missSub.join(', '));
      }
    }
  }

  // ── 5. Compatibilidad PANEL_CONTABLE ─────────────────────────
  a.push('\n📦 COMPATIBILIDAD PANEL_CONTABLE:');
  var oldIt = DriveApp.getRootFolder().getFoldersByName('Contabilidad');
  if (oldIt.hasNext()) {
    var of = oldIt.next();
    a.push('  ✅ Carpeta contable antigua detectada: "Contabilidad" (' + of.getId() + ')');
    a.push('     No modificada. PANEL_CONTABLE sigue funcionando.');
  } else {
    a.push('  ⚪ No existe carpeta "Contabilidad" en raíz (se crea al generar el primer paquete)');
  }
  var newContabId = _dmGetConfig_(ss, 'CONTABILIDAD_FOLDER_ID');
  if (newContabId) {
    a.push('  ✅ Nueva estructura ALVARO_EMPRESA/2026/CONTABILIDAD preparada');
  } else {
    a.push('  ⚪ Nueva estructura contable pendiente (ejecuta crearEstructuraDriveAlvaro)');
  }

  // ── 6. Resumen ────────────────────────────────────────────────
  a.push('\n📊 RESUMEN:');
  a.push('  Total carpetas verificadas: ' + totalVerified);
  a.push('  Advertencias: ' + warn.length);

  a.push('');
  var status = warn.length === 0
    ? '✅ DRIVE ALVARO listo'
    : '⚠ DRIVE ALVARO con advertencias — ' + warn.length + ' incidencia(s)';
  a.push(status);
  if (warn.length > 0) {
    warn.forEach(function(w) { a.push('  · ' + w); });
  }

  var report = a.join('\n');
  Logger.log(report);
  logEvent('INFO', 'diagnosticarDriveAlvaro', status, totalVerified + ' verificadas');
  SpreadsheetApp.getUi().alert('📁 Diagnóstico DRIVE ÁLVARO\n\n' + report);
}

// ============================================================
// DRIVE_CONFIG — Hoja de configuración
// ============================================================

/** Garantiza que la hoja DRIVE_CONFIG existe con cabeceras. */
function _dmEnsureConfigSheet_(ss) {
  var sh = ss.getSheetByName(DM_CFG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DM_CFG_SHEET);
    var hdr = sh.getRange(1, 1, 1, 4);
    hdr.setValues([['Clave', 'Valor', 'Descripción', 'Fecha_Actualizacion']]);
    hdr.setBackground('#4a148c').setFontColor('#ffffff')
       .setFontWeight('bold').setFontSize(11);
    sh.setColumnWidth(1, 220);
    sh.setColumnWidth(2, 340);
    sh.setColumnWidth(3, 240);
    sh.setColumnWidth(4, 150);
    sh.setFrozenRows(1);
  }
  return sh;
}

/**
 * Escribe en bloque un objeto {clave: valor} en DRIVE_CONFIG.
 * Actualiza filas existentes; añade las nuevas. Operación segura.
 */
function _dmFlushConfig_(ss, cfg) {
  if (!cfg || Object.keys(cfg).length === 0) return;

  var sh   = _dmEnsureConfigSheet_(ss);
  var tz   = Session.getScriptTimeZone();
  var ts   = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  var rows = sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()
    : [];

  // Index existing rows by key
  var idx = {};
  rows.forEach(function(r, i) { if (r[0]) idx[String(r[0]).trim()] = i; });

  Object.keys(cfg).forEach(function(key) {
    var val  = String(cfg[key] || '');
    var desc = _dmCfgDesc_(key);
    if (key in idx) {
      // Update existing row
      sh.getRange(idx[key] + 2, 2, 1, 3).setValues([[val, desc, ts]]);
    } else {
      // Append new row
      sh.appendRow([key, val, desc, ts]);
      idx[key] = sh.getLastRow() - 2; // keep index consistent
    }
  });
}

/** Lee un valor de DRIVE_CONFIG por clave. */
function _dmGetConfig_(ss, key) {
  var sh = ss.getSheetByName(DM_CFG_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) return data[i][1] || null;
  }
  return null;
}

/** Devuelve una descripción legible para cada clave de config. */
function _dmCfgDesc_(key) {
  var map = {
    ROOT_FOLDER_ID:        'Carpeta raíz ALVARO_EMPRESA',
    ROOT_FOLDER_NAME:      'Nombre carpeta raíz',
    SISTEMA_FOLDER_ID:     'Carpeta SISTEMA (config, logs, etc.)',
    CLIENTES_FOLDER_ID:    'Carpeta raíz de clientes',
    YEAR_2026_FOLDER_ID:   'Carpeta año 2026',
    YEAR_2027_FOLDER_ID:   'Carpeta año 2027',
    YEAR_2028_FOLDER_ID:   'Carpeta año 2028',
    FACTURAS_FOLDER_ID:    'Facturas emitidas 2026',
    GASTOS_FOLDER_ID:      'Gastos 2026',
    FOTOS_TRABAJOS_FOLDER_ID: 'Fotos de trabajos 2026',
    PRESUPUESTOS_FOLDER_ID:'Presupuestos 2026',
    PROFORMAS_FOLDER_ID:   'Proformas 2026',
    INFORMES_FOLDER_ID:    'Informes 2026',
    IVA_FOLDER_ID:         'Declaraciones IVA 2026',
    CONTABILIDAD_FOLDER_ID:'Contabilidad 2026',
    BACKUP_FOLDER_ID:      'Backups 2026'
  };
  return map[key] || key;
}

// ============================================================
// HELPERS
// ============================================================

/**
 * Sanitiza un nombre para usarlo como nombre de carpeta Drive.
 * Solo elimina caracteres inválidos; no borra contenido existente.
 */
function _dmSanitizeFolderName_(name) {
  return String(name || '').trim()
    .replace(/[\/\\:*?"<>|#@]/g, '-')
    .replace(/\s{2,}/g, ' ')
    .substring(0, 80)
    .trim();
}

/** Escribe en el log de Apps Script. */
function _dmLog_(msg) {
  Logger.log('[DriveManager] ' + msg);
}
