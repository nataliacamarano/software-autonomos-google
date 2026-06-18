// ============================================================
// RecuperarCarpetas.gs — Recuperación y gestión de carpetas Drive
//
// REGLAS:
//   - NUNCA borra ni mueve carpetas ni archivos.
//   - NUNCA crea carpetas raíz (FACTURAS, GASTOS…); solo las busca.
//   - Crea subcarpetas AÑO/MES si faltan dentro de una raíz ya existente.
//   - Escribe IDs en la hoja CONFIG (no en DRIVE_CONFIG).
// ============================================================

var RC_ = {
  // Mapping de carpetas → claves en CONFIG
  CATEGORIAS: [
    { nombre: 'FACTURAS',     key: 'carpeta_facturas_id',     label: 'Facturas'     },
    { nombre: 'PROFORMAS',    key: 'carpeta_proformas_id',    label: 'Proformas'    },
    { nombre: 'PRESUPUESTOS', key: 'carpeta_presupuestos_id', label: 'Presupuestos' },
    { nombre: 'TRABAJOS',     key: 'carpeta_trabajos_id',     label: 'Trabajos'     },
    { nombre: 'GASTOS',       key: 'gastos_folder_id',        label: 'Gastos'       },
    { nombre: 'INFORMES',     key: 'carpeta_informes_id',     label: 'Informes'     }
  ],
  MESES: [
    '01_Enero','02_Febrero','03_Marzo','04_Abril',
    '05_Mayo','06_Junio','07_Julio','08_Agosto',
    '09_Septiembre','10_Octubre','11_Noviembre','12_Diciembre'
  ],
  ANOS_ESPERADOS: ['2025', '2026', '2027', '2028']
};

// ── Punto de entrada ─────────────────────────────────────────

/**
 * Busca en Drive las carpetas FACTURAS, PROFORMAS, PRESUPUESTOS, TRABAJOS,
 * GASTOS e INFORMES. Registra sus IDs en CONFIG. No crea carpetas raíz.
 * Si hay varias coincidencias sin estructura clara, reporta sin elegir.
 * Al terminar crea la hoja DIAGNOSTICO_CARPETAS_DRIVE y muestra un resumen.
 */
function recuperarCarpetasDrive() {
  var ss  = getTargetSS_();
  var ui  = SpreadsheetApp.getUi();
  var ts  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  var log = [];  // filas para la hoja diagnóstico

  var resumen = { asignadas: 0, multiples: [], noEncontradas: [], yaConfig: 0 };

  log.push(['DIAGNÓSTICO CARPETAS DRIVE — ' + ts, '', '', '', '', '']);
  log.push(['Tipo', 'Carpeta encontrada', 'ID', 'Ruta', 'Estado', 'Acción']);

  RC_.CATEGORIAS.forEach(function(cat) {
    var existingId = getConfig(cat.key);

    // Si ya está configurada, verificar que sigue siendo accesible
    if (existingId) {
      try {
        var ef = DriveApp.getFolderById(existingId);
        log.push([cat.label, ef.getName(), existingId, _rc_buildPath_(ef), '✓ Ya configurada', 'Sin cambios']);
        resumen.yaConfig++;
        return;
      } catch (_) {
        Logger.log('[RC] ID inválida para ' + cat.nombre + ': ' + existingId + ' — buscando de nuevo');
      }
    }

    // Buscar por nombre en todo Drive del usuario
    var coincidencias = _rc_buscarCarpetasPorNombre_(cat.nombre);
    Logger.log('[RC] "' + cat.nombre + '" — ' + coincidencias.length + ' coincidencia(s)');

    if (coincidencias.length === 0) {
      log.push([cat.label, '(no encontrada)', '', '', '❌ No encontrada',
        'Crea la carpeta en Drive y vuelve a ejecutar']);
      resumen.noEncontradas.push(cat.label);
      return;
    }

    if (coincidencias.length === 1) {
      var f = coincidencias[0];
      setConfig(cat.key, f.getId());
      var ruta = _rc_buildPath_(f);
      log.push([cat.label, f.getName(), f.getId(), ruta, '✅ Asignada', 'ID guardado en CONFIG']);
      resumen.asignadas++;
      Logger.log('[RC] ✅ ' + cat.nombre + ' → ' + f.getId());
      return;
    }

    // Múltiples: intentar filtrar por estructura de años
    var conEstructura = coincidencias.filter(_rc_tieneEstructuraAnios_);
    Logger.log('[RC] ' + cat.nombre + ' con estructura años: ' + conEstructura.length);

    if (conEstructura.length === 1) {
      var f2 = conEstructura[0];
      setConfig(cat.key, f2.getId());
      log.push([cat.label, f2.getName(), f2.getId(), _rc_buildPath_(f2),
        '✅ Asignada (por estructura AÑO/MES)', 'ID guardado en CONFIG']);
      resumen.asignadas++;
      Logger.log('[RC] ✅ ' + cat.nombre + ' (filtrado) → ' + f2.getId());
      return;
    }

    // Sigue siendo ambiguo — no elegir, reportar IDs para revisión manual
    var detalle = coincidencias.map(function(f) {
      return f.getId() + ' (' + _rc_buildPath_(f) + ')';
    }).join('\n  ');
    log.push([cat.label, coincidencias.length + ' coincidencias', detalle, '',
      '⚠️ Ambiguo — no asignado', 'Añade el ID correcto manualmente en CONFIG\nclave: ' + cat.key]);
    resumen.multiples.push(cat.label + ' (' + coincidencias.length + ')');
    Logger.log('[RC] ⚠️ ' + cat.nombre + ' — ambiguo:\n  ' + detalle);
  });

  // Asegurar estructura AÑO/MES en las carpetas ya configuradas
  log.push(['', '', '', '', '', '']);
  log.push(['=== ESTRUCTURA AÑO/MES ===', '', '', '', '', '']);
  _rc_asegurarEstructuraAnioMes_(log);

  // Crear hoja diagnóstico
  _rc_crearHojaDiagnostico_(ss, log, ts);

  // Resumen
  var msg = '📁 RECUPERACIÓN DE CARPETAS DRIVE\n\n' +
    '✅ Asignadas ahora: '    + resumen.asignadas  + '\n' +
    '🔄 Ya configuradas: '   + resumen.yaConfig   + '\n';
  if (resumen.multiples.length)
    msg += '⚠️ Ambiguas (revisar): ' + resumen.multiples.join(', ') + '\n';
  if (resumen.noEncontradas.length)
    msg += '❌ No encontradas:     ' + resumen.noEncontradas.join(', ') + '\n';
  msg += '\nVer hoja DIAGNOSTICO_CARPETAS_DRIVE para el informe completo.';

  logEvent('INFO', 'recuperarCarpetasDrive',
    'asignadas=' + resumen.asignadas + ' yaConfig=' + resumen.yaConfig +
    ' multiples=' + resumen.multiples.length + ' noEncontradas=' + resumen.noEncontradas.length, '');

  ui.alert('📁 Drive — Recuperación completada', msg, ui.ButtonSet.OK);
}

/**
 * Muestra en la hoja DIAGNOSTICO_CARPETAS_DRIVE el estado actual de cada carpeta
 * configurada (o no) en CONFIG, sin modificar nada.
 */
function diagnosticarCarpetasDrive() {
  var ss  = getTargetSS_();
  var ts  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  var log = [];

  log.push(['DIAGNÓSTICO CARPETAS DRIVE — ' + ts, '', '', '', '', '']);
  log.push(['Tipo', 'Clave CONFIG', 'ID configurado', 'Estado acceso', 'Ruta en Drive', 'Estructura AÑO/MES']);

  RC_.CATEGORIAS.forEach(function(cat) {
    var id  = getConfig(cat.key);
    var fila;

    if (!id) {
      fila = [cat.label, cat.key, '(no configurado)', '❌ Sin ID', '—', '—'];
    } else {
      try {
        var f       = DriveApp.getFolderById(id);
        var ruta    = _rc_buildPath_(f);
        var tieneAn = _rc_tieneEstructuraAnios_(f) ? '✓ Tiene años' : '⚠️ Sin subcarpetas de año';
        fila = [cat.label, cat.key, id, '✅ Accesible', ruta, tieneAn];
      } catch (e) {
        fila = [cat.label, cat.key, id, '❌ ID no accesible: ' + e.message, '—', '—'];
      }
    }
    log.push(fila);
  });

  // Estado de las carpetas AÑO/MES del año actual
  log.push(['', '', '', '', '', '']);
  log.push(['=== SUBCARPETAS AÑO/MES (año actual) ===', '', '', '', '', '']);
  _rc_asegurarEstructuraAnioMes_(log, true);  // true = solo diagnóstico, no crea

  _rc_crearHojaDiagnostico_(ss, log, ts);
  SpreadsheetApp.getUi().alert(
    '📁 Diagnóstico completado',
    'Revisa la hoja DIAGNOSTICO_CARPETAS_DRIVE para ver el estado de cada carpeta.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ── Función pública usada por los generadores de PDF ─────────

/**
 * Devuelve la subcarpeta AÑO/MES dentro de una carpeta raíz.
 * La crea si no existe. Nunca crea la carpeta raíz.
 *
 * @param {string} rootFolderId  ID de la carpeta raíz (FACTURAS, PROFORMAS, etc.)
 * @param {Date}   fecha         Fecha del documento (determina año y mes)
 * @returns {Folder}             Carpeta de destino: ROOT/YYYY/MM_Mes
 */
function _getOrCreateSubcarpetaAnioMes_(rootFolderId, fecha) {
  var year   = String(fecha.getFullYear());
  var mesStr = RC_.MESES[fecha.getMonth()];

  var rootFolder = DriveApp.getFolderById(rootFolderId);

  var yearIt     = rootFolder.getFoldersByName(year);
  var yearFolder = yearIt.hasNext() ? yearIt.next() : rootFolder.createFolder(year);

  var mesIt     = yearFolder.getFoldersByName(mesStr);
  var mesFolder = mesIt.hasNext() ? mesIt.next() : yearFolder.createFolder(mesStr);

  Logger.log('  📁 Destino: ' + rootFolder.getName() + '/' + year + '/' + mesStr);
  return mesFolder;
}

// ── Helpers privados ─────────────────────────────────────────

function _rc_buscarCarpetasPorNombre_(nombre) {
  var results = [];
  try {
    var it = DriveApp.getFoldersByName(nombre);
    while (it.hasNext()) results.push(it.next());
  } catch (e) {
    Logger.log('[RC] Error buscando "' + nombre + '": ' + e.message);
  }
  return results;
}

function _rc_tieneEstructuraAnios_(folder) {
  try {
    for (var i = 0; i < RC_.ANOS_ESPERADOS.length; i++) {
      if (folder.getFoldersByName(RC_.ANOS_ESPERADOS[i]).hasNext()) return true;
    }
  } catch (_) {}
  return false;
}

function _rc_buildPath_(folder) {
  try {
    var parts   = [folder.getName()];
    var current = folder;
    var rootId  = DriveApp.getRootFolder().getId();
    for (var i = 0; i < 4; i++) {
      var parents = current.getParents();
      if (!parents.hasNext()) break;
      var p = parents.next();
      if (p.getId() === rootId) { parts.unshift('Mi unidad'); break; }
      parts.unshift(p.getName());
      current = p;
    }
    return parts.join(' / ');
  } catch (_) {
    try { return folder.getName(); } catch (_2) { return '—'; }
  }
}

/**
 * Verifica (y opcionalmente crea) las subcarpetas AÑO/MES del año actual
 * para cada carpeta raíz ya configurada en CONFIG.
 * @param {Array}   log       Filas a añadir al informe
 * @param {boolean} soloLeer  Si true, solo verifica (no crea subcarpetas)
 */
function _rc_asegurarEstructuraAnioMes_(log, soloLeer) {
  var hoy    = new Date();
  var year   = String(hoy.getFullYear());
  var mesStr = RC_.MESES[hoy.getMonth()];
  var cols   = 6;

  RC_.CATEGORIAS.forEach(function(cat) {
    var rootId = getConfig(cat.key);
    if (!rootId) {
      log.push([cat.label, '—', '—', '—', '⚪ Sin ID configurado', '—']);
      return;
    }
    try {
      var root    = DriveApp.getFolderById(rootId);
      var yearIt  = root.getFoldersByName(year);
      var yearOk  = yearIt.hasNext();
      var yearF;
      if (yearOk) {
        yearF = yearIt.next();
      } else if (!soloLeer) {
        yearF = root.createFolder(year);
        Logger.log('[RC] Creada carpeta año: ' + root.getName() + '/' + year);
      }

      var mesOk = false;
      if (yearF) {
        var mesIt = yearF.getFoldersByName(mesStr);
        mesOk = mesIt.hasNext();
        if (!mesOk && !soloLeer) {
          yearF.createFolder(mesStr);
          Logger.log('[RC] Creada carpeta mes: ' + root.getName() + '/' + year + '/' + mesStr);
        }
      }

      var estado  = (yearOk && mesOk) ? '✅ Completa' : (soloLeer ? '⚠️ Falta estructura' : '✅ Creada');
      var detalle = root.getName() + '/' + year + (yearOk ? '' : ' (nueva)') +
                    '/' + mesStr + (mesOk ? '' : ' (nueva)');
      log.push([cat.label, cat.key, rootId, detalle, estado,
        soloLeer ? (yearOk && mesOk ? 'Existe' : 'Falta — ejecuta recuperarCarpetasDrive') : 'Verificado/creado']);
    } catch (e) {
      log.push([cat.label, cat.key, rootId, 'Error: ' + e.message, '❌ Error', 'Verificar ID']);
    }
  });
}

function _rc_crearHojaDiagnostico_(ss, log, ts) {
  var sh = getOrCreateSheet('DIAGNOSTICO_CARPETAS_DRIVE');
  sh.clearContents();
  sh.clearFormats();

  var ncols = 6;
  sh.getRange(1, 1, log.length, ncols).setValues(log);

  // Estilo filas de cabecera
  sh.getRange(1, 1, 1, ncols)
    .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange(2, 1, 1, ncols)
    .setBackground('#4285f4').setFontColor('#ffffff').setFontWeight('bold');

  // Color por estado
  for (var r = 2; r < log.length; r++) {
    var estado = String(log[r][4] || '') + String(log[r][0] || '');
    var bg;
    if (estado.indexOf('✅') >= 0 || estado.indexOf('✓') >= 0) bg = '#e6f4ea';
    else if (estado.indexOf('❌') >= 0)                         bg = '#fce8e6';
    else if (estado.indexOf('⚠️') >= 0)                        bg = '#fff3e0';
    else if (estado.indexOf('===') >= 0)                        bg = '#e8eaed';
    if (bg) sh.getRange(r + 1, 1, 1, ncols).setBackground(bg);
  }

  sh.autoResizeColumns(1, ncols);
  ss.setActiveSheet(sh);
}
