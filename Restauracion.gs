// ================================================================
// Restauracion.gs — Restauración segura desde hojas BKP_*
//
// REGLAS ABSOLUTAS:
//   - NO eliminar nada (ni en hojas actuales ni en backups)
//   - NO modificar las hojas de backup
//   - Solo AÑADIR filas que estaban en el backup y ya no están
//   - Comparar por ID (col 1) si existe; si no, por hash de fila
//   - Todo lo dudoso va a RESTAURACION_REVISION_MANUAL
//
// HOJAS QUE SE RESTAURAN:
//   CLIENTES, TRABAJOS, FACTURAS, PROFORMAS,
//   PRESUPUESTOS, GASTOS, INFORMES
// ================================================================

var RST_HOJAS = [
  { nombre: 'CLIENTES',     abrev: 'CLI', idCol: 1, nomCol: [3, 4] },
  { nombre: 'TRABAJOS',     abrev: 'TRA', idCol: 1, nomCol: [4]    },
  { nombre: 'FACTURAS',     abrev: 'FAC', idCol: 1, nomCol: [6]    },
  { nombre: 'PROFORMAS',    abrev: 'PRO', idCol: 1, nomCol: [6]    },
  { nombre: 'PRESUPUESTOS', abrev: 'PRE', idCol: 1, nomCol: [3]    },
  { nombre: 'GASTOS',       abrev: 'GAS', idCol: 1, nomCol: [5]    },
  { nombre: 'INFORMES',     abrev: 'INF', idCol: 1, nomCol: [4]    }
];

// ================================================================
// PUNTO DE ENTRADA PRINCIPAL
// ================================================================

function ejecutarRestauracionDesdeBackup() {
  var ui = SpreadsheetApp.getUi();
  var ss = getTargetSS_();

  // ── Localizar el backup más reciente ────────────────────────
  var prefijo = _rst_encontrarPrefijoBkpMasReciente_(ss);

  if (!prefijo) {
    ui.alert('❌ No se encontró ningún backup',
      'No hay hojas con el prefijo BKP_YYYYMMDD_HHMM_ en este spreadsheet.\n\n'
      + 'Sin backup no es posible restaurar datos automáticamente.',
      ui.ButtonSet.OK);
    return;
  }

  ui.alert(
    '🔍 BACKUP ENCONTRADO',
    'Backup más reciente localizado:\n\n'
    + '  Prefijo: ' + prefijo + '\n\n'
    + 'Se generará primero el diagnóstico (RESTAURACION_DIAGNOSTICO)\n'
    + 'para que puedas ver qué se recuperaría antes de aplicar nada.\n\n'
    + '¿Continuar con el diagnóstico?',
    ui.ButtonSet.OK
  );

  // ── PASO 1: Diagnóstico ──────────────────────────────────────
  var diagnostico = _rst_generarDiagnostico_(ss, prefijo);

  var shDiag = ss.getSheetByName('RESTAURACION_DIAGNOSTICO');
  if (shDiag) ss.setActiveSheet(shDiag);

  var resumen = _rst_resumenDiagnostico_(diagnostico);
  var conf = ui.alert(
    '📋 DIAGNÓSTICO GENERADO — RESTAURACION_DIAGNOSTICO',
    'Filas a restaurar por hoja:\n\n' + resumen
    + '\n\n⚠️ Solo se AÑADIRÁN filas. No se elimina nada.\n'
    + '¿Ejecutar la restauración ahora?',
    ui.ButtonSet.YES_NO
  );
  if (conf !== ui.Button.YES) {
    ui.alert('Restauración cancelada. El diagnóstico está en RESTAURACION_DIAGNOSTICO.');
    return;
  }

  // ── PASO 2: Restaurar ────────────────────────────────────────
  var logRestaurado = [];
  var logRevision   = [];

  // Orden: CLIENTES primero (para integridad referencial)
  RST_HOJAS.forEach(function(cfg) {
    _rst_restaurarHoja_(ss, cfg, prefijo, logRestaurado, logRevision);
  });

  SpreadsheetApp.flush();

  // ── PASO 3: Generar hojas de resultado ───────────────────────
  _rst_generarHojaRealizada_(ss, logRestaurado, prefijo);
  _rst_generarHojaRevision_(ss, logRevision);

  // ── PASO 4: Sincronizar sistema ──────────────────────────────
  try { sincronizarClientesSistema(); } catch (e) {
    logEvent('WARN', 'ejecutarRestauracionDesdeBackup',
      'sincronizarClientesSistema falló: ' + e.message, '');
  }
  try { actualizarPanelGeneral(); } catch (e) { /* panel opcional */ }

  var shFinal = ss.getSheetByName('RESTAURACION_REALIZADA');
  if (shFinal) ss.setActiveSheet(shFinal);

  ui.alert(
    '✅ RESTAURACIÓN COMPLETADA',
    'Filas restauradas:         ' + logRestaurado.length + '\n'
    + 'Filas en revisión manual:  ' + logRevision.length + '\n\n'
    + 'Ver:\n'
    + '  • RESTAURACION_REALIZADA\n'
    + '  • RESTAURACION_REVISION_MANUAL\n\n'
    + 'Se han sincronizado los desplegables de clientes.',
    ui.ButtonSet.OK
  );
}

// ================================================================
// LOCALIZAR BACKUP MÁS RECIENTE
// ================================================================

function _rst_encontrarPrefijoBkpMasReciente_(ss) {
  var hojas = ss.getSheets().map(function(s) { return s.getName(); });
  var timestamps = {};

  hojas.forEach(function(nombre) {
    var m = nombre.match(/^(BKP_\d{8}_\d{4}_)/);
    if (m) timestamps[m[1]] = true;
  });

  var prefijos = Object.keys(timestamps).sort();
  return prefijos.length > 0 ? prefijos[prefijos.length - 1] : null;
}

// ================================================================
// DIAGNÓSTICO
// ================================================================

function _rst_generarDiagnostico_(ss, prefijo) {
  var diagnostico = [];

  RST_HOJAS.forEach(function(cfg) {
    var nombreBkp = prefijo + cfg.abrev;
    var shActual  = ss.getSheetByName(cfg.nombre);
    var shBkp     = ss.getSheetByName(nombreBkp);

    var filasActual = shActual ? Math.max(0, shActual.getLastRow() - 1) : 0;
    var filasBkp    = shBkp    ? Math.max(0, shBkp.getLastRow()    - 1) : 0;
    var diferencia  = filasBkp - filasActual;

    // Contar exactamente cuántas filas del backup faltan en la hoja actual
    var filasFaltantes = 0;
    var errorHeaders   = '';

    if (shBkp && shActual && shBkp.getLastRow() >= 2 && shActual.getLastRow() >= 1) {
      try {
        // Verificar que los headers son compatibles
        var hdrActual = shActual.getRange(1, 1, 1, shActual.getLastColumn()).getValues()[0]
                          .map(function(v) { return String(v).trim(); });
        var hdrBkp    = shBkp.getRange(1, 1, 1, shBkp.getLastColumn()).getValues()[0]
                          .map(function(v) { return String(v).trim(); });

        if (hdrActual.join('|') !== hdrBkp.join('|')) {
          errorHeaders = '⚠️ Cabeceras distintas entre hoja y backup';
        }

        // Índice de IDs actuales
        var idsActuales = _rst_construirIndiceIds_(shActual);
        var hashesActuales = _rst_construirIndiceHashes_(shActual);

        var nColsBkp = shBkp.getLastColumn();
        var dataBkp  = shBkp.getLastRow() >= 2
          ? shBkp.getRange(2, 1, shBkp.getLastRow() - 1, nColsBkp).getValues()
          : [];

        dataBkp.forEach(function(row) {
          var id   = String(row[cfg.idCol - 1] || '').trim();
          var hash = _rst_hashFila_(row);
          if (id && !idsActuales[id]) {
            filasFaltantes++;
          } else if (!id && !hashesActuales[hash]) {
            filasFaltantes++;
          }
        });
      } catch (e) {
        errorHeaders = '❌ Error al comparar: ' + e.message;
      }
    } else if (!shBkp) {
      errorHeaders = '⚠️ Backup no encontrado';
    }

    var recomendacion;
    if (!shBkp) {
      recomendacion = 'Sin backup disponible — no se puede restaurar';
    } else if (filasFaltantes === 0) {
      recomendacion = '✅ Sin diferencias — no se necesita restaurar';
    } else {
      recomendacion = '🔴 RESTAURAR ' + filasFaltantes + ' fila(s) faltante(s)';
    }

    diagnostico.push({
      nombre: cfg.nombre, nombreBkp: nombreBkp,
      filasActual: filasActual, filasBkp: filasBkp,
      diferencia: diferencia, filasFaltantes: filasFaltantes,
      errorHeaders: errorHeaders, recomendacion: recomendacion
    });
  });

  // Generar hoja RESTAURACION_DIAGNOSTICO
  var sh = ss.getSheetByName('RESTAURACION_DIAGNOSTICO');
  if (sh) { sh.clear(); sh.clearFormats(); }
  else    { sh = ss.insertSheet('RESTAURACION_DIAGNOSTICO'); }

  var tz    = Session.getScriptTimeZone();
  var ahora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');

  sh.setColumnWidth(1, 175); sh.setColumnWidth(2, 80);
  sh.setColumnWidth(3, 185); sh.setColumnWidth(4, 80);
  sh.setColumnWidth(5, 85);  sh.setColumnWidth(6, 100);
  sh.setColumnWidth(7, 300);

  var r = 1;
  sh.getRange(r, 1, 1, 7).merge()
    .setValue('🔍 DIAGNÓSTICO DE RESTAURACIÓN — ' + ahora)
    .setBackground('#b71c1c').setFontColor('#ffffff')
    .setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(r, 42); r++;

  sh.getRange(r, 1, 1, 7).merge()
    .setValue('Backup utilizado: ' + prefijo)
    .setBackground('#ffebee').setFontColor('#b71c1c').setFontWeight('bold').setHorizontalAlignment('center');
  r += 2;

  sh.getRange(r, 1, 1, 7).setValues([[
    'Hoja actual', 'Filas act.', 'Hoja backup', 'Filas bkp',
    'Diferencia', 'Faltan', 'Recomendación'
  ]]).setBackground('#202124').setFontColor('#ffffff').setFontWeight('bold');
  r++;

  diagnostico.forEach(function(d, i) {
    var bg = i % 2 === 0 ? '#f9f9f9' : '#f0f0f0';
    var colorFalt = d.filasFaltantes > 0 ? '#c5221f' : '#137333';
    sh.getRange(r, 1, 1, 7).setBackground(bg);
    sh.getRange(r, 1).setValue(d.nombre).setFontWeight('bold');
    sh.getRange(r, 2).setValue(d.filasActual).setHorizontalAlignment('center');
    sh.getRange(r, 3).setValue(d.nombreBkp).setFontSize(9);
    sh.getRange(r, 4).setValue(d.filasBkp).setHorizontalAlignment('center');
    sh.getRange(r, 5).setValue(d.diferencia).setHorizontalAlignment('center')
      .setFontColor(d.diferencia > 0 ? '#c5221f' : '#137333');
    sh.getRange(r, 6).setValue(d.filasFaltantes).setHorizontalAlignment('center')
      .setFontColor(colorFalt).setFontWeight('bold');
    sh.getRange(r, 7).setValue((d.errorHeaders ? d.errorHeaders + ' | ' : '') + d.recomendacion)
      .setWrap(true).setFontSize(9);
    sh.setRowHeight(r, 32); r++;
  });

  SpreadsheetApp.flush();
  return diagnostico;
}

function _rst_resumenDiagnostico_(diagnostico) {
  var lineas = [];
  diagnostico.forEach(function(d) {
    if (d.filasFaltantes > 0) {
      lineas.push('  • ' + d.nombre + ': ' + d.filasFaltantes + ' fila(s) a restaurar');
    } else {
      lineas.push('  • ' + d.nombre + ': sin diferencias');
    }
  });
  return lineas.join('\n');
}

// ================================================================
// RESTAURAR UNA HOJA DESDE SU BACKUP
// ================================================================

function _rst_restaurarHoja_(ss, cfg, prefijo, logRestaurado, logRevision) {
  var nombreBkp = prefijo + cfg.abrev;
  var shBkp     = ss.getSheetByName(nombreBkp);
  var shActual  = ss.getSheetByName(cfg.nombre);

  if (!shBkp) {
    logEvent('WARN', '_rst_restaurarHoja_', 'Backup no encontrado: ' + nombreBkp, '');
    return;
  }
  if (!shActual) {
    logEvent('WARN', '_rst_restaurarHoja_', 'Hoja actual no encontrada: ' + cfg.nombre, '');
    return;
  }
  if (shBkp.getLastRow() < 2) return; // backup vacío (solo cabecera o nada)

  // Verificar compatibilidad de cabeceras
  var nColsActual = shActual.getLastColumn();
  var nColsBkp    = shBkp.getLastColumn();

  var hdrActual = shActual.getLastRow() >= 1
    ? shActual.getRange(1, 1, 1, nColsActual).getValues()[0].map(function(v) { return String(v).trim(); })
    : [];
  var hdrBkp = shBkp.getRange(1, 1, 1, nColsBkp).getValues()[0]
    .map(function(v) { return String(v).trim(); });

  var headersMismatch = hdrActual.join('|') !== hdrBkp.join('|');

  // Construir índices de la hoja actual
  var idsActuales    = _rst_construirIndiceIds_(shActual);
  var hashesActuales = _rst_construirIndiceHashes_(shActual);

  // Leer todas las filas del backup
  var dataBkp = shBkp.getRange(2, 1, shBkp.getLastRow() - 1, nColsBkp).getValues();

  dataBkp.forEach(function(row, idx) {
    var filaBkp = idx + 2;
    var id      = String(row[cfg.idCol - 1] || '').trim();
    var hash    = _rst_hashFila_(row);

    // Nombre del cliente/contacto para el log
    var nomMostrar = cfg.nomCol.map(function(c) {
      return String(row[c - 1] || '').trim();
    }).filter(function(v) { return v; }).join(' ');

    // Si las cabeceras no coinciden → revisión manual siempre
    if (headersMismatch) {
      logRevision.push({
        hoja: cfg.nombre, filaBkp: filaBkp, id: id || '(sin ID)',
        nombre: nomMostrar,
        motivo: 'Cabeceras de hoja actual y backup no coinciden. Verificar antes de restaurar.'
      });
      return;
    }

    // Decidir si falta en la hoja actual
    var faltaById   = id && !idsActuales[id];
    var faltaByHash = !id && !hashesActuales[hash];

    if (!faltaById && !faltaByHash) return; // ya existe, no restaurar

    // Fila vacía → ignorar
    var esVacia = row.every(function(v) { return v === '' || v === null || v === undefined; });
    if (esVacia) return;

    // Si hay ID pero también hay hash colisión dudosa → revisión
    if (id && hashesActuales[hash] && !idsActuales[id]) {
      // Mismo hash (fila idéntica) pero distinto ID → posible duplicado con ID cambiado
      logRevision.push({
        hoja: cfg.nombre, filaBkp: filaBkp, id: id,
        nombre: nomMostrar,
        motivo: 'ID no existe en hoja actual, pero existe una fila con contenido idéntico. Posible duplicado.'
      });
      return;
    }

    // Restaurar: añadir la fila al final
    try {
      // Normalizar longitud de fila al número de columnas actual
      var filaAInsertar = row.slice(0, Math.max(nColsActual, nColsBkp));
      // Rellenar con vacíos si la fila del backup tiene menos columnas que la actual
      while (filaAInsertar.length < nColsActual) filaAInsertar.push('');

      shActual.appendRow(filaAInsertar);
      var filaDestino = shActual.getLastRow();

      logRestaurado.push({
        hoja: cfg.nombre, filaBkp: filaBkp, id: id || '(sin ID)',
        nombre: nomMostrar, filaDestino: filaDestino,
        motivo: id ? 'ID ausente en hoja actual' : 'Fila ausente (sin ID, por hash)'
      });

      // Actualizar índice para evitar duplicados en la misma pasada
      if (id) idsActuales[id] = true;
      hashesActuales[hash] = true;

      logEvent('INFO', '_rst_restaurarHoja_',
        cfg.nombre + ' — restaurada fila bkp ' + filaBkp + ' → actual ' + filaDestino,
        'ID=' + id + ' | ' + nomMostrar);

    } catch (e) {
      logRevision.push({
        hoja: cfg.nombre, filaBkp: filaBkp, id: id || '(sin ID)',
        nombre: nomMostrar,
        motivo: 'Error al restaurar: ' + e.message
      });
      logEvent('ERROR', '_rst_restaurarHoja_', 'Error restaurando ' + cfg.nombre, e.message);
    }
  });
}

// ================================================================
// ÍNDICES AUXILIARES
// ================================================================

function _rst_construirIndiceIds_(sh) {
  var ids = {};
  if (!sh || sh.getLastRow() < 2) return ids;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  vals.forEach(function(r) {
    var id = String(r[0] || '').trim();
    if (id) ids[id] = true;
  });
  return ids;
}

function _rst_construirIndiceHashes_(sh) {
  var hashes = {};
  if (!sh || sh.getLastRow() < 2) return hashes;
  var nCols = sh.getLastColumn();
  var data  = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();
  data.forEach(function(row) {
    hashes[_rst_hashFila_(row)] = true;
  });
  return hashes;
}

function _rst_hashFila_(row) {
  return row.map(function(v) {
    if (v instanceof Date) return v.getTime().toString();
    return String(v === null || v === undefined ? '' : v).trim();
  }).join('|||');
}

// ================================================================
// HOJA RESTAURACION_REALIZADA
// ================================================================

function _rst_generarHojaRealizada_(ss, logRestaurado, prefijo) {
  var sh = ss.getSheetByName('RESTAURACION_REALIZADA');
  if (sh) { sh.clear(); sh.clearFormats(); }
  else    { sh = ss.insertSheet('RESTAURACION_REALIZADA'); }

  var tz    = Session.getScriptTimeZone();
  var ahora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');

  sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 80);
  sh.setColumnWidth(3, 80);  sh.setColumnWidth(4, 200);
  sh.setColumnWidth(5, 80);  sh.setColumnWidth(6, 300);

  var r = 1;
  sh.getRange(r, 1, 1, 6).merge()
    .setValue('✅ RESTAURACIÓN REALIZADA — ' + ahora)
    .setBackground('#1b5e20').setFontColor('#ffffff')
    .setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(r, 42); r++;

  sh.getRange(r, 1, 1, 6).merge()
    .setValue('Backup: ' + prefijo + '   |   Filas restauradas: ' + logRestaurado.length)
    .setBackground('#e8f5e9').setFontColor('#1b5e20').setHorizontalAlignment('center');
  r += 2;

  if (logRestaurado.length === 0) {
    sh.getRange(r, 1, 1, 6).merge()
      .setValue('No se restauró ninguna fila — las hojas ya contenían todos los registros del backup.')
      .setBackground('#f0fff4').setFontColor('#137333').setHorizontalAlignment('center');
    SpreadsheetApp.flush();
    return;
  }

  sh.getRange(r, 1, 1, 6).setValues([[
    'Hoja restaurada', 'Fila backup', 'Fila destino', 'Cliente / Nombre', 'ID restaurado', 'Motivo'
  ]]).setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
  r++;

  logRestaurado.forEach(function(item, i) {
    var bg = i % 2 === 0 ? '#f0fff4' : '#e8f5e9';
    sh.getRange(r, 1, 1, 6).setBackground(bg);
    sh.getRange(r, 1).setValue(item.hoja).setFontWeight('bold');
    sh.getRange(r, 2).setValue(item.filaBkp).setHorizontalAlignment('center');
    sh.getRange(r, 3).setValue(item.filaDestino).setHorizontalAlignment('center');
    sh.getRange(r, 4).setValue(item.nombre);
    sh.getRange(r, 5).setValue(item.id).setFontSize(9);
    sh.getRange(r, 6).setValue(item.motivo).setFontSize(9).setWrap(true);
    sh.setRowHeight(r, 26); r++;
  });

  // Resumen por hoja
  r++;
  sh.getRange(r, 1, 1, 6).merge()
    .setValue('📊 RESUMEN POR HOJA')
    .setBackground('#202124').setFontColor('#ffffff').setFontWeight('bold');
  r++;

  var porHoja = {};
  logRestaurado.forEach(function(item) {
    porHoja[item.hoja] = (porHoja[item.hoja] || 0) + 1;
  });
  Object.keys(porHoja).forEach(function(hoja, i) {
    var bg = i % 2 === 0 ? '#e8eaf6' : '#f3f4f6';
    sh.getRange(r, 1, 1, 2).setBackground(bg);
    sh.getRange(r, 1).setValue(hoja).setFontWeight('bold');
    sh.getRange(r, 2).setValue(porHoja[hoja]).setFontColor('#137333').setFontWeight('bold')
      .setHorizontalAlignment('center');
    r++;
  });

  SpreadsheetApp.flush();
}

// ================================================================
// HOJA RESTAURACION_REVISION_MANUAL
// ================================================================

function _rst_generarHojaRevision_(ss, logRevision) {
  var sh = ss.getSheetByName('RESTAURACION_REVISION_MANUAL');
  if (sh) { sh.clear(); sh.clearFormats(); }
  else    { sh = ss.insertSheet('RESTAURACION_REVISION_MANUAL'); }

  var tz    = Session.getScriptTimeZone();
  var ahora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');

  sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 80);
  sh.setColumnWidth(3, 160); sh.setColumnWidth(4, 200);
  sh.setColumnWidth(5, 380);

  var r = 1;
  sh.getRange(r, 1, 1, 5).merge()
    .setValue('⚠️ RESTAURACIÓN — REVISIÓN MANUAL REQUERIDA — ' + ahora)
    .setBackground('#e65100').setFontColor('#ffffff')
    .setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(r, 42); r++;

  sh.getRange(r, 1, 1, 5).merge()
    .setValue('Estas filas NO se restauraron automáticamente. Requieren intervención manual.')
    .setBackground('#fff3e0').setFontColor('#e65100').setHorizontalAlignment('center');
  r += 2;

  if (logRevision.length === 0) {
    sh.getRange(r, 1, 1, 5).merge()
      .setValue('✅ Sin casos para revisión manual.')
      .setBackground('#f0fff4').setFontColor('#137333').setHorizontalAlignment('center');
    SpreadsheetApp.flush();
    return;
  }

  sh.getRange(r, 1, 1, 5).setValues([[
    'Hoja', 'Fila backup', 'ID', 'Cliente / Nombre', 'Motivo / Acción recomendada'
  ]]).setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
  r++;

  logRevision.forEach(function(item, i) {
    var bg = i % 2 === 0 ? '#fff8e1' : '#fff3cd';
    sh.getRange(r, 1, 1, 5).setBackground(bg);
    sh.getRange(r, 1).setValue(item.hoja).setFontWeight('bold');
    sh.getRange(r, 2).setValue(item.filaBkp).setHorizontalAlignment('center');
    sh.getRange(r, 3).setValue(item.id).setFontSize(9);
    sh.getRange(r, 4).setValue(item.nombre);
    sh.getRange(r, 5).setValue(item.motivo).setFontSize(9).setWrap(true);
    sh.setRowHeight(r, 30); r++;
  });

  SpreadsheetApp.flush();
}

// ================================================================
// SOLO DIAGNÓSTICO (sin restaurar) — para inspección previa
// ================================================================

function ejecutarSoloDiagnosticoRestauracion() {
  var ui = SpreadsheetApp.getUi();
  var ss = getTargetSS_();

  var prefijo = _rst_encontrarPrefijoBkpMasReciente_(ss);
  if (!prefijo) {
    ui.alert('No se encontró ningún backup (hojas BKP_*).'); return;
  }

  var diagnostico = _rst_generarDiagnostico_(ss, prefijo);
  var sh = ss.getSheetByName('RESTAURACION_DIAGNOSTICO');
  if (sh) ss.setActiveSheet(sh);

  ui.alert('📋 Diagnóstico generado',
    'Backup: ' + prefijo + '\n\n'
    + _rst_resumenDiagnostico_(diagnostico) + '\n\n'
    + 'Ver hoja: RESTAURACION_DIAGNOSTICO', ui.ButtonSet.OK);
}

// ================================================================
// RESTAURAR SOLO TRABAJOS desde BKP_20260618_1601_TRA
//
// Columnas clave (base-1):
//   1  ID_Trabajo
//   2  Fecha_Creacion
//   3  ID_Cliente
//   4  Nombre_Cliente
//   7  Descripcion
//   10 Fecha_Inicio
//   11 Fecha_Fin_Estimada
//
// CLAVE DE DEDUPLICACIÓN:
//   ID_Trabajo + Nombre_Cliente + Fecha_Creacion +
//   Fecha_Inicio + Fecha_Fin_Estimada + Descripcion
//
// LÓGICA:
//   1. Unión de backup + actual (priorizando filas del backup
//      para recuperar las borradas).
//   2. Deduplicación por clave: se conserva la primera aparición.
//   3. Reescritura limpia de TRABAJOS (cabecera intacta).
//   4. Informe en TRABAJOS_RESTAURADOS.
//
// NO toca ninguna otra hoja.
// ================================================================

var RST_TRA_BACKUP   = 'BKP_20260618_1601_TRA';
var RST_TRA_HOJA     = 'TRABAJOS';
var RST_TRA_INFORME  = 'TRABAJOS_RESTAURADOS';

// Columnas que forman la clave de deduplicación (base-1)
var RST_TRA_DEDUP_COLS = [1, 4, 2, 10, 11, 7];

function restaurarSoloTrabajos() {
  var ui = SpreadsheetApp.getUi();
  var ss = getTargetSS_();

  var shBkp    = ss.getSheetByName(RST_TRA_BACKUP);
  var shActual = ss.getSheetByName(RST_TRA_HOJA);

  if (!shBkp) {
    ui.alert('❌ Backup no encontrado',
      'No existe la hoja: ' + RST_TRA_BACKUP + '\n\n'
      + 'Verifica que el nombre es correcto.', ui.ButtonSet.OK);
    return;
  }
  if (!shActual) {
    ui.alert('❌ Hoja TRABAJOS no encontrada.', '', ui.ButtonSet.OK);
    return;
  }

  // ── Leer cabecera actual ─────────────────────────────────────
  var nColsActual = Math.max(shActual.getLastColumn(), 19);
  var nColsBkp    = Math.max(shBkp.getLastColumn(), 19);
  var nCols       = Math.max(nColsActual, nColsBkp);

  var hdrActual = shActual.getRange(1, 1, 1, nColsActual).getValues()[0];
  var hdrBkp    = shBkp.getRange(1, 1, 1, nColsBkp).getValues()[0];

  // ── Leer datos ───────────────────────────────────────────────
  var dataBkp = shBkp.getLastRow() >= 2
    ? shBkp.getRange(2, 1, shBkp.getLastRow() - 1, nColsBkp).getValues()
    : [];
  var dataActual = shActual.getLastRow() >= 2
    ? shActual.getRange(2, 1, shActual.getLastRow() - 1, nColsActual).getValues()
    : [];

  // ── Normalizar filas a nCols columnas ────────────────────────
  function normalizar(row) {
    var r = row.slice(0, nCols);
    while (r.length < nCols) r.push('');
    return r;
  }

  // ── Clave de deduplicación ───────────────────────────────────
  function claveDedup(row) {
    return RST_TRA_DEDUP_COLS.map(function(c) {
      var v = row[c - 1];
      if (v instanceof Date) {
        return isNaN(v.getTime()) ? '' : v.getTime().toString();
      }
      return String(v === null || v === undefined ? '' : v).trim().toLowerCase();
    }).join('|||');
  }

  function idDe(row)      { return String(row[0] || '').trim(); }
  function clienteDe(row) { return String(row[3] || '').trim(); }

  // ── Construir unión: backup primero (fuente de verdad), luego ──
  // filas actuales que no estén en el backup por ID.
  // Esto garantiza que las filas borradas por la limpieza vuelven.

  var idsBkp = {};
  dataBkp.forEach(function(row) {
    var id = idDe(row);
    if (id) idsBkp[id] = true;
  });

  // Filas de current que el backup NO tiene (añadidas después del backup)
  var soloEnActual = dataActual.filter(function(row) {
    var id = idDe(row);
    return id && !idsBkp[id];
  });

  // Unión: backup + filas exclusivas de actual
  var union = dataBkp.concat(soloEnActual).map(normalizar);

  // ── Deduplicación ────────────────────────────────────────────
  var vistos     = {};    // clave → índice conservado
  var vistosId   = {};    // id    → primera aparición
  var deduped    = [];    // filas finales
  var logAcciones = [];   // para el informe

  // Detectar conflictos: mismo ID, distinta clave de dedup
  var idAClave = {};  // primera clave vista para cada ID
  union.forEach(function(row) {
    var id    = idDe(row);
    var clave = claveDedup(row);
    if (id && !idAClave[id]) idAClave[id] = clave;
  });

  union.forEach(function(row, idx) {
    var id     = idDe(row);
    var clave  = claveDedup(row);
    var cliente = clienteDe(row);

    // Fila completamente vacía → ignorar
    var esVacia = row.every(function(v) {
      return v === '' || v === null || v === undefined;
    });
    if (esVacia) return;

    if (!vistos[clave]) {
      // Primera vez que vemos esta clave → conservar
      vistos[clave] = true;
      deduped.push(row);

      // ¿Venía del backup o del actual?
      var eraEnBkp  = idx < dataBkp.length;
      var eraEnActual = !eraEnBkp || (id && dataActual.some(function(r) { return idDe(r) === id; }));
      var accion;
      if (eraEnBkp && !dataActual.some(function(r) { return idDe(r) === id; })) {
        accion = 'RESTAURADO';  // estaba en backup, no en actual
      } else {
        accion = 'CONSERVADO';  // estaba en ambos o solo en actual
      }

      logAcciones.push({
        id: id || '(sin ID)', cliente: cliente,
        accion: accion, eliminado: false,
        nota: accion === 'RESTAURADO' ? 'Recuperado desde ' + RST_TRA_BACKUP : ''
      });

    } else {
      // Duplicado → eliminar
      logAcciones.push({
        id: id || '(sin ID)', cliente: cliente,
        accion: 'DUPLICADO_ELIMINADO', eliminado: true,
        nota: 'Clave: ' + RST_TRA_DEDUP_COLS.map(function(c) {
          return String(row[c-1] || '').trim();
        }).join(' | ')
      });
    }
  });

  // ── Resumen para confirmación ────────────────────────────────
  var nRestaurados = logAcciones.filter(function(a) { return a.accion === 'RESTAURADO'; }).length;
  var nDuplicados  = logAcciones.filter(function(a) { return a.eliminado; }).length;
  var nConservados = logAcciones.filter(function(a) { return a.accion === 'CONSERVADO'; }).length;

  var conf = ui.alert(
    '📋 RESTAURAR TRABAJOS — CONFIRMACIÓN',
    'Backup fuente: ' + RST_TRA_BACKUP + '\n\n'
    + 'Filas en backup:          ' + dataBkp.length + '\n'
    + 'Filas en TRABAJOS actual: ' + dataActual.length + '\n'
    + 'Solo en actual (nuevas):  ' + soloEnActual.length + '\n\n'
    + '────────────────────────────\n'
    + 'A restaurar desde backup: ' + nRestaurados + '\n'
    + 'Duplicados a eliminar:    ' + nDuplicados + '\n'
    + 'Conservados sin cambio:   ' + nConservados + '\n'
    + 'TOTAL final en TRABAJOS:  ' + deduped.length + '\n'
    + '────────────────────────────\n\n'
    + '⚠️ Se reescribirá la hoja TRABAJOS.\n'
    + 'El backup ' + RST_TRA_BACKUP + ' NO se modifica.\n\n'
    + '¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (conf !== ui.Button.YES) {
    ui.alert('Cancelado. No se modificó ninguna hoja.'); return;
  }

  // ── Reescribir TRABAJOS ──────────────────────────────────────
  // 1. Limpiar filas de datos (dejar cabecera intacta en fila 1)
  if (shActual.getLastRow() > 1) {
    shActual.deleteRows(2, shActual.getLastRow() - 1);
  }

  // 2. Escribir datos deduplicados a partir de fila 2
  if (deduped.length > 0) {
    shActual.getRange(2, 1, deduped.length, nCols).setValues(deduped);
  }

  SpreadsheetApp.flush();

  // ── Generar informe TRABAJOS_RESTAURADOS ─────────────────────
  _rst_tra_generarInforme_(ss, logAcciones, dataBkp.length, dataActual.length, deduped.length);

  var shInforme = ss.getSheetByName(RST_TRA_INFORME);
  if (shInforme) ss.setActiveSheet(shInforme);

  ui.alert(
    '✅ TRABAJOS RESTAURADO',
    'Filas restauradas desde backup: ' + nRestaurados + '\n'
    + 'Duplicados eliminados:         ' + nDuplicados + '\n'
    + 'Conservados sin cambio:        ' + nConservados + '\n'
    + 'Total final en TRABAJOS:       ' + deduped.length + '\n\n'
    + 'Ver hoja: ' + RST_TRA_INFORME,
    ui.ButtonSet.OK
  );
}

// ── Informe TRABAJOS_RESTAURADOS ─────────────────────────────────

function _rst_tra_generarInforme_(ss, logAcciones, nBkp, nActual, nFinal) {
  var sh = ss.getSheetByName(RST_TRA_INFORME);
  if (sh) { sh.clear(); sh.clearFormats(); }
  else    { sh = ss.insertSheet(RST_TRA_INFORME); }

  var tz    = Session.getScriptTimeZone();
  var ahora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');

  var NC = 6;
  sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 160); sh.setColumnWidth(4, 90);
  sh.setColumnWidth(5, 90);  sh.setColumnWidth(6, 350);

  var r = 1;

  // Cabecera
  sh.getRange(r, 1, 1, NC).merge()
    .setValue('📋 TRABAJOS_RESTAURADOS — ' + ahora)
    .setBackground('#1a237e').setFontColor('#ffffff')
    .setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(r, 42); r++;

  sh.getRange(r, 1, 1, NC).merge()
    .setValue('Backup: ' + RST_TRA_BACKUP
      + '   |   Filas backup: ' + nBkp
      + '   |   Filas antes: ' + nActual
      + '   |   Filas después: ' + nFinal)
    .setBackground('#e8eaf6').setFontColor('#1a237e').setHorizontalAlignment('center');
  r += 2;

  // Resumen numérico
  var nR = logAcciones.filter(function(a) { return a.accion === 'RESTAURADO'; }).length;
  var nD = logAcciones.filter(function(a) { return a.eliminado; }).length;
  var nC = logAcciones.filter(function(a) { return a.accion === 'CONSERVADO'; }).length;

  var resumen = [
    ['🟢 Restaurados desde backup', nR, '#137333'],
    ['🔴 Duplicados eliminados',     nD, '#c5221f'],
    ['⚪ Conservados sin cambio',    nC, '#5f6368']
  ];
  resumen.forEach(function(d, i) {
    var bg = i % 2 === 0 ? '#e8eaf6' : '#f3f4f6';
    sh.getRange(r, 1, 1, NC).setBackground(bg);
    sh.getRange(r, 1).setValue(d[0]).setFontWeight('bold');
    sh.getRange(r, 2).setValue(d[1])
      .setFontColor(d[2]).setFontWeight('bold')
      .setHorizontalAlignment('center').setFontSize(12);
    r++;
  });
  r++;

  // Tabla de detalle
  sh.getRange(r, 1, 1, NC).setValues([[
    'ID Trabajo', 'Cliente', 'Acción', 'Restaurado', 'Duplicado elim.', 'Nota'
  ]]).setBackground('#202124').setFontColor('#ffffff').setFontWeight('bold');
  r++;

  // Primero los restaurados, luego eliminados, luego conservados
  var orden = logAcciones.slice().sort(function(a, b) {
    var prioridad = { RESTAURADO: 0, DUPLICADO_ELIMINADO: 1, CONSERVADO: 2 };
    return prioridad[a.accion] - prioridad[b.accion];
  });

  orden.forEach(function(item, i) {
    var bg, colorAccion;
    if (item.accion === 'RESTAURADO') {
      bg = i % 2 === 0 ? '#f0fff4' : '#e8f5e9';
      colorAccion = '#137333';
    } else if (item.accion === 'DUPLICADO_ELIMINADO') {
      bg = i % 2 === 0 ? '#fff3f3' : '#ffebee';
      colorAccion = '#c5221f';
    } else {
      bg = i % 2 === 0 ? '#f9f9f9' : '#f3f3f3';
      colorAccion = '#5f6368';
    }

    sh.getRange(r, 1, 1, NC).setBackground(bg);
    sh.getRange(r, 1).setValue(item.id).setFontSize(9);
    sh.getRange(r, 2).setValue(item.cliente);
    sh.getRange(r, 3).setValue(item.accion).setFontColor(colorAccion).setFontWeight('bold').setFontSize(9);
    sh.getRange(r, 4).setValue(item.accion === 'RESTAURADO' ? 'SÍ' : '').setHorizontalAlignment('center')
      .setFontColor('#137333').setFontWeight('bold');
    sh.getRange(r, 5).setValue(item.accion === 'DUPLICADO_ELIMINADO' ? 'SÍ' : '').setHorizontalAlignment('center')
      .setFontColor('#c5221f').setFontWeight('bold');
    sh.getRange(r, 6).setValue(item.nota).setFontSize(8).setWrap(true);
    sh.setRowHeight(r, 24); r++;
  });

  SpreadsheetApp.flush();
  logEvent('INFO', '_rst_tra_generarInforme_',
    'TRABAJOS_RESTAURADOS: restaurados=' + nR + ' | duplicados=' + nD + ' | conservados=' + nC, '');
}
