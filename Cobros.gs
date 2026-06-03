// ============================================================
// Cobros.gs — Gestión de cobros sobre la hoja FACTURAS
// ============================================================
//
// Orden de ejecución:
//   1. installCobrosModule()  — añade columnas y validaciones
//   2. validateCobrosModule() — verifica que todo está correcto
//   3. markInvoiceAsPaid()    — marcar una factura como cobrada
//
// Columnas añadidas a FACTURAS:
//   Estado Cobro      → Pendiente | Cobrado | Parcial | Cancelado
//   Fecha Cobro       → fecha en que se recibió el pago
//   Método Cobro      → Transferencia | Efectivo | Bizum | Tarjeta | Otro
//   Observaciones Cobro → texto libre
// ============================================================

var ESTADOS_COBRO  = ['Pendiente', 'Cobrado', 'Parcial', 'Cancelado'];
var METODOS_COBRO  = ['Transferencia', 'Efectivo', 'Bizum', 'Tarjeta', 'Otro'];
var COLS_COBROS    = ['Estado Cobro', 'Fecha Cobro', 'Método Cobro', 'Observaciones Cobro'];

// ─────────────────────────────────────────────────────────────
// installCobrosModule
// ─────────────────────────────────────────────────────────────

/**
 * Añade las columnas de cobro a FACTURAS si no existen,
 * aplica validaciones desplegables y pone Estado Cobro = Pendiente
 * en facturas existentes que tengan ese campo vacío.
 *
 * NO toca CLIENTES, OPERARIOS ni TRABAJOS.
 * NO modifica PDFs existentes.
 * Seguro ejecutar varias veces (idempotente).
 */
function installCobrosModule() {
  Logger.log('================================================');
  Logger.log('installCobrosModule()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var sheet = getSheet('FACTURAS');

  // ── 1. Columnas actuales ──────────────────────────────────────
  Logger.log('\n[1] Columnas actuales en FACTURAS:');
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); })
    : [];
  Logger.log('  Total columnas: ' + lastCol);

  // ── 2. Añadir columnas faltantes ──────────────────────────────
  Logger.log('\n[2] Comprobando columnas de cobro:');
  var newCols = [];
  for (var i = 0; i < COLS_COBROS.length; i++) {
    var col = COLS_COBROS[i];
    if (headers.indexOf(col) === -1) {
      lastCol++;
      var headerCell = sheet.getRange(1, lastCol);
      headerCell.setValue(col);
      headerCell.setBackground('#1a73e8')
                .setFontColor('#ffffff')
                .setFontWeight('bold')
                .setFontSize(10);
      headers.push(col);
      newCols.push({ name: col, col: lastCol });
      Logger.log('  + Añadida: "' + col + '" en columna ' + lastCol);
    } else {
      Logger.log('  ✓ Ya existe: "' + col + '"');
    }
  }
  if (newCols.length === 0) Logger.log('  ✓ Todas las columnas ya estaban presentes.');

  // ── 3. Localizar columnas de cobro ────────────────────────────
  var cols = _findCobrosColumns_(sheet);
  Logger.log('\n[3] Posiciones de columnas de cobro:');
  Logger.log('  Estado Cobro:       col ' + (cols.estadoCobro  || '❌'));
  Logger.log('  Fecha Cobro:        col ' + (cols.fechaCobro   || '❌'));
  Logger.log('  Método Cobro:       col ' + (cols.metodoCobro  || '❌'));
  Logger.log('  Observaciones Cobro: col ' + (cols.observaciones || '❌'));

  // ── 4. Validaciones desplegables ──────────────────────────────
  Logger.log('\n[4] Aplicando validaciones desplegables:');
  var maxRows = sheet.getMaxRows();
  var dataRows = Math.max(maxRows - 1, 1);

  if (cols.estadoCobro) {
    var ruleEstado = SpreadsheetApp.newDataValidation()
      .requireValueInList(ESTADOS_COBRO, true)
      .setAllowInvalid(false)
      .setHelpText('Selecciona el estado del cobro.')
      .build();
    sheet.getRange(2, cols.estadoCobro, dataRows, 1).setDataValidation(ruleEstado);
    Logger.log('  ✓ "Estado Cobro" → ' + ESTADOS_COBRO.join(' | '));
  }

  if (cols.metodoCobro) {
    var ruleMetodo = SpreadsheetApp.newDataValidation()
      .requireValueInList(METODOS_COBRO, true)
      .setAllowInvalid(false)
      .setHelpText('Selecciona el método de cobro.')
      .build();
    sheet.getRange(2, cols.metodoCobro, dataRows, 1).setDataValidation(ruleMetodo);
    Logger.log('  ✓ "Método Cobro"  → ' + METODOS_COBRO.join(' | '));
  }

  // ── 5. Estado Cobro = Pendiente en filas existentes vacías ─────
  Logger.log('\n[5] Inicializando Estado Cobro en facturas existentes...');
  var lastRow = sheet.getLastRow();
  var actualizadas = 0;

  if (lastRow > 1 && cols.estadoCobro) {
    var estadoRange  = sheet.getRange(2, cols.estadoCobro, lastRow - 1, 1);
    var estadoValues = estadoRange.getValues();
    for (var r = 0; r < estadoValues.length; r++) {
      if (!estadoValues[r][0] || String(estadoValues[r][0]).trim() === '') {
        estadoValues[r][0] = 'Pendiente';
        actualizadas++;
      }
    }
    if (actualizadas > 0) {
      estadoRange.setValues(estadoValues);
      Logger.log('  ✓ ' + actualizadas + ' fila(s) puestas a "Pendiente".');
    } else {
      Logger.log('  ✓ Todas las facturas ya tenían Estado Cobro asignado.');
    }
  } else {
    Logger.log('  ✓ Sin facturas existentes — nada que inicializar.');
  }

  // ── 6. Inicializar claves CONFIG para markInvoiceAsPaid ───────
  Logger.log('\n[6] Inicializando claves CONFIG para markInvoiceAsPaid:');
  var clavesCobro = [
    'factura_a_marcar',
    'fecha_cobro',
    'metodo_cobro',
    'observaciones_cobro'
  ];
  for (var c = 0; c < clavesCobro.length; c++) {
    if (getConfig(clavesCobro[c]) === null) {
      setConfig(clavesCobro[c], '');
      Logger.log('  + Creada: ' + clavesCobro[c]);
    } else {
      Logger.log('  ✓ Ya existe: ' + clavesCobro[c]);
    }
  }

  // ── Resumen ───────────────────────────────────────────────────
  Logger.log('\n================================================');
  Logger.log('MÓDULO COBROS INSTALADO.');
  Logger.log('  Columnas añadidas:   ' + newCols.length);
  Logger.log('  Filas inicializadas: ' + actualizadas);
  Logger.log('  Siguiente paso: ejecuta validateCobrosModule()');
  Logger.log('================================================');
}

// ─────────────────────────────────────────────────────────────
// validateCobrosModule
// ─────────────────────────────────────────────────────────────

/**
 * Verifica que el módulo de cobros está correctamente instalado.
 * Comprueba columnas, validaciones, datos existentes y claves CONFIG.
 * SOLO LECTURA — no modifica nada.
 */
function validateCobrosModule() {
  Logger.log('================================================');
  Logger.log('validateCobrosModule()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var ok = true;

  // 1. Hoja FACTURAS
  Logger.log('\n[1] Hoja FACTURAS:');
  var sheet;
  try {
    sheet = getSheet('FACTURAS');
    var totalRegistros = Math.max(0, sheet.getLastRow() - 1);
    Logger.log('  ✓ GID=' + sheet.getSheetId() + '  registros=' + totalRegistros);
  } catch (e) {
    Logger.log('  ❌ NO EXISTE: ' + e.message);
    return;
  }

  // 2. Columnas de cobro
  Logger.log('\n[2] Columnas de cobro:');
  var cols = _findCobrosColumns_(sheet);
  var colsOk = true;
  for (var i = 0; i < COLS_COBROS.length; i++) {
    var nombre = COLS_COBROS[i];
    var clave  = _colNameToKey_(nombre);
    var colNum = cols[clave];
    if (colNum) {
      Logger.log('  ✓ "' + nombre + '" — columna ' + colNum);
    } else {
      Logger.log('  ❌ "' + nombre + '" — NO EXISTE. Ejecuta installCobrosModule().');
      colsOk = false;
      ok     = false;
    }
  }

  // 3. Validaciones desplegables
  Logger.log('\n[3] Validaciones desplegables:');
  if (cols.estadoCobro) {
    var valEstado = sheet.getRange(2, cols.estadoCobro).getDataValidation();
    if (valEstado) {
      Logger.log('  ✓ "Estado Cobro" tiene validación.');
      Logger.log('    Valores: ' + ESTADOS_COBRO.join(' | '));
    } else {
      Logger.log('  ⚠️  "Estado Cobro" sin validación — ejecuta installCobrosModule().');
    }
  }
  if (cols.metodoCobro) {
    var valMetodo = sheet.getRange(2, cols.metodoCobro).getDataValidation();
    if (valMetodo) {
      Logger.log('  ✓ "Método Cobro" tiene validación.');
      Logger.log('    Valores: ' + METODOS_COBRO.join(' | '));
    } else {
      Logger.log('  ⚠️  "Método Cobro" sin validación — ejecuta installCobrosModule().');
    }
  }

  // 4. Facturas existentes conservadas
  Logger.log('\n[4] Integridad de datos existentes:');
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log('  ✓ Sin facturas — hoja limpia.');
  } else {
    var sample = sheet.getRange(2, 1, Math.min(lastRow - 1, 5), 1).getValues();
    Logger.log('  ✓ ' + (lastRow - 1) + ' factura(s) encontrada(s). Primeras IDs:');
    for (var r = 0; r < sample.length; r++) {
      Logger.log('    · ' + sample[r][0]);
    }

    // Verificar que Estado Cobro está relleno en facturas existentes
    if (cols.estadoCobro) {
      var estadoVals = sheet.getRange(2, cols.estadoCobro, lastRow - 1, 1).getValues();
      var sinEstado  = estadoVals.filter(function(v) { return !String(v[0]).trim(); }).length;
      if (sinEstado > 0) {
        Logger.log('  ⚠️  ' + sinEstado + ' factura(s) sin Estado Cobro — ejecuta installCobrosModule().');
      } else {
        Logger.log('  ✓ Todas las facturas tienen Estado Cobro asignado.');
      }
    }
  }

  // 5. CONFIG — claves de markInvoiceAsPaid
  Logger.log('\n[5] CONFIG — claves para markInvoiceAsPaid:');
  var clavesCobro = [
    ['factura_a_marcar',    'ID de la factura (ej: FAC-2026-001)'],
    ['fecha_cobro',         'DD/MM/YYYY'],
    ['metodo_cobro',        METODOS_COBRO.join(' | ')],
    ['observaciones_cobro', 'Texto libre (opcional)']
  ];
  for (var c = 0; c < clavesCobro.length; c++) {
    var clave = clavesCobro[c][0];
    var desc  = clavesCobro[c][1];
    var val   = getConfig(clave);
    Logger.log('  ' + clave + ': ' + (val ? '"' + val + '"' : '(vacío) — ' + desc));
  }

  // Resumen
  Logger.log('\n================================================');
  Logger.log('RESULTADO:');
  Logger.log('  Columnas: ' + (colsOk ? '✓' : '❌'));
  Logger.log('  ' + (ok ? 'MÓDULO COBROS LISTO.' : 'HAY PROBLEMAS — revisa los ❌ anteriores.'));
  if (ok) Logger.log('  Rellena CONFIG ([5]) y ejecuta markInvoiceAsPaid().');
  Logger.log('================================================');
}

// ─────────────────────────────────────────────────────────────
// markInvoiceAsPaid
// ─────────────────────────────────────────────────────────────

/**
 * Marca una factura como cobrada actualizando las columnas de cobro en FACTURAS.
 *
 * Lee de CONFIG:
 *   factura_a_marcar    → ID de la factura (ej: FAC-2026-001)
 *   fecha_cobro         → DD/MM/YYYY
 *   metodo_cobro        → Transferencia | Efectivo | Bizum | Tarjeta | Otro
 *   observaciones_cobro → texto libre (opcional)
 */
function markInvoiceAsPaid() {
  Logger.log('================================================');
  Logger.log('markInvoiceAsPaid()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  // ── 1. Leer CONFIG ────────────────────────────────────────────
  var facturaId     = getConfig('factura_a_marcar');
  var fechaCobroStr = getConfig('fecha_cobro');
  var metodoCobro   = getConfig('metodo_cobro');
  var observaciones = getConfig('observaciones_cobro') || '';

  Logger.log('\n[1] Parámetros de CONFIG:');
  Logger.log('  factura_a_marcar:    ' + (facturaId     || '❌ NO CONFIGURADO'));
  Logger.log('  fecha_cobro:         ' + (fechaCobroStr || '❌ NO CONFIGURADO'));
  Logger.log('  metodo_cobro:        ' + (metodoCobro   || '❌ NO CONFIGURADO'));
  Logger.log('  observaciones_cobro: ' + (observaciones || '(vacío)'));

  if (!facturaId || !fechaCobroStr || !metodoCobro) {
    Logger.log('\nABORTANDO. Rellena factura_a_marcar, fecha_cobro y metodo_cobro en CONFIG.');
    return;
  }

  // Validar método de cobro
  if (METODOS_COBRO.indexOf(metodoCobro) === -1) {
    Logger.log('ERROR: metodo_cobro "' + metodoCobro + '" no válido.');
    Logger.log('  Valores aceptados: ' + METODOS_COBRO.join(' | '));
    return;
  }

  // Parsear fecha de cobro
  var fechaCobro = _parseReportDate_(fechaCobroStr);
  if (!fechaCobro) {
    Logger.log('ERROR: fecha_cobro "' + fechaCobroStr + '" no válida. Usa DD/MM/YYYY.');
    return;
  }

  // ── 2. Abrir FACTURAS y localizar columnas ────────────────────
  Logger.log('\n[2] Localizando columnas de cobro...');
  var sheet = getSheet('FACTURAS');
  var cols  = _findCobrosColumns_(sheet);

  if (!cols.estadoCobro || !cols.fechaCobro || !cols.metodoCobro) {
    Logger.log('  ❌ Columnas de cobro no encontradas. Ejecuta installCobrosModule() primero.');
    return;
  }
  Logger.log('  ✓ Estado Cobro: col ' + cols.estadoCobro);
  Logger.log('  ✓ Fecha Cobro:  col ' + cols.fechaCobro);
  Logger.log('  ✓ Método Cobro: col ' + cols.metodoCobro);

  // ── 3. Buscar la factura por ID ───────────────────────────────
  Logger.log('\n[3] Buscando factura "' + facturaId + '"...');
  var lastRow = sheet.getLastRow();
  var filaFactura = null;

  if (lastRow > 1) {
    var ids = sheet.getRange(2, cols.id, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === facturaId.trim()) {
        filaFactura = i + 2;
        break;
      }
    }
  }

  if (!filaFactura) {
    Logger.log('  ❌ Factura "' + facturaId + '" no encontrada en FACTURAS.');
    return;
  }
  Logger.log('  ✓ Encontrada en fila ' + filaFactura);

  // Estado actual antes del cambio
  var estadoActual = sheet.getRange(filaFactura, cols.estadoCobro).getValue();
  Logger.log('  Estado Cobro actual: "' + (estadoActual || '(vacío)') + '"');

  // ── 4. Actualizar columnas de cobro ───────────────────────────
  Logger.log('\n[4] Actualizando factura...');
  sheet.getRange(filaFactura, cols.estadoCobro).setValue('Cobrado');
  sheet.getRange(filaFactura, cols.fechaCobro).setValue(fechaCobro);
  sheet.getRange(filaFactura, cols.metodoCobro).setValue(metodoCobro);
  if (cols.observaciones) {
    sheet.getRange(filaFactura, cols.observaciones).setValue(observaciones);
  }

  Logger.log('  ✓ Estado Cobro:       "Cobrado"');
  Logger.log('  ✓ Fecha Cobro:        ' + fechaCobroStr);
  Logger.log('  ✓ Método Cobro:       ' + metodoCobro);
  Logger.log('  ✓ Observaciones:      ' + (observaciones || '(sin observaciones)'));

  logEvent('INFO', 'markInvoiceAsPaid',
    facturaId + ' → Cobrado | ' + fechaCobroStr + ' | ' + metodoCobro,
    observaciones);

  // ── Resumen ───────────────────────────────────────────────────
  Logger.log('\n================================================');
  Logger.log('FACTURA MARCADA COMO COBRADA');
  Logger.log('  ID:           ' + facturaId);
  Logger.log('  Fecha cobro:  ' + fechaCobroStr);
  Logger.log('  Método:       ' + metodoCobro);
  Logger.log('  Observaciones:' + (observaciones || '—'));
  Logger.log('================================================');
}

// ── Funciones privadas ────────────────────────────────────────

/**
 * Devuelve un objeto con las posiciones (base 1) de cada columna de cobro.
 * Busca por nombre exacto de cabecera. Devuelve null en los campos no encontrados.
 */
function _findCobrosColumns_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return { id: null, estadoCobro: null, fechaCobro: null, metodoCobro: null, observaciones: null };

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var result  = { id: null, estadoCobro: null, fechaCobro: null, metodoCobro: null, observaciones: null };

  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim();
    if (h === 'ID_Factura')          result.id           = i + 1;
    if (h === 'Estado Cobro')        result.estadoCobro  = i + 1;
    if (h === 'Fecha Cobro')         result.fechaCobro   = i + 1;
    if (h === 'Método Cobro')        result.metodoCobro  = i + 1;
    if (h === 'Observaciones Cobro') result.observaciones = i + 1;
  }
  return result;
}

/**
 * Convierte el nombre de columna al nombre de clave en el objeto de _findCobrosColumns_.
 */
function _colNameToKey_(nombre) {
  var map = {
    'Estado Cobro':        'estadoCobro',
    'Fecha Cobro':         'fechaCobro',
    'Método Cobro':        'metodoCobro',
    'Observaciones Cobro': 'observaciones'
  };
  return map[nombre] || null;
}
