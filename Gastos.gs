// ============================================================
// Gastos.gs — Registro de gastos e IVA soportado (IVA In)
// ============================================================
//
// Orden de ejecución:
//   1. installGastosModule()   — crea hoja GASTOS y claves CONFIG
//   2. validateGastosModule()  — verifica instalación
//   3. registerExpense()       — registrar un gasto
//   4. calculateVATSummary()   — resumen IVA Out / IVA In / IVA Neto
//
// Claves CONFIG para registerExpense (rellenar en hoja CONFIG):
//   proveedor_gasto     → nombre del proveedor
//   concepto_gasto      → descripción del gasto
//   base_gasto          → base imponible en EUR (ej: 100)
//   iva_gasto           → importe de IVA en EUR (ej: 21)
//   estado_pago_gasto   → Pendiente | Pagado
// ============================================================

// Columnas de la hoja GASTOS (base 1)
var COL_GAS = {
  ID:            1,
  FECHA:         2,
  PROVEEDOR:     3,
  CONCEPTO:      4,
  BASE:          5,
  IVA:           6,
  TOTAL:         7,
  FACTURA_URL:   8,
  ESTADO_PAGO:   9,
  OBSERVACIONES: 10
};

var HEADERS_GASTOS = [
  'ID Gasto', 'Fecha', 'Proveedor', 'Concepto',
  'Base Imponible', 'IVA', 'Total',
  'Factura PDF/URL', 'Estado Pago', 'Observaciones'
];

var ESTADOS_PAGO_GASTO = ['Pendiente', 'Pagado'];

// ─────────────────────────────────────────────────────────────
// installGastosModule
// ─────────────────────────────────────────────────────────────

/**
 * Crea la hoja GASTOS con sus cabeceras y validaciones.
 * Crea las claves de CONFIG necesarias para registerExpense.
 * Seguro ejecutar varias veces (idempotente).
 */
function installGastosModule() {
  Logger.log('================================================');
  Logger.log('installGastosModule()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var ss = getTargetSS_();

  // ── 1. Hoja GASTOS ────────────────────────────────────────────
  Logger.log('\n[1] Hoja GASTOS:');
  var sheet = ss.getSheetByName('GASTOS');
  if (!sheet) {
    sheet = ss.insertSheet('GASTOS');
    Logger.log('  + Hoja creada.');
  } else {
    Logger.log('  ✓ Ya existe.');
  }

  // Escribir cabeceras si la hoja está vacía
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS_GASTOS);
    applyHeaderStyle_(sheet, HEADERS_GASTOS.length);
    Logger.log('  ✓ Cabeceras escritas: ' + HEADERS_GASTOS.join(' | '));
  } else {
    Logger.log('  ✓ Cabeceras existentes conservadas.');
    Logger.log('    ' + sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].join(' | '));
  }

  // ── 2. Ancho de columnas ──────────────────────────────────────
  Logger.log('\n[2] Ajustando anchos de columna:');
  sheet.setColumnWidth(COL_GAS.ID,            110);
  sheet.setColumnWidth(COL_GAS.FECHA,          90);
  sheet.setColumnWidth(COL_GAS.PROVEEDOR,     160);
  sheet.setColumnWidth(COL_GAS.CONCEPTO,      220);
  sheet.setColumnWidth(COL_GAS.BASE,           90);
  sheet.setColumnWidth(COL_GAS.IVA,            70);
  sheet.setColumnWidth(COL_GAS.TOTAL,          90);
  sheet.setColumnWidth(COL_GAS.FACTURA_URL,   200);
  sheet.setColumnWidth(COL_GAS.ESTADO_PAGO,   110);
  sheet.setColumnWidth(COL_GAS.OBSERVACIONES, 200);
  Logger.log('  ✓ Anchos ajustados.');

  // ── 3. Validación Estado Pago ─────────────────────────────────
  Logger.log('\n[3] Validación Estado Pago:');
  var maxRows   = sheet.getMaxRows();
  var dataRows  = Math.max(maxRows - 1, 1);
  var rulePago  = SpreadsheetApp.newDataValidation()
    .requireValueInList(ESTADOS_PAGO_GASTO, true)
    .setAllowInvalid(false)
    .setHelpText('Selecciona el estado del pago.')
    .build();
  sheet.getRange(2, COL_GAS.ESTADO_PAGO, dataRows, 1).setDataValidation(rulePago);
  Logger.log('  ✓ "Estado Pago" → ' + ESTADOS_PAGO_GASTO.join(' | '));

  // ── 4. Claves CONFIG para registerExpense ─────────────────────
  Logger.log('\n[4] Inicializando claves CONFIG:');
  var clavesGasto = [
    ['proveedor_gasto',   'Nombre del proveedor'],
    ['concepto_gasto',    'Descripción del gasto'],
    ['base_gasto',        'Base imponible en EUR (ej: 100)'],
    ['iva_gasto',         'Importe de IVA en EUR (ej: 21)'],
    ['estado_pago_gasto', 'Pendiente | Pagado']
  ];
  for (var i = 0; i < clavesGasto.length; i++) {
    var clave = clavesGasto[i][0];
    if (getConfig(clave) === null) {
      setConfig(clave, '');
      Logger.log('  + Creada: ' + clave + ' — ' + clavesGasto[i][1]);
    } else {
      Logger.log('  ✓ Ya existe: ' + clave);
    }
  }

  // ── Resumen ───────────────────────────────────────────────────
  Logger.log('\n================================================');
  Logger.log('MÓDULO GASTOS INSTALADO.');
  Logger.log('  Siguiente paso: ejecuta validateGastosModule()');
  Logger.log('================================================');
}

// ─────────────────────────────────────────────────────────────
// validateGastosModule
// ─────────────────────────────────────────────────────────────

/**
 * Verifica que el módulo de gastos está correctamente instalado.
 * SOLO LECTURA — no modifica nada.
 */
function validateGastosModule() {
  Logger.log('================================================');
  Logger.log('validateGastosModule()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var ok = true;

  // 1. Hoja GASTOS
  Logger.log('\n[1] Hoja GASTOS:');
  var sheet;
  try {
    sheet = getSheet('GASTOS');
    var numGastos = Math.max(0, sheet.getLastRow() - 1);
    Logger.log('  ✓ GID=' + sheet.getSheetId() + '  registros=' + numGastos);
  } catch (e) {
    Logger.log('  ❌ NO EXISTE: ' + e.message);
    Logger.log('     Ejecuta installGastosModule() primero.');
    return;
  }

  // 2. Cabeceras
  Logger.log('\n[2] Cabeceras:');
  var lastCol  = sheet.getLastColumn();
  var hdrs     = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var hdrsOk   = true;
  for (var i = 0; i < HEADERS_GASTOS.length; i++) {
    var esperada  = HEADERS_GASTOS[i];
    var presente  = hdrs.indexOf(esperada) !== -1;
    Logger.log('  ' + (presente ? '✓' : '❌') + ' "' + esperada + '"');
    if (!presente) { hdrsOk = false; ok = false; }
  }

  // 3. Validación Estado Pago
  Logger.log('\n[3] Validación "Estado Pago":');
  var valPago = sheet.getRange(2, COL_GAS.ESTADO_PAGO).getDataValidation();
  if (valPago) {
    Logger.log('  ✓ Validación presente → ' + ESTADOS_PAGO_GASTO.join(' | '));
  } else {
    Logger.log('  ⚠️  Sin validación — ejecuta installGastosModule().');
  }

  // 4. Hoja FACTURAS (necesaria para calculateVATSummary)
  Logger.log('\n[4] Hoja FACTURAS (para IVA Out):');
  try {
    var shFac   = getSheet('FACTURAS');
    var numFac  = Math.max(0, shFac.getLastRow() - 1);
    Logger.log('  ✓ GID=' + shFac.getSheetId() + '  registros=' + numFac);
  } catch (e) {
    Logger.log('  ❌ NO EXISTE: ' + e.message);
    ok = false;
  }

  // 5. CONFIG — claves de gasto
  Logger.log('\n[5] CONFIG — claves de registerExpense:');
  var clavesGasto = [
    ['proveedor_gasto',   'Nombre del proveedor'],
    ['concepto_gasto',    'Descripción del gasto'],
    ['base_gasto',        'Base imponible en EUR'],
    ['iva_gasto',         'Importe de IVA en EUR'],
    ['estado_pago_gasto', 'Pendiente | Pagado']
  ];
  for (var c = 0; c < clavesGasto.length; c++) {
    var clave = clavesGasto[c][0];
    var desc  = clavesGasto[c][1];
    var val   = getConfig(clave);
    Logger.log('  ' + clave + ': ' + (val ? '"' + val + '"' : '(vacío) — ' + desc));
  }

  // 6. Siguiente ID de gasto
  Logger.log('\n[6] Siguiente ID de gasto:');
  try {
    var sig = _nextExpenseId_();
    Logger.log('  ' + sig);
  } catch (e) {
    Logger.log('  Error: ' + e.message);
  }

  // Resumen
  Logger.log('\n================================================');
  Logger.log('RESULTADO:');
  Logger.log('  GASTOS:    ' + (ok ? '✓' : '❌'));
  Logger.log('  Cabeceras: ' + (hdrsOk ? '✓' : '❌'));
  if (ok) {
    Logger.log('  MÓDULO LISTO.');
    Logger.log('  Rellena CONFIG ([5]) y ejecuta registerExpense().');
    Logger.log('  Para el resumen de IVA: calculateVATSummary().');
  } else {
    Logger.log('  HAY PROBLEMAS — revisa los ❌ anteriores.');
  }
  Logger.log('================================================');
}

// ─────────────────────────────────────────────────────────────
// registerExpense
// ─────────────────────────────────────────────────────────────

/**
 * Registra un gasto en la hoja GASTOS leyendo los parámetros de CONFIG.
 *
 * Claves requeridas:
 *   proveedor_gasto, concepto_gasto, base_gasto, iva_gasto, estado_pago_gasto
 */
function registerExpense() {
  Logger.log('================================================');
  Logger.log('registerExpense()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  // ── 1. Leer CONFIG ────────────────────────────────────────────
  var proveedor   = getConfig('proveedor_gasto');
  var concepto    = getConfig('concepto_gasto');
  var baseStr     = getConfig('base_gasto');
  var ivaStr      = getConfig('iva_gasto');
  var estadoPago  = getConfig('estado_pago_gasto') || 'Pendiente';

  Logger.log('\n[1] Parámetros de CONFIG:');
  Logger.log('  proveedor_gasto:   ' + (proveedor  || '❌ NO CONFIGURADO'));
  Logger.log('  concepto_gasto:    ' + (concepto   || '❌ NO CONFIGURADO'));
  Logger.log('  base_gasto:        ' + (baseStr    || '❌ NO CONFIGURADO'));
  Logger.log('  iva_gasto:         ' + (ivaStr     || '❌ NO CONFIGURADO'));
  Logger.log('  estado_pago_gasto: ' + estadoPago);

  if (!proveedor || !concepto || !baseStr || !ivaStr) {
    Logger.log('\nABORTANDO. Rellena las claves en CONFIG y vuelve a ejecutar.');
    return;
  }

  // Validar estado de pago
  if (ESTADOS_PAGO_GASTO.indexOf(estadoPago) === -1) {
    Logger.log('ERROR: estado_pago_gasto "' + estadoPago + '" no válido.');
    Logger.log('  Valores aceptados: ' + ESTADOS_PAGO_GASTO.join(' | '));
    return;
  }

  var base  = parseFloat(String(baseStr).replace(',', '.')) || 0;
  var iva   = parseFloat(String(ivaStr).replace(',', '.'))  || 0;
  var total = base + iva;

  Logger.log('\n[2] Importes calculados:');
  Logger.log('  Base:  ' + formatCurrency(base));
  Logger.log('  IVA:   ' + formatCurrency(iva));
  Logger.log('  Total: ' + formatCurrency(total));

  // ── 2. Generar ID ─────────────────────────────────────────────
  Logger.log('\n[3] Generando ID de gasto...');
  var id  = _nextExpenseId_();
  var hoy = new Date();
  Logger.log('  ID: ' + id);

  // ── 3. Registrar en GASTOS ────────────────────────────────────
  Logger.log('\n[4] Registrando en GASTOS...');
  var sheet = getSheet('GASTOS');
  sheet.appendRow([
    id,           // ID Gasto
    hoy,          // Fecha
    proveedor,    // Proveedor
    concepto,     // Concepto
    base,         // Base Imponible
    iva,          // IVA
    total,        // Total
    '',           // Factura PDF/URL (rellenar manualmente si se tiene)
    estadoPago,   // Estado Pago
    ''            // Observaciones (rellenar manualmente)
  ]);
  Logger.log('  ✓ Registrado.');

  logEvent('INFO', 'registerExpense',
    id + ' | ' + proveedor + ' | ' + concepto + ' | Total: ' + total.toFixed(2) + ' €',
    estadoPago);

  // ── Resumen ───────────────────────────────────────────────────
  Logger.log('\n================================================');
  Logger.log('GASTO REGISTRADO');
  Logger.log('  ID:        ' + id);
  Logger.log('  Fecha:     ' + Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'dd/MM/yyyy'));
  Logger.log('  Proveedor: ' + proveedor);
  Logger.log('  Concepto:  ' + concepto);
  Logger.log('  Base:      ' + formatCurrency(base));
  Logger.log('  IVA:       ' + formatCurrency(iva));
  Logger.log('  Total:     ' + formatCurrency(total));
  Logger.log('  Estado:    ' + estadoPago);
  Logger.log('================================================');
}

// ─────────────────────────────────────────────────────────────
// calculateVATSummary
// ─────────────────────────────────────────────────────────────

/**
 * Muestra el resumen de IVA repercutido (Out), soportado (In) y neto.
 * Desglosa por año y muestra totales globales.
 * SOLO LECTURA — no modifica datos.
 */
function calculateVATSummary() {
  Logger.log('================================================');
  Logger.log('calculateVATSummary()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  // ── 1. IVA OUT — leer de FACTURAS (IVA_Importe) ──────────────
  Logger.log('\n[1] IVA OUT — Facturas emitidas (IVA_Importe):');
  var ivaOutPorAnio = {};
  var ivaOutTotal   = 0;

  try {
    var shFac   = getSheet('FACTURAS');
    var lastFac = shFac.getLastRow();

    if (lastFac > 1) {
      // Localizar columnas por nombre
      var facHdrs = shFac.getRange(1, 1, 1, shFac.getLastColumn()).getValues()[0];
      var colFechaFac = -1, colIvaImporte = -1;
      for (var h = 0; h < facHdrs.length; h++) {
        var hn = String(facHdrs[h]).trim();
        if (hn === 'Fecha_Emision') colFechaFac   = h + 1;
        if (hn === 'IVA_Importe')  colIvaImporte = h + 1;
      }

      if (colIvaImporte === -1) {
        Logger.log('  ⚠️  Columna IVA_Importe no encontrada en FACTURAS.');
      } else {
        var facData = shFac.getRange(2, 1, lastFac - 1, shFac.getLastColumn()).getValues();
        for (var f = 0; f < facData.length; f++) {
          var importe = parseFloat(facData[f][colIvaImporte - 1]) || 0;
          var anio    = '—';
          if (colFechaFac !== -1) {
            var fechaVal = facData[f][colFechaFac - 1];
            if (fechaVal instanceof Date && !isNaN(fechaVal.getTime())) {
              anio = String(fechaVal.getFullYear());
            } else if (fechaVal) {
              var y = String(fechaVal).match(/\d{4}/);
              if (y) anio = y[0];
            }
          }
          ivaOutPorAnio[anio] = (ivaOutPorAnio[anio] || 0) + importe;
          ivaOutTotal += importe;
        }
        var aniosOut = Object.keys(ivaOutPorAnio).sort();
        for (var ao = 0; ao < aniosOut.length; ao++) {
          Logger.log('  ' + aniosOut[ao] + ': ' + formatCurrency(ivaOutPorAnio[aniosOut[ao]]));
        }
        Logger.log('  TOTAL IVA OUT: ' + formatCurrency(ivaOutTotal));
      }
    } else {
      Logger.log('  ⚠️  Sin facturas registradas.');
    }
  } catch (e) {
    Logger.log('  ❌ Error leyendo FACTURAS: ' + e.message);
  }

  // ── 2. IVA IN — leer de GASTOS (columna IVA) ─────────────────
  Logger.log('\n[2] IVA IN — Gastos (IVA soportado):');
  var ivaInPorAnio = {};
  var ivaInTotal   = 0;

  try {
    var shGas   = getSheet('GASTOS');
    var lastGas = shGas.getLastRow();

    if (lastGas > 1) {
      var gasData = shGas.getRange(2, 1, lastGas - 1, HEADERS_GASTOS.length).getValues();
      for (var g = 0; g < gasData.length; g++) {
        var ivaGasto = parseFloat(gasData[g][COL_GAS.IVA - 1]) || 0;
        var fechaGas = gasData[g][COL_GAS.FECHA - 1];
        var anioG    = '—';
        if (fechaGas instanceof Date && !isNaN(fechaGas.getTime())) {
          anioG = String(fechaGas.getFullYear());
        } else if (fechaGas) {
          var yg = String(fechaGas).match(/\d{4}/);
          if (yg) anioG = yg[0];
        }
        ivaInPorAnio[anioG] = (ivaInPorAnio[anioG] || 0) + ivaGasto;
        ivaInTotal += ivaGasto;
      }
      var aniosIn = Object.keys(ivaInPorAnio).sort();
      for (var ai = 0; ai < aniosIn.length; ai++) {
        Logger.log('  ' + aniosIn[ai] + ': ' + formatCurrency(ivaInPorAnio[aniosIn[ai]]));
      }
      Logger.log('  TOTAL IVA IN:  ' + formatCurrency(ivaInTotal));
    } else {
      Logger.log('  ⚠️  Sin gastos registrados.');
    }
  } catch (e) {
    Logger.log('  ❌ Error leyendo GASTOS: ' + e.message);
  }

  // ── 3. IVA Neto ───────────────────────────────────────────────
  Logger.log('\n[3] IVA NETO (Out − In):');
  var anioActual = String(new Date().getFullYear());
  var ivaOutAnio = ivaOutPorAnio[anioActual] || 0;
  var ivaInAnio  = ivaInPorAnio[anioActual]  || 0;
  var ivaNetoAnio  = ivaOutAnio  - ivaInAnio;
  var ivaNetoTotal = ivaOutTotal - ivaInTotal;

  Logger.log('  ' + anioActual + ':');
  Logger.log('    IVA Out:  ' + formatCurrency(ivaOutAnio));
  Logger.log('    IVA In:   ' + formatCurrency(ivaInAnio));
  Logger.log('    IVA Neto: ' + formatCurrency(ivaNetoAnio) +
             (ivaNetoAnio >= 0 ? ' (a ingresar)' : ' (a compensar/devolver)'));

  Logger.log('\n================================================');
  Logger.log('RESUMEN GLOBAL');
  Logger.log('  IVA OUT (repercutido): ' + formatCurrency(ivaOutTotal));
  Logger.log('  IVA IN  (soportado):   ' + formatCurrency(ivaInTotal));
  Logger.log('  IVA NETO:              ' + formatCurrency(ivaNetoTotal) +
             (ivaNetoTotal >= 0 ? ' → a ingresar a Hacienda' : ' → a compensar o devolver'));
  Logger.log('================================================');
}

// ─────────────────────────────────────────────────────────────
// setGastosConfig  — alternativa a editar CONFIG manualmente
// ─────────────────────────────────────────────────────────────

/**
 * Escribe los valores de un gasto directamente en CONFIG.
 * Usar cuando no sea posible editar la hoja CONFIG a mano.
 *
 * Edita los valores dentro del código antes de ejecutar:
 *   proveedor, concepto, base, iva, estadoPago
 * Después ejecuta registerExpense().
 */
function setGastosConfig() {
  // ── EDITAR ESTOS VALORES ANTES DE EJECUTAR ────────────────────
  var proveedor  = 'Ferretería López';   // nombre del proveedor
  var concepto   = 'Material de obra';   // descripción del gasto
  var base       = '100';                // base imponible en EUR
  var iva        = '21';                 // importe de IVA en EUR
  var estadoPago = 'Pendiente';          // Pendiente | Pagado
  // ─────────────────────────────────────────────────────────────

  Logger.log('================================================');
  Logger.log('setGastosConfig()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  setConfig('proveedor_gasto',   proveedor);
  setConfig('concepto_gasto',    concepto);
  setConfig('base_gasto',        base);
  setConfig('iva_gasto',         iva);
  setConfig('estado_pago_gasto', estadoPago);

  Logger.log('\n  ✓ proveedor_gasto:   "' + proveedor  + '"');
  Logger.log('  ✓ concepto_gasto:    "' + concepto   + '"');
  Logger.log('  ✓ base_gasto:        "' + base       + '"');
  Logger.log('  ✓ iva_gasto:         "' + iva        + '"');
  Logger.log('  ✓ estado_pago_gasto: "' + estadoPago + '"');
  Logger.log('\n  Siguiente paso: ejecuta registerExpense()');
  Logger.log('================================================');
}

// ── Funciones privadas ────────────────────────────────────────

/**
 * Genera el siguiente ID de gasto del año en curso.
 * Formato: GAS-YYYY-NNN
 */
function _nextExpenseId_() {
  var year   = String(new Date().getFullYear());
  var prefix = 'GAS-' + year + '-';

  var sheet;
  try { sheet = getSheet('GASTOS'); } catch (e) { return prefix + '001'; }

  var lastRow = sheet.getLastRow();
  var maxNum  = 0;

  if (lastRow > 1) {
    var ids = sheet.getRange(2, COL_GAS.ID, lastRow - 1, 1).getValues().flat();
    for (var i = 0; i < ids.length; i++) {
      var idStr = String(ids[i]);
      if (idStr.indexOf(prefix) === 0) {
        var n = parseInt(idStr.split('-').pop(), 10);
        if (!isNaN(n) && n > maxNum) maxNum = n;
      }
    }
  }

  return prefix + String(maxNum + 1).padStart(3, '0');
}
