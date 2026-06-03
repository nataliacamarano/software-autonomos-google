// ============================================================
// Gastos.gs — Módulo de gastos FASE 1: Estructura base
// ============================================================

// Columnas de la hoja GASTOS (base 1) — 23 columnas
var COL_GAS = {
  ID_GASTO:              1,
  FECHA_ALTA:            2,
  TIPO_GASTO:            3,
  FECHA_GASTO:           4,
  PROVEEDOR:             5,
  CIF_NIF_PROVEEDOR:     6,
  NUMERO_FACTURA_TICKET: 7,
  CONCEPTO:              8,
  CATEGORIA:             9,
  BASE_IMPONIBLE:        10,
  IVA_PORCENTAJE:        11,
  IVA_IMPORTE:           12,
  TOTAL:                 13,
  METODO_PAGO:           14,
  ESTADO_PAGO:           15,
  FOTO_URL:              16,
  ARCHIVO_ID:            17,
  OCR_TEXTO:             18,
  CONFIANZA_OCR:         19,
  REVISADO:              20,
  ENVIAR_GESTOR:         21,
  OBSERVACIONES:         22,
  ESTADO_PROCESAMIENTO:  23
};

var HEADERS_GASTOS = [
  'ID_Gasto', 'Fecha_Alta', 'Tipo_Gasto', 'Fecha_Gasto',
  'Proveedor', 'CIF_NIF_Proveedor', 'Numero_Factura_Ticket', 'Concepto', 'Categoria',
  'Base_Imponible', 'IVA_Porcentaje', 'IVA_Importe', 'Total',
  'Metodo_Pago', 'Estado_Pago',
  'Foto_URL', 'Archivo_ID', 'OCR_Texto', 'Confianza_OCR',
  'Revisado', 'Enviar_Gestor', 'Observaciones', 'Estado_Procesamiento'
];

var TIPOS_GASTO = [
  'Materiales de obra',
  'Herramientas y maquinaria',
  'Material de oficina',
  'Combustible y transporte',
  'Electricidad',
  'Agua',
  'Internet y teléfono',
  'Alquiler local/almacén',
  'Seguros',
  'Asesoría y gestoría',
  'Formación y cursos',
  'Software y licencias',
  'Mantenimiento y reparaciones',
  'Publicidad y marketing',
  'Dietas y representación',
  'Viajes y desplazamientos',
  'Ropa de trabajo y EPIs',
  'Subcontrataciones',
  'Gastos bancarios',
  'Otros servicios profesionales',
  'Otros'
];

var ESTADOS_PAGO_GASTO = ['Pendiente', 'Pagado'];

var GAS_CFG = {
  FORM_ID:   'gastos_form_id',
  FORM_URL:  'gastos_form_url',
  FOLDER_ID: 'gastos_folder_id'
};

// ─────────────────────────────────────────────────────────────
// installGastosModule  — ejecutar una sola vez (idempotente)
// ─────────────────────────────────────────────────────────────

/**
 * Crea la hoja GASTOS (23 columnas), una carpeta Drive para tickets,
 * el formulario Google con 4 campos, lo vincula al spreadsheet y
 * guarda los IDs en CONFIG.
 * Seguro ejecutar varias veces.
 */
function installGastosModule() {
  Logger.log('================================================');
  Logger.log('installGastosModule() — FASE 1');
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
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS_GASTOS);
    applyHeaderStyle_(sheet, HEADERS_GASTOS.length);
    Logger.log('  ✓ Cabeceras escritas (' + HEADERS_GASTOS.length + ' columnas).');
  } else {
    Logger.log('  ✓ Cabeceras existentes conservadas.');
  }

  // ── 2. Anchos de columna ──────────────────────────────────────
  Logger.log('\n[2] Anchos de columna:');
  var widths = [
    [COL_GAS.ID_GASTO,              120],
    [COL_GAS.FECHA_ALTA,            110],
    [COL_GAS.TIPO_GASTO,            175],
    [COL_GAS.FECHA_GASTO,           110],
    [COL_GAS.PROVEEDOR,             160],
    [COL_GAS.CIF_NIF_PROVEEDOR,     120],
    [COL_GAS.NUMERO_FACTURA_TICKET, 150],
    [COL_GAS.CONCEPTO,              220],
    [COL_GAS.CATEGORIA,             140],
    [COL_GAS.BASE_IMPONIBLE,        100],
    [COL_GAS.IVA_PORCENTAJE,         80],
    [COL_GAS.IVA_IMPORTE,            90],
    [COL_GAS.TOTAL,                  90],
    [COL_GAS.METODO_PAGO,           130],
    [COL_GAS.ESTADO_PAGO,           100],
    [COL_GAS.FOTO_URL,              220],
    [COL_GAS.ARCHIVO_ID,            160],
    [COL_GAS.OCR_TEXTO,             200],
    [COL_GAS.CONFIANZA_OCR,         100],
    [COL_GAS.REVISADO,               80],
    [COL_GAS.ENVIAR_GESTOR,         100],
    [COL_GAS.OBSERVACIONES,         200],
    [COL_GAS.ESTADO_PROCESAMIENTO,  150]
  ];
  for (var w = 0; w < widths.length; w++) {
    sheet.setColumnWidth(widths[w][0], widths[w][1]);
  }
  Logger.log('  ✓ Anchos ajustados.');

  // ── 3. Validaciones de datos ──────────────────────────────────
  Logger.log('\n[3] Validaciones:');
  var maxRows  = sheet.getMaxRows();
  var dataRows = Math.max(maxRows - 1, 1);

  sheet.getRange(2, COL_GAS.TIPO_GASTO, dataRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(TIPOS_GASTO, true).setAllowInvalid(true).build()
  );
  Logger.log('  ✓ Tipo_Gasto (' + TIPOS_GASTO.length + ' opciones).');

  sheet.getRange(2, COL_GAS.ESTADO_PAGO, dataRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(ESTADOS_PAGO_GASTO, true).setAllowInvalid(false).build()
  );
  Logger.log('  ✓ Estado_Pago → ' + ESTADOS_PAGO_GASTO.join(' | '));

  var ruleCheck = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange(2, COL_GAS.REVISADO,      dataRows, 1).setDataValidation(ruleCheck);
  sheet.getRange(2, COL_GAS.ENVIAR_GESTOR, dataRows, 1).setDataValidation(ruleCheck);
  Logger.log('  ✓ Revisado y Enviar_Gestor → checkbox.');

  // ── 4. Carpeta Drive para tickets ─────────────────────────────
  Logger.log('\n[4] Carpeta Drive para tickets:');
  var folderId = getConfig(GAS_CFG.FOLDER_ID);
  if (folderId) {
    try {
      DriveApp.getFolderById(folderId);
      Logger.log('  ✓ Ya existe. ID: ' + folderId);
    } catch (e) {
      Logger.log('  ⚠️  Carpeta no encontrada, recreando...');
      folderId = null;
    }
  }
  if (!folderId) {
    var folder = DriveApp.createFolder('Álvaro Gestión — Tickets y Facturas de Gastos');
    folderId = folder.getId();
    setConfig(GAS_CFG.FOLDER_ID, folderId);
    Logger.log('  + Carpeta creada. ID: ' + folderId);
  }

  // ── 5. Formulario Google ──────────────────────────────────────
  Logger.log('\n[5] Formulario GASTOS:');
  var formId = getConfig(GAS_CFG.FORM_ID);
  if (formId) {
    try {
      var existingForm = FormApp.openById(formId);
      Logger.log('  ✓ Ya existe.');
      Logger.log('    URL: ' + existingForm.getPublishedUrl());
    } catch (e) {
      Logger.log('  ⚠️  Formulario no encontrado, recreando...');
      formId = null;
    }
  }
  if (!formId) {
    var form = FormApp.create('GASTOS — Álvaro Gestión');
    form.setDescription('Registra un nuevo gasto: tipo, foto del ticket/factura y estado.');
    form.setCollectEmail(false);
    form.setAllowResponseEdits(false);
    form.setShowLinkToRespondAgain(true);

    // Campo 1: Tipo de gasto
    form.addListItem()
      .setTitle('Tipo de gasto')
      .setChoiceValues(TIPOS_GASTO)
      .setRequired(true);

    // Campo 2: URL de la foto del ticket / factura (Drive o fotos Google)
    form.addTextItem()
      .setTitle('URL de la foto del ticket / factura')
      .setHelpText('Sube la foto a Google Drive o Google Fotos y pega aquí el enlace compartido.')
      .setRequired(false);

    // Campo 3: Estado del gasto
    form.addListItem()
      .setTitle('Estado del gasto')
      .setChoiceValues(ESTADOS_PAGO_GASTO)
      .setRequired(true);

    // Campo 4: Observaciones
    form.addParagraphTextItem()
      .setTitle('Observaciones')
      .setRequired(false);

    formId = form.getId();
    setConfig(GAS_CFG.FORM_ID,  formId);
    setConfig(GAS_CFG.FORM_URL, form.getPublishedUrl());

    // Vincular al spreadsheet y renombrar hoja de respuestas
    var snapshot = ss.getSheets().map(function(s) { return s.getName(); });
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    renameNewSheet_(ss, snapshot, 'RESP_GASTOS');

    Logger.log('  + Formulario creado.');
    Logger.log('    URL: ' + form.getPublishedUrl());
    Logger.log('    ID:  ' + formId);
  }

  // ── Resumen ───────────────────────────────────────────────────
  Logger.log('\n================================================');
  Logger.log('MÓDULO GASTOS FASE 1 INSTALADO.');
  Logger.log('  Hoja GASTOS:    ✓  (' + HEADERS_GASTOS.length + ' columnas)');
  Logger.log('  Carpeta Drive:  ✓  ID=' + folderId);
  Logger.log('  Formulario:     ✓  ID=' + formId);
  Logger.log('  URL formulario: ' + (getConfig(GAS_CFG.FORM_URL) || '—'));
  Logger.log('\nSiguiente paso: el trigger onFormSubmitHandler ya');
  Logger.log('procesará las respuestas del formulario automáticamente.');
  Logger.log('================================================');
}

// ─────────────────────────────────────────────────────────────
// validateGastosModule
// ─────────────────────────────────────────────────────────────

/**
 * Verifica que el módulo esté instalado correctamente.
 * SOLO LECTURA.
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
    Logger.log('  ✓ GID=' + sheet.getSheetId() + '  registros=' + Math.max(0, sheet.getLastRow() - 1));
  } catch (e) {
    Logger.log('  ❌ NO EXISTE — ejecuta installGastosModule().');
    return;
  }

  // 2. Cabeceras
  Logger.log('\n[2] Cabeceras (' + HEADERS_GASTOS.length + ' columnas):');
  var lastCol = sheet.getLastColumn();
  var hdrs    = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var hdrsOk  = true;
  for (var i = 0; i < HEADERS_GASTOS.length; i++) {
    var presente = hdrs.indexOf(HEADERS_GASTOS[i]) !== -1;
    Logger.log('  ' + (presente ? '✓' : '❌') + ' ' + HEADERS_GASTOS[i]);
    if (!presente) { hdrsOk = false; ok = false; }
  }

  // 3. Carpeta Drive
  Logger.log('\n[3] Carpeta Drive:');
  var folderId = getConfig(GAS_CFG.FOLDER_ID);
  if (folderId) {
    try {
      var folder = DriveApp.getFolderById(folderId);
      Logger.log('  ✓ "' + folder.getName() + '"  ID=' + folderId);
    } catch (e) {
      Logger.log('  ❌ Carpeta no encontrada (ID=' + folderId + ').');
      ok = false;
    }
  } else {
    Logger.log('  ❌ Sin ID de carpeta — ejecuta installGastosModule().');
    ok = false;
  }

  // 4. Formulario
  Logger.log('\n[4] Formulario GASTOS:');
  var formId = getConfig(GAS_CFG.FORM_ID);
  if (formId) {
    try {
      var form = FormApp.openById(formId);
      Logger.log('  ✓ "' + form.getTitle() + '"');
      Logger.log('    URL: ' + form.getPublishedUrl());
      Logger.log('    Campos: ' + form.getItems().length);
    } catch (e) {
      Logger.log('  ❌ Formulario no encontrado (ID=' + formId + ').');
      ok = false;
    }
  } else {
    Logger.log('  ❌ Sin ID de formulario — ejecuta installGastosModule().');
    ok = false;
  }

  // 5. Hoja RESP_GASTOS
  Logger.log('\n[5] Hoja RESP_GASTOS (respuestas del formulario):');
  var respSheet = getTargetSS_().getSheetByName('RESP_GASTOS');
  if (respSheet) {
    Logger.log('  ✓ Existe. Respuestas: ' + Math.max(0, respSheet.getLastRow() - 1));
  } else {
    Logger.log('  ⚠️  No encontrada — se creará al vincular el formulario.');
  }

  // 6. Siguiente ID
  Logger.log('\n[6] Siguiente ID de gasto:');
  try {
    Logger.log('  ' + _nextGastoId_());
  } catch (e) {
    Logger.log('  Error: ' + e.message);
  }

  Logger.log('\n================================================');
  Logger.log('RESULTADO: ' + (ok ? '✅ MÓDULO LISTO.' : '❌ HAY PROBLEMAS — revisa los ❌.'));
  Logger.log('================================================');
}

// ─────────────────────────────────────────────────────────────
// onGastoFormSubmit  — llamado desde onFormSubmitHandler
// ─────────────────────────────────────────────────────────────

/**
 * Procesa una respuesta del formulario GASTOS y escribe en la hoja GASTOS.
 * Llamado desde onFormSubmitHandler() en Triggers.gs cuando la hoja es RESP_GASTOS.
 */
function onGastoFormSubmit(namedValues) {
  try {
    var sheet  = getSheet('GASTOS');
    var id     = _nextGastoId_();
    var ahora  = new Date();

    var tipoGasto   = String((namedValues['Tipo de gasto']                      || [''])[0]).trim();
    var fotoUrl     = String((namedValues['URL de la foto del ticket / factura']|| [''])[0]).trim();
    var estadoPago  = String((namedValues['Estado del gasto']                   || [''])[0]).trim() || 'Pendiente';
    var observaciones = String((namedValues['Observaciones']                    || [''])[0]).trim();

    // Extraer el ID del archivo de Drive desde la URL
    var archivoId = '';
    var m = fotoUrl.match(/[-\w]{25,}/);
    if (m) archivoId = m[0];

    var row = new Array(HEADERS_GASTOS.length).fill('');
    row[COL_GAS.ID_GASTO              - 1] = id;
    row[COL_GAS.FECHA_ALTA            - 1] = ahora;
    row[COL_GAS.TIPO_GASTO            - 1] = tipoGasto;
    row[COL_GAS.ESTADO_PAGO           - 1] = estadoPago;
    row[COL_GAS.FOTO_URL              - 1] = fotoUrl;
    row[COL_GAS.ARCHIVO_ID            - 1] = archivoId;
    row[COL_GAS.REVISADO              - 1] = false;
    row[COL_GAS.ENVIAR_GESTOR         - 1] = false;
    row[COL_GAS.OBSERVACIONES         - 1] = observaciones;
    row[COL_GAS.ESTADO_PROCESAMIENTO  - 1] = 'Pendiente OCR';

    sheet.appendRow(row);
    logEvent('INFO', 'onGastoFormSubmit', 'Gasto registrado: ' + id, tipoGasto);
    return { success: true, id: id };
  } catch (e) {
    logEvent('ERROR', 'onGastoFormSubmit', e.message, JSON.stringify(namedValues));
    return { success: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
// calculateVATSummary
// ─────────────────────────────────────────────────────────────

/**
 * Resumen de IVA repercutido (Out, de FACTURAS), soportado (In, de GASTOS) y neto.
 * SOLO LECTURA.
 */
function calculateVATSummary() {
  Logger.log('================================================');
  Logger.log('calculateVATSummary()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  // ── IVA OUT — FACTURAS ────────────────────────────────────────
  Logger.log('\n[1] IVA OUT — Facturas emitidas:');
  var ivaOutPorAnio = {}, ivaOutTotal = 0;
  try {
    var shFac   = getSheet('FACTURAS');
    var lastFac = shFac.getLastRow();
    if (lastFac > 1) {
      var facHdrs       = shFac.getRange(1, 1, 1, shFac.getLastColumn()).getValues()[0];
      var colFechaFac   = -1, colIvaImporte = -1;
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
            var fv = facData[f][colFechaFac - 1];
            if (fv instanceof Date && !isNaN(fv.getTime())) anio = String(fv.getFullYear());
            else if (fv) { var y = String(fv).match(/\d{4}/); if (y) anio = y[0]; }
          }
          ivaOutPorAnio[anio] = (ivaOutPorAnio[anio] || 0) + importe;
          ivaOutTotal += importe;
        }
        Object.keys(ivaOutPorAnio).sort().forEach(function(a) {
          Logger.log('  ' + a + ': ' + formatCurrency(ivaOutPorAnio[a]));
        });
        Logger.log('  TOTAL IVA OUT: ' + formatCurrency(ivaOutTotal));
      }
    } else {
      Logger.log('  ⚠️  Sin facturas registradas.');
    }
  } catch (e) {
    Logger.log('  ❌ Error: ' + e.message);
  }

  // ── IVA IN — GASTOS ───────────────────────────────────────────
  Logger.log('\n[2] IVA IN — Gastos (IVA soportado):');
  var ivaInPorAnio = {}, ivaInTotal = 0;
  try {
    var shGas   = getSheet('GASTOS');
    var lastGas = shGas.getLastRow();
    if (lastGas > 1) {
      var gasData = shGas.getRange(2, 1, lastGas - 1, HEADERS_GASTOS.length).getValues();
      for (var g = 0; g < gasData.length; g++) {
        var ivaGasto = parseFloat(gasData[g][COL_GAS.IVA_IMPORTE - 1]) || 0;
        var fechaGas = gasData[g][COL_GAS.FECHA_GASTO - 1] || gasData[g][COL_GAS.FECHA_ALTA - 1];
        var anioG    = '—';
        if (fechaGas instanceof Date && !isNaN(fechaGas.getTime())) {
          anioG = String(fechaGas.getFullYear());
        } else if (fechaGas) {
          var yg = String(fechaGas).match(/\d{4}/); if (yg) anioG = yg[0];
        }
        ivaInPorAnio[anioG] = (ivaInPorAnio[anioG] || 0) + ivaGasto;
        ivaInTotal += ivaGasto;
      }
      Object.keys(ivaInPorAnio).sort().forEach(function(a) {
        Logger.log('  ' + a + ': ' + formatCurrency(ivaInPorAnio[a]));
      });
      Logger.log('  TOTAL IVA IN: ' + formatCurrency(ivaInTotal));
    } else {
      Logger.log('  ⚠️  Sin gastos registrados.');
    }
  } catch (e) {
    Logger.log('  ❌ Error: ' + e.message);
  }

  // ── IVA Neto ──────────────────────────────────────────────────
  Logger.log('\n[3] IVA NETO (Out − In):');
  var anioActual   = String(new Date().getFullYear());
  var ivaOutAnio   = ivaOutPorAnio[anioActual] || 0;
  var ivaInAnio    = ivaInPorAnio[anioActual]  || 0;
  var ivaNetoAnio  = ivaOutAnio - ivaInAnio;
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

// ── Privadas ──────────────────────────────────────────────────

/**
 * Genera el siguiente ID de gasto del año en curso.
 * Formato: GAS-YYYY-NNN
 */
function _nextGastoId_() {
  var year   = String(new Date().getFullYear());
  var prefix = 'GAS-' + year + '-';
  var sheet;
  try { sheet = getSheet('GASTOS'); } catch (e) { return prefix + '001'; }

  var lastRow = sheet.getLastRow();
  var maxNum  = 0;
  if (lastRow > 1) {
    var ids = sheet.getRange(2, COL_GAS.ID_GASTO, lastRow - 1, 1).getValues().flat();
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
