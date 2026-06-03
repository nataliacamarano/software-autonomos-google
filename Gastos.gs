// ============================================================
// Gastos.gs — Módulo de gastos FASE 2: Recepción y normalización
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
  'Combustible',
  'Materiales',
  'Herramientas',
  'Recambios',
  'Productos químicos',
  'Transporte',
  'Parking',
  'Peajes',
  'Comidas',
  'Alojamiento',
  'Telefonía',
  'Internet',
  'Seguros',
  'Gestoría',
  'Alquiler maquinaria',
  'Vehículos',
  'Ropa laboral',
  'Publicidad',
  'Formación',
  'Gasto recurrente',
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
 * FASE 2: copia foto a carpeta Drive, detecta falta de foto, evita duplicados.
 */
function onGastoFormSubmit(namedValues) {
  try {
    var sheet = getSheet('GASTOS');
    var ahora = new Date();

    var tipoGasto     = String((namedValues['Tipo de gasto']                      || [''])[0]).trim();
    var fotoUrlRaw    = String((namedValues['URL de la foto del ticket / factura'] || [''])[0]).trim();
    var estadoPago    = String((namedValues['Estado del gasto']                   || [''])[0]).trim() || 'Pendiente';
    var observaciones = String((namedValues['Observaciones']                      || [''])[0]).trim();

    // Deduplicación: mismo tipo + foto + estado en los últimos 10 minutos
    if (_gastoYaRegistrado_(sheet, tipoGasto, fotoUrlRaw, ahora)) {
      logEvent('WARN', 'onGastoFormSubmit', 'Duplicado detectado — omitido.', fotoUrlRaw);
      return { success: true, id: null, skipped: true };
    }

    var id         = _nextGastoId_();
    var fotoUrl    = fotoUrlRaw;
    var archivoId  = '';
    var estadoProc = 'Pendiente OCR';

    if (fotoUrlRaw) {
      // Intentar copiar el archivo a la carpeta de gastos
      var copia = _copiarArchivoAGastos_(fotoUrlRaw);
      if (copia) {
        fotoUrl   = copia.url;
        archivoId = copia.id;
      } else {
        // Fallback: extraer ID del archivo de la URL original
        var m = fotoUrlRaw.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
        if (!m) m = fotoUrlRaw.match(/[?&]id=([a-zA-Z0-9_-]{25,})/);
        if (m) archivoId = m[1];
      }
    } else {
      estadoProc    = 'Revisar';
      var notaFalta = 'Falta foto del ticket/factura';
      observaciones = observaciones ? (observaciones + ' | ' + notaFalta) : notaFalta;
    }

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
    row[COL_GAS.ESTADO_PROCESAMIENTO  - 1] = estadoProc;

    sheet.appendRow(row);
    logEvent('INFO', 'onGastoFormSubmit', 'Gasto registrado: ' + id, tipoGasto + ' | ' + estadoProc);
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

// ─────────────────────────────────────────────────────────────
// repararGastosFase1  — ejecutar una sola vez para corregir
// ─────────────────────────────────────────────────────────────

/**
 * Corrige la instalación cuando la hoja GASTOS tiene el esquema antiguo
 * y RESP_GASTOS no existe.
 *
 * Acciones:
 *   1. Si GASTOS no tiene cabecera ID_Gasto → la renombra a GASTOS_LEGACY.
 *   2. Crea GASTOS nueva con 23 columnas, validaciones y anchos.
 *   3. Desvincula el formulario de su destino actual y lo re-vincula al
 *      spreadsheet → crea RESP_GASTOS.
 */
function repararGastosFase1() {
  Logger.log('================================================');
  Logger.log('repararGastosFase1()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var ss = getTargetSS_();

  // ── 1. Renombrar hoja GASTOS antigua ─────────────────────────
  Logger.log('\n[1] Hoja GASTOS:');
  var existente = ss.getSheetByName('GASTOS');
  if (existente) {
    var primeraCol = existente.getLastRow() > 0
      ? String(existente.getRange(1, 1).getValue()).trim() : '';
    if (primeraCol !== 'ID_Gasto') {
      var legacyName = 'GASTOS_LEGACY';
      if (ss.getSheetByName(legacyName)) legacyName = 'GASTOS_LEGACY_' + Date.now();
      existente.setName(legacyName);
      Logger.log('  ✓ Hoja antigua renombrada a "' + legacyName + '" (' +
                 Math.max(0, existente.getLastRow() - 1) + ' registros preservados).');
    } else {
      Logger.log('  ✓ Ya tiene el esquema correcto (ID_Gasto en col 1). Sin cambios.');
    }
  } else {
    Logger.log('  ℹ️  No existía hoja GASTOS.');
  }

  // ── 2. Crear GASTOS nueva con 23 columnas ────────────────────
  Logger.log('\n[2] Creando nueva hoja GASTOS:');
  var sheet = ss.getSheetByName('GASTOS') || ss.insertSheet('GASTOS');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS_GASTOS);
    applyHeaderStyle_(sheet, HEADERS_GASTOS.length);
    Logger.log('  ✓ Cabeceras escritas (' + HEADERS_GASTOS.length + ' columnas).');
  } else {
    Logger.log('  ✓ Cabeceras existentes conservadas.');
  }

  // Anchos
  var widths = [
    [COL_GAS.ID_GASTO,120],[COL_GAS.FECHA_ALTA,110],[COL_GAS.TIPO_GASTO,175],
    [COL_GAS.FECHA_GASTO,110],[COL_GAS.PROVEEDOR,160],[COL_GAS.CIF_NIF_PROVEEDOR,120],
    [COL_GAS.NUMERO_FACTURA_TICKET,150],[COL_GAS.CONCEPTO,220],[COL_GAS.CATEGORIA,140],
    [COL_GAS.BASE_IMPONIBLE,100],[COL_GAS.IVA_PORCENTAJE,80],[COL_GAS.IVA_IMPORTE,90],
    [COL_GAS.TOTAL,90],[COL_GAS.METODO_PAGO,130],[COL_GAS.ESTADO_PAGO,100],
    [COL_GAS.FOTO_URL,220],[COL_GAS.ARCHIVO_ID,160],[COL_GAS.OCR_TEXTO,200],
    [COL_GAS.CONFIANZA_OCR,100],[COL_GAS.REVISADO,80],[COL_GAS.ENVIAR_GESTOR,100],
    [COL_GAS.OBSERVACIONES,200],[COL_GAS.ESTADO_PROCESAMIENTO,150]
  ];
  widths.forEach(function(w) { sheet.setColumnWidth(w[0], w[1]); });

  // Validaciones
  var maxRows  = sheet.getMaxRows();
  var dataRows = Math.max(maxRows - 1, 1);
  sheet.getRange(2, COL_GAS.TIPO_GASTO, dataRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(TIPOS_GASTO, true).setAllowInvalid(true).build());
  sheet.getRange(2, COL_GAS.ESTADO_PAGO, dataRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS_PAGO_GASTO, true).setAllowInvalid(false).build());
  var ruleCheck = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange(2, COL_GAS.REVISADO,      dataRows, 1).setDataValidation(ruleCheck);
  sheet.getRange(2, COL_GAS.ENVIAR_GESTOR, dataRows, 1).setDataValidation(ruleCheck);
  Logger.log('  ✓ Validaciones y anchos aplicados.');

  // ── 3. Reconectar formulario → RESP_GASTOS ───────────────────
  Logger.log('\n[3] Reconectando formulario:');
  var formId = getConfig(GAS_CFG.FORM_ID);
  if (!formId) {
    Logger.log('  ❌ Sin form_id en CONFIG. Ejecuta installGastosModule() primero.');
  } else {
    try {
      var form = FormApp.openById(formId);
      try { form.removeDestination(); Utilities.sleep(1000); } catch (_) {}

      var snapshot = ss.getSheets().map(function(s) { return s.getName(); });
      form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
      var ok = renameNewSheet_(ss, snapshot, 'RESP_GASTOS');
      if (ok) {
        Logger.log('  ✓ Formulario reconectado → RESP_GASTOS creada.');
      } else {
        Logger.log('  ⚠️  setDestination enviado pero RESP_GASTOS no apareció en 12s.');
        Logger.log('      Espera 30s y ejecuta validateGastosModule() para verificar.');
      }
    } catch (e) {
      Logger.log('  ❌ Error: ' + e.message);
      logEvent('ERROR', 'repararGastosFase1', e.message, '');
    }
  }

  Logger.log('\n================================================');
  Logger.log('REPARACIÓN COMPLETADA.');
  Logger.log('  Ejecuta validateGastosModule() para verificar.');
  Logger.log('================================================');
}

// ─────────────────────────────────────────────────────────────
// reprocesarRespGastos  — procesa respuestas no grabadas en GASTOS
// ─────────────────────────────────────────────────────────────

/**
 * Lee RESP_GASTOS fila a fila y llama a onGastoFormSubmit para cada respuesta
 * que no tenga ya un ID correspondiente en GASTOS.
 * Útil cuando el trigger no pudo procesar respuestas anteriores.
 */
function reprocesarRespGastos() {
  Logger.log('=== reprocesarRespGastos() ===');
  var ss       = getTargetSS_();
  var respSh   = ss.getSheetByName('RESP_GASTOS');
  if (!respSh) { Logger.log('  ❌ RESP_GASTOS no encontrada.'); return; }

  var lastRow = respSh.getLastRow();
  if (lastRow < 2) { Logger.log('  Sin respuestas.'); return; }

  // Leer cabeceras y datos
  var all     = respSh.getRange(1, 1, lastRow, respSh.getLastColumn()).getValues();
  var headers = all[0].map(function(h) { return String(h).trim(); });
  Logger.log('  Cabeceras RESP_GASTOS: ' + headers.join(' | '));
  Logger.log('  Respuestas encontradas: ' + (lastRow - 1));

  var procesadas = 0, omitidas = 0;

  for (var i = 1; i < all.length; i++) {
    var row = all[i];
    // Construir namedValues igual que lo haría el trigger
    var nv = {};
    headers.forEach(function(h, idx) {
      nv[h] = [String(row[idx] || '').trim()];
    });

    var fotoUrl = String((nv['URL de la foto del ticket / factura'] || [''])[0]).trim();
    var tipo    = String((nv['Tipo de gasto'] || [''])[0]).trim();
    Logger.log('  Fila ' + (i+1) + ': tipo="' + tipo + '"  foto="' + fotoUrl.substring(0,60) + '"');

    var result = onGastoFormSubmit(nv);
    if (result.skipped) {
      Logger.log('    → omitida (ya existe en GASTOS).');
      omitidas++;
    } else if (result.success) {
      Logger.log('    → registrada como ' + result.id);
      procesadas++;
    } else {
      Logger.log('    → ERROR: ' + result.error);
    }

    if (i < all.length - 1) Utilities.sleep(500);
  }

  Logger.log('\nRESUMEN: procesadas=' + procesadas + '  omitidas=' + omitidas);
  Logger.log('=== FIN reprocesarRespGastos() ===');
}

// ─────────────────────────────────────────────────────────────
// vincularFormularioGastos  — solo reconecta el formulario
// ─────────────────────────────────────────────────────────────

/**
 * Busca la hoja de respuestas del formulario (puede haberse creado con
 * otro nombre) y la renombra a RESP_GASTOS.
 * Si no existe ninguna, vuelve a llamar a setDestination y espera 30s.
 */
function vincularFormularioGastos() {
  Logger.log('=== vincularFormularioGastos() ===');
  var ss = getTargetSS_();

  // Si ya existe, perfecto
  if (ss.getSheetByName('RESP_GASTOS')) {
    Logger.log('  ✓ RESP_GASTOS ya existe.');
    return;
  }

  // Buscar hoja de respuestas de formulario ya creada (cualquier nombre genérico)
  var sheets  = ss.getSheets();
  var renamed = false;
  for (var i = sheets.length - 1; i >= 0; i--) {
    var n = sheets[i].getName();
    if (/respuestas.de.formulario|form.responses/i.test(n)) {
      sheets[i].setName('RESP_GASTOS');
      Logger.log('  ✓ "' + n + '" renombrada a RESP_GASTOS.');
      renamed = true;
      break;
    }
  }
  if (renamed) return;

  // No hay ninguna — volver a llamar setDestination y esperar 30s
  Logger.log('  No encontrada. Re-enviando setDestination (espera 30s)...');
  var formId = getConfig(GAS_CFG.FORM_ID);
  if (!formId) { Logger.log('  ❌ Sin form_id en CONFIG.'); return; }

  try {
    var form     = FormApp.openById(formId);
    var snapshot = sheets.map(function(s) { return s.getName(); });
    try { form.removeDestination(); Utilities.sleep(2000); } catch (_) {}
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

    // Esperar hasta 30s
    for (var t = 0; t < 15; t++) {
      Utilities.sleep(2000);
      var nueva = ss.getSheets().find(function(s) {
        return !snapshot.includes(s.getName()) ||
               /respuestas.de.formulario|form.responses/i.test(s.getName());
      });
      if (nueva) { nueva.setName('RESP_GASTOS'); Logger.log('  ✓ RESP_GASTOS creada.'); return; }
      if (ss.getSheetByName('RESP_GASTOS')) { Logger.log('  ✓ RESP_GASTOS creada.'); return; }
    }
    Logger.log('  ⚠️  Timeout. Comprueba manualmente si apareció una hoja nueva y renómbrala a RESP_GASTOS.');
  } catch (e) {
    Logger.log('  ❌ Error: ' + e.message);
  }
}

// ─────────────────────────────────────────────────────────────
// _borrarFormulariosGastosHuerfanos_  — ejecutar una sola vez
// ─────────────────────────────────────────────────────────────

/**
 * Mueve a la papelera todos los formularios llamados
 * "GASTOS — Álvaro Gestión" excepto el guardado en CONFIG.
 */
function borrarFormulariosGastosHuerfanos() {
  return _borrarFormulariosGastosHuerfanos_();
}

function _borrarFormulariosGastosHuerfanos_() {
  var formIdCorrecto = getConfig(GAS_CFG.FORM_ID);
  if (!formIdCorrecto) { Logger.log('❌ Sin form_id en CONFIG.'); return; }

  var files   = DriveApp.getFilesByName('GASTOS — Álvaro Gestión');
  var borrados = 0;
  while (files.hasNext()) {
    var f = files.next();
    if (f.getId() !== formIdCorrecto) {
      f.setTrashed(true);
      Logger.log('🗑️  Papelera: ' + f.getName() + ' (' + f.getId() + ')');
      borrados++;
    } else {
      Logger.log('✓  Conservado (correcto): ' + f.getId());
    }
  }
  Logger.log('Formularios movidos a papelera: ' + borrados);
}

// ─────────────────────────────────────────────────────────────
// actualizarFormularioGastos — sincroniza tipos en el form existente
// ─────────────────────────────────────────────────────────────

/**
 * Actualiza las opciones del campo "Tipo de gasto" en el formulario existente
 * para que coincidan con TIPOS_GASTO. Ejecutar una vez tras cambiar la lista.
 */
function actualizarFormularioGastos() {
  var formId = getConfig(GAS_CFG.FORM_ID);
  if (!formId) { Logger.log('❌ Sin form_id en CONFIG. Ejecuta installGastosModule().'); return; }
  try {
    var form  = FormApp.openById(formId);
    var items = form.getItems(FormApp.ItemType.LIST);
    var updated = 0;
    for (var i = 0; i < items.length; i++) {
      if (items[i].getTitle() === 'Tipo de gasto') {
        items[i].asListItem().setChoiceValues(TIPOS_GASTO);
        Logger.log('✓ "Tipo de gasto" actualizado (' + TIPOS_GASTO.length + ' opciones).');
        updated++;
      }
      if (items[i].getTitle() === 'Estado del gasto') {
        items[i].asListItem().setChoiceValues(ESTADOS_PAGO_GASTO);
        Logger.log('✓ "Estado del gasto" actualizado.');
        updated++;
      }
    }
    if (updated === 0) Logger.log('⚠️  No se encontraron campos LIST en el formulario.');
    Logger.log('URL: ' + form.getPublishedUrl());
  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
    logEvent('ERROR', 'actualizarFormularioGastos', e.message, '');
  }
}

// ─────────────────────────────────────────────────────────────
// testGastosFase2
// ─────────────────────────────────────────────────────────────

/**
 * Prueba onGastoFormSubmit con datos simulados.
 * Crea 2 filas en GASTOS: una con foto, otra sin foto.
 * Verifica deduplicación ejecutando el mismo registro dos veces.
 */
function testGastosFase2() {
  Logger.log('================================================');
  Logger.log('testGastosFase2()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  // URL única por ejecución para evitar colisión con runs previos
  var fotoUrlTest = 'https://drive.google.com/file/d/TEST_' + Date.now() + '/view';

  // Test 1: gasto con URL de foto
  Logger.log('\n[1] Gasto con foto (URL Drive simulada):');
  var nv1 = {
    'Tipo de gasto':                       ['Combustible'],
    'URL de la foto del ticket / factura': [fotoUrlTest],
    'Estado del gasto':                    ['Pagado'],
    'Observaciones':                       ['Test FASE 2 — con foto']
  };
  var r1 = onGastoFormSubmit(nv1);
  Logger.log('  Resultado: ' + JSON.stringify(r1));

  // Test 2: gasto sin foto → Estado_Procesamiento = "Revisar"
  Logger.log('\n[2] Gasto sin foto:');
  var nv2 = {
    'Tipo de gasto':                       ['Materiales'],
    'URL de la foto del ticket / factura': [''],
    'Estado del gasto':                    ['Pendiente'],
    'Observaciones':                       ['Test FASE 2 — sin foto']
  };
  var r2 = onGastoFormSubmit(nv2);
  Logger.log('  Resultado: ' + JSON.stringify(r2));

  // Test 3: duplicado — misma URL que Test 1, dentro de 3 min → debe omitirse
  Logger.log('\n[3] Duplicado del Test 1 (debe skipped=true):');
  var r3 = onGastoFormSubmit(nv1);
  Logger.log('  Resultado: ' + JSON.stringify(r3));

  Logger.log('\n================================================');
  Logger.log('VERIFICAR EN HOJA GASTOS:');
  Logger.log('  Fila ' + r1.id + ': Estado_Procesamiento = "Pendiente OCR"');
  Logger.log('  Fila ' + r2.id + ': Estado_Procesamiento = "Revisar"');
  Logger.log('               Observaciones contiene "Falta foto..."');
  Logger.log('  Test 3: skipped=true (no se añadió fila duplicada)');
  Logger.log('================================================');
}

// ── Privadas ──────────────────────────────────────────────────

/**
 * Copia un archivo de Drive (dado por URL) a la carpeta de gastos.
 * Devuelve { id, url } del archivo copiado, o null si falla.
 */
function _copiarArchivoAGastos_(url) {
  try {
    var m = url.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
    if (!m) m = url.match(/[?&]id=([a-zA-Z0-9_-]{25,})/);
    if (!m) return null;

    var fileId   = m[1];
    var folderId = getConfig(GAS_CFG.FOLDER_ID);
    if (!folderId) return null;

    var folder = DriveApp.getFolderById(folderId);
    var file   = DriveApp.getFileById(fileId);
    var copia  = file.makeCopy(file.getName(), folder);
    copia.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return { id: copia.getId(), url: copia.getUrl() };
  } catch (e) {
    logEvent('WARN', '_copiarArchivoAGastos_', e.message, url);
    return null;
  }
}

/**
 * Devuelve true si ya existe en GASTOS un registro con la misma fotoUrl
 * creado en los últimos 3 minutos (doble disparo del trigger).
 * Solo aplica cuando fotoUrl no está vacía: dos envíos sin foto son legítimos.
 */
function _gastoYaRegistrado_(sheet, tipoGasto, fotoUrl, ahora) {
  if (!fotoUrl) return false; // sin foto no se deduplica
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var data   = sheet.getRange(2, 1, lastRow - 1, HEADERS_GASTOS.length).getValues();
  var limite = 3 * 60 * 1000; // 3 minutos en ms
  for (var i = data.length - 1; i >= 0; i--) {
    var rowFoto  = String(data[i][COL_GAS.FOTO_URL  - 1]);
    var rowFecha = data[i][COL_GAS.FECHA_ALTA - 1];
    if (rowFoto === fotoUrl && rowFecha instanceof Date) {
      if (ahora.getTime() - rowFecha.getTime() < limite) return true;
    }
  }
  return false;
}

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

// ============================================================
// FASE 3 — OCR automático de tickets
// ============================================================

// ─────────────────────────────────────────────────────────────
// procesarOCRGasto  — procesa un gasto por ID
// ─────────────────────────────────────────────────────────────

/**
 * Extrae texto OCR de la imagen del ticket, parsea los datos y
 * actualiza la fila correspondiente en la hoja GASTOS.
 * No sobrescribe si Revisado = TRUE.
 */
function procesarOCRGasto(idGasto) {
  Logger.log('=== procesarOCRGasto(' + idGasto + ') ===');
  try {
    var sheet  = getSheet('GASTOS');
    var rowNum = findRowById_(sheet, idGasto);
    if (!rowNum) {
      Logger.log('  ❌ ID no encontrado: ' + idGasto);
      return { success: false, error: 'ID no encontrado: ' + idGasto };
    }

    var rowData  = sheet.getRange(rowNum, 1, 1, HEADERS_GASTOS.length).getValues()[0];
    var revisado = rowData[COL_GAS.REVISADO - 1];
    if (revisado === true || String(revisado).toLowerCase() === 'true') {
      Logger.log('  ⚠️  Revisado=TRUE — omitido.');
      return { success: true, skipped: true, reason: 'Revisado manualmente' };
    }

    var archivoId = String(rowData[COL_GAS.ARCHIVO_ID - 1] || '').trim();
    if (!archivoId) {
      Logger.log('  ⚠️  Sin Archivo_ID.');
      sheet.getRange(rowNum, COL_GAS.ESTADO_PROCESAMIENTO).setValue('OCR fallido');
      return { success: true, estado: 'OCR fallido', reason: 'Sin Archivo_ID' };
    }

    // OCR
    Logger.log('  OCR en archivo: ' + archivoId);
    var ocr      = _extraerTextoOCR_(archivoId);
    Logger.log('  OCR exito=' + ocr.exito + '  chars=' + ocr.texto.length + (ocr.error ? '  err=' + ocr.error : ''));

    // Parsear y calcular
    var datos     = _parsearTextoTicket_(ocr.texto);
    var confianza = _calcularConfianzaOCR_(datos, ocr.exito);
    var estado    = _determinarEstadoOCR_(datos, ocr.exito);

    Logger.log('  total=' + datos.total + '  base=' + datos.base +
               '  iva=' + datos.ivaImporte + '  fecha=' + datos.fecha +
               '  proveedor="' + datos.proveedor + '"');
    Logger.log('  Confianza=' + confianza + '  Estado=' + estado);

    _actualizarGastoConOCR_(sheet, rowNum, datos, ocr.texto, confianza, estado);

    logEvent('INFO', 'procesarOCRGasto', idGasto + ' → ' + estado, 'confianza=' + confianza);
    return { success: true, id: idGasto, estado: estado, confianza: confianza };

  } catch (e) {
    logEvent('ERROR', 'procesarOCRGasto', e.message, idGasto);
    return { success: false, error: e.message };
  }
}

/**
 * Procesa el OCR del último gasto registrado en GASTOS.
 * Útil para probar tras registrar un nuevo gasto.
 */
function procesarOCRUltimoGasto() {
  var sheet   = getSheet('GASTOS');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Sin gastos en GASTOS.'); return; }

  // Buscar el último ID_Gasto no vacío
  var ids = sheet.getRange(2, COL_GAS.ID_GASTO, lastRow - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    var id = String(ids[i][0]).trim();
    if (id) return procesarOCRGasto(id);
  }
  Logger.log('No se encontró ningún ID_Gasto.');
}

// ─────────────────────────────────────────────────────────────
// procesarGastosPendientesOCR  — batch sobre todos los pendientes
// ─────────────────────────────────────────────────────────────

/**
 * Procesa todos los gastos con Estado_Procesamiento = "Pendiente OCR"
 * que no estén marcados como Revisado.
 */
function procesarGastosPendientesOCR() {
  Logger.log('================================================');
  Logger.log('procesarGastosPendientesOCR()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var sheet   = getSheet('GASTOS');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('  Sin gastos registrados.'); return; }

  // Leer solo columnas relevantes para el filtro (ID, Revisado, Estado)
  var colsIds    = sheet.getRange(2, COL_GAS.ID_GASTO,             lastRow - 1, 1).getValues();
  var colsRev    = sheet.getRange(2, COL_GAS.REVISADO,             lastRow - 1, 1).getValues();
  var colsEst    = sheet.getRange(2, COL_GAS.ESTADO_PROCESAMIENTO, lastRow - 1, 1).getValues();
  var pendientes = [];
  for (var i = 0; i < colsIds.length; i++) {
    var id       = String(colsIds[i][0]).trim();
    var estado   = String(colsEst[i][0]).trim();
    var revisado = colsRev[i][0];
    if (id && estado === 'Pendiente OCR' &&
        revisado !== true && String(revisado).toLowerCase() !== 'true') {
      pendientes.push(id);
    }
  }

  Logger.log('  Gastos pendientes: ' + pendientes.length);
  var ok = 0, err = 0;
  for (var j = 0; j < pendientes.length; j++) {
    var r = procesarOCRGasto(pendientes[j]);
    if (r.success) ok++; else err++;
    if (j < pendientes.length - 1) Utilities.sleep(1500);
  }

  Logger.log('\nRESUMEN: procesados=' + ok + '  errores=' + err);
  Logger.log('================================================');
}

// ─────────────────────────────────────────────────────────────
// testGastosFase3
// ─────────────────────────────────────────────────────────────

/**
 * [1] Prueba el parser con un ticket simulado (sin necesidad de imagen real).
 * [2] Llama a procesarGastosPendientesOCR() sobre los gastos reales.
 */
function testGastosFase3() {
  Logger.log('================================================');
  Logger.log('testGastosFase3()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  // ── 1. Parser con texto simulado ──────────────────────────────
  Logger.log('\n[1] Parser con texto de ticket simulado:');
  var textoDemo = [
    'FERRETERÍA MARTÍNEZ S.L.',
    'CIF: B12345678',
    'C/ Mayor 45, Madrid',
    'TICKET Nº: T-2026-1234',
    'Fecha: 03/06/2026',
    '',
    'Tornillos acero M8 x50      8,26',
    'Cinta métrica 5m            12,40',
    '',
    'BASE IMPONIBLE:   16,74 €',
    'IVA 21%:           3,52 €',
    'TOTAL:            20,26 €',
    '',
    'FORMA DE PAGO: TARJETA'
  ].join('\n');

  var datos     = _parsearTextoTicket_(textoDemo);
  var confianza = _calcularConfianzaOCR_(datos, true);
  var estado    = _determinarEstadoOCR_(datos, true);

  Logger.log('  Proveedor:   ' + datos.proveedor);
  Logger.log('  CIF/NIF:     ' + datos.cifNif);
  Logger.log('  Nº Ticket:   ' + datos.numFactura);
  Logger.log('  Fecha:       ' + datos.fecha);
  Logger.log('  Base:        ' + datos.base);
  Logger.log('  IVA %:       ' + datos.ivaPorcentaje);
  Logger.log('  IVA Importe: ' + datos.ivaImporte);
  Logger.log('  Total:       ' + datos.total);
  Logger.log('  Método pago: ' + datos.metodoPago);
  Logger.log('  Confianza:   ' + confianza);
  Logger.log('  Estado:      ' + estado);

  // ── 2. OCR sobre gastos reales pendientes ─────────────────────
  Logger.log('\n[2] procesarGastosPendientesOCR():');
  procesarGastosPendientesOCR();

  Logger.log('\n================================================');
  Logger.log('FIN testGastosFase3()');
  Logger.log('================================================');
}

// ── Privadas FASE 3 ───────────────────────────────────────────

/**
 * Copia el archivo de Drive como Google Doc (activa OCR automáticamente),
 * lee el texto resultante y elimina el doc temporal.
 * Devuelve { texto, exito, error? }.
 */
function _extraerTextoOCR_(archivoId) {
  if (!archivoId) return { texto: '', exito: false, error: 'Sin archivoId' };

  var token     = ScriptApp.getOAuthToken();
  var tempDocId = null;

  try {
    var resp = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + archivoId + '/copy',
      {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type':  'application/json'
        },
        payload: JSON.stringify({
          name:     '_OCR_TEMP_' + archivoId,
          mimeType: 'application/vnd.google-apps.document'
        }),
        muteHttpExceptions: true
      }
    );

    if (resp.getResponseCode() !== 200) {
      return { texto: '', exito: false,
               error: 'Drive API ' + resp.getResponseCode() + ': ' +
                      resp.getContentText().substring(0, 200) };
    }

    tempDocId = JSON.parse(resp.getContentText()).id;
    if (!tempDocId) return { texto: '', exito: false, error: 'Sin ID en respuesta Drive' };

    // Esperar hasta 10 s a que el doc esté listo
    var texto = '';
    for (var i = 0; i < 5; i++) {
      Utilities.sleep(2000);
      try {
        texto = DocumentApp.openById(tempDocId).getBody().getText().trim();
        if (texto.length > 5) break;
      } catch (_) {}
    }

    return { texto: texto, exito: texto.length > 10 };

  } catch (e) {
    return { texto: '', exito: false, error: e.message };
  } finally {
    if (tempDocId) {
      try { DriveApp.getFileById(tempDocId).setTrashed(true); } catch (_) {}
    }
  }
}

/**
 * Parsea texto OCR de un ticket o factura en español.
 * Devuelve un objeto con los campos extraídos (vacíos si no se detectan).
 */
function _parsearTextoTicket_(texto) {
  var datos = {
    fecha: '', proveedor: '', cifNif: '', numFactura: '', concepto: '',
    base: '', ivaPorcentaje: '', ivaImporte: '', total: '', metodoPago: ''
  };
  if (!texto || texto.trim().length < 5) return datos;

  var t      = texto.toUpperCase();
  var lineas = texto.split('\n').map(function(l) { return l.trim(); })
                                .filter(function(l) { return l.length > 0; });

  // Fecha
  var mFecha = texto.match(/\b(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})\b/);
  if (mFecha) {
    var yy = mFecha[3].length === 2 ? '20' + mFecha[3] : mFecha[3];
    datos.fecha = mFecha[1].padStart(2,'0') + '/' + mFecha[2].padStart(2,'0') + '/' + yy;
  }

  // Total
  var mTotal = t.match(/TOTAL\s*A\s*PAGAR[:\s]*[€$]?\s*(\d+[.,]\d{1,2})/) ||
               t.match(/IMPORTE\s+TOTAL[:\s]*[€$]?\s*(\d+[.,]\d{1,2})/)   ||
               t.match(/A\s+PAGAR[:\s]*[€$]?\s*(\d+[.,]\d{1,2})/)         ||
               t.match(/TOTAL[:\s]*[€$]?\s*(\d+[.,]\d{1,2})/);
  if (mTotal) datos.total = _normNum_(mTotal[1]);

  // Base imponible
  var mBase = t.match(/BASE\s*IMPONIBLE[:\s]*[€$]?\s*(\d+[.,]\d{1,2})/) ||
              t.match(/BASE[:\s]*[€$]?\s*(\d+[.,]\d{1,2})/);
  if (mBase) datos.base = _normNum_(mBase[1]);

  // IVA porcentaje
  var mIvaPct = t.match(/IVA\s*(\d{1,2})\s*%/);
  if (mIvaPct) datos.ivaPorcentaje = parseInt(mIvaPct[1], 10);

  // IVA importe
  var mIvaImp = t.match(/CUOTA\s+IVA[:\s]*[€$]?\s*(\d+[.,]\d{1,2})/)     ||
                t.match(/IMPORTE\s+IVA[:\s]*[€$]?\s*(\d+[.,]\d{1,2})/)    ||
                t.match(/IVA\s*\d*\s*%[:\s]*[€$]?\s*(\d+[.,]\d{1,2})/)    ||
                t.match(/IVA[:\s]+[€$]?\s*(\d+[.,]\d{1,2})/);
  if (mIvaImp) datos.ivaImporte = _normNum_(mIvaImp[1]);

  // CIF / NIF
  var mCif = texto.match(/\b([A-Z]\d{7}[A-Z0-9]|\d{8}[A-Z]|[XYZ]\d{7}[A-Z])\b/);
  if (mCif) datos.cifNif = mCif[1];

  // Número de factura / ticket
  var mNum = t.match(/(?:FACTURA|TICKET|ALBAR[AÁ]N|N[ÚU]M(?:ERO)?\.?)\s*[:\s#]*([A-Z0-9][A-Z0-9\/\-]{1,20})/);
  if (mNum) datos.numFactura = mNum[1].trim();

  // Método de pago
  if      (/\bEFECTIVO\b/.test(t))     datos.metodoPago = 'Efectivo';
  else if (/\bBIZUM\b/.test(t))        datos.metodoPago = 'Bizum';
  else if (/\bVISA\b/.test(t))         datos.metodoPago = 'Tarjeta Visa';
  else if (/MASTERCARD/.test(t))       datos.metodoPago = 'Tarjeta Mastercard';
  else if (/\bTARJETA\b/.test(t))      datos.metodoPago = 'Tarjeta';
  else if (/TRANSFERENCIA/.test(t))    datos.metodoPago = 'Transferencia';

  // Proveedor (primera línea con texto relevante, no empieza por dígito ni palabra clave)
  for (var p = 0; p < Math.min(lineas.length, 6); p++) {
    var l = lineas[p];
    if (l.length >= 3 && !/^\d/.test(l) &&
        !/^(fecha|date|hora|time|ticket|factura|nif|cif|www\.|tel)/i.test(l)) {
      datos.proveedor = l.substring(0, 60);
      break;
    }
  }

  // Concepto
  var mConc = t.match(/(?:DESCRIPCI[OÓ]N|CONCEPTO|ART[IÍ]CULO)[:\s]+(.{3,80})/);
  if (mConc) datos.concepto = mConc[1].trim().substring(0, 100);

  return datos;
}

/**
 * Puntuación 0-1 basada en los campos extraídos.
 * El total es el campo más importante (0.35).
 */
function _calcularConfianzaOCR_(datos, ocrExito) {
  if (!ocrExito) return 0.0;
  var score = 0;
  if (datos.total       !== '') score += 0.35;
  if (datos.base        !== '') score += 0.20;
  if (datos.ivaImporte  !== '') score += 0.15;
  if (datos.fecha             ) score += 0.10;
  if (datos.proveedor         ) score += 0.10;
  if (datos.cifNif            ) score += 0.05;
  if (datos.numFactura        ) score += 0.05;
  return Math.round(score * 100) / 100;
}

/**
 * Determina Estado_Procesamiento según los datos extraídos.
 * Reglas (spec FASE 3):
 *   - total + (proveedor o fecha) + IVA → "Procesado"
 *   - total + (proveedor o fecha) sin IVA → "Revisar"
 *   - poco texto o sin total → "OCR fallido"
 */
function _determinarEstadoOCR_(datos, ocrExito) {
  if (!ocrExito) return 'OCR fallido';
  var tieneContexto = datos.proveedor || datos.fecha;
  if (datos.total !== '' && tieneContexto) {
    return datos.ivaImporte !== '' ? 'Procesado' : 'Revisar';
  }
  return 'OCR fallido';
}

/**
 * Escribe los datos OCR en la fila de GASTOS.
 * Solo actualiza campos no vacíos (nunca borra datos existentes).
 */
function _actualizarGastoConOCR_(sheet, rowNum, datos, textoOCR, confianza, estado) {
  var up = {};
  up[COL_GAS.OCR_TEXTO]            = textoOCR;
  up[COL_GAS.CONFIANZA_OCR]        = confianza;
  up[COL_GAS.ESTADO_PROCESAMIENTO] = estado;

  if (datos.fecha)            up[COL_GAS.FECHA_GASTO]           = datos.fecha;
  if (datos.proveedor)        up[COL_GAS.PROVEEDOR]              = datos.proveedor;
  if (datos.cifNif)           up[COL_GAS.CIF_NIF_PROVEEDOR]      = datos.cifNif;
  if (datos.numFactura)       up[COL_GAS.NUMERO_FACTURA_TICKET]  = datos.numFactura;
  if (datos.concepto)         up[COL_GAS.CONCEPTO]               = datos.concepto;
  if (datos.base        !== '') up[COL_GAS.BASE_IMPONIBLE]       = datos.base;
  if (datos.ivaPorcentaje !== '') up[COL_GAS.IVA_PORCENTAJE]     = datos.ivaPorcentaje;
  if (datos.ivaImporte  !== '') up[COL_GAS.IVA_IMPORTE]          = datos.ivaImporte;
  if (datos.total       !== '') up[COL_GAS.TOTAL]                = datos.total;
  if (datos.metodoPago)       up[COL_GAS.METODO_PAGO]            = datos.metodoPago;

  for (var col in up) {
    sheet.getRange(rowNum, parseInt(col, 10)).setValue(up[col]);
  }
}

/**
 * Normaliza un string numérico con coma decimal a float.
 */
function _normNum_(str) {
  var v = parseFloat(String(str || '').replace(',', '.'));
  return isNaN(v) ? '' : v;
}
