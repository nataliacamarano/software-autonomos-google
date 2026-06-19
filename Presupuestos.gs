// ================================================================
// Presupuestos.gs — Gestión de presupuestos PDF con líneas libres
// ================================================================

var HEADERS_PRESUPUESTOS = [
  'ID_Presupuesto', 'Fecha', 'Cliente', 'Num_Lineas',
  'Subtotal', 'IVA_Porcentaje', 'IVA_Importe', 'Total',
  'Estado', 'URL_PDF', 'Lineas_JSON',
  'Fecha_Emision', 'Fecha_Caducidad'
];

// Columnas PRESUPUESTOS base-1
var COL_PRE = {
  ID:              1,
  FECHA:           2,
  CLIENTE:         3,
  NUM_LINEAS:      4,
  SUBTOTAL:        5,
  IVA_PCT:         6,
  IVA_IMPORTE:     7,
  TOTAL:           8,
  ESTADO:          9,
  URL_PDF:         10,
  LINEAS_JSON:     11,
  FECHA_EMISION:   12,
  FECHA_CADUCIDAD: 13
};

var ESTADOS_PRESUPUESTO = ['Pendiente', 'Aceptado', 'Rechazado', 'Facturado'];

// ================================================================
// INSTALACIÓN
// ================================================================

function installPresupuestosModule() {
  Logger.log('================================================');
  Logger.log('installPresupuestosModule()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var ss = getTargetSS_();

  Logger.log('\n[1] Hoja PRESUPUESTOS:');
  var sh = ss.getSheetByName('PRESUPUESTOS');
  if (!sh) {
    sh = ss.insertSheet('PRESUPUESTOS');
    Logger.log('  Creada.');
  } else {
    Logger.log('  Ya existe. Verificando cabeceras...');
  }

  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS_PRESUPUESTOS);
    Logger.log('  Cabeceras escritas: ' + HEADERS_PRESUPUESTOS.join(' | '));
  } else {
    var existing = sh.getRange(1, 1, 1, Math.min(sh.getLastColumn(), HEADERS_PRESUPUESTOS.length)).getValues()[0];
    Logger.log('  Cabeceras actuales: ' + existing.join(' | '));
  }

  applyHeaderStyle_(sh, HEADERS_PRESUPUESTOS.length);

  sh.setColumnWidth(COL_PRE.ID,              130);
  sh.setColumnWidth(COL_PRE.FECHA,           100);
  sh.setColumnWidth(COL_PRE.CLIENTE,         200);
  sh.setColumnWidth(COL_PRE.NUM_LINEAS,       80);
  sh.setColumnWidth(COL_PRE.SUBTOTAL,        100);
  sh.setColumnWidth(COL_PRE.IVA_PCT,          70);
  sh.setColumnWidth(COL_PRE.IVA_IMPORTE,     100);
  sh.setColumnWidth(COL_PRE.TOTAL,           100);
  sh.setColumnWidth(COL_PRE.ESTADO,          110);
  sh.setColumnWidth(COL_PRE.URL_PDF,         280);
  sh.setColumnWidth(COL_PRE.LINEAS_JSON,     400);
  sh.setColumnWidth(COL_PRE.FECHA_EMISION,   110);
  sh.setColumnWidth(COL_PRE.FECHA_CADUCIDAD, 110);

  Logger.log('\n[2] Estados: ' + ESTADOS_PRESUPUESTO.join(' | '));
  _applyEstadoValidation_(sh);
  Logger.log('  Validación de Estados aplicada a columna ' + COL_PRE.ESTADO + '.');

  var siguiente = _nextPresupuestoNumber_();
  Logger.log('\n[3] Próximo ID: ' + siguiente.id);
  Logger.log('\n================================================');
  Logger.log('MÓDULO PRESUPUESTOS INSTALADO.');
  Logger.log('================================================');
}

// ================================================================
// GENERACIÓN DE PDF
// ================================================================

/**
 * Genera el PDF de un presupuesto con líneas libres.
 *
 * @param {Object} params
 *   .cliente      {string}
 *   .lineItems    {Array<{desc,cant,pu,imp}>}
 *   .conIva       {boolean}
 *   .fechaEmision {Date}   — fecha de emisión introducida por el usuario
 *   .validezDias  {number} — 7 | 15 | 30 | 60 | 90
 *
 * @returns {{ id: string, url: string }}
 */
function generatePresupuestoPDF(params) {
  var cliente      = String(params.cliente   || '').trim();
  var lineItems    = params.lineItems  || [];
  var conIva       = !!params.conIva;
  var fechaEmision = (params.fechaEmision instanceof Date && !isNaN(params.fechaEmision.getTime()))
                       ? params.fechaEmision : new Date();
  var validezDias  = Number(params.validezDias) || 30;

  Logger.log('================================================');
  Logger.log('generatePresupuestoPDF()');
  Logger.log('  Cliente:      ' + cliente);
  Logger.log('  Líneas:       ' + lineItems.length);
  Logger.log('  IVA:          ' + (conIva ? 'Sí' : 'No'));
  Logger.log('  Fecha emisión:' + Utilities.formatDate(fechaEmision, Session.getScriptTimeZone(), 'dd/MM/yyyy'));
  Logger.log('  Validez:      ' + validezDias + ' días');
  Logger.log('================================================');

  // ── Datos empresa ─────────────────────────────────────────────
  var empresa      = getConfig('empresa_nombre')    || 'Mi Empresa';
  var empresaCif   = getConfig('empresa_cif')       || '';
  var empresaDir   = getConfig('empresa_direccion') || '';
  var empresaEmail = getConfig('empresa_email')     || '';
  var empresaTel   = getConfig('empresa_telefono')  || '';
  var ivaPct       = parseFloat(getConfig('iva_porcentaje') || '21');

  // ── Datos cliente ─────────────────────────────────────────────
  var cliData   = _getClienteByLabel_(cliente);
  var cliNombre = cliData ? cliData.nombre    : cliente;
  var cliCif    = cliData ? cliData.cif       : '';
  var cliDir    = cliData ? cliData.direccion : '';
  var cliEmail  = cliData ? cliData.email     : '';

  // ── Fechas ────────────────────────────────────────────────────
  var tz           = Session.getScriptTimeZone();
  var fechaCaduc   = new Date(fechaEmision.getTime());
  fechaCaduc.setDate(fechaCaduc.getDate() + validezDias);
  var fechaEmiStr  = Utilities.formatDate(fechaEmision, tz, 'dd/MM/yyyy');
  var fechaCaducStr= Utilities.formatDate(fechaCaduc,   tz, 'dd/MM/yyyy');
  var hoy          = new Date();

  // ── Cálculos ──────────────────────────────────────────────────
  var subtotal = 0;
  for (var i = 0; i < lineItems.length; i++) subtotal += lineItems[i].imp || 0;
  subtotal       = Math.round(subtotal * 100) / 100;
  var ivaImporte = conIva ? Math.round(subtotal * ivaPct) / 100 : 0;
  var total      = subtotal + ivaImporte;

  // ── Numeración ────────────────────────────────────────────────
  var numPre = _nextPresupuestoNumber_();
  Logger.log('  Número: ' + numPre.id);

  // ── Crear documento Google Docs ───────────────────────────────
  var docTitulo = numPre.id + '_' + _sanitizeFilename_(cliente);
  var doc       = DocumentApp.create(docTitulo);
  var body      = doc.getBody();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);

  // Cabecera empresa (logo izquierda + info derecha)
  _cabecera_empresa_(body, empresa, empresaCif, empresaTel, empresaEmail, empresaDir);

  // Número y fechas
  body.appendParagraph('PRESUPUESTO   ' + numPre.id)
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph('Fecha de emisión: ' + fechaEmiStr)
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
    .editAsText().setFontSize(10);
  body.appendParagraph('Válido hasta: ' + fechaCaducStr)
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
    .editAsText().setFontSize(10).setBold(true);

  body.appendParagraph('');

  // Datos del cliente
  body.appendParagraph('CLIENTE').editAsText().setFontSize(9).setBold(true).setForegroundColor('#555555');
  body.appendParagraph(cliNombre).editAsText().setFontSize(11).setBold(true);
  if (cliCif)   body.appendParagraph('CIF/NIF: ' + cliCif).editAsText().setFontSize(10);
  if (cliDir)   body.appendParagraph(cliDir).editAsText().setFontSize(10);
  if (cliEmail) body.appendParagraph(cliEmail).editAsText().setFontSize(10);

  body.appendHorizontalRule();

  // ── Tabla de conceptos ────────────────────────────────────────
  body.appendParagraph('CONCEPTOS').editAsText().setFontSize(11).setBold(true);
  body.appendParagraph('');

  var tableData = [['Descripción', 'Cant.', 'P.U. €', 'Importe']];
  for (var j = 0; j < lineItems.length; j++) {
    var li = lineItems[j];
    tableData.push([
      li.desc || '',
      li.cant > 0 ? String(li.cant) : '',
      li.pu   > 0 ? formatCurrency(li.pu) : '',
      li.imp  > 0 ? formatCurrency(li.imp) : '0,00 €'
    ]);
  }
  tableData.push(['', '', 'Subtotal', formatCurrency(subtotal)]);
  if (conIva) {
    tableData.push(['', '', 'IVA ' + ivaPct + '%', formatCurrency(ivaImporte)]);
  }
  tableData.push(['', '', 'TOTAL', formatCurrency(total)]);

  var tbl = body.appendTable(tableData);

  // Estilo cabecera de tabla
  var hdrRow = tbl.getRow(0);
  for (var c = 0; c < 4; c++) {
    hdrRow.getCell(c)
      .setBackgroundColor('#34a853')
      .editAsText().setBold(true).setFontSize(10).setForegroundColor('#ffffff');
  }

  // Estilo filas de datos
  for (var k = 1; k <= lineItems.length; k++) {
    var dRow = tbl.getRow(k);
    dRow.getCell(0).editAsText().setFontSize(10);
    dRow.getCell(1).editAsText().setFontSize(10);
    dRow.getCell(2).editAsText().setFontSize(10);
    dRow.getCell(3).editAsText().setFontSize(10);
  }

  // Estilo filas de totales
  var totalRowStart = 1 + lineItems.length;
  for (var tr = totalRowStart; tr < tbl.getNumRows(); tr++) {
    var totRow = tbl.getRow(tr);
    var isLast = (tr === tbl.getNumRows() - 1);
    for (var tc = 0; tc < 4; tc++) {
      var cell = totRow.getCell(tc);
      if (isLast) {
        cell.setBackgroundColor('#34a853');
        cell.editAsText().setBold(true).setFontSize(11).setForegroundColor('#ffffff');
      } else {
        cell.editAsText().setBold(true).setFontSize(10).setForegroundColor('#333333');
      }
    }
  }

  // Nota de validez
  body.appendParagraph('');
  body.appendParagraph(
    'Este presupuesto tiene validez hasta la fecha indicada.'
  ).editAsText().setFontSize(9).setForegroundColor('#555555').setItalic(true);

  // Pie
  body.appendParagraph('');
  body.appendHorizontalRule();
  body.appendParagraph(
    'Documento generado el ' +
    Utilities.formatDate(hoy, tz, 'dd/MM/yyyy HH:mm') +
    '   ·   ' + empresa
  ).editAsText().setFontSize(8).setForegroundColor('#888888');

  doc.saveAndClose();

  // ── Convertir a PDF ───────────────────────────────────────────
  var docFile = DriveApp.getFileById(doc.getId());
  var pdfBlob = docFile.getAs('application/pdf');
  pdfBlob.setName(docTitulo + '.pdf');

  // ── Guardar en Drive ──────────────────────────────────────────
  var _prCarpetaId_ = getConfig('carpeta_presupuestos_id') || getConfig('carpeta_facturas_id');
  var carpeta;
  if (!_prCarpetaId_) {
    Logger.log('  ⚠️ carpeta_presupuestos_id no configurado. Guardando en raíz.');
    carpeta = DriveApp.getRootFolder();
  } else {
    try {
      carpeta = _getOrCreateSubcarpetaAnioMes_(_prCarpetaId_, new Date());
    } catch (e) {
      Logger.log('  ⚠️ Error subcarpeta AÑO/MES: ' + e.message + '. Guardando en raíz.');
      carpeta = DriveApp.getRootFolder();
    }
  }
  var pdfFile = carpeta.createFile(pdfBlob);
  var pdfUrl  = pdfFile.getUrl();
  docFile.setTrashed(true);
  Logger.log('  PDF: ' + pdfUrl);

  // ── Registrar en PRESUPUESTOS ─────────────────────────────────
  _ensurePresupuestosSheet_();
  var preSheet = getSheet('PRESUPUESTOS');
  _ensureLineasJsonCol_(preSheet);   // garantiza que la columna existe antes de mapear
  var cm = _buildPreColMap_(preSheet);
  var values = {
    'ID_Presupuesto': numPre.id,
    'Fecha':          hoy,
    'Cliente':        cliNombre,
    'Num_Lineas':     lineItems.length,
    'Subtotal':       subtotal,
    'IVA_Porcentaje': conIva ? ivaPct : 0,
    'IVA_Importe':    ivaImporte,
    'Total':          total,
    'Estado':         'Pendiente',
    'URL_PDF':        pdfUrl,
    'Lineas_JSON':    JSON.stringify(lineItems),
    'Fecha_Emision':  fechaEmision,
    'Fecha_Caducidad':fechaCaduc
  };
  var maxCol = 0;
  for (var k in cm) { if (cm[k] > maxCol) maxCol = cm[k]; }
  var rowArr = new Array(maxCol).fill('');
  for (var key in values) {
    if (cm[key]) rowArr[cm[key] - 1] = values[key];
  }
  preSheet.appendRow(rowArr);

  logEvent('INFO', 'generatePresupuestoPDF',
    numPre.id + ' | ' + cliNombre + ' | ' + lineItems.length + ' líneas | Total: ' + formatCurrency(total) +
    ' | Válido hasta: ' + fechaCaducStr,
    pdfUrl);

  return { id: numPre.id, url: pdfUrl };
}

// ================================================================
// GESTIÓN DE ESTADOS
// ================================================================

function updatePresupuestoEstado(id, nuevoEstado) {
  if (ESTADOS_PRESUPUESTO.indexOf(nuevoEstado) === -1) {
    return { success: false, error: 'Estado inválido: ' + nuevoEstado };
  }
  try {
    var sh        = getSheet('PRESUPUESTOS');
    var cm        = _buildPreColMap_(sh);
    var colId     = cm['ID_Presupuesto'] || COL_PRE.ID;
    var colEstado = cm['Estado']         || COL_PRE.ESTADO;
    var lastRow   = sh.getLastRow();
    if (lastRow < 2) return { success: false, error: 'Hoja PRESUPUESTOS vacía' };
    var ids = sh.getRange(2, colId, lastRow - 1, 1).getValues().flat();
    var idx = -1;
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i]) === String(id)) { idx = i; break; }
    }
    if (idx === -1) return { success: false, error: 'ID no encontrado: ' + id };
    sh.getRange(idx + 2, colEstado).setValue(nuevoEstado);
    logEvent('INFO', 'updatePresupuestoEstado', id + ' → ' + nuevoEstado, '');
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ================================================================
// CONVERSIÓN A FACTURA (preparación de datos)
// ================================================================

/**
 * Devuelve los datos de un presupuesto listos para convertir a factura.
 * lineItems contiene el array deserializado de Lineas_JSON.
 *
 * @param {string} idPresupuesto  Ej: 'PRE-2026-001'
 * @returns {Object|null}
 */
function getPresupuestoParaFacturar(idPresupuesto) {
  try {
    var sh      = getSheet('PRESUPUESTOS');
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return null;

    // Mapa de columnas real (tolerante a diferencias de estructura)
    var cm  = _buildPreColMap_(sh);
    var ncols = sh.getLastColumn();
    var data  = sh.getRange(2, 1, lastRow - 1, ncols).getValues();

    // Helper: obtiene valor de fila por nombre de cabecera, con fallback
    function gv(row, name, def) {
      var c = cm[name];
      return (c !== undefined) ? row[c - 1] : (def !== undefined ? def : '');
    }

    var colId = cm['ID_Presupuesto'] || COL_PRE.ID;

    for (var i = 0; i < data.length; i++) {
      if (String(data[i][colId - 1]) !== String(idPresupuesto)) continue;
      var row       = data[i];
      var lineItems = [];
      var jsonStr   = String(gv(row, 'Lineas_JSON', '[]'));
      try { lineItems = JSON.parse(jsonStr); } catch (_) {}

      // Fallback: si la columna Lineas_JSON tiene datos desalineados (URL en lugar de JSON),
      // escanear toda la fila buscando una celda que sea un array JSON válido y no vacío.
      if (!Array.isArray(lineItems) || !lineItems.length) {
        for (var fc = 0; fc < row.length; fc++) {
          var fcVal = String(row[fc] || '').trim();
          if (fcVal.charAt(0) !== '[') continue;
          try {
            var fcParsed = JSON.parse(fcVal);
            if (Array.isArray(fcParsed) && fcParsed.length > 0 &&
                typeof fcParsed[0] === 'object' && fcParsed[0] !== null) {
              lineItems = fcParsed;
              Logger.log('getPresupuestoParaFacturar: lineItems recuperado de columna ' + (fc + 1) + ' (fallback desalineamiento).');
              break;
            }
          } catch (_) {}
        }
      }
      return {
        idPresupuesto:  String(gv(row, 'ID_Presupuesto')),
        fecha:          gv(row, 'Fecha'),
        cliente:        String(gv(row, 'Cliente')),
        numLineas:      Number(gv(row, 'Num_Lineas',      0)) || 0,
        subtotal:       Number(gv(row, 'Subtotal',        0)) || 0,
        ivaPct:         Number(gv(row, 'IVA_Porcentaje',  0)) || 0,
        ivaImporte:     Number(gv(row, 'IVA_Importe',     0)) || 0,
        total:          Number(gv(row, 'Total',           0)) || 0,
        estado:         String(gv(row, 'Estado')),
        urlPdf:         String(gv(row, 'URL_PDF')),
        lineItems:      lineItems,
        fechaEmision:   gv(row, 'Fecha_Emision'),
        fechaCaducidad: gv(row, 'Fecha_Caducidad')
      };
    }
  } catch (e) {
    Logger.log('getPresupuestoParaFacturar: ' + e.message);
  }
  return null;
}

// ================================================================
// HELPERS PRIVADOS
// ================================================================

/**
 * Añade la columna 'Lineas_JSON' al final de PRESUPUESTOS si no existe.
 * Seguro de llamar varias veces — idempotente.
 * @param {Sheet} sh  hoja PRESUPUESTOS ya obtenida
 */
function _ensureLineasJsonCol_(sh) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return;
  var hdrs = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var h = 0; h < hdrs.length; h++) {
    if (String(hdrs[h]).trim() === 'Lineas_JSON') return; // ya existe
  }
  var newCol = lastCol + 1;
  var hdrCell = sh.getRange(1, newCol);
  hdrCell.setValue('Lineas_JSON');
  hdrCell.setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  sh.setColumnWidth(newCol, 400);
  Logger.log('_ensureLineasJsonCol_: columna Lineas_JSON añadida en col ' + newCol);
}

/**
 * Wrapper público para llamar desde el menú.
 * Añade Lineas_JSON a PRESUPUESTOS y muestra confirmación.
 */
function menuEnsureLineasJson() {
  var sh = getSheet('PRESUPUESTOS');
  var lastCol = sh.getLastColumn();
  var hdrs = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var h = 0; h < hdrs.length; h++) {
    if (String(hdrs[h]).trim() === 'Lineas_JSON') {
      SpreadsheetApp.getUi().alert('✅ La columna Lineas_JSON ya existe en col ' + (h + 1) + '. No es necesario hacer nada.');
      return;
    }
  }
  _ensureLineasJsonCol_(sh);
  SpreadsheetApp.getUi().alert(
    '✅  Columna Lineas_JSON añadida (col ' + (sh.getLastColumn()) + ').\n\n' +
    'Los presupuestos NUEVOS ya guardarán el detalle completo de líneas.\n' +
    'Los presupuestos existentes no se modifican — crea uno nuevo para probar.'
  );
}

/**
 * Lee la fila 1 de la hoja y devuelve un mapa { NombreCabecera: índice_base1 }.
 * Permite trabajar con la estructura real de columnas en lugar de constantes fijas.
 */
function _buildPreColMap_(sh) {
  var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var map  = {};
  for (var i = 0; i < hdrs.length; i++) {
    var h = String(hdrs[i]).trim();
    if (h) map[h] = i + 1;  // índice base-1
  }
  return map;
}

function _nextPresupuestoNumber_() {
  var year   = String(new Date().getFullYear());
  var prefix = 'PRE-' + year + '-';
  var maxNum = 0;

  try {
    var sh      = getSheet('PRESUPUESTOS');
    var lastRow = sh.getLastRow();
    if (lastRow > 1) {
      var ids = sh.getRange(2, COL_PRE.ID, lastRow - 1, 1).getValues().flat();
      for (var i = 0; i < ids.length; i++) {
        var idStr = String(ids[i]);
        if (idStr.indexOf(prefix) === 0) {
          var n = parseInt(idStr.split('-').pop(), 10);
          if (!isNaN(n) && n > maxNum) maxNum = n;
        }
      }
    }
  } catch (_) {}

  var numero = maxNum + 1;
  return { id: prefix + String(numero).padStart(3, '0'), numero: numero };
}

function _ensurePresupuestosSheet_() {
  var ss = getTargetSS_();
  var sh = ss.getSheetByName('PRESUPUESTOS');
  if (!sh) {
    sh = ss.insertSheet('PRESUPUESTOS');
    sh.appendRow(HEADERS_PRESUPUESTOS);
    applyHeaderStyle_(sh, HEADERS_PRESUPUESTOS.length);
    Logger.log('  PRESUPUESTOS creada automáticamente.');
  }
  _applyEstadoValidation_(sh);
  return sh;
}

/**
 * Aplica validación de lista a la columna Estado real de la hoja (excepto cabecera).
 * Detecta la columna por nombre de cabecera, no por índice fijo.
 * Idempotente — se puede llamar varias veces sin efecto secundario.
 */
function _applyEstadoValidation_(sh) {
  var cm        = _buildPreColMap_(sh);
  var colEstado = cm['Estado'];
  if (!colEstado) {
    Logger.log('  ⚠️  Columna "Estado" no encontrada en ' + sh.getName() + '. Cabeceras: ' + JSON.stringify(Object.keys(cm)));
    return;
  }
  Logger.log('  Estado validation → col ' + colEstado + ' de hoja "' + sh.getName() + '"');
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(ESTADOS_PRESUPUESTO, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, colEstado, sh.getMaxRows() - 1, 1).setDataValidation(rule);
  Logger.log('  Estado validation aplicada OK.');
}

// ================================================================
// REPARACIÓN DE FILAS DESALINEADAS
// ================================================================

/**
 * Detecta y corrige filas donde los datos fueron escritos por índice fijo
 * (orden de HEADERS_PRESUPUESTOS) en lugar de por columna real de la hoja.
 * Caso típico: la hoja tenía Estado en col 12 pero appendRow lo escribía
 * en la posición 9 (Estado en el array), poniendo una fecha en col 12.
 *
 * Seguro de ejecutar varias veces — omite filas ya correctas.
 * NO toca otras hojas (facturas, trabajos, informes, proformas).
 *
 * @returns {{ fixed: number, skipped: number, errors: string[] }}
 */
function repairPresupuestosRows() {
  Logger.log('================================================');
  Logger.log('repairPresupuestosRows()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var sh = getSheet('PRESUPUESTOS');
  var lastRow = sh.getLastRow();
  if (lastRow < 2) {
    Logger.log('  Sin filas de datos. Nada que reparar.');
    return { fixed: 0, skipped: 0, errors: [] };
  }

  var cm = _buildPreColMap_(sh);
  var colEstado = cm['Estado'];
  if (!colEstado) {
    var msg = 'Columna "Estado" no encontrada en cabeceras: ' + JSON.stringify(Object.keys(cm));
    Logger.log('  ERROR: ' + msg);
    return { fixed: 0, skipped: 0, errors: [msg] };
  }

  Logger.log('  Mapa de columnas reales: ' + JSON.stringify(cm));

  var ncols  = sh.getLastColumn();
  var data   = sh.getRange(2, 1, lastRow - 1, ncols).getValues();
  var fixed  = 0, skipped = 0;
  var errors = [];

  for (var r = 0; r < data.length; r++) {
    var row    = data[r];
    var rowNum = r + 2;

    // Fila vacía (sin ID)
    if (!String(row[0] || '').trim()) { skipped++; continue; }

    var estadoActual = String(row[colEstado - 1] || '').trim();

    if (ESTADOS_PRESUPUESTO.indexOf(estadoActual) !== -1) {
      Logger.log('  fila ' + rowNum + ': Estado="' + estadoActual + '" — OK, sin cambios.');
      skipped++;
      continue;
    }

    // Estado inválido: reconstruir la fila usando HEADERS_PRESUPUESTOS como
    // clave de posición (appendRow escribió los datos en ese orden exacto).
    Logger.log('  fila ' + rowNum + ': Estado="' + estadoActual + '" — DESALINEADA. Reparando…');

    try {
      // Extraer cada valor por su posición en HEADERS_PRESUPUESTOS
      var extracted = {};
      for (var h = 0; h < HEADERS_PRESUPUESTOS.length; h++) {
        extracted[HEADERS_PRESUPUESTOS[h]] = (h < row.length) ? row[h] : '';
      }
      Logger.log('  Valores extraídos: Estado="' + extracted['Estado'] + '" URL_PDF="' + extracted['URL_PDF'] + '"');

      // Si Lineas_JSON no es un array válido, buscar JSON en cualquier columna de la fila
      var jsonCandidate = String(extracted['Lineas_JSON'] || '').trim();
      var itemsArr = [];
      try { itemsArr = JSON.parse(jsonCandidate); } catch (_) {}
      if (!Array.isArray(itemsArr) || !itemsArr.length) {
        for (var fc = 0; fc < row.length; fc++) {
          var fcv = String(row[fc] || '').trim();
          if (fcv.charAt(0) !== '[') continue;
          try {
            var fcp = JSON.parse(fcv);
            if (Array.isArray(fcp) && fcp.length > 0 && typeof fcp[0] === 'object' && fcp[0] !== null) {
              itemsArr = fcp;
              Logger.log('  Lineas_JSON recuperado de columna ' + (fc + 1) + ' (fallback).');
              break;
            }
          } catch (_) {}
        }
      }
      // Último recurso: reconstruir línea única desde Subtotal
      if (!Array.isArray(itemsArr) || !itemsArr.length) {
        var sub = Number(extracted['Subtotal'] || extracted['Total'] || 0);
        if (sub > 0) {
          itemsArr = [{ desc: 'Trabajos según presupuesto', cant: 1, pu: sub, imp: sub }];
          Logger.log('  Lineas_JSON reconstruido desde Subtotal=' + sub + ' (datos originales perdidos).');
        }
      }
      extracted['Lineas_JSON'] = JSON.stringify(itemsArr);

      // Construir array de fila con columnas en el orden real de la hoja
      var maxCol = 0;
      for (var colName in cm) { if (cm[colName] > maxCol) maxCol = cm[colName]; }
      var correctedRow = sh.getRange(rowNum, 1, 1, maxCol).getValues()[0]; // copia actual

      for (var fieldName in extracted) {
        var realCol = cm[fieldName];
        if (realCol) correctedRow[realCol - 1] = extracted[fieldName];
      }

      // Garantizar que Estado es válido
      var estadoFinal = String(correctedRow[colEstado - 1] || '').trim();
      if (ESTADOS_PRESUPUESTO.indexOf(estadoFinal) === -1) {
        correctedRow[colEstado - 1] = 'Pendiente';
        Logger.log('  Forzado Estado="Pendiente" (valor extraído no válido: "' + estadoFinal + '")');
      }

      sh.getRange(rowNum, 1, 1, maxCol).setValues([correctedRow]);
      Logger.log('  fila ' + rowNum + ' reparada → Estado="' + correctedRow[colEstado - 1] + '"');
      fixed++;
    } catch (e) {
      var errMsg = 'fila ' + rowNum + ': ' + e.message;
      Logger.log('  ERROR: ' + errMsg);
      errors.push(errMsg);
    }
  }

  // Re-aplicar validación de Estado tras la reparación
  _applyEstadoValidation_(sh);

  Logger.log('================================================');
  Logger.log('repairPresupuestosRows — reparadas: ' + fixed + '  omitidas: ' + skipped + '  errores: ' + errors.length);
  Logger.log('================================================');
  return { fixed: fixed, skipped: skipped, errors: errors };
}
