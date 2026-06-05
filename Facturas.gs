// ============================================================
// Facturas.gs — Generación de facturas PDF con numeración automática
// ============================================================
//
// Claves de CONFIG requeridas (rellenar antes de ejecutar):
//   cliente_factura        → nombre del cliente (igual que en TRABAJOS)
//   fecha_inicio_factura   → DD/MM/YYYY
//   fecha_fin_factura      → DD/MM/YYYY
//   precio_hora            → precio por hora en EUR (ej: 35)
//
// Orden de ejecución:
//   1. validateInvoiceModule()  — verificar infraestructura y ver claves CONFIG
//   2. generateInvoicePDF()     — generar la factura
// ============================================================

/**
 * Filtra TRABAJOS por cliente y rango de fechas, calcula base imponible,
 * IVA e IRPF, genera un PDF de factura profesional, lo guarda en Drive
 * y registra la factura en la hoja FACTURAS.
 */
function generateInvoicePDF(tipo, selectedJobIds, metodoPagoOverride) {
  tipo = tipo || 'Factura';
  Logger.log('================================================');
  Logger.log('generateInvoicePDF(' + tipo + ')');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  // ── 1. Parámetros de CONFIG ──────────────────────────────────
  var cliente       = getConfig('cliente_factura');
  var fechaIniStr   = getConfig('fecha_inicio_factura');
  var fechaFinStr   = getConfig('fecha_fin_factura');
  var precioHoraStr = getConfig('precio_hora');

  Logger.log('\n[1] Parámetros de CONFIG:');
  Logger.log('  cliente_factura:       ' + (cliente       || '❌ NO CONFIGURADO'));
  Logger.log('  fecha_inicio_factura:  ' + (fechaIniStr   || '❌ NO CONFIGURADO'));
  Logger.log('  fecha_fin_factura:     ' + (fechaFinStr   || '❌ NO CONFIGURADO'));
  Logger.log('  precio_hora:           ' + (precioHoraStr || '❌ NO CONFIGURADO'));

  if (!cliente || !fechaIniStr || !fechaFinStr || !precioHoraStr) {
    Logger.log('\nABORTANDO. Rellena las claves en CONFIG y vuelve a ejecutar.');
    Logger.log('  Ejecuta validateInvoiceModule() para ver el estado de cada clave.');
    return;
  }

  var fechaIni = _parseReportDate_(fechaIniStr);
  var fechaFin = _parseReportDate_(fechaFinStr);
  if (!fechaIni || !fechaFin) {
    Logger.log('ERROR: Formato de fecha incorrecto. Usa DD/MM/YYYY.');
    return;
  }
  fechaFin.setHours(23, 59, 59, 999);

  var precioHora = parseFloat(String(precioHoraStr).replace(',', '.')) || 0;

  // ── 2. Datos de empresa ──────────────────────────────────────
  var empresa      = getConfig('empresa_nombre')    || 'Mi Empresa';
  var empresaCif   = getConfig('empresa_cif')       || '';
  var empresaDir   = getConfig('empresa_direccion') || '';
  var empresaEmail = getConfig('empresa_email')     || '';
  var empresaTel   = getConfig('empresa_telefono')  || '';
  var ivaPct = parseFloat(getConfig('iva_porcentaje') || '21');

  // ── 3. Datos del cliente (lookup en CLIENTES) ────────────────
  Logger.log('\n[2] Buscando datos del cliente en CLIENTES...');
  var cliData   = _getClienteByLabel_(cliente);
  var cliNombre = cliData ? cliData.nombre    : cliente;
  var cliCif    = cliData ? cliData.cif       : '';
  var cliDir    = cliData ? cliData.direccion : '';
  var cliEmail  = cliData ? cliData.email     : '';
  var cliId      = cliData ? cliData.id    : '';
  var metodosPago = cliData ? (cliData.metodosPago || []) : [];

  // ── Selección de método de pago ──────────────────────────────
  var pagoSeleccionado = null;
  if (metodoPagoOverride !== undefined) {
    pagoSeleccionado = metodoPagoOverride;
    Logger.log('  Método de pago (override): ' + (pagoSeleccionado ? pagoSeleccionado.tipo : 'ninguno'));
  } else if (metodosPago.length === 1) {
    pagoSeleccionado = metodosPago[0];
    Logger.log('  Método de pago: ' + pagoSeleccionado.tipo + ' (automático)');
  } else if (metodosPago.length > 1) {
    var opsTxt = metodosPago.map(function(m, i) {
      var extra = m.iban ? ' (' + m.iban.replace(/\s/g,'').substring(0, 10) + '…)' :
                  m.numero ? ' (' + m.numero + ')' : m.email ? ' (' + m.email + ')' : '';
      return (i + 1) + '. ' + m.tipo + extra;
    }).join('\n');
    var selResp = SpreadsheetApp.getUi().prompt(
      'Método de pago para este ' + tipo,
      'El cliente tiene varios métodos. Escribe el número:\n\n' + opsTxt + '\n\n0. No incluir datos de pago',
      SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
    );
    if (selResp.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
      var selIdx = parseInt(selResp.getResponseText()) - 1;
      if (selIdx >= 0 && selIdx < metodosPago.length) pagoSeleccionado = metodosPago[selIdx];
    }
    if (pagoSeleccionado) Logger.log('  Método elegido: ' + pagoSeleccionado.tipo);
  }

  Logger.log('  ' + (cliData
    ? '✓ ' + cliNombre + (cliCif ? ' | CIF: ' + cliCif : '')
    : '⚠️  No encontrado en CLIENTES — se usará el nombre del dropdown.'));

  // ── 4. Filtrar TRABAJOS ──────────────────────────────────────
  Logger.log('\n[3] Filtrando TRABAJOS...');
  var shTra   = getSheet('TRABAJOS');
  var lastRow = shTra.getLastRow();
  var trabajos = [];

  if (lastRow > 1) {
    var numCols = Math.min(shTra.getLastColumn(), COL_TRA.MATERIALES);
    var data = shTra.getRange(2, 1, lastRow - 1, numCols).getValues();
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (String(row[COL_TRA.CLIENTE - 1] || '').trim().toLowerCase() !== cliente.toLowerCase()) continue;
      var ft = _parseTrabajosDate_(row[COL_TRA.FECHA_TRABAJO - 1]);
      if (!ft || ft < fechaIni || ft > fechaFin) continue;
      trabajos.push({
        id:               String(row[COL_TRA.ID               - 1] || ''),
        fechaTrabajo:     ft,
        horaInicio:       row[COL_TRA.HORA_INICIO       - 1],
        horaFin:          row[COL_TRA.HORA_FIN          - 1],
        trabajoRealizado: String(row[COL_TRA.TRABAJO_REALIZADO - 1] || '').trim(),
        incidencias:      String(row[COL_TRA.INCIDENCIAS       - 1] || '').trim(),
        materiales:       Number(row[COL_TRA.MATERIALES - 1] || 0)
      });
    }
  }

  Logger.log('  Trabajos encontrados en el período: ' + trabajos.length);

  // Filtrar por IDs seleccionados si vienen del panel con selección manual
  if (Array.isArray(selectedJobIds) && selectedJobIds.length > 0) {
    var selSet = {};
    for (var si = 0; si < selectedJobIds.length; si++) selSet[String(selectedJobIds[si])] = true;
    trabajos = trabajos.filter(function(t) { return !!selSet[String(t.id)]; });
    Logger.log('  Filtrando por IDs seleccionados (' + selectedJobIds.length + ') → ' + trabajos.length + ' trabajos');
  }

  if (trabajos.length === 0) {
    Logger.log('  Sin trabajos seleccionados para facturar. Factura no generada.');
    return;
  }

  var matsTrabajos = trabajos.reduce(function(sum, t) { return sum + (t.materiales || 0); }, 0);
  var matsAdicStr  = getConfig('materiales_adicionales') || '[]';
  var matsAdicItems;
  try {
    matsAdicItems = JSON.parse(matsAdicStr);
    if (!Array.isArray(matsAdicItems)) matsAdicItems = [];
  } catch (_) { matsAdicItems = []; }
  var otrosMateriales = matsAdicItems.reduce(function(s, m) { return s + (Number(m.imp) || 0); }, 0);
  var materiales      = matsTrabajos + otrosMateriales;

  // ── 5. Calcular importes ─────────────────────────────────────
  Logger.log('\n[4] Calculando importes...');
  var totalHoras = 0;
  for (var j = 0; j < trabajos.length; j++) {
    var h = _calcHoras_(trabajos[j].horaInicio, trabajos[j].horaFin);
    trabajos[j].horas = h;
    totalHoras += h;
  }
  var baseHoras     = totalHoras * precioHora;
  var baseImponible = baseHoras + materiales;
  var ivaImporte    = (tipo === 'Factura') ? baseImponible * ivaPct / 100 : 0;
  var total         = baseImponible + ivaImporte;

  Logger.log('  Horas totales:   ' + totalHoras.toFixed(2) + ' h');
  Logger.log('  Precio/hora:     ' + formatCurrency(precioHora));
  Logger.log('  Mano de obra:    ' + formatCurrency(baseHoras));
  Logger.log('  Mat. trabajos:   ' + formatCurrency(matsTrabajos));
  Logger.log('  Otros mat.:      ' + formatCurrency(otrosMateriales));
  Logger.log('  Materiales:      ' + formatCurrency(materiales));
  Logger.log('  Base imponible:  ' + formatCurrency(baseImponible));
  if (tipo === 'Factura') Logger.log('  IVA (' + ivaPct + '%):        ' + formatCurrency(ivaImporte));
  Logger.log('  TOTAL:           ' + formatCurrency(total));

  // ── 6. Número de factura ─────────────────────────────────────
  Logger.log('\n[5] Generando número de factura...');
  var numFac = tipo === 'Proforma' ? _nextProformaNumber_() : _nextInvoiceNumber_();
  Logger.log('  Número: ' + numFac.id);

  // ── 7. Crear PDF ─────────────────────────────────────────────
  Logger.log('\n[6] Creando PDF...');
  var hoy       = new Date();
  var hoyStr    = Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var docTitulo = numFac.id + '_' + _sanitizeFilename_(cliente);
  var doc       = DocumentApp.create(docTitulo);
  var body      = doc.getBody();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);

  // Cabecera empresa
  body.appendParagraph(empresa)
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .editAsText().setFontSize(16).setBold(true);

  var infoEmp = [];
  if (empresaCif)   infoEmp.push('CIF: ' + empresaCif);
  if (empresaTel)   infoEmp.push('Tel: ' + empresaTel);
  if (empresaEmail) infoEmp.push(empresaEmail);
  if (infoEmp.length > 0)
    body.appendParagraph(infoEmp.join('   ·   ')).editAsText().setFontSize(9);
  if (empresaDir)
    body.appendParagraph(empresaDir).editAsText().setFontSize(9);

  body.appendHorizontalRule();

  // Número y fecha (derecha)
  var tituloDoc = tipo === 'Proforma' ? 'PROFORMA   ' + numFac.id : 'FACTURA   ' + numFac.id;
  body.appendParagraph(tituloDoc)
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph('Fecha de emisión: ' + hoyStr)
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
    .editAsText().setFontSize(10);
  body.appendParagraph('Periodo: ' + fechaIniStr + ' — ' + fechaFinStr)
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
    .editAsText().setFontSize(10);

  body.appendParagraph('');

  // Datos del cliente
  body.appendParagraph('CLIENTE').editAsText().setFontSize(9).setBold(true).setForegroundColor('#555555');
  body.appendParagraph(cliNombre).editAsText().setFontSize(11).setBold(true);
  if (cliCif)   body.appendParagraph('CIF/NIF: ' + cliCif).editAsText().setFontSize(10);
  if (cliDir)   body.appendParagraph(cliDir).editAsText().setFontSize(10);
  if (cliEmail) body.appendParagraph(cliEmail).editAsText().setFontSize(10);

  body.appendHorizontalRule();

  // Tabla de trabajos
  body.appendParagraph('DETALLE DE SERVICIOS').editAsText().setFontSize(11).setBold(true);
  body.appendParagraph('');

  var tableData = [['Fecha', 'Trabajo realizado', 'Horas', 'Incidencias']];
  for (var k = 0; k < trabajos.length; k++) {
    var t = trabajos[k];
    tableData.push([
      formatDate(t.fechaTrabajo),
      t.trabajoRealizado ? t.trabajoRealizado.substring(0, 80) : '—',
      t.horas > 0 ? t.horas.toFixed(2).replace('.', ',') + ' h' : '0,00 h',
      t.incidencias ? t.incidencias.substring(0, 80) : 'Sin incidencias'
    ]);
  }

  var tbl = body.appendTable(tableData);
  for (var col = 0; col < 4; col++) {
    var hdrCell = tbl.getRow(0).getCell(col);
    hdrCell.setBackgroundColor('#1a73e8');
    hdrCell.editAsText().setBold(true).setForegroundColor('#ffffff').setFontSize(9);
  }
  for (var tr = 1; tr < tbl.getNumRows(); tr++) {
    tbl.getRow(tr).editAsText().setFontSize(9);
  }

  body.appendParagraph('');
  body.appendHorizontalRule();

  // Sección materiales adicionales (solo si hay entradas)
  if (matsAdicItems.length > 0) {
    body.appendParagraph('MATERIALES ADICIONALES').editAsText().setFontSize(11).setBold(true);
    body.appendParagraph('');
    var matsRows = matsAdicItems.map(function(m) {
      return [String(m.desc || '(sin descripción)').substring(0, 80), formatCurrency(Number(m.imp) || 0)];
    });
    matsRows.push(['Total materiales', formatCurrency(otrosMateriales)]);
    var tblMats = body.appendTable(matsRows);
    for (var mr = 0; mr < tblMats.getNumRows(); mr++) {
      tblMats.getRow(mr).editAsText().setFontSize(10);
      if (mr === tblMats.getNumRows() - 1) {
        tblMats.getRow(mr).editAsText().setBold(true);
      }
    }
    body.appendParagraph('');
    body.appendHorizontalRule();
  }

  // Tabla de importes
  body.appendParagraph('IMPORTES').editAsText().setFontSize(11).setBold(true);
  body.appendParagraph('');

  var importesData = [
    ['Horas (' + totalHoras.toFixed(2) + ' h × ' + formatCurrency(precioHora) + '/h)', formatCurrency(baseHoras)]
  ];
  if (matsTrabajos > 0) {
    importesData.push(['Materiales de los trabajos', formatCurrency(matsTrabajos)]);
  }
  if (otrosMateriales > 0) {
    importesData.push(['Materiales adicionales', formatCurrency(otrosMateriales)]);
  }
  if (tipo === 'Factura') {
    importesData.push(['Base imponible',        formatCurrency(baseImponible)]);
    importesData.push(['IVA ' + ivaPct + '%',   formatCurrency(ivaImporte)]);
  }
  importesData.push([
    tipo === 'Proforma' ? 'TOTAL (sin impuestos)' : 'TOTAL A PAGAR',
    formatCurrency(total)
  ]);

  var tblImp = body.appendTable(importesData);
  for (var ir = 0; ir < tblImp.getNumRows(); ir++) {
    tblImp.getRow(ir).editAsText().setFontSize(10);
    if (ir === tblImp.getNumRows() - 1) {
      for (var ic = 0; ic < 2; ic++) {
        var totalCell = tblImp.getRow(ir).getCell(ic);
        totalCell.setBackgroundColor('#f0f0f0');
        totalCell.editAsText().setBold(true).setFontSize(12);
      }
    }
  }

  // Sección DATOS DE PAGO (solo si hay método seleccionado)
  if (pagoSeleccionado) {
    body.appendParagraph('');
    body.appendHorizontalRule();
    body.appendParagraph('DATOS DE PAGO').editAsText().setFontSize(11).setBold(true);
    body.appendParagraph('');
    var pagoData = _buildPagoRows_(pagoSeleccionado);
    var tblPago = body.appendTable(pagoData);
    for (var pr = 0; pr < tblPago.getNumRows(); pr++) {
      tblPago.getRow(pr).editAsText().setFontSize(10);
      tblPago.getRow(pr).getCell(0).editAsText().setBold(true);
    }
  }

  // Aviso legal para proformas
  if (tipo === 'Proforma') {
    body.appendParagraph('');
    body.appendParagraph(
      'Este documento no tiene validez fiscal y se emite únicamente como solicitud de pago.'
    ).editAsText().setFontSize(9).setForegroundColor('#c0392b').setItalic(true);
  }

  // Pie de documento
  body.appendParagraph('');
  body.appendHorizontalRule();
  body.appendParagraph(
    'Documento generado el ' +
    Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') +
    '   ·   ' + empresa
  ).editAsText().setFontSize(8).setForegroundColor('#888888');

  doc.saveAndClose();
  Logger.log('  ✓ Doc: "' + docTitulo + '"');

  // ── 8. Convertir a PDF ───────────────────────────────────────
  Logger.log('\n[7] Convirtiendo a PDF...');
  var docFile = DriveApp.getFileById(doc.getId());
  var pdfBlob = docFile.getAs('application/pdf');
  pdfBlob.setName(docTitulo + '.pdf');

  // ── 9. Guardar en Drive ──────────────────────────────────────
  Logger.log('\n[8] Guardando PDF en Drive...');
  var carpetaId = getConfig('carpeta_facturas_id');
  var carpeta;
  if (carpetaId) {
    try {
      carpeta = DriveApp.getFolderById(carpetaId);
      Logger.log('  Carpeta: "' + carpeta.getName() + '"');
    } catch (e) {
      Logger.log('  ⚠️  carpeta_facturas_id no accesible. Guardando en raíz.');
      carpeta = DriveApp.getRootFolder();
    }
  } else {
    Logger.log('  ⚠️  carpeta_facturas_id no configurado. Guardando en raíz.');
    carpeta = DriveApp.getRootFolder();
  }
  var pdfFile = carpeta.createFile(pdfBlob);
  var pdfUrl  = pdfFile.getUrl();
  Logger.log('  ✓ PDF: ' + pdfUrl);
  docFile.setTrashed(true);

  // ── 10. Registrar en hoja ────────────────────────────────────
  var idsTrabajos = trabajos.map(function(t) { return t.id; }).join(', ');

  if (tipo === 'Proforma') {
    Logger.log('\n[9] Registrando en PROFORMAS...');
    var proSheet = _ensureProformasSheet_();
    proSheet.appendRow([
      numFac.id,                                        // ID_Proforma
      hoy,                                              // Fecha_Emision
      numFac.serie,                                     // Serie
      numFac.numero,                                    // Numero
      cliId,                                            // ID_Cliente
      cliNombre,                                        // Nombre_Cliente
      cliEmail,                                         // Email_Cliente
      cliDir,                                           // Direccion_Cliente
      cliCif,                                           // CIF_Cliente
      idsTrabajos,                                      // IDs_Trabajos
      'Servicios ' + fechaIniStr + ' — ' + fechaFinStr, // Concepto
      baseImponible,                                    // Base_Imponible
      total,                                            // Total
      'Emitida',                                        // Estado
      pdfUrl,                                           // URL_PDF
      '',                                               // Fecha_Envio
      '',                                               // Fecha_Pago
      '',                                               // ID_Presupuesto_Origen (vacío — origen trabajos)
      'Pendiente de cobro',                             // Estado_Cobro
      pagoSeleccionado ? pagoSeleccionado.tipo : '',    // Metodo_Pago_Usado
      pagoSeleccionado ? JSON.stringify(pagoSeleccionado) : '' // Datos_Pago_Usados
    ]);
    _guardarMaterialesJson_(proSheet, matsAdicStr);
    Logger.log('  ✓ Registrada en PROFORMAS.');
  } else {
    Logger.log('\n[9] Registrando en FACTURAS...');
    var facSheet = getSheet('FACTURAS');
    facSheet.appendRow([
      numFac.id,                                        // ID_Factura
      hoy,                                              // Fecha_Emision
      numFac.serie,                                     // Serie
      numFac.numero,                                    // Numero
      cliId,                                            // ID_Cliente
      cliNombre,                                        // Nombre_Cliente
      cliEmail,                                         // Email_Cliente
      cliDir,                                           // Direccion_Cliente
      cliCif,                                           // CIF_Cliente
      idsTrabajos,                                      // IDs_Trabajos
      'Servicios ' + fechaIniStr + ' — ' + fechaFinStr, // Concepto
      baseImponible,                                    // Base_Imponible
      ivaPct,                                           // IVA_Porcentaje
      ivaImporte,                                       // IVA_Importe
      0,                                                // IRPF_Porcentaje
      0,                                                // IRPF_Importe
      total,                                            // Total
      'Emitida',                                        // Estado
      pdfUrl,                                           // URL_PDF
      '',                                               // Fecha_Envio
      '',                                               // Fecha_Pago
      pagoSeleccionado ? pagoSeleccionado.tipo : '',    // Metodo_Pago_Usado
      pagoSeleccionado ? JSON.stringify(pagoSeleccionado) : '' // Datos_Pago_Usados
    ]);
    _guardarMaterialesJson_(facSheet, matsAdicStr);

    // Poner Estado Cobro = Pendiente si el módulo de cobros está instalado
    var facHdrs_ = facSheet.getRange(1, 1, 1, facSheet.getLastColumn()).getValues()[0];
    for (var ch_ = 0; ch_ < facHdrs_.length; ch_++) {
      if (String(facHdrs_[ch_]).trim() === 'Estado Cobro') {
        facSheet.getRange(facSheet.getLastRow(), ch_ + 1).setValue('Pendiente');
        break;
      }
    }
    Logger.log('  ✓ Registrada en FACTURAS.');
  }

  logEvent('INFO', 'generateInvoicePDF',
    numFac.id + ' | ' + tipo + ' | ' + cliNombre +
    ' | ' + fechaIniStr + ' → ' + fechaFinStr +
    ' | ' + trabajos.length + ' trabajos' +
    ' | Total: ' + total.toFixed(2) + ' €',
    pdfUrl);

  // ── Resumen ──────────────────────────────────────────────────
  Logger.log('\n================================================');
  Logger.log(tipo === 'Proforma' ? 'PROFORMA GENERADA' : 'FACTURA GENERADA');
  Logger.log('  Número:         ' + numFac.id);
  Logger.log('  Cliente:        ' + cliNombre);
  Logger.log('  Periodo:        ' + fechaIniStr + ' → ' + fechaFinStr);
  Logger.log('  Trabajos:       ' + trabajos.length);
  Logger.log('  Horas:          ' + totalHoras.toFixed(2) + ' h');
  Logger.log('  Base imponible: ' + formatCurrency(baseImponible));
  if (tipo === 'Factura') Logger.log('  IVA (' + ivaPct + '%):       ' + formatCurrency(ivaImporte));
  Logger.log('  TOTAL:          ' + formatCurrency(total));
  Logger.log('  URL PDF:        ' + pdfUrl);
  Logger.log('================================================');
}

// ─────────────────────────────────────────────────────────────

/**
 * Configura las claves de factura en CONFIG y genera el PDF.
 * Solo toca CONFIG — no modifica clientes, operarios, trabajos ni formularios.
 */
function prepareAndGenerateInvoice() {
  setConfig('cliente_factura',      'David Moreno Dias (Sunset Beach Club)');
  setConfig('fecha_inicio_factura', '01/06/2026');
  setConfig('fecha_fin_factura',    '02/06/2026');
  setConfig('precio_hora',          '35');
  setConfig('materiales',           '0');

  Logger.log('CONFIG actualizada:');
  Logger.log('  cliente_factura:       David Moreno Dias (Sunset Beach Club)');
  Logger.log('  fecha_inicio_factura:  01/06/2026');
  Logger.log('  fecha_fin_factura:     02/06/2026');
  Logger.log('  precio_hora:           35');
  Logger.log('  materiales:            0');

  generateInvoicePDF();
}

// ─────────────────────────────────────────────────────────────

/**
 * Verifica que el módulo de facturas puede ejecutarse correctamente.
 * Crea las claves de CONFIG necesarias si no existen (vacías),
 * para que sean visibles en la hoja CONFIG y se puedan rellenar.
 * SOLO LECTURA de datos — no genera facturas.
 */
function validateInvoiceModule() {
  Logger.log('================================================');
  Logger.log('validateInvoiceModule()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var ok = true;

  // 1. Hoja FACTURAS
  Logger.log('\n[1] Hoja FACTURAS:');
  try {
    var shFac = getSheet('FACTURAS');
    Logger.log('  ✓ GID=' + shFac.getSheetId() +
               '  registros=' + Math.max(0, shFac.getLastRow() - 1));
  } catch (e) {
    Logger.log('  ❌ NO EXISTE: ' + e.message);
    ok = false;
  }

  // 2. Hoja TRABAJOS
  Logger.log('\n[2] Hoja TRABAJOS:');
  try {
    var shTra = getSheet('TRABAJOS');
    Logger.log('  ✓ GID=' + shTra.getSheetId() +
               '  registros=' + Math.max(0, shTra.getLastRow() - 1));
  } catch (e) {
    Logger.log('  ❌ NO EXISTE: ' + e.message);
    ok = false;
  }

  // 3. Hoja LOGS
  Logger.log('\n[3] Hoja LOGS:');
  try {
    var shLog = getSheet('LOGS');
    Logger.log('  ✓ GID=' + shLog.getSheetId());
  } catch (e) {
    Logger.log('  ❌ NO EXISTE: ' + e.message);
    ok = false;
  }

  // 4. Carpeta Drive
  Logger.log('\n[4] Carpeta Facturas (carpeta_facturas_id):');
  var carpetaId = getConfig('carpeta_facturas_id');
  if (!carpetaId) {
    Logger.log('  ⚠️  No configurada. PDFs se guardarán en la raíz de Drive.');
  } else {
    try {
      var carpeta = DriveApp.getFolderById(carpetaId);
      Logger.log('  ✓ "' + carpeta.getName() + '"  ID=' + carpetaId);
    } catch (e) {
      Logger.log('  ❌ No accesible: ' + e.message);
      ok = false;
    }
  }

  // 5. DocumentApp y DriveApp
  Logger.log('\n[5] DocumentApp y DriveApp:');
  try {
    var testDoc = DocumentApp.create('_validateInvoiceModule_test_');
    testDoc.saveAndClose();
    DriveApp.getFileById(testDoc.getId()).setTrashed(true);
    Logger.log('  ✓ Accesibles.');
  } catch (e) {
    Logger.log('  ❌ ' + e.message);
    ok = false;
  }

  // 6. CONFIG — datos de empresa
  Logger.log('\n[6] CONFIG — datos de empresa:');
  var clavesEmp = [
    'empresa_nombre', 'empresa_cif', 'empresa_direccion',
    'empresa_email', 'empresa_telefono', 'iva_porcentaje', 'irpf_porcentaje'
  ];
  for (var i = 0; i < clavesEmp.length; i++) {
    var v = getConfig(clavesEmp[i]);
    Logger.log('  ' + clavesEmp[i] + ': ' + (v ? '"' + v + '"' : '⚠️  (vacío — rellenar en CONFIG)'));
  }

  // 7. CONFIG — claves de factura (inicializar vacías si no existen)
  Logger.log('\n[7] CONFIG — claves de factura (rellenar antes de ejecutar):');
  var clavesFac = [
    ['cliente_factura',      'Nombre del cliente tal como aparece en TRABAJOS'],
    ['fecha_inicio_factura', 'DD/MM/YYYY — inicio del periodo a facturar'],
    ['fecha_fin_factura',    'DD/MM/YYYY — fin del periodo a facturar'],
    ['precio_hora',          'Precio por hora en EUR (ej: 35)']
  ];
  for (var j = 0; j < clavesFac.length; j++) {
    var clave = clavesFac[j][0];
    var desc  = clavesFac[j][1];
    if (getConfig(clave) === null) setConfig(clave, '');
    var val = getConfig(clave);
    Logger.log('  ' + clave + ': ' + (val ? '"' + val + '"' : '(vacío) — ' + desc));
  }

  // 8. Siguiente número de factura
  Logger.log('\n[8] Siguiente número de factura:');
  try {
    var sig = _nextInvoiceNumber_();
    Logger.log('  ' + sig.id);
  } catch (e) {
    Logger.log('  Error: ' + e.message);
  }

  // Resumen
  Logger.log('\n================================================');
  Logger.log('RESULTADO:');
  Logger.log('  Infraestructura: ' + (ok ? '✓' : '❌'));
  if (ok) {
    Logger.log('  MÓDULO LISTO.');
    Logger.log('  Rellena las claves de CONFIG mostradas en [7] y ejecuta generateInvoicePDF().');
  } else {
    Logger.log('  HAY PROBLEMAS — revisa los ❌ anteriores.');
  }
  Logger.log('================================================');
}

// ── Funciones privadas ────────────────────────────────────────

/**
 * Escribe el JSON de materiales adicionales en la columna Materiales_JSON
 * de la última fila de la hoja dada. Crea la cabecera si no existe.
 */
function _guardarMaterialesJson_(sheet, matsJson) {
  try {
    if (!matsJson || matsJson === '[]') return;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var col = -1;
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim() === 'Materiales_JSON') { col = i + 1; break; }
    }
    if (col === -1) {
      col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue('Materiales_JSON');
    }
    sheet.getRange(sheet.getLastRow(), col).setValue(matsJson);
  } catch (e) {
    Logger.log('  ⚠️  _guardarMaterialesJson_: ' + e.message);
  }
}

/**
 * Devuelve el siguiente número de factura del año en curso.
 * Formato: SERIE-YYYY-NNN (ej: FAC-2026-001).
 */
function _nextInvoiceNumber_(serie) {
  var year   = String(new Date().getFullYear());
  serie = serie || getConfig('serie_factura') || 'FAC';
  var prefix = serie + '-' + year + '-';

  var sheet   = getSheet('FACTURAS');
  var lastRow = sheet.getLastRow();
  var maxNum  = 0;

  if (lastRow > 1) {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    for (var i = 0; i < ids.length; i++) {
      var idStr = String(ids[i]);
      if (idStr.indexOf(prefix) === 0) {
        var n = parseInt(idStr.split('-').pop(), 10);
        if (!isNaN(n) && n > maxNum) maxNum = n;
      }
    }
  }

  var numero = maxNum + 1;
  return {
    id:     prefix + String(numero).padStart(3, '0'),
    serie:  serie,
    numero: numero
  };
}

/**
 * Devuelve el siguiente número de proforma del año en curso.
 * Formato: PRO-YYYY-NNN. Lee de la hoja PROFORMAS.
 */
function _nextProformaNumber_() {
  var year   = String(new Date().getFullYear());
  var serie  = 'PRO';
  var prefix = serie + '-' + year + '-';
  var sheet  = _ensureProformasSheet_();
  var lastRow = sheet.getLastRow();
  var maxNum  = 0;
  if (lastRow > 1) {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    for (var i = 0; i < ids.length; i++) {
      var idStr = String(ids[i]);
      if (idStr.indexOf(prefix) === 0) {
        var n = parseInt(idStr.split('-').pop(), 10);
        if (!isNaN(n) && n > maxNum) maxNum = n;
      }
    }
  }
  var numero = maxNum + 1;
  return {
    id:     prefix + String(numero).padStart(3, '0'),
    serie:  serie,
    numero: numero
  };
}

/**
 * Devuelve la hoja PROFORMAS, creándola con cabecera si no existe.
 * Siempre asegura la columna Estado_Cobro con validación y migración.
 */
function _ensureProformasSheet_() {
  var ss    = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('PROFORMAS');
  if (!sheet) {
    sheet = ss.insertSheet('PROFORMAS');
    var hdrs = SHEET_HEADERS.PROFORMAS;
    sheet.appendRow(hdrs);
    sheet.getRange(1, 1, 1, hdrs.length)
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  } else if (sheet.getLastRow() === 0) {
    var hdrs2 = SHEET_HEADERS.PROFORMAS;
    sheet.appendRow(hdrs2);
    sheet.getRange(1, 1, 1, hdrs2.length)
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  }
  _ensureProformasEstadoCobro_(sheet);
  return sheet;
}

/**
 * Asegura que la hoja PROFORMAS tiene la columna Estado_Cobro con:
 *  - Cabecera en la posición correcta
 *  - Validación de datos ['Pendiente de cobro', 'Pagada', 'Anulada'] en toda la columna
 *  - Migración: filas existentes vacías o con 'Emitida' → 'Pendiente de cobro'
 */
function _ensureProformasEstadoCobro_(sheet) {
  var COBRO_VALS = ['Pendiente de cobro', 'Pagada', 'Anulada'];

  // 1. Buscar o añadir columna Estado_Cobro
  var lastCol  = sheet.getLastColumn();
  var hdrs     = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var colCobro = -1;
  for (var h = 0; h < hdrs.length; h++) {
    if (String(hdrs[h]).trim() === 'Estado_Cobro') { colCobro = h + 1; break; }
  }
  if (colCobro < 0) {
    colCobro = lastCol + 1;
    sheet.getRange(1, colCobro)
      .setValue('Estado_Cobro')
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  }

  // 2. Validación en toda la columna (fila 2 hasta el final de la hoja)
  var validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(COBRO_VALS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, colCobro, sheet.getMaxRows() - 1, 1).setDataValidation(validation);

  // 3. Migrar filas existentes: vacío o 'Emitida' → 'Pendiente de cobro'
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var dataRows  = lastRow - 1;
  var idVals    = sheet.getRange(2, 1, dataRows, 1).getValues();
  var cobroVals = sheet.getRange(2, colCobro, dataRows, 1).getValues();
  var updates   = [];
  var changed   = false;
  for (var i = 0; i < dataRows; i++) {
    var hasId = String(idVals[i][0] || '').trim() !== '';
    var cobro = String(cobroVals[i][0] || '').trim();
    if (hasId && (cobro === '' || cobro === 'Emitida')) {
      updates.push(['Pendiente de cobro']);
      changed = true;
    } else {
      updates.push([cobroVals[i][0]]);
    }
  }
  if (changed) {
    sheet.getRange(2, colCobro, dataRows, 1).setValues(updates);
  }
}

/**
 * Convierte un objeto método de pago en filas [['campo', 'valor'], ...] para tabla PDF.
 */
function _buildPagoRows_(m) {
  var rows = [['Método de pago', m.tipo || '']];
  var tl = (m.tipo || '').toLowerCase();
  if (tl === 'transferencia bancaria' || tl === 'domiciliación bancaria') {
    if (m.titular) rows.push(['Titular', m.titular]);
    if (m.banco)   rows.push(['Banco',   m.banco]);
    if (m.iban)    rows.push(['IBAN',    m.iban]);
    if (m.swift)   rows.push(['SWIFT / BIC', m.swift]);
  } else if (tl === 'bizum') {
    if (m.numero)  rows.push(['Número Bizum', m.numero]);
  } else if (tl === 'paypal') {
    if (m.email)   rows.push(['Email PayPal', m.email]);
  } else if (m.observaciones) {
    rows.push(['Observaciones', m.observaciones]);
  }
  return rows;
}

/**
 * Busca en CLIENTES la fila que corresponde a la etiqueta del dropdown.
 * Formato de etiqueta: "Nombre Apellidos (Empresa)" o "Nombre Apellidos".
 * Devuelve { id, nombre, email, telefono, direccion, cif, metodosPago } o null.
 */
function _getClienteByLabel_(label) {
  try {
    var sheet = getSheet(SHEET_NAMES.CLIENTES);
    var data  = sheet.getDataRange().getValues();
    if (data.length < 2) return null;
    var labelLow = String(label).toLowerCase().trim();

    for (var i = 1; i < data.length; i++) {
      var row       = data[i];
      if (!row.some(function(c) { return c !== ''; })) continue; // fila vacía
      var nombre    = String(row[COL_CLI.NOMBRE    - 1] || '').trim();
      var apellidos = String(row[COL_CLI.APELLIDOS - 1] || '').trim();
      var empresa_  = String(row[COL_CLI.EMPRESA   - 1] || '').trim();
      var dropLabel = (nombre + ' ' + apellidos).trim();
      if (empresa_) dropLabel += ' (' + empresa_ + ')';
      if (dropLabel.toLowerCase() !== labelLow) continue;

      // Leer metodosPago por posición de columna para evitar dependencia del nombre de cabecera
      var metodosPago = [];
      try {
        var mpRaw = row.length >= COL_CLI.METODOS_PAGO ? row[COL_CLI.METODOS_PAGO - 1] : '';
        var mp_ = JSON.parse(String(mpRaw || '[]'));
        if (Array.isArray(mp_)) metodosPago = mp_;
      } catch (_) {}

      return {
        id:          String(row[COL_CLI.ID        - 1] || ''),
        nombre:      dropLabel,
        email:       String(row[COL_CLI.EMAIL     - 1] || '').trim(),
        telefono:    String(row[COL_CLI.TELEFONO  - 1] || '').trim(),
        direccion:   String(row[COL_CLI.DIRECCION - 1] || '').trim(),
        cif:         String(row[COL_CLI.CIF_NIF   - 1] || '').trim(),
        metodosPago: metodosPago
      };
    }
  } catch (e) {
    Logger.log('  Error buscando cliente: ' + e.message);
  }
  return null;
}
