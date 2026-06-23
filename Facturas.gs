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

  // ── 0. Validación anti-datos-de-prueba ────────────────────────
  var _errEmpresa = _validarDatosEmpresa_();
  var _errCliente = _validarClienteParaFactura_(cliente);
  var _todosErrores = _errEmpresa.concat(_errCliente);
  if (_todosErrores.length > 0) {
    var _msgErr = '❌ FACTURA NO GENERADA\n\n' + _todosErrores.join('\n') +
                  '\n\nCorrige estos datos antes de generar la factura.';
    Logger.log(_msgErr);
    try { SpreadsheetApp.getUi().alert(_msgErr); } catch (_e_) {}
    return;
  }

  // ── 0b. Validar carpeta Drive de destino ─────────────────────
  var _carpetaKey_    = (tipo === 'Proforma') ? 'carpeta_proformas_id' : 'carpeta_facturas_id';
  var _carpetaRootId_ = getConfig(_carpetaKey_);
  if (!_carpetaRootId_) {
    var _msgCarp_ = '❌ ' + tipo.toUpperCase() + ' NO GENERADA\n\n' +
                    'Carpeta Drive no configurada en CONFIG.\n' +
                    'Clave requerida: ' + _carpetaKey_ + '\n\n' +
                    'Ejecuta: ⚙️ Sistema → 📁 DRIVE — Recuperar carpetas existentes';
    Logger.log('[BLOQUEADO] ' + _msgCarp_);
    try { SpreadsheetApp.getUi().alert(_msgCarp_); } catch (_e2_) {}
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
    // Leer hasta la última columna del empleado 2 (col 22) para incluir sus horas.
    // En hojas antiguas sin esas columnas, getLastColumn() las acota y el empleado 2
    // queda vacío (horas = 0), manteniendo el comportamiento previo.
    var numCols = Math.min(shTra.getLastColumn(), COL_TRA.HORA_FIN_2);
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
        // Horas totales del parte = empleado 1 + empleado 2 (si lo hay).
        horas:            _horasTotalesTrabajoRow_(row),
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
    // trabajos[j].horas ya incluye empleado 1 + empleado 2 (calculado al leer la fila).
    totalHoras += trabajos[j].horas;
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

  // Cabecera empresa (logo izquierda + info derecha)
  _cabecera_empresa_(body, empresa, empresaCif, empresaTel, empresaEmail, empresaDir, doc.getId());

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

  // ── 9. Guardar en Drive (AÑO/MES automático) ─────────────────
  Logger.log('\n[8] Guardando PDF en Drive...');
  var carpeta;
  try {
    carpeta = _getOrCreateSubcarpetaAnioMes_(_carpetaRootId_, hoy);
  } catch (e) {
    Logger.log('  ⚠️ Error al obtener subcarpeta AÑO/MES: ' + e.message + '. Intentando carpeta raíz.');
    try { carpeta = DriveApp.getFolderById(_carpetaRootId_); }
    catch (e2) { carpeta = DriveApp.getRootFolder(); }
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
 * DESACTIVADA — función de prueba que hardcodeaba datos de "David Moreno Dias".
 * Si esta función se ejecutó en el pasado, puede haber contaminado CONFIG.
 * Para generar facturas: usa PANEL_FACTURAS → selecciona cliente → busca trabajos → generar.
 */
function prepareAndGenerateInvoice() {
  SpreadsheetApp.getUi().alert(
    '⛔ FUNCIÓN DESACTIVADA\n\n' +
    'Esta función de prueba escribía datos hardcodeados (David Moreno Dias) en CONFIG.\n\n' +
    'Para generar facturas usa:\n' +
    'PANEL_FACTURAS → selecciona cliente → selecciona fechas → BUSCAR TRABAJOS → GENERAR DOCUMENTO'
  );
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
 * Valida que los datos de empresa en CONFIG no sean valores de prueba.
 * Solo bloquea valores de prueba literales, no campos vacíos opcionales.
 * Devuelve array de mensajes de error (vacío = OK).
 */
function _validarDatosEmpresa_() {
  var errores = [];
  var nombre = String(getConfig('empresa_nombre') || '').trim();
  var cif    = String(getConfig('empresa_cif')    || '').trim();
  var dir    = String(getConfig('empresa_direccion') || '').trim();

  // Bloquear nombre vacío o valor de prueba
  if (!nombre || nombre.toLowerCase() === 'mi empresa') {
    errores.push('empresa_nombre: "' + (nombre || '(vacío)') + '" — configura el nombre real de la empresa en la hoja CONFIG');
  }
  // Bloquear CIF vacío o valor de prueba
  if (!cif || cif.toLowerCase() === 'a00000000') {
    errores.push('empresa_cif: "' + (cif || '(vacío)') + '" — configura el CIF/NIF real en la hoja CONFIG');
  }
  // Bloquear solo si contiene "Calle Ejemplo" (dirección vacía es permitida)
  if (dir && dir.toLowerCase().indexOf('calle ejemplo') >= 0) {
    errores.push('empresa_direccion: "' + dir + '" — contiene valor de prueba. Configura la dirección real en CONFIG');
  }

  return errores;
}

/**
 * Muestra en un alert los valores exactos de los campos de empresa en CONFIG.
 * Úsalo para verificar que los datos se guardaron correctamente.
 * Menú: ⚙️ Sistema → 🏢 VER — Datos empresa en CONFIG
 */
function verDatosEmpresaConfig() {
  var claves = [
    'empresa_nombre', 'empresa_cif', 'empresa_direccion',
    'empresa_email', 'empresa_telefono', 'iva_porcentaje'
  ];
  var lineas = ['Valores actuales en CONFIG:\n'];
  for (var i = 0; i < claves.length; i++) {
    var val = getConfig(claves[i]);
    lineas.push(claves[i] + ':');
    lineas.push('  → ' + (val !== null ? '"' + val + '"' : '(clave no existe en CONFIG)'));
  }
  var errEmp = _validarDatosEmpresa_();
  lineas.push('');
  lineas.push(errEmp.length === 0
    ? '✅ Datos de empresa válidos — se puede generar factura'
    : '❌ Errores:\n  ' + errEmp.join('\n  '));
  SpreadsheetApp.getUi().alert('DATOS EMPRESA EN CONFIG', lineas.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Valida que el cliente exista en CLIENTES y esté activo.
 * Devuelve array de mensajes de error (vacío = OK).
 */
function _validarClienteParaFactura_(cliente) {
  var errores = [];
  if (!cliente) {
    errores.push('No hay cliente seleccionado en PANEL_FACTURAS!C5');
    return errores;
  }

  var CLIENTES_PRUEBA = ['david moreno', 'ana fernandez', 'sofia torres', 'carlos ruiz', 'cristina vega', 'juan martinez', 'ana f'];
  var clienteLow = cliente.toLowerCase();
  for (var i = 0; i < CLIENTES_PRUEBA.length; i++) {
    if (clienteLow.indexOf(CLIENTES_PRUEBA[i]) >= 0) {
      errores.push('El cliente "' + cliente + '" es un cliente de prueba eliminado del sistema. No se puede facturar.');
      return errores;
    }
  }

  var cliData = _getClienteByLabel_(cliente);
  if (!cliData) {
    errores.push('Cliente "' + cliente + '" no encontrado en la hoja CLIENTES — puede ser un problema de formato del nombre. Ejecuta el diagnóstico de factura para más detalles.');
    return errores;
  }

  try {
    var sheet = getSheet(SHEET_NAMES.CLIENTES);
    var data  = sheet.getDataRange().getValues();
    var clienteLowNorm = clienteLow.replace(/\s+/g, ' ');
    for (var j = 1; j < data.length; j++) {
      var row       = data[j];
      var nombre_   = String(row[COL_CLI.NOMBRE    - 1] || '').trim();
      var apells_   = String(row[COL_CLI.APELLIDOS - 1] || '').trim();
      var emp_      = String(row[COL_CLI.EMPRESA   - 1] || '').trim();
      var label_    = (nombre_ + ' ' + apells_).trim();
      if (emp_) label_ += ' (' + emp_ + ')';
      if (label_.toLowerCase().replace(/\s+/g, ' ') !== clienteLowNorm) continue;
      var activo = row[COL_CLI.ACTIVO - 1];
      if (activo !== true && String(activo).toLowerCase() !== 'true') {
        errores.push('El cliente "' + cliente + '" está marcado como INACTIVO en CLIENTES (columna J). Actívalo antes de facturar.');
      }
      break;
    }
  } catch (e) {
    Logger.log('_validarClienteParaFactura_: ' + e.message);
  }

  return errores;
}

/**
 * Genera un informe diagnóstico completo en la hoja DIAGNOSTICO_FACTURA_DATOS.
 * Muestra: celda cliente leída, cliente encontrado en CLIENTES, trabajos seleccionados,
 * datos empresa usados, errores detectados.
 * Menú: ⚙️ Sistema → 🔍 DIAGNÓSTICO — Datos de factura
 */
function diagnosticarFacturaDatos() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('PANEL_FACTURAS');
  var resultado = [];
  var errores   = [];
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');

  resultado.push(['DIAGNÓSTICO DATOS FACTURA — ' + ts, '']);
  resultado.push(['', '']);

  // 1. Panel FACTURAS
  resultado.push(['=== PANEL_FACTURAS ===', '']);
  var cliente = '', fechaIni = '', fechaFin = '', tipo = '', precioHora = 0;
  if (!sh) {
    resultado.push(['Hoja PANEL_FACTURAS', '❌ NO EXISTE']);
    errores.push('La hoja PANEL_FACTURAS no existe');
  } else {
    cliente    = String(sh.getRange(P_.CLIENTE[0],     P_.CLIENTE[1]).getValue()     || '').trim();
    fechaIni   = sh.getRange(P_.FECHA_INI[0],   P_.FECHA_INI[1]).getDisplayValue().trim();
    fechaFin   = sh.getRange(P_.FECHA_FIN[0],   P_.FECHA_FIN[1]).getDisplayValue().trim();
    tipo       = String(sh.getRange(P_.TIPO[0],         P_.TIPO[1]).getValue()         || '').trim();
    precioHora = Number(sh.getRange(P_.PRECIO_HORA[0], P_.PRECIO_HORA[1]).getValue()) || 0;
    resultado.push(['Celda cliente leída', 'PANEL_FACTURAS!C5']);
    resultado.push(['Cliente seleccionado (C5)', cliente || '(vacío)']);
    resultado.push(['Fecha Inicio (C6)', fechaIni || '(vacío)']);
    resultado.push(['Fecha Fin (C7)', fechaFin || '(vacío)']);
    resultado.push(['Tipo documento (C8)', tipo || '(vacío)']);
    resultado.push(['Precio/hora (C10)', precioHora > 0 ? precioHora + ' €/h' : '(no configurado)']);
    if (!cliente) errores.push('No hay cliente seleccionado en C5');
    if (!fechaIni) errores.push('Fecha Inicio vacía en C6');
    if (!fechaFin) errores.push('Fecha Fin vacía en C7');
    if (precioHora <= 0 && tipo !== 'Informe') errores.push('Precio/hora debe ser mayor que 0');
  }

  // 2. Cliente en CLIENTES
  resultado.push(['', '']);
  resultado.push(['=== CLIENTE EN CLIENTES ===', '']);
  if (cliente) {
    var errCli = _validarClienteParaFactura_(cliente);
    var cliData = _getClienteByLabel_(cliente);
    if (cliData) {
      resultado.push(['Encontrado en CLIENTES', '✓ SÍ']);
      resultado.push(['ID', cliData.id]);
      resultado.push(['Nombre completo', cliData.nombre]);
      resultado.push(['Email', cliData.email || '(vacío)']);
      resultado.push(['Teléfono', cliData.telefono || '(vacío)']);
      resultado.push(['Dirección', cliData.direccion || '(vacío)']);
      resultado.push(['CIF/NIF', cliData.cif || '(vacío)']);
      resultado.push(['Estado CLIENTES', errCli.length === 0 ? '✓ ACTIVO' : '❌ INACTIVO']);
    } else {
      resultado.push(['Encontrado en CLIENTES', '❌ NO']);
    }
    for (var ec = 0; ec < errCli.length; ec++) errores.push(errCli[ec]);
  } else {
    resultado.push(['Encontrado en CLIENTES', '— sin cliente seleccionado']);
  }

  // 3. Trabajos seleccionados
  resultado.push(['', '']);
  resultado.push(['=== TRABAJOS SELECCIONADOS ===', '']);
  var nSeleccionados = 0;
  if (sh) {
    var jobsData = sh.getRange(
      JOB_TABLE.START_ROW, JOB_TABLE.COL_CHECK,
      JOB_TABLE.MAX_ROWS,  JOB_TABLE.COL_DATA - JOB_TABLE.COL_CHECK + 1
    ).getValues();
    for (var jj = 0; jj < jobsData.length; jj++) {
      var checked = jobsData[jj][0];
      if (checked !== true && String(checked) !== 'TRUE') continue;
      var jsonStr = String(jobsData[jj][JOB_TABLE.COL_DATA - JOB_TABLE.COL_CHECK] || '').trim();
      nSeleccionados++;
      try {
        var jd = JSON.parse(jsonStr);
        resultado.push(['Trabajo ' + nSeleccionados,
          (jd.id || '?') + ' | ' + formatDate(new Date(jd.fechaTrabajo)) +
          ' | ' + String(jd.trabajoRealizado || '').substring(0, 50)]);
      } catch (_) {
        resultado.push(['Trabajo ' + nSeleccionados, jsonStr.substring(0, 80)]);
      }
    }
    resultado.push(['Total trabajos seleccionados', nSeleccionados]);
    if (nSeleccionados === 0) errores.push('No hay trabajos seleccionados en el panel (pulsa BUSCAR TRABAJOS y marca trabajos)');
  }

  // 4. Datos empresa en CONFIG
  resultado.push(['', '']);
  resultado.push(['=== DATOS EMPRESA (CONFIG) ===', '']);
  var empNombre = String(getConfig('empresa_nombre')    || '').trim();
  var empCif    = String(getConfig('empresa_cif')       || '').trim();
  var empDir    = String(getConfig('empresa_direccion') || '').trim();
  var empEmail  = String(getConfig('empresa_email')     || '').trim();
  var empTel    = String(getConfig('empresa_telefono')  || '').trim();
  var empIva    = String(getConfig('iva_porcentaje')    || '21').trim();
  resultado.push(['empresa_nombre',    empNombre || '(vacío)']);
  resultado.push(['empresa_cif',       empCif    || '(vacío)']);
  resultado.push(['empresa_direccion', empDir    || '(vacío)']);
  resultado.push(['empresa_email',     empEmail  || '(vacío)']);
  resultado.push(['empresa_telefono',  empTel    || '(vacío)']);
  resultado.push(['iva_porcentaje',    empIva + '%']);
  var errEmp = _validarDatosEmpresa_();
  for (var ee = 0; ee < errEmp.length; ee++) {
    resultado.push(['⚠️ ERROR EMPRESA', errEmp[ee]]);
    errores.push(errEmp[ee]);
  }
  if (errEmp.length === 0) resultado.push(['Estado empresa', '✓ Datos de empresa configurados correctamente']);

  // 5. CONFIG claves factura (estado actual)
  resultado.push(['', '']);
  resultado.push(['=== CONFIG CLAVES FACTURA (estado actual) ===', '']);
  var cfgCliente  = String(getConfig('cliente_factura')      || '').trim();
  var cfgFechaIni = String(getConfig('fecha_inicio_factura') || '').trim();
  var cfgFechaFin = String(getConfig('fecha_fin_factura')    || '').trim();
  var cfgPrecio   = String(getConfig('precio_hora')          || '').trim();
  resultado.push(['cliente_factura (CONFIG)',      cfgCliente  || '(vacío)']);
  resultado.push(['fecha_inicio_factura (CONFIG)', cfgFechaIni || '(vacío)']);
  resultado.push(['fecha_fin_factura (CONFIG)',     cfgFechaFin || '(vacío)']);
  resultado.push(['precio_hora (CONFIG)',           cfgPrecio   || '(vacío)']);
  if (cfgCliente && cliente && cfgCliente.toLowerCase() !== cliente.toLowerCase()) {
    resultado.push(['⚠️ DESINCRONIZACIÓN', 'Panel tiene "' + cliente + '" pero CONFIG tiene "' + cfgCliente + '"']);
    errores.push('DESINCRONIZACIÓN: panel tiene "' + cliente + '" pero CONFIG tiene "' + cfgCliente + '" (puede haber sido contaminado por una ejecución anterior)');
  } else if (cfgCliente && cliente && cfgCliente.toLowerCase() === cliente.toLowerCase()) {
    resultado.push(['Sincronización panel↔CONFIG', '✓ Coinciden']);
  }

  // 6. Resumen errores
  resultado.push(['', '']);
  resultado.push(['=== RESUMEN ===', '']);
  if (errores.length === 0) {
    resultado.push(['Estado general', '✅ Sin errores — se puede generar la factura']);
  } else {
    resultado.push(['Total errores', errores.length + ' error(es) detectado(s)']);
    for (var ei = 0; ei < errores.length; ei++) {
      resultado.push(['❌ Error ' + (ei + 1), errores[ei]]);
    }
    resultado.push(['Acción requerida', 'Corrige los errores antes de generar la factura']);
  }

  // Escribir hoja
  var diagSheet = getOrCreateSheet('DIAGNOSTICO_FACTURA_DATOS');
  diagSheet.clearContents();
  diagSheet.clearFormats();
  diagSheet.getRange(1, 1, resultado.length, 2).setValues(resultado);
  diagSheet.getRange(1, 1, 1, 2)
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  for (var rr = 0; rr < resultado.length; rr++) {
    var k = String(resultado[rr][0]);
    var bg = null;
    if (k.indexOf('❌') >= 0 || k.indexOf('⚠️') >= 0)       bg = '#fce8e6';
    else if (k.indexOf('✓') >= 0 || k.indexOf('✅') >= 0)   bg = '#e6f4ea';
    else if (k.indexOf('===') >= 0)                          bg = '#e8eaed';
    if (bg) diagSheet.getRange(rr + 1, 1, 1, 2).setBackground(bg).setFontWeight('bold');
  }
  diagSheet.autoResizeColumns(1, 2);
  ss.setActiveSheet(diagSheet);

  SpreadsheetApp.getUi().alert(
    errores.length === 0
      ? '✅ DIAGNÓSTICO OK\n\nNo se detectaron errores.\nPuedes generar la factura.'
      : '❌ DIAGNÓSTICO — ' + errores.length + ' ERROR(ES)\n\n' +
        errores.join('\n\n') +
        '\n\nRevisa la hoja DIAGNOSTICO_FACTURA_DATOS para el informe completo.'
  );
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
    // Normaliza espacios múltiples (la fórmula FILTER usa "  (" con doble espacio)
    var labelLow = String(label).toLowerCase().trim().replace(/\s+/g, ' ');

    for (var i = 1; i < data.length; i++) {
      var row       = data[i];
      if (!row.some(function(c) { return c !== ''; })) continue; // fila vacía
      var nombre    = String(row[COL_CLI.NOMBRE    - 1] || '').trim();
      var apellidos = String(row[COL_CLI.APELLIDOS - 1] || '').trim();
      var empresa_  = String(row[COL_CLI.EMPRESA   - 1] || '').trim();
      var dropLabel = (nombre + ' ' + apellidos).trim();
      if (empresa_) dropLabel += ' (' + empresa_ + ')';
      if (dropLabel.toLowerCase().replace(/\s+/g, ' ') !== labelLow) continue;

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
