// ============================================================
// FacturasPresupuesto.gs — Conversión de presupuesto en factura PDF
// ============================================================
//
// Punto de entrada:
//   generateInvoiceFromPresupuestoPDF(params)
//     params.presupuestoId  — ID del presupuesto origen
//     params.cliente        — etiqueta del cliente (tal como sale en el dropdown)
//     params.lineItems      — array de { desc, cant, pu, imp }
//
// Efectos secundarios:
//   · Crea PDF en carpeta carpeta_facturas_id (o raíz si no está configurado)
//   · Registra la factura en hoja FACTURAS (cols 1-21 + col 22 ID_Presupuesto_Origen)
//   · Marca el presupuesto como 'Facturado' via updatePresupuestoEstado()
// ============================================================

/**
 * Genera una factura PDF a partir de un presupuesto y la registra en FACTURAS.
 * @param {{presupuestoId:string, cliente:string, lineItems:Array}} params
 * @returns {{id:string, url:string}}
 */
function generateInvoiceFromPresupuestoPDF(params) {
  var presupuestoId = String(params.presupuestoId || '').trim();
  var clienteLabel  = String(params.cliente       || '').trim();
  var lineItems     = params.lineItems || [];
  var tipo          = String(params.tipo          || 'Factura').trim();

  if (!presupuestoId) throw new Error('presupuestoId es obligatorio');
  if (!lineItems.length) throw new Error('lineItems está vacío — no se puede emitir una factura sin líneas');

  // ── 1. Datos de empresa ──────────────────────────────────────
  var empresa      = getConfig('empresa_nombre')    || 'Mi Empresa';
  var empresaCif   = getConfig('empresa_cif')       || '';
  var empresaDir   = getConfig('empresa_direccion') || '';
  var empresaEmail = getConfig('empresa_email')     || '';
  var empresaTel   = getConfig('empresa_telefono')  || '';
  var ivaPct       = parseFloat(getConfig('iva_porcentaje') || '21');

  // ── 2. Datos del cliente ─────────────────────────────────────
  var cliData   = _getClienteByLabel_(clienteLabel);
  var cliNombre = cliData ? cliData.nombre    : clienteLabel;
  var cliCif    = cliData ? cliData.cif       : '';
  var cliDir    = cliData ? cliData.direccion : '';
  var cliEmail  = cliData ? cliData.email     : '';
  var cliId     = cliData ? cliData.id        : '';

  // ── Selección de método de pago ──────────────────────────────
  var metodosPago_fp = cliData ? (cliData.metodosPago || []) : [];
  var pagoSeleccionado = null;
  if (params.metodoPago !== undefined) {
    pagoSeleccionado = params.metodoPago;
    Logger.log('  Método de pago (override): ' + (pagoSeleccionado ? pagoSeleccionado.tipo : 'ninguno'));
  } else if (metodosPago_fp.length === 1) {
    pagoSeleccionado = metodosPago_fp[0];
  } else if (metodosPago_fp.length > 1) {
    var opsTxt_fp = metodosPago_fp.map(function(m, i) {
      var extra = m.iban ? ' (' + m.iban.replace(/\s/g,'').substring(0,10) + '…)' :
                  m.numero ? ' (' + m.numero + ')' : m.email ? ' (' + m.email + ')' : '';
      return (i + 1) + '. ' + m.tipo + extra;
    }).join('\n');
    var selResp_fp = SpreadsheetApp.getUi().prompt(
      'Método de pago para esta ' + tipo,
      'El cliente tiene varios métodos. Escribe el número:\n\n' + opsTxt_fp + '\n\n0. No incluir datos de pago',
      SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
    );
    if (selResp_fp.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
      var selIdx_fp = parseInt(selResp_fp.getResponseText()) - 1;
      if (selIdx_fp >= 0 && selIdx_fp < metodosPago_fp.length) pagoSeleccionado = metodosPago_fp[selIdx_fp];
    }
  }

  // ── 3. Cálculo de importes desde las líneas de la tabla ──────
  // IVA nunca heredado del presupuesto — se aplica una sola vez aquí.
  // Fallback: si cant o pu están vacíos, usa lineItems[i].imp directamente.
  var subtotal = 0;
  for (var i = 0; i < lineItems.length; i++) {
    var li = lineItems[i];
    var lineImp = (li.cant > 0 && li.pu > 0)
      ? Math.round((li.cant || 0) * (li.pu || 0) * 100) / 100
      : (Number(li.imp) || 0);
    subtotal += lineImp;
  }
  subtotal         = Math.round(subtotal * 100) / 100;
  var ivaImporte   = tipo === 'Proforma' ? 0 : Math.round(subtotal * ivaPct) / 100;
  var total        = Math.round((subtotal + ivaImporte) * 100) / 100;

  // ── 4. Número de documento ───────────────────────────────────
  var numFac = tipo === 'Proforma' ? _nextProformaNumber_() : _nextInvoiceNumber_();

  // ── 5. Crear Google Doc → PDF ────────────────────────────────
  var hoy       = new Date();
  var tz        = Session.getScriptTimeZone();
  var hoyStr    = Utilities.formatDate(hoy, tz, 'dd/MM/yyyy');
  var docTitulo = numFac.id + '_' + _sanitizeFilename_(clienteLabel);
  var doc       = DocumentApp.create(docTitulo);
  var body      = doc.getBody();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);

  // Cabecera empresa (logo izquierda + info derecha)
  _cabecera_empresa_(body, empresa, empresaCif, empresaTel, empresaEmail, empresaDir, doc.getId());

  // Número y fecha (alineado a la derecha)
  body.appendParagraph((tipo === 'Proforma' ? 'PROFORMA   ' : 'FACTURA   ') + numFac.id)
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph('Fecha de emisión: ' + hoyStr)
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
    .editAsText().setFontSize(10);
  body.appendParagraph('Ref. presupuesto: ' + presupuestoId)
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
    .editAsText().setFontSize(10).setForegroundColor('#555555');

  body.appendParagraph('');

  // Datos del cliente
  body.appendParagraph('CLIENTE').editAsText().setFontSize(9).setBold(true).setForegroundColor('#555555');
  body.appendParagraph(cliNombre).editAsText().setFontSize(11).setBold(true);
  if (cliCif)   body.appendParagraph('CIF/NIF: ' + cliCif).editAsText().setFontSize(10);
  if (cliDir)   body.appendParagraph(cliDir).editAsText().setFontSize(10);
  if (cliEmail) body.appendParagraph(cliEmail).editAsText().setFontSize(10);

  body.appendHorizontalRule();

  // Tabla de líneas
  body.appendParagraph('DETALLE DE SERVICIOS').editAsText().setFontSize(11).setBold(true);
  body.appendParagraph('');

  var tableData = [['Descripción', 'Cantidad', 'Precio unit.', 'Importe']];
  for (var j = 0; j < lineItems.length; j++) {
    var li = lineItems[j];
    var imp = Math.round((li.cant || 0) * (li.pu || 0) * 100) / 100;
    tableData.push([
      String(li.desc || '').substring(0, 100),
      li.cant > 0 ? String(li.cant)  : '—',
      li.pu   > 0 ? formatCurrency(li.pu) : '—',
      imp > 0 ? formatCurrency(imp) : '—'
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

  // Tabla de importes
  body.appendParagraph('IMPORTES').editAsText().setFontSize(11).setBold(true);
  body.appendParagraph('');

  var importesData = [['Base imponible', formatCurrency(subtotal)]];
  if (tipo !== 'Proforma') {
    importesData.push(['IVA ' + ivaPct + '%', formatCurrency(ivaImporte)]);
  }
  importesData.push([tipo === 'Proforma' ? 'TOTAL (sin impuestos)' : 'TOTAL A PAGAR', formatCurrency(total)]);

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
    var pagoData_fp = _buildPagoRows_(pagoSeleccionado);
    var tblPago_fp = body.appendTable(pagoData_fp);
    for (var pr_fp = 0; pr_fp < tblPago_fp.getNumRows(); pr_fp++) {
      tblPago_fp.getRow(pr_fp).editAsText().setFontSize(10);
      tblPago_fp.getRow(pr_fp).getCell(0).editAsText().setBold(true);
    }
  }

  // Aviso legal para proformas
  if (tipo === 'Proforma') {
    body.appendParagraph('');
    body.appendParagraph(
      'Este documento no tiene validez fiscal y se emite únicamente como solicitud de pago.'
    ).editAsText().setFontSize(9).setForegroundColor('#c0392b').setItalic(true);
  }

  // Pie
  body.appendParagraph('');
  body.appendHorizontalRule();
  body.appendParagraph(
    'Documento generado el ' +
    Utilities.formatDate(hoy, tz, 'dd/MM/yyyy HH:mm') +
    '   ·   ' + empresa
  ).editAsText().setFontSize(8).setForegroundColor('#888888');

  doc.saveAndClose();

  // ── 6. Convertir a PDF y guardar en Drive ────────────────────
  var docFile = DriveApp.getFileById(doc.getId());
  var pdfBlob = docFile.getAs('application/pdf');
  pdfBlob.setName(docTitulo + '.pdf');

  var _fpCarpetaKey_  = (tipo === 'Proforma') ? 'carpeta_proformas_id' : 'carpeta_facturas_id';
  var _fpCarpetaId_   = getConfig(_fpCarpetaKey_);
  var carpeta;
  if (!_fpCarpetaId_) {
    Logger.log('  ⚠️ ' + _fpCarpetaKey_ + ' no configurado. Guardando en raíz.');
    carpeta = DriveApp.getRootFolder();
  } else {
    try {
      carpeta = _getOrCreateSubcarpetaAnioMes_(_fpCarpetaId_, new Date());
    } catch (e) {
      Logger.log('  ⚠️ Error subcarpeta AÑO/MES: ' + e.message + '. Guardando en raíz.');
      carpeta = DriveApp.getRootFolder();
    }
  }
  var pdfFile = carpeta.createFile(pdfBlob);
  var pdfUrl  = pdfFile.getUrl();
  docFile.setTrashed(true);

  // ── 7. Registrar en hoja ─────────────────────────────────────
  var concepto = (tipo === 'Proforma' ? 'Proforma' : 'Factura') + ' desde presupuesto ' + presupuestoId;

  if (tipo === 'Proforma') {
    var proSheet = _ensureProformasSheet_();
    proSheet.appendRow([
      numFac.id,         // 1  ID_Proforma
      hoy,               // 2  Fecha_Emision
      numFac.serie,      // 3  Serie
      numFac.numero,     // 4  Numero
      cliId,             // 5  ID_Cliente
      cliNombre,         // 6  Nombre_Cliente
      cliEmail,          // 7  Email_Cliente
      cliDir,            // 8  Direccion_Cliente
      cliCif,            // 9  CIF_Cliente
      '',                // 10 IDs_Trabajos (vacío — origen es presupuesto)
      concepto,          // 11 Concepto
      subtotal,          // 12 Base_Imponible
      total,             // 13 Total
      'Emitida',         // 14 Estado
      pdfUrl,            // 15 URL_PDF
      '',                // 16 Fecha_Envio
      '',                // 17 Fecha_Pago
      presupuestoId,     // 18 ID_Presupuesto_Origen
      'Pendiente de cobro', // 19 Estado_Cobro
      pagoSeleccionado ? pagoSeleccionado.tipo : '', // 20 Metodo_Pago_Usado
      pagoSeleccionado ? JSON.stringify(pagoSeleccionado) : '' // 21 Datos_Pago_Usados
    ]);
  } else {
    var facSheet = getSheet('FACTURAS');
    facSheet.appendRow([
      numFac.id,         // 1  ID_Factura
      hoy,               // 2  Fecha_Emision
      numFac.serie,      // 3  Serie
      numFac.numero,     // 4  Numero
      cliId,             // 5  ID_Cliente
      cliNombre,         // 6  Nombre_Cliente
      cliEmail,          // 7  Email_Cliente
      cliDir,            // 8  Direccion_Cliente
      cliCif,            // 9  CIF_Cliente
      '',                // 10 IDs_Trabajos (vacío — origen es presupuesto)
      concepto,          // 11 Concepto
      subtotal,          // 12 Base_Imponible
      ivaPct,            // 13 IVA_Porcentaje
      ivaImporte,        // 14 IVA_Importe
      0,                 // 15 IRPF_Porcentaje
      0,                 // 16 IRPF_Importe
      total,             // 17 Total
      'Emitida',         // 18 Estado
      pdfUrl,            // 19 URL_PDF
      '',                // 20 Fecha_Envio
      '',                // 21 Fecha_Pago
      pagoSeleccionado ? pagoSeleccionado.tipo : '', // 22 Metodo_Pago_Usado
      pagoSeleccionado ? JSON.stringify(pagoSeleccionado) : '' // 23 Datos_Pago_Usados
    ]);

    var newFacRow = facSheet.getLastRow();
    // ID_Presupuesto_Origen: columna dinámica (header scan + crear si no existe)
    var presCol = _ensureFacPrespCol_(facSheet);
    facSheet.getRange(newFacRow, presCol).setValue(presupuestoId);

    // Estado Cobro: columna dinámica del módulo de cobros
    var facHdrs = facSheet.getRange(1, 1, 1, facSheet.getLastColumn()).getValues()[0];
    for (var ch = 0; ch < facHdrs.length; ch++) {
      if (String(facHdrs[ch]).trim() === 'Estado Cobro') {
        facSheet.getRange(newFacRow, ch + 1).setValue('Pendiente');
        break;
      }
    }
  }

  // ── 8. El presupuesto NO se modifica — la factura/proforma es un documento independiente ──

  logEvent('INFO', 'generateInvoiceFromPresupuestoPDF',
    numFac.id + ' | ' + tipo + ' | ' + cliNombre +
    ' | Presupuesto: ' + presupuestoId +
    ' | ' + lineItems.length + ' líneas' +
    ' | Total: ' + total.toFixed(2) + ' €',
    pdfUrl);

  return { id: numFac.id, url: pdfUrl };
}

/**
 * Busca la columna 'ID_Presupuesto_Origen' en FACTURAS por nombre de cabecera.
 * Si no existe la crea al final. Devuelve el número de columna (base 1).
 */
function _ensureFacPrespCol_(facSheet) {
  var lastCol = Math.max(facSheet.getLastColumn(), 1);
  var hdrs = facSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < hdrs.length; i++) {
    if (String(hdrs[i]).trim() === 'ID_Presupuesto_Origen') return i + 1;
  }
  var newCol = lastCol + 1;
  facSheet.getRange(1, newCol).setValue('ID_Presupuesto_Origen');
  return newCol;
}
