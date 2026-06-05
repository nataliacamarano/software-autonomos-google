// ============================================================
// PanelGastosFacturas.gs — FASE 16.3
// Subida de facturas PDF recibidas por email desde PANEL_GASTOS.
// Reutiliza el motor OCR de Gastos.gs (_extraerTextoOCR_,
// _parsearTextoTicket_, _calcularConfianzaOCR_, _determinarEstadoOCR_).
// ============================================================

// ─────────────────────────────────────────────────────────────
// mostrarModalSubirFacturaPdfGasto_
// Abre el modal HTML. Llamado desde onEditPanelGastos_ al pulsar A12.
// ─────────────────────────────────────────────────────────────

function mostrarModalSubirFacturaPdfGasto_() {
  try {
    var template = HtmlService.createTemplateFromFile('SubirFacturaPdf');
    var html = template.evaluate().setWidth(500).setHeight(360);
    SpreadsheetApp.getUi().showModalDialog(html, '📥 Subir factura PDF recibida');
  } catch (err) {
    Logger.log('ERROR mostrarModalSubirFacturaPdfGasto_: ' + err.message);
    try {
      SpreadsheetApp.getUi().alert('Error al abrir el modal:\n' + err.message);
    } catch (_) {}
  }
}

// ─────────────────────────────────────────────────────────────
// subirFacturaPdfGastoDesdePanel  — pública (llamada desde HTML)
// ─────────────────────────────────────────────────────────────

function subirFacturaPdfGastoDesdePanel(nombreArchivo, base64) {
  try {
    Logger.log('=== subirFacturaPdfGastoDesdePanel ===');
    Logger.log('  Archivo: ' + nombreArchivo);
    Logger.log('  Base64 chars: ' + base64.length +
               ' (~' + Math.round(base64.length * 0.75 / 1024) + ' KB)');

    // 1. Guardar PDF en Drive
    var bytes = Utilities.base64Decode(base64);
    Logger.log('  Bytes decodificados: ' + bytes.length);

    var blob    = Utilities.newBlob(bytes, 'application/pdf', nombreArchivo);
    var file    = guardarPdfFacturaGasto_(blob);
    var fileId  = file.getId();
    var fileUrl = file.getUrl();

    var carpetaNombre = '';
    try { carpetaNombre = file.getParents().next().getName(); } catch (_) {}
    Logger.log('  Carpeta destino: ' + carpetaNombre);
    Logger.log('  URL Drive: ' + fileUrl);

    // 2. Extraer datos del PDF via OCR (puede tardar 15-30 s)
    var datosOCR = extraerDatosFacturaPdf_(fileId);
    Logger.log('  Proveedor detectado: ' + (datosOCR.proveedor || '—'));
    Logger.log('  Fecha detectada:     ' + (datosOCR.fecha    || '—'));
    Logger.log('  Total detectado:     ' + (datosOCR.total    || '—'));
    Logger.log('  Base imponible:      ' + (datosOCR.base     || '—'));
    Logger.log('  IVA %:               ' + (datosOCR.ivaPorcentaje || '—'));
    Logger.log('  CIF/NIF:             ' + (datosOCR.cifNif   || '—'));
    Logger.log('  Nº Factura:          ' + (datosOCR.numFactura || '—'));
    Logger.log('  Confianza OCR:       ' + datosOCR.confianza);
    Logger.log('  Estado OCR:          ' + datosOCR.estadoOCR);

    // 3. Crear fila en GASTOS
    var gastoId = crearGastoDesdeFacturaPdf_(datosOCR, fileUrl, fileId, nombreArchivo);
    Logger.log('  Fila creada en GASTOS: ' + gastoId);

    // 4. Actualizar celdas de estado en la sección del panel
    _actualizarEstadoPanelFac_(nombreArchivo, gastoId, datosOCR);

    // 5. Refrescar tabla de gastos
    _actualizarPanelGastos_();

    logEvent('INFO', 'subirFacturaPdfGastoDesdePanel',
      gastoId + ' | prov=' + (datosOCR.proveedor || '—') +
      ' | total=' + (datosOCR.total || '—') +
      ' | estado=' + datosOCR.estadoOCR,
      nombreArchivo);

    return {
      success:   true,
      id:        gastoId,
      proveedor: datosOCR.proveedor || '',
      total:     datosOCR.total     || '',
      estadoOCR: datosOCR.estadoOCR,
      url:       fileUrl
    };

  } catch (err) {
    Logger.log('ERROR subirFacturaPdfGastoDesdePanel: ' + err.message);
    logEvent('ERROR', 'subirFacturaPdfGastoDesdePanel', err.message, nombreArchivo);
    throw err;
  }
}

// ─────────────────────────────────────────────────────────────
// guardarPdfFacturaGasto_
// Guarda el blob en Drive: GASTOS_FOLDER / Facturas_PDF /
// ─────────────────────────────────────────────────────────────

function guardarPdfFacturaGasto_(blob) {
  var ss = getTargetSS_();

  // Buscar carpeta raíz de gastos en DRIVE_CONFIG → CONFIG
  var gastosFolderId = '';
  var dcSh = ss.getSheetByName('DRIVE_CONFIG');
  if (dcSh) {
    var dcData = dcSh.getDataRange().getValues();
    for (var i = 0; i < dcData.length; i++) {
      if (String(dcData[i][0]).trim() === 'GASTOS_FOLDER_ID') {
        gastosFolderId = String(dcData[i][1]).trim();
        break;
      }
    }
  }
  if (!gastosFolderId) {
    gastosFolderId = String(getConfig(GAS_CFG.FOLDER_ID) || '').trim();
  }
  if (!gastosFolderId) {
    throw new Error('No se encontró carpeta de Gastos en DRIVE_CONFIG ni en CONFIG.');
  }

  var gastosFolder = DriveApp.getFolderById(gastosFolderId);

  // Subcarpeta Facturas_PDF (crear si no existe)
  var subIt = gastosFolder.getFoldersByName('Facturas_PDF');
  var carpeta = subIt.hasNext() ? subIt.next() : gastosFolder.createFolder('Facturas_PDF');

  // Eliminar archivo anterior con el mismo nombre si existe
  var existing = carpeta.getFilesByName(blob.getName());
  while (existing.hasNext()) { existing.next().setTrashed(true); }

  var file = carpeta.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file;
}

// ─────────────────────────────────────────────────────────────
// extraerDatosFacturaPdf_
// Usa el motor OCR de Gastos.gs para extraer texto del PDF
// y parsear los campos de la factura.
// ─────────────────────────────────────────────────────────────

function extraerDatosFacturaPdf_(fileId) {
  // Reutilizar funciones privadas de Gastos.gs (mismo proyecto → mismo scope)
  var ocr    = _extraerTextoOCR_(fileId);
  var datos  = _parsearTextoTicket_(ocr.texto || '');

  datos.ocrTexto  = ocr.texto  || '';
  datos.estadoOCR = _determinarEstadoOCR_(datos, ocr.exito);
  datos.confianza = _calcularConfianzaOCR_(datos, ocr.exito);
  datos.ocrExito  = ocr.exito;

  // Para facturas recibidas nunca marcar "Procesado" sin revisión manual
  if (datos.estadoOCR === 'Procesado') datos.estadoOCR = 'Revisar';

  return datos;
}

// ─────────────────────────────────────────────────────────────
// crearGastoDesdeFacturaPdf_
// Escribe la nueva fila en la hoja GASTOS con los datos extraídos.
// ─────────────────────────────────────────────────────────────

function crearGastoDesdeFacturaPdf_(datos, fileUrl, fileId, nombreArchivo) {
  var sheet   = getSheet('GASTOS');
  var gastoId = _nextGastoId_();
  var ahora   = new Date();

  // Convertir fecha string (dd/MM/yyyy) a Date
  var fechaGasto = '';
  if (datos.fecha) {
    var partes = datos.fecha.split('/');
    if (partes.length === 3) {
      try {
        fechaGasto = new Date(parseInt(partes[2], 10),
                              parseInt(partes[1], 10) - 1,
                              parseInt(partes[0], 10));
      } catch (_) { fechaGasto = ''; }
    }
  }

  var fila = new Array(HEADERS_GASTOS.length).fill('');
  fila[COL_GAS.ID_GASTO              - 1] = gastoId;
  fila[COL_GAS.FECHA_ALTA            - 1] = ahora;
  fila[COL_GAS.TIPO_GASTO            - 1] = 'Otros';
  fila[COL_GAS.FECHA_GASTO           - 1] = fechaGasto;
  fila[COL_GAS.PROVEEDOR             - 1] = datos.proveedor    || '';
  fila[COL_GAS.CIF_NIF_PROVEEDOR     - 1] = datos.cifNif       || '';
  fila[COL_GAS.NUMERO_FACTURA_TICKET - 1] = datos.numFactura   || '';
  fila[COL_GAS.CONCEPTO              - 1] = datos.concepto     || ('Factura PDF: ' + nombreArchivo);
  fila[COL_GAS.BASE_IMPONIBLE        - 1] = datos.base         ? parseFloat(datos.base)        : '';
  fila[COL_GAS.IVA_PORCENTAJE        - 1] = datos.ivaPorcentaje ? Number(datos.ivaPorcentaje)  : '';
  fila[COL_GAS.IVA_IMPORTE           - 1] = datos.ivaImporte   ? parseFloat(datos.ivaImporte)  : '';
  fila[COL_GAS.TOTAL                 - 1] = datos.total        ? parseFloat(datos.total)       : '';
  fila[COL_GAS.METODO_PAGO           - 1] = datos.metodoPago   || '';
  fila[COL_GAS.ESTADO_PAGO           - 1] = 'Pendiente';
  fila[COL_GAS.FOTO_URL              - 1] = fileUrl;
  fila[COL_GAS.ARCHIVO_ID            - 1] = fileId;
  fila[COL_GAS.OCR_TEXTO             - 1] = datos.ocrTexto     || '';
  fila[COL_GAS.CONFIANZA_OCR         - 1] = datos.confianza    || 0;
  fila[COL_GAS.REVISADO              - 1] = false;
  fila[COL_GAS.ENVIAR_GESTOR         - 1] = false;
  fila[COL_GAS.OBSERVACIONES         - 1] = 'Factura recibida por email/PDF';
  fila[COL_GAS.ESTADO_PROCESAMIENTO  - 1] = datos.estadoOCR;
  fila[COL_GAS.DOCUMENTO             - 1] = fileUrl;

  var targetRow = _nextEmptyGastosRow_(sheet);
  sheet.getRange(targetRow, 1, 1, fila.length).setValues([fila]);
  SpreadsheetApp.flush();

  Logger.log('[crearGastoDesdeFacturaPdf_] fila=' + targetRow + '  id=' + gastoId);
  return gastoId;
}

// ─────────────────────────────────────────────────────────────
// _actualizarEstadoPanelFac_
// Escribe el resultado en las celdas de estado de la sección.
// ─────────────────────────────────────────────────────────────

function _actualizarEstadoPanelFac_(nombreArchivo, gastoId, datos) {
  try {
    var ss = getTargetSS_();
    var sh = ss.getSheetByName('PANEL_GASTOS');
    if (!sh) return;

    var tz = Session.getScriptTimeZone();
    var ts = Utilities.formatDate(new Date(), tz, 'dd/MM HH:mm');

    var estado = ts + ' — ' + gastoId + ' (' + (datos.estadoOCR || 'Revisar') + ')';
    sh.getRange(PG_.FAC_ESTADO[0], PG_.FAC_ESTADO[1]).setValue(estado);

    var resultado = '';
    if (datos.proveedor) resultado += 'Prov: ' + datos.proveedor + '  ';
    if (datos.total)     resultado += 'Total: ' + datos.total + ' €  ';
    if (datos.base)      resultado += 'Base: ' + datos.base + ' €  ';
    if (!resultado)      resultado  = datos.ocrExito ? 'OCR sin datos claros — revisa manualmente'
                                                     : 'PDF sin texto legible — rellena manualmente';
    sh.getRange(PG_.FAC_RESULT[0], PG_.FAC_RESULT[1]).setValue(resultado.trim());
  } catch (e) {
    Logger.log('WARN _actualizarEstadoPanelFac_: ' + e.message);
  }
}

// ─────────────────────────────────────────────────────────────
// Wrapper público para el menú
// ─────────────────────────────────────────────────────────────

function abrirSubirFacturaPdfDesdeMenu() {
  mostrarModalSubirFacturaPdfGasto_();
}
