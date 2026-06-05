// ============================================================
// PanelGastosFacturas.gs — FASE 16.3.1
// Subida de facturas PDF recibidas por email desde PANEL_GASTOS.
// Reutiliza el motor OCR de Gastos.gs (_extraerTextoOCR_,
// _calcularConfianzaOCR_, _determinarEstadoOCR_).
// Usa _parsearTextoFacturaPdf_ (propio) en lugar de _parsearTextoTicket_.
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
// y parsear los campos con el parser de facturas (multilingual).
// ─────────────────────────────────────────────────────────────

function extraerDatosFacturaPdf_(fileId) {
  var ocr   = _extraerTextoOCR_(fileId);
  var datos = _parsearTextoFacturaPdf_(ocr.texto || '');

  datos.ocrTexto  = ocr.texto  || '';
  datos.ocrExito  = ocr.exito;
  datos.confianza = _calcularConfianzaOCR_(datos, ocr.exito);
  datos.estadoOCR = _determinarEstadoOCR_(datos, ocr.exito);

  // Facturas recibidas nunca pasan a "Procesado" sin revisión manual
  if (datos.estadoOCR === 'Procesado') datos.estadoOCR = 'Revisar';

  // Si no se extrajo ningún campo significativo → marcar "OCR revisar"
  var tieneDatos = datos.proveedor || datos.total || datos.fecha || datos.numFactura;
  if (!tieneDatos) {
    datos.estadoOCR     = 'OCR revisar';
    datos.observaciones = 'Revisar datos extraídos del PDF';
  }

  return datos;
}

// ─────────────────────────────────────────────────────────────
// _parsearTextoFacturaPdf_
// Parser multilingual (ES/EN) para facturas PDF recibidas.
// Devuelve el mismo objeto que _parsearTextoTicket_ para ser
// compatible con crearGastoDesdeFacturaPdf_ y _calcularConfianzaOCR_.
// ─────────────────────────────────────────────────────────────

function _parsearTextoFacturaPdf_(texto) {
  var datos = {
    proveedor: '', fecha: '', total: '', base: '',
    ivaPorcentaje: '', ivaImporte: '', cifNif: '',
    numFactura: '', concepto: '', metodoPago: '',
    observaciones: '', _reglas: []
  };

  if (!texto || !texto.trim()) return datos;

  var lineas = texto.split('\n');

  // ── Palabras/frases a excluir de la detección de proveedor ──
  var EXCLUIR_PROV = [
    /need help/i, /call free/i, /account name/i, /account number/i,
    /customer/i, /bill number/i, /payment due/i, /invoice number/i,
    /statement/i, /^date[:\s]/i, /^dear\b/i, /^hello\b/i,
    /^page \d/i, /^invoice date/i, /^bill date/i, /^amount due/i,
    /^total due/i, /^grand total/i, /^balance due/i,
    /^ref(erence)?[:\s]/i, /^vat reg/i, /^registered/i,
    /^tel[:\s]/i, /^fax[:\s]/i, /^email[:\s]/i,
    /^www\./i, /^http/i, /\.co\.uk/i, /\.com/i,
    /^\d[\d\s\-\/]+$/, // solo números/fechas
    /^[+]?\d{6,}/      // teléfonos
  ];

  // Meses EN → número
  var MESES_EN = {
    jan:1,feb:2,mar:3,apr:4,may:5,jun:6,
    jul:7,aug:8,sep:9,oct:10,nov:11,dec:12
  };

  // ── Helper: normalizar número ──
  function normNum(s) {
    if (!s) return '';
    s = String(s).trim().replace(/[£€$\s]/g, '').replace(',', '.');
    return isNaN(parseFloat(s)) ? '' : s;
  }

  // ── Helper: parsear fecha inglesa "5 Apr 26" / "5 April 2026" ──
  function parseFechaEN(s) {
    var m = s.match(/(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{2,4})/);
    if (!m) return '';
    var mes = MESES_EN[m[2].toLowerCase().substring(0, 3)];
    if (!mes) return '';
    var anio = parseInt(m[3], 10);
    if (anio < 100) anio += 2000;
    var dia = parseInt(m[1], 10);
    return (dia < 10 ? '0' : '') + dia + '/' +
           (mes < 10 ? '0' : '') + mes + '/' + anio;
  }

  // ── Helper: parsear fecha española "dd/MM/yyyy" o "dd-MM-yyyy" ──
  function parseFechaES(s) {
    var m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (!m) return '';
    var anio = parseInt(m[3], 10);
    if (anio < 100) anio += 2000;
    var dia = m[1], mes = m[2];
    return (dia.length < 2 ? '0' : '') + dia + '/' +
           (mes.length < 2 ? '0' : '') + mes + '/' + anio;
  }

  // ── 1. PROVEEDOR ── primera línea no vacía que no sea ruido ──
  for (var i = 0; i < Math.min(lineas.length, 25); i++) {
    var ln = lineas[i].trim();
    if (!ln || ln.length < 2) continue;
    var esRuido = EXCLUIR_PROV.some(function(re) { return re.test(ln); });
    if (!esRuido && ln.length >= 2 && ln.length <= 60) {
      datos.proveedor = ln;
      datos._reglas.push('proveedor → línea ' + i + ': "' + ln + '"');
      break;
    }
  }

  // ── 2. NÚMERO DE FACTURA ──
  for (var i = 0; i < lineas.length; i++) {
    var ln = lineas[i].trim();
    // EN: "Bill number: 12345" / "Invoice number: INV-001"
    var mBill = ln.match(/(?:bill|invoice|statement)\s+(?:number|no\.?|ref)[:\s]+([A-Z0-9\-\/]+)/i);
    if (mBill) {
      datos.numFactura = mBill[1].trim();
      datos._reglas.push('numFactura → "' + ln + '"');
      break;
    }
    // ES: "Nº Factura: F2024-001" / "Factura: F2024-001"
    var mFac = ln.match(/(?:n[úu]mero\s+)?factura[:\s]+([A-Z0-9\-\/]+)/i);
    if (mFac) {
      datos.numFactura = mFac[1].trim();
      datos._reglas.push('numFactura ES → "' + ln + '"');
      break;
    }
  }

  // ── 3. FECHA ──
  for (var i = 0; i < lineas.length; i++) {
    var ln = lineas[i].trim();
    // EN: "Bill date: 5 Apr 26" / "Invoice date: 5 Apr 2026"
    var mDateEN = ln.match(/(?:bill|invoice|statement|issue)\s+date[:\s]+(.+)/i);
    if (mDateEN) {
      var f = parseFechaEN(mDateEN[1]) || parseFechaES(mDateEN[1]);
      if (f) { datos.fecha = f; datos._reglas.push('fecha EN label → "' + ln + '"'); break; }
    }
    // ES: "Fecha: 05/06/2024" / "Fecha factura:"
    var mDateES = ln.match(/fecha(?:\s+(?:factura|emisi[oó]n))?[:\s]+(.+)/i);
    if (mDateES) {
      var f = parseFechaES(mDateES[1]) || parseFechaEN(mDateES[1]);
      if (f) { datos.fecha = f; datos._reglas.push('fecha ES label → "' + ln + '"'); break; }
    }
  }
  // Fallback: cualquier línea con patrón de fecha
  if (!datos.fecha) {
    for (var i = 0; i < Math.min(lineas.length, 30); i++) {
      var ln = lineas[i].trim();
      var f = parseFechaES(ln) || parseFechaEN(ln);
      if (f) { datos.fecha = f; datos._reglas.push('fecha fallback → "' + ln + '"'); break; }
    }
  }

  // ── 4. IVA ESPECIAL: "This is not a VAT Invoice" ──
  var noVat = lineas.some(function(l) {
    return /not\s+a\s+vat\s+invoice/i.test(l) || /no\s+vat/i.test(l);
  });
  if (noVat) {
    datos.ivaPorcentaje = '0';
    datos._reglas.push('IVA → "not a VAT invoice" detectado → ivaPorcentaje=0');
  }

  // ── 5. TOTAL / IMPORTE ──
  for (var i = 0; i < lineas.length; i++) {
    var ln = lineas[i].trim();
    // EN: "Total due: £27.30" / "Amount due: €55.00" / "Grand total"
    var mTotalEN = ln.match(/(?:total\s+due|amount\s+due|grand\s+total|balance\s+due)[:\s]+[£€$]?([\d,\.]+)/i);
    if (mTotalEN) {
      var v = normNum(mTotalEN[1]);
      if (v) { datos.total = v; datos._reglas.push('total EN label → "' + ln + '"'); break; }
    }
    // ES: "Total: 55,00 €" / "Importe total:"
    var mTotalES = ln.match(/(?:total|importe\s+total|total\s+factura)[:\s]+[€£$]?([\d,\.]+)/i);
    if (mTotalES) {
      var v = normNum(mTotalES[1]);
      if (v) { datos.total = v; datos._reglas.push('total ES label → "' + ln + '"'); break; }
    }
    // Importe con símbolo: "£27.30" o "€55,00" solos en la línea
    var mSimbolo = ln.match(/^[£€]\s*([\d,\.]+)$/);
    if (mSimbolo && !datos.total) {
      var v = normNum(mSimbolo[1]);
      if (v) { datos.total = v; datos._reglas.push('total símbolo → "' + ln + '"'); }
    }
  }

  // ── 6. BASE IMPONIBLE ──
  for (var i = 0; i < lineas.length; i++) {
    var ln = lineas[i].trim();
    var mBase = ln.match(/(?:base\s+imponible|subtotal|net\s+(?:total|amount))[:\s]+[€£$]?([\d,\.]+)/i);
    if (mBase) {
      var v = normNum(mBase[1]);
      if (v) { datos.base = v; datos._reglas.push('base → "' + ln + '"'); break; }
    }
  }
  // Si no hay VAT → base = total
  if (noVat && datos.total && !datos.base) {
    datos.base = datos.total;
    datos._reglas.push('base = total (no VAT invoice)');
  }

  // ── 7. IVA IMPORTE ──
  if (!noVat) {
    for (var i = 0; i < lineas.length; i++) {
      var ln = lineas[i].trim();
      var mIva = ln.match(/(?:iva|vat)[:\s]+(?:\d+%\s+)?[€£$]?([\d,\.]+)/i);
      if (mIva) {
        var v = normNum(mIva[1]);
        if (v) { datos.ivaImporte = v; datos._reglas.push('ivaImporte → "' + ln + '"'); break; }
      }
    }
    // IVA porcentaje
    for (var i = 0; i < lineas.length; i++) {
      var ln = lineas[i].trim();
      var mPct = ln.match(/(?:iva|vat)\s+(\d+)\s*%/i);
      if (mPct && !datos.ivaPorcentaje) {
        datos.ivaPorcentaje = mPct[1];
        datos._reglas.push('ivaPorcentaje → "' + ln + '"');
        break;
      }
    }
  }

  // ── 8. CIF/NIF ──
  for (var i = 0; i < lineas.length; i++) {
    var ln = lineas[i].trim();
    // ES: "CIF: B12345678" / "NIF: 12345678A"
    var mCif = ln.match(/(?:cif|nif|vat\s+(?:no|number|reg)[.:\s]*)[:\s]+([A-Z0-9]{5,15})/i);
    if (mCif) {
      datos.cifNif = mCif[1].trim().toUpperCase();
      datos._reglas.push('cifNif → "' + ln + '"');
      break;
    }
    // UK VAT: "GB123456789"
    var mVatUK = ln.match(/\b(GB\d{9})\b/);
    if (mVatUK) {
      datos.cifNif = mVatUK[1];
      datos._reglas.push('cifNif GB → "' + ln + '"');
      break;
    }
  }

  // ── 9. MÉTODO DE PAGO ──
  for (var i = 0; i < lineas.length; i++) {
    var ln = lineas[i].trim();
    if (/direct\s+debit/i.test(ln)) {
      datos.metodoPago = 'Domiciliación';
      datos._reglas.push('metodoPago → direct debit');
      break;
    }
    if (/tarjeta|credit\s+card|debit\s+card/i.test(ln)) {
      datos.metodoPago = 'Tarjeta';
      datos._reglas.push('metodoPago → tarjeta/card');
      break;
    }
    if (/transf|transfer|wire/i.test(ln)) {
      datos.metodoPago = 'Transferencia';
      datos._reglas.push('metodoPago → transferencia');
      break;
    }
  }

  // ── 10. CONCEPTO ── patrón "X: dd Mon yy to dd Mon yy" o nombre archivo ──
  for (var i = 0; i < lineas.length; i++) {
    var ln = lineas[i].trim();
    if (/\d{1,2}\s+[A-Za-z]{3}\s+\d{2,4}\s+to\s+\d{1,2}\s+[A-Za-z]{3}/i.test(ln)) {
      datos.concepto = ln.replace(/^[^:]+:\s*/, '').trim() || ln;
      datos._reglas.push('concepto → "' + ln + '"');
      break;
    }
    // ES: "Descripción:" / "Concepto:"
    var mConc = ln.match(/^(?:descripci[oó]n|concepto|detail|description)[:\s]+(.+)/i);
    if (mConc && mConc[1].trim().length > 3) {
      datos.concepto = mConc[1].trim();
      datos._reglas.push('concepto label → "' + ln + '"');
      break;
    }
  }

  // ── 11. OBSERVACIONES ──
  if (noVat) {
    datos.observaciones = 'No VAT invoice';
  }

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
  var obs = datos.observaciones || 'Factura recibida por email/PDF';
  fila[COL_GAS.OBSERVACIONES         - 1] = obs;
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
// diagnosticarLecturaPdfFacturaGasto  — pública, llamar desde el editor
// Uso: diagnosticarLecturaPdfFacturaGasto("DRIVE_FILE_ID")
// Muestra en Logger: primeras 40 líneas del OCR + todos los campos
// detectados con la regla que los originó.
// ─────────────────────────────────────────────────────────────

function diagnosticarLecturaPdfFacturaGasto(fileId) {
  Logger.log('=== DIAGNÓSTICO FACTURA PDF ===');
  Logger.log('File ID: ' + fileId);

  var ocr = _extraerTextoOCR_(fileId);
  Logger.log('OCR exitoso: ' + ocr.exito);
  Logger.log('Chars texto: ' + (ocr.texto || '').length);
  Logger.log('');

  var lineas = (ocr.texto || '').split('\n');
  Logger.log('── Primeras 40 líneas del texto OCR ──');
  for (var i = 0; i < Math.min(lineas.length, 40); i++) {
    Logger.log('  [' + i + '] ' + lineas[i]);
  }
  Logger.log('');

  var datos = _parsearTextoFacturaPdf_(ocr.texto || '');
  Logger.log('── Campos extraídos ──');
  Logger.log('  proveedor:     ' + (datos.proveedor     || '(vacío)'));
  Logger.log('  fecha:         ' + (datos.fecha         || '(vacío)'));
  Logger.log('  numFactura:    ' + (datos.numFactura    || '(vacío)'));
  Logger.log('  total:         ' + (datos.total         || '(vacío)'));
  Logger.log('  base:          ' + (datos.base          || '(vacío)'));
  Logger.log('  ivaPorcentaje: ' + (datos.ivaPorcentaje || '(vacío)'));
  Logger.log('  ivaImporte:    ' + (datos.ivaImporte    || '(vacío)'));
  Logger.log('  cifNif:        ' + (datos.cifNif        || '(vacío)'));
  Logger.log('  concepto:      ' + (datos.concepto      || '(vacío)'));
  Logger.log('  metodoPago:    ' + (datos.metodoPago    || '(vacío)'));
  Logger.log('  observaciones: ' + (datos.observaciones || '(vacío)'));
  Logger.log('');
  Logger.log('── Reglas aplicadas ──');
  (datos._reglas || []).forEach(function(r) { Logger.log('  • ' + r); });
  Logger.log('');

  var confianza = _calcularConfianzaOCR_(datos, ocr.exito);
  var estado    = _determinarEstadoOCR_(datos, ocr.exito);
  Logger.log('  Confianza OCR: ' + confianza);
  Logger.log('  Estado OCR:    ' + estado);
  Logger.log('=== FIN DIAGNÓSTICO ===');
}

// ─────────────────────────────────────────────────────────────
// Wrapper público para el menú
// ─────────────────────────────────────────────────────────────

function abrirSubirFacturaPdfDesdeMenu() {
  mostrarModalSubirFacturaPdfGasto_();
}
