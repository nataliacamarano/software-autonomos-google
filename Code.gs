// ============================================================
// Code.gs — Inicialización del sistema y punto de entrada web
// ============================================================

/**
 * SETUP — Ejecutar UNA SOLA VEZ desde el editor de Apps Script
 * para crear todas las hojas, configuración y carpetas de Drive.
 */
function setup() {
  try {
    Logger.log('=== INICIO SETUP ===');

    createAllSheets_();
    Logger.log('✓ Hojas creadas');

    initConfig();
    Logger.log('✓ CONFIG inicializada');

    createDriveFolders_();
    Logger.log('✓ Carpetas Drive creadas');

    applySheetFormatting_();
    Logger.log('✓ Formato aplicado');

    Logger.log('=== SETUP COMPLETADO ===');
    Logger.log('Siguiente paso: abre la hoja CONFIG y completa los datos de tu empresa.');
  } catch (e) {
    logEvent('ERROR', 'setup', e.message, e.stack);
    Logger.log('❌ ERROR en setup: ' + e.message);
    throw e;
  }
}

// ---- Creación de hojas ----

function createAllSheets_() {
  const ss = getTargetSS_();

  Object.entries(SHEET_HEADERS).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  });

  // Eliminar hojas vacías creadas por defecto por Google Sheets
  ['Sheet1', 'Hoja 1', 'Hoja1'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s && ss.getSheets().length > 1) {
      try { ss.deleteSheet(s); } catch (_) {}
    }
  });
}

// ---- Estructura de carpetas en Drive ----

function createDriveFolders_() {
  const year = getCurrentYear();
  const root = DriveApp.getRootFolder();

  const main = getOrCreateDriveFolder_(root, 'Álvaro — Gestión');

  const facturas = getOrCreateDriveFolder_(main, 'Facturas');
  getOrCreateDriveFolder_(facturas, year);

  const trabajos = getOrCreateDriveFolder_(main, 'Trabajos');
  getOrCreateDriveFolder_(trabajos, year);

  setConfig('carpeta_facturas_id', facturas.getId());
  setConfig('carpeta_trabajos_id', trabajos.getId());

  Logger.log(`Drive Facturas: ${facturas.getId()} | Trabajos: ${trabajos.getId()}`);
}

function getOrCreateDriveFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// ---- Formato de hojas ----

function applySheetFormatting_() {
  const ss = getTargetSS_();

  // Cabeceras
  Object.entries(SHEET_HEADERS).forEach(([name, headers]) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    applyHeaderStyle_(sheet, headers.length);
  });

  // Anchos de columna específicos por hoja
  const colWidths = {
    CLIENTES:    { 1: 100, 3: 150, 4: 150, 5: 180, 6: 200, 8: 250 },
    OPERARIOS:   { 1: 110, 3: 150, 4: 150, 6: 130, 7: 150 },
    TRABAJOS:    { 1: 110, 4: 160, 6: 160, 7: 250, 9: 200, 12: 110 },
    FACTURAS:    { 1: 140, 6: 160, 8: 220, 11: 250, 19: 300 },
    PENDIENTES:  { 3: 160, 4: 160, 5: 220 },
    COMPLETADOS: { 3: 160, 4: 160, 5: 220 },
    LOGS:        { 1: 160, 4: 280, 5: 300 }
  };

  Object.entries(colWidths).forEach(([sheetName, widths]) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    Object.entries(widths).forEach(([col, width]) => {
      sheet.setColumnWidth(Number(col), width);
    });
  });

  SpreadsheetApp.flush();
}

// ---- PANEL_GASTOS — funciones de menú ----

function crearPanelGastos() {
  _crearPanelGastos_();
}

function actualizarPanelGastos() {
  _actualizarPanelGastos_();
}

function diagnosticarPanelGastos() {
  _diagnosticarPanelGastos_();
}

function testDiagnosticoGastos() {
  _diagnosticarPanelGastos_();
}

function obtenerGasto(id) {
  Logger.log('obtenerGasto() id="' + id + '"');
  var gasSheet = getSheet('GASTOS');
  var row = findRowById_(gasSheet, id);
  if (!row) { Logger.log('obtenerGasto: no encontrado'); return null; }
  var data = gasSheet.getRange(row, 1, 1, HEADERS_GASTOS.length).getValues()[0];
  var tz   = Session.getScriptTimeZone();
  var fGas = data[COL_GAS.FECHA_GASTO - 1];
  return {
    id:            String(data[COL_GAS.ID_GASTO              - 1]).trim(),
    fechaGasto:    (fGas instanceof Date && !isNaN(fGas.getTime()))
                     ? Utilities.formatDate(fGas, tz, 'yyyy-MM-dd') : '',
    tipoGasto:     String(data[COL_GAS.TIPO_GASTO            - 1]).trim(),
    proveedor:     String(data[COL_GAS.PROVEEDOR             - 1]).trim(),
    cifNif:        String(data[COL_GAS.CIF_NIF_PROVEEDOR     - 1]).trim(),
    numFactura:    String(data[COL_GAS.NUMERO_FACTURA_TICKET  - 1]).trim(),
    concepto:      String(data[COL_GAS.CONCEPTO              - 1]).trim(),
    categoria:     String(data[COL_GAS.CATEGORIA             - 1]).trim(),
    baseImponible: data[COL_GAS.BASE_IMPONIBLE - 1] !== '' ? data[COL_GAS.BASE_IMPONIBLE - 1] : '',
    ivaPorcentaje: data[COL_GAS.IVA_PORCENTAJE - 1] !== '' ? data[COL_GAS.IVA_PORCENTAJE - 1] : '',
    ivaImporte:    data[COL_GAS.IVA_IMPORTE    - 1] !== '' ? data[COL_GAS.IVA_IMPORTE    - 1] : '',
    total:         data[COL_GAS.TOTAL          - 1] !== '' ? data[COL_GAS.TOTAL          - 1] : '',
    metodoPago:    String(data[COL_GAS.METODO_PAGO           - 1]).trim(),
    estadoPago:    String(data[COL_GAS.ESTADO_PAGO           - 1]).trim(),
    observaciones: String(data[COL_GAS.OBSERVACIONES         - 1]).trim(),
    estadoProc:    String(data[COL_GAS.ESTADO_PROCESAMIENTO  - 1]).trim()
  };
}

function guardarCambiosGasto(data) {
  try {
    Logger.log('guardarCambiosGasto() id="' + data.id + '"');
    var gasSheet = getSheet('GASTOS');
    var row = findRowById_(gasSheet, data.id);
    if (!row) return { success: false, error: 'ID no encontrado: ' + data.id };
    var fechaGasto = '';
    if (data.fechaGasto) {
      var d = new Date(data.fechaGasto);
      if (!isNaN(d.getTime())) fechaGasto = d;
    }
    gasSheet.getRange(row, COL_GAS.FECHA_GASTO          ).setValue(fechaGasto);
    gasSheet.getRange(row, COL_GAS.TIPO_GASTO           ).setValue(data.tipoGasto     || '');
    gasSheet.getRange(row, COL_GAS.PROVEEDOR            ).setValue(data.proveedor     || '');
    gasSheet.getRange(row, COL_GAS.CIF_NIF_PROVEEDOR    ).setValue(data.cifNif        || '');
    gasSheet.getRange(row, COL_GAS.NUMERO_FACTURA_TICKET).setValue(data.numFactura    || '');
    gasSheet.getRange(row, COL_GAS.CONCEPTO             ).setValue(data.concepto      || '');
    gasSheet.getRange(row, COL_GAS.CATEGORIA            ).setValue(data.categoria     || '');
    gasSheet.getRange(row, COL_GAS.BASE_IMPONIBLE       ).setValue(parseFloat(data.baseImponible) || '');
    gasSheet.getRange(row, COL_GAS.IVA_PORCENTAJE       ).setValue(parseFloat(data.ivaPorcentaje) || '');
    gasSheet.getRange(row, COL_GAS.IVA_IMPORTE          ).setValue(parseFloat(data.ivaImporte)    || '');
    gasSheet.getRange(row, COL_GAS.TOTAL                ).setValue(parseFloat(data.total)         || '');
    gasSheet.getRange(row, COL_GAS.METODO_PAGO          ).setValue(data.metodoPago    || '');
    gasSheet.getRange(row, COL_GAS.ESTADO_PAGO          ).setValue(data.estadoPago    || '');
    gasSheet.getRange(row, COL_GAS.OBSERVACIONES        ).setValue(data.observaciones || '');
    gasSheet.getRange(row, COL_GAS.ESTADO_PROCESAMIENTO ).setValue(data.estadoProc    || '');
    SpreadsheetApp.flush();
    logEvent('INFO', 'guardarCambiosGasto', data.id + ' actualizado', '');
    return { success: true };
  } catch (err) {
    logEvent('ERROR', 'guardarCambiosGasto', err.message, err.stack);
    return { success: false, error: err.message };
  }
}

// ---- FASE 6 — Envío de gastos al gestor ----

function enviarGastosAlGestor() {
  var ui = SpreadsheetApp.getUi();

  // 1. Email del gestor
  var emailGestor = getConfig('email_gestor') || '';
  if (!emailGestor) {
    var res = ui.prompt('Email del gestor',
      'No hay email configurado.\nIntroduce el email del gestor/asesor:', ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return;
    emailGestor = res.getResponseText().trim();
    if (!emailGestor) { ui.alert('Email vacío. Operación cancelada.'); return; }
    setConfig('email_gestor', emailGestor);
  }

  // 2. Leer gastos marcados con ENVIAR_GESTOR = true
  var gasSheet  = getSheet('GASTOS');
  var lastRow   = gasSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('La hoja GASTOS está vacía.'); return;
  }
  var allData = gasSheet.getRange(2, 1, lastRow - 1, HEADERS_GASTOS.length).getValues();
  var gastos  = allData.filter(function(row) {
    var id  = String(row[COL_GAS.ID_GASTO      - 1]).trim();
    var env = row[COL_GAS.ENVIAR_GESTOR - 1];
    return id.indexOf('GAS-') === 0 && (env === true || String(env).toLowerCase() === 'true');
  });

  if (gastos.length === 0) {
    ui.alert('Sin gastos para enviar',
      'No hay gastos con "Enviar gestor" marcado.\n'
      + 'Márcalos en PANEL_GASTOS (columna K) primero.', ui.ButtonSet.OK);
    return;
  }

  // 3. Calcular totales
  var totalImporte = 0, totalBase = 0, totalIva = 0;
  gastos.forEach(function(row) {
    totalImporte += parseFloat(row[COL_GAS.TOTAL          - 1]) || 0;
    totalBase    += parseFloat(row[COL_GAS.BASE_IMPONIBLE - 1]) || 0;
    totalIva     += parseFloat(row[COL_GAS.IVA_IMPORTE    - 1]) || 0;
  });

  // 4. Confirmación
  var conf = ui.alert('Confirmar envío al gestor',
    gastos.length + ' gastos  |  Total: ' + totalImporte.toFixed(2) + ' €\n'
    + 'Destinatario: ' + emailGestor + '\n\n¿Enviar ahora?',
    ui.ButtonSet.YES_NO);
  if (conf !== ui.Button.YES) return;

  // 5. Construir y enviar email
  var tz      = Session.getScriptTimeZone();
  var empresa = getConfig('empresa_nombre') || 'Empresa';
  var fechaStr = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
  var subject  = 'Gastos para revisión — ' + fechaStr
                 + ' — ' + gastos.length + ' registros — '
                 + totalImporte.toFixed(2) + ' €';
  var htmlBody = _construirEmailGestor_(gastos, totalBase, totalIva, totalImporte, empresa, tz);
  GmailApp.sendEmail(emailGestor, subject, '', { htmlBody: htmlBody });

  // 6. Marcar como enviados en GASTOS
  var fechaEnvio = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  gastos.forEach(function(row) {
    var id    = String(row[COL_GAS.ID_GASTO - 1]).trim();
    var fila  = findRowById_(gasSheet, id);
    if (!fila) return;
    gasSheet.getRange(fila, COL_GAS.ENVIAR_GESTOR).setValue(false);
    var obs    = String(gasSheet.getRange(fila, COL_GAS.OBSERVACIONES).getValue() || '').trim();
    var newObs = obs ? obs + ' | Enviado gestor: ' + fechaEnvio : 'Enviado gestor: ' + fechaEnvio;
    gasSheet.getRange(fila, COL_GAS.OBSERVACIONES).setValue(newObs);
  });
  SpreadsheetApp.flush();

  logEvent('INFO', 'enviarGastosAlGestor',
    gastos.length + ' gastos enviados a ' + emailGestor + '  total=' + totalImporte.toFixed(2), '');

  // 7. Refrescar panel
  _actualizarPanelGastos_();

  ui.alert('✅ Email enviado',
    'Destinatario: ' + emailGestor + '\n'
    + gastos.length + ' gastos  |  ' + totalImporte.toFixed(2) + ' €\n\n'
    + 'Los gastos han sido desmarcados de "Enviar gestor".', ui.ButtonSet.OK);
}

function _construirEmailGestor_(gastos, totalBase, totalIva, totalImporte, empresa, tz) {
  var fecha = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');

  var filas = gastos.map(function(row) {
    var fGas     = row[COL_GAS.FECHA_GASTO - 1];
    var fechaGas = (fGas instanceof Date && !isNaN(fGas.getTime()))
                   ? Utilities.formatDate(fGas, tz, 'dd/MM/yyyy') : '—';
    var total    = parseFloat(row[COL_GAS.TOTAL          - 1]) || 0;
    var base     = parseFloat(row[COL_GAS.BASE_IMPONIBLE - 1]) || 0;
    var iva      = parseFloat(row[COL_GAS.IVA_IMPORTE    - 1]) || 0;
    var ep       = String(row[COL_GAS.ESTADO_PAGO        - 1]).trim();
    var colorEp  = ep === 'Pagado' ? '#137333' : '#c5221f';
    var numFac   = String(row[COL_GAS.NUMERO_FACTURA_TICKET - 1]).trim();
    return '<tr>'
      + td(String(row[COL_GAS.ID_GASTO   - 1]).trim())
      + td(fechaGas)
      + td(String(row[COL_GAS.TIPO_GASTO - 1]).trim())
      + td(String(row[COL_GAS.PROVEEDOR  - 1]).trim())
      + td(numFac || '—')
      + td(String(row[COL_GAS.CONCEPTO   - 1]).trim())
      + tdr(base.toFixed(2)  + ' €')
      + tdr(iva.toFixed(2)   + ' €')
      + '<td style="' + TD + 'text-align:right;font-weight:600">'  + total.toFixed(2) + ' €</td>'
      + '<td style="' + TD + 'color:' + colorEp + '">' + ep + '</td>'
      + '</tr>';
  }).join('');

  return '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;color:#202124;max-width:960px;margin:0 auto;padding:24px">'
    + '<h2 style="color:#1a73e8;border-bottom:2px solid #1a73e8;padding-bottom:8px;margin-bottom:4px">'
    + '📋 Gastos para revisión — ' + empresa + '</h2>'
    + '<p style="color:#5f6368;margin:0 0 24px">Generado el ' + fecha + '</p>'

    // Resumen
    + '<table style="background:#f8f9fa;border-radius:6px;padding:16px 24px;margin-bottom:28px">'
    + resRow('Nº gastos',      gastos.length)
    + resRow('Base imponible', totalBase.toFixed(2)    + ' €')
    + resRow('IVA soportado',  totalIva.toFixed(2)     + ' €')
    + resRowBig('TOTAL',       totalImporte.toFixed(2) + ' €')
    + '</table>'

    // Tabla
    + '<table style="width:100%;border-collapse:collapse;font-size:12px">'
    + '<thead><tr style="background:#202124;color:#fff">'
    + th('ID') + th('Fecha') + th('Tipo') + th('Proveedor') + th('Nº Factura')
    + th('Concepto') + thr('Base') + thr('IVA') + thr('Total') + th('Pago')
    + '</tr></thead><tbody>' + filas + '</tbody></table>'

    + '<p style="color:#9aa0a6;font-size:11px;margin-top:24px">'
    + 'Email generado automáticamente por el sistema de gestión de ' + empresa + '.</p>'
    + '</body></html>';
}

// helpers de construcción HTML
var TD  = 'padding:6px 10px;border-bottom:1px solid #e8eaed;';
function td(v)  { return '<td style="' + TD + '">' + (v || '—') + '</td>'; }
function tdr(v) { return '<td style="' + TD + 'text-align:right">' + v + '</td>'; }
function th(v)  { return '<th style="padding:8px 10px;text-align:left">' + v + '</th>'; }
function thr(v) { return '<th style="padding:8px 10px;text-align:right">' + v + '</th>'; }
function resRow(label, val) {
  return '<tr><td style="padding:3px 24px 3px 0;font-weight:600">' + label
       + '</td><td style="padding:3px 0">' + val + '</td></tr>';
}
function resRowBig(label, val) {
  return '<tr><td style="padding:6px 24px 3px 0;font-weight:700;font-size:15px">' + label
       + '</td><td style="padding:6px 0;font-weight:700;font-size:15px;color:#1a73e8">'
       + val + '</td></tr>';
}

// ---- Diagnóstico de columnas GASTOS -------------------------

function diagnosticarColumnasGastos() {
  var ui      = SpreadsheetApp.getUi();
  var gasSheet;
  try { gasSheet = getSheet('GASTOS'); } catch (e) {
    ui.alert('⚠ No se pudo abrir la hoja GASTOS:\n' + e.message); return;
  }

  var lastRow = gasSheet.getLastRow();
  if (lastRow < 1) { ui.alert('La hoja GASTOS está vacía.'); return; }

  // ── 1. Leer cabecera real ────────────────────────────────────
  var lastCol   = gasSheet.getLastColumn();
  var headerRow = gasSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  var lines = [];
  lines.push('=== DIAGNÓSTICO COLUMNAS GASTOS ===');
  lines.push('Total columnas detectadas: ' + lastCol);
  lines.push('Total filas con datos: ' + (lastRow - 1));
  lines.push('');

  lines.push('── HEADERS REALES (col base-1 → nombre) ──');
  headerRow.forEach(function(h, idx) {
    var marker = '';
    if (idx + 1 === COL_GAS.BASE_IMPONIBLE) marker = ' ← COL_GAS.BASE_IMPONIBLE';
    if (idx + 1 === COL_GAS.IVA_IMPORTE)    marker = ' ← COL_GAS.IVA_IMPORTE';
    if (idx + 1 === COL_GAS.TOTAL)          marker = ' ← COL_GAS.TOTAL';
    lines.push('  Col ' + (idx + 1) + ': "' + h + '"' + marker);
  });

  // ── 2. Buscar Total/Base/IVA por nombre ──────────────────────
  var colTotalNombre   = -1;
  var colBaseNombre    = -1;
  var colIvaNombre     = -1;
  headerRow.forEach(function(h, idx) {
    var hn = String(h).trim().toLowerCase();
    if (hn === 'total')          colTotalNombre = idx + 1;
    if (hn === 'base_imponible' || hn === 'base imponible') colBaseNombre = idx + 1;
    if (hn === 'iva_importe'    || hn === 'iva importe')    colIvaNombre  = idx + 1;
  });

  lines.push('');
  lines.push('── BÚSQUEDA POR NOMBRE EXACTO ──');
  lines.push('  "Total"         → col por nombre: ' + (colTotalNombre > 0 ? colTotalNombre : 'NO ENCONTRADA'));
  lines.push('  "Base_Imponible"→ col por nombre: ' + (colBaseNombre  > 0 ? colBaseNombre  : 'NO ENCONTRADA'));
  lines.push('  "IVA_Importe"   → col por nombre: ' + (colIvaNombre   > 0 ? colIvaNombre   : 'NO ENCONTRADA'));
  lines.push('');
  lines.push('── CONSTANTES COL_GAS vs. NOMBRE REAL ──');
  lines.push('  COL_GAS.BASE_IMPONIBLE = ' + COL_GAS.BASE_IMPONIBLE +
             ' → header real: "' + (headerRow[COL_GAS.BASE_IMPONIBLE - 1] || 'FUERA DE RANGO') + '"');
  lines.push('  COL_GAS.IVA_IMPORTE    = ' + COL_GAS.IVA_IMPORTE +
             ' → header real: "' + (headerRow[COL_GAS.IVA_IMPORTE - 1]    || 'FUERA DE RANGO') + '"');
  lines.push('  COL_GAS.TOTAL          = ' + COL_GAS.TOTAL +
             ' → header real: "' + (headerRow[COL_GAS.TOTAL - 1]          || 'FUERA DE RANGO') + '"');

  // ── 3. Leer primeras 5 filas de datos ───────────────────────
  var nMuestra = Math.min(5, lastRow - 1);
  if (nMuestra > 0) {
    var muestra = gasSheet.getRange(2, 1, nMuestra, lastCol).getValues();

    lines.push('');
    lines.push('── PRIMERAS ' + nMuestra + ' FILAS — COL_GAS.TOTAL (col ' + COL_GAS.TOTAL + ', idx ' + (COL_GAS.TOTAL - 1) + ') ──');
    muestra.forEach(function(row, i) {
      var v = row[COL_GAS.TOTAL - 1];
      lines.push('  Fila ' + (i + 2) + ': valor="' + v + '"  tipo=' + typeof v +
                 '  parseFloat=' + (parseFloat(v) || 0));
    });

    if (colTotalNombre > 0 && colTotalNombre !== COL_GAS.TOTAL) {
      lines.push('');
      lines.push('── PRIMERAS ' + nMuestra + ' FILAS — "Total" detectado por nombre (col ' + colTotalNombre + ', idx ' + (colTotalNombre - 1) + ') ──');
      muestra.forEach(function(row, i) {
        var v = row[colTotalNombre - 1];
        lines.push('  Fila ' + (i + 2) + ': valor="' + v + '"  tipo=' + typeof v +
                   '  parseFloat=' + (parseFloat(v) || 0));
      });
    } else if (colTotalNombre === COL_GAS.TOTAL) {
      lines.push('');
      lines.push('  ℹ "Total" por nombre coincide con COL_GAS.TOTAL — misma columna.');
    }
  }

  // ── 4. Resumen del problema ──────────────────────────────────
  lines.push('');
  lines.push('── DIAGNÓSTICO ──');
  if (COL_GAS.TOTAL > lastCol) {
    lines.push('  ❌ COL_GAS.TOTAL (' + COL_GAS.TOTAL + ') SUPERA el número de columnas (' + lastCol + ')');
  } else if (String(headerRow[COL_GAS.TOTAL - 1]).trim().toLowerCase() !== 'total') {
    lines.push('  ⚠ COL_GAS.TOTAL apunta a "' + headerRow[COL_GAS.TOTAL - 1] + '" — NO es "Total"');
    lines.push('  → El importe se lee de la columna INCORRECTA');
    if (colTotalNombre > 0) {
      lines.push('  → "Total" real está en columna ' + colTotalNombre + ' (el código usa ' + COL_GAS.TOTAL + ')');
    }
  } else {
    lines.push('  ✅ COL_GAS.TOTAL apunta correctamente a "Total"');
    lines.push('  → El problema puede ser formato de datos (texto vs. número)');
  }

  var txt = lines.join('\n');
  Logger.log(txt);
  ui.alert('🔍 Diagnóstico columnas GASTOS\n\n' + txt);
}

// ---- Diagnóstico de filtro de fechas GASTOS -----------------

function diagnosticarFiltroGastos() {
  var ui  = SpreadsheetApp.getUi();
  var now = new Date();
  var mes  = now.getMonth();
  var anyo = now.getFullYear();
  var tz   = Session.getScriptTimeZone();

  // Período: mes actual (mismo que opción 1 en generarInformeGastos)
  var desde = new Date(anyo, mes, 1);
  var hasta = new Date(anyo, mes + 1, 0, 23, 59, 59);

  var gasSheet;
  try { gasSheet = getSheet('GASTOS'); } catch (e) {
    ui.alert('⚠ No se pudo abrir GASTOS:\n' + e.message); return;
  }

  var lastRow = gasSheet.getLastRow();
  if (lastRow < 2) { ui.alert('GASTOS está vacía.'); return; }

  var lastCol = gasSheet.getLastColumn();
  var nLeer   = Math.min(lastRow - 1, 30);   // máx 30 filas para el diagnóstico
  var allData = gasSheet.getRange(2, 1, nLeer, lastCol).getValues();

  var lines = [];
  lines.push('=== DIAGNÓSTICO FILTRO GASTOS — ' +
    Utilities.formatDate(desde, tz, 'dd/MM/yyyy') + ' a ' +
    Utilities.formatDate(hasta, tz, 'dd/MM/yyyy') + ' ===');
  lines.push('Filas analizadas: ' + nLeer + ' (de ' + (lastRow - 1) + ' totales)');
  lines.push('');
  lines.push('ID_Gasto | Fecha_Alta | Fecha_Gasto | Total | ID ok? | fRef tipo | Incluido?');
  lines.push('─'.repeat(90));

  var nIncluidos = 0, nExcluidosId = 0, nExcluidosFecha = 0, nSinFecha = 0;
  var sumaIncluidos = 0;

  allData.forEach(function(row, i) {
    var idGasto   = String(row[COL_GAS.ID_GASTO    - 1] || '').trim();
    var fAltaRaw  = row[COL_GAS.FECHA_ALTA  - 1];
    var fGasRaw   = row[COL_GAS.FECHA_GASTO - 1];
    var total     = parseFloat(row[COL_GAS.TOTAL    - 1]) || 0;

    // Paso 1: filtro por ID
    var idOk = idGasto.indexOf('GAS-') === 0;
    if (!idOk) {
      nExcluidosId++;
      lines.push('[' + (i+2) + '] ' + (idGasto || '(vacío)').substring(0,12) +
        ' | ID sin GAS- → EXCLUIDO');
      return;
    }

    // Paso 2: determinar fRef (igual que en generarInformeGastos)
    var fRef, fRefTipo;
    if (fGasRaw instanceof Date && !isNaN(fGasRaw.getTime())) {
      fRef     = fGasRaw;
      fRefTipo = 'Fecha_Gasto (Date)';
    } else if (fAltaRaw instanceof Date && !isNaN(fAltaRaw.getTime())) {
      fRef     = fAltaRaw;
      fRefTipo = 'Fecha_Alta (Date, fallback)';
    } else {
      fRef     = null;
      fRefTipo = 'NINGUNA (tipo=' + typeof fGasRaw + ' / ' + typeof fAltaRaw + ')';
    }

    // Paso 3: aplicar filtro de fecha
    var incluido;
    if (!fRef) {
      incluido = true;   // sin fecha → incluir siempre
      nSinFecha++;
    } else {
      incluido = fRef >= desde && fRef <= hasta;
      if (!incluido) nExcluidosFecha++;
    }

    if (incluido) {
      nIncluidos++;
      sumaIncluidos += total;
    }

    var fAltaStr = (fAltaRaw instanceof Date && !isNaN(fAltaRaw.getTime()))
      ? Utilities.formatDate(fAltaRaw, tz, 'dd/MM/yy') : String(fAltaRaw || '—');
    var fGasStr  = (fGasRaw instanceof Date && !isNaN(fGasRaw.getTime()))
      ? Utilities.formatDate(fGasRaw,  tz, 'dd/MM/yy') : String(fGasRaw  || '—');
    var fRefStr  = fRef
      ? Utilities.formatDate(fRef, tz, 'dd/MM/yy') + ' [' + fRefTipo + ']'
      : fRefTipo;

    lines.push(
      '[' + (i+2) + '] ' + idGasto.substring(0, 12) +
      ' | Alta=' + fAltaStr +
      ' | Gasto=' + fGasStr +
      ' | fRef=' + fRefStr +
      ' | Total=' + total +
      ' | ' + (incluido ? '✅ SÍ' : '❌ NO (fuera de rango)')
    );
  });

  lines.push('─'.repeat(90));
  lines.push('RESUMEN:');
  lines.push('  Incluidos:              ' + nIncluidos + '  → suma Total = ' + sumaIncluidos.toFixed(2));
  lines.push('  Excluidos (sin GAS-):   ' + nExcluidosId);
  lines.push('  Excluidos (fuera rango):' + nExcluidosFecha);
  lines.push('  Incluidos sin fecha:    ' + nSinFecha + ' (sin fecha → siempre incluir)');
  lines.push('');

  // Diagnóstico final
  if (nIncluidos === 0) {
    lines.push('❌ NINGUNA fila pasa el filtro — el informe muestra 0 gastos y 0,00€');
    lines.push('   Causa probable: las fechas de los gastos no caen dentro del período seleccionado.');
  } else if (sumaIncluidos === 0) {
    lines.push('⚠ ' + nIncluidos + ' fila(s) pasan el filtro pero la suma es 0,00€');
    lines.push('   Causa probable: las filas incluidas tienen Total = 0 o vacío.');
  } else {
    lines.push('✅ ' + nIncluidos + ' fila(s) incluidas con suma = ' + sumaIncluidos.toFixed(2) + '€');
    lines.push('   Si el informe muestra 0, hay una discrepancia con la lectura en _igConstruirHoja_.');
  }

  var txt = lines.join('\n');
  Logger.log(txt);
  // Mostrar en dos partes si es muy largo
  var MAX = 1200;
  if (txt.length <= MAX) {
    ui.alert('🔍 Diagnóstico filtro GASTOS\n\n' + txt);
  } else {
    ui.alert('🔍 Diagnóstico filtro GASTOS (1/2)\n\n' + txt.substring(0, MAX) + '\n\n[continúa en logs]');
    ui.alert('🔍 Diagnóstico filtro GASTOS (2/2)\n\n' + txt.substring(MAX));
  }
}

// ---- FASE 7 — Reportes y control financiero por período ----

function generarInformeGastos() {
  var ui  = SpreadsheetApp.getUi();
  var now = new Date();
  var mes = now.getMonth();
  var anyo = now.getFullYear();
  var trimestre = Math.floor(mes / 3) + 1;
  var MESES_NOM = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
                   'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

  var res = ui.prompt('Generar informe de gastos',
    'Selecciona el período:\n\n'
    + '  1 — Mes actual (' + MESES_NOM[mes] + ' ' + anyo + ')\n'
    + '  2 — Trimestre actual (T' + trimestre + ' ' + anyo + ')\n'
    + '  3 — Año actual (' + anyo + ')\n'
    + '  4 — Todo\n\n'
    + 'Escribe el número (1-4):',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  var op = res.getResponseText().trim();
  if (['1','2','3','4'].indexOf(op) === -1) {
    ui.alert('Opción no válida. Escribe 1, 2, 3 o 4.'); return;
  }

  var desde, hasta, label;
  if (op === '1') {
    desde = new Date(anyo, mes, 1);
    hasta = new Date(anyo, mes + 1, 0, 23, 59, 59);
    label = MESES_NOM[mes] + ' ' + anyo;
  } else if (op === '2') {
    var q0 = (trimestre - 1) * 3;
    desde = new Date(anyo, q0, 1);
    hasta = new Date(anyo, q0 + 3, 0, 23, 59, 59);
    label = 'T' + trimestre + ' ' + anyo;
  } else if (op === '3') {
    desde = new Date(anyo, 0, 1);
    hasta = new Date(anyo, 11, 31, 23, 59, 59);
    label = 'Año ' + anyo;
  } else {
    desde = null; hasta = null;
    label = 'Todos los períodos';
  }

  // Leer gastos
  var gasSheet = getSheet('GASTOS');
  var lastRow  = gasSheet.getLastRow();
  if (lastRow < 2) { ui.alert('No hay gastos en GASTOS.'); return; }
  var allData = gasSheet.getRange(2, 1, lastRow - 1, HEADERS_GASTOS.length).getValues();

  var gastos = allData.filter(function(row) {
    if (String(row[COL_GAS.ID_GASTO - 1]).trim().indexOf('GAS-') !== 0) return false;
    if (!desde) return true;
    var fGas = row[COL_GAS.FECHA_GASTO - 1];
    var fRef = (fGas instanceof Date && !isNaN(fGas.getTime()))
               ? fGas : row[COL_GAS.FECHA_ALTA - 1];
    if (!(fRef instanceof Date) || isNaN(fRef.getTime())) return true;
    return fRef >= desde && fRef <= hasta;
  });

  if (gastos.length === 0) {
    ui.alert('No hay gastos para: ' + label); return;
  }

  _igConstruirHoja_(gastos, label, op, desde, hasta);

  var ss = getTargetSS_();
  var sh = ss.getSheetByName('INFORME_GASTOS');
  if (sh) ss.setActiveSheet(sh);

  ui.alert('✅ Informe generado',
    'Período: ' + label + '\n' + gastos.length + ' gastos analizados.\n\nVer hoja: INFORME_GASTOS',
    ui.ButtonSet.OK);
}

function _igConstruirHoja_(gastos, label, op, desde, hasta) {
  var ss  = getTargetSS_();
  var sh  = ss.getSheetByName('INFORME_GASTOS');
  if (sh) { sh.clear(); sh.clearFormats(); sh.clearDataValidations(); }
  else    { sh = ss.insertSheet('INFORME_GASTOS'); }

  var tz    = Session.getScriptTimeZone();
  var ahora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  var NC    = 7;

  sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 55);
  sh.setColumnWidth(3, 115); sh.setColumnWidth(4, 95);
  sh.setColumnWidth(5, 115); sh.setColumnWidth(6, 70);
  sh.setColumnWidth(7, 130);

  var r = 1;

  // ── Título ────────────────────────────────────────────────────
  sh.getRange(r,1,1,NC).merge()
    .setValue('INFORME DE GASTOS — ' + label.toUpperCase())
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(r, 40); r++;

  var periodoStr = desde
    ? Utilities.formatDate(desde, tz, 'dd/MM/yyyy') + ' – ' + Utilities.formatDate(hasta, tz, 'dd/MM/yyyy')
    : 'Todos';
  sh.getRange(r,1,1,NC).merge()
    .setValue('Generado el ' + ahora + '   |   Período: ' + periodoStr)
    .setFontColor('#5f6368').setFontSize(10).setHorizontalAlignment('center');
  r++; r++;

  // ── Totales globales ──────────────────────────────────────────
  var totalImporte = 0, totalBase = 0, totalIva = 0, nPend = 0;
  gastos.forEach(function(row) {
    totalImporte += parseFloat(row[COL_GAS.TOTAL          - 1]) || 0;
    totalBase    += parseFloat(row[COL_GAS.BASE_IMPONIBLE - 1]) || 0;
    totalIva     += parseFloat(row[COL_GAS.IVA_IMPORTE    - 1]) || 0;
    if (String(row[COL_GAS.ESTADO_PAGO - 1]).trim() !== 'Pagado') nPend++;
  });

  _igSeccion_(sh, r, NC, 'RESUMEN DEL PERÍODO'); r++;
  var res = [
    ['Nº de gastos',             gastos.length, null, null],
    ['Total (Base + IVA)',        null, null, totalImporte],
    ['Base imponible',            null, totalBase,  null],
    ['IVA soportado',             null, null, totalIva],
    ['Gastos pendientes de pago', nPend, null, null]
  ];
  res.forEach(function(it) {
    sh.getRange(r,1).setValue(it[0]).setFontWeight('bold');
    if (it[1] !== null) sh.getRange(r,2).setValue(it[1]);
    if (it[2] !== null) sh.getRange(r,3).setValue(Math.round(it[2]*100)/100).setNumberFormat('#,##0.00 €');
    if (it[3] !== null) sh.getRange(r,5).setValue(Math.round(it[3]*100)/100).setNumberFormat('#,##0.00 €').setFontWeight('bold').setFontColor('#1a73e8');
    r++;
  });
  r++;

  // ── Desglose por tipo ─────────────────────────────────────────
  _igSeccion_(sh, r, NC, 'DESGLOSE POR TIPO DE GASTO'); r++;
  _igCabecera_(sh, r, NC, ['Tipo de gasto','Nº','Base (€)','IVA (€)','Total (€)','% s/total','Pend. pago']); r++;

  var tipos = {};
  gastos.forEach(function(row) {
    var t = String(row[COL_GAS.TIPO_GASTO - 1]).trim() || 'Sin tipo';
    if (!tipos[t]) tipos[t] = {n:0,base:0,iva:0,total:0,pend:0};
    tipos[t].n++;
    tipos[t].base  += parseFloat(row[COL_GAS.BASE_IMPONIBLE - 1]) || 0;
    tipos[t].iva   += parseFloat(row[COL_GAS.IVA_IMPORTE    - 1]) || 0;
    tipos[t].total += parseFloat(row[COL_GAS.TOTAL          - 1]) || 0;
    if (String(row[COL_GAS.ESTADO_PAGO - 1]).trim() !== 'Pagado') tipos[t].pend++;
  });
  Object.keys(tipos).sort(function(a,b){ return tipos[b].total - tipos[a].total; })
    .forEach(function(tipo, idx) {
      var t = tipos[tipo];
      var pct = totalImporte > 0 ? t.total / totalImporte : 0;
      sh.getRange(r,1,1,NC).setBackground(idx%2===0?'#ffffff':'#f8f9fa');
      sh.getRange(r,1).setValue(tipo);
      sh.getRange(r,2).setValue(t.n).setHorizontalAlignment('center');
      sh.getRange(r,3).setValue(Math.round(t.base *100)/100).setNumberFormat('#,##0.00 €');
      sh.getRange(r,4).setValue(Math.round(t.iva  *100)/100).setNumberFormat('#,##0.00 €');
      sh.getRange(r,5).setValue(Math.round(t.total*100)/100).setNumberFormat('#,##0.00 €');
      sh.getRange(r,6).setValue(pct).setNumberFormat('0.0%');
      sh.getRange(r,7).setValue(t.pend>0?t.pend:'—').setHorizontalAlignment('center');
      r++;
    });
  // Fila total tipos
  sh.getRange(r,1,1,NC).setBackground('#e8f0fe');
  sh.getRange(r,1).setValue('TOTAL').setFontWeight('bold');
  sh.getRange(r,2).setValue(gastos.length).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(r,3).setValue(Math.round(totalBase    *100)/100).setNumberFormat('#,##0.00 €').setFontWeight('bold');
  sh.getRange(r,4).setValue(Math.round(totalIva     *100)/100).setNumberFormat('#,##0.00 €').setFontWeight('bold');
  sh.getRange(r,5).setValue(Math.round(totalImporte *100)/100).setNumberFormat('#,##0.00 €').setFontWeight('bold');
  sh.getRange(r,6).setValue(1).setNumberFormat('0.0%').setFontWeight('bold');
  sh.getRange(r,7).setValue(nPend>0?nPend:'—').setFontWeight('bold').setHorizontalAlignment('center');
  r++; r++;

  // ── Desglose mensual (año o todo) ─────────────────────────────
  if (op === '3' || op === '4') {
    var MESES_CORTO = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
    _igSeccion_(sh, r, NC, 'DESGLOSE MENSUAL'); r++;
    _igCabecera_(sh, r, NC, ['Mes','Nº','Base (€)','IVA (€)','Total (€)','% s/total','Pend. pago']); r++;

    var meses = {};
    gastos.forEach(function(row) {
      var fGas = row[COL_GAS.FECHA_GASTO - 1];
      var fRef = (fGas instanceof Date && !isNaN(fGas.getTime())) ? fGas : row[COL_GAS.FECHA_ALTA - 1];
      var key  = (fRef instanceof Date && !isNaN(fRef.getTime()))
                 ? fRef.getFullYear() + '-' + String(fRef.getMonth()+1).padStart(2,'0') : '0000-00';
      var lbl  = key === '0000-00' ? 'Sin fecha'
                 : MESES_CORTO[parseInt(key.split('-')[1])-1] + ' ' + key.split('-')[0];
      if (!meses[key]) meses[key] = {label:lbl,n:0,base:0,iva:0,total:0,pend:0};
      meses[key].n++;
      meses[key].base  += parseFloat(row[COL_GAS.BASE_IMPONIBLE - 1]) || 0;
      meses[key].iva   += parseFloat(row[COL_GAS.IVA_IMPORTE    - 1]) || 0;
      meses[key].total += parseFloat(row[COL_GAS.TOTAL          - 1]) || 0;
      if (String(row[COL_GAS.ESTADO_PAGO - 1]).trim() !== 'Pagado') meses[key].pend++;
    });
    Object.keys(meses).sort().forEach(function(key, idx) {
      var m = meses[key];
      var pct = totalImporte > 0 ? m.total / totalImporte : 0;
      sh.getRange(r,1,1,NC).setBackground(idx%2===0?'#ffffff':'#f8f9fa');
      sh.getRange(r,1).setValue(m.label);
      sh.getRange(r,2).setValue(m.n).setHorizontalAlignment('center');
      sh.getRange(r,3).setValue(Math.round(m.base *100)/100).setNumberFormat('#,##0.00 €');
      sh.getRange(r,4).setValue(Math.round(m.iva  *100)/100).setNumberFormat('#,##0.00 €');
      sh.getRange(r,5).setValue(Math.round(m.total*100)/100).setNumberFormat('#,##0.00 €');
      sh.getRange(r,6).setValue(pct).setNumberFormat('0.0%');
      sh.getRange(r,7).setValue(m.pend>0?m.pend:'—').setHorizontalAlignment('center');
      r++;
    });
    r++;
  }

  // ── Pendientes de pago ────────────────────────────────────────
  var pendientes = gastos.filter(function(row) {
    return String(row[COL_GAS.ESTADO_PAGO - 1]).trim() !== 'Pagado';
  });
  if (pendientes.length > 0) {
    _igSeccion_(sh, r, NC, 'GASTOS PENDIENTES DE PAGO (' + pendientes.length + ')', '#fce8b2', '#b06000'); r++;
    _igCabecera_(sh, r, NC, ['ID','Fecha','Tipo','Proveedor','Concepto','Total (€)','Estado']); r++;
    pendientes.forEach(function(row, idx) {
      var fGas    = row[COL_GAS.FECHA_GASTO - 1];
      var fechaStr = (fGas instanceof Date && !isNaN(fGas.getTime()))
                     ? Utilities.formatDate(fGas, tz, 'dd/MM/yyyy') : '—';
      sh.getRange(r,1,1,NC).setBackground(idx%2===0?'#fffde7':'#fff9c4');
      sh.getRange(r,1).setValue(String(row[COL_GAS.ID_GASTO   -1]).trim());
      sh.getRange(r,2).setValue(fechaStr);
      sh.getRange(r,3).setValue(String(row[COL_GAS.TIPO_GASTO -1]).trim());
      sh.getRange(r,4).setValue(String(row[COL_GAS.PROVEEDOR  -1]).trim());
      sh.getRange(r,5).setValue(String(row[COL_GAS.CONCEPTO   -1]).trim());
      sh.getRange(r,6).setValue(parseFloat(row[COL_GAS.TOTAL  -1])||0).setNumberFormat('#,##0.00 €');
      sh.getRange(r,7).setValue(String(row[COL_GAS.ESTADO_PAGO-1]).trim()).setFontColor('#c5221f');
      r++;
    });
  }

  SpreadsheetApp.flush();
  Logger.log('INFORME_GASTOS generado: ' + gastos.length + ' gastos, ' + r + ' filas.');
}

function _igSeccion_(sh, r, nc, txt, bg, fg) {
  sh.getRange(r,1,1,nc).merge()
    .setValue(txt)
    .setBackground(bg || '#e8f0fe').setFontColor(fg || '#1a73e8')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(r, 26);
}

function _igCabecera_(sh, r, nc, hdrs) {
  sh.getRange(r,1,1,nc).setValues([hdrs])
    .setBackground('#202124').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');
}

// ---- Web App (Fase 5 — placeholder) ----

function doGet(e) {
  return HtmlService
    .createHtmlOutput(
      '<html><body style="font-family:sans-serif;padding:40px;color:#444">' +
      '<h2>⚙️ Panel en construcción</h2>' +
      '<p>El panel web completo se implementa en la Fase 5.</p>' +
      '</body></html>'
    )
    .setTitle('Sistema de Gestión — Álvaro');
}

function doPost(e) {
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = body.action || '';
    var token  = body.token  || '';

    if (action === 'migrate_m1m4' && token === DM185_WEB_TOKEN) {
      var result = _dm185_runSilent_();
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'migrate_explicit' && token === DM185_WEB_TOKEN) {
      var result2 = _dm185_runExplicit_(body.facSourceId || '', body.gasSourceId || '');
      return ContentService
        .createTextOutput(JSON.stringify(result2))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'list_folder' && token === DM185_WEB_TOKEN) {
      var fid = body.folderId || '';
      var out = { folderId: fid, files: [], error: null };
      try {
        var folder = DriveApp.getFolderById(fid);
        out.folderName = folder.getName();
        var iter = folder.getFiles();
        while (iter.hasNext()) {
          var ff = iter.next();
          out.files.push({ id: ff.getId(), name: ff.getName(), mime: ff.getMimeType() });
        }
      } catch(e) { out.error = e.message; }
      return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'read_config' && token === DM185_WEB_TOKEN) {
      var ss2   = getTargetSS_();
      var keys  = ['carpeta_facturas_id', 'gastos_folder_id'];
      var vals  = {};
      keys.forEach(function(k) { vals[k] = getConfig(k) || null; });
      vals['FACTURAS_FOLDER_ID']  = String(_dmGetConfig_(ss2, 'FACTURAS_FOLDER_ID')  || '');
      vals['GASTOS_FOLDER_ID']    = String(_dmGetConfig_(ss2, 'GASTOS_FOLDER_ID')    || '');
      return ContentService.createTextOutput(JSON.stringify(vals)).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'run_panel_check' && token === DM185_WEB_TOKEN) {
      var checks = [
        { name: 'PANEL_GENERAL',   fn: function() { _actualizarPanelGeneral_(); } },
        { name: 'PANEL_FACTURAS',  fn: function() {
            var ss7 = getTargetSS_();
            var sh7 = ss7.getSheetByName('PANEL_FACTURAS');
            if (!sh7) throw new Error('Hoja PANEL_FACTURAS no encontrada');
            var rows7 = sh7.getLastRow(); var cols7 = sh7.getLastColumn();
            if (rows7 < 10 || cols7 < 5) throw new Error('PANEL_FACTURAS parece vacía (' + rows7 + 'r × ' + cols7 + 'c)');
          }
        },
        { name: 'PANEL_CLIENTE',   fn: function() { refreshPanelCliente(); } },
        { name: 'PANEL_GASTOS',    fn: function() { _actualizarPanelGastos_(); } },
        { name: 'PANEL_CONTABLE',  fn: function() { _actualizarPanelContable_(); } }
      ];
      var results7 = checks.map(function(c) {
        try   { c.fn(); return { name: c.name, ok: true };  }
        catch (ce) { return { name: c.name, ok: false, error: ce.message }; }
      });
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({ ok: true, results: results7 })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'get_pcn_kpis' && token === DM185_WEB_TOKEN) {
      var ss9 = getTargetSS_();
      var sh9 = ss9.getSheetByName('PANEL_CONTABLE');
      if (!sh9) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'PANEL_CONTABLE no encontrado' })).setMimeType(ContentService.MimeType.JSON);
      var rango9 = _pcnObtenerRango_(sh9);
      var fac9   = _pcnLeerFacturas_(ss9, rango9);
      var gas9   = _pcnLeerGastos_(ss9, rango9);
      var periodo9 = _pcnPeriodoEtiqueta_(_pcnLeerTipoPeriodo_(sh9));
      // Leer también los valores visibles actuales del panel (fila KPI2_VAL = 20)
      var kpiCells = sh9.getRange(20, 2, 1, 12).getValues()[0];
      return ContentService.createTextOutput(JSON.stringify({
        ok: true,
        periodo: periodo9,
        rango: rango9 ? {
          ini: Utilities.formatDate(rango9.ini, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
          fin: Utilities.formatDate(rango9.fin, Session.getScriptTimeZone(), 'dd/MM/yyyy')
        } : null,
        facturas: { count: fac9.count, total: fac9.total, ivaRep: fac9.ivaRep, pendiente: fac9.pendiente },
        gastos:   { count: gas9.count, total: gas9.total, ivaSop: gas9.ivaSop },
        kpiCellsRow20: kpiCells
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'set_periodo_contable' && token === DM185_WEB_TOKEN) {
      var ss8  = getTargetSS_();
      var sh8  = ss8.getSheetByName('PANEL_CONTABLE');
      if (!sh8) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'PANEL_CONTABLE no encontrado' })).setMimeType(ContentService.MimeType.JSON);
      var periodo8 = body.periodo || 'Año actual';
      sh8.getRange(5, 4, 1, 4).merge().setValue(periodo8);
      SpreadsheetApp.flush();
      _actualizarPanelContable_();
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({ ok: true, periodo: periodo8 })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'peek_sheet' && token === DM185_WEB_TOKEN) {
      var ss6   = getTargetSS_();
      var name6 = body.name || '';
      var sh6   = ss6.getSheetByName(name6);
      if (!sh6) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'not found: ' + name6 })).setMimeType(ContentService.MimeType.JSON);
      var lastRow = sh6.getLastRow();
      var lastCol = sh6.getLastColumn();
      var sample  = [];
      if (lastRow > 0 && lastCol > 0) {
        var nRow = Math.min(lastRow, 3);
        var nCol = Math.min(lastCol, 5);
        sample = sh6.getRange(1, 1, nRow, nCol).getValues();
      }
      return ContentService.createTextOutput(JSON.stringify({ ok: true, name: name6, lastRow: lastRow, lastCol: lastCol, sample: sample, hidden: sh6.isSheetHidden() })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'list_sheets' && token === DM185_WEB_TOKEN) {
      var ss4   = getTargetSS_();
      var sheets4 = ss4.getSheets();
      var out4  = sheets4.map(function(s) {
        return { name: s.getName(), index: s.getIndex(), hidden: s.isSheetHidden() };
      });
      return ContentService.createTextOutput(JSON.stringify({ ok: true, sheets: out4 })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'hide_reorder_sheets' && token === DM185_WEB_TOKEN) {
      var ss5      = getTargetSS_();
      var toHide   = body.hide   || [];
      var toShow   = body.show   || [];
      var order    = body.order  || [];
      var renamed  = body.rename || {};
      var log5     = { renamed: [], shown: [], hidden: [], ordered: [], errors: [] };

      // 1. Rename sheets
      Object.keys(renamed).forEach(function(oldName) {
        var newName = renamed[oldName];
        var s = ss5.getSheetByName(oldName);
        if (s) { s.setName(newName); log5.renamed.push(oldName + ' → ' + newName); }
        else    { log5.errors.push('rename: not found: ' + oldName); }
      });

      // 2. Show sheets
      toShow.forEach(function(name) {
        var s = ss5.getSheetByName(name);
        if (s) { s.showSheet(); log5.shown.push(name); }
        else   { log5.errors.push('show: not found: ' + name); }
      });

      // 3. Hide sheets
      toHide.forEach(function(name) {
        var s = ss5.getSheetByName(name);
        if (s) {
          try { s.hideSheet(); log5.hidden.push(name); }
          catch(he) { log5.errors.push('hide ' + name + ': ' + he.message); }
        } else { log5.errors.push('hide: not found: ' + name); }
      });

      // 4. Reorder visible sheets (order = array of names in desired order)
      if (order.length > 0) {
        order.forEach(function(name, pos) {
          var s = ss5.getSheetByName(name);
          if (s) {
            try { ss5.moveActiveSheet; ss5.setActiveSheet(s); ss5.moveActiveSheet(pos + 1); log5.ordered.push(name); }
            catch(oe) { log5.errors.push('order ' + name + ': ' + oe.message); }
          } else { log5.errors.push('order: not found: ' + name); }
        });
      }

      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({ ok: true, log: log5 })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'move_items' && token === DM185_WEB_TOKEN) {
      var moves = body.moves || [];
      var log4  = [];
      moves.forEach(function(m) {
        var srcId = m.id || '';
        var dstId = m.targetParentId || '';
        if (!srcId || !dstId) {
          log4.push({ id: srcId, ok: false, error: 'missing id or targetParentId' });
          return;
        }
        try {
          var dst = DriveApp.getFolderById(dstId);
          var moved;
          try {
            var folder = DriveApp.getFolderById(srcId);
            moved = folder.moveTo(dst);
            log4.push({ id: srcId, ok: true, name: folder.getName(), type: 'folder' });
          } catch(_) {
            var file = DriveApp.getFileById(srcId);
            moved = file.moveTo(dst);
            log4.push({ id: srcId, ok: true, name: file.getName(), type: 'file' });
          }
        } catch(me) {
          log4.push({ id: srcId, ok: false, error: me.message });
        }
      });
      return ContentService.createTextOutput(JSON.stringify({ ok: true, log: log4 })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'trash_items' && token === DM185_WEB_TOKEN) {
      var ids  = body.ids  || [];
      var log3 = [];
      ids.forEach(function(fid) {
        try {
          var f = DriveApp.getFileById(fid);
          f.setTrashed(true);
          log3.push({ id: fid, ok: true,  name: f.getName() });
        } catch(fe) {
          try {
            var fo = DriveApp.getFolderById(fid);
            fo.setTrashed(true);
            log3.push({ id: fid, ok: true,  name: fo.getName() });
          } catch(fe2) {
            log3.push({ id: fid, ok: false, error: fe2.message });
          }
        }
      });
      return ContentService.createTextOutput(JSON.stringify({ ok: true, log: log3 })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'dump_sheets' && token === DM185_WEB_TOKEN) {
      var ss3 = getTargetSS_();
      var out3 = { driveConfig: [], historial: [], ok: true };
      var dcSh = ss3.getSheetByName('DRIVE_CONFIG');
      if (dcSh) {
        var dcRows = dcSh.getDataRange().getValues();
        dcRows.forEach(function(r) { out3.driveConfig.push(r.map(String)); });
      }
      var hSh = ss3.getSheetByName('HISTORIAL_PAQUETES');
      if (hSh) {
        var hLast = hSh.getLastRow();
        var hCols = hSh.getLastColumn();
        if (hLast >= 2 && hCols >= 1) {
          var hStart = Math.max(2, hLast - 9);
          var hRows  = hSh.getRange(hStart, 1, hLast - hStart + 1, hCols).getValues();
          hRows.forEach(function(r) { out3.historial.push(r.map(String)); });
        }
      }
      return ContentService.createTextOutput(JSON.stringify(out3)).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: 'Unauthorized or unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: e.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
