// ============================================================
// PanelContableHistorial.gs — FASE 18.3
// Historial de paquetes contables + modal de selección HTML
// ============================================================

var PCN_HIST_SHEET = 'HISTORIAL_PAQUETES';

// ============================================================
// WRAPPERS PÚBLICOS
// ============================================================

function verHistorialPaquetes() {
  _pcnVerHistorial_();
}

function mostrarSeleccionActual() {
  _pcnMostrarSeleccionActual_();
}

// ============================================================
// HOJA HISTORIAL_PAQUETES
// ============================================================

/**
 * Devuelve la hoja HISTORIAL_PAQUETES, creándola si no existe.
 * Columnas: Fecha · Periodo · Facturas IDs · Nº Fac · Gastos IDs · Nº Gas
 *           · PDFs OK · Adj OK · Carpeta Drive · Enviado al gestor
 */
function _pcnEnsureHistorialSheet_(ss) {
  var sh = ss.getSheetByName(PCN_HIST_SHEET);
  if (sh) return sh;

  sh = ss.insertSheet(PCN_HIST_SHEET);
  var hdrs = [
    'Fecha generación', 'Periodo',
    'Facturas incluidas (IDs)', 'Nº Fac',
    'Gastos incluidos (IDs)',   'Nº Gas',
    'PDFs copiados', 'Adjuntos copiados',
    'Carpeta Drive', 'Enviado al gestor'
  ];
  sh.getRange(1, 1, 1, hdrs.length)
    .setValues([hdrs])
    .setFontWeight('bold')
    .setBackground('#4a148c')
    .setFontColor('#ffffff');
  sh.setFrozenRows(1);

  [140, 110, 380, 65, 380, 65, 110, 130, 340, 200]
    .forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });

  logEvent('INFO', 'ensureHistorialSheet', PCN_HIST_SHEET + ' creado', '');
  return sh;
}

/**
 * Registra un paquete generado.
 * facRows / gasRows son los .rows de los datos filtrados (ya seleccionados).
 */
function _pcnLogPaquete_(ss, periodo, facRows, gasRows, stats) {
  var sh  = _pcnEnsureHistorialSheet_(ss);
  var now = new Date();

  var facIds = facRows.map(function(r) {
    return String(r[0] || '').trim();
  }).filter(Boolean).join(', ');

  var gasIds = gasRows.map(function(r) {
    return String(r[0] || '').trim();
  }).filter(Boolean).join(', ');

  sh.appendRow([
    now,
    periodo,
    facIds,
    facRows.length,
    gasIds,
    gasRows.length,
    stats.facPdfOk  || 0,
    stats.gasAdjOk  || 0,
    stats.folderUrl || '',
    ''
  ]);
  sh.getRange(sh.getLastRow(), 1).setNumberFormat('dd/MM/yyyy HH:mm');
  logEvent('INFO', 'logPaquete', periodo,
    facRows.length + ' fac / ' + gasRows.length + ' gas');
}

/**
 * Marca el último paquete del periodo como enviado al gestor.
 * Llamar desde _pcnEnviarEmailGestor_ tras envío exitoso.
 */
function _pcnMarcarEnviado_(ss, periodo, emailGestor) {
  var sh = ss.getSheetByName(PCN_HIST_SHEET);
  if (!sh || sh.getLastRow() < 2) return;

  var tz   = Session.getScriptTimeZone();
  var now  = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  var data = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();

  for (var i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]).trim() === periodo) {
      sh.getRange(i + 2, 10)
        .setValue('✅ ' + now + (emailGestor ? ' → ' + emailGestor : ''));
      return;
    }
  }
}

// ============================================================
// VER HISTORIAL
// ============================================================

function _pcnVerHistorial_() {
  var ss = getTargetSS_();
  var sh = ss.getSheetByName(PCN_HIST_SHEET);

  if (!sh || sh.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert(
      '⚪ Historial vacío.\n\nGenera el primer paquete contable para empezar el registro.'
    );
    return;
  }

  var tz      = Session.getScriptTimeZone();
  var lastRow = sh.getLastRow();
  var n       = Math.min(lastRow - 1, 10);
  var start   = Math.max(2, lastRow - n + 1);
  var data    = sh.getRange(start, 1, n, 10).getValues();

  var a = ['=== HISTORIAL PAQUETES CONTABLES (últimos ' + n + ') ===\n'];
  data.slice().reverse().forEach(function(r) {
    var fecha = r[0] instanceof Date
      ? Utilities.formatDate(r[0], tz, 'dd/MM/yyyy HH:mm') : String(r[0] || '—');
    a.push('📦 ' + fecha + '  —  ' + (r[1] || '?'));
    a.push('   Facturas: ' + (r[3] || 0) + '  ·  Gastos: ' + (r[5] || 0) +
           '  ·  PDFs: ' + (r[6] || 0) + '  ·  Adj: ' + (r[7] || 0));
    a.push('   Enviado: ' + (String(r[9] || '') || '—'));
    a.push('');
  });

  logEvent('INFO', 'verHistorialPaquetes', n + ' registros mostrados', '');
  SpreadsheetApp.getUi().alert('📂 Historial paquetes\n\n' + a.join('\n'));
}

// ============================================================
// MODAL — SELECCIÓN ACTUAL
// ============================================================

/**
 * Muestra un modal HTML con la lista de documentos actualmente
 * seleccionados en el selector del periodo activo.
 */
function _pcnMostrarSeleccionActual_() {
  var ss = getTargetSS_();
  var sh = ss.getSheetByName(PCN_SHEET);
  if (!sh) {
    SpreadsheetApp.getUi().alert('⚠ PANEL_CONTABLE no existe.');
    return;
  }

  var tipo    = _pcnLeerTipoPeriodo_(sh);
  var periodo = _pcnPeriodoEtiqueta_(tipo);
  var rango   = _pcnObtenerRango_(sh);
  var facData = _pcnGetFacturasData_(ss, rango);
  var gasData = _pcnGetGastosData_(ss, rango);
  var sel     = _pcnGetDocumentosSeleccionados_(sh, facData, gasData);
  var tz      = Session.getScriptTimeZone();

  var html = _pcnBuildSeleccionHtml_(
    sel.facDataFilt.rows, sel.gasDataFilt.rows,
    periodo, sel, tz
  );

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(660).setHeight(560),
    '📋 Selección actual — ' + periodo
  );
}

function _pcnBuildSeleccionHtml_(facRows, gasRows, periodo, sel, tz) {
  var css =
    '<style>' +
    'body{font-family:Arial,sans-serif;font-size:12px;color:#202124;margin:0;padding:10px 14px}' +
    'h3{margin:10px 0 4px;color:#4a148c;font-size:13px;border-bottom:1px solid #e0e0e0;' +
       'padding-bottom:3px}' +
    'table{width:100%;border-collapse:collapse;margin-bottom:10px;font-size:11px}' +
    'th{background:#4a148c;color:#fff;padding:5px 7px;text-align:left}' +
    'tr:nth-child(even){background:#f8f9fa}' +
    'td{padding:4px 7px;border-bottom:1px solid #e8eaed}' +
    '.num{text-align:right}.cen{text-align:center}' +
    '.ok{color:#137333;font-weight:bold}.warn{color:#e37400}' +
    '.banner{padding:8px 10px;margin-bottom:10px;font-size:11px;border-radius:3px}' +
    '.stale{background:#fef7e0;border-left:3px solid #e37400}' +
    '.nosel{background:#e8f0fe;border-left:3px solid #1a73e8}' +
    '.foot{font-size:10px;color:#5f6368;border-top:1px solid #e0e0e0;' +
          'padding-top:6px;margin-top:4px}' +
    '</style>';

  var banner = '';
  if (sel.stale) {
    banner = '<div class="banner stale">⚠ <b>Selector desactualizado</b> — ' +
      'mostrando todos los documentos del periodo. Pulsa <b>Actualizar</b> para ' +
      'refrescar y elegir individualmente.</div>';
  } else if (!sel.hasSelector) {
    banner = '<div class="banner nosel">ℹ️ Selector no renderizado — ' +
      'pulsa <b>Actualizar</b> en el panel para ver y elegir documentos ' +
      'individualmente.</div>';
  }

  // Tabla facturas
  var facHtml = '';
  facRows.forEach(function(r) {
    var id  = _h(r[0]);
    var fch = r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'dd/MM/yyyy') : '—';
    var cli = _h(String(r[5] || '').substring(0, 30));
    var tot = typeof r[16] === 'number' ? _pcnFmtEur_(r[16]) : '—';
    var ok  = !!(r[18] && _pcnExtraerFileId_(String(r[18])));
    facHtml +=
      '<tr><td><b>' + id + '</b></td><td>' + fch + '</td><td>' + cli + '</td>' +
      '<td class="num">' + tot + '</td>' +
      '<td class="cen ' + (ok ? 'ok' : 'warn') + '">' + (ok ? '✅' : '⚠') + '</td></tr>';
  });
  if (!facHtml) {
    facHtml = '<tr><td colspan="5" class="cen" style="color:#5f6368;padding:8px">' +
      'Sin facturas seleccionadas</td></tr>';
  }

  // Tabla gastos
  var gasHtml = '';
  gasRows.forEach(function(r) {
    var id  = _h(r[0]);
    var fg  = r[3];
    var fch = (fg instanceof Date && !isNaN(fg))
      ? Utilities.formatDate(fg, tz, 'dd/MM/yyyy')
      : (r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'dd/MM/yyyy') : '—');
    var prov = _h(String(r[4] || '').substring(0, 28));
    var tot  = typeof r[12] === 'number' ? _pcnFmtEur_(r[12]) : '—';
    var hasFile = !!(String(r[16] || '').trim() ||
                     _pcnExtraerFileId_(String(r[15] || '')));
    gasHtml +=
      '<tr><td><b>' + id + '</b></td><td>' + fch + '</td><td>' + prov + '</td>' +
      '<td class="num">' + tot + '</td>' +
      '<td class="cen ' + (hasFile ? 'ok' : 'warn') + '">' +
        (hasFile ? '✅' : '⚠') + '</td></tr>';
  });
  if (!gasHtml) {
    gasHtml = '<tr><td colspan="5" class="cen" style="color:#5f6368;padding:8px">' +
      'Sin gastos seleccionados</td></tr>';
  }

  var thRow = '<tr><th>ID</th><th>Fecha</th><th>Nombre</th>' +
              '<th class="num">Total</th><th class="cen">Archivo</th></tr>';

  return '<!DOCTYPE html><html><head>' + css + '</head><body>' +
    banner +
    '<h3>📄 Facturas seleccionadas (' +
      facRows.length + ' de ' + sel.facTotal + ')</h3>' +
    '<table><thead>' + thRow + '</thead><tbody>' + facHtml + '</tbody></table>' +
    '<h3>🧾 Gastos seleccionados (' +
      gasRows.length + ' de ' + sel.gasTotal + ')</h3>' +
    '<table><thead>' + thRow.replace('Nombre', 'Proveedor') +
      '</thead><tbody>' + gasHtml + '</tbody></table>' +
    '<p class="foot">Periodo: <b>' + _h(periodo) + '</b>  ·  ' +
      'Para cambiar la selección: cierra este modal y usa los checkboxes en el panel.</p>' +
    '</body></html>';
}

/** Escapa caracteres HTML para evitar inyección en el modal. */
function _h(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
