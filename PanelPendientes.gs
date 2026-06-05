// ============================================================
// PanelPendientes.gs — Vista global: clientes con trabajos sin facturar
// ============================================================

var PP_ = {
  BOTON:     [4, 3],   // C4 — checkbox Actualizar
  ESTADO:    [5, 2],   // B5 — estado
  KPI_LBL:   7,
  KPI_VAL:   8,
  DATA_START: 11
};

// ================================================================
// HELPERS INTERNOS
// ================================================================
function _ppEstado_(sh, msg, bgColor) {
  var fc = (bgColor === '#f9ab00' || bgColor === '#ea4335' || bgColor === '#34a853')
           ? '#ffffff' : '#5f6368';
  sh.getRange(PP_.ESTADO[0], PP_.ESTADO[1], 1, 6).merge()
    .setValue(msg)
    .setBackground(bgColor || '#f8f9fa').setFontColor(fc)
    .setFontSize(9).setHorizontalAlignment('left').setVerticalAlignment('middle');
}

// ================================================================
// SETUP — Crea la hoja con estructura fija
// ================================================================
function setupPanelPendientes(silent) {
  var ss  = SpreadsheetApp.getActive();
  var old = ss.getSheetByName('PANEL_PENDIENTES');
  if (old) { try { ss.deleteSheet(old); } catch(e) {} }
  var sh = ss.insertSheet('PANEL_PENDIENTES');
  ss.setActiveSheet(sh);
  SpreadsheetApp.flush();

  sh.setColumnWidth(1, 10);   // A — margen
  sh.setColumnWidth(2, 240);  // B — Cliente
  sh.setColumnWidth(3, 110);  // C — Total trabajos
  sh.setColumnWidth(4, 110);  // D — Sin facturar
  sh.setColumnWidth(5, 110);  // E — Horas pendientes
  sh.setColumnWidth(6, 130);  // F — Importe estimado
  sh.setColumnWidth(7, 120);  // G — Estado / semáforo

  sh.setRowHeight(1, 50);
  sh.setRowHeight(2, 8);
  sh.setRowHeight(3, 26);
  sh.setRowHeight(4, 46);
  sh.setRowHeight(5, 32);
  sh.setRowHeight(6, 8);

  // Fila 1 — Título
  sh.getRange(1, 2, 1, 6).merge()
    .setValue('PANEL PENDIENTES DE FACTURAR')
    .setBackground('#c62828').setFontColor('#ffffff')
    .setFontSize(18).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');

  // Fila 2 — spacer
  sh.getRange(2, 1, 1, 7).setBackground('#ffffff');

  // Fila 3 — descripción
  sh.getRange(3, 2, 1, 6).merge()
    .setValue('Vista global de todos los clientes · trabajos realizados sin factura emitida')
    .setBackground('#f8f9fa').setFontColor('#5f6368').setFontSize(9)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');

  // Fila 4 — botón Actualizar (etiqueta B4 + checkbox C4)
  sh.getRange(4, 2).setValue('ACTUALIZAR PANEL PENDIENTES')
    .setBackground('#c62828').setFontColor('#ffffff')
    .setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(4, 3).insertCheckboxes()
    .setBackground('#b71c1c').setFontColor('#ffffff').setHorizontalAlignment('center');
  sh.getRange(4, 4, 1, 4).setBackground('#c62828');

  // Fila 5 — estado
  _ppEstado_(sh, 'Activa el checkbox para cargar los datos.', '#f8f9fa');

  // Fila 6 — spacer
  sh.getRange(6, 1, 1, 7).setBackground('#ffffff');

  SpreadsheetApp.flush();
  if (!silent) {
    SpreadsheetApp.getUi().alert(
      'PANEL_PENDIENTES creado.\n\nActiva el checkbox de la fila 4 para cargar todos los datos.'
    );
  }
}

// ================================================================
// REFRESH — Escanea TRABAJOS + FACTURAS/PROFORMAS y renderiza tabla global
// ================================================================
function refreshPanelPendientes() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('PANEL_PENDIENTES');
  if (!sh) {
    // Auto-crear la hoja la primera vez (silencioso)
    setupPanelPendientes(true);
    sh = ss.getSheetByName('PANEL_PENDIENTES');
    if (!sh) return; // no debería pasar
  }

  _ppEstado_(sh, 'Cargando datos...', '#f9ab00');
  SpreadsheetApp.flush();

  try {
    // ── 1. TRABAJOS → agrupar por cliente ────────────────────────
    var clienteTrabajos = {}; // { clienteLabel: [{id,fecha,horas,estado,materiales}] }
    try {
      var shTra    = getSheet('TRABAJOS');
      var lastTr   = shTra.getLastRow();
      if (lastTr > 1) {
        var numColsTra = Math.min(shTra.getLastColumn(), COL_TRA.MATERIALES);
        var dataTra    = shTra.getRange(2, 1, lastTr - 1, numColsTra).getValues();
        for (var i = 0; i < dataTra.length; i++) {
          var trow       = dataTra[i];
          var clienteLbl = String(trow[COL_TRA.CLIENTE - 1] || '').trim();
          if (!clienteLbl) continue;
          var tId = String(trow[COL_TRA.ID - 1] || '').trim();
          var ft  = _parseTrabajosDate_(trow[COL_TRA.FECHA_TRABAJO - 1]);
          if (!tId || !ft) continue;
          if (!clienteTrabajos[clienteLbl]) clienteTrabajos[clienteLbl] = [];
          clienteTrabajos[clienteLbl].push({
            id:         tId,
            fecha:      ft,
            horas:      _calcHoras_(trow[COL_TRA.HORA_INICIO - 1], trow[COL_TRA.HORA_FIN - 1]),
            estado:     String(trow[COL_TRA.ESTADO         - 1] || '').trim(),
            materiales: Number(trow[COL_TRA.MATERIALES     - 1] || 0)
          });
        }
      }
    } catch(eTra) { Logger.log('[PP] TRABAJOS: ' + eTra.message); }

    // ── 2. FACTURAS + PROFORMAS → mapa {trabajoId: true} ─────────
    var facturaPorTrabajo = {};
    var _idxIds_ = function(idsStr) {
      if (!idsStr) return;
      String(idsStr).split(',').forEach(function(tid) {
        var t = tid.trim(); if (t) facturaPorTrabajo[t] = true;
      });
    };
    try {
      var shFac  = getSheet('FACTURAS');
      var lastFr = shFac.getLastRow();
      if (lastFr > 1) {
        var dataFac = shFac.getRange(2, 1, lastFr - 1, 10).getValues();
        for (var f = 0; f < dataFac.length; f++) _idxIds_(dataFac[f][9]);
      }
    } catch(eFac) {}
    try {
      var shPro  = getSheet('PROFORMAS');
      var lastPr = shPro.getLastRow();
      if (lastPr > 1) {
        var dataPro = shPro.getRange(2, 1, lastPr - 1, 10).getValues();
        for (var p = 0; p < dataPro.length; p++) _idxIds_(dataPro[p][9]);
      }
    } catch(ePro) {}

    // ── 3. Estadísticas por cliente ───────────────────────────────
    var clienteStats = [];
    Object.keys(clienteTrabajos).forEach(function(clienteLbl) {
      var trabajos   = clienteTrabajos[clienteLbl];
      var pendientes = trabajos.filter(function(t) { return !facturaPorTrabajo[t.id]; });
      clienteStats.push({
        cliente:     clienteLbl,
        total:       trabajos.length,
        pendCount:   pendientes.length,
        pendHoras:   pendientes.reduce(function(s, t) { return s + t.horas;       }, 0),
        pendImporte: pendientes.reduce(function(s, t) { return s + (t.materiales || 0); }, 0)
      });
    });

    // Ordenar: más pendientes primero; empate → alfabético
    clienteStats.sort(function(a, b) {
      if (b.pendCount !== a.pendCount) return b.pendCount - a.pendCount;
      return a.cliente.localeCompare(b.cliente);
    });

    // ── 4. KPIs globales ──────────────────────────────────────────
    var kpiCliPend    = clienteStats.filter(function(c) { return c.pendCount > 0; }).length;
    var kpiTotalPend  = clienteStats.reduce(function(s, c) { return s + c.pendCount;   }, 0);
    var kpiHorasPend  = clienteStats.reduce(function(s, c) { return s + c.pendHoras;   }, 0);

    var kpiLbls  = ['Clientes con pendientes', 'Trabajos sin facturar', 'Horas pendientes'];
    var kpiVals  = [
      kpiCliPend   + ' cliente(s)',
      kpiTotalPend + ' trabajo(s)',
      kpiHorasPend.toFixed(1).replace('.', ',') + ' h'
    ];
    var kpiClaro = [
      kpiCliPend  === 0 ? '#34a853' : kpiCliPend  <= 3 ? '#ef6c00' : '#c62828',
      kpiTotalPend === 0 ? '#34a853' : kpiTotalPend <= 5 ? '#ef6c00' : '#c62828',
      '#1a73e8'
    ];
    var kpiOscuro = [
      kpiCliPend  === 0 ? '#2d8e47' : kpiCliPend  <= 3 ? '#e65100' : '#b71c1c',
      kpiTotalPend === 0 ? '#2d8e47' : kpiTotalPend <= 5 ? '#e65100' : '#b71c1c',
      '#1565c0'
    ];

    // Render KPIs (rows 7-8)
    sh.getRange(PP_.KPI_LBL, 1, 2, 7).clearContent().clearFormat().setBackground('#ffffff');
    var kpiCols = [2, 4, 6];
    kpiLbls.forEach(function(lbl, ki) {
      var col = kpiCols[ki];
      sh.getRange(PP_.KPI_LBL, col, 1, 2).merge()
        .setValue(lbl)
        .setBackground(kpiClaro[ki]).setFontColor('#ffffff')
        .setFontSize(9).setFontWeight('bold')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
      sh.getRange(PP_.KPI_VAL, col, 1, 2).merge()
        .setValue(kpiVals[ki])
        .setBackground(kpiOscuro[ki]).setFontColor('#ffffff')
        .setFontSize(20).setFontWeight('bold')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
    });
    sh.setRowHeight(PP_.KPI_LBL, 22);
    sh.setRowHeight(PP_.KPI_VAL, 46);

    // Spacer fila 9-10
    sh.getRange(9, 1, 2, 7).clearContent().clearFormat().setBackground('#ffffff');
    sh.setRowHeight(9, 8);
    sh.setRowHeight(10, 8);

    // ── 5. Zona dinámica (fila DATA_START = 11) ───────────────────
    var shMaxRow  = sh.getMaxRows();
    var clearRows = Math.max(shMaxRow - PP_.DATA_START + 1, 200);
    var clearRng  = sh.getRange(PP_.DATA_START, 1, clearRows, 7);
    clearRng.breakApart();
    SpreadsheetApp.flush();
    clearRng.clearContent().clearFormat().clearDataValidations().clearNote();
    SpreadsheetApp.flush();

    var currentRow = PP_.DATA_START;

    // ── Sección: clientes CON pendientes ─────────────────────────
    var conPend = clienteStats.filter(function(c) { return c.pendCount > 0; });
    var sinPend = clienteStats.filter(function(c) { return c.pendCount === 0; });

    // Cabecera roja
    sh.getRange(currentRow, 2, 1, 6).merge()
      .setValue('🔴  CLIENTES CON TRABAJOS SIN FACTURAR (' + conPend.length + ')')
      .setBackground('#c62828').setFontColor('#ffffff')
      .setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.setRowHeight(currentRow, 32);
    currentRow++;

    if (conPend.length === 0) {
      sh.getRange(currentRow, 2, 1, 6).merge()
        .setValue('✅  Todos los clientes tienen sus trabajos facturados.')
        .setBackground('#e8f5e9').setFontColor('#2d8e47')
        .setFontSize(10).setFontWeight('bold').setVerticalAlignment('middle');
      sh.setRowHeight(currentRow, 32);
      currentRow++;
    } else {
      // Cabecera de tabla
      sh.getRange(currentRow, 2, 1, 6).setValues([[
        'Cliente', 'Total trabajos', 'Sin facturar', 'Horas pend.', 'Importe est.', 'Estado'
      ]]).setBackground('#b71c1c').setFontColor('#ffffff')
        .setFontSize(9).setFontWeight('bold').setVerticalAlignment('middle');
      sh.setRowHeight(currentRow, 24);
      currentRow++;

      for (var ci = 0; ci < conPend.length; ci++) {
        var cs     = conPend[ci];
        var claro  = cs.pendCount <= 3 ? '#ef6c00' : '#c62828';
        var bg     = ci % 2 === 0 ? '#ffebee' : '#fce4ec';
        var estado = cs.pendCount <= 3 ? '🟠 ' + cs.pendCount + ' pend.' : '🔴 ' + cs.pendCount + ' pend.';
        sh.getRange(currentRow, 2, 1, 6).setValues([[
          cs.cliente,
          cs.total,
          cs.pendCount,
          cs.pendHoras.toFixed(1).replace('.', ',') + ' h',
          cs.pendImporte > 0 ? formatCurrency(cs.pendImporte) : '—',
          estado
        ]]).setBackground(bg).setFontSize(10).setVerticalAlignment('middle');
        // Destacar columna "Sin facturar" y "Estado" en color semáforo
        sh.getRange(currentRow, 4).setFontWeight('bold').setFontColor(claro);
        sh.getRange(currentRow, 7).setFontWeight('bold').setFontColor(claro);
        sh.setRowHeight(currentRow, 28);
        currentRow++;
      }
    }

    currentRow++; // spacer entre secciones
    sh.setRowHeight(currentRow - 1, 8);

    // ── Sección: clientes AL DÍA ──────────────────────────────────
    if (sinPend.length > 0) {
      sh.getRange(currentRow, 2, 1, 6).merge()
        .setValue('✅  CLIENTES AL DÍA (' + sinPend.length + ')')
        .setBackground('#34a853').setFontColor('#ffffff')
        .setFontSize(11).setFontWeight('bold')
        .setHorizontalAlignment('left').setVerticalAlignment('middle');
      sh.setRowHeight(currentRow, 32);
      currentRow++;

      // Cabecera tabla verde
      sh.getRange(currentRow, 2, 1, 6).setValues([[
        'Cliente', 'Total trabajos', 'Sin facturar', '', '', 'Estado'
      ]]).setBackground('#2d8e47').setFontColor('#ffffff')
        .setFontSize(9).setFontWeight('bold').setVerticalAlignment('middle');
      sh.setRowHeight(currentRow, 24);
      currentRow++;

      for (var si2 = 0; si2 < sinPend.length; si2++) {
        var sp  = sinPend[si2];
        var bg2 = si2 % 2 === 0 ? '#e8f5e9' : '#f1f8f4';
        sh.getRange(currentRow, 2, 1, 6).setValues([[
          sp.cliente, sp.total, 0, '—', '—', '✅ Al día'
        ]]).setBackground(bg2).setFontSize(10).setVerticalAlignment('middle');
        sh.getRange(currentRow, 4).setFontColor('#9aa0a6');
        sh.getRange(currentRow, 7).setFontWeight('bold').setFontColor('#2d8e47');
        sh.setRowHeight(currentRow, 24);
        currentRow++;
      }
    }

    // ── Estado final ──────────────────────────────────────────────
    var ts = new Date();
    _ppEstado_(sh, '✅  Actualizado ' +
      Utilities.formatDate(ts, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') +
      '  ·  ' + conPend.length + ' cliente(s) con pendientes  ·  ' + kpiTotalPend + ' trabajo(s) sin facturar',
      '#34a853');

    logEvent('INFO', 'refreshPanelPendientes',
      'conPend=' + conPend.length + ' | totalSinFac=' + kpiTotalPend + ' | horas=' + kpiHorasPend.toFixed(1), '');

  } catch (ex) {
    _ppEstado_(sh, '⚠  Error: ' + ex.message, '#ea4335');
    logEvent('ERROR', 'refreshPanelPendientes', ex.message, ex.stack || '');
  }
}
