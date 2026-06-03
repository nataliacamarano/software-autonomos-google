// ============================================================
// PanelCliente.gs — Vista consolidada por cliente (Fase D)
// ============================================================

var PC_ = {
  CLIENTE:        [4, 3],
  FECHA_INI:      [5, 3],
  FECHA_FIN:      [6, 3],
  BOTON:          [8, 3],
  ESTADO:         [9, 2],
  KPI_ROW_LBL:    11,
  KPI_ROW_VAL:    12,
  // Resumen ampliado: filas 15-25
  RES_NOMBRE:     [15, 3],
  RES_TEL:        [16, 3],
  RES_EMAIL:      [17, 3],
  RES_PRIMERO:    [18, 3],
  RES_ULTIMO:     [19, 3],
  RES_HRS_FAC:    [20, 3],
  RES_HRS_PTE:    [21, 3],
  RES_HIST_TOTAL: [22, 3],
  RES_HIST_COB:   [23, 3],
  RES_HIST_PTE:   [24, 3],
  RES_EHR:        [25, 3],
  DATA_START:     27
};

function _r5_(v) { return [v, v, v, v, v]; }

var KPI_COLORS_ = {
  FAC:       ['#34a853', '#2d8e47'],
  PTE:       ['#f9ab00', '#e09200'],
  PTE_ALERT: ['#ea4335', '#c62828'],
  FRA:       ['#1a73e8', '#1565c0'],
  TRA:       ['#00695c', '#004d40'],
  HRS:       ['#7b1fa2', '#6a1b9a']
};

// ================================================================
// FASE 1 — Crear hoja, columnas, filas 1-10
// ================================================================
function setupPanelClienteFase1(silent) {
  var FASE = 'Fase 1';
  try {
    var ss = SpreadsheetApp.getActive();
    if (!ss) throw new Error('SpreadsheetApp.getActive() es null.');

    var old = ss.getSheetByName('PANEL_CLIENTE');
    if (old) { try { ss.deleteSheet(old); } catch(e) {} }
    var sh = ss.insertSheet('PANEL_CLIENTE');
    ss.setActiveSheet(sh);
    SpreadsheetApp.flush();

    sh.setColumnWidth(1,  10);  // col A — margen izquierdo (vacío, sin datos técnicos)
    sh.setColumnWidth(2, 155);
    sh.setColumnWidth(3, 155);
    sh.setColumnWidth(4, 155);
    sh.setColumnWidth(5, 155);
    sh.setColumnWidth(6, 155);
    sh.setColumnWidth(7, 100);  // col G — Informe
    sh.setColumnWidth(8, 100);  // col H — Proforma
    sh.setColumnWidth(9, 100);  // col I — Factura
    sh.setColumnWidth(10, 55);  // col J — 🔍 Ver

    sh.setRowHeight(1, 50); sh.setRowHeight(2,  8); sh.setRowHeight(3, 26);
    sh.setRowHeight(4, 38); sh.setRowHeight(5, 38); sh.setRowHeight(6, 38);
    sh.setRowHeight(7,  8); sh.setRowHeight(8, 46); sh.setRowHeight(9, 34);
    sh.setRowHeight(10, 8);

    var E = _r5_(''), W = '#ffffff', LBL = '#f8f9fa', DAT = '#e8f0fe';

    sh.getRange(1, 2, 10, 5).setValues([
      ['PANEL CLIENTE','','','',''],
      E, ['FILTROS','','','',''],
      ['Cliente','','','',''], ['Fecha desde','','','',''], ['Fecha hasta','','','',''],
      E, ['ACTUALIZAR PANEL CLIENTE','','','',''],
      ['Selecciona cliente y fechas, luego activa el checkbox.','','','',''], E
    ]);

    sh.getRange(1, 2, 10, 5).setBackgrounds([
      _r5_('#1a73e8'), _r5_(W), _r5_(DAT),
      [LBL,W,W,W,W], [LBL,DAT,W,W,W], [LBL,DAT,W,W,W],
      _r5_(W), ['#1a73e8','#1765cc',W,W,W], _r5_(LBL), _r5_(W)
    ]);
    sh.getRange(1, 2, 10, 5).setFontColors([
      _r5_('#ffffff'), _r5_('#3c4043'), _r5_('#1a73e8'),
      _r5_('#3c4043'), _r5_('#3c4043'), _r5_('#3c4043'),
      _r5_('#3c4043'), _r5_('#ffffff'), _r5_('#9aa0a6'), _r5_('#3c4043')
    ]);
    sh.getRange(1, 2, 10, 5).setFontWeights([
      _r5_('bold'), _r5_('normal'), _r5_('bold'),
      ['bold','normal','normal','normal','normal'], ['bold','normal','normal','normal','normal'],
      ['bold','normal','normal','normal','normal'], _r5_('normal'),
      ['bold','normal','normal','normal','normal'], _r5_('normal'), _r5_('normal')
    ]);
    sh.getRange(1, 2, 10, 5).setFontSizes([
      _r5_(18), _r5_(10), _r5_(9), _r5_(11), _r5_(10), _r5_(10),
      _r5_(10), _r5_(12), _r5_(10), _r5_(10)
    ]);
    sh.getRange(1, 2, 10, 5).setHorizontalAlignments([
      _r5_('center'), _r5_('left'), _r5_('center'), _r5_('left'), _r5_('left'), _r5_('left'),
      _r5_('left'), _r5_('left'), _r5_('center'), _r5_('left')
    ]);
    sh.getRange(1, 2, 10, 5).setVerticalAlignments([
      _r5_('middle'),_r5_('middle'),_r5_('middle'),_r5_('middle'),_r5_('middle'),
      _r5_('middle'),_r5_('middle'),_r5_('middle'),_r5_('middle'),_r5_('middle')
    ]);

    SpreadsheetApp.flush();
    sh.getRange(1, 2, 1, 5).merge();
    sh.getRange(3, 2, 1, 5).merge();
    sh.getRange(9, 2, 1, 5).merge();
    SpreadsheetApp.flush();
    ss.setActiveSheet(sh);

    if (!silent) SpreadsheetApp.getUi().alert('Fase 1 OK\n\nEjecuta ahora: Sistema → PANEL_CLIENTE Fase 2');
  } catch(e) {
    SpreadsheetApp.getUi().alert('Error en ' + FASE + ':\n\n' + e.message);
    if (silent) throw e;
  }
}

// ================================================================
// FASE 2 — KPI strip (11-12) + Resumen ampliado (14-25)
// ================================================================
function setupPanelClienteFase2(silent) {
  var FASE = 'Fase 2';
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName('PANEL_CLIENTE');
    if (!sh) throw new Error('PANEL_CLIENTE no existe. Ejecuta primero Fase 1.');

    sh.setRowHeight(11, 28); sh.setRowHeight(12, 52); sh.setRowHeight(13,  8);
    sh.setRowHeight(14, 30);
    for (var r = 15; r <= 25; r++) sh.setRowHeight(r, 26);
    sh.setRowHeight(26, 8);

    var E = _r5_(''), W = '#ffffff', LBL = '#f8f9fa';
    var FW = '#ffffff', FD = '#5f6368', FB = '#3c4043';

    sh.getRange(11, 2, 16, 5).setValues([
      ['💰 Facturado','⏳ Pendiente cobro','🧾 Facturas','🔧 Trabajos','⏱ Horas'],
      ['—','—','—','—','—'], E,
      ['RESUMEN CLIENTE','','','',''],
      ['Cliente','','','',''], ['Teléfono','','','',''], ['Email','','','',''],
      ['Primer trabajo','','','',''], ['Último trabajo','','','',''],
      ['Horas facturadas','','','',''], ['Horas pendientes','','','',''],
      ['Total histórico','','','',''], ['Pagado histórico','','','',''],
      ['Pendiente histórico','','','',''], ['Promedio €/hora','','','',''], E
    ]);
    sh.getRange(11, 2, 16, 5).setBackgrounds([
      ['#34a853','#f9ab00','#1a73e8','#00695c','#7b1fa2'],
      ['#2d8e47','#e09200','#1565c0','#004d40','#6a1b9a'],
      _r5_(W), _r5_('#34a853'),
      [LBL,W,W,W,W],[LBL,W,W,W,W],[LBL,W,W,W,W],
      [LBL,W,W,W,W],[LBL,W,W,W,W],[LBL,W,W,W,W],[LBL,W,W,W,W],
      [LBL,W,W,W,W],[LBL,W,W,W,W],[LBL,W,W,W,W],[LBL,W,W,W,W], _r5_(W)
    ]);
    sh.getRange(11, 2, 16, 5).setFontColors([
      _r5_(FW), _r5_(FW), _r5_(FB), _r5_(FW),
      [FD,FB,FB,FB,FB],[FD,FB,FB,FB,FB],[FD,FB,FB,FB,FB],
      [FD,FB,FB,FB,FB],[FD,FB,FB,FB,FB],[FD,FB,FB,FB,FB],[FD,FB,FB,FB,FB],
      [FD,FB,FB,FB,FB],[FD,FB,FB,FB,FB],[FD,FB,FB,FB,FB],[FD,FB,FB,FB,FB], _r5_(FB)
    ]);
    sh.getRange(11, 2, 16, 5).setFontWeights([
      _r5_('bold'), _r5_('bold'), _r5_('normal'), _r5_('bold'),
      ['bold','normal','normal','normal','normal'],['bold','normal','normal','normal','normal'],
      ['bold','normal','normal','normal','normal'],['bold','normal','normal','normal','normal'],
      ['bold','normal','normal','normal','normal'],['bold','normal','normal','normal','normal'],
      ['bold','normal','normal','normal','normal'],['bold','normal','normal','normal','normal'],
      ['bold','normal','normal','normal','normal'],['bold','normal','normal','normal','normal'],
      ['bold','normal','normal','normal','normal'], _r5_('normal')
    ]);
    sh.getRange(11, 2, 16, 5).setFontSizes([
      _r5_(9), _r5_(18), _r5_(10), _r5_(11),
      _r5_(10),_r5_(10),_r5_(10),_r5_(10),_r5_(10),
      _r5_(10),_r5_(10),_r5_(10),_r5_(10),_r5_(10), _r5_(10), _r5_(10)
    ]);
    sh.getRange(11, 2, 16, 5).setHorizontalAlignments([
      _r5_('center'), _r5_('center'), _r5_('left'), _r5_('center'),
      _r5_('left'),_r5_('left'),_r5_('left'),_r5_('left'),_r5_('left'),
      _r5_('left'),_r5_('left'),_r5_('left'),_r5_('left'),_r5_('left'), _r5_('left'), _r5_('left')
    ]);
    var vas2 = []; for (var i2 = 0; i2 < 16; i2++) vas2.push(_r5_('middle'));
    sh.getRange(11, 2, 16, 5).setVerticalAlignments(vas2);

    SpreadsheetApp.flush();
    sh.getRange(14, 2, 1, 5).merge();
    for (var mr = 15; mr <= 25; mr++) sh.getRange(mr, 3, 1, 4).merge();
    SpreadsheetApp.flush();

    if (!silent) SpreadsheetApp.getUi().alert('Fase 2 OK\n\nEjecuta ahora: Sistema → PANEL_CLIENTE Fase 3');
  } catch(e) {
    SpreadsheetApp.getUi().alert('Error en ' + FASE + ':\n\n' + e.message);
    if (silent) throw e;
  }
}

// ================================================================
// FASE 3 — Secciones placeholder desde DATA_START (fila 27)
// ================================================================
function setupPanelClienteFase3(silent) {
  var FASE = 'Fase 3';
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName('PANEL_CLIENTE');
    if (!sh) throw new Error('PANEL_CLIENTE no existe. Ejecuta primero Fase 1.');

    var S = PC_.DATA_START;
    var secciones = [
      ['TRABAJOS DEL PERIODO',    '#e65100'],
      ['FACTURAS',                '#1a73e8'],
      ['PROFORMAS',               '#7b1fa2'],
      ['INFORMES',                '#00695c'],
      ['COBROS PENDIENTES',       '#b71c1c'],
      ['HISTÓRICO DEL CLIENTE',   '#37474f'],
      ['HISTORIAL CRONOLÓGICO',   '#455a64']
    ];
    for (var s = 0; s < secciones.length; s++) {
      sh.setRowHeight(S + s * 3,     32);
      sh.setRowHeight(S + s * 3 + 1, 26);
      sh.setRowHeight(S + s * 3 + 2,  8);
    }

    var vals = [], bgs = [], fcs = [], fws = [], fss = [], has = [], vas = [];
    var FA = '#fafafa', W = '#ffffff';
    var PLACEHOLDER = 'Pulsa ACTUALIZAR para cargar datos.';
    for (var si = 0; si < secciones.length; si++) {
      var col = secciones[si][1];
      vals.push([secciones[si][0],'','','','']); vals.push([PLACEHOLDER,'','','','']); vals.push(['','','','','']);
      bgs.push(_r5_(col)); bgs.push(_r5_(FA)); bgs.push(_r5_(W));
      fcs.push(_r5_('#ffffff')); fcs.push(_r5_('#9aa0a6')); fcs.push(_r5_('#3c4043'));
      fws.push(_r5_('bold'));   fws.push(_r5_('normal')); fws.push(_r5_('normal'));
      fss.push(_r5_(11));       fss.push(_r5_(10));       fss.push(_r5_(10));
      has.push(_r5_('left'));   has.push(_r5_('left'));   has.push(_r5_('left'));
      vas.push(_r5_('middle')); vas.push(_r5_('middle')); vas.push(_r5_('middle'));
    }
    sh.getRange(S, 2, secciones.length * 3, 5)
      .setValues(vals).setBackgrounds(bgs).setFontColors(fcs)
      .setFontWeights(fws).setFontSizes(fss)
      .setHorizontalAlignments(has).setVerticalAlignments(vas);

    SpreadsheetApp.flush();
    for (var mi = 0; mi < secciones.length; mi++) {
      sh.getRange(S + mi * 3,     2, 1, 5).merge();
      sh.getRange(S + mi * 3 + 1, 2, 1, 5).merge();
    }
    SpreadsheetApp.flush();

    if (!silent) SpreadsheetApp.getUi().alert('Fase 3 OK\n\nEjecuta ahora: Sistema → PANEL_CLIENTE Fase 4');
  } catch(e) {
    SpreadsheetApp.getUi().alert('Error en ' + FASE + ':\n\n' + e.message);
    if (silent) throw e;
  }
}

// ================================================================
// FASE 4 — Validaciones, dropdown y checkbox ACTUALIZAR
// ================================================================
function setupPanelClienteFase4(silent) {
  var FASE = 'Fase 4';
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName('PANEL_CLIENTE');
    if (!sh) throw new Error('PANEL_CLIENTE no existe. Ejecuta primero Fase 1.');

    sh.getRange(PC_.BOTON[0], PC_.BOTON[1])
      .insertCheckboxes().setValue(false)
      .setBackground('#1765cc').setHorizontalAlignment('center').setVerticalAlignment('middle');

    try {
      var names = _panelGetClienteNames_(ss);
      if (names.length > 0) {
        sh.getRange(PC_.CLIENTE[0], PC_.CLIENTE[1]).setDataValidation(
          SpreadsheetApp.newDataValidation().requireValueInList(names, true).setAllowInvalid(false).build()
        );
      }
    } catch(eV) { Logger.log('[Fase4] Dropdown: ' + eV.message); }

    try {
      var dv = SpreadsheetApp.newDataValidation()
        .requireDateOnOrAfter(new Date(1900, 0, 1)).setAllowInvalid(false).build();
      sh.getRange(PC_.FECHA_INI[0], PC_.FECHA_INI[1]).setDataValidation(dv).setNumberFormat('dd/MM/yyyy');
      sh.getRange(PC_.FECHA_FIN[0], PC_.FECHA_FIN[1]).setDataValidation(dv).setNumberFormat('dd/MM/yyyy');
    } catch(eD) { Logger.log('[Fase4] Fechas: ' + eD.message); }

    SpreadsheetApp.flush();
    ss.setActiveSheet(sh);

    if (!silent) SpreadsheetApp.getUi().alert(
      'Fase 4 OK — PANEL_CLIENTE completamente configurado.\n\n' +
      'Selecciona cliente y fechas, luego activa el checkbox ACTUALIZAR.'
    );
  } catch(e) {
    SpreadsheetApp.getUi().alert('Error en ' + FASE + ':\n\n' + e.message);
    if (silent) throw e;
  }
}

// ================================================================
// CREAR COMPLETO
// ================================================================
function setupPanelClienteCompleto() {
  var t0 = new Date().getTime();
  var fases = [
    { n: 1, fn: setupPanelClienteFase1 }, { n: 2, fn: setupPanelClienteFase2 },
    { n: 3, fn: setupPanelClienteFase3 }, { n: 4, fn: setupPanelClienteFase4 }
  ];
  for (var i = 0; i < fases.length; i++) {
    var t1 = new Date().getTime();
    try { fases[i].fn(true); } catch(e) { return; }
    if (new Date().getTime() - t1 > 25000) {
      SpreadsheetApp.getUi().alert('Fase ' + fases[i].n + ' tardó demasiado.\nEjecuta las fases restantes manualmente.');
      return;
    }
  }
  var total = ((new Date().getTime() - t0) / 1000).toFixed(1);
  SpreadsheetApp.getUi().alert('PANEL_CLIENTE creado en ' + total + 's.\n\nSelecciona cliente y fechas, luego activa el checkbox ACTUALIZAR.');
}

// ================================================================
// REFRESH — Carga y muestra todos los datos del cliente
// ================================================================
function refreshPanelCliente() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('PANEL_CLIENTE');
  if (!sh) {
    SpreadsheetApp.getUi().alert('PANEL_CLIENTE no existe.\nEjecuta: Sistema → PANEL_CLIENTE Crear completo');
    return;
  }

  var cliente    = String(sh.getRange(PC_.CLIENTE[0],   PC_.CLIENTE[1]).getValue()   || '').trim();
  var fechaIniDv = sh.getRange(PC_.FECHA_INI[0], PC_.FECHA_INI[1]).getDisplayValue().trim();
  var fechaFinDv = sh.getRange(PC_.FECHA_FIN[0], PC_.FECHA_FIN[1]).getDisplayValue().trim();

  var err = [];
  if (!cliente)                       err.push('Selecciona un cliente');
  if (!_parseReportDate_(fechaIniDv)) err.push('Fecha desde inválida (DD/MM/YYYY)');
  if (!_parseReportDate_(fechaFinDv)) err.push('Fecha hasta inválida (DD/MM/YYYY)');
  if (err.length) { _pcEstado_(sh, '⚠  ' + err.join('  ·  '), '#ffffff', '#ea4335'); return; }

  var fi = _parseReportDate_(fechaIniDv);
  var ff = _parseReportDate_(fechaFinDv);
  if (fi > ff) { _pcEstado_(sh, '⚠  Fecha desde es posterior a Fecha hasta', '#ffffff', '#ea4335'); return; }
  ff.setHours(23, 59, 59, 999);
  var hoy = new Date();

  _pcEstado_(sh, 'Cargando datos de ' + cliente + '...', '#ffffff', '#f9ab00');
  SpreadsheetApp.flush();

  try {
    // ── 1. TRABAJOS ──────────────────────────────────────────────
    var trabajos    = [];
    var allTrabajos = [];
    var allTrabajosMap = {};
    try {
      var shTra  = getSheet('TRABAJOS');
      var lastTr = shTra.getLastRow();
      if (lastTr > 1) {
        var numColsTra = Math.min(shTra.getLastColumn(), COL_TRA.MATERIALES);
        var dataTra    = shTra.getRange(2, 1, lastTr - 1, numColsTra).getValues();
        for (var i = 0; i < dataTra.length; i++) {
          var trow = dataTra[i];
          if (String(trow[COL_TRA.CLIENTE - 1] || '').trim().toLowerCase() !== cliente.toLowerCase()) continue;
          var ft         = _parseTrabajosDate_(trow[COL_TRA.FECHA_TRABAJO - 1]);
          var tId        = String(trow[COL_TRA.ID - 1] || '').trim();
          var horasTra   = _calcHoras_(trow[COL_TRA.HORA_INICIO - 1], trow[COL_TRA.HORA_FIN - 1]);
          var trabajoStr = String(trow[COL_TRA.TRABAJO_REALIZADO - 1] || '').trim();
          if (tId && ft) {
            allTrabajosMap[tId] = { fecha: ft, trabajo: trabajoStr, horas: horasTra };
            allTrabajos.push({ id: tId, fecha: ft, horas: horasTra, trabajo: trabajoStr });
          }
          if (!ft || ft < fi || ft > ff) continue;
          trabajos.push({
            id:        tId,
            fecha:     ft,
            operario:  String(trow[COL_TRA.OPERARIO - 1]          || '').trim(),
            trabajo:   trabajoStr,
            horas:     horasTra,
            estado:    String(trow[COL_TRA.ESTADO - 1]            || '').trim(),
            materiales: Number(trow[COL_TRA.MATERIALES - 1]       || 0)
          });
        }
        trabajos.sort(function(a, b) { return a.fecha - b.fecha; });
      }
    } catch (eTra) { Logger.log('[PC] TRABAJOS: ' + eTra.message); }

    // ── 2. PRESUPUESTOS (lineas JSON) ─────────────────────────────
    var presupuestosMap = {};
    try {
      var shPre    = getSheet('PRESUPUESTOS');
      var lastPreR = shPre.getLastRow();
      if (lastPreR > 1) {
        var numColsPre = Math.min(shPre.getLastColumn(), COL_PRE.LINEAS_JSON);
        var dataPre    = shPre.getRange(2, 1, lastPreR - 1, numColsPre).getValues();
        for (var pi = 0; pi < dataPre.length; pi++) {
          var prow    = dataPre[pi];
          var preId   = String(prow[COL_PRE.ID - 1] || '').trim();
          var jsonStr = String(prow[COL_PRE.LINEAS_JSON - 1] || '').trim();
          if (preId && jsonStr && jsonStr !== '[]') {
            try { var lineas = JSON.parse(jsonStr); if (Array.isArray(lineas) && lineas.length) presupuestosMap[preId] = lineas; } catch(_) {}
          }
        }
      }
    } catch (ePre) { Logger.log('[PC] PRESUPUESTOS: ' + ePre.message); }

    // ── 3. FACTURAS ──────────────────────────────────────────────
    var facturas    = [];
    var allFacturas = [];
    var colFacEstadoCobro = -1;
    try {
      var shFac  = getSheet('FACTURAS');
      var lastFr = shFac.getLastRow();
      if (lastFr > 1) {
        var facHdrs = shFac.getRange(1, 1, 1, shFac.getLastColumn()).getValues()[0];
        for (var fh = 0; fh < facHdrs.length; fh++) {
          if (String(facHdrs[fh]).trim().toLowerCase().replace(/_/g, ' ') === 'estado cobro') { colFacEstadoCobro = fh; break; }
        }
        var dataFac = shFac.getRange(2, 1, lastFr - 1, facHdrs.length).getValues();
        for (var j = 0; j < dataFac.length; j++) {
          var frow  = dataFac[j];
          if (String(frow[5] || '').trim().toLowerCase() !== cliente.toLowerCase()) continue;
          var fDate = frow[1] instanceof Date ? frow[1] : _parseReportDate_(String(frow[1] || ''));
          if (!fDate) continue;
          var ecFac = colFacEstadoCobro >= 0 ? String(frow[colFacEstadoCobro] || '').trim() : '';
          if (!ecFac) ecFac = String(frow[17] || '').trim() === 'Emitida' ? 'Pendiente' : String(frow[17] || '').trim();
          var facObj = {
            id: String(frow[0] || ''), fecha: fDate, total: Number(frow[16] || 0),
            estado: String(frow[17] || '').trim(), estadoCobro: ecFac,
            pdfUrl: String(frow[18] || '').trim(),
            idsTrabajos: String(frow[9]  || '').trim(),
            idPresup:    String(frow[21] || '').trim()
          };
          allFacturas.push(facObj);
          if (fDate >= fi && fDate <= ff) facturas.push(facObj);
        }
        facturas.sort(function(a, b) { return a.fecha - b.fecha; });
      }
    } catch (eFac) { Logger.log('[PC] FACTURAS: ' + eFac.message); }

    // ── 4. PROFORMAS ─────────────────────────────────────────────
    var proformas    = [];
    var allProformas = [];
    try {
      var shPro  = getSheet('PROFORMAS');
      var lastPr = shPro.getLastRow();
      if (lastPr > 1) {
        var numColsPro = shPro.getLastColumn();
        var dataPro    = shPro.getRange(2, 1, lastPr - 1, numColsPro).getValues();
        for (var k = 0; k < dataPro.length; k++) {
          var prow2 = dataPro[k];
          if (String(prow2[5] || '').trim().toLowerCase() !== cliente.toLowerCase()) continue;
          var pDate = prow2[1] instanceof Date ? prow2[1] : _parseReportDate_(String(prow2[1] || ''));
          if (!pDate) continue;
          var proObj = {
            id: String(prow2[0] || ''), fecha: pDate, total: Number(prow2[12] || 0),
            estadoCobro: String(prow2[18] || 'Pendiente de cobro').trim(),
            pdfUrl: String(prow2[14] || '').trim(),
            idsTrabajos: String(prow2[9]  || '').trim(),
            idPresup:    String(prow2[17] || '').trim()
          };
          allProformas.push(proObj);
          if (pDate >= fi && pDate <= ff) proformas.push(proObj);
        }
        proformas.sort(function(a, b) { return a.fecha - b.fecha; });
      }
    } catch (ePro) { Logger.log('[PC] PROFORMAS: ' + ePro.message); }

    // ── 5. INFORMES ──────────────────────────────────────────────
    var informes = [];
    var allInformes = [];
    try {
      var shInf  = getSheet('INFORMES');
      var lastIr = shInf.getLastRow();
      if (lastIr > 1) {
        var numColsInf = shInf.getLastColumn();
        var dataInf    = shInf.getRange(2, 1, lastIr - 1, numColsInf).getValues();
        for (var m = 0; m < dataInf.length; m++) {
          var irow  = dataInf[m];
          if (String(irow[3] || '').trim().toLowerCase() !== cliente.toLowerCase()) continue;
          var iDate = irow[1] instanceof Date ? irow[1] : _parseReportDate_(String(irow[1] || ''));
          if (!iDate) continue;
          var infObj = {
            id:          String(irow[0] || ''),
            fecha:       iDate,
            pdfUrl:      String(irow[8] || '').trim(),
            idsTrabajos: numColsInf >= 11 ? String(irow[10] || '').trim() : ''
          };
          allInformes.push(infObj);
          if (iDate >= fi && iDate <= ff) informes.push(infObj);
        }
        informes.sort(function(a, b) { return a.fecha - b.fecha; });
      }
    } catch (eInf) { Logger.log('[PC] INFORMES: ' + eInf.message); }

    // ── 6. Índices de facturación + detección de duplicados ──────
    var facturadosSet = {}, docPorTrabajo = {}, trabajoDocsMap = {};

    // Registra un trabajo en los tres índices; acumula todos los docs para detectar duplicados
    var _indexarTrabajo_ = function(tid, docId) {
      if (!tid || !docId) return;
      facturadosSet[tid] = true;
      if (!docPorTrabajo[tid]) docPorTrabajo[tid] = docId; // primer doc encontrado
      if (!trabajoDocsMap[tid]) trabajoDocsMap[tid] = [];
      if (trabajoDocsMap[tid].indexOf(docId) < 0) trabajoDocsMap[tid].push(docId);
    };

    var procesarIds = function(docs) {
      for (var d = 0; d < docs.length; d++) {
        if (!docs[d].idsTrabajos) continue;
        var ids = docs[d].idsTrabajos.split(',');
        for (var x = 0; x < ids.length; x++) {
          _indexarTrabajo_(ids[x].trim(), docs[d].id);
        }
      }
    };
    procesarIds(allFacturas);
    procesarIds(allProformas);

    // Mapas separados por tipo de documento (para columnas Informe / Proforma / Factura)
    var facturaPorTrabajo  = {};
    var proformaPorTrabajo = {};
    var informePorTrabajo  = {};
    (function() {
      var _idx_ = function(docs, map) {
        for (var d = 0; d < docs.length; d++) {
          if (!docs[d].idsTrabajos) continue;
          var ids = docs[d].idsTrabajos.split(',');
          for (var x = 0; x < ids.length; x++) {
            var tid = ids[x].trim();
            if (tid && !map[tid]) map[tid] = docs[d].id;
          }
        }
      };
      _idx_(allFacturas,  facturaPorTrabajo);
      _idx_(allProformas, proformaPorTrabajo);
      _idx_(allInformes,  informePorTrabajo);
    })();

    // Trabajos que aparecen en más de un documento (duplicados reales)
    var trabajosDuplicados = [];
    for (var dupTid in trabajoDocsMap) {
      if (trabajoDocsMap[dupTid].length > 1) {
        trabajosDuplicados.push(dupTid + ' → ' + trabajoDocsMap[dupTid].join(', '));
      }
    }

    // ── 7. Cálculos del periodo ───────────────────────────────────
    var totalHoras     = trabajos.reduce(function(s, t) { return s + t.horas; }, 0);
    var totalFacturado = facturas.reduce(function(s, f) { return s + f.total; }, 0);
    var pendFact = facturas.filter(function(f) {
      return f.estadoCobro === 'Pendiente' || f.estadoCobro === 'Pendiente de cobro';
    });
    var pendPro = proformas.filter(function(p) {
      var ec = (p.estadoCobro || '').toLowerCase();
      return ec !== 'pagada' && ec !== 'cobrado' && ec !== 'pagado' && ec !== 'anulada';
    });
    var totalPendiente = pendFact.reduce(function(s, f) { return s + f.total; }, 0)
                       + pendPro.reduce(function(s, p)  { return s + p.total; }, 0);
    var cobradoPeriodo = facturas.filter(function(f) {
      return f.estadoCobro === 'Cobrado' || f.estadoCobro === 'Pagado' || f.estadoCobro === 'Pagada';
    }).reduce(function(s, f) { return s + f.total; }, 0)
    + proformas.filter(function(p) {
      return p.estadoCobro === 'Pagada' || p.estadoCobro === 'Cobrado' || p.estadoCobro === 'Pagado';
    }).reduce(function(s, p) { return s + p.total; }, 0);

    // ── 8. Cálculos históricos ────────────────────────────────────
    var histHorasTotal = allTrabajos.reduce(function(s, t) { return s + t.horas; }, 0);
    var histHorasFacturadas = 0;
    for (var hi = 0; hi < allTrabajos.length; hi++) {
      if (facturadosSet[allTrabajos[hi].id]) histHorasFacturadas += allTrabajos[hi].horas;
    }
    var histHorasPendientes = Math.max(0, histHorasTotal - histHorasFacturadas);
    var histTotalFacturado  = allFacturas.reduce(function(s, f) { return s + f.total; }, 0)
                            + allProformas.reduce(function(s, p) { return s + p.total; }, 0);
    var histCobrado = allFacturas.filter(function(f) {
      return f.estadoCobro === 'Cobrado' || f.estadoCobro === 'Pagado' || f.estadoCobro === 'Pagada';
    }).reduce(function(s, f) { return s + f.total; }, 0)
    + allProformas.filter(function(p) {
      return p.estadoCobro === 'Pagada' || p.estadoCobro === 'Cobrado' || p.estadoCobro === 'Pagado';
    }).reduce(function(s, p) { return s + p.total; }, 0);
    var histPendiente = Math.max(0, histTotalFacturado - histCobrado);
    var promedioEHR   = histHorasTotal > 0 ? histTotalFacturado / histHorasTotal : 0;

    var primerTrabajoAll = allTrabajos.length > 0
      ? allTrabajos.reduce(function(a, b) { return a.fecha <= b.fecha ? a : b; }) : null;
    var ultimoTrabajoAll = allTrabajos.length > 0
      ? allTrabajos.reduce(function(a, b) { return a.fecha >= b.fecha ? a : b; }) : null;
    var primerFacAll = allFacturas.length > 0
      ? allFacturas.reduce(function(a, b) { return a.fecha <= b.fecha ? a : b; }) : null;
    var ultimaFacAll = allFacturas.length > 0
      ? allFacturas.reduce(function(a, b) { return a.fecha >= b.fecha ? a : b; }) : null;

    // ── 9. KPIs ───────────────────────────────────────────────────
    var docsText = facturas.length + ' F  ·  ' + proformas.length + ' P  ·  ' + informes.length + ' I';
    sh.getRange(PC_.KPI_ROW_LBL, 2, 1, 5).setValues([[
      '🧾 Facturas emitidas', '✅ Pagado', '⏳ Pendiente cobro', '📄 Documentos', '🔧 Trabajos'
    ]]).setFontColor('#ffffff').setFontWeight('bold').setFontSize(9)
       .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(PC_.KPI_ROW_VAL, 2, 1, 5).setValues([[
      formatCurrency(totalFacturado), formatCurrency(cobradoPeriodo),
      formatCurrency(totalPendiente), docsText, trabajos.length
    ]]);
    var kpiLblBgs = totalPendiente > 0
      ? ['#34a853', '#1a73e8', KPI_COLORS_.PTE_ALERT[0], '#7b1fa2', '#00695c']
      : ['#34a853', '#1a73e8', KPI_COLORS_.PTE[0],        '#7b1fa2', '#00695c'];
    var kpiValBgs = totalPendiente > 0
      ? ['#2d8e47', '#1565c0', KPI_COLORS_.PTE_ALERT[1], '#6a1b9a', '#004d40']
      : ['#2d8e47', '#1565c0', KPI_COLORS_.PTE[1],        '#6a1b9a', '#004d40'];
    sh.getRange(PC_.KPI_ROW_LBL, 2, 1, 5).setBackgrounds([kpiLblBgs]);
    sh.getRange(PC_.KPI_ROW_VAL, 2, 1, 5).setBackgrounds([kpiValBgs]);

    // ── 10. Barra de estado ──────────────────────────────────────
    if (totalPendiente > 0) {
      _pcEstado_(sh, '⚠  CLIENTE CON ' + formatCurrency(totalPendiente) + ' PENDIENTES DE PAGO', '#ffffff', '#ea4335');
    } else {
      _pcEstado_(sh,
        cliente + '  ·  ' + fechaIniDv + ' — ' + fechaFinDv +
        '  ·  ' + trabajos.length + ' trabajo(s)' +
        (facturas.length  ? '  ·  ' + facturas.length  + ' factura(s)'  : '') +
        (proformas.length ? '  ·  ' + proformas.length + ' proforma(s)' : '') +
        (informes.length  ? '  ·  ' + informes.length  + ' informe(s)'  : ''),
        '#3c4043', '#e8f5e9'
      );
    }

    // ── 11. Resumen ampliado (2 API calls) ────────────────────────
    var cliInfo = null; try { cliInfo = _getClienteByLabel_(cliente); } catch(eCli) {}
    sh.getRange(15, 2, 11, 1).setValues([
      ['Cliente'], ['Teléfono'], ['Email'], ['Primer trabajo'], ['Último trabajo'],
      ['Horas facturadas'], ['Horas pendientes'], ['Total histórico'],
      ['Pagado histórico'], ['Pendiente histórico'], ['Promedio €/hora']
    ]);
    sh.getRange(15, 3, 11, 1).setValues([
      [cliInfo ? cliInfo.nombre              : cliente],
      [cliInfo ? (cliInfo.telefono || '—')  : '—'],
      [cliInfo ? (cliInfo.email    || '—')  : '—'],
      [primerTrabajoAll ? formatDate(primerTrabajoAll.fecha) + '  —  ' + primerTrabajoAll.trabajo.substring(0, 50) : '—'],
      [ultimoTrabajoAll ? formatDate(ultimoTrabajoAll.fecha) + '  —  ' + ultimoTrabajoAll.trabajo.substring(0, 50) : '—'],
      [histHorasFacturadas.toFixed(1).replace('.', ',') + ' h'],
      [histHorasPendientes.toFixed(1).replace('.', ',') + ' h'],
      [formatCurrency(histTotalFacturado)],
      [formatCurrency(histCobrado)],
      [formatCurrency(histPendiente)],
      [promedioEHR > 0 ? formatCurrency(promedioEHR) + ' / h' : '—']
    ]);

    // ── 12. Limpiar secciones dinámicas ──────────────────────────
    var lastUsedRow = Math.max(PC_.DATA_START, sh.getLastRow());
    sh.getRange(PC_.DATA_START, 1, lastUsedRow - PC_.DATA_START + 50, 11).breakApart().clear();
    // Targeted clear for col J: insertCheckboxes() leaves data validation without a value,
    // so getLastRow() may not count those rows — explicit wide clear removes orphan checkboxes.
    sh.getRange(PC_.DATA_START, 10, 300, 1).clear();
    SpreadsheetApp.flush();
    var currentRow = PC_.DATA_START;

    // ── DATOS DE PAGO ─────────────────────────────────────────────
    currentRow = _pcSectionHeader_(sh, currentRow, 'DATOS DE PAGO', '#455a64');
    var metodosPago_ = (cliInfo && Array.isArray(cliInfo.metodosPago)) ? cliInfo.metodosPago : [];
    if (metodosPago_.length === 0) {
      sh.getRange(currentRow, 2, 1, 6).merge()
        .setValue('Sin datos de pago registrados.')
        .setFontSize(10).setFontColor('#9aa0a6').setFontStyle('italic')
        .setVerticalAlignment('middle').setBackground('#ffffff');
      sh.setRowHeight(currentRow, 26);
      currentRow++;
    } else {
      for (var mi = 0; mi < metodosPago_.length; mi++) {
        var mPago = metodosPago_[mi];
        var pagoRows_ = _buildPagoRows_(mPago);
        var headerBg = mi === 0 ? '#ecf3fb' : '#f0f4c3';
        var pVals_ = pagoRows_.map(function(r) { return [r[0], r[1], '', '', '', '']; });
        var pBgs_  = pagoRows_.map(function(_, pi) {
          if (pi === 0) return [headerBg, headerBg, headerBg, headerBg, headerBg, headerBg];
          var bg = pi % 2 === 0 ? '#f8f9fa' : '#ffffff';
          return [bg, bg, bg, bg, bg, bg];
        });
        sh.getRange(currentRow, 2, pagoRows_.length, 6).setValues(pVals_).setFontSize(10).setBackgrounds(pBgs_);
        sh.getRange(currentRow, 2, pagoRows_.length, 1).setFontWeight('bold').setFontColor('#5f6368');
        sh.getRange(currentRow, 3, pagoRows_.length, 1).setFontColor('#3c4043');
        sh.getRange(currentRow, 2, 1, 1).setFontColor('#1a73e8');
        for (var pRi = 0; pRi < pagoRows_.length; pRi++) sh.setRowHeight(currentRow + pRi, 26);
        currentRow += pagoRows_.length;
        if (mi < metodosPago_.length - 1) currentRow++; // spacer between methods
      }
    }
    // Botón editar datos de pago
    sh.getRange(currentRow, 2, 1, 1).setValue('✏  Editar datos de pago')
      .setBackground('#e8f0fe').setFontColor('#1a73e8')
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.getRange(currentRow, 3).insertCheckboxes().setNote('EDITAR_PAGO').setBackground('#e8f0fe');
    sh.getRange(currentRow, 4, 1, 4).setBackground('#e8f0fe');
    sh.setRowHeight(currentRow, 28);
    currentRow++;
    currentRow++; // spacer

    // ── ALERTAS ───────────────────────────────────────────────────
    var alertasBlock = [];
    // Sin trabajos en 60+ días
    if (ultimoTrabajoAll) {
      var diasSinTra = Math.floor((hoy.getTime() - ultimoTrabajoAll.fecha.getTime()) / 86400000);
      if (diasSinTra >= 60) alertasBlock.push({
        msg: '⚠  Sin trabajos en los últimos ' + diasSinTra + ' días (último: ' + formatDate(ultimoTrabajoAll.fecha) + ')',
        color: '#e65100'
      });
    } else {
      alertasBlock.push({ msg: '⚠  Este cliente no tiene trabajos registrados', color: '#9e9e9e' });
    }
    // Proformas pendientes de cobro
    var prosPend = allProformas.filter(function(p) {
      var ec = (p.estadoCobro || '').toLowerCase();
      return ec !== 'pagada' && ec !== 'cobrado' && ec !== 'pagado' && ec !== 'anulada';
    });
    if (prosPend.length > 0) alertasBlock.push({
      msg: '⚠  ' + prosPend.length + ' proforma(s) pendiente(s) de cobro',
      color: '#7b1fa2'
    });
    // Saldo histórico supera umbral
    var umbralAlerta = parseFloat(getConfig('alerta_pendiente_max') || '500');
    if (histPendiente > umbralAlerta) alertasBlock.push({
      msg: '⚠  Saldo histórico pendiente: ' + formatCurrency(histPendiente) + '  (umbral configurado: ' + formatCurrency(umbralAlerta) + ')',
      color: '#b71c1c'
    });
    // Trabajos duplicados en varios documentos — integridad de datos comprometida
    for (var dupAi = 0; dupAi < trabajosDuplicados.length; dupAi++) {
      alertasBlock.push({
        msg: '⚠  Trabajo duplicado en documentos: ' + trabajosDuplicados[dupAi],
        color: '#c62828'
      });
    }

    if (alertasBlock.length > 0) {
      var aVals = alertasBlock.map(function(a) { return [a.msg,'','','','','']; });
      var aBgs  = alertasBlock.map(function(a) { return [a.color,a.color,a.color,a.color,a.color,a.color]; });
      sh.getRange(currentRow, 2, alertasBlock.length, 6)
        .setValues(aVals).setBackgrounds(aBgs)
        .setFontColor('#ffffff').setFontWeight('bold').setFontSize(10)
        .setVerticalAlignment('middle');
      for (var ai2 = 0; ai2 < alertasBlock.length; ai2++) sh.setRowHeight(currentRow + ai2, 28);
      currentRow += alertasBlock.length;
    }

    // ── TRABAJOS ─────────────────────────────────────────────────
    _pcTableHeader_(sh, currentRow++,
      ['Fecha', 'Operario', 'Trabajo realizado', 'Horas', 'Estado', 'Informe', 'Proforma', 'Factura', '🔍 Ver'], '#ef6c00');
    if (trabajos.length === 0) {
      _pcNoData_(sh, currentRow++);
    } else {
      var traRows = trabajos.map(function(t) {
        var inf    = informePorTrabajo[t.id]  || '';
        var pro    = proformaPorTrabajo[t.id] || '';
        var fac    = facturaPorTrabajo[t.id]  || '';
        var sinDoc = !inf && !pro && !fac;
        return [
          formatDate(t.fecha), t.operario || '—',
          t.trabajo.substring(0, 55) || '—',
          t.horas > 0 ? t.horas.toFixed(1).replace('.', ',') + ' h' : '0 h',
          t.estado || '—',
          inf || (sinDoc ? 'Sin documento' : '—'),
          pro || '—',
          fac || '—'
        ];
      });
      sh.getRange(currentRow, 2, traRows.length, 8).setValues(traRows).setFontSize(10);
      // Fondos por tipo de documento (prioridad: Factura > Proforma > Informe > Sin doc)
      var traBgs = trabajos.map(function(t) {
        var hasFac = !!facturaPorTrabajo[t.id];
        var hasPro = !!proformaPorTrabajo[t.id];
        var hasInf = !!informePorTrabajo[t.id];
        var bg = hasFac ? '#e8f5e9'   // verde suave
               : hasPro ? '#f3e5f5'   // morado suave
               : hasInf ? '#e0f2f1'   // teal suave
               :          '#fffde7';  // amarillo suave — sin documento
        return [bg, bg, bg, bg, bg, bg, bg, bg, bg];
      });
      sh.getRange(currentRow, 2, traRows.length, 9).setBackgrounds(traBgs);
      // Estado colors (col F=6) batch
      var estFcs = trabajos.map(function(t) { return [_pcEstadoColor_(t.estado) || '#3c4043']; });
      var estFws = trabajos.map(function(t) { return [_pcEstadoColor_(t.estado) ? 'bold' : 'normal']; });
      sh.getRange(currentRow, 6, traRows.length, 1).setFontColors(estFcs).setFontWeights(estFws);
      // Informe colors (col G=7) — teal; "Sin documento" en gris itálica
      sh.getRange(currentRow, 7, traRows.length, 1)
        .setFontColors(trabajos.map(function(t) {
          return [informePorTrabajo[t.id] ? '#00695c' : '#9aa0a6'];
        }))
        .setFontWeights(trabajos.map(function(t) {
          return [informePorTrabajo[t.id] ? 'bold' : 'normal'];
        }))
        .setFontStyles(trabajos.map(function(t) {
          var sinDoc = !informePorTrabajo[t.id] && !proformaPorTrabajo[t.id] && !facturaPorTrabajo[t.id];
          return [sinDoc ? 'italic' : 'normal'];
        }));
      // Proforma colors (col H=8) — morado
      sh.getRange(currentRow, 8, traRows.length, 1)
        .setFontColors(trabajos.map(function(t) { return [proformaPorTrabajo[t.id] ? '#7b1fa2' : '#9aa0a6']; }))
        .setFontWeights(trabajos.map(function(t) { return [proformaPorTrabajo[t.id] ? 'bold' : 'normal']; }));
      // Factura colors (col I=9) — verde
      sh.getRange(currentRow, 9, traRows.length, 1)
        .setFontColors(trabajos.map(function(t) { return [facturaPorTrabajo[t.id]  ? '#2e7d32' : '#9aa0a6']; }))
        .setFontWeights(trabajos.map(function(t) { return [facturaPorTrabajo[t.id]  ? 'bold' : 'normal']; }));
      // Notas/tooltip en columnas de documentos — lista de trabajos incluidos en cada doc
      var _docNote_ = function(tipo, docId, idsStr) {
        var lines = [tipo + '  ' + docId, ''];
        if (idsStr) {
          lines.push('Incluye:');
          idsStr.split(',').forEach(function(tid) { var t = tid.trim(); if (t) lines.push('  · ' + t); });
        }
        return lines.join('\n');
      };
      var facIdsMap = {}; allFacturas.forEach(function(f)  { if (f.id)  facIdsMap[f.id]  = f.idsTrabajos  || ''; });
      var proIdsMap = {}; allProformas.forEach(function(p) { if (p.id)  proIdsMap[p.id]  = p.idsTrabajos  || ''; });
      var infIdsMap = {}; allInformes.forEach(function(i)  { if (i.id)  infIdsMap[i.id]  = i.idsTrabajos  || ''; });
      sh.getRange(currentRow, 7, traRows.length, 1).setNotes(
        trabajos.map(function(t) { var d = informePorTrabajo[t.id];  return [d ? _docNote_('Informe',  d, infIdsMap[d]) : '']; })
      );
      sh.getRange(currentRow, 8, traRows.length, 1).setNotes(
        trabajos.map(function(t) { var d = proformaPorTrabajo[t.id]; return [d ? _docNote_('Proforma', d, proIdsMap[d]) : '']; })
      );
      sh.getRange(currentRow, 9, traRows.length, 1).setNotes(
        trabajos.map(function(t) { var d = facturaPorTrabajo[t.id];  return [d ? _docNote_('Factura',  d, facIdsMap[d]) : '']; })
      );
      // Checkboxes en col J=10 — ID del trabajo guardado en nota de celda
      sh.getRange(currentRow, 10, traRows.length, 1)
        .insertCheckboxes()
        .setBackground('#e8f0fe')
        .setFontColor('#1a73e8')
        .setHorizontalAlignment('center')
        .setNotes(trabajos.map(function(t) { return [t.id]; }));
      sh.setColumnWidth(10, 55);
      sh.showColumns(10);
      SpreadsheetApp.flush();
      currentRow += traRows.length;
    }
    currentRow++;

    // ── FACTURAS ─────────────────────────────────────────────────
    currentRow = _pcSectionHeader_(sh, currentRow, 'FACTURAS', '#1a73e8');
    if (facturas.length === 0) {
      _pcNoData_(sh, currentRow++);
    } else {
      _pcTableHeader_(sh, currentRow++, ['Número', 'Fecha', 'Total', 'Estado cobro', 'PDF', '✓ Pagada'], '#1565c0');
      for (var facIdx = 0; facIdx < facturas.length; facIdx++) {
        var f       = facturas[facIdx];
        var fbg     = facIdx % 2 === 0 ? '#ffffff' : '#f8f9fa';
        var fPagada = (f.estadoCobro === 'Pagada' || f.estadoCobro === 'Cobrado' || f.estadoCobro === 'Pagado');
        // Número clicable → abre PDF
        var fIdCell = sh.getRange(currentRow, 2).setFontSize(10).setBackground(fbg);
        if (f.pdfUrl) {
          fIdCell.setFormula('=HYPERLINK("' + f.pdfUrl.replace(/"/g, '') + '","' + f.id + '")');
          fIdCell.setFontColor('#1a73e8').setFontWeight('bold');
        } else {
          fIdCell.setValue(f.id).setFontColor('#3c4043');
        }
        sh.getRange(currentRow, 3, 1, 3).setValues([[
          formatDate(f.fecha), formatCurrency(f.total), f.estadoCobro || f.estado
        ]]).setFontSize(10).setBackground(fbg);
        _pcPdfLink_(sh, currentRow, 6, f.pdfUrl, fbg);
        // Checkbox ✓ Cobrada en col G (7) — ID del documento guardado en nota de celda
        sh.getRange(currentRow, 7)
          .insertCheckboxes().setValue(fPagada)
          .setBackground(fPagada ? '#e6f4ea' : fbg)
          .setFontColor(fPagada ? '#34a853' : '#9aa0a6')
          .setHorizontalAlignment('center')
          .setNote('FAC:' + f.id);
        currentRow++;
        var detFac = _pcItemsForDoc_(f, allTrabajosMap, presupuestosMap);
        currentRow += _pcWriteIncludes_(sh, currentRow, detFac, '#e8f0fe', '#f0f4ff');
      }
    }
    currentRow++;

    // ── PROFORMAS ────────────────────────────────────────────────
    currentRow = _pcSectionHeader_(sh, currentRow, 'PROFORMAS', '#7b1fa2');
    if (proformas.length === 0) {
      _pcNoData_(sh, currentRow++);
    } else {
      _pcTableHeader_(sh, currentRow++, ['Número', 'Fecha', 'Total', 'Estado cobro', 'PDF', '✓ Pagada'], '#6a1b9a');
      for (var proIdx = 0; proIdx < proformas.length; proIdx++) {
        var p       = proformas[proIdx];
        var pbg     = proIdx % 2 === 0 ? '#ffffff' : '#f8f9fa';
        var pPagada = (p.estadoCobro === 'Pagada' || p.estadoCobro === 'Cobrado' || p.estadoCobro === 'Pagado');
        // Número clicable → abre PDF
        var pIdCell = sh.getRange(currentRow, 2).setFontSize(10).setBackground(pbg);
        if (p.pdfUrl) {
          pIdCell.setFormula('=HYPERLINK("' + p.pdfUrl.replace(/"/g, '') + '","' + p.id + '")');
          pIdCell.setFontColor('#7b1fa2').setFontWeight('bold');
        } else {
          pIdCell.setValue(p.id).setFontColor('#3c4043');
        }
        sh.getRange(currentRow, 3, 1, 3).setValues([[
          formatDate(p.fecha), formatCurrency(p.total), p.estadoCobro
        ]]).setFontSize(10).setBackground(pbg);
        _pcPdfLink_(sh, currentRow, 6, p.pdfUrl, pbg);
        // Checkbox ✓ Cobrada en col G (7) — ID del documento guardado en nota de celda
        sh.getRange(currentRow, 7)
          .insertCheckboxes().setValue(pPagada)
          .setBackground(pPagada ? '#f3e5f5' : pbg)
          .setFontColor(pPagada ? '#7b1fa2' : '#9aa0a6')
          .setHorizontalAlignment('center')
          .setNote('PRO:' + p.id);
        currentRow++;
        var detPro = _pcItemsForDoc_(p, allTrabajosMap, presupuestosMap);
        currentRow += _pcWriteIncludes_(sh, currentRow, detPro, '#f3e5f5', '#faf0fc');
      }
    }
    currentRow++;

    // ── INFORMES ─────────────────────────────────────────────────
    currentRow = _pcSectionHeader_(sh, currentRow, 'INFORMES', '#00695c');
    if (informes.length === 0) {
      _pcNoData_(sh, currentRow++);
    } else {
      _pcTableHeader_(sh, currentRow++, ['Número', 'Fecha', 'PDF', '', '', ''], '#00897b');
      var infVals = informes.map(function(inf) { return [inf.id, formatDate(inf.fecha), '', '', '', '']; });
      var infBgs  = informes.map(function(inf, ii) { var bg = ii % 2 === 0 ? '#ffffff' : '#f8f9fa'; return [bg,bg,bg,bg,bg,bg]; });
      sh.getRange(currentRow, 2, informes.length, 6).setValues(infVals).setFontSize(10).setBackgrounds(infBgs);
      for (var ii = 0; ii < informes.length; ii++) {
        _pcPdfLink_(sh, currentRow + ii, 4, informes[ii].pdfUrl, ii % 2 === 0 ? '#ffffff' : '#f8f9fa');
      }
      currentRow += informes.length;
    }
    currentRow++;

    // ── COBROS PENDIENTES ────────────────────────────────────────
    currentRow = _pcSectionHeader_(sh, currentRow, 'COBROS PENDIENTES', '#b71c1c');
    if (pendFact.length === 0 && pendPro.length === 0) {
      sh.getRange(currentRow, 2, 1, 6).merge()
        .setValue('Sin importes pendientes de cobro en el periodo.')
        .setFontSize(10).setFontColor('#34a853').setFontWeight('bold')
        .setVerticalAlignment('middle').setBackground('#f1f8f1');
      currentRow++;
    } else {
      _pcTableHeader_(sh, currentRow++, ['Documento', 'Tipo', 'Importe', 'Vencido', 'Semáforo', ''], '#c62828');
      var cobroVals = [], cobroBgs = [], cobroFcs = [];
      var cobIdx = 0;
      var buildCobro = function(doc, tipo) {
        var diasVenc = Math.floor((hoy.getTime() - doc.fecha.getTime()) / 86400000);
        var diasStr  = diasVenc <= 0 ? 'Al día' : diasVenc + ' días';
        var semColor = diasVenc <= 0 ? '#34a853' : diasVenc <= 15 ? '#f9ab00' : '#ea4335';
        var semaforo = diasVenc <= 0 ? '🟢' : diasVenc <= 15 ? '🟡' : '🔴';
        var bg = cobIdx % 2 === 0 ? '#fff8f8' : '#ffffff';
        cobIdx++;
        cobroVals.push([doc.id, tipo, formatCurrency(doc.total), diasStr, semaforo, '']);
        cobroBgs.push([bg,bg,bg,bg,bg,bg]);
        cobroFcs.push(['#3c4043','#3c4043','#3c4043','#3c4043',semColor,'#3c4043']);
      };
      for (var cfi = 0; cfi < pendFact.length; cfi++) buildCobro(pendFact[cfi], 'Factura');
      for (var cpi = 0; cpi < pendPro.length;  cpi++) buildCobro(pendPro[cpi],  'Proforma');
      sh.getRange(currentRow, 2, cobroVals.length, 6)
        .setValues(cobroVals).setFontSize(10).setBackgrounds(cobroBgs).setFontColors(cobroFcs);
      currentRow += cobroVals.length;
      sh.getRange(currentRow, 2, 1, 2).merge()
        .setValue('PENDIENTE TOTAL').setFontSize(11).setFontWeight('bold')
        .setBackground('#fce4e4').setFontColor('#b71c1c').setVerticalAlignment('middle');
      sh.getRange(currentRow, 4).setValue(formatCurrency(totalPendiente))
        .setFontSize(11).setFontWeight('bold')
        .setBackground('#fce4e4').setFontColor('#b71c1c').setVerticalAlignment('middle');
      sh.setRowHeight(currentRow, 32);
      currentRow++;
    }
    currentRow++;

    // ── HISTÓRICO DEL CLIENTE ─────────────────────────────────────
    currentRow = _pcSectionHeader_(sh, currentRow, 'HISTÓRICO DEL CLIENTE', '#37474f');
    var histRows = [
      ['Total trabajos realizados', String(allTrabajos.length)],
      ['Total horas',              histHorasTotal.toFixed(1).replace('.', ',') + ' h'],
      ['Total facturado',          formatCurrency(histTotalFacturado)],
      ['Total cobrado',            formatCurrency(histCobrado)],
      ['Total pendiente de cobro', formatCurrency(histPendiente)],
      ['Primera factura',  primerFacAll ? formatDate(primerFacAll.fecha) + '  ·  ' + primerFacAll.id : '—'],
      ['Última factura',   ultimaFacAll ? formatDate(ultimaFacAll.fecha) + '  ·  ' + ultimaFacAll.id : '—']
    ];
    var hVals = histRows.map(function(r) { return [r[0], r[1], '', '', '', '']; });
    var hBgs  = histRows.map(function(r, hi2) { var bg = hi2 % 2 === 0 ? '#f5f5f5' : '#ffffff'; return [bg,bg,bg,bg,bg,bg]; });
    sh.getRange(currentRow, 2, histRows.length, 6).setValues(hVals).setFontSize(10).setBackgrounds(hBgs);
    sh.getRange(currentRow, 2, histRows.length, 1).setFontWeight('bold').setFontColor('#5f6368');
    sh.getRange(currentRow, 3, histRows.length, 1).setFontWeight('bold').setFontColor('#1a73e8');
    currentRow += histRows.length;
    currentRow++;

    // ── HISTORIAL CRONOLÓGICO ─────────────────────────────────────
    currentRow = _pcSectionHeader_(sh, currentRow, 'HISTORIAL CRONOLÓGICO', '#455a64');
    var eventos = [];
    for (var tei = 0; tei < allTrabajos.length; tei++) {
      var te = allTrabajos[tei];
      eventos.push({
        fecha: te.fecha, tipo: '🔧 Trabajo',
        desc:    (te.trabajo || '—').substring(0, 70),
        importe: te.horas > 0 ? te.horas.toFixed(1).replace('.', ',') + ' h' : '',
        estado:  te.estado || ''
      });
    }
    for (var fei = 0; fei < allFacturas.length; fei++) {
      var fe = allFacturas[fei];
      var feEstado = fe.estadoCobro || fe.estado;
      eventos.push({
        fecha: fe.fecha, tipo: '🧾 Factura', desc: fe.id,
        importe: formatCurrency(fe.total),
        estado:  feEstado === 'Cobrado' ? 'Pagado' : feEstado
      });
    }
    for (var pei = 0; pei < allProformas.length; pei++) {
      var pe = allProformas[pei];
      eventos.push({
        fecha: pe.fecha, tipo: '📋 Proforma', desc: pe.id,
        importe: formatCurrency(pe.total),
        estado:  pe.estadoCobro
      });
    }
    for (var iei = 0; iei < allInformes.length; iei++) {
      var ie = allInformes[iei];
      eventos.push({
        fecha: ie.fecha, tipo: '📄 Informe', desc: ie.id,
        importe: '—',
        estado:  'Emitido'
      });
    }
    eventos.sort(function(a, b) { return b.fecha - a.fecha; });
    if (eventos.length > 150) eventos = eventos.slice(0, 150);

    if (eventos.length === 0) {
      _pcNoData_(sh, currentRow++);
    } else {
      _pcTableHeader_(sh, currentRow++, ['Fecha', 'Tipo', 'Descripción', 'Valor', 'Estado', ''], '#546e7a');
      var evVals = [], evBgs = [], evFcs = [];
      var prevFechaStr = null;
      for (var evi = 0; evi < eventos.length; evi++) {
        var ev       = eventos[evi];
        var fechaStr = formatDate(ev.fecha);
        var showFecha = (fechaStr !== prevFechaStr);
        prevFechaStr = fechaStr;
        // Fondo: nuevo grupo de fecha = ligeramente destacado
        var evbg = showFecha ? '#eceff1' : (evi % 2 === 0 ? '#ffffff' : '#f5f5f5');
        var tColor = ev.tipo.indexOf('Factura')  >= 0 ? '#1565c0'
                   : ev.tipo.indexOf('Proforma') >= 0 ? '#7b1fa2'
                   : ev.tipo.indexOf('Informe')  >= 0 ? '#00695c'
                   : '#e65100'; // Trabajo
        var impColor = ev.tipo.indexOf('Trabajo') >= 0 ? '#5f6368' : '#1a73e8';
        evVals.push([showFecha ? fechaStr : '', ev.tipo, ev.desc, ev.importe, ev.estado, '']);
        evBgs.push([evbg,evbg,evbg,evbg,evbg,evbg]);
        evFcs.push(['#9aa0a6', tColor, '#3c4043', impColor, '#5f6368', '#3c4043']);
      }
      sh.getRange(currentRow, 2, eventos.length, 6)
        .setValues(evVals).setFontSize(9).setBackgrounds(evBgs).setFontColors(evFcs);
      currentRow += eventos.length;
    }

    // ── Layout final: garantizar col J visible y anchos de columnas doc ─
    sh.setColumnWidth(7,  100);  // G — Informe
    sh.setColumnWidth(8,  100);  // H — Proforma
    sh.setColumnWidth(9,  100);  // I — Factura
    sh.setColumnWidth(10,  55);  // J — 🔍 Ver
    sh.showColumns(10);

    logEvent('INFO', 'refreshPanelCliente',
      cliente + ' | ' + fechaIniDv + ' -> ' + fechaFinDv +
      ' | t=' + trabajos.length + ' f=' + facturas.length + ' p=' + proformas.length +
      ' | hist: t=' + allTrabajos.length + ' f=' + allFacturas.length + ' ev=' + eventos.length, '');

  } catch (ex) {
    _pcEstado_(sh, '⚠  Error: ' + ex.message, '#ffffff', '#ea4335');
    logEvent('ERROR', 'refreshPanelCliente', ex.message, ex.stack || '');
  }
}

// ================================================================
// EXPORTAR HISTORIAL PDF
// ================================================================
function exportarHistorialClientePDF() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('PANEL_CLIENTE');
  if (!sh) { SpreadsheetApp.getUi().alert('PANEL_CLIENTE no existe.'); return; }

  var cliente = String(sh.getRange(PC_.CLIENTE[0], PC_.CLIENTE[1]).getValue() || '').trim();
  if (!cliente) { SpreadsheetApp.getUi().alert('Selecciona primero un cliente en PANEL_CLIENTE.'); return; }

  _pcEstado_(sh, 'Generando PDF del historial de ' + cliente + '...', '#ffffff', '#f9ab00');
  SpreadsheetApp.flush();

  try {
    // ── Cargar datos completos (sin filtro de fechas) ─────────────
    var allTrabHist = [], allFacHist = [], allProHist = [];

    try {
      var shT = getSheet('TRABAJOS'); var lrT = shT.getLastRow();
      if (lrT > 1) {
        var dT = shT.getRange(2, 1, lrT - 1, Math.min(shT.getLastColumn(), COL_TRA.MATERIALES)).getValues();
        for (var i = 0; i < dT.length; i++) {
          if (String(dT[i][COL_TRA.CLIENTE-1]||'').trim().toLowerCase() !== cliente.toLowerCase()) continue;
          var ft2 = _parseTrabajosDate_(dT[i][COL_TRA.FECHA_TRABAJO-1]); if (!ft2) continue;
          var h2  = _calcHoras_(dT[i][COL_TRA.HORA_INICIO-1], dT[i][COL_TRA.HORA_FIN-1]);
          allTrabHist.push({
            fecha: ft2, operario: String(dT[i][COL_TRA.OPERARIO-1]||'').trim(),
            trabajo: String(dT[i][COL_TRA.TRABAJO_REALIZADO-1]||'').trim(),
            horas: h2, estado: String(dT[i][COL_TRA.ESTADO-1]||'').trim()
          });
        }
        allTrabHist.sort(function(a,b){return b.fecha-a.fecha;});
      }
    } catch(et){}

    var colECFac = -1;
    try {
      var shF = getSheet('FACTURAS'); var lrF = shF.getLastRow();
      if (lrF > 1) {
        var hdrsF = shF.getRange(1,1,1,shF.getLastColumn()).getValues()[0];
        for (var fh2=0;fh2<hdrsF.length;fh2++){if(String(hdrsF[fh2]).trim()==='Estado Cobro'){colECFac=fh2;break;}}
        var dF = shF.getRange(2,1,lrF-1,hdrsF.length).getValues();
        for (var j2=0;j2<dF.length;j2++){
          if (String(dF[j2][5]||'').trim().toLowerCase()!==cliente.toLowerCase()) continue;
          var fd2=dF[j2][1] instanceof Date?dF[j2][1]:_parseReportDate_(String(dF[j2][1]||'')); if(!fd2) continue;
          var ec2=colECFac>=0?String(dF[j2][colECFac]||'').trim():'';
          if(!ec2) ec2=String(dF[j2][17]||'').trim()==='Emitida'?'Pendiente':String(dF[j2][17]||'').trim();
          allFacHist.push({id:String(dF[j2][0]||''),fecha:fd2,total:Number(dF[j2][16]||0),estadoCobro:ec2});
        }
        allFacHist.sort(function(a,b){return b.fecha-a.fecha;});
      }
    } catch(ef){}

    try {
      var shP = getSheet('PROFORMAS'); var lrP = shP.getLastRow();
      if (lrP > 1) {
        var dP = shP.getRange(2,1,lrP-1,shP.getLastColumn()).getValues();
        for (var k2=0;k2<dP.length;k2++){
          if(String(dP[k2][5]||'').trim().toLowerCase()!==cliente.toLowerCase()) continue;
          var pd2=dP[k2][1] instanceof Date?dP[k2][1]:_parseReportDate_(String(dP[k2][1]||'')); if(!pd2) continue;
          allProHist.push({id:String(dP[k2][0]||''),fecha:pd2,total:Number(dP[k2][12]||0),estadoCobro:String(dP[k2][18]||'Pendiente de cobro')});
        }
        allProHist.sort(function(a,b){return b.fecha-a.fecha;});
      }
    } catch(ep){}

    // ── Estadísticas ─────────────────────────────────────────────
    var totalHorasHist = allTrabHist.reduce(function(s,t){return s+t.horas;},0);
    var totalFacHist   = allFacHist.reduce(function(s,f){return s+f.total;},0);
    var totalCobHist   = allFacHist.filter(function(f){return f.estadoCobro==='Cobrado'||f.estadoCobro==='Pagado';}).reduce(function(s,f){return s+f.total;},0);
    var totalPteHist   = Math.max(0, totalFacHist - totalCobHist);
    var pendientesFac  = allFacHist.filter(function(f){return f.estadoCobro!=='Cobrado'&&f.estadoCobro!=='Pagado';});

    // ── Crear Google Doc ─────────────────────────────────────────
    var empresa  = getConfig('empresa_nombre') || 'Mi Empresa';
    var hoy2     = new Date();
    var tz       = Session.getScriptTimeZone();
    var hoyStr2  = Utilities.formatDate(hoy2, tz, 'dd/MM/yyyy HH:mm');
    var docTitle = 'Historial_' + _sanitizeFilename_(cliente) + '_' + Utilities.formatDate(hoy2, tz, 'yyyyMMdd');
    var doc      = DocumentApp.create(docTitle);
    var body     = doc.getBody();
    body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);

    body.appendParagraph(empresa).setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .editAsText().setFontSize(16).setBold(true);
    body.appendParagraph('Historial de cliente · Generado el ' + hoyStr2)
      .editAsText().setFontSize(9).setForegroundColor('#888888');
    body.appendHorizontalRule();

    var cliI2 = null; try { cliI2 = _getClienteByLabel_(cliente); } catch(ecc){}
    body.appendParagraph('CLIENTE: ' + (cliI2 ? cliI2.nombre : cliente))
      .setHeading(DocumentApp.ParagraphHeading.HEADING2)
      .editAsText().setFontSize(14).setBold(true);
    if (cliI2) {
      if (cliI2.cif)      body.appendParagraph('CIF/NIF: ' + cliI2.cif).editAsText().setFontSize(10);
      if (cliI2.telefono) body.appendParagraph('Tel: ' + cliI2.telefono).editAsText().setFontSize(10);
      if (cliI2.email)    body.appendParagraph('Email: ' + cliI2.email).editAsText().setFontSize(10);
      if (cliI2.direccion) body.appendParagraph(cliI2.direccion).editAsText().setFontSize(10);
    }
    body.appendParagraph('');

    // Resumen
    body.appendParagraph('RESUMEN HISTÓRICO').editAsText().setFontSize(12).setBold(true);
    var resData = [
      ['Total trabajos',  String(allTrabHist.length)],
      ['Total horas',     totalHorasHist.toFixed(1).replace('.',',') + ' h'],
      ['Total facturado', formatCurrency(totalFacHist)],
      ['Total cobrado',   formatCurrency(totalCobHist)],
      ['Pendiente cobro', formatCurrency(totalPteHist)]
    ];
    var tblRes = body.appendTable(resData);
    for (var ri=0;ri<tblRes.getNumRows();ri++) tblRes.getRow(ri).editAsText().setFontSize(10);
    body.appendParagraph('');

    // Cobros pendientes
    if (pendientesFac.length > 0) {
      body.appendParagraph('COBROS PENDIENTES (' + pendientesFac.length + ')')
        .editAsText().setFontSize(12).setBold(true).setForegroundColor('#b71c1c');
      var pteData = [['Factura','Total','Estado']];
      for (var pi2=0;pi2<pendientesFac.length;pi2++) {
        var pf=pendientesFac[pi2];
        pteData.push([pf.id, formatCurrency(pf.total), pf.estadoCobro]);
      }
      var tblPte = body.appendTable(pteData);
      tblPte.getRow(0).editAsText().setBold(true).setFontSize(9);
      for (var pr=1;pr<tblPte.getNumRows();pr++) tblPte.getRow(pr).editAsText().setFontSize(9);
      body.appendParagraph('');
    }

    // Trabajos
    if (allTrabHist.length > 0) {
      body.appendHorizontalRule();
      body.appendParagraph('TRABAJOS (' + allTrabHist.length + ')').editAsText().setFontSize(12).setBold(true);
      var traData = [['Fecha','Operario','Trabajo realizado','Horas','Estado']];
      for (var ti2=0;ti2<allTrabHist.length;ti2++) {
        var t2=allTrabHist[ti2];
        traData.push([formatDate(t2.fecha), t2.operario||'—', t2.trabajo.substring(0,80)||'—',
          t2.horas>0?t2.horas.toFixed(1).replace('.',',')+'h':'0h', t2.estado||'—']);
      }
      var tblTra = body.appendTable(traData);
      tblTra.getRow(0).editAsText().setBold(true).setFontSize(9);
      for (var tr=1;tr<tblTra.getNumRows();tr++) tblTra.getRow(tr).editAsText().setFontSize(9);
      body.appendParagraph('');
    }

    // Facturas
    if (allFacHist.length > 0) {
      body.appendHorizontalRule();
      body.appendParagraph('FACTURAS (' + allFacHist.length + ')').editAsText().setFontSize(12).setBold(true);
      var facData = [['Número','Fecha','Total','Estado cobro']];
      for (var fi4=0;fi4<allFacHist.length;fi4++) {
        var f2=allFacHist[fi4];
        facData.push([f2.id, formatDate(f2.fecha), formatCurrency(f2.total), f2.estadoCobro]);
      }
      var tblFac = body.appendTable(facData);
      tblFac.getRow(0).editAsText().setBold(true).setFontSize(9);
      for (var fr=1;fr<tblFac.getNumRows();fr++) tblFac.getRow(fr).editAsText().setFontSize(9);
      body.appendParagraph('');
    }

    // Proformas
    if (allProHist.length > 0) {
      body.appendHorizontalRule();
      body.appendParagraph('PROFORMAS (' + allProHist.length + ')').editAsText().setFontSize(12).setBold(true);
      var proData = [['Número','Fecha','Total','Estado']];
      for (var pi4=0;pi4<allProHist.length;pi4++) {
        var p2=allProHist[pi4];
        proData.push([p2.id, formatDate(p2.fecha), formatCurrency(p2.total), p2.estadoCobro]);
      }
      var tblPro = body.appendTable(proData);
      tblPro.getRow(0).editAsText().setBold(true).setFontSize(9);
      for (var prr=1;prr<tblPro.getNumRows();prr++) tblPro.getRow(prr).editAsText().setFontSize(9);
      body.appendParagraph('');
    }

    body.appendHorizontalRule();
    body.appendParagraph('Generado el ' + hoyStr2 + '  ·  ' + empresa)
      .editAsText().setFontSize(8).setForegroundColor('#888888');

    doc.saveAndClose();

    // ── Convertir a PDF y guardar ─────────────────────────────────
    var docFile = DriveApp.getFileById(doc.getId());
    var pdfBlob = docFile.getAs('application/pdf');
    pdfBlob.setName(docTitle + '.pdf');
    var carpetaId = getConfig('carpeta_facturas_id');
    var carpeta;
    if (carpetaId) { try { carpeta = DriveApp.getFolderById(carpetaId); } catch(e2) { carpeta = DriveApp.getRootFolder(); } }
    else { carpeta = DriveApp.getRootFolder(); }
    var pdfFile = carpeta.createFile(pdfBlob);
    docFile.setTrashed(true);

    _pcEstado_(sh, '✓ PDF generado: ' + pdfFile.getName(), '#1a5e20', '#e8f5e9');
    logEvent('INFO', 'exportarHistorialClientePDF',
      cliente + ' | t=' + allTrabHist.length + ' f=' + allFacHist.length, pdfFile.getUrl());
    SpreadsheetApp.getUi().alert(
      '✓ PDF generado correctamente.\n\n' +
      'Cliente: ' + cliente + '\n' +
      'Trabajos: ' + allTrabHist.length + '  ·  Facturas: ' + allFacHist.length + '\n\n' +
      'URL:\n' + pdfFile.getUrl()
    );
  } catch(ex2) {
    _pcEstado_(sh, '⚠  Error exportando: ' + ex2.message, '#ffffff', '#ea4335');
    logEvent('ERROR', 'exportarHistorialClientePDF', ex2.message, '');
    SpreadsheetApp.getUi().alert('Error generando PDF:\n\n' + ex2.message);
  }
}

// ================================================================
// TOGGLE ESTADO COBRO — llamado desde onPanelEdit (col G checkbox)
// ================================================================

/**
 * Actualiza el estado de cobro de una factura o proforma en su hoja.
 * Al finalizar refresca PANEL_CLIENTE y muestra confirmación en la barra de estado.
 *
 * @param {string}  docId    — ID del documento (e.g. "FAC-2026-001")
 * @param {string}  tipo     — 'Factura' o 'Proforma'
 * @param {boolean} isPagada — true → Pagada; false → Pendiente
 */
function actualizarEstadoCobro(docId, tipo, isPagada, filaEditada) {
  // Usar valores que respetan la validación de datos de cada hoja:
  //   PROFORMAS → Estado_Cobro: ['Pendiente de cobro', 'Pagada', 'Anulada']
  //   FACTURAS  → Estado Cobro:  ['Pendiente', 'Cobrado', 'Parcial', 'Cancelado']
  var nuevoEstado = tipo === 'Proforma'
    ? (isPagada ? 'Pagada' : 'Pendiente de cobro')
    : (isPagada ? 'Cobrado' : 'Pendiente');
  Logger.log('[actualizarEstadoCobro] ── INICIO ──────────────────────────────');
  Logger.log('[actualizarEstadoCobro] Fila editada (PANEL_CLIENTE): ' + (filaEditada || '?'));
  Logger.log('[actualizarEstadoCobro] Tipo documento : ' + tipo);
  Logger.log('[actualizarEstadoCobro] Número documento: ' + docId);
  Logger.log('[actualizarEstadoCobro] Nuevo valor checkbox (isPagada): ' + isPagada
    + ' → estado a escribir: "' + nuevoEstado + '"');

  var shName = tipo === 'Proforma' ? 'PROFORMAS' : 'FACTURAS';
  Logger.log('[actualizarEstadoCobro] Hoja destino: ' + shName);
  var sh;
  try { sh = getSheet(shName); } catch (e) {
    Logger.log('[actualizarEstadoCobro] ERROR: Hoja no encontrada: ' + shName);
    return;
  }

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  // Localizar fila del documento
  var ids       = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  var targetRow = -1;
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === docId) { targetRow = i + 2; break; }
  }
  Logger.log('[actualizarEstadoCobro] Fila destino encontrada: '
    + (targetRow > 0 ? targetRow : 'NO ENCONTRADA — docId=' + docId));
  if (targetRow < 0) {
    Logger.log('[actualizarEstadoCobro] ERROR: Documento no encontrado en ' + shName);
    return;
  }

  // Localizar columnas por cabecera (tolerante a mayúsculas y guiones bajos)
  var hdrs     = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var colCobro = -1, colPago = -1, colMetodoPago = -1, colDatosPago = -1;
  for (var h = 0; h < hdrs.length; h++) {
    var hdrNorm = String(hdrs[h]).trim().toLowerCase().replace(/_/g, ' ');
    if (hdrNorm === 'estado cobro')       colCobro      = h + 1;
    if (hdrNorm === 'fecha pago')         colPago       = h + 1;
    if (hdrNorm === 'metodo pago usado')  colMetodoPago = h + 1;
    if (hdrNorm === 'datos pago usados')  colDatosPago  = h + 1;
  }

  if (tipo === 'Proforma') {
    if (colCobro < 0) colCobro = 19; // posición estándar PROFORMAS (col Estado_Cobro)
    if (colPago  < 0) colPago  = 17; // posición estándar PROFORMAS (col Fecha_Pago)
    sh.getRange(targetRow, colCobro).setValue(nuevoEstado);
    sh.getRange(targetRow, colPago ).setValue(isPagada ? new Date() : '');
  } else {
    // FACTURAS: crear columna "Estado Cobro" si aún no existe
    if (colCobro < 0) {
      colCobro = sh.getLastColumn() + 1;
      sh.getRange(1, colCobro)
        .setValue('Estado Cobro')
        .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    }
    if (colPago < 0) colPago = 21; // posición estándar FACTURAS (col Fecha_Pago)
    sh.getRange(targetRow, colCobro).setValue(nuevoEstado);
    sh.getRange(targetRow, colPago ).setValue(isPagada ? new Date() : '');
  }

  // Registrar o borrar el método de pago usado
  if (colMetodoPago > 0 || colDatosPago > 0) {
    if (!isPagada) {
      if (colMetodoPago > 0) sh.getRange(targetRow, colMetodoPago).setValue('');
      if (colDatosPago  > 0) sh.getRange(targetRow, colDatosPago ).setValue('');
    } else {
      // Obtener cliente y sus métodos de pago
      var cliNombreDoc_ = '';
      try {
        var rowData_ = sh.getRange(targetRow, 1, 1, sh.getLastColumn()).getValues()[0];
        // Nombre_Cliente está en col 6 (índice 5) para FACTURAS y PROFORMAS
        cliNombreDoc_ = String(rowData_[5] || '').trim();
      } catch (_) {}
      var cliDocData_  = cliNombreDoc_ ? _getClienteByLabel_(cliNombreDoc_) : null;
      var metodosDoc_  = cliDocData_ ? (cliDocData_.metodosPago || []) : [];
      var pagoUsado_   = null;

      if (metodosDoc_.length === 1) {
        pagoUsado_ = metodosDoc_[0];
      } else if (metodosDoc_.length > 1) {
        var opsTxt_ = metodosDoc_.map(function(m, i) {
          return (i + 1) + '. ' + m.tipo + (m.iban ? ' (' + m.iban + ')' : m.numero ? ' (' + m.numero + ')' : '');
        }).join('\n');
        var selResp_ = SpreadsheetApp.getUi().prompt(
          'Método de pago utilizado',
          'Selecciona el número del método de pago:\n\n' + opsTxt_,
          SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
        );
        if (selResp_.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
          var selIdx_ = parseInt(selResp_.getResponseText()) - 1;
          if (selIdx_ >= 0 && selIdx_ < metodosDoc_.length) pagoUsado_ = metodosDoc_[selIdx_];
        }
      }

      if (colMetodoPago > 0) sh.getRange(targetRow, colMetodoPago).setValue(pagoUsado_ ? pagoUsado_.tipo : '');
      if (colDatosPago  > 0) sh.getRange(targetRow, colDatosPago ).setValue(pagoUsado_ ? JSON.stringify(pagoUsado_) : '');
    }
  }

  Logger.log('[actualizarEstadoCobro] Estado escrito: "' + nuevoEstado
    + '" → ' + shName + ' fila ' + targetRow + ' col ' + colCobro);
  logEvent('INFO', 'actualizarEstadoCobro',
    tipo + ' ' + docId + ' → ' + nuevoEstado, '');

  // flush() garantiza que la escritura sea visible antes del refresh
  SpreadsheetApp.flush();
  Logger.log('[actualizarEstadoCobro] flush() completado — iniciando refreshPanelCliente');

  refreshPanelCliente();

  var pcSh = SpreadsheetApp.getActive().getSheetByName('PANEL_CLIENTE');
  if (pcSh) {
    _pcEstado_(pcSh,
      '✅  ' + docId + '  →  ' + nuevoEstado,
      '#ffffff', '#34a853');
  }
  Logger.log('[actualizarEstadoCobro] ── FIN ──────────────────────────────────');
}

// ================================================================
// HELPERS INTERNOS
// ================================================================

function _pcItemsForDoc_(doc, allTrabajosMap, presupuestosMap) {
  // Prioridad 1: IDs_Trabajos en la fila del documento (FACTURAS/PROFORMAS col 10)
  if (doc.idsTrabajos) {
    var ids = doc.idsTrabajos.split(','), found2 = [], horas2 = 0;
    for (var i = 0; i < ids.length; i++) {
      var id = ids[i].trim();
      if (id && allTrabajosMap[id]) {
        var t = allTrabajosMap[id];
        found2.push(formatDate(t.fecha) + '  —  ' + (t.trabajo.substring(0, 60) || '—'));
        horas2 += t.horas || 0;
      }
    }
    if (found2.length > 0) return { items: found2, type: 'trabajos', horas: horas2 };
  }
  // Prioridad 3: Lineas del presupuesto origen (solo docs generados desde presupuesto)
  if (doc.idPresup && presupuestosMap[doc.idPresup]) {
    var lineas = presupuestosMap[doc.idPresup], descs = [];
    for (var j = 0; j < lineas.length; j++) {
      var l = lineas[j];
      var txt = (String(l.desc || '').substring(0, 60) || '—');
      if (l.cant > 0) txt += '  ×' + l.cant;
      descs.push(txt);
    }
    if (descs.length > 0) return { items: descs, type: 'presup', horas: 0 };
  }
  return null;
}

function _pcWriteIncludes_(sh, startRow, detail, hdrBg, itemBg) {
  if (!detail || !detail.items || !detail.items.length) return 0;
  var items    = detail.items;
  var label    = detail.type === 'trabajos'
    ? '   📋 Trabajos incluidos (' + items.length + '):'
    : '   📋 Servicios del presupuesto (' + items.length + '):';
  var showHoras = detail.horas > 0;
  var nRows  = items.length + 1 + (showHoras ? 1 : 0);
  var vals   = [[label,'','','','','']];
  var bgs    = [[hdrBg,hdrBg,hdrBg,hdrBg,hdrBg,hdrBg]];
  for (var i = 0; i < items.length; i++) {
    vals.push(['        ' + items[i],'','','','','']);
    bgs.push([itemBg,itemBg,itemBg,itemBg,itemBg,itemBg]);
  }
  if (showHoras) {
    vals.push(['   ⏱ Total horas incluidas: ' + detail.horas.toFixed(1).replace('.', ',') + ' h','','','','','']);
    bgs.push([hdrBg,hdrBg,hdrBg,hdrBg,hdrBg,hdrBg]);
  }
  sh.getRange(startRow, 2, nRows, 6)
    .setValues(vals).setFontSize(9).setFontColor('#5f6368').setFontStyle('italic').setBackgrounds(bgs);
  return nRows;
}

function _pcEstado_(sh, msg, fontColor, bgColor) {
  sh.getRange(PC_.ESTADO[0], PC_.ESTADO[1], 1, 5).merge().setValue(msg)
    .setFontColor(fontColor || '#3c4043').setBackground(bgColor || '#f8f9fa').setFontSize(10)
    .setFontWeight(bgColor && bgColor !== '#f8f9fa' && bgColor !== '#e8f5e9' ? 'bold' : 'normal')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
}

function _pcResVal_(sh, coords, val) {
  sh.getRange(coords[0], 3, 1, 4).merge().setValue(val)
    .setFontSize(10).setVerticalAlignment('middle').setBackground('#ffffff');
}

function _pcEstadoColor_(estado) {
  switch ((estado || '').toLowerCase()) {
    case 'completado': return '#34a853';
    case 'pendiente':  return '#f9ab00';
    case 'en curso':   return '#1a73e8';
    case 'cancelado':  return '#ea4335';
    default:           return null;
  }
}

function _pcSectionHeader_(sh, row, title, color, numCols) {
  sh.getRange(row, 2, 1, numCols || 6).merge().setValue(title)
    .setFontSize(11).setFontWeight('bold').setFontColor('#ffffff')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setBackground(color || '#5f6368');
  sh.setRowHeight(row, 32);
  return row + 1;
}

function _pcTableHeader_(sh, row, headers, color) {
  sh.getRange(row, 2, 1, headers.length).setValues([headers])
    .setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground(color || '#757575').setVerticalAlignment('middle');
  sh.setRowHeight(row, 24);
}

function _pcNoData_(sh, row) {
  sh.getRange(row, 2, 1, 6).merge()
    .setValue('Sin registros para el periodo seleccionado.')
    .setFontSize(10).setFontColor('#9aa0a6').setFontStyle('italic')
    .setVerticalAlignment('middle').setBackground('#fafafa');
  sh.setRowHeight(row, 26);
}

function _pcStripe_(sh, startRow, numRows, numCols) {
  var bgs = [];
  for (var i = 0; i < numRows; i++) {
    var bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa', row = [];
    for (var c = 0; c < numCols; c++) row.push(bg);
    bgs.push(row);
  }
  sh.getRange(startRow, 2, numRows, numCols).setBackgrounds(bgs);
}

function _pcPdfLink_(sh, row, col, url, bg) {
  var cell = sh.getRange(row, col);
  if (url) {
    cell.setFormula('=HYPERLINK("' + url.replace(/"/g, '') + '","📄 Abrir PDF")');
    cell.setFontColor('#1a73e8').setFontSize(10);
  } else {
    cell.setValue('—').setFontColor('#9aa0a6').setFontSize(10);
  }
  if (bg) cell.setBackground(bg);
}

// ================================================================
// POPUP: DETALLE DEL TRABAJO
// ================================================================

/**
 * Busca el trabajo por ID en la hoja TRABAJOS y abre un modal con sus datos.
 * Llamado desde onPanelEdit cuando el usuario activa el checkbox de la col H.
 * @param {string} trabajoId
 * @param {string} docAsocId  ID de factura/proforma en col G (puede estar vacío)
 */
function mostrarDetalleTrabajoPC(trabajoId, docAsocId) {
  try {
    var shTra  = getSheet('TRABAJOS');
    var lastR  = shTra.getLastRow();
    if (lastR < 2) { SpreadsheetApp.getUi().alert('No hay trabajos registrados.'); return; }

    var data = shTra.getRange(2, 1, lastR - 1, 19).getValues();
    var row  = null;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][COL_TRA.ID - 1] || '').trim() === trabajoId) { row = data[i]; break; }
    }
    if (!row) { SpreadsheetApp.getUi().alert('Trabajo no encontrado: ' + trabajoId); return; }

    var html = _pcDetalleHtml_(row, trabajoId, docAsocId || '');
    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput(html).setWidth(490).setHeight(580),
      'DETALLE DEL TRABAJO'
    );
  } catch(e) {
    SpreadsheetApp.getUi().alert('Error al abrir detalle:\n\n' + e.message);
  }
}

/**
 * Genera el HTML del popup con todos los datos del trabajo.
 */
function _pcDetalleHtml_(row, trabajoId, docAsocId) {
  var fecha      = _parseTrabajosDate_(row[COL_TRA.FECHA_TRABAJO - 1]);
  var cliente    = String(row[COL_TRA.CLIENTE            - 1] || '').trim();
  var operario   = String(row[COL_TRA.OPERARIO           - 1] || '').trim();
  var horaIni    = row[COL_TRA.HORA_INICIO - 1];
  var horaFin    = row[COL_TRA.HORA_FIN   - 1];
  var horas      = _calcHoras_(horaIni, horaFin);
  var trabajo    = String(row[COL_TRA.TRABAJO_REALIZADO  - 1] || '').trim();
  var incids     = String(row[COL_TRA.INCIDENCIAS        - 1] || '').trim();
  var estado     = String(row[COL_TRA.ESTADO             - 1] || '').trim();
  var materiales = Number(row[COL_TRA.MATERIALES         - 1] || 0);
  var carpeta    = String(row[COL_TRA.CARPETA_FOTOS      - 1] || '').trim();

  var fotos = [];
  for (var fi = 13; fi <= 17; fi++) {
    var url = String(row[fi - 1] || '').trim();
    if (url) fotos.push({ label: 'Foto ' + (fi - 12), url: url });
  }

  var estadoColor = '#757575';
  switch ((estado || '').toLowerCase()) {
    case 'completado': estadoColor = '#1e8e3e'; break;
    case 'pendiente':  estadoColor = '#e37400'; break;
    case 'en curso':   estadoColor = '#1558d6'; break;
    case 'cancelado':  estadoColor = '#c5221f'; break;
  }

  var docLabel = '';
  if (docAsocId && docAsocId !== 'Sin facturar') {
    var prefix = docAsocId.toUpperCase().indexOf('FAC') >= 0 ? 'Factura'
               : docAsocId.toUpperCase().indexOf('PRO') >= 0 ? 'Proforma' : 'Documento';
    docLabel = prefix + ': ' + docAsocId;
  }

  var H = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>';
  H += 'body{font-family:Google Sans,Arial,sans-serif;font-size:13px;color:#3c4043;margin:0;padding:0;background:#fff;}';
  H += '.hdr{background:#1a73e8;color:#fff;padding:14px 20px;font-size:15px;font-weight:bold;letter-spacing:.2px;}';
  H += '.body{padding:12px 20px 8px;max-height:470px;overflow-y:auto;}';
  H += '.row{display:flex;align-items:flex-start;padding:7px 0;border-bottom:1px solid #f1f3f4;}';
  H += '.row:last-child{border-bottom:none;}';
  H += '.lbl{color:#5f6368;font-size:11px;font-weight:bold;width:120px;min-width:120px;padding-top:2px;text-transform:uppercase;letter-spacing:.4px;}';
  H += '.val{flex:1;color:#3c4043;word-break:break-word;line-height:1.5;}';
  H += '.badge{display:inline-block;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:bold;color:#fff;}';
  H += '.warn{background:#fff8e1;border-radius:4px;padding:8px 12px;color:#e65100;font-size:12px;margin:6px 0;line-height:1.5;}';
  H += '.sec{font-size:11px;font-weight:bold;color:#5f6368;text-transform:uppercase;letter-spacing:.4px;margin:12px 0 6px;padding-bottom:4px;border-bottom:2px solid #e8eaed;}';
  H += '.flink{display:block;margin:4px 0;color:#1a73e8;text-decoration:none;font-size:12px;padding:3px 0;}';
  H += '.flink:hover{text-decoration:underline;}';
  H += '.doc-badge{background:#e8f0fe;color:#1a73e8;border-radius:4px;padding:3px 10px;font-size:12px;font-weight:bold;display:inline-block;}';
  H += '.ftr{padding:12px 20px;text-align:right;border-top:1px solid #e8eaed;background:#f8f9fa;}';
  H += '.btn{background:#1a73e8;color:#fff;border:none;padding:8px 22px;border-radius:4px;cursor:pointer;font-size:13px;font-weight:500;}';
  H += '.btn:hover{background:#1558b0;}';
  H += '</style></head><body>';

  H += '<div class="hdr">🔍 DETALLE DEL TRABAJO</div>';
  H += '<div class="body">';

  // ID (subtle)
  H += '<div class="row"><div class="lbl">ID</div>';
  H += '<div class="val" style="color:#9aa0a6;font-size:11px;">' + _escHtml_(trabajoId) + '</div></div>';

  // Fecha
  H += '<div class="row"><div class="lbl">Fecha</div>';
  H += '<div class="val"><b>' + (fecha ? formatDate(fecha) : '—') + '</b></div></div>';

  // Cliente
  H += '<div class="row"><div class="lbl">Cliente</div>';
  H += '<div class="val"><b>' + _escHtml_(cliente || '—') + '</b></div></div>';

  // Operario
  if (operario) {
    H += '<div class="row"><div class="lbl">Operario</div>';
    H += '<div class="val">' + _escHtml_(operario) + '</div></div>';
  }

  // Horas
  H += '<div class="row"><div class="lbl">Horas</div><div class="val">';
  H += horas > 0 ? '<b>' + horas.toFixed(1).replace('.', ',') + ' h</b>' : '—';
  if (horaIni && horaFin) {
    H += '&nbsp;&nbsp;<span style="color:#9aa0a6;font-size:11px;">' +
         _horaStr_(horaIni) + ' → ' + _horaStr_(horaFin) + '</span>';
  }
  H += '</div></div>';

  // Estado
  H += '<div class="row"><div class="lbl">Estado</div>';
  H += '<div class="val"><span class="badge" style="background:' + estadoColor + ';">' +
       _escHtml_(estado || '—') + '</span></div></div>';

  // Materiales
  if (materiales > 0) {
    H += '<div class="row"><div class="lbl">Materiales</div>';
    H += '<div class="val">' + formatCurrency(materiales) + '</div></div>';
  }

  // Documento asociado
  if (docLabel) {
    H += '<div class="row"><div class="lbl">Facturado</div>';
    H += '<div class="val"><span class="doc-badge">' + _escHtml_(docLabel) + '</span></div></div>';
  }

  // Descripción
  if (trabajo) {
    H += '<div class="sec">Descripción</div>';
    H += '<div style="padding:4px 0 8px;line-height:1.6;white-space:pre-wrap;">' + _escHtml_(trabajo) + '</div>';
  }

  // Incidencias
  if (incids) {
    H += '<div class="sec">Incidencias</div>';
    H += '<div class="warn">⚠ ' + _escHtml_(incids) + '</div>';
  }

  // Fotos
  if (fotos.length > 0 || carpeta) {
    H += '<div class="sec">📸 Fotos</div>';
    for (var fj = 0; fj < fotos.length; fj++) {
      H += '<a class="flink" href="' + _escAttr_(fotos[fj].url) + '" target="_blank">📷 ' + _escHtml_(fotos[fj].label) + '</a>';
    }
    if (carpeta) {
      H += '<a class="flink" href="' + _escAttr_(carpeta) + '" target="_blank">📁 Ver carpeta de fotos</a>';
    }
  }

  H += '</div>'; // body
  H += '<div class="ftr"><button class="btn" onclick="google.script.host.close()">✕ Cerrar</button></div>';
  H += '</body></html>';
  return H;
}

/** Escapa texto para incluir en HTML */
function _escHtml_(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

/** Escapa texto para atributo href */
function _escAttr_(s) {
  return String(s || '').replace(/"/g,'%22').replace(/'/g,'%27');
}

/** Formatea un valor de hora (Date, número fraccionario o string) como HH:MM */
function _horaStr_(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
  if (typeof val === 'number') {
    var totalMin = Math.round(val * 24 * 60);
    var h = Math.floor(totalMin / 60);
    var mn = totalMin % 60;
    return (h < 10 ? '0' : '') + h + ':' + (mn < 10 ? '0' : '') + mn;
  }
  return String(val);
}

// ================================================================
// DIAGNÓSTICO
// ================================================================
function testPanelClienteDiag() {
  var lines = [];
  var ss = SpreadsheetApp.getActive();
  if (!ss) { SpreadsheetApp.getUi().alert('CRITICO: SpreadsheetApp.getActive() es null.'); return; }
  lines.push('Spreadsheet: ' + ss.getName());
  var pcSheet = ss.getSheetByName('PANEL_CLIENTE');
  lines.push('PANEL_CLIENTE: ' + (pcSheet ? 'EXISTE (pos ' + pcSheet.getIndex() + ')' : 'NO EXISTE'));
  lines.push('COL_TRA.OPERARIO: ' + (typeof COL_TRA !== 'undefined' ? COL_TRA.OPERARIO : 'N/D'));
  lines.push('COL_TRA.ESTADO:   ' + (typeof COL_TRA !== 'undefined' ? COL_TRA.ESTADO   : 'N/D'));
  lines.push('COL_PRE.LINEAS_JSON: ' + (typeof COL_PRE !== 'undefined' ? COL_PRE.LINEAS_JSON : 'N/D'));
  lines.push('DATA_START: ' + PC_.DATA_START);
  SpreadsheetApp.getUi().alert(lines.join('\n'));
}
