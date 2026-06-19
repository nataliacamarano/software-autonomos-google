// ============================================================
// PanelHorasEmpleados.gs — Panel de horas trabajadas por operario
// ============================================================

// ── Layout: 8 columnas con anchos legibles ───────────────────
// A(1)=120 · B(2)=190 · C(3)=200 · D(4)=90 · E(5)=120 · F(6)=90 · G(7)=90 · H(8)=80

var PHE_ = {
  HOJA: 'PANEL_HORAS_EMPLEADOS',
  NCOLS: 8,

  // Filtros — fila 4
  OPERARIO:  [4, 2],   // B4
  FECHA_INI: [4, 4],   // D4
  FECHA_FIN: [4, 6],   // F4

  // Botón y estado — filas 6-7
  BTN_ACTUALIZAR: [6, 1],   // A6 — checkbox
  ESTADO:         [7, 1],   // A7 — mensaje

  // KPIs — filas 10(label) y 11(valor), 2 cols cada uno
  KPI_ROW_LBL:      10,
  KPI_ROW_VAL:      11,
  KPI_COL_HORAS:     1,   // A:B merged
  KPI_COL_TRABAJOS:  3,   // C:D merged
  KPI_COL_CLIENTES:  5,   // E:F merged
  KPI_COL_PROMEDIO:  7,   // G:H merged

  // Tabla detalle — cabecera fila 14, datos desde 15
  TABLE_HDR_ROW:   14,
  TABLE_START_ROW: 15,
  TABLE_ROWS:      60,
  TC_FECHA:    1,   // A
  TC_CLIENTE:  2,   // B
  TC_TRABAJO:  3,   // C
  TC_INICIO:   4,   // D
  TC_FIN:      5,   // E
  TC_HORAS:    6,   // F

  // Resumen clientes y mensual (calculados dinámicamente en setup)
  CLI_TITLE_ROW: 77,
  CLI_HDR_ROW:   78,
  CLI_START_ROW: 79,
  CLI_ROWS:      25,
  MES_TITLE_ROW: 106,
  MES_HDR_ROW:   107,
  MES_START_ROW: 108,
  MES_ROWS:      13,
  EXP_ROW:       123
};

// ── Índices columnas TRABAJOS (0-based) ──────────────────────
// Col1=ID, Col2=FECHA_ALTA, Col3=CLIENTE, Col4=OPERARIO,
// Col5=FECHA_TRABAJO, Col6=HORA_INICIO, Col7=HORA_FIN,
// Col8=TRABAJO_REALIZADO, Col9=INCIDENCIAS, Col10=ESTADO
var PHE_TRA_ = {
  FECHA_CRE:   1,   // Col 2: FECHA_ALTA
  NOMBRE_CLI:  2,   // Col 3: CLIENTE
  NOMBRE_OPE:  3,   // Col 4: OPERARIO
  FECHA_TRA:   4,   // Col 5: FECHA_TRABAJO
  HORA_INI:    5,   // Col 6: HORA_INICIO (string "HH:MM")
  HORA_FIN:    6,   // Col 7: HORA_FIN (string "HH:MM")
  DESC:        7    // Col 8: TRABAJO_REALIZADO
};

// ════════════════════════════════════════════════════════════════
// SETUP
// ════════════════════════════════════════════════════════════════

function setupPanelHorasEmpleados() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();

  var sh = ss.getSheetByName(PHE_.HOJA);
  if (!sh) sh = ss.insertSheet(PHE_.HOJA);
  else { sh.clearContents(); sh.clearFormats(); sh.clearNotes(); }

  ss.setActiveSheet(sh);

  // ── Anchos de columna (8 cols, todos legibles) ───────────
  var anchos = [120, 190, 200, 90, 120, 90, 90, 80];
  anchos.forEach(function(w, i) { try { sh.setColumnWidth(i + 1, w); } catch (_) {} });

  var N = PHE_.NCOLS;

  // ── Fila 1: Título ───────────────────────────────────────
  sh.setRowHeight(1, 50);
  sh.getRange(1, 1, 1, N).merge()
    .setValue('👷 PANEL HORAS EMPLEADOS')
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontSize(18).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── Fila 2: Subtítulo ────────────────────────────────────
  sh.setRowHeight(2, 26);
  sh.getRange(2, 1, 1, N).merge()
    .setValue('Consulta de horas trabajadas por operario en un rango de fechas')
    .setBackground('#e8f0fe').setFontColor('#1a73e8')
    .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── Fila 3: Separador ────────────────────────────────────
  sh.setRowHeight(3, 8);
  sh.getRange(3, 1, 1, N).setBackground('#dadce0');

  // ── Fila 4: Filtros ──────────────────────────────────────
  sh.setRowHeight(4, 38);

  // Operario
  _phe_lbl_(sh, 4, 1, 'Operario');
  sh.getRange(4, 2)
    .setBackground('#ffffff').setFontSize(11).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#9aa0a6', SpreadsheetApp.BorderStyle.SOLID);

  var dvFecha = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build();

  // Desde
  _phe_lbl_(sh, 4, 3, 'Desde');
  sh.getRange(PHE_.FECHA_INI[0], PHE_.FECHA_INI[1])
    .setBackground('#ffffff').setFontSize(11).setVerticalAlignment('middle')
    .setNumberFormat('dd/MM/yyyy')
    .setDataValidation(dvFecha)
    .setBorder(true, true, true, true, false, false, '#9aa0a6', SpreadsheetApp.BorderStyle.SOLID);

  // Hasta
  _phe_lbl_(sh, 4, 5, 'Hasta');
  sh.getRange(PHE_.FECHA_FIN[0], PHE_.FECHA_FIN[1])
    .setBackground('#ffffff').setFontSize(11).setVerticalAlignment('middle')
    .setNumberFormat('dd/MM/yyyy')
    .setDataValidation(dvFecha)
    .setBorder(true, true, true, true, false, false, '#9aa0a6', SpreadsheetApp.BorderStyle.SOLID);

  // Celda H4 vacía
  sh.getRange(4, 8).setBackground('#f8f9fa');

  // Defaults: mes actual
  var hoy = new Date();
  sh.getRange(PHE_.FECHA_INI[0], PHE_.FECHA_INI[1])
    .setValue(new Date(hoy.getFullYear(), hoy.getMonth(), 1));
  sh.getRange(PHE_.FECHA_FIN[0], PHE_.FECHA_FIN[1])
    .setValue(new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0));

  // ── Fila 5: Separador fino ───────────────────────────────
  sh.setRowHeight(5, 6);
  sh.getRange(5, 1, 1, N).setBackground('#e8eaed');

  // ── Fila 6: Botón ACTUALIZAR ─────────────────────────────
  sh.setRowHeight(6, 36);
  sh.getRange(PHE_.BTN_ACTUALIZAR[0], PHE_.BTN_ACTUALIZAR[1])
    .insertCheckboxes().setValue(false);
  sh.getRange(6, 2, 1, N - 1).merge()
    .setValue('  ✅  Activa el checkbox para cargar datos')
    .setBackground('#e6f4ea').setFontColor('#137333')
    .setFontSize(11).setVerticalAlignment('middle');

  // ── Fila 7: Estado ───────────────────────────────────────
  sh.setRowHeight(7, 26);
  sh.getRange(PHE_.ESTADO[0], PHE_.ESTADO[1], 1, N).merge()
    .setValue('Selecciona operario y fechas, luego activa el checkbox.')
    .setFontColor('#5f6368').setFontSize(10).setFontStyle('italic')
    .setBackground('#fafafa').setVerticalAlignment('middle');

  // ── Fila 8: Separador ────────────────────────────────────
  sh.setRowHeight(8, 8);
  sh.getRange(8, 1, 1, N).setBackground('#dadce0');
  sh.setRowHeight(9, 10);

  // ── Filas 10-11: KPIs ────────────────────────────────────
  sh.setRowHeight(PHE_.KPI_ROW_LBL, 26);
  sh.setRowHeight(PHE_.KPI_ROW_VAL, 48);

  var kpis = [
    { col: PHE_.KPI_COL_HORAS,    lbl: '⏱  HORAS TOTALES',    bg: '#e8f0fe', txt: '#1a73e8' },
    { col: PHE_.KPI_COL_TRABAJOS, lbl: '🔨  Nº TRABAJOS',      bg: '#fce8b2', txt: '#b06000' },
    { col: PHE_.KPI_COL_CLIENTES, lbl: '👤  CLIENTES',         bg: '#e6f4ea', txt: '#137333' },
    { col: PHE_.KPI_COL_PROMEDIO, lbl: '📊  PROM. h/trabajo',  bg: '#f3e8fd', txt: '#7b1fa2' }
  ];
  kpis.forEach(function(k) {
    sh.getRange(PHE_.KPI_ROW_LBL, k.col, 1, 2).merge()
      .setValue(k.lbl).setBackground(k.bg).setFontColor(k.txt)
      .setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sh.getRange(PHE_.KPI_ROW_VAL, k.col, 1, 2).merge()
      .setValue('—').setBackground(k.bg).setFontColor(k.txt)
      .setFontSize(24).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });

  // ── Fila 12: Separador ───────────────────────────────────
  sh.setRowHeight(12, 8);
  sh.getRange(12, 1, 1, N).setBackground('#dadce0');
  sh.setRowHeight(13, 10);

  // ── Fila 14: Título tabla detalle ────────────────────────
  sh.setRowHeight(PHE_.TABLE_HDR_ROW, 32);
  sh.getRange(PHE_.TABLE_HDR_ROW, 1, 1, N).merge()
    .setValue('📋  DETALLE DE TRABAJOS')
    .setBackground('#4285f4').setFontColor('#ffffff')
    .setFontSize(11).setFontWeight('bold').setVerticalAlignment('middle');

  // Fila 15: Cabeceras de columna (= TABLE_HDR_ROW + 1, pero TABLE_START_ROW = 15)
  // Reajuste: la cabecera de columnas va en fila 14+1=15, datos desde 16
  // Simplificación: cabecera en fila 14 (merged title) y columnas en fila 15:
  sh.getRange(PHE_.TABLE_START_ROW, 1, 1, N)
    .setValues([['Fecha', 'Cliente', 'Trabajo / Descripción', 'Inicio', 'Fin', 'Horas', '', '']])
    .setBackground('#d2e3fc').setFontColor('#1a73e8').setFontWeight('bold').setFontSize(10)
    .setBorder(true, true, true, true, true, false, '#a8c7fa', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(PHE_.TABLE_START_ROW, 28);

  // Ajuste: datos reales empiezan en fila 16
  PHE_.TABLE_START_ROW = 16;

  // Filas de datos vacías (fila 16 a 75)
  _phe_limpiarTablaDetalle_(sh);

  // ── Bloques de resumen (ubicados tras la tabla) ───────────
  var afterTable = PHE_.TABLE_START_ROW + PHE_.TABLE_ROWS;  // fila 76+

  // Separador
  sh.getRange(afterTable, 1, 1, N).setBackground('#dadce0');
  sh.setRowHeight(afterTable, 6);

  // Resumen por cliente
  var cliTitle = afterTable + 1;
  PHE_.CLI_TITLE_ROW = cliTitle;
  PHE_.CLI_HDR_ROW   = cliTitle + 1;
  PHE_.CLI_START_ROW = cliTitle + 2;

  sh.setRowHeight(cliTitle, 32);
  sh.getRange(cliTitle, 1, 1, N).merge()
    .setValue('👤  RESUMEN POR CLIENTE')
    .setBackground('#4285f4').setFontColor('#ffffff')
    .setFontSize(11).setFontWeight('bold').setVerticalAlignment('middle');
  sh.getRange(cliTitle + 1, 1, 1, N)
    .setValues([['Cliente', 'Horas', 'Trabajos', '', '', '', '', '']])
    .setBackground('#d2e3fc').setFontColor('#1a73e8').setFontWeight('bold').setFontSize(10);
  sh.getRange(cliTitle + 2, 1, PHE_.CLI_ROWS, N)
    .setBackground('#ffffff').setFontSize(10);
  sh.setRowHeight(cliTitle + 1, 26);

  // Resumen mensual
  var mesTitle = cliTitle + 2 + PHE_.CLI_ROWS + 1;
  PHE_.MES_TITLE_ROW = mesTitle;
  PHE_.MES_HDR_ROW   = mesTitle + 1;
  PHE_.MES_START_ROW = mesTitle + 2;

  sh.getRange(mesTitle - 1, 1, 1, N).setBackground('#dadce0');
  sh.setRowHeight(mesTitle - 1, 6);
  sh.setRowHeight(mesTitle, 32);
  sh.getRange(mesTitle, 1, 1, N).merge()
    .setValue('📅  RESUMEN MENSUAL')
    .setBackground('#4285f4').setFontColor('#ffffff')
    .setFontSize(11).setFontWeight('bold').setVerticalAlignment('middle');
  sh.getRange(mesTitle + 1, 1, 1, N)
    .setValues([['Mes', 'Horas', 'Trabajos', '', '', '', '', '']])
    .setBackground('#d2e3fc').setFontColor('#1a73e8').setFontWeight('bold').setFontSize(10);
  sh.getRange(mesTitle + 2, 1, PHE_.MES_ROWS, N)
    .setBackground('#ffffff').setFontSize(10);
  sh.setRowHeight(mesTitle + 1, 26);

  // Exportar
  var expRow = mesTitle + 2 + PHE_.MES_ROWS + 2;
  PHE_.EXP_ROW = expRow;
  sh.getRange(expRow - 1, 1, 1, N).setBackground('#dadce0');
  sh.setRowHeight(expRow - 1, 6);
  sh.getRange(expRow, 1, 1, N).merge()
    .setValue('📄 PDF: Menú ⚙️ → 👷 HORAS EMPLEADOS — Exportar PDF    |    📊 Excel: Menú ⚙️ → 👷 HORAS EMPLEADOS — Exportar Excel')
    .setFontColor('#1a73e8').setFontSize(10).setFontStyle('italic')
    .setBackground('#e8f0fe').setVerticalAlignment('middle');
  sh.setRowHeight(expRow, 28);

  // ── Dropdown operarios ────────────────────────────────────
  _phe_refreshDropdownOperarios_(sh);

  // ── Freeze fila 4 (filtros siempre visibles) ─────────────
  sh.setFrozenRows(4);

  SpreadsheetApp.flush();
  ui.alert('👷 Panel Horas Empleados',
    '✅ PANEL_HORAS_EMPLEADOS creado.\n\n' +
    '1. Selecciona operario en B4\n' +
    '2. Ajusta fechas en D4 y F4\n' +
    '3. Activa el checkbox en A6 para cargar datos',
    ui.ButtonSet.OK);
}

// ════════════════════════════════════════════════════════════════
// ACTUALIZAR
// ════════════════════════════════════════════════════════════════

function actualizarPanelHorasEmpleados() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(PHE_.HOJA);
  if (!sh) {
    SpreadsheetApp.getUi().alert('Crea primero el panel:\nSistema → 👷 HORAS EMPLEADOS — Crear / Recrear panel');
    return;
  }

  _phe_estado_(sh, '⏳  Cargando datos…', '#f9ab00');
  SpreadsheetApp.flush();

  // ── Leer filtros ─────────────────────────────────────────
  var operario = String(sh.getRange(PHE_.OPERARIO[0], PHE_.OPERARIO[1]).getValue() || '').trim().replace(/\s+/g, ' ');
  var fechaIni = sh.getRange(PHE_.FECHA_INI[0], PHE_.FECHA_INI[1]).getValue();
  var fechaFin = sh.getRange(PHE_.FECHA_FIN[0], PHE_.FECHA_FIN[1]).getValue();

  if (!operario || operario === '(sin operarios)') {
    _phe_estado_(sh, '⚠️  Selecciona un operario en B4.', '#f9ab00'); return;
  }
  if (!(fechaIni instanceof Date) || !(fechaFin instanceof Date)) {
    _phe_estado_(sh, '⚠️  Introduce fechas válidas en D4 (Desde) y F4 (Hasta).', '#f9ab00'); return;
  }

  var dIni = new Date(fechaIni.getFullYear(), fechaIni.getMonth(), fechaIni.getDate());
  var dFin = new Date(fechaFin.getFullYear(), fechaFin.getMonth(), fechaFin.getDate(), 23, 59, 59);

  // ── Leer TRABAJOS ─────────────────────────────────────────
  var shTra = ss.getSheetByName('TRABAJOS');
  if (!shTra || shTra.getLastRow() < 2) {
    _phe_estado_(sh, '❌  Hoja TRABAJOS vacía o no encontrada.', '#ea4335'); return;
  }
  var data = shTra.getRange(2, 1, shTra.getLastRow() - 1, 10).getValues();

  // ── Filtrar ───────────────────────────────────────────────
  var MESES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  var esTodos = operario.toLowerCase() === 'todos';
  var filas   = [];

  data.forEach(function(row) {
    var ope = String(row[PHE_TRA_.NOMBRE_OPE] || '').trim().replace(/\s+/g, ' ');
    if (!esTodos && ope.toLowerCase().indexOf(operario.toLowerCase()) < 0) return;

    // Fecha del trabajo
    var fecha = row[PHE_TRA_.FECHA_TRA];
    if (!(fecha instanceof Date) || isNaN(fecha.getTime()))
      fecha = row[PHE_TRA_.FECHA_CRE];
    if (!(fecha instanceof Date) || isNaN(fecha.getTime())) return;

    // Comparar solo fecha (sin hora)
    var dFecha = new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate());
    if (dFecha < dIni || dFecha > dFin) return;

    // Horas: calcular desde HORA_INICIO y HORA_FIN (strings "HH:MM")
    var inicioStr = String(row[PHE_TRA_.HORA_INI] || '').trim();
    var finStr    = String(row[PHE_TRA_.HORA_FIN]  || '').trim();
    var horas     = _phe_calcHoras_(inicioStr, finStr);
    if (!inicioStr) inicioStr = '—';
    if (!finStr)    finStr    = '—';

    filas.push({
      fecha:    dFecha,
      cliente:  String(row[PHE_TRA_.NOMBRE_CLI] || '').trim(),
      trabajo:  String(row[PHE_TRA_.DESC]        || '').trim(),
      operario: ope,
      inicio:   inicioStr,
      fin:      finStr,
      horas:    horas,
      mes:      dFecha.getFullYear() + '-' + _phe_pad2_(dFecha.getMonth() + 1),
      mesLabel: MESES[dFecha.getMonth()] + ' ' + dFecha.getFullYear()
    });
  });

  filas.sort(function(a, b) { return b.fecha - a.fecha; });

  // ── KPIs ─────────────────────────────────────────────────
  var totalH   = filas.reduce(function(s, f) { return s + f.horas; }, 0);
  var nTrab    = filas.length;
  var cliSet   = {};
  filas.forEach(function(f) { cliSet[f.cliente] = true; });
  var nCli     = Object.keys(cliSet).length;
  var prom     = nTrab > 0 ? Math.round(totalH / nTrab * 10) / 10 : 0;

  sh.getRange(PHE_.KPI_ROW_VAL, PHE_.KPI_COL_HORAS,    1, 2).merge().setValue(Math.round(totalH * 10) / 10 + ' h');
  sh.getRange(PHE_.KPI_ROW_VAL, PHE_.KPI_COL_TRABAJOS, 1, 2).merge().setValue(nTrab);
  sh.getRange(PHE_.KPI_ROW_VAL, PHE_.KPI_COL_CLIENTES, 1, 2).merge().setValue(nCli);
  sh.getRange(PHE_.KPI_ROW_VAL, PHE_.KPI_COL_PROMEDIO, 1, 2).merge().setValue(prom + ' h');

  // ── Tabla detalle ─────────────────────────────────────────
  _phe_limpiarTablaDetalle_(sh);
  var maxDet = Math.min(filas.length, PHE_.TABLE_ROWS);
  if (maxDet > 0) {
    var tz  = Session.getScriptTimeZone();
    var detData = filas.slice(0, maxDet).map(function(f) {
      var fechaStr = Utilities.formatDate(f.fecha, tz, 'dd/MM/yyyy');
      return [
        fechaStr,
        f.cliente,
        f.trabajo,
        f.inicio,
        f.fin,
        f.horas > 0 ? f.horas : '—',
        esTodos ? f.operario : '',
        ''
      ];
    });
    sh.getRange(PHE_.TABLE_START_ROW, 1, maxDet, PHE_.NCOLS).setValues(detData);
    for (var i = 0; i < maxDet; i++) {
      sh.getRange(PHE_.TABLE_START_ROW + i, 1, 1, PHE_.NCOLS)
        .setBackground(i % 2 === 0 ? '#ffffff' : '#f8f9fa');
    }
  }

  // ── Resumen por cliente ───────────────────────────────────
  var cliMap = {};
  filas.forEach(function(f) {
    if (!cliMap[f.cliente]) cliMap[f.cliente] = { h: 0, t: 0 };
    cliMap[f.cliente].h += f.horas;
    cliMap[f.cliente].t += 1;
  });
  var cliArr = Object.keys(cliMap).sort().map(function(c) {
    return [c, Math.round(cliMap[c].h * 10) / 10, cliMap[c].t, '', '', '', '', ''];
  });

  sh.getRange(PHE_.CLI_START_ROW, 1, PHE_.CLI_ROWS, PHE_.NCOLS).clearContent().setBackground('#ffffff');
  var maxCli = Math.min(cliArr.length, PHE_.CLI_ROWS);
  if (maxCli > 0) {
    sh.getRange(PHE_.CLI_START_ROW, 1, maxCli, PHE_.NCOLS).setValues(cliArr);
    for (var j = 0; j < maxCli; j++) {
      sh.getRange(PHE_.CLI_START_ROW + j, 1, 1, 3)
        .setBackground(j % 2 === 0 ? '#ffffff' : '#f8f9fa');
    }
  }

  // ── Resumen mensual ───────────────────────────────────────
  var mesMap = {}, mesOrd = [];
  filas.forEach(function(f) {
    if (!mesMap[f.mes]) { mesMap[f.mes] = { lbl: f.mesLabel, h: 0, t: 0 }; mesOrd.push(f.mes); }
    mesMap[f.mes].h += f.horas;
    mesMap[f.mes].t += 1;
  });
  mesOrd = mesOrd.filter(function(v, i, a) { return a.indexOf(v) === i; }).sort();
  var mesArr = mesOrd.map(function(m) {
    return [mesMap[m].lbl, Math.round(mesMap[m].h * 10) / 10, mesMap[m].t, '', '', '', '', ''];
  });

  sh.getRange(PHE_.MES_START_ROW, 1, PHE_.MES_ROWS, PHE_.NCOLS).clearContent().setBackground('#ffffff');
  var maxMes = Math.min(mesArr.length, PHE_.MES_ROWS - 1);
  if (maxMes > 0) {
    sh.getRange(PHE_.MES_START_ROW, 1, maxMes, PHE_.NCOLS).setValues(mesArr);
    sh.getRange(PHE_.MES_START_ROW + maxMes, 1, 1, PHE_.NCOLS)
      .setValues([['TOTAL', Math.round(totalH * 10) / 10, nTrab, '', '', '', '', '']])
      .setBackground('#e8eaed').setFontWeight('bold');
  }

  // ── Estado ───────────────────────────────────────────────
  var tz2 = Session.getScriptTimeZone();
  var per = Utilities.formatDate(dIni, tz2, 'dd/MM/yyyy') + ' → ' +
            Utilities.formatDate(dFin, tz2, 'dd/MM/yyyy');
  _phe_estado_(sh,
    '✅  ' + operario + '  |  ' + per + '  |  ' +
    nTrab + ' trabajos  |  ' + Math.round(totalH * 10) / 10 + ' h totales',
    '#137333');

  logEvent('INFO', 'actualizarPanelHorasEmpleados',
    'operario=' + operario + ' trabajos=' + nTrab + ' horas=' + totalH, '');
}

// ════════════════════════════════════════════════════════════════
// EXPORTAR PDF / EXCEL
// ════════════════════════════════════════════════════════════════

function exportarHorasEmpleadosPDF() {
  var ss  = SpreadsheetApp.getActive();
  var sh  = ss.getSheetByName(PHE_.HOJA);
  if (!sh) { SpreadsheetApp.getUi().alert('Crea primero el panel.'); return; }
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
            '/export?format=pdf&gid=' + sh.getSheetId() +
            '&size=A4&portrait=false&fitw=true&gridlines=false' +
            '&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5';
  sh.getRange(PHE_.EXP_ROW + 1, 1, 1, PHE_.NCOLS).merge()
    .setValue('📄 PDF → ' + url).setFontColor('#1a73e8').setFontSize(9);
  SpreadsheetApp.getUi().alert('📄 Exportar PDF',
    'Enlace generado en la hoja.\nTambién: Archivo → Descargar → PDF', SpreadsheetApp.getUi().ButtonSet.OK);
}

function exportarHorasEmpleadosExcel() {
  var ss  = SpreadsheetApp.getActive();
  var sh  = ss.getSheetByName(PHE_.HOJA);
  if (!sh) { SpreadsheetApp.getUi().alert('Crea primero el panel.'); return; }
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
            '/export?format=xlsx&gid=' + sh.getSheetId();
  sh.getRange(PHE_.EXP_ROW + 2, 1, 1, PHE_.NCOLS).merge()
    .setValue('📊 Excel → ' + url).setFontColor('#1a73e8').setFontSize(9);
  SpreadsheetApp.getUi().alert('📊 Exportar Excel',
    'Enlace generado en la hoja.\nTambién: Archivo → Descargar → Microsoft Excel', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ════════════════════════════════════════════════════════════════
// DROPDOWN OPERARIOS
// ════════════════════════════════════════════════════════════════

function _phe_refreshDropdownOperarios_(sh) {
  try {
    var nombres = _phe_getNombresOperarios_();
    sh.getRange(PHE_.OPERARIO[0], PHE_.OPERARIO[1])
      .setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(nombres, true).setAllowInvalid(true).build());
  } catch (e) {
    Logger.log('[PHE] dropdown operarios: ' + e.message);
  }
}

function _phe_getNombresOperarios_() {
  var nombres = ['Todos'];
  var rows = SpreadsheetApp.getActive().getSheetByName('OPERARIOS');
  if (rows && rows.getLastRow() > 1) {
    rows.getRange(2, 1, rows.getLastRow() - 1, 8).getValues().forEach(function(r) {
      var activo = r[7];
      if (activo === false || String(activo).toUpperCase() === 'FALSE') return;
      var n = (String(r[2] || '') + ' ' + String(r[3] || '')).trim().replace(/\s+/g, ' ');
      if (n) nombres.push(n);
    });
  }
  if (nombres.length === 1) nombres.push('(sin operarios)');
  return nombres;
}

function refrescarDropdownOperarios() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(PHE_.HOJA);
  if (!sh) {
    SpreadsheetApp.getUi().alert('Crea primero el panel:\nSistema → 👷 HORAS EMPLEADOS — Crear / Recrear panel');
    return;
  }
  var nombres = _phe_getNombresOperarios_();
  sh.getRange(PHE_.OPERARIO[0], PHE_.OPERARIO[1])
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(nombres, true).setAllowInvalid(true).build());
  SpreadsheetApp.getUi().alert(
    '✅ Dropdown actualizado\n\nOperarios encontrados:\n• ' + nombres.join('\n• ')
  );
}

// ════════════════════════════════════════════════════════════════
// HANDLER onEdit
// ════════════════════════════════════════════════════════════════

function onEditPanelHorasEmpleados_(e) {
  var row = e.range.getRow();
  var col = e.range.getColumn();
  if (row === PHE_.BTN_ACTUALIZAR[0] && col === PHE_.BTN_ACTUALIZAR[1] && e.value === 'TRUE') {
    e.range.setValue(false);
    actualizarPanelHorasEmpleados();
  }
}

// ════════════════════════════════════════════════════════════════
// HELPERS
// ════════════════════════════════════════════════════════════════

function _phe_lbl_(sh, row, col, text) {
  sh.getRange(row, col)
    .setValue(text)
    .setBackground('#f1f3f4').setFontColor('#5f6368')
    .setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
}

function _phe_estado_(sh, msg, color) {
  sh.getRange(PHE_.ESTADO[0], PHE_.ESTADO[1], 1, PHE_.NCOLS).merge()
    .setValue(msg).setFontColor(color || '#3c4043')
    .setFontSize(10).setFontStyle('normal').setVerticalAlignment('middle');
}

function _phe_limpiarTablaDetalle_(sh) {
  var blanks = [];
  for (var i = 0; i < PHE_.TABLE_ROWS; i++)
    blanks.push(['', '', '', '', '', '', '', '']);
  sh.getRange(PHE_.TABLE_START_ROW, 1, PHE_.TABLE_ROWS, PHE_.NCOLS)
    .setValues(blanks).setBackground('#ffffff')
    .setFontSize(10).setFontWeight('normal').setFontColor('#3c4043');
}

function _phe_pad2_(n) { return n < 10 ? '0' + n : String(n); }

function _phe_charCodes_(s) {
  var out = [];
  for (var i = 0; i < s.length; i++) out.push(s.charCodeAt(i));
  return out.join(',');
}

function _phe_calcHoras_(ini, fin) {
  if (!ini || !fin) return 0;
  var pi = ini.match(/^(\d{1,2}):(\d{2})/);
  var pf = fin.match(/^(\d{1,2}):(\d{2})/);
  if (!pi || !pf) return 0;
  var mIni = parseInt(pi[1]) * 60 + parseInt(pi[2]);
  var mFin = parseInt(pf[1]) * 60 + parseInt(pf[2]);
  if (mFin <= mIni) return 0;
  return Math.round((mFin - mIni) / 60 * 10) / 10;
}

// ════════════════════════════════════════════════════════════════
// DIAGNÓSTICO TRABAJOS
// ════════════════════════════════════════════════════════════════

function diagnosticarTrabajosPHE() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var sh = ss.getSheetByName(PHE_.HOJA);

  // Leer filtro actual del panel
  var operario = sh ? String(sh.getRange(PHE_.OPERARIO[0], PHE_.OPERARIO[1]).getValue() || '').trim() : '(panel no encontrado)';
  var fechaIni = sh ? sh.getRange(PHE_.FECHA_INI[0], PHE_.FECHA_INI[1]).getValue() : null;
  var fechaFin = sh ? sh.getRange(PHE_.FECHA_FIN[0], PHE_.FECHA_FIN[1]).getValue() : null;

  var shTra = ss.getSheetByName('TRABAJOS');
  if (!shTra || shTra.getLastRow() < 2) {
    ui.alert('❌ TRABAJOS vacía o no encontrada');
    return;
  }

  var lastRow = shTra.getLastRow();
  var data = shTra.getRange(2, 1, Math.min(lastRow - 1, 20), 10).getValues();

  var msg = '🔍 FILTRO ACTIVO:\n';
  msg += '  Operario: "' + operario + '"\n';
  msg += '  Desde: ' + fechaIni + '\n';
  msg += '  Hasta: ' + fechaFin + '\n\n';
  msg += 'TRABAJOS (primeras ' + data.length + ' filas):\n';
  msg += 'Col3(CLIENTE) | Col4(OPERARIO) | Col5(FECHA_TRA)\n';
  msg += '─────────────────────────────────────────\n';

  var coincidencias = 0;
  data.forEach(function(r, i) {
    var cli  = String(r[2] || '').trim();
    var ope  = String(r[3] || '').trim();
    var fTra = r[4];
    var match = ope.toLowerCase().indexOf(operario.toLowerCase()) >= 0;
    if (match) coincidencias++;
    msg += (i+2) + ': cli="' + cli + '" ope="' + ope + '" fecha=' + fTra + (match ? ' ✅' : '') + '\n';
  });

  msg += '\n✅ Coincidencias con "' + operario + '": ' + coincidencias;
  msg += '\nTotal filas TRABAJOS: ' + (lastRow - 1);

  // Códigos de carácter para detectar invisibles
  var opeRef = data.length > 0 ? String(data[0][3] || '').trim() : '';
  msg += '\n\nCódigos (filtro B4):\n' + _phe_charCodes_(operario);
  msg += '\nCódigos (TRABAJOS fila 2):\n' + _phe_charCodes_(opeRef);

  ui.alert('Diagnóstico TRABAJOS', msg, ui.ButtonSet.OK);
}

// ════════════════════════════════════════════════════════════════
// DIAGNÓSTICO OPERARIOS
// ════════════════════════════════════════════════════════════════

function diagnosticarOperariosPHE() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var sh = ss.getSheetByName('OPERARIOS');

  if (!sh) {
    ui.alert('❌ No existe la hoja OPERARIOS');
    return;
  }

  var lastRow = sh.getLastRow();
  var msg = '📋 Hoja OPERARIOS\n';
  msg += '   Última fila: ' + lastRow + '\n\n';

  if (lastRow < 2) {
    ui.alert(msg + '⚠️  Sin datos (solo cabecera o vacía)');
    return;
  }

  var data = sh.getRange(2, 1, lastRow - 1, 8).getValues();
  msg += 'Fila | Col3(Nombre) | Col4(Apellidos) | Col8(Activo)\n';
  msg += '----------------------------------------------\n';
  data.forEach(function(r, i) {
    msg += (i + 2) + '    | ' + r[2] + ' | ' + r[3] + ' | ' + JSON.stringify(r[7]) + '\n';
  });

  // También mostrar lo que pondrá el dropdown
  var nombres = ['Todos'];
  data.forEach(function(r) {
    var activo = r[7];
    if (activo === false || String(activo).toUpperCase() === 'FALSE') return;
    var n = (String(r[2] || '') + ' ' + String(r[3] || '')).trim();
    if (n) nombres.push(n);
  });
  msg += '\n✅ Dropdown mostraría:\n' + nombres.join(', ');

  ui.alert('Diagnóstico Operarios PHE', msg, ui.ButtonSet.OK);
}
