// ============================================================
// PanelHorasEmpleados.gs — Panel de horas trabajadas por operario
// ============================================================

// ── Coordenadas del panel ────────────────────────────────────
var PHE_ = {
  HOJA: 'PANEL_HORAS_EMPLEADOS',

  // Filtros — fila 4
  OPERARIO:  [4, 3],   // C4
  FECHA_INI: [4, 6],   // F4
  FECHA_FIN: [4, 9],   // I4

  // Botón y estado — filas 6-7
  BTN_ACTUALIZAR: [6, 3],   // C6 — checkbox
  ESTADO:         [7, 3],   // C7 — mensaje estado

  // KPIs — fila 10 (valores en fila 11)
  KPI_ROW_LBL: 10,
  KPI_ROW_VAL: 11,
  KPI_COL_HORAS:    2,   // B
  KPI_COL_TRABAJOS: 5,   // E
  KPI_COL_CLIENTES: 8,   // H
  KPI_COL_PROMEDIO: 11,  // K

  // Tabla detalle
  TABLE_HDR_ROW:   14,
  TABLE_START_ROW: 15,
  TABLE_ROWS:      60,  // filas 15-74
  TC_FECHA:    1,  // A
  TC_CLIENTE:  2,  // B
  TC_TRABAJO:  3,  // C (ancho)
  TC_INICIO:   5,  // E
  TC_FIN:      6,  // F
  TC_HORAS:    7,  // G

  // Resumen por cliente (tras tabla)
  CLI_TITLE_ROW: 77,
  CLI_HDR_ROW:   78,
  CLI_START_ROW: 79,
  CLI_ROWS:      25,

  // Resumen mensual
  MES_TITLE_ROW: 106,
  MES_HDR_ROW:   107,
  MES_START_ROW: 108,
  MES_ROWS:      13,  // 12 meses + total

  // Exportar
  EXP_ROW: 123
};

// ── Índices columnas TRABAJOS (0-based) ──────────────────────
var PHE_TRA_ = {
  ID:           0,
  FECHA_CRE:    1,
  ID_CLI:       2,
  NOMBRE_CLI:   3,
  ID_OPE:       4,
  NOMBRE_OPE:   5,
  DESC:         6,
  TIPO:         7,
  FECHA_INI:    9,
  FECHA_FIN:   10,
  ESTADO:      11,
  HORAS:       12
};

// ════════════════════════════════════════════════════════════════
// SETUP — Crea o recrea la hoja PANEL_HORAS_EMPLEADOS
// ════════════════════════════════════════════════════════════════

function setupPanelHorasEmpleados() {
  var ss  = SpreadsheetApp.getActive();
  var ui  = SpreadsheetApp.getUi();

  // Crear o limpiar hoja
  var sh = ss.getSheetByName(PHE_.HOJA);
  if (!sh) {
    sh = ss.insertSheet(PHE_.HOJA);
  } else {
    sh.clearContents();
    sh.clearFormats();
    sh.clearNotes();
  }

  ss.setActiveSheet(sh);

  // Anchos de columna
  var anchos = [20, 140, 140, 20, 120, 90, 90, 90, 90, 120, 20, 120];
  anchos.forEach(function(w, i) {
    try { sh.setColumnWidth(i + 1, w); } catch (_) {}
  });

  // ── Título ───────────────────────────────────────────────
  sh.getRange(1, 1, 1, 12).merge()
    .setValue('👷 PANEL HORAS EMPLEADOS')
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 48);

  sh.getRange(2, 1, 1, 12).merge()
    .setValue('Consulta de horas trabajadas por operario en un rango de fechas')
    .setBackground('#e8f0fe').setFontColor('#1a73e8')
    .setFontSize(10).setHorizontalAlignment('center');
  sh.setRowHeight(2, 24);

  // ── Separador ────────────────────────────────────────────
  sh.getRange(3, 1, 1, 12).setBackground('#dadce0');
  sh.setRowHeight(3, 6);

  // ── Filtros — fila 4 ─────────────────────────────────────
  sh.setRowHeight(4, 36);
  _phe_label_(sh, 4, 1, 'Operario');
  sh.getRange(4, 2, 1, 2).merge()
    .setBackground('#ffffff').setBorder(true, true, true, true, false, false, '#dadce0', SpreadsheetApp.BorderStyle.SOLID)
    .setFontSize(11).setVerticalAlignment('middle');

  _phe_label_(sh, 4, 4, 'Desde');
  sh.getRange(4, 5, 1, 2).merge()
    .setBackground('#ffffff').setBorder(true, true, true, true, false, false, '#dadce0', SpreadsheetApp.BorderStyle.SOLID)
    .setNumberFormat('dd/MM/yyyy').setFontSize(11).setVerticalAlignment('middle');

  _phe_label_(sh, 4, 7, 'Hasta');
  sh.getRange(4, 8, 1, 2).merge()
    .setBackground('#ffffff').setBorder(true, true, true, true, false, false, '#dadce0', SpreadsheetApp.BorderStyle.SOLID)
    .setNumberFormat('dd/MM/yyyy').setFontSize(11).setVerticalAlignment('middle');

  // Defaults de fecha: mes actual
  var hoy   = new Date();
  var ini   = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
  var fin   = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0);
  sh.getRange(PHE_.FECHA_INI[0], PHE_.FECHA_INI[1]).setValue(ini);
  sh.getRange(PHE_.FECHA_FIN[0], PHE_.FECHA_FIN[1]).setValue(fin);

  // ── Botón Actualizar — fila 6 ────────────────────────────
  sh.setRowHeight(5, 8);
  sh.setRowHeight(6, 34);
  sh.getRange(6, 1, 1, 2).merge()
    .setValue('□ ACTUALIZAR PANEL')
    .setBackground('#34a853').setFontColor('#ffffff')
    .setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  // Checkbox real en C6
  sh.getRange(PHE_.BTN_ACTUALIZAR[0], PHE_.BTN_ACTUALIZAR[1])
    .insertCheckboxes().setValue(false);
  sh.getRange(6, 4, 1, 9).merge().setBackground('#f8f9fa');

  // Estado — fila 7
  sh.setRowHeight(7, 28);
  sh.getRange(PHE_.ESTADO[0], PHE_.ESTADO[1], 1, 10).merge()
    .setValue('Selecciona operario y fechas, luego activa el checkbox ACTUALIZAR.')
    .setFontColor('#5f6368').setFontSize(10).setFontStyle('italic')
    .setVerticalAlignment('middle');

  // ── Separador ────────────────────────────────────────────
  sh.getRange(8, 1, 1, 12).setBackground('#dadce0');
  sh.setRowHeight(8, 6);
  sh.setRowHeight(9, 12);

  // ── KPIs — filas 10-11 ───────────────────────────────────
  sh.setRowHeight(10, 24);
  sh.setRowHeight(11, 44);

  var kpis = [
    { col: PHE_.KPI_COL_HORAS,    lbl: '⏱ HORAS TOTALES',    bg: '#e8f0fe', clr: '#1a73e8', key: 'kpi_horas'    },
    { col: PHE_.KPI_COL_TRABAJOS, lbl: '🔨 Nº TRABAJOS',      bg: '#fce8b2', clr: '#b06000', key: 'kpi_trabajos' },
    { col: PHE_.KPI_COL_CLIENTES, lbl: '👤 CLIENTES',         bg: '#e6f4ea', clr: '#137333', key: 'kpi_clientes' },
    { col: PHE_.KPI_COL_PROMEDIO, lbl: '📊 PROMEDIO h/Trab.', bg: '#f3e8fd', clr: '#7b1fa2', key: 'kpi_promedio' }
  ];
  kpis.forEach(function(k) {
    sh.getRange(PHE_.KPI_ROW_LBL, k.col, 1, 3).merge()
      .setValue(k.lbl).setBackground(k.bg).setFontColor(k.clr)
      .setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(PHE_.KPI_ROW_VAL, k.col, 1, 3).merge()
      .setValue('—').setBackground(k.bg).setFontColor(k.clr)
      .setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  });

  // ── Separador ────────────────────────────────────────────
  sh.getRange(12, 1, 1, 12).setBackground('#dadce0');
  sh.setRowHeight(12, 6);
  sh.setRowHeight(13, 16);

  // ── Tabla detalle — cabecera fila 14 ─────────────────────
  sh.setRowHeight(PHE_.TABLE_HDR_ROW, 30);
  sh.getRange(PHE_.TABLE_HDR_ROW, 1, 1, 12).merge()
    .setValue('📋 DETALLE DE TRABAJOS')
    .setBackground('#4285f4').setFontColor('#ffffff')
    .setFontSize(11).setFontWeight('bold')
    .setVerticalAlignment('middle');

  var tableHdr = ['Fecha', 'Cliente', 'Trabajo / Descripción', '', 'Inicio', 'Fin', 'Horas', '', '', '', '', ''];
  sh.getRange(PHE_.TABLE_HDR_ROW + 1, 1, 1, 12)
    .setValues([tableHdr])
    .setBackground('#d2e3fc').setFontColor('#1a73e8')
    .setFontWeight('bold').setFontSize(10)
    .setBorder(true, true, true, true, true, true, '#a8c7fa', SpreadsheetApp.BorderStyle.SOLID);

  // Filas de datos vacías
  _phe_limpiarTablaDetalle_(sh);

  // ── Resumen clientes ─────────────────────────────────────
  var cliTitleRow = PHE_.TABLE_START_ROW + PHE_.TABLE_ROWS + 2;
  PHE_.CLI_TITLE_ROW = cliTitleRow;
  PHE_.CLI_HDR_ROW   = cliTitleRow + 1;
  PHE_.CLI_START_ROW = cliTitleRow + 2;

  sh.setRowHeight(cliTitleRow, 30);
  sh.getRange(cliTitleRow, 1, 1, 12).merge()
    .setValue('👤 RESUMEN POR CLIENTE')
    .setBackground('#4285f4').setFontColor('#ffffff')
    .setFontSize(11).setFontWeight('bold').setVerticalAlignment('middle');
  sh.getRange(cliTitleRow + 1, 1, 1, 3)
    .setValues([['Cliente', 'Horas', 'Trabajos']])
    .setBackground('#d2e3fc').setFontColor('#1a73e8')
    .setFontWeight('bold').setFontSize(10);
  sh.getRange(cliTitleRow + 2, 1, PHE_.CLI_ROWS, 3).clearContent()
    .setBackground('#ffffff').setFontColor('#3c4043').setFontSize(10);

  // ── Resumen mensual ───────────────────────────────────────
  var mesTitleRow = cliTitleRow + 2 + PHE_.CLI_ROWS + 2;
  PHE_.MES_TITLE_ROW = mesTitleRow;
  PHE_.MES_HDR_ROW   = mesTitleRow + 1;
  PHE_.MES_START_ROW = mesTitleRow + 2;

  sh.setRowHeight(mesTitleRow, 30);
  sh.getRange(mesTitleRow, 1, 1, 12).merge()
    .setValue('📅 RESUMEN MENSUAL')
    .setBackground('#4285f4').setFontColor('#ffffff')
    .setFontSize(11).setFontWeight('bold').setVerticalAlignment('middle');
  sh.getRange(mesTitleRow + 1, 1, 1, 3)
    .setValues([['Mes', 'Horas', 'Trabajos']])
    .setBackground('#d2e3fc').setFontColor('#1a73e8')
    .setFontWeight('bold').setFontSize(10);
  sh.getRange(mesTitleRow + 2, 1, PHE_.MES_ROWS, 3).clearContent()
    .setBackground('#ffffff').setFontColor('#3c4043').setFontSize(10);

  // ── Exportar ─────────────────────────────────────────────
  var expRow = mesTitleRow + 2 + PHE_.MES_ROWS + 2;
  PHE_.EXP_ROW = expRow;

  sh.getRange(expRow, 1, 1, 12).setBackground('#dadce0');
  sh.setRowHeight(expRow, 6);
  sh.getRange(expRow + 1, 1, 1, 3).merge()
    .setValue('📄 Exportar PDF → menú ⚙️ Sistema → 👷 HORAS EMPLEADOS — Exportar PDF')
    .setFontColor('#1a73e8').setFontSize(10).setFontStyle('italic');
  sh.getRange(expRow + 2, 1, 1, 3).merge()
    .setValue('📊 Exportar Excel → menú ⚙️ Sistema → 👷 HORAS EMPLEADOS — Exportar Excel')
    .setFontColor('#1a73e8').setFontSize(10).setFontStyle('italic');

  // ── Dropdown operarios ────────────────────────────────────
  _phe_refreshDropdownOperarios_(sh);

  // Freeze filtros
  sh.setFrozenRows(4);

  SpreadsheetApp.flush();
  ui.alert('👷 Panel Horas Empleados',
    '✅ PANEL_HORAS_EMPLEADOS creado.\n\n' +
    '1. Selecciona un operario en la fila 4.\n' +
    '2. Ajusta las fechas Desde / Hasta.\n' +
    '3. Activa el checkbox ACTUALIZAR.',
    ui.ButtonSet.OK);
}

// ════════════════════════════════════════════════════════════════
// ACTUALIZAR — lee filtros, consulta TRABAJOS, escribe resultados
// ════════════════════════════════════════════════════════════════

function actualizarPanelHorasEmpleados() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(PHE_.HOJA);
  if (!sh) {
    SpreadsheetApp.getUi().alert('Hoja PANEL_HORAS_EMPLEADOS no encontrada.\nEjecuta primero: Sistema → 👷 HORAS EMPLEADOS — Crear panel.');
    return;
  }

  _phe_estado_(sh, '⏳ Actualizando…', '#f9ab00');
  SpreadsheetApp.flush();

  // ── Leer filtros ─────────────────────────────────────────
  var operario  = String(sh.getRange(PHE_.OPERARIO[0],  PHE_.OPERARIO[1]).getValue()  || '').trim();
  var fechaIni  = sh.getRange(PHE_.FECHA_INI[0], PHE_.FECHA_INI[1]).getValue();
  var fechaFin  = sh.getRange(PHE_.FECHA_FIN[0], PHE_.FECHA_FIN[1]).getValue();

  if (!operario) {
    _phe_estado_(sh, '⚠️ Selecciona un operario en C4.', '#f9ab00');
    return;
  }
  if (!(fechaIni instanceof Date) || !(fechaFin instanceof Date)) {
    _phe_estado_(sh, '⚠️ Introduce fechas válidas en F4 e I4.', '#f9ab00');
    return;
  }

  // Normalizar: inicio a las 00:00, fin a las 23:59
  var dIni = new Date(fechaIni.getFullYear(), fechaIni.getMonth(), fechaIni.getDate());
  var dFin = new Date(fechaFin.getFullYear(), fechaFin.getMonth(), fechaFin.getDate(), 23, 59, 59);

  // ── Leer TRABAJOS ─────────────────────────────────────────
  var shTra = ss.getSheetByName('TRABAJOS');
  if (!shTra || shTra.getLastRow() < 2) {
    _phe_estado_(sh, '❌ Hoja TRABAJOS vacía o no encontrada.', '#ea4335');
    return;
  }

  var data = shTra.getRange(2, 1, shTra.getLastRow() - 1, 13).getValues();

  // ── Filtrar por operario y fechas ────────────────────────
  var MESES_ES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  var filas = [];

  data.forEach(function(row) {
    var nombreOpe = String(row[PHE_TRA_.NOMBRE_OPE] || '').trim();
    if (!nombreOpe || nombreOpe.toLowerCase().indexOf(operario.toLowerCase()) < 0) return;

    var fecha = row[PHE_TRA_.FECHA_INI];
    if (!(fecha instanceof Date) || isNaN(fecha.getTime())) {
      fecha = row[PHE_TRA_.FECHA_CRE];
    }
    if (!(fecha instanceof Date) || isNaN(fecha.getTime())) return;
    if (fecha < dIni || fecha > dFin) return;

    var horas = parseFloat(row[PHE_TRA_.HORAS]) || 0;

    // Calcular horas de inicio-fin si hay datetime con hora
    var fechaFecha = row[PHE_TRA_.FECHA_INI];
    var fechaFinV  = row[PHE_TRA_.FECHA_FIN];
    var inicioStr  = '—';
    var finStr     = '—';
    if (fechaFecha instanceof Date && fechaFinV instanceof Date) {
      var difH = (fechaFinV.getTime() - fechaFecha.getTime()) / 3600000;
      if (difH > 0 && difH < 24 && horas === 0) horas = Math.round(difH * 100) / 100;
      var hI = fechaFecha.getHours(), mI = fechaFecha.getMinutes();
      var hF = fechaFinV.getHours(),  mF = fechaFinV.getMinutes();
      if (hI > 0 || mI > 0) inicioStr = _phe_pad2_(hI) + ':' + _phe_pad2_(mI);
      if (hF > 0 || mF > 0) finStr    = _phe_pad2_(hF) + ':' + _phe_pad2_(mF);
    }

    filas.push({
      fecha:    fecha,
      cliente:  String(row[PHE_TRA_.NOMBRE_CLI] || '').trim(),
      trabajo:  String(row[PHE_TRA_.DESC]        || '').trim(),
      inicio:   inicioStr,
      fin:      finStr,
      horas:    horas,
      mes:      fecha.getFullYear() + '-' + _phe_pad2_(fecha.getMonth() + 1),
      mesLabel: MESES_ES[fecha.getMonth()] + ' ' + fecha.getFullYear()
    });
  });

  // Ordenar por fecha desc
  filas.sort(function(a, b) { return b.fecha - a.fecha; });

  // ── KPIs ─────────────────────────────────────────────────
  var totalHoras   = filas.reduce(function(s, f) { return s + f.horas; }, 0);
  var numTrabajos  = filas.length;
  var clientesSet  = {};
  filas.forEach(function(f) { clientesSet[f.cliente] = true; });
  var numClientes  = Object.keys(clientesSet).length;
  var promedio     = numTrabajos > 0 ? Math.round(totalHoras / numTrabajos * 10) / 10 : 0;

  sh.getRange(PHE_.KPI_ROW_VAL, PHE_.KPI_COL_HORAS,    1, 3).merge().setValue(Math.round(totalHoras * 10) / 10 + ' h');
  sh.getRange(PHE_.KPI_ROW_VAL, PHE_.KPI_COL_TRABAJOS, 1, 3).merge().setValue(numTrabajos);
  sh.getRange(PHE_.KPI_ROW_VAL, PHE_.KPI_COL_CLIENTES, 1, 3).merge().setValue(numClientes);
  sh.getRange(PHE_.KPI_ROW_VAL, PHE_.KPI_COL_PROMEDIO, 1, 3).merge().setValue(promedio + ' h');

  // ── Tabla detalle ─────────────────────────────────────────
  _phe_limpiarTablaDetalle_(sh);
  var maxDet = Math.min(filas.length, PHE_.TABLE_ROWS);
  if (maxDet > 0) {
    var detData = filas.slice(0, maxDet).map(function(f, i) {
      return [
        Utilities.formatDate(f.fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
        f.cliente, f.trabajo, '', f.inicio, f.fin,
        f.horas > 0 ? f.horas : '—', '', '', '', '', ''
      ];
    });
    sh.getRange(PHE_.TABLE_START_ROW, 1, maxDet, 12).setValues(detData);
    // Colores alternos
    for (var r = 0; r < maxDet; r++) {
      var bg = r % 2 === 0 ? '#ffffff' : '#f8f9fa';
      sh.getRange(PHE_.TABLE_START_ROW + r, 1, 1, 12).setBackground(bg);
    }
  }

  // ── Resumen por cliente ───────────────────────────────────
  var cliMap = {};
  filas.forEach(function(f) {
    if (!cliMap[f.cliente]) cliMap[f.cliente] = { horas: 0, trabajos: 0 };
    cliMap[f.cliente].horas    += f.horas;
    cliMap[f.cliente].trabajos += 1;
  });
  var cliRows = Object.keys(cliMap).sort().map(function(c) {
    return [c, Math.round(cliMap[c].horas * 10) / 10, cliMap[c].trabajos, '', '', '', '', '', '', '', '', ''];
  });

  var cliStart = PHE_.CLI_START_ROW;
  sh.getRange(cliStart, 1, PHE_.CLI_ROWS, 12).clearContent().setBackground('#ffffff');
  var maxCli = Math.min(cliRows.length, PHE_.CLI_ROWS);
  if (maxCli > 0) {
    sh.getRange(cliStart, 1, maxCli, 12).setValues(cliRows);
    for (var r2 = 0; r2 < maxCli; r2++) {
      sh.getRange(cliStart + r2, 1, 1, 3).setBackground(r2 % 2 === 0 ? '#ffffff' : '#f8f9fa');
    }
  }

  // ── Resumen mensual ───────────────────────────────────────
  var mesMap = {};
  var mesOrder = [];
  filas.forEach(function(f) {
    if (!mesMap[f.mes]) { mesMap[f.mes] = { label: f.mesLabel, horas: 0, trabajos: 0 }; mesOrder.push(f.mes); }
    mesMap[f.mes].horas    += f.horas;
    mesMap[f.mes].trabajos += 1;
  });
  mesOrder = mesOrder.filter(function(v, i, a) { return a.indexOf(v) === i; }).sort();
  var mesRows = mesOrder.map(function(m) {
    return [mesMap[m].label, Math.round(mesMap[m].horas * 10) / 10, mesMap[m].trabajos,
            '', '', '', '', '', '', '', '', ''];
  });

  var mesStart = PHE_.MES_START_ROW;
  sh.getRange(mesStart, 1, PHE_.MES_ROWS, 12).clearContent().setBackground('#ffffff');
  var maxMes = Math.min(mesRows.length, PHE_.MES_ROWS - 1);
  if (maxMes > 0) {
    sh.getRange(mesStart, 1, maxMes, 12).setValues(mesRows);
    // Fila total
    var totalRow = [['TOTAL', Math.round(totalHoras * 10) / 10, numTrabajos,
                     '', '', '', '', '', '', '', '', '']];
    sh.getRange(mesStart + maxMes, 1, 1, 12).setValues(totalRow)
      .setBackground('#e8eaed').setFontWeight('bold');
  }

  // ── Estado final ─────────────────────────────────────────
  var periodo = Utilities.formatDate(dIni, Session.getScriptTimeZone(), 'dd/MM/yyyy') +
                ' → ' +
                Utilities.formatDate(dFin, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  _phe_estado_(sh,
    '✅ ' + operario + ' | ' + periodo + ' | ' +
    numTrabajos + ' trabajos | ' + Math.round(totalHoras * 10) / 10 + ' h totales',
    '#34a853');

  logEvent('INFO', 'actualizarPanelHorasEmpleados',
    'operario=' + operario + ' trabajos=' + numTrabajos + ' horas=' + totalHoras, '');
}

// ════════════════════════════════════════════════════════════════
// EXPORTAR PDF
// ════════════════════════════════════════════════════════════════

function exportarHorasEmpleadosPDF() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(PHE_.HOJA);
  if (!sh) { SpreadsheetApp.getUi().alert('Crea primero el panel.'); return; }

  var ssId  = ss.getId();
  var shId  = sh.getSheetId();
  var url   = 'https://docs.google.com/spreadsheets/d/' + ssId +
              '/export?format=pdf&gid=' + shId +
              '&size=A4&portrait=false&fitw=true&gridlines=false' +
              '&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5';

  sh.getRange(PHE_.EXP_ROW + 1, 1, 1, 6).merge()
    .setValue('📄 PDF → ' + url)
    .setFontColor('#1a73e8').setFontSize(10);

  SpreadsheetApp.getUi().alert(
    '📄 Exportar PDF',
    'Enlace generado en la hoja (fila ' + (PHE_.EXP_ROW + 1) + ').\n\n' +
    'También puedes usar:\n' +
    'Archivo → Descargar → PDF (.pdf)\n\n' +
    'URL directa copiada en la celda de exportación.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ════════════════════════════════════════════════════════════════
// EXPORTAR EXCEL
// ════════════════════════════════════════════════════════════════

function exportarHorasEmpleadosExcel() {
  var ss   = SpreadsheetApp.getActive();
  var sh   = ss.getSheetByName(PHE_.HOJA);
  if (!sh) { SpreadsheetApp.getUi().alert('Crea primero el panel.'); return; }

  var ssId = ss.getId();
  var shId = sh.getSheetId();
  var url  = 'https://docs.google.com/spreadsheets/d/' + ssId +
             '/export?format=xlsx&gid=' + shId;

  sh.getRange(PHE_.EXP_ROW + 2, 1, 1, 6).merge()
    .setValue('📊 Excel → ' + url)
    .setFontColor('#1a73e8').setFontSize(10);

  SpreadsheetApp.getUi().alert(
    '📊 Exportar Excel',
    'Enlace generado en la hoja (fila ' + (PHE_.EXP_ROW + 2) + ').\n\n' +
    'También puedes usar:\n' +
    'Archivo → Descargar → Microsoft Excel (.xlsx)\n\n' +
    'URL directa copiada en la celda de exportación.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ════════════════════════════════════════════════════════════════
// REFRESH DROPDOWN OPERARIOS
// ════════════════════════════════════════════════════════════════

function _phe_refreshDropdownOperarios_(sh) {
  try {
    var ss      = SpreadsheetApp.getActive();
    var shOpe   = ss.getSheetByName('OPERARIOS');
    var nombres = [];
    if (shOpe && shOpe.getLastRow() > 1) {
      var rows = shOpe.getRange(2, 1, shOpe.getLastRow() - 1, 8).getValues();
      rows.forEach(function(row) {
        var activo = row[7]; // col H — Activo
        if (activo === true || String(activo).toUpperCase() === 'TRUE') {
          var nombre = String(row[2] || '').trim() + ' ' + String(row[3] || '').trim();
          if (nombre.trim()) nombres.push(nombre.trim());
        }
      });
    }
    if (nombres.length === 0) nombres = ['(sin operarios)'];
    nombres.unshift('Todos');

    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(nombres, true)
      .setAllowInvalid(true)
      .build();
    sh.getRange(PHE_.OPERARIO[0], PHE_.OPERARIO[1], 1, 2).setDataValidation(rule);
  } catch (e) {
    Logger.log('[PHE] Error dropdown operarios: ' + e.message);
  }
}

// ════════════════════════════════════════════════════════════════
// HANDLER onEdit
// ════════════════════════════════════════════════════════════════

function onEditPanelHorasEmpleados_(e) {
  var sh  = e.range.getSheet();
  var row = e.range.getRow();
  var col = e.range.getColumn();

  // Checkbox ACTUALIZAR (C6)
  if (row === PHE_.BTN_ACTUALIZAR[0] && col === PHE_.BTN_ACTUALIZAR[1] && e.value === 'TRUE') {
    e.range.setValue(false);
    actualizarPanelHorasEmpleados();
  }
}

// ════════════════════════════════════════════════════════════════
// HELPERS PRIVADOS
// ════════════════════════════════════════════════════════════════

function _phe_label_(sh, row, col, text) {
  sh.getRange(row, col)
    .setValue(text)
    .setBackground('#f1f3f4').setFontColor('#5f6368')
    .setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
}

function _phe_estado_(sh, msg, color) {
  sh.getRange(PHE_.ESTADO[0], PHE_.ESTADO[1], 1, 10).merge()
    .setValue(msg).setFontColor(color || '#3c4043')
    .setFontSize(10).setFontStyle('normal');
}

function _phe_limpiarTablaDetalle_(sh) {
  var blanks = [];
  for (var i = 0; i < PHE_.TABLE_ROWS; i++) {
    blanks.push(['', '', '', '', '', '', '', '', '', '', '', '']);
  }
  sh.getRange(PHE_.TABLE_START_ROW, 1, PHE_.TABLE_ROWS, 12)
    .setValues(blanks)
    .setBackground('#ffffff').setFontColor('#3c4043')
    .setFontSize(10).setFontWeight('normal');
}

function _phe_pad2_(n) {
  return n < 10 ? '0' + n : String(n);
}
