// ============================================================
// PanelGastos.gs — PANEL_GASTOS  (FASE 4)
// Panel de revisión de gastos: filtros, resumen y tabla.
// ============================================================

// ── Constantes de layout ─────────────────────────────────────

var PG_ = {
  // Filtros: etiquetas col A (1), valores col B (2)
  F_FECHA_DESDE:  [3,  2],
  F_FECHA_HASTA:  [4,  2],
  F_TIPO:         [5,  2],
  F_ESTADO_PROC:  [6,  2],
  F_ESTADO_PAGO:  [7,  2],
  F_REVISADO:     [8,  2],
  F_ENVIAR_GES:   [9,  2],
  BTN_ACTUALIZAR: [10, 2],   // B10 — checkbox-botón

  // Resumen: etiquetas col D (4), valores col E (5)
  S_TOTAL:  [3, 5],
  S_BASE:   [4, 5],
  S_IVA:    [5, 5],
  S_N:      [6, 5],
  S_N_REV:  [7, 5],
  S_N_PAGO: [8, 5],
  S_N_GES:  [9, 5],

  // Tabla
  TABLE_HDR:   13,   // fila de cabeceras
  TABLE_START: 14,   // primera fila de datos

  // Columnas de la tabla (números de columna absolutos)
  TC_ID:       1,   // A
  TC_FECHA:    2,   // B
  TC_TIPO:     3,   // C
  TC_PROV:     4,   // D
  TC_CONC:     5,   // E
  TC_TOTAL:    6,   // F
  TC_IVA:      7,   // G
  TC_EST_PAGO: 8,   // H
  TC_EST_PROC: 9,   // I
  TC_REVISADO: 10,  // J ← checkbox interactivo
  TC_ENVIAR:   11,  // K ← checkbox interactivo
  TC_VER_FOTO: 12,  // L ← hyperlink
  TC_VER_OCR:  13,  // M ← checkbox "Ver OCR"
  TC_OBS:      14   // N
};

var PG_N_COLS = 14;

// Colores de fila según Estado_Procesamiento
var PG_COLORS = {
  'Procesado':    '#c6efce',
  'Revisar':      '#fce8b2',
  'OCR fallido':  '#f4c7c3',
  'Pendiente OCR':'#eeeeee'
};

// ─────────────────────────────────────────────────────────────
// crearPanelGastos — construye la estructura del panel
// ─────────────────────────────────────────────────────────────

/**
 * Crea (o reconstruye) la hoja PANEL_GASTOS con toda su estructura:
 * título, filtros, resumen y cabeceras de tabla.
 * Idempotente: borra y rehace si ya existe.
 */
function crearPanelGastos() {
  Logger.log('=== crearPanelGastos() ===');
  var ss = getTargetSS_();
  var sh = ss.getSheetByName('PANEL_GASTOS');

  if (!sh) {
    sh = ss.insertSheet('PANEL_GASTOS');
    Logger.log('  + Hoja PANEL_GASTOS creada.');
  } else {
    sh.clear();
    var full = sh.getRange(1, 1, sh.getMaxRows(), Math.min(sh.getMaxColumns(), 30));
    full.breakApart();
    full.clearDataValidations();
    Logger.log('  ✓ Hoja PANEL_GASTOS limpiada y reconstruida.');
  }

  // ── Anchos de columna ─────────────────────────────────────────
  var colWidths = [
    [1, 115], [2, 100], [3, 145], [4, 180], [5, 220],
    [6, 90],  [7, 80],  [8, 105], [9, 145],
    [10, 80], [11, 80], [12, 75], [13, 75], [14, 200]
  ];
  colWidths.forEach(function(w) { sh.setColumnWidth(w[0], w[1]); });

  // ── Fila 1: Título ────────────────────────────────────────────
  sh.getRange(1, 1, 1, PG_N_COLS).merge()
    .setValue('PANEL_GASTOS — Revisión de gastos')
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 40);

  // ── Fila 2: Encabezados de sección ───────────────────────────
  sh.getRange(2, 1, 1, 3).merge()
    .setValue('FILTROS');
  _pgSectionHeader_(sh, 2, 1, 3);

  sh.getRange(2, 4, 1, 2).merge()
    .setValue('RESUMEN');
  _pgSectionHeader_(sh, 2, 4, 2);

  // ── Filtros: etiquetas (filas 3-10, col A) ───────────────────
  var filtroLabels = [
    [3,  'Fecha desde'],
    [4,  'Fecha hasta'],
    [5,  'Tipo de gasto'],
    [6,  'Estado procesamiento'],
    [7,  'Estado pago'],
    [8,  'Revisado'],
    [9,  'Enviar gestor'],
    [10, '▶ Actualizar panel']
  ];
  filtroLabels.forEach(function(f) {
    sh.getRange(f[0], 1).setValue(f[1])
      .setFontWeight('bold').setHorizontalAlignment('right');
  });

  // Validaciones de los desplegables de filtro
  var opTodos = ['Todos'];
  sh.getRange(PG_.F_FECHA_DESDE[0], PG_.F_FECHA_DESDE[1]).setValue('');
  sh.getRange(PG_.F_FECHA_HASTA[0], PG_.F_FECHA_HASTA[1]).setValue('');

  sh.getRange(PG_.F_TIPO[0], PG_.F_TIPO[1])
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(opTodos.concat(TIPOS_GASTO), true)
      .setAllowInvalid(false).build())
    .setValue('Todos');

  sh.getRange(PG_.F_ESTADO_PROC[0], PG_.F_ESTADO_PROC[1])
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Todos', 'Pendiente OCR', 'Procesado', 'Revisar',
                           'OCR fallido', 'Revisado manualmente'], true)
      .setAllowInvalid(false).build())
    .setValue('Todos');

  sh.getRange(PG_.F_ESTADO_PAGO[0], PG_.F_ESTADO_PAGO[1])
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(opTodos.concat(ESTADOS_PAGO_GASTO), true)
      .setAllowInvalid(false).build())
    .setValue('Todos');

  var ruleBool = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Todos', 'Sí', 'No'], true).setAllowInvalid(false).build();
  sh.getRange(PG_.F_REVISADO[0],   PG_.F_REVISADO[1]).setDataValidation(ruleBool).setValue('Todos');
  sh.getRange(PG_.F_ENVIAR_GES[0], PG_.F_ENVIAR_GES[1]).setDataValidation(ruleBool).setValue('Todos');

  // Botón Actualizar (checkbox)
  sh.getRange(PG_.BTN_ACTUALIZAR[0], PG_.BTN_ACTUALIZAR[1])
    .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build())
    .setValue(false);

  // Estilo celdas de valor (filtros)
  sh.getRange(3, 2, 8, 1)
    .setBackground('#f8f9fa')
    .setBorder(true, true, true, true, false, false,
               '#cccccc', SpreadsheetApp.BorderStyle.SOLID);

  // ── Resumen: etiquetas (col D) y valores (col E) ─────────────
  var resumenLabels = [
    [PG_.S_TOTAL[0],  'Total gastos (€)'],
    [PG_.S_BASE[0],   'Base imponible (€)'],
    [PG_.S_IVA[0],    'IVA soportado (€)'],
    [PG_.S_N[0],      'Nº gastos'],
    [PG_.S_N_REV[0],  'Nº pend. revisión'],
    [PG_.S_N_PAGO[0], 'Nº pend. pago'],
    [PG_.S_N_GES[0],  'Nº para gestor']
  ];
  resumenLabels.forEach(function(r) {
    sh.getRange(r[0], 4).setValue(r[1])
      .setFontWeight('bold').setHorizontalAlignment('right');
    sh.getRange(r[0], 5).setValue('—')
      .setBackground('#f8f9fa').setHorizontalAlignment('right')
      .setBorder(true, true, true, true, false, false,
                 '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  });
  sh.getRange(PG_.S_TOTAL[0], PG_.S_TOTAL[1]).setNumberFormat('#,##0.00 €');
  sh.getRange(PG_.S_BASE[0],  PG_.S_BASE[1]).setNumberFormat('#,##0.00 €');
  sh.getRange(PG_.S_IVA[0],   PG_.S_IVA[1]).setNumberFormat('#,##0.00 €');

  // ── Fila 11: espaciador, Fila 12: encabezado tabla ───────────
  sh.getRange(11, 1, 1, PG_N_COLS).merge().setBackground('#ffffff');
  sh.getRange(12, 1, 1, PG_N_COLS).merge()
    .setValue('TABLA DE GASTOS FILTRADOS');
  _pgSectionHeader_(sh, 12, 1, PG_N_COLS);

  // ── Fila 13: cabeceras de la tabla ───────────────────────────
  var hdrs = [
    'ID_Gasto', 'Fecha', 'Tipo_Gasto', 'Proveedor', 'Concepto',
    'Total', 'IVA', 'Estado\nPago', 'Estado\nProc.',
    'Revisado', 'Enviar\nGestor', 'Ver\nFoto', 'Ver\nOCR', 'Observaciones'
  ];
  sh.getRange(PG_.TABLE_HDR, 1, 1, hdrs.length).setValues([hdrs])
    .setBackground('#202124').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sh.setRowHeight(PG_.TABLE_HDR, 42);

  // Fijar las primeras 13 filas
  sh.setFrozenRows(PG_.TABLE_HDR);
  sh.setFrozenColumns(1);

  SpreadsheetApp.flush();
  Logger.log('  ✓ Estructura creada.');
  Logger.log('  → Ejecuta: ⚙️ Sistema → PANEL_GASTOS — Actualizar datos');
  Logger.log('=== FIN crearPanelGastos() ===');
}

// ─────────────────────────────────────────────────────────────
// actualizarPanelGastos — carga y filtra datos de GASTOS
// ─────────────────────────────────────────────────────────────

/**
 * Lee los filtros del panel, obtiene los gastos de GASTOS,
 * aplica los filtros, escribe el resumen y la tabla con colores.
 */
function actualizarPanelGastos() {
  Logger.log('=== actualizarPanelGastos() ===');
  var ss = getTargetSS_();
  var sh = ss.getSheetByName('PANEL_GASTOS');
  if (!sh) {
    Logger.log('  PANEL_GASTOS no existe — creando...');
    crearPanelGastos();
    sh = ss.getSheetByName('PANEL_GASTOS');
  }

  // 1. Leer filtros
  var f = _pgGetFiltros_(sh);
  Logger.log('  Filtros: tipo=' + f.tipo + '  estadoProc=' + f.estadoProc +
             '  estadoPago=' + f.estadoPago + '  revisado=' + f.revisado +
             '  enviarGes=' + f.enviarGes);

  // 2. Leer todos los gastos reales de GASTOS
  var gasSheet = getSheet('GASTOS');
  var maxRows  = gasSheet.getMaxRows();
  var gastos   = [];
  if (maxRows > 1) {
    var allData = gasSheet.getRange(2, 1, maxRows - 1, HEADERS_GASTOS.length).getValues();
    gastos = allData.filter(function(row) {
      return String(row[COL_GAS.ID_GASTO - 1]).trim().indexOf('GAS-') === 0;
    });
  }
  Logger.log('  Total gastos en GASTOS: ' + gastos.length);

  // 3. Aplicar filtros
  var filtrados = _pgAplicarFiltros_(gastos, f);
  Logger.log('  Gastos tras filtro: ' + filtrados.length);

  // 4. Resumen
  _pgEscribirResumen_(sh, filtrados);

  // 5. Limpiar tabla y escribir datos
  _pgLimpiarTabla_(sh);
  if (filtrados.length > 0) {
    _pgEscribirTabla_(sh, filtrados);
  } else {
    sh.getRange(PG_.TABLE_START, 1).setValue('— Sin gastos para los filtros seleccionados —')
      .setFontStyle('italic').setFontColor('#888888');
  }

  Logger.log('=== FIN actualizarPanelGastos() ===');
}

// ─────────────────────────────────────────────────────────────
// Privadas de datos
// ─────────────────────────────────────────────────────────────

function _pgGetFiltros_(sh) {
  var v = function(r, c) { return sh.getRange(r, c).getValue(); };
  var fd = v(PG_.F_FECHA_DESDE[0], PG_.F_FECHA_DESDE[1]);
  var fh = v(PG_.F_FECHA_HASTA[0], PG_.F_FECHA_HASTA[1]);
  return {
    fechaDesde: (fd instanceof Date && !isNaN(fd.getTime())) ? fd : null,
    fechaHasta: (fh instanceof Date && !isNaN(fh.getTime())) ? fh : null,
    tipo:        String(v(PG_.F_TIPO[0],        PG_.F_TIPO[1])        || 'Todos').trim(),
    estadoProc:  String(v(PG_.F_ESTADO_PROC[0], PG_.F_ESTADO_PROC[1]) || 'Todos').trim(),
    estadoPago:  String(v(PG_.F_ESTADO_PAGO[0], PG_.F_ESTADO_PAGO[1]) || 'Todos').trim(),
    revisado:    String(v(PG_.F_REVISADO[0],    PG_.F_REVISADO[1])    || 'Todos').trim(),
    enviarGes:   String(v(PG_.F_ENVIAR_GES[0],  PG_.F_ENVIAR_GES[1]) || 'Todos').trim()
  };
}

function _pgAplicarFiltros_(gastos, f) {
  return gastos.filter(function(row) {
    // Fecha: usa Fecha_Gasto si existe, si no Fecha_Alta
    var fGas = row[COL_GAS.FECHA_GASTO - 1];
    var fRef = (fGas instanceof Date && !isNaN(fGas.getTime()))
             ? fGas : row[COL_GAS.FECHA_ALTA - 1];

    if (f.fechaDesde && fRef instanceof Date && !isNaN(fRef.getTime())) {
      if (fRef < f.fechaDesde) return false;
    }
    if (f.fechaHasta && fRef instanceof Date && !isNaN(fRef.getTime())) {
      var hasta = new Date(f.fechaHasta);
      hasta.setHours(23, 59, 59, 999);
      if (fRef > hasta) return false;
    }
    if (f.tipo !== 'Todos') {
      if (String(row[COL_GAS.TIPO_GASTO - 1]).trim() !== f.tipo) return false;
    }
    if (f.estadoProc !== 'Todos') {
      if (String(row[COL_GAS.ESTADO_PROCESAMIENTO - 1]).trim() !== f.estadoProc) return false;
    }
    if (f.estadoPago !== 'Todos') {
      if (String(row[COL_GAS.ESTADO_PAGO - 1]).trim() !== f.estadoPago) return false;
    }
    if (f.revisado !== 'Todos') {
      var rev     = row[COL_GAS.REVISADO - 1];
      var revBool = (rev === true || String(rev).toLowerCase() === 'true');
      if (f.revisado === 'Sí' && !revBool) return false;
      if (f.revisado === 'No' &&  revBool) return false;
    }
    if (f.enviarGes !== 'Todos') {
      var env     = row[COL_GAS.ENVIAR_GESTOR - 1];
      var envBool = (env === true || String(env).toLowerCase() === 'true');
      if (f.enviarGes === 'Sí' && !envBool) return false;
      if (f.enviarGes === 'No' &&  envBool) return false;
    }
    return true;
  });
}

function _pgEscribirResumen_(sh, gastos) {
  var total = 0, base = 0, iva = 0, nRev = 0, nPago = 0, nGes = 0;
  gastos.forEach(function(row) {
    total += parseFloat(row[COL_GAS.TOTAL          - 1]) || 0;
    base  += parseFloat(row[COL_GAS.BASE_IMPONIBLE  - 1]) || 0;
    iva   += parseFloat(row[COL_GAS.IVA_IMPORTE     - 1]) || 0;

    var eProc   = String(row[COL_GAS.ESTADO_PROCESAMIENTO - 1]).trim();
    var ePago   = String(row[COL_GAS.ESTADO_PAGO           - 1]).trim();
    var rev     = row[COL_GAS.REVISADO      - 1];
    var env     = row[COL_GAS.ENVIAR_GESTOR - 1];
    var revBool = (rev === true || String(rev).toLowerCase() === 'true');
    var envBool = (env === true || String(env).toLowerCase() === 'true');

    if (!revBool && eProc !== 'Revisado manualmente') nRev++;
    if (ePago !== 'Pagado') nPago++;
    if (envBool) nGes++;
  });

  sh.getRange(PG_.S_TOTAL[0],  PG_.S_TOTAL[1]).setValue(Math.round(total * 100) / 100);
  sh.getRange(PG_.S_BASE[0],   PG_.S_BASE[1]).setValue(Math.round(base  * 100) / 100);
  sh.getRange(PG_.S_IVA[0],    PG_.S_IVA[1]).setValue(Math.round(iva   * 100) / 100);
  sh.getRange(PG_.S_N[0],      PG_.S_N[1]).setValue(gastos.length);
  sh.getRange(PG_.S_N_REV[0],  PG_.S_N_REV[1]).setValue(nRev);
  sh.getRange(PG_.S_N_PAGO[0], PG_.S_N_PAGO[1]).setValue(nPago);
  sh.getRange(PG_.S_N_GES[0],  PG_.S_N_GES[1]).setValue(nGes);
}

function _pgLimpiarTabla_(sh) {
  var lastRow = sh.getLastRow();
  if (lastRow < PG_.TABLE_START) return;
  var n = lastRow - PG_.TABLE_START + 1;
  var rng = sh.getRange(PG_.TABLE_START, 1, n, PG_N_COLS);
  rng.clearContent();
  rng.clearDataValidations();
  rng.clearFormat();
}

function _pgEscribirTabla_(sh, gastos) {
  var n         = gastos.length;
  var ruleCheck = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  var tz        = Session.getScriptTimeZone();

  // ── Paso 1: preparar datos y escribir en batch ────────────────
  var tableData = [];
  for (var i = 0; i < n; i++) {
    var row     = gastos[i];
    var fGas    = row[COL_GAS.FECHA_GASTO - 1];
    var fechaStr = (fGas instanceof Date && !isNaN(fGas.getTime()))
                   ? Utilities.formatDate(fGas, tz, 'dd/MM/yyyy') : '';
    var rev     = row[COL_GAS.REVISADO      - 1];
    var env     = row[COL_GAS.ENVIAR_GESTOR - 1];
    var revBool = (rev === true || String(rev).toLowerCase() === 'true');
    var envBool = (env === true || String(env).toLowerCase() === 'true');
    var total   = parseFloat(row[COL_GAS.TOTAL       - 1]);
    var iva     = parseFloat(row[COL_GAS.IVA_IMPORTE - 1]);

    tableData.push([
      String(row[COL_GAS.ID_GASTO    - 1]).trim(),
      fechaStr,
      String(row[COL_GAS.TIPO_GASTO  - 1]).trim(),
      String(row[COL_GAS.PROVEEDOR   - 1]).trim(),
      String(row[COL_GAS.CONCEPTO    - 1]).trim(),
      isNaN(total) ? '' : total,
      isNaN(iva)   ? '' : iva,
      String(row[COL_GAS.ESTADO_PAGO              - 1]).trim(),
      String(row[COL_GAS.ESTADO_PROCESAMIENTO      - 1]).trim(),
      revBool,   // Revisado (checkbox)
      envBool,   // Enviar_Gestor (checkbox)
      '',        // Ver foto (se rellena con rich text abajo)
      false,     // Ver OCR (checkbox)
      String(row[COL_GAS.OBSERVACIONES - 1]).trim()
    ]);
  }
  sh.getRange(PG_.TABLE_START, 1, n, PG_N_COLS).setValues(tableData);

  // ── Paso 2: formatos de número ────────────────────────────────
  sh.getRange(PG_.TABLE_START, PG_.TC_TOTAL, n, 1).setNumberFormat('#,##0.00 €');
  sh.getRange(PG_.TABLE_START, PG_.TC_IVA,   n, 1).setNumberFormat('#,##0.00 €');

  // ── Paso 3: validaciones checkbox en batch ────────────────────
  sh.getRange(PG_.TABLE_START, PG_.TC_REVISADO, n, 1).setDataValidation(ruleCheck);
  sh.getRange(PG_.TABLE_START, PG_.TC_ENVIAR,   n, 1).setDataValidation(ruleCheck);
  sh.getRange(PG_.TABLE_START, PG_.TC_VER_OCR,  n, 1).setDataValidation(ruleCheck);

  // ── Paso 4: alineaciones en batch ────────────────────────────
  sh.getRange(PG_.TABLE_START, PG_.TC_TOTAL,    n, 1).setHorizontalAlignment('right');
  sh.getRange(PG_.TABLE_START, PG_.TC_IVA,      n, 1).setHorizontalAlignment('right');
  sh.getRange(PG_.TABLE_START, PG_.TC_REVISADO, n, 1).setHorizontalAlignment('center');
  sh.getRange(PG_.TABLE_START, PG_.TC_ENVIAR,   n, 1).setHorizontalAlignment('center');
  sh.getRange(PG_.TABLE_START, PG_.TC_VER_FOTO, n, 1).setHorizontalAlignment('center');
  sh.getRange(PG_.TABLE_START, PG_.TC_VER_OCR,  n, 1).setHorizontalAlignment('center');

  // ── Paso 5: por fila — rich text (foto), notas, color ────────
  for (var j = 0; j < n; j++) {
    var row2  = gastos[j];
    var fila  = PG_.TABLE_START + j;
    var id    = String(row2[COL_GAS.ID_GASTO              - 1]).trim();
    var foto  = String(row2[COL_GAS.FOTO_URL              - 1]).trim();
    var eProc = String(row2[COL_GAS.ESTADO_PROCESAMIENTO  - 1]).trim();

    // Hyperlink Ver foto
    var fotoCell = sh.getRange(fila, PG_.TC_VER_FOTO);
    if (foto) {
      fotoCell.setRichTextValue(
        SpreadsheetApp.newRichTextValue().setText('Ver foto').setLinkUrl(foto).build()
      ).setFontColor('#1a73e8');
    } else {
      fotoCell.setValue('—').setFontColor('#999999');
    }

    // Notas con ID para identificar la fila en onEdit
    sh.getRange(fila, PG_.TC_REVISADO).setNote(id);
    sh.getRange(fila, PG_.TC_ENVIAR).setNote(id);
    sh.getRange(fila, PG_.TC_VER_OCR).setNote(id);

    // Color de fila según estado
    sh.getRange(fila, 1, 1, PG_N_COLS)
      .setBackground(PG_COLORS[eProc] || '#ffffff');
  }

  SpreadsheetApp.flush();
  Logger.log('  Tabla escrita: ' + n + ' filas (desde fila ' + PG_.TABLE_START + ').');
}

// ─────────────────────────────────────────────────────────────
// onEditPanelGastos_ — llamado desde onPanelEdit en Panel.gs
// ─────────────────────────────────────────────────────────────

/**
 * Gestiona todos los eventos de edición en PANEL_GASTOS:
 *  - Checkbox BTN_ACTUALIZAR → actualiza el panel
 *  - Checkbox Revisado (col J) → actualiza GASTOS
 *  - Checkbox Enviar_Gestor (col K) → actualiza GASTOS
 *  - Checkbox Ver OCR (col M) → muestra diálogo con texto OCR
 */
function onEditPanelGastos_(e) {
  var sh  = e.range.getSheet();
  var row = e.range.getRow();
  var col = e.range.getColumn();
  var val = e.value;

  // Botón Actualizar (B10)
  if (row === PG_.BTN_ACTUALIZAR[0] && col === PG_.BTN_ACTUALIZAR[1] && val === 'TRUE') {
    e.range.setValue(false);
    actualizarPanelGastos();
    return;
  }

  // Solo actúa en filas de datos de la tabla
  if (row < PG_.TABLE_START) return;
  var id = String(sh.getRange(row, PG_.TC_ID).getValue()).trim();
  if (!id || id.indexOf('GAS-') !== 0) {
    // Fallback: leer ID desde la nota de la celda editada
    id = String(e.range.getNote() || '').trim();
    if (!id || id.indexOf('GAS-') !== 0) return;
  }

  // Checkbox Revisado
  if (col === PG_.TC_REVISADO) {
    var gasSheet = getSheet('GASTOS');
    var gasRow   = findRowById_(gasSheet, id);
    if (gasRow) {
      gasSheet.getRange(gasRow, COL_GAS.REVISADO).setValue(val === 'TRUE');
      logEvent('INFO', 'onEditPanelGastos_', id + ' Revisado → ' + val, '');
    }
    return;
  }

  // Checkbox Enviar_Gestor
  if (col === PG_.TC_ENVIAR) {
    var gasSheet2 = getSheet('GASTOS');
    var gasRow2   = findRowById_(gasSheet2, id);
    if (gasRow2) {
      gasSheet2.getRange(gasRow2, COL_GAS.ENVIAR_GESTOR).setValue(val === 'TRUE');
      logEvent('INFO', 'onEditPanelGastos_', id + ' Enviar_Gestor → ' + val, '');
    }
    return;
  }

  // Checkbox Ver OCR
  if (col === PG_.TC_VER_OCR && val === 'TRUE') {
    e.range.setValue(false);
    var gasSheet3 = getSheet('GASTOS');
    var gasRow3   = findRowById_(gasSheet3, id);
    if (!gasRow3) {
      SpreadsheetApp.getUi().alert('ID no encontrado en GASTOS: ' + id);
      return;
    }
    var ocrTexto = String(gasSheet3.getRange(gasRow3, COL_GAS.OCR_TEXTO).getValue() || '').trim();
    var proveedor = String(gasSheet3.getRange(gasRow3, COL_GAS.PROVEEDOR).getValue() || '').trim();
    var confianza = gasSheet3.getRange(gasRow3, COL_GAS.CONFIANZA_OCR).getValue();
    var estadoOCR = String(gasSheet3.getRange(gasRow3, COL_GAS.ESTADO_PROCESAMIENTO).getValue() || '').trim();
    var ui = SpreadsheetApp.getUi();
    if (!ocrTexto) {
      ui.alert(id + ' — Sin texto OCR',
               'No hay texto OCR disponible para este gasto.\nEstado: ' + estadoOCR,
               ui.ButtonSet.OK);
      return;
    }
    var cabecera = id + (proveedor ? ' — ' + proveedor : '') +
                   '\nConfianza OCR: ' + confianza + '  |  Estado: ' + estadoOCR +
                   '\n──────────────────────────────────────\n';
    var cuerpo   = ocrTexto.length > 900
                   ? ocrTexto.substring(0, 900) + '\n[... texto truncado]'
                   : ocrTexto;
    ui.alert('Texto OCR — ' + id, cabecera + cuerpo, ui.ButtonSet.OK);
    return;
  }
}

// ─────────────────────────────────────────────────────────────
// Utilidad de estilo
// ─────────────────────────────────────────────────────────────

function _pgSectionHeader_(sh, row, col, numCols) {
  sh.getRange(row, col, 1, numCols)
    .setBackground('#e8f0fe').setFontColor('#1a73e8')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(row, 26);
}
