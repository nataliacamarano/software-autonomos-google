// ============================================================
// PanelContable.gs — PANEL_CONTABLE / Envío al Gestor
// FASE 17 — Módulo de preparación contable
// ============================================================

// ── Constantes principales ────────────────────────────────────

var PCN_SHEET   = 'PANEL_CONTABLE';
var PCN_NC      = 14;
var PCN_EUR_FMT = '#,##0.00 "€"';

// Alias de hojas fuente
var PCN_SHEET_ALIASES = {
  FACTURAS:     ['FACTURAS'],
  GASTOS:       ['GASTOS'],
  PROFORMAS:    ['PROFORMAS'],
  PRESUPUESTOS: ['PRESUPUESTOS'],
  ENVIOS_LOG:   ['ENVIOS_CONTABLE_LOG']
};

// Paleta de colores (morado como color corporativo del módulo)
var PCN_C = {
  MORADO:      '#6a1b9a',
  MORADO_DARK: '#4a148c',
  MORADO_BG:   '#f3e5f5',
  AZUL_CORP:   '#1a73e8',
  AZUL_LIGHT:  '#e8f0fe',
  VERDE:       '#137333',
  VERDE_BG:    '#e6f4ea',
  VERDE_DARK:  '#0f5132',
  ROJO:        '#c5221f',
  ROJO_BG:     '#fad0cd',
  NARANJA:     '#e37400',
  NARANJA_BG:  '#fef7e0',
  GRIS:        '#5f6368',
  GRIS_BG:     '#f8f9fa',
  GRIS_DARK:   '#343a40',
  SEP_BG:      '#e8eaed',
  BLANCO:      '#ffffff',
  NEGRO:       '#202124'
};

// Layout de filas
var PCN_ROWS = {
  TITULO:      1,
  SUBTITULO:   2,
  SEP_0:       3,
  // Sección 1: Configuración del periodo
  SEC_CONFIG:  4,
  CFG_TIPO:    5,
  CFG_FECHAS:  6,
  CFG_GESTOR:  7,
  CFG_OBS:     8,
  SEP_1:       9,
  // Sección 2: Documentos a incluir
  SEC_DOCS:    10,
  DOCS_R1:     11,
  DOCS_R2:     12,
  DOCS_R3:     13,
  DOCS_R4:     14,
  SEP_2:       15,
  // Sección 3: Resumen del periodo
  SEC_RESUMEN: 16,
  KPI1_LBL:    17,
  KPI1_VAL:    18,
  KPI2_LBL:    19,
  KPI2_VAL:    20,
  SEP_3:       21,
  // Sección 4: Carpeta Drive
  SEC_DRIVE:   22,
  DRIVE_ROW:   23,
  SEP_4:       24,
  // Sección 5: Email al gestor
  SEC_EMAIL:   25,
  EMAIL_ROW:   26,
  SEP_5:       27,
  // Sección 6: Historial de envíos
  SEC_HIST:    28,
  HIST_HDR:    29,
  HIST_START:  30,   // 5 filas: 30-34
  SEP_6:       35,
  // Sección 7: Acciones rápidas
  SEC_ACCESOS: 36,
  ACCS_R1:     37,
  ACCS_R2:     38,
  ACCS_R3:     39,
  ACCS_NOTA:   40
};

// Posiciones de celdas de configuración [fila, col]
// Se leen en FASE 17.2 y siguientes
var PCN_CFG = {
  TIPO_PERIODO:  [5, 4],    // dropdown: Mes/Trimestre/Año/Personalizado
  FECHA_INICIO:  [6, 4],    // editable solo en modo Personalizado
  FECHA_FIN:     [6, 11],
  EMAIL_GESTOR:  [7, 4],
  NOMBRE_GESTOR: [7, 11],
  OBSERVACIONES: [8, 4]
};

// Posiciones de checkboxes de configuración de documentos [fila, col]
var PCN_DOCS = {
  FACTURAS:     [11, 1],
  GASTOS:       [11, 8],
  INF_GASTOS:   [12, 1],
  RESUMEN_IVA:  [12, 8],
  CSV_FAC:      [13, 1],
  CSV_GAS:      [13, 8],
  PROFORMAS:    [14, 1],
  PRESUPUESTOS: [14, 8]
};

// Posiciones de botones de acción (auto-reset) [fila, col]
var PCN_BTNS = {
  PREP_CARPETA: [37, 1],
  GEN_RESUMEN:  [37, 8],
  ENVIAR:       [38, 1],
  ABRIR_DRIVE:  [38, 8],
  DIAGNOSTICAR: [39, 1]
};

// ============================================================
// FUNCIONES PÚBLICAS — Wrappers de menú
// ============================================================

function crearPanelContable() {
  _crearPanelContable_();
}

function actualizarPanelContable() {
  _actualizarPanelContable_();
}

function diagnosticarPanelContable() {
  _diagnosticarPanelContable_();
}

// ============================================================
// _crearPanelContable_ — construye la hoja completa
// ============================================================

function _crearPanelContable_() {
  Logger.log('=== crearPanelContable() INICIO ===');
  var ss = getTargetSS_();
  var sh = ss.getSheetByName(PCN_SHEET);

  if (!sh) {
    sh = ss.insertSheet(PCN_SHEET);
    Logger.log('  + Hoja ' + PCN_SHEET + ' creada.');
  } else {
    // breakApart() ANTES de clear() para evitar conflictos con celdas combinadas
    var maxRows = sh.getMaxRows();
    var maxCols = sh.getMaxColumns();
    sh.getRange(1, 1, maxRows, maxCols).breakApart();
    sh.clear();
    Logger.log('  ✓ Hoja ' + PCN_SHEET + ' limpiada y reconstruida.');
  }

  _pcnSetColumnWidths_(sh);
  _pcnBuildStructure_(sh);
  sh.setFrozenRows(3);

  SpreadsheetApp.flush();
  logEvent('INFO', 'crearPanelContable', 'PANEL_CONTABLE creado.', '');
  Logger.log('=== crearPanelContable() COMPLETADO (40 filas) ===');

  SpreadsheetApp.getUi().alert(
    '✅ PANEL_CONTABLE creado.\n\n' +
    '1. Configura el email del gestor en la fila 7.\n' +
    '2. Elige el periodo en la fila 5.\n' +
    '3. Ejecuta ⚙️ Sistema → 📦 PANEL_CONTABLE — Actualizar para cargar datos.'
  );
}

function _pcnSetColumnWidths_(sh) {
  [[1,15],[2,145],[3,115],[4,15],[5,145],[6,115],
   [7,15],[8,145],[9,115],[10,15],[11,145],[12,115],[13,15],[14,15]]
    .forEach(function(w) { sh.setColumnWidth(w[0], w[1]); });
}

// ── Orquestador ───────────────────────────────────────────────

function _pcnBuildStructure_(sh) {
  _pcnBuildHeader_(sh);
  _pcnBuildConfig_(sh);
  _pcnBuildDocumentos_(sh);
  _pcnBuildResumen_(sh);
  _pcnBuildDrive_(sh);
  _pcnBuildEmail_(sh);
  _pcnBuildHistorial_(sh);
  _pcnBuildAccesos_(sh);
  Logger.log('  ✓ Estructura completa: 40 filas, 14 columnas.');
}

// ── Cabecera (filas 1-3) ──────────────────────────────────────

function _pcnBuildHeader_(sh) {
  var R  = PCN_ROWS;
  var C  = PCN_C;
  var nc = PCN_NC;
  var empresa = _pcnGetEmpresaNombre_();

  // Fila 1: Título
  sh.getRange(R.TITULO, 1, 1, nc).merge()
    .setValue('📦  PANEL_CONTABLE — Envío al gestor')
    .setBackground(C.MORADO_DARK).setFontColor(C.BLANCO)
    .setFontSize(17).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(R.TITULO, 54);

  // Fila 2: Empresa | Estado periodo | Timestamp
  sh.getRange(R.SUBTITULO, 1, 1, 5).merge()
    .setValue('  ' + empresa)
    .setBackground(C.NEGRO).setFontColor(C.GRIS)
    .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(R.SUBTITULO, 6, 1, 4).merge()
    .setValue('⚪ Sin datos')
    .setBackground(C.GRIS_BG).setFontColor(C.GRIS)
    .setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(R.SUBTITULO, 10, 1, 5).merge()
    .setValue('Última actualización: —')
    .setBackground(C.NEGRO).setFontColor(C.GRIS)
    .setFontSize(10).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sh.setRowHeight(R.SUBTITULO, 24);

  // Fila 3: Separador
  _pcnSepRow_(sh, R.SEP_0, nc);
}

// ── Sección 1: Configuración del periodo (filas 4-9) ──────────

function _pcnBuildConfig_(sh) {
  var R  = PCN_ROWS;
  var C  = PCN_C;
  var nc = PCN_NC;

  _pcnSecHdr_(sh, R.SEC_CONFIG, nc,
    '📅  CONFIGURACIÓN DEL PERIODO', C.MORADO_DARK, C.BLANCO);

  // Validación dropdown tipo de periodo
  var dvTipo = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Mes actual', 'Trimestre actual', 'Año actual', 'Personalizado'], true)
    .setAllowInvalid(false).build();

  // Fila 5: Tipo de periodo + periodo detectado
  sh.getRange(R.CFG_TIPO, 1, 1, 3).merge()
    .setValue('  Tipo de periodo')
    .setBackground(C.MORADO_BG).setFontColor(C.NEGRO).setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(R.CFG_TIPO, 4, 1, 4).merge()
    .setValue('Mes actual').setDataValidation(dvTipo)
    .setBackground(C.BLANCO).setFontColor(C.MORADO_DARK)
    .setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(R.CFG_TIPO, 8, 1, 3).merge()
    .setValue('  Periodo detectado:')
    .setBackground(C.MORADO_BG).setFontColor(C.GRIS).setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(R.CFG_TIPO, 11, 1, 4).merge()
    .setValue('—')
    .setBackground(C.BLANCO).setFontColor(C.MORADO)
    .setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(R.CFG_TIPO, 30);

  // Fila 6: Fechas inicio / fin
  // Fila 6: label(1-3) | fechaIni(4-7) | label(8-10) | fechaFin(11-14)
  sh.getRange(R.CFG_FECHAS, 1, 1, 3).merge()
    .setValue('  Fecha inicio')
    .setBackground(C.GRIS_BG).setFontColor(C.NEGRO).setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(R.CFG_FECHAS, 4, 1, 4).merge()   // cols 4-7 (4 columnas, sin solapar col 8)
    .setValue('').setNumberFormat('dd/MM/yyyy')
    .setBackground(C.BLANCO).setFontColor(C.MORADO_DARK).setFontSize(11)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(R.CFG_FECHAS, 8, 1, 3).merge()   // cols 8-10
    .setValue('  Fecha fin')
    .setBackground(C.GRIS_BG).setFontColor(C.NEGRO).setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(R.CFG_FECHAS, 11, 1, 4).merge()  // cols 11-14
    .setValue('').setNumberFormat('dd/MM/yyyy')
    .setBackground(C.BLANCO).setFontColor(C.MORADO_DARK).setFontSize(11)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(R.CFG_FECHAS, 28);

  // Fila 7: Email gestor / Nombre gestor
  var emailVal  = _pcnReadConfig_('email_gestor')   || '';
  var nombreVal = _pcnReadConfig_('nombre_gestor')  || '';

  // Fila 7: label(1-3) | email(4-7) | label(8-10) | nombre(11-14)
  sh.getRange(R.CFG_GESTOR, 1, 1, 3).merge()
    .setValue('  Email gestor')
    .setBackground(C.MORADO_BG).setFontColor(C.NEGRO).setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(R.CFG_GESTOR, 4, 1, 4).merge()   // cols 4-7 (4 columnas, sin solapar col 8)
    .setValue(emailVal)
    .setBackground(C.BLANCO).setFontColor(C.NEGRO).setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(R.CFG_GESTOR, 8, 1, 3).merge()   // cols 8-10
    .setValue('  Nombre gestor')
    .setBackground(C.MORADO_BG).setFontColor(C.NEGRO).setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(R.CFG_GESTOR, 11, 1, 4).merge()  // cols 11-14
    .setValue(nombreVal)
    .setBackground(C.BLANCO).setFontColor(C.NEGRO).setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(R.CFG_GESTOR, 28);

  // Fila 8: Observaciones (ancho completo)
  sh.getRange(R.CFG_OBS, 1, 1, 3).merge()
    .setValue('  Observaciones')
    .setBackground(C.GRIS_BG).setFontColor(C.NEGRO).setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(R.CFG_OBS, 4, 1, 11).merge()
    .setValue('')
    .setBackground(C.BLANCO).setFontColor(C.NEGRO).setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(R.CFG_OBS, 28);

  _pcnSepRow_(sh, R.SEP_1, nc);
}

// ── Sección 2: Documentos a incluir (filas 10-15) ─────────────

function _pcnBuildDocumentos_(sh) {
  var R  = PCN_ROWS;
  var C  = PCN_C;
  var nc = PCN_NC;

  _pcnSecHdr_(sh, R.SEC_DOCS, nc,
    '📄  DOCUMENTOS A INCLUIR', C.AZUL_CORP, C.BLANCO);

  var dvCb = SpreadsheetApp.newDataValidation().requireCheckbox().build();

  var docs = [
    [R.DOCS_R1, '🧾  Facturas emitidas',   '💸  Gastos / tickets'],
    [R.DOCS_R2, '📊  Informe de gastos',    '📋  Resumen IVA'],
    [R.DOCS_R3, '📑  CSV facturas',         '📑  CSV gastos'],
    [R.DOCS_R4, '📃  Proformas',            '📃  Presupuestos']
  ];

  docs.forEach(function(d) {
    var row = d[0];
    var bg  = row % 2 === 0 ? C.BLANCO : C.GRIS_BG;

    // Checkbox + label izquierda
    sh.getRange(row, 1).setDataValidation(dvCb).setValue(true)
      .setBackground(C.AZUL_LIGHT).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(row, 2, 1, 6).merge()
      .setValue(d[1])
      .setBackground(bg).setFontColor(C.NEGRO).setFontSize(10)
      .setHorizontalAlignment('left').setVerticalAlignment('middle');

    // Checkbox + label derecha
    sh.getRange(row, 8).setDataValidation(dvCb).setValue(true)
      .setBackground(C.AZUL_LIGHT).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(row, 9, 1, 6).merge()
      .setValue(d[2])
      .setBackground(bg).setFontColor(C.NEGRO).setFontSize(10)
      .setHorizontalAlignment('left').setVerticalAlignment('middle');

    sh.setRowHeight(row, 26);
  });

  _pcnSepRow_(sh, R.SEP_2, nc);
}

// ── Sección 3: Resumen del periodo (filas 16-21) ──────────────

function _pcnBuildResumen_(sh) {
  var R = PCN_ROWS;
  var C = PCN_C;

  _pcnSecHdr_(sh, R.SEC_RESUMEN, PCN_NC,
    '📊  RESUMEN DEL PERIODO', C.AZUL_CORP, C.BLANCO);

  // Fila 1 KPI: Facturas | Facturado | IVA repercutido | Pendiente
  _pcnKpiRow_(sh, R.KPI1_LBL, R.KPI1_VAL, [
    { lbl: 'Facturas emitidas', val: '—', bg: C.AZUL_LIGHT, fg: C.AZUL_CORP },
    { lbl: 'Total facturado',   val: '—', bg: C.AZUL_LIGHT, fg: C.AZUL_CORP },
    { lbl: 'IVA repercutido',   val: '—', bg: C.VERDE_BG,   fg: C.VERDE     },
    { lbl: 'Pendiente cobro',   val: '—', bg: C.ROJO_BG,    fg: C.ROJO      }
  ], 18, 48);

  // Fila 2 KPI: Nº gastos | Total gastos | IVA soportado | IVA neto
  _pcnKpiRow_(sh, R.KPI2_LBL, R.KPI2_VAL, [
    { lbl: 'Nº gastos',         val: '—', bg: C.NARANJA_BG, fg: C.NARANJA   },
    { lbl: 'Total gastos',      val: '—', bg: C.NARANJA_BG, fg: C.NARANJA   },
    { lbl: 'IVA soportado',     val: '—', bg: C.MORADO_BG,  fg: C.MORADO    },
    { lbl: 'IVA neto',          val: '—', bg: C.MORADO_BG,  fg: C.MORADO    }
  ], 18, 48);

  _pcnSepRow_(sh, R.SEP_3, PCN_NC);
}

// ── Sección 4: Carpeta Drive (filas 22-24) ────────────────────

function _pcnBuildDrive_(sh) {
  var R  = PCN_ROWS;
  var C  = PCN_C;
  var nc = PCN_NC;

  _pcnSecHdr_(sh, R.SEC_DRIVE, nc,
    '📁  CARPETA DRIVE', C.VERDE_DARK, C.BLANCO);

  var bg = C.GRIS_BG;
  sh.getRange(R.DRIVE_ROW, 1, 1, 3).merge()
    .setValue('  Estado:')
    .setBackground(bg).setFontColor(C.NEGRO).setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(R.DRIVE_ROW, 4, 1, 4).merge()
    .setValue('⚪ No preparada')
    .setBackground(bg).setFontColor(C.GRIS).setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(R.DRIVE_ROW, 8, 1, 3).merge()
    .setValue('  URL carpeta:')
    .setBackground(bg).setFontColor(C.NEGRO).setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(R.DRIVE_ROW, 11, 1, 4).merge()
    .setValue('—')
    .setBackground(bg).setFontColor(C.GRIS).setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(R.DRIVE_ROW, 28);

  _pcnSepRow_(sh, R.SEP_4, nc);
}

// ── Sección 5: Email al gestor (filas 25-27) ──────────────────

function _pcnBuildEmail_(sh) {
  var R  = PCN_ROWS;
  var C  = PCN_C;
  var nc = PCN_NC;

  _pcnSecHdr_(sh, R.SEC_EMAIL, nc,
    '✉️  EMAIL AL GESTOR', C.AZUL_CORP, C.BLANCO);

  var bg = C.BLANCO;
  sh.getRange(R.EMAIL_ROW, 1, 1, 3).merge()
    .setValue('  Último envío:')
    .setBackground(bg).setFontColor(C.NEGRO).setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(R.EMAIL_ROW, 4, 1, 4).merge()
    .setValue('—')
    .setBackground(bg).setFontColor(C.GRIS).setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(R.EMAIL_ROW, 8, 1, 3).merge()
    .setValue('  Estado:')
    .setBackground(bg).setFontColor(C.NEGRO).setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(R.EMAIL_ROW, 11, 1, 4).merge()
    .setValue('Sin envíos registrados')
    .setBackground(bg).setFontColor(C.GRIS).setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(R.EMAIL_ROW, 28);

  _pcnSepRow_(sh, R.SEP_5, nc);
}

// ── Sección 6: Historial de envíos (filas 28-35) ─────────────

function _pcnBuildHistorial_(sh) {
  var R  = PCN_ROWS;
  var C  = PCN_C;
  var nc = PCN_NC;

  _pcnSecHdr_(sh, R.SEC_HIST, nc,
    '📋  HISTORIAL DE ENVÍOS', C.GRIS_DARK, C.BLANCO);

  // Cabecera de columnas del historial
  var hdrDefs = [
    [1, 2, 'Fecha envío'],
    [3, 2, 'Periodo'],
    [5, 2, 'Gestor'],
    [7, 4, 'Estado'],
    [11, 4, 'Carpeta Drive']
  ];
  hdrDefs.forEach(function(h) {
    sh.getRange(R.HIST_HDR, h[0], 1, h[1]).merge()
      .setValue(h[2])
      .setBackground('#e9ecef').setFontColor(C.GRIS_DARK)
      .setFontSize(9).setFontWeight('bold')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
  });
  sh.setRowHeight(R.HIST_HDR, 20);

  // 5 filas placeholder
  for (var i = 0; i < 5; i++) {
    var rh  = R.HIST_START + i;
    var bgh = i % 2 === 0 ? C.BLANCO : C.GRIS_BG;
    sh.getRange(rh, 1, 1, nc).merge()
      .setValue('—').setBackground(bgh).setFontColor(C.GRIS)
      .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.setRowHeight(rh, 24);
  }

  _pcnSepRow_(sh, R.SEP_6, nc);
}

// ── Sección 7: Acciones rápidas (filas 36-40) ─────────────────

function _pcnBuildAccesos_(sh) {
  var R  = PCN_ROWS;
  var C  = PCN_C;
  var nc = PCN_NC;

  _pcnSecHdr_(sh, R.SEC_ACCESOS, nc,
    '⚡  ACCIONES RÁPIDAS', C.GRIS_DARK, C.BLANCO);

  var dvAct = SpreadsheetApp.newDataValidation().requireCheckbox().build();

  var filas = [
    [R.ACCS_R1,
     '📁  Preparar carpeta contable',      C.VERDE,      C.BLANCO,
     '📊  Generar paquete contable',       C.AZUL_CORP,  C.BLANCO],
    [R.ACCS_R2,
     '✉️  Enviar email al gestor',         C.MORADO,     C.BLANCO,
     '🔗  Abrir carpeta Drive',            C.VERDE_BG,   C.VERDE_DARK],
    [R.ACCS_R3,
     '🔍  Diagnosticar PANEL_CONTABLE',    C.GRIS_BG,    C.NEGRO,
     '',                                   C.GRIS_BG,    C.GRIS]
  ];

  filas.forEach(function(b) {
    var row = b[0];
    sh.getRange(row, 1).setDataValidation(dvAct).setValue(false)
      .setBackground(b[2]).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(row, 2, 1, 6).merge()
      .setValue(b[1]).setBackground(b[2]).setFontColor(b[3])
      .setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(row, 8).setDataValidation(dvAct).setValue(false)
      .setBackground(b[5]).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(row, 9, 1, 6).merge()
      .setValue(b[4]).setBackground(b[5]).setFontColor(b[6])
      .setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.setRowHeight(row, 40);
  });

  sh.getRange(R.ACCS_NOTA, 1, 1, nc).merge()
    .setValue('ℹ️  Marca el checkbox ☑ de cualquier botón de acción para ejecutarlo.')
    .setBackground(C.GRIS_BG).setFontColor(C.GRIS)
    .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(R.ACCS_NOTA, 22);
}

// ============================================================
// _actualizarPanelContable_ — refresca timestamp, periodo y KPIs
// ============================================================

function _actualizarPanelContable_() {
  Logger.log('=== actualizarPanelContable() ===');
  var ss = getTargetSS_();
  var sh = ss.getSheetByName(PCN_SHEET);

  if (!sh) {
    SpreadsheetApp.getUi().alert(
      '⚠ PANEL_CONTABLE no existe.\n\n' +
      'Ejecuta primero: ⚙️ Sistema → 📦 PANEL_CONTABLE — Crear completo'
    );
    return;
  }

  var tz  = Session.getScriptTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');

  // Timestamp
  sh.getRange(PCN_ROWS.SUBTITULO, 10, 1, 5).merge()
    .setValue('Última actualización: ' + now)
    .setBackground(PCN_C.NEGRO).setFontColor(PCN_C.GRIS)
    .setFontSize(10).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sh.setRowHeight(PCN_ROWS.SUBTITULO, 24);

  // Tipo de periodo y etiqueta
  var tipo    = _pcnLeerTipoPeriodo_(sh);
  var periodo = _pcnPeriodoEtiqueta_(tipo);

  // "Periodo detectado" (fila 5, cols 11-14)
  sh.getRange(PCN_ROWS.CFG_TIPO, 11, 1, 4).merge()
    .setValue(periodo)
    .setBackground(PCN_C.BLANCO).setFontColor(PCN_C.MORADO)
    .setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');

  // Fechas automáticas para periodos fijos
  if (tipo !== 'PERSONALIZADO') {
    var rango   = _pcnCalcularRango_(tipo);
    var fmtIni  = Utilities.formatDate(rango.ini, tz, 'dd/MM/yyyy');
    var fmtFin  = Utilities.formatDate(rango.fin, tz, 'dd/MM/yyyy');
    sh.getRange(PCN_ROWS.CFG_FECHAS, 4, 1, 4).merge()
      .setValue(fmtIni).setFontColor(PCN_C.GRIS);
    sh.getRange(PCN_ROWS.CFG_FECHAS, 11, 1, 4).merge()
      .setValue(fmtFin).setFontColor(PCN_C.GRIS);
  }

  // FASE 17.2 — Calcular y mostrar KPIs reales
  var kpis = _pcnActualizarResumen_(ss, sh);

  // Indicador de estado en cabecera (fila 2, cols 6-9)
  var estado = _pcnEstadoIndicador_(kpis);
  sh.getRange(PCN_ROWS.SUBTITULO, 6, 1, 4).merge()
    .setValue(estado.texto)
    .setBackground(estado.bg).setFontColor(estado.fg)
    .setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  SpreadsheetApp.flush();

  // FASE 18.1: Renderizar selector de documentos
  _pcnRenderizarSelector_(ss, sh);

  logEvent('INFO', 'actualizarPanelContable',
    '✓ Actualizado ' + now + ' — periodo=' + periodo, '');
  Logger.log('=== actualizarPanelContable() OK — periodo=' + periodo + ' ===');
}

// ============================================================
// FASE 17.2 — Lectura de datos reales y KPIs del periodo
// ============================================================

/**
 * Orquestador: lee las hojas fuente, calcula KPIs y actualiza
 * la sección de resumen (filas 17-20).
 * Devuelve el objeto kpis para que el caller actualice el indicador.
 */
function _pcnActualizarResumen_(ss, sh) {
  var rango = _pcnObtenerRango_(sh);

  var fac = _pcnLeerFacturas_(ss, rango);
  var gas = _pcnLeerGastos_(ss, rango);
  var kpis = {
    facCount:    fac.count,
    facTotal:    fac.total,
    ivaRep:      fac.ivaRep,
    pendiente:   fac.pendiente,
    gasCount:    gas.count,
    gasTotal:    gas.total,
    ivaSop:      gas.ivaSop,
    ivaNeto:     fac.ivaRep - gas.ivaSop,
    tieneDatos:  (fac.count > 0 || gas.count > 0),
    hayAnomalias: fac.hayAnomalias || gas.hayAnomalias
  };

  _pcnEscribirKpis_(sh, kpis);
  return kpis;
}

/**
 * Resuelve el rango de fechas efectivo.
 * Para PERSONALIZADO lee las celdas; si están vacías devuelve rango=null (todo).
 * Para el resto usa _pcnCalcularRango_.
 */
function _pcnObtenerRango_(sh) {
  var tipo = _pcnLeerTipoPeriodo_(sh);

  if (tipo !== 'PERSONALIZADO') {
    return _pcnCalcularRango_(tipo);  // {ini, fin}
  }

  // Personalizado: leer celdas de fecha
  var iniVal = sh.getRange(PCN_CFG.FECHA_INICIO[0], PCN_CFG.FECHA_INICIO[1]).getValue();
  var finVal = sh.getRange(PCN_CFG.FECHA_FIN[0],    PCN_CFG.FECHA_FIN[1]   ).getValue();

  if (iniVal instanceof Date && finVal instanceof Date &&
      !isNaN(iniVal.getTime()) && !isNaN(finVal.getTime())) {
    // Normalizar: ini a 00:00:00, fin a 23:59:59
    var ini = new Date(iniVal.getFullYear(), iniVal.getMonth(), iniVal.getDate(), 0, 0, 0);
    var fin = new Date(finVal.getFullYear(), finVal.getMonth(), finVal.getDate(), 23, 59, 59);
    return { ini: ini, fin: fin };
  }

  // Sin fechas → mostrar todo (rango null = sin filtro de fecha)
  return null;
}

/** Comprueba si una fecha cae dentro del rango (rango null = todo). */
function _pcnEnRango_(fecha, rango) {
  if (!(fecha instanceof Date) || isNaN(fecha.getTime())) return false;
  if (!rango) return true;  // sin filtro
  return fecha >= rango.ini && fecha <= rango.fin;
}

// ── Lectura FACTURAS ──────────────────────────────────────────

/**
 * Lee la hoja FACTURAS y calcula KPIs del periodo.
 * Columnas (base 1): 2=Fecha_Emision, 14=IVA_Importe, 17=Total, 18=Estado, 21=Fecha_Pago
 */
function _pcnLeerFacturas_(ss, rango) {
  var result = { count:0, total:0, ivaRep:0, pendiente:0, hayAnomalias:false };
  var sh = ss.getSheetByName('FACTURAS');
  if (!sh || sh.getLastRow() < 2) return result;

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 23).getValues();

  data.forEach(function(row) {
    var idFac  = String(row[0]).trim();
    if (!idFac) return;                    // fila vacía

    var fecha  = row[1];                   // col 2 Fecha_Emision
    var ivaImp = parseFloat(row[13]) || 0; // col 14 IVA_Importe
    var total  = parseFloat(row[16]) || 0; // col 17 Total
    var estado = String(row[17]).trim();   // col 18 Estado

    if (estado === 'Cancelada') return;    // excluir canceladas

    if (!_pcnEnRango_(fecha, rango)) return;

    result.count++;
    result.total  += total;
    result.ivaRep += ivaImp;

    // Pendiente de cobro: facturada pero sin fecha de pago registrada
    var fechaPago = row[20];               // col 21 Fecha_Pago
    var estPago   = String(estado).toLowerCase();
    var sinPago   = !(fechaPago instanceof Date && !isNaN(fechaPago.getTime())) &&
                    estPago !== 'pagada' && estPago !== 'cobrada';
    if (sinPago) result.pendiente += total;

    // Anomalía: total negativo o IVA inconsistente
    if (total < 0 || ivaImp < 0) result.hayAnomalias = true;
  });

  return result;
}

// ── Lectura GASTOS ────────────────────────────────────────────

/**
 * Lee la hoja GASTOS y calcula KPIs del periodo.
 * Columnas según COL_GAS: 4=Fecha_Gasto, 12=IVA_Importe, 13=Total, 15=Estado_Pago
 */
function _pcnLeerGastos_(ss, rango) {
  var result = { count:0, total:0, ivaSop:0, hayAnomalias:false };
  var sh = ss.getSheetByName('GASTOS');
  if (!sh || sh.getLastRow() < 2) return result;

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 23).getValues();

  data.forEach(function(row) {
    var idGas = String(row[0]).trim();
    if (!idGas) return;

    var fecha  = row[3];                   // col 4  Fecha_Gasto
    var ivaImp = parseFloat(row[11]) || 0; // col 12 IVA_Importe
    var total  = parseFloat(row[12]) || 0; // col 13 Total

    if (!_pcnEnRango_(fecha, rango)) return;

    result.count++;
    result.total  += total;
    result.ivaSop += ivaImp;

    if (total < 0 || ivaImp < 0) result.hayAnomalias = true;
  });

  return result;
}

// ── Escritura KPIs en panel ───────────────────────────────────

function _pcnEscribirKpis_(sh, kpis) {
  var R = PCN_ROWS;
  var C = PCN_C;

  // Fila KPI1: Facturas emitidas | Total facturado | IVA repercutido | Pendiente cobro
  _pcnKpiRow_(sh, R.KPI1_LBL, R.KPI1_VAL, [
    { lbl: 'Facturas emitidas',
      val: kpis.facCount > 0 ? String(kpis.facCount) : '—',
      bg:  C.AZUL_LIGHT, fg: C.AZUL_CORP },
    { lbl: 'Total facturado',
      val: kpis.facCount > 0 ? _pcnFmtEur_(kpis.facTotal) : '—',
      bg:  C.AZUL_LIGHT, fg: C.AZUL_CORP },
    { lbl: 'IVA repercutido',
      val: kpis.facCount > 0 ? _pcnFmtEur_(kpis.ivaRep) : '—',
      bg:  C.VERDE_BG,   fg: C.VERDE     },
    { lbl: 'Pendiente cobro',
      val: kpis.facCount > 0 ? _pcnFmtEur_(kpis.pendiente) : '—',
      bg:  kpis.pendiente > 0 ? C.ROJO_BG : C.VERDE_BG,
      fg:  kpis.pendiente > 0 ? C.ROJO    : C.VERDE    }
  ], 18, 48);

  // Fila KPI2: Nº gastos | Total gastos | IVA soportado | IVA neto
  var ivaNetoPos = kpis.ivaNeto >= 0;
  _pcnKpiRow_(sh, R.KPI2_LBL, R.KPI2_VAL, [
    { lbl: 'Nº gastos',
      val: kpis.gasCount > 0 ? String(kpis.gasCount) : '—',
      bg:  C.NARANJA_BG, fg: C.NARANJA },
    { lbl: 'Total gastos',
      val: kpis.gasCount > 0 ? _pcnFmtEur_(kpis.gasTotal) : '—',
      bg:  C.NARANJA_BG, fg: C.NARANJA },
    { lbl: 'IVA soportado',
      val: kpis.gasCount > 0 ? _pcnFmtEur_(kpis.ivaSop) : '—',
      bg:  C.MORADO_BG, fg: C.MORADO },
    { lbl: 'IVA a pagar (neto)',
      val: kpis.tieneDatos ? _pcnFmtEur_(kpis.ivaNeto) : '—',
      bg:  ivaNetoPos ? C.ROJO_BG : C.VERDE_BG,
      fg:  ivaNetoPos ? C.ROJO    : C.VERDE    }
  ], 18, 48);
}

// ── Indicador de estado global ────────────────────────────────

function _pcnEstadoIndicador_(kpis) {
  var C = PCN_C;
  if (!kpis.tieneDatos) {
    return { texto: '⚪ Sin datos en el periodo', bg: C.GRIS_BG, fg: C.GRIS };
  }
  if (kpis.hayAnomalias) {
    return { texto: '🟠 Revisar — hay anomalías', bg: C.NARANJA_BG, fg: C.NARANJA };
  }
  return { texto: '🟢 Sistema OK — ' + kpis.facCount + ' facturas · ' +
                  kpis.gasCount + ' gastos', bg: C.VERDE_BG, fg: C.VERDE_DARK };
}

// ============================================================
// _diagnosticarPanelContable_ — diagnóstico completo FASE 17.2+
// ============================================================

function _diagnosticarPanelContable_() {
  var tz  = Session.getScriptTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  Logger.log('=== diagnosticarPanelContable() ' + now + ' ===');

  var ss    = getTargetSS_();
  var a     = [];
  var warns = [];

  a.push('=== DIAGNÓSTICO PANEL_CONTABLE ===');
  a.push(now);
  a.push('');

  // 1. PANEL_CONTABLE
  var panelSh = ss.getSheetByName(PCN_SHEET);
  if (panelSh) {
    a.push('✅ PANEL_CONTABLE — ' + panelSh.getLastRow() + ' filas usadas');
  } else {
    a.push('❌ PANEL_CONTABLE — no existe (ejecuta "Crear completo")');
    warns.push('PANEL_CONTABLE no existe');
  }

  // 2. Hojas fuente
  a.push('');
  a.push('📋 HOJAS FUENTE:');
  var hojasCheck = [
    ['FACTURAS',     PCN_SHEET_ALIASES.FACTURAS    ],
    ['GASTOS',       PCN_SHEET_ALIASES.GASTOS       ],
    ['PROFORMAS',    PCN_SHEET_ALIASES.PROFORMAS    ],
    ['PRESUPUESTOS', PCN_SHEET_ALIASES.PRESUPUESTOS ],
    ['ENVIOS_LOG',   PCN_SHEET_ALIASES.ENVIOS_LOG   ]
  ];

  hojasCheck.forEach(function(h) {
    var found = null;
    h[1].forEach(function(nombre) {
      if (!found) found = ss.getSheetByName(nombre);
    });
    if (found) {
      var nReg = Math.max(0, found.getLastRow() - 1);
      a.push('  ✅ ' + h[0] + ' — ' + nReg + ' registros');
    } else {
      a.push('  ⚠ ' + h[0] + ' — no encontrada');
      if (h[0] !== 'ENVIOS_LOG') warns.push(h[0] + ' no encontrada');
    }
  });

  // 3. Configuración
  a.push('');
  a.push('⚙️ CONFIGURACIÓN:');

  if (panelSh) {
    var tipo    = _pcnLeerTipoPeriodo_(panelSh);
    var periodo = _pcnPeriodoEtiqueta_(tipo);
    a.push('  Tipo periodo: ' + tipo + ' → ' + periodo);

    var email = String(
      panelSh.getRange(PCN_CFG.EMAIL_GESTOR[0], PCN_CFG.EMAIL_GESTOR[1]).getValue() || ''
    ).trim();
    if (email) {
      a.push('  ✅ Email gestor: ' + email);
    } else {
      a.push('  ⚠ Email gestor: no configurado (fila 7, campo derecho)');
      warns.push('Email gestor no configurado');
    }

    var nombre = String(
      panelSh.getRange(PCN_CFG.NOMBRE_GESTOR[0], PCN_CFG.NOMBRE_GESTOR[1]).getValue() || ''
    ).trim();
    a.push('  Nombre gestor: ' + (nombre || '(no configurado)'));
  } else {
    a.push('  (panel no existe — crear primero)');
  }

  // 4. KPIs del periodo
  if (panelSh) {
    a.push('');
    a.push('📊 KPIs DEL PERIODO:');

    var rango = _pcnObtenerRango_(panelSh);
    if (rango) {
      var fmtIni = Utilities.formatDate(rango.ini, tz, 'dd/MM/yyyy');
      var fmtFin = Utilities.formatDate(rango.fin, tz, 'dd/MM/yyyy');
      a.push('  Rango detectado:    ' + fmtIni + ' — ' + fmtFin);
    } else {
      a.push('  Rango detectado:    Todo el historial (sin filtro de fecha)');
    }

    var fac = _pcnLeerFacturas_(ss, rango);
    var gas = _pcnLeerGastos_(ss, rango);

    a.push('  Facturas incluidas: ' + fac.count);
    a.push('  Total facturas:     ' + _pcnFmtEur_(fac.total));
    a.push('  IVA repercutido:    ' + _pcnFmtEur_(fac.ivaRep));
    a.push('  Pendiente cobro:    ' + _pcnFmtEur_(fac.pendiente));
    a.push('');
    a.push('  Gastos incluidos:   ' + gas.count);
    a.push('  Total gastos:       ' + _pcnFmtEur_(gas.total));
    a.push('  IVA soportado:      ' + _pcnFmtEur_(gas.ivaSop));
    a.push('');
    a.push('  IVA neto (a pagar): ' + _pcnFmtEur_(fac.ivaRep - gas.ivaSop));

    if (fac.hayAnomalias) warns.push('Anomalías en FACTURAS (totales negativos)');
    if (gas.hayAnomalias) warns.push('Anomalías en GASTOS (totales negativos)');
  }

  // 5. Validación de columnas clave
  a.push('');
  a.push('🔬 COLUMNAS CLAVE:');
  var facSh2 = ss.getSheetByName('FACTURAS');
  var gasSh2 = ss.getSheetByName('GASTOS');
  if (facSh2 && facSh2.getLastRow() >= 1) {
    var facCols = facSh2.getLastColumn();
    var marker  = facCols >= 21 ? '✅' : '⚠';
    a.push('  ' + marker + ' FACTURAS: ' + facCols + ' columnas' +
      (facCols < 21 ? ' (esperadas ≥21 para todos los campos)' : ''));
    if (facCols < 21) warns.push('FACTURAS tiene menos columnas de las esperadas (' + facCols + '<21)');
  }
  if (gasSh2 && gasSh2.getLastRow() >= 1) {
    var gasCols = gasSh2.getLastColumn();
    var markerG = gasCols >= 13 ? '✅' : '⚠';
    a.push('  ' + markerG + ' GASTOS: ' + gasCols + ' columnas' +
      (gasCols < 13 ? ' (esperadas ≥13 para todos los campos)' : ''));
    if (gasCols < 13) warns.push('GASTOS tiene menos columnas de las esperadas (' + gasCols + '<13)');
  }

  // 6. Carpeta Drive
  if (panelSh) {
    a.push('');
    a.push('📁 CARPETA DRIVE:');
    var estadoDrive = String(panelSh.getRange(PCN_ROWS.DRIVE_ROW, 4).getValue() || '').trim();
    var urlDrive    = String(panelSh.getRange(PCN_ROWS.DRIVE_ROW, 11).getValue() || '').trim();
    a.push('  Estado: ' + (estadoDrive || '—'));
    a.push('  URL:    ' + (urlDrive    || '—'));
  }

  // 7. Log de envíos
  a.push('');
  a.push('📋 LOG DE ENVÍOS:');
  var logSh = ss.getSheetByName('ENVIOS_CONTABLE_LOG');
  if (logSh) {
    var nEnvios = Math.max(0, logSh.getLastRow() - 1);
    a.push('  ✅ ENVIOS_CONTABLE_LOG — ' + nEnvios + ' envío(s) registrado(s)');
  } else {
    a.push('  ⚪ ENVIOS_CONTABLE_LOG — no existe aún (se crea al enviar el primer email)');
  }

  // 8. Selector y paquete (FASE 18.1)
  a.push('');
  a.push('📦 SELECTOR / PAQUETE:');
  var selRaw  = PropertiesService.getScriptProperties().getProperty(PCN_PROP_SEL);
  var pkgRaw  = PropertiesService.getScriptProperties().getProperty(PCN_PROP_STATS);
  if (selRaw) {
    try {
      var sel = JSON.parse(selRaw);
      a.push('  Selector activo para:  ' + sel.periodo);
      a.push('  Facturas en selector:  ' + sel.facCount);
      a.push('  Gastos en selector:    ' + sel.gasCount);
    } catch (e) { a.push('  Selector: error al leer layout'); }
  } else {
    a.push('  Selector: no renderizado aún (pulsa "Actualizar")');
  }
  if (pkgRaw) {
    try {
      var pkg = JSON.parse(pkgRaw);
      a.push('');
      a.push('  Último paquete:        ' + pkg.periodo + ' (' + pkg.ts + ')');
      a.push('  PDFs facturas:         ' + pkg.facPdfOk + ' copiados / ' +
        (pkg.facSelected || pkg.facTotal || '?') + ' seleccionadas');
      a.push('  Adjuntos gastos:       ' + pkg.gasAdjOk + ' copiados / ' +
        (pkg.gasSelected || pkg.gasTotal || '?') + ' seleccionados');
      if ((pkg.facPdfError || 0) + (pkg.gasAdjError || 0) > 0) {
        warns.push('Archivos no localizados en el último paquete: ' +
          ((pkg.facPdfError || 0) + (pkg.gasAdjError || 0)));
      }
    } catch (e) { a.push('  Paquete: error al leer stats'); }
  } else {
    a.push('  Paquete: no generado aún (pulsa "Generar paquete contable")');
    warns.push('No se ha generado ningún paquete contable');
  }

  // 9. Estado final
  a.push('');
  var resumen = warns.length === 0
    ? '✅ PANEL_CONTABLE listo'
    : '⚠ PANEL_CONTABLE con advertencias — ' + warns.length + ' incidencia(s)';
  a.push(resumen);

  if (warns.length > 0) {
    a.push('');
    warns.forEach(function(w) { a.push('  · ' + w); });
  }

  Logger.log(a.join('\n'));
  logEvent('INFO', 'diagnosticarPanelContable', resumen, warns.slice(0, 5).join(' | '));
  SpreadsheetApp.getUi().alert('🔍 Diagnóstico PANEL_CONTABLE\n\n' + a.join('\n'));
}

// ============================================================
// onEditPanelContable_ — gestiona botones de acción
// ============================================================

function onEditPanelContable_(e) {
  if (!e || !e.range) return;
  if (e.range.getSheet().getName() !== PCN_SHEET) return;
  if (e.value !== 'TRUE') return;

  var row = e.range.getRow();
  var col = e.range.getColumn();

  // Solo reaccionar a botones de acción (PCN_BTNS), no a checkboxes de documentos
  var B   = PCN_BTNS;
  var esBtn = (row === B.PREP_CARPETA[0] && col === B.PREP_CARPETA[1]) ||
              (row === B.GEN_RESUMEN[0]  && col === B.GEN_RESUMEN[1])  ||
              (row === B.ENVIAR[0]       && col === B.ENVIAR[1])       ||
              (row === B.ABRIR_DRIVE[0]  && col === B.ABRIR_DRIVE[1])  ||
              (row === B.DIAGNOSTICAR[0] && col === B.DIAGNOSTICAR[1]);

  if (!esBtn) return;

  e.range.setValue(false);  // reset inmediato
  _pcnEjecutarAccion_(row, col);
}

function _pcnEjecutarAccion_(row, col) {
  var ui = SpreadsheetApp.getUi();
  var B  = PCN_BTNS;

  if (row === B.PREP_CARPETA[0] && col === B.PREP_CARPETA[1]) {
    _pcnPrepararCarpetaContable_();
  } else if (row === B.GEN_RESUMEN[0] && col === B.GEN_RESUMEN[1]) {
    _pcnGenerarPaqueteContable_();
  } else if (row === B.ENVIAR[0] && col === B.ENVIAR[1]) {
    _pcnEnviarEmailGestor_();
  } else if (row === B.ABRIR_DRIVE[0] && col === B.ABRIR_DRIVE[1]) {
    _pcnAbrirCarpetaDrive_();
  } else if (row === B.DIAGNOSTICAR[0] && col === B.DIAGNOSTICAR[1]) {
    _diagnosticarPanelContable_();
  }
}

// ============================================================
// HELPERS
// ============================================================

function _pcnSecHdr_(sh, row, nc, texto, bg, fg) {
  sh.getRange(row, 1, 1, nc).merge()
    .setValue('  ' + texto)
    .setBackground(bg || PCN_C.MORADO_DARK).setFontColor(fg || PCN_C.BLANCO)
    .setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(row, 36);
}

function _pcnSepRow_(sh, row, nc) {
  sh.getRange(row, 1, 1, nc).merge().setValue('').setBackground(PCN_C.SEP_BG);
  sh.setRowHeight(row, 8);
}

// 4 KPI cards en columnas B-C | E-F | H-I | K-L (misma lógica que PANEL_GENERAL)
function _pcnKpiRow_(sh, lblRow, valRow, kpis, valFontSize, valRowHeight) {
  var colStarts = [2, 5, 8, 11];
  var C   = PCN_C;
  var vfs = valFontSize  || 18;
  var vrh = valRowHeight || 48;
  var bStyle = SpreadsheetApp.BorderStyle.SOLID;
  var bColor = '#dadce0';

  [1, 4, 7, 10, 13, 14].forEach(function(c) {
    sh.getRange(lblRow, c).setBackground(C.BLANCO);
    sh.getRange(valRow, c).setBackground(C.BLANCO);
  });

  kpis.forEach(function(kpi, i) {
    var col  = colStarts[i];
    var lblR = sh.getRange(lblRow, col, 1, 2).merge();
    var valR = sh.getRange(valRow, col, 1, 2).merge();

    lblR.setValue(kpi.lbl)
      .setBackground(kpi.bg).setFontColor(C.GRIS)
      .setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('bottom');
    valR.setValue(kpi.val)
      .setBackground(kpi.bg).setFontColor(kpi.fg)
      .setFontSize(vfs).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');

    lblR.setBorder(true, true, false, true, null, null, bColor, bStyle);
    valR.setBorder(false, true, true, true, null, null, bColor, bStyle);
  });

  sh.setRowHeight(lblRow, 20);
  sh.setRowHeight(valRow, vrh);
}

// Lee el tipo de periodo desde la celda de configuración
function _pcnLeerTipoPeriodo_(sh) {
  var val = String(
    sh.getRange(PCN_CFG.TIPO_PERIODO[0], PCN_CFG.TIPO_PERIODO[1]).getValue() || ''
  ).trim();
  var map = {
    'Mes actual':       'MES_ACTUAL',
    'Trimestre actual': 'TRIMESTRE_ACTUAL',
    'Año actual':       'ANIO_ACTUAL',
    'Personalizado':    'PERSONALIZADO'
  };
  return map[val] || 'MES_ACTUAL';
}

// Devuelve la etiqueta legible del periodo
function _pcnPeriodoEtiqueta_(tipo) {
  var d    = new Date();
  var anio = d.getFullYear();
  var mes  = d.getMonth();
  var meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
               'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var trim = Math.floor(mes / 3) + 1;
  switch (tipo) {
    case 'MES_ACTUAL':       return meses[mes] + ' ' + anio;
    case 'TRIMESTRE_ACTUAL': return 'T' + trim + ' ' + anio;
    case 'ANIO_ACTUAL':      return 'Año ' + anio;
    case 'PERSONALIZADO':    return 'Personalizado';
    default:                 return meses[mes] + ' ' + anio;
  }
}

// Calcula fecha inicio y fin según tipo (para autocompletado de fechas)
function _pcnCalcularRango_(tipo) {
  var d    = new Date();
  var anio = d.getFullYear();
  var mes  = d.getMonth();
  var trim = Math.floor(mes / 3);
  var ini, fin;

  switch (tipo) {
    case 'MES_ACTUAL':
      ini = new Date(anio, mes, 1);
      fin = new Date(anio, mes + 1, 0);
      break;
    case 'TRIMESTRE_ACTUAL':
      ini = new Date(anio, trim * 3, 1);
      fin = new Date(anio, trim * 3 + 3, 0);
      break;
    case 'ANIO_ACTUAL':
      ini = new Date(anio, 0, 1);
      fin = new Date(anio, 11, 31);
      break;
    default:
      ini = new Date(anio, mes, 1);
      fin = new Date(anio, mes + 1, 0);
  }
  return { ini: ini, fin: fin };
}

// Lee valor de CONFIG (con fallback silencioso)
function _pcnReadConfig_(key) {
  try { return String(getConfig(key) || ''); }
  catch (_) { return ''; }
}

function _pcnGetEmpresaNombre_() {
  try { return getConfig('empresa_nombre') || 'Gestión'; }
  catch (_) { return 'Gestión'; }
}

function _pcnFmtEur_(val) {
  var n   = parseFloat(val) || 0;
  var sgn = n < 0 ? '-' : '';
  var abs = Math.abs(n).toFixed(2).split('.');
  var int = abs[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return sgn + int + ',' + abs[1] + ' €';
}

// ============================================================
// FASE 17.3 — Carpeta contable en Google Drive
// ============================================================

var PCN_DRIVE_ROOT      = 'Contabilidad';
var PCN_DRIVE_SUBFOLDERS = [
  'Facturas Emitidas',
  'Gastos',
  'IVA',
  'Presupuestos',
  'Proformas',
  'Resumen'
];

/**
 * Wrapper público (menú y test directo).
 */
function prepararCarpetaContable() {
  _pcnPrepararCarpetaContable_();
}

/**
 * Crea (o localiza si ya existe) la estructura de carpetas en Drive:
 *   Contabilidad / AÑO / PERIODO / [Facturas Emitidas, Gastos, IVA, ...]
 * Actualiza las celdas Estado y URL en la sección CARPETA DRIVE del panel.
 */
function _pcnPrepararCarpetaContable_() {
  var ss = getTargetSS_();
  var sh = ss.getSheetByName(PCN_SHEET);
  if (!sh) {
    SpreadsheetApp.getUi().alert(
      '⚠ PANEL_CONTABLE no existe.\nEjecuta primero: ⚙️ Sistema → 📦 PANEL_CONTABLE — Crear completo.'
    );
    return;
  }

  var tipo    = _pcnLeerTipoPeriodo_(sh);
  var anio    = String(new Date().getFullYear());
  var periodo = _pcnPeriodoEtiqueta_(tipo);  // e.g. "Junio 2026" / "T2 2026" / "Año 2026"

  try {
    var rootDrive  = DriveApp.getRootFolder();
    var contFolder = _pcnGetOrCreateFolder_(rootDrive, PCN_DRIVE_ROOT);
    var anioFolder = _pcnGetOrCreateFolder_(contFolder, anio);
    var perFolder  = _pcnGetOrCreateFolder_(anioFolder, periodo);

    PCN_DRIVE_SUBFOLDERS.forEach(function(nombre) {
      _pcnGetOrCreateFolder_(perFolder, nombre);
    });

    var url = perFolder.getUrl();

    // Actualizar sección CARPETA DRIVE en el panel
    var C = PCN_C;
    var R = PCN_ROWS;
    sh.getRange(R.DRIVE_ROW, 4, 1, 4).merge()
      .setValue('✅ Preparada — ' + periodo)
      .setBackground(C.VERDE_BG).setFontColor(C.VERDE_DARK)
      .setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.getRange(R.DRIVE_ROW, 11, 1, 4).merge()
      .setValue(url)
      .setBackground(C.AZUL_LIGHT).setFontColor(C.AZUL_CORP)
      .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');

    SpreadsheetApp.flush();

    // FASE 18.1: Renderizar selector de documentos
    _pcnRenderizarSelector_(ss, sh);

    logEvent('INFO', 'prepararCarpetaContable',
      'Carpeta: ' + PCN_DRIVE_ROOT + '/' + anio + '/' + periodo, url);

    var html = HtmlService.createHtmlOutput(
      '<p style="font-family:sans-serif">✅ <b>Carpeta contable preparada.</b></p>' +
      '<p style="font-family:sans-serif">Ruta: <b>' +
        PCN_DRIVE_ROOT + ' / ' + anio + ' / ' + periodo + '</b></p>' +
      '<p style="font-family:sans-serif">Subcarpetas creadas:<br>' +
        PCN_DRIVE_SUBFOLDERS.join(', ') + '</p>' +
      '<p><a href="' + url + '" target="_blank" ' +
        'style="font-family:sans-serif;font-size:14px">' +
        '📂 Abrir carpeta en Drive</a></p>'
    ).setWidth(440).setHeight(180);
    SpreadsheetApp.getUi().showModalDialog(html, '📁 Carpeta contable preparada');

  } catch (e) {
    Logger.log('ERROR _pcnPrepararCarpetaContable_: ' + e.message);
    logEvent('ERROR', 'prepararCarpetaContable', e.message, '');
    SpreadsheetApp.getUi().alert('❌ Error al preparar la carpeta:\n\n' + e.message);
  }
}

/**
 * Abre un modal con el enlace a la carpeta Drive del periodo actual.
 * Si todavía no está preparada, avisa al usuario.
 */
function _pcnAbrirCarpetaDrive_() {
  var ss = getTargetSS_();
  var sh = ss.getSheetByName(PCN_SHEET);
  if (!sh) {
    SpreadsheetApp.getUi().alert('⚠ PANEL_CONTABLE no existe.');
    return;
  }

  var urlVal = String(sh.getRange(PCN_ROWS.DRIVE_ROW, 11).getValue() || '').trim();

  if (!urlVal || urlVal === '—') {
    SpreadsheetApp.getUi().alert(
      '⚠ La carpeta contable aún no ha sido preparada.\n\n' +
      'Ejecuta primero: ☑ Preparar carpeta contable'
    );
    return;
  }

  var html = HtmlService.createHtmlOutput(
    '<p style="font-family:sans-serif">📂 Carpeta contable del periodo:</p>' +
    '<p><a href="' + urlVal + '" target="_blank" ' +
      'style="font-family:sans-serif;font-size:14px">Abrir en Google Drive</a></p>'
  ).setWidth(360).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, '🔗 Carpeta Drive');
}

/** Obtiene una subcarpeta por nombre o la crea si no existe. */
function _pcnGetOrCreateFolder_(parent, nombre) {
  var iter = parent.getFoldersByName(nombre);
  if (iter.hasNext()) return iter.next();
  return parent.createFolder(nombre);
}
