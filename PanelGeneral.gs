// ============================================================
// PanelGeneral.gs — PANEL_GENERAL / Dashboard Ejecutivo
// FASE 16 — Dashboard principal del sistema
// Mejoras estructurales + FASE 16.2 (Resumen Financiero real)
// ============================================================

// ── Constantes principales ────────────────────────────────────

var PGN_SHEET = 'PANEL_GENERAL';
var PGN_NC    = 14;  // columnas A..N

// Alias de hojas para compatibilidad con formularios Google
var PGN_SHEET_ALIASES = {
  CLIENTES:          ['CLIENTES', 'Alta clientes de formulario 2', 'RESP_CLIENTES'],
  TRABAJOS:          ['TRABAJOS', 'Trabajo de formulario 1', 'RESP_TRABAJOS'],
  OPERARIOS:         ['OPERARIOS', 'Alta empleados de formulario 3', 'RESP_OPERARIOS'],
  FACTURAS:          ['FACTURAS'],
  GASTOS:            ['GASTOS'],
  PROFORMAS:         ['PROFORMAS'],
  PRESUPUESTOS:      ['PRESUPUESTOS'],
  INFORMES:          ['INFORMES'],
  RECORDATORIOS_LOG: ['RECORDATORIOS_LOG'],
  ALERTAS_LOG:       ['ALERTAS_LOG']
};

// ── Paleta de colores ─────────────────────────────────────────
var PGN_C = {
  AZUL_CORP:  '#1a73e8',
  AZUL_DARK:  '#1557b0',
  AZUL_LIGHT: '#e8f0fe',
  VERDE:      '#137333',
  VERDE_BG:   '#e6f4ea',
  VERDE_DARK: '#0f5132',
  ROJO:       '#c5221f',
  ROJO_BG:    '#fad0cd',
  NARANJA:    '#e37400',
  NARANJA_BG: '#fef7e0',
  MORADO:     '#7b1fa2',
  MORADO_BG:  '#f3e5f5',
  GRIS:       '#5f6368',
  GRIS_BG:    '#f8f9fa',
  GRIS_DARK:  '#343a40',
  SEP_BG:     '#e8eaed',
  BLANCO:     '#ffffff',
  NEGRO:      '#202124',
  SLATE:      '#212529'
};

// ── Layout de filas (constantes configurables) ────────────────
// Modificar estos valores si el dashboard crece en fases futuras.
var PGN_ROWS = {
  TITULO:         1,
  SUBTITULO:      2,
  SEP_0:          3,
  // KPI Destacados (prioridad máxima — fuente grande)
  SEC_DESTACADOS: 4,
  KDEST_LBL:      5,
  KDEST_VAL:      6,
  SEP_TOP:        7,
  // Resumen Financiero (8 KPIs)
  SEC_FINANCIERO: 8,
  KPI1_LBL:       9,
  KPI1_VAL:       10,
  KPI2_LBL:       11,
  KPI2_VAL:       12,
  SEP_1:          13,
  // Facturas (izq cols 1-7) + Gastos (der cols 8-14)
  SEC_FAC_GAS:    14,
  FAC_GAS_START:  15,   // 5 KPI rows: 15-19
  FAC_GAS_ULT:    20,   // cabecera "Últimas..."
                        // placeholder rows: 21-25
  SEP_2:          26,
  // Trabajos (izq) + Clientes (der)
  SEC_TRAB_CLI:   27,
  TRAB_CLI_START: 28,   // 4 KPI rows: 28-31
  TRAB_CLI_ULT:   32,   // cabecera "Últimos..."
                        // placeholder rows: 33-37
  SEP_3:          38,
  // Presupuestos / Proformas
  SEC_PRESP:      39,
  PRESP_START:    40,   // 4 KPI rows: 40-43
  PRESP_ULT:      44,   // cabecera "Últimos..."
                        // placeholder rows: 45-47
  SEP_4:          48,
  // Alertas
  SEC_ALERTAS:    49,
  ALERTAS_START:  50,   // 8 alert rows: 50-57
  SEP_5:          58,
  // Accesos rápidos
  SEC_ACCESOS:    59,
  ACCESOS_START:  60,   // 4 button rows: 60-63
  ACCESOS_NOTA:   64
};

// ── Posiciones de checkboxes (onEdit handler) ─────────────────
// Formato: [fila, columna]
var PGN_BTNS = {
  ACTUALIZAR:       [60, 1],
  IR_PANEL_CLIENTE: [60, 8],
  IR_PANEL_GASTOS:  [61, 1],
  IR_PANEL_FAC:     [61, 8],
  INF_GASTOS:       [62, 1],
  IVA_TRIMESTRAL:   [62, 8],
  VER_RECORD:       [63, 1],
  VER_ALERTAS:      [63, 8]
};

// ── Formato de moneda ─────────────────────────────────────────
var PGN_EUR_FMT = '#,##0.00 "€"';   // #,##0.00 "€"

// ============================================================
// FUNCIONES PÚBLICAS — Wrappers de menú
// ============================================================

function crearPanelGeneral() {
  _crearPanelGeneral_();
}

function actualizarPanelGeneral() {
  _actualizarPanelGeneral_();
}

function diagnosticarPanelGeneral() {
  _diagnosticarPanelGeneral_();
}

// ============================================================
// pgObtenerHoja_  —  helper público de resolución de hojas
// Devuelve la primera hoja encontrada entre los nombres dados.
// Útil para compatibilidad con formularios Google que renombran hojas.
// ============================================================

function pgObtenerHoja_(nombres) {
  var ss = getTargetSS_();
  for (var i = 0; i < nombres.length; i++) {
    var sh = ss.getSheetByName(nombres[i]);
    if (sh) return sh;
  }
  return null;
}

// ============================================================
// _crearPanelGeneral_  —  construye la hoja completa
// ============================================================

function _crearPanelGeneral_() {
  Logger.log('=== crearPanelGeneral() INICIO ===');
  var ss = getTargetSS_();
  var sh = ss.getSheetByName(PGN_SHEET);

  if (!sh) {
    sh = ss.insertSheet(PGN_SHEET, 0);
    Logger.log('  + Hoja ' + PGN_SHEET + ' creada en posición 0.');
  } else {
    sh.clear();
    var maxCols = Math.min(sh.getMaxColumns(), PGN_NC + 4);
    var maxRows = sh.getMaxRows();
    sh.getRange(1, 1, maxRows, maxCols).breakApart();
    sh.getRange(1, 1, maxRows, maxCols).clearDataValidations();
    Logger.log('  ✓ Hoja ' + PGN_SHEET + ' limpiada y reconstruida.');
  }

  _pgnSetColumnWidths_(sh);
  _pgnBuildStructure_(sh);
  sh.setFrozenRows(3);

  SpreadsheetApp.flush();
  logEvent('INFO', 'crearPanelGeneral', 'PANEL_GENERAL creado.', '');
  Logger.log('=== crearPanelGeneral() COMPLETADO (64 filas) ===');

  SpreadsheetApp.getUi().alert(
    '✅ PANEL_GENERAL creado.\n\n' +
    'Ejecuta ⚙️ Sistema → 🏠 PANEL_GENERAL — Actualizar datos\n' +
    'para cargar los datos financieros reales.'
  );
}

function _pgnSetColumnWidths_(sh) {
  [[1,15],[2,145],[3,115],[4,15],[5,145],[6,115],
   [7,15],[8,145],[9,115],[10,15],[11,145],[12,115],[13,15],[14,15]]
    .forEach(function(w) { sh.setColumnWidth(w[0], w[1]); });
}

// ── Orquestador de secciones ──────────────────────────────────

function _pgnBuildStructure_(sh) {
  _pgnBuildHeader_(sh);
  _pgnBuildKpiDestacados_(sh);
  _pgnBuildResumenFinanciero_(sh);
  _pgnBuildFacturasGastos_(sh);
  _pgnBuildTrabajosClientes_(sh);
  _pgnBuildPresupuestos_(sh);
  _pgnBuildAlertas_(sh);
  _pgnBuildAccesos_(sh);
  Logger.log('  ✓ Estructura completa: 64 filas, 14 columnas.');
}

// ── Cabecera (filas 1-3) ──────────────────────────────────────

function _pgnBuildHeader_(sh) {
  var nc = PGN_NC;
  var C  = PGN_C;
  var R  = PGN_ROWS;
  var empresa = _pgnGetEmpresaNombre_();

  // Fila 1: Título principal
  sh.getRange(R.TITULO, 1, 1, nc).merge()
    .setValue('🏠  PANEL_GENERAL — Dashboard Ejecutivo')
    .setBackground(C.AZUL_CORP).setFontColor(C.BLANCO)
    .setFontSize(17).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(R.TITULO, 54);

  // Fila 2: Empresa | Estado sistema | Última actualización
  sh.getRange(R.SUBTITULO, 1, 1, 5).merge()
    .setValue('  ' + empresa)
    .setBackground(C.NEGRO).setFontColor(C.GRIS)
    .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(R.SUBTITULO, 6, 1, 4).merge()
    .setValue('🟢 Sistema OK')
    .setBackground('#e6f4ea').setFontColor(C.VERDE)
    .setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(R.SUBTITULO, 10, 1, 5).merge()
    .setValue('Última actualización: —')
    .setBackground(C.NEGRO).setFontColor(C.GRIS)
    .setFontSize(10).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sh.setRowHeight(R.SUBTITULO, 24);

  // Fila 3: Separador
  _pgnSepRow_(sh, R.SEP_0, nc);
}

// ── KPI Destacados (filas 4-7) ────────────────────────────────
// Los 4 indicadores más importantes — fuente 26px

function _pgnBuildKpiDestacados_(sh) {
  var R = PGN_ROWS;
  var C = PGN_C;

  _pgnSecHdr_(sh, R.SEC_DESTACADOS, PGN_NC,
    '🎯  INDICADORES CLAVE', C.SLATE, C.BLANCO);

  _pgnKpiRow_(sh, R.KDEST_LBL, R.KDEST_VAL, [
    { lbl: 'Facturado este mes',   val: '—', bg: C.AZUL_LIGHT, fg: C.AZUL_CORP },
    { lbl: 'Pendiente de cobro',   val: '—', bg: C.ROJO_BG,    fg: C.ROJO      },
    { lbl: 'Trabajos pendientes',  val: '—', bg: C.NARANJA_BG, fg: C.NARANJA   },
    { lbl: 'Alertas activas',      val: '—', bg: '#fce8b2',    fg: C.ROJO      }
  ], 28, 60);

  _pgnSepRow_(sh, R.SEP_TOP, PGN_NC);
}

// ── Resumen Financiero (filas 8-13) ──────────────────────────

function _pgnBuildResumenFinanciero_(sh) {
  var R = PGN_ROWS;
  var C = PGN_C;

  _pgnSecHdr_(sh, R.SEC_FINANCIERO, PGN_NC,
    '💰  RESUMEN FINANCIERO — ' + _pgnMesActual_(),
    C.AZUL_CORP, C.BLANCO);

  _pgnKpiRow_(sh, R.KPI1_LBL, R.KPI1_VAL, [
    { lbl: 'Facturado este mes',   val: '—', bg: C.AZUL_LIGHT, fg: C.AZUL_CORP },
    { lbl: 'Cobrado este mes',     val: '—', bg: C.VERDE_BG,   fg: C.VERDE     },
    { lbl: 'Pendiente de cobro',   val: '—', bg: C.ROJO_BG,    fg: C.ROJO      },
    { lbl: 'Gastos este mes',      val: '—', bg: C.NARANJA_BG, fg: C.NARANJA   }
  ], 18, 48);

  _pgnKpiRow_(sh, R.KPI2_LBL, R.KPI2_VAL, [
    { lbl: 'Beneficio estimado',   val: '—', bg: C.VERDE_BG,   fg: C.VERDE     },
    { lbl: 'IVA repercutido',      val: '—', bg: C.AZUL_LIGHT, fg: C.AZUL_CORP },
    { lbl: 'IVA soportado',        val: '—', bg: C.AZUL_LIGHT, fg: C.AZUL_CORP },
    { lbl: 'IVA neto',             val: '—', bg: C.MORADO_BG,  fg: C.MORADO    }
  ], 18, 48);

  _pgnSepRow_(sh, R.SEP_1, PGN_NC);
}

// ── Facturas (izq) + Gastos (der) (filas 14-26) ──────────────

function _pgnBuildFacturasGastos_(sh) {
  var R  = PGN_ROWS;
  var C  = PGN_C;
  var nc = PGN_NC;

  _pgnSplitHdr_(sh, R.SEC_FAC_GAS, nc,
    '🧯  FACTURAS', C.AZUL_DARK, C.BLANCO,
    '💸  GASTOS',   C.NARANJA,   C.BLANCO);

  var kpiFac = [
    ['Emitidas este mes',   '—', C.AZUL_CORP],
    ['Pendientes de cobro', '—', C.ROJO     ],
    ['Pagadas',             '—', C.VERDE    ],
    ['Vencidas',            '—', C.ROJO     ],
    ['Total pendiente',     '—', C.ROJO     ]
  ];
  var kpiGas = [
    ['Total gastos mes',    '—', C.NARANJA  ],
    ['Nº gastos',      '—', C.AZUL_CORP],
    ['Pendientes de pago',  '—', C.ROJO     ],
    ['Sin revisar',         '—', C.NARANJA  ],
    ['OCR fallido',         '—', C.ROJO     ]
  ];

  for (var i = 0; i < 5; i++) {
    var ri  = R.FAC_GAS_START + i;
    var bgi = i % 2 === 0 ? C.BLANCO : C.GRIS_BG;
    _pgnSplitKpi_(sh, ri, bgi, kpiFac[i], kpiGas[i]);
  }

  sh.getRange(R.FAC_GAS_ULT, 1, 1, 7).merge()
    .setValue('  Últimas 5 facturas emitidas')
    .setBackground(C.AZUL_LIGHT).setFontColor(C.AZUL_CORP)
    .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('left');
  sh.getRange(R.FAC_GAS_ULT, 8, 1, 7).merge()
    .setValue('  Últimos 5 gastos registrados')
    .setBackground(C.NARANJA_BG).setFontColor(C.NARANJA)
    .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('left');
  sh.setRowHeight(R.FAC_GAS_ULT, 22);

  for (var j = 0; j < 5; j++) {
    var rj  = R.FAC_GAS_ULT + 1 + j;
    var bgj = j % 2 === 0 ? C.BLANCO : C.GRIS_BG;
    sh.getRange(rj, 1, 1, 7).merge()
      .setValue('—').setBackground(bgj).setFontColor(C.GRIS)
      .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.getRange(rj, 8, 1, 7).merge()
      .setValue('—').setBackground(bgj).setFontColor(C.GRIS)
      .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.setRowHeight(rj, 24);
  }

  _pgnSepRow_(sh, R.SEP_2, nc);
}

// ── Trabajos (izq) + Clientes (der) (filas 27-38) ────────────

function _pgnBuildTrabajosClientes_(sh) {
  var R  = PGN_ROWS;
  var C  = PGN_C;
  var nc = PGN_NC;

  _pgnSplitHdr_(sh, R.SEC_TRAB_CLI, nc,
    '🔧  TRABAJOS', C.VERDE_DARK, C.BLANCO,
    '👥  CLIENTES', C.GRIS_DARK,  C.BLANCO);

  var kpiTrab = [
    ['Realizados este mes',  '—', C.VERDE_DARK],
    ['Pendientes',           '—', C.NARANJA   ],
    ['Completados',          '—', C.VERDE     ],
    ['Horas trabajadas mes', '—', C.AZUL_CORP ]
  ];
  var kpiCli = [
    ['Clientes activos',        '—', C.GRIS_DARK],
    ['Con facturas pendientes', '—', C.ROJO     ],
    ['Sin actividad (60 días)', '—', C.GRIS],
    ['Facturación total',  '—', C.VERDE    ]
  ];

  for (var k = 0; k < 4; k++) {
    var rk  = R.TRAB_CLI_START + k;
    var bgk = k % 2 === 0 ? C.BLANCO : C.GRIS_BG;
    _pgnSplitKpi_(sh, rk, bgk, kpiTrab[k], kpiCli[k]);
  }

  sh.getRange(R.TRAB_CLI_ULT, 1, 1, 7).merge()
    .setValue('  Últimos 5 trabajos')
    .setBackground('#d1e7dd').setFontColor(C.VERDE_DARK)
    .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('left');
  sh.getRange(R.TRAB_CLI_ULT, 8, 1, 7).merge()
    .setValue('  Top 5 clientes por facturación')
    .setBackground('#e9ecef').setFontColor(C.GRIS_DARK)
    .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('left');
  sh.setRowHeight(R.TRAB_CLI_ULT, 22);

  for (var t = 0; t < 5; t++) {
    var rt  = R.TRAB_CLI_ULT + 1 + t;
    var bgt = t % 2 === 0 ? C.BLANCO : C.GRIS_BG;
    sh.getRange(rt, 1, 1, 7).merge()
      .setValue('—').setBackground(bgt).setFontColor(C.GRIS)
      .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.getRange(rt, 8, 1, 7).merge()
      .setValue('—').setBackground(bgt).setFontColor(C.GRIS)
      .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.setRowHeight(rt, 24);
  }

  _pgnSepRow_(sh, R.SEP_3, nc);
}

// ── Presupuestos / Proformas (filas 39-48) ───────────────────

function _pgnBuildPresupuestos_(sh) {
  var R  = PGN_ROWS;
  var C  = PGN_C;
  var nc = PGN_NC;

  _pgnSecHdr_(sh, R.SEC_PRESP, nc,
    '📋  PRESUPUESTOS / PROFORMAS', C.MORADO, C.BLANCO);

  var kpiPresp = [
    ['Presupuestos pendientes',    '—', C.NARANJA   ],
    ['Presupuestos aceptados',     '—', C.VERDE     ],
    ['Presupuestos rechazados',    '—', C.ROJO      ],
    ['Proformas pendientes de pago','—',C.AZUL_CORP ]
  ];

  for (var p = 0; p < 4; p++) {
    var rp  = R.PRESP_START + p;
    var bgp = p % 2 === 0 ? C.BLANCO : C.GRIS_BG;
    sh.getRange(rp, 1, 1, 8).merge()
      .setValue('  ' + kpiPresp[p][0])
      .setBackground(bgp).setFontColor(C.NEGRO).setFontSize(10)
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.getRange(rp, 9, 1, 6).merge()
      .setValue(kpiPresp[p][1])
      .setBackground(bgp).setFontColor(kpiPresp[p][2])
      .setFontSize(14).setFontWeight('bold')
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
    sh.setRowHeight(rp, 28);
  }

  sh.getRange(R.PRESP_ULT, 1, 1, nc).merge()
    .setValue('  Últimos presupuestos / proformas emitidos')
    .setBackground(C.MORADO_BG).setFontColor(C.MORADO)
    .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('left');
  sh.setRowHeight(R.PRESP_ULT, 22);

  for (var q = 0; q < 3; q++) {
    var rq  = R.PRESP_ULT + 1 + q;
    var bgq = q % 2 === 0 ? C.BLANCO : C.GRIS_BG;
    sh.getRange(rq, 1, 1, nc).merge()
      .setValue('—').setBackground(bgq).setFontColor(C.GRIS)
      .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.setRowHeight(rq, 24);
  }

  _pgnSepRow_(sh, R.SEP_4, nc);
}

// ── Alertas con semáforos (filas 49-58) ──────────────────────
// Estructura: col 1 = tira de color | cols 2-10 = etiqueta | cols 11-14 = semáforo + valor

function _pgnBuildAlertas_(sh) {
  var R  = PGN_ROWS;
  var C  = PGN_C;
  var nc = PGN_NC;

  _pgnSecHdr_(sh, R.SEC_ALERTAS, nc,
    '🚨  ALERTAS ACTIVAS', C.ROJO, C.BLANCO);

  var defs = [
    { lbl: 'Facturas vencidas',           val: '—', sem: '🔴', sev: 'r' },
    { lbl: 'Facturas próximas a vencer', val: '—', sem: '🟠', sev: 'n' },
    { lbl: 'Gastos sin revisar',          val: '—', sem: '🟠', sev: 'n' },
    { lbl: 'OCR fallidos',                val: '—', sem: '🔴', sev: 'r' },
    { lbl: 'Gastos sin ticket / foto',    val: '—', sem: '🟠', sev: 'n' },
    { lbl: 'Presupuestos sin respuesta',  val: '—', sem: '🟠', sev: 'm' },
    { lbl: 'Recordatorios enviados',      val: 'No disponible', sem: '⚪', sev: 'g' },
    { lbl: 'Alertas automáticas',    val: 'No disponible', sem: '⚪', sev: 'g' }
  ];

  // Colores por severidad: r=rojo, n=naranja, m=morado, g=gris
  var sevStrip = { r: C.ROJO_BG,    n: C.NARANJA_BG, m: C.MORADO_BG, g: C.GRIS_BG };
  var sevFg    = { r: C.ROJO,       n: C.NARANJA,    m: C.MORADO,    g: C.GRIS    };

  for (var a = 0; a < defs.length; a++) {
    var ra   = R.ALERTAS_START + a;
    var def  = defs[a];
    var bgRow = a % 2 === 0 ? '#fff8f8' : C.BLANCO;

    // Col 1: tira de color (indicador visual de severidad)
    sh.getRange(ra, 1).setBackground(sevStrip[def.sev]);

    // Cols 2-10: etiqueta
    sh.getRange(ra, 2, 1, 9).merge()
      .setValue('  ' + def.lbl)
      .setBackground(bgRow).setFontColor(C.NEGRO).setFontSize(10)
      .setHorizontalAlignment('left').setVerticalAlignment('middle');

    // Cols 11-14: semáforo + valor
    sh.getRange(ra, 11, 1, 4).merge()
      .setValue(def.sem + '  ' + def.val)
      .setBackground(bgRow).setFontColor(sevFg[def.sev])
      .setFontSize(12).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');

    sh.setRowHeight(ra, 28);
  }

  _pgnSepRow_(sh, R.SEP_5, nc);
}

// ── Accesos rápidos (filas 59-64) ────────────────────────────
// Estructura: col 1 = checkbox | cols 2-7 = label izq
//             col 8 = checkbox | cols 9-14 = label der

function _pgnBuildAccesos_(sh) {
  var R  = PGN_ROWS;
  var C  = PGN_C;
  var nc = PGN_NC;

  _pgnSecHdr_(sh, R.SEC_ACCESOS, nc,
    '⚡  ACCESOS RÁPIDOS', C.GRIS_DARK, C.BLANCO);

  var dv = SpreadsheetApp.newDataValidation().requireCheckbox().build();

  var filas = [
    [60,
     '🔄  Actualizar PANEL_GENERAL', C.AZUL_CORP, C.BLANCO,
     '👥  Ir a PANEL_CLIENTE',       C.GRIS_DARK, C.BLANCO],
    [61,
     '💸  Ir a PANEL_GASTOS',        C.NARANJA_BG, C.NARANJA,
     '🧯  Ir a PANEL_FACTURAS',      C.AZUL_LIGHT, C.AZUL_CORP],
    [62,
     '📊  Generar informe de gastos', C.NARANJA_BG, C.NARANJA,
     '📋  Generar IVA trimestral',   C.MORADO_BG,  C.MORADO],
    [63,
     '⏰  Ver recordatorios',             C.GRIS_BG,   C.NEGRO,
     '🚨  Ver alertas',              '#fce8b2',   C.ROJO]
  ];

  filas.forEach(function(b) {
    var row = b[0];
    // Checkbox izquierdo + label izquierdo
    sh.getRange(row, 1).setDataValidation(dv).setValue(false)
      .setBackground(b[2]).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(row, 2, 1, 6).merge()
      .setValue(b[1]).setBackground(b[2]).setFontColor(b[3])
      .setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    // Checkbox derecho + label derecho
    sh.getRange(row, 8).setDataValidation(dv).setValue(false)
      .setBackground(b[5]).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(row, 9, 1, 6).merge()
      .setValue(b[4]).setBackground(b[5]).setFontColor(b[6])
      .setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.setRowHeight(row, 40);
  });

  sh.getRange(R.ACCESOS_NOTA, 1, 1, nc).merge()
    .setValue('ℹ️  Marca el checkbox ☑ de cualquier botón para ejecutar la acción.')
    .setBackground(C.GRIS_BG).setFontColor(C.GRIS)
    .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(R.ACCESOS_NOTA, 22);
}

// ============================================================
// _actualizarPanelGeneral_  —  actualiza timestamp + datos reales
// ============================================================

function _actualizarPanelGeneral_() {
  Logger.log('=== actualizarPanelGeneral() ===');
  var ss = getTargetSS_();
  var sh = ss.getSheetByName(PGN_SHEET);

  if (!sh) {
    SpreadsheetApp.getUi().alert(
      '⚠ PANEL_GENERAL no existe.\n\n' +
      'Ejecuta primero: ⚙️ Sistema → 🏠 PANEL_GENERAL — Crear completo'
    );
    return;
  }

  var tz  = Session.getScriptTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');

  // Actualizar timestamp (fila 2, segmento derecho cols 10-14)
  sh.getRange(PGN_ROWS.SUBTITULO, 10, 1, 5).merge()
    .setValue('Última actualización: ' + now)
    .setBackground(PGN_C.NEGRO).setFontColor(PGN_C.GRIS)
    .setFontSize(10).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sh.setRowHeight(PGN_ROWS.SUBTITULO, 24);

  // Estado del sistema: se actualizará con valor real en FASE 16.8 (ALERTAS)
  _pgnSetStatus_(sh, 'ok');

  // Actualizar mes en cabecera RESUMEN FINANCIERO
  sh.getRange(PGN_ROWS.SEC_FINANCIERO, 1, 1, PGN_NC).merge()
    .setValue('💰  RESUMEN FINANCIERO — ' + _pgnMesActual_())
    .setBackground(PGN_C.AZUL_CORP).setFontColor(PGN_C.BLANCO)
    .setFontSize(12).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');

  // ── FASE 16.2: Calcular y volcar datos financieros ────────
  var fin = _pgnCalcResumenFinanciero_();
  _pgnVolcarResumenFinanciero_(sh, fin);

  // ── FASE 16.3: Bloque FACTURAS con datos reales ───────────
  var datFac = _pgnCalcFacturas_();
  _pgnActualizarBloqueFacturas_(sh, datFac);

  // ── FASE 16.4: Bloque GASTOS con datos reales ─────────────
  var datGas = _pgnCalcGastos_();
  _pgnActualizarBloqueGastos_(sh, datGas);

  // ── FASE 16.5: Bloque TRABAJOS con datos reales ───────────
  var datTrab = _pgnCalcTrabajos_();
  _pgnActualizarBloqueTrabajos_(sh, datTrab);

  // ── FASE 16.6: Bloque CLIENTES con datos reales ───────────
  var datCli = _pgnCalcClientes_();
  _pgnActualizarBloqueClientes_(sh, datCli);

  // ── FASE 16.7: Bloque PRESUPUESTOS/PROFORMAS ──────────────
  var datPresp = _pgnCalcPresupuestos_();
  _pgnActualizarBloquePresupuestos_(sh, datPresp);

  // ── FASE 16.8: Bloque ALERTAS (usa datos ya calculados) ───
  var datAler = _pgnCalcAlertas_(datFac, datGas, datTrab, datPresp);
  _pgnActualizarBloqueAlertas_(sh, datAler);

  SpreadsheetApp.flush();
  logEvent('INFO', 'actualizarPanelGeneral',
    '✓ Actualizado ' + now,
    fin.logs.concat(datFac.logs).concat(datGas.logs)
            .concat(datTrab.logs).concat(datCli.logs)
            .concat(datPresp.logs).concat(datAler.logs).join(' | '));
  Logger.log('=== actualizarPanelGeneral() OK ===');
}

// ── Volcar datos financieros en las celdas KPI ───────────────

function _pgnVolcarResumenFinanciero_(sh, fin) {
  var R  = PGN_ROWS;
  var C  = PGN_C;
  var cs = [2, 5, 8, 11];  // columnas de inicio de cada tarjeta KPI
  var F  = PGN_EUR_FMT;

  // KPI1: Facturado | Cobrado | Pendiente | Gastos
  var kpi1 = [
    { v: fin.facturadoMes,   fg: C.AZUL_CORP },
    { v: fin.cobradoMes,     fg: C.VERDE     },
    { v: fin.pendienteCobro, fg: fin.pendienteCobro > 0 ? C.ROJO : C.VERDE },
    { v: fin.gastosMes,      fg: C.NARANJA   }
  ];
  kpi1.forEach(function(kpi, i) {
    sh.getRange(R.KPI1_VAL, cs[i], 1, 2).merge()
      .setValue(kpi.v).setNumberFormat(F).setFontColor(kpi.fg)
      .setFontSize(18).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });

  // KPI2: Beneficio | IVA Rep | IVA Sop | IVA Neto
  var kpi2 = [
    { v: fin.beneficioMes,    fg: fin.beneficioMes  >= 0 ? C.VERDE     : C.ROJO   },
    { v: fin.ivaRepercutido,  fg: C.AZUL_CORP },
    { v: fin.ivaSoportado,    fg: C.AZUL_CORP },
    { v: fin.ivaNeto,         fg: fin.ivaNeto >= 0 ? C.VERDE : C.ROJO }
  ];
  kpi2.forEach(function(kpi, i) {
    sh.getRange(R.KPI2_VAL, cs[i], 1, 2).merge()
      .setValue(kpi.v).setNumberFormat(F).setFontColor(kpi.fg)
      .setFontSize(18).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });

  // KDEST: Facturado | Pendiente | Trabajos* | Alertas*
  sh.getRange(R.KDEST_VAL, cs[0], 1, 2).merge()
    .setValue(fin.facturadoMes).setNumberFormat(F)
    .setFontColor(C.AZUL_CORP).setFontSize(26).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(R.KDEST_VAL, cs[1], 1, 2).merge()
    .setValue(fin.pendienteCobro).setNumberFormat(F)
    .setFontColor(fin.pendienteCobro > 0 ? C.ROJO : C.VERDE)
    .setFontSize(26).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  // Trabajos pendientes y Alertas activas: se actualizan en FASE 16.5 y 16.8

  Logger.log('  ✓ KPIs financieros volcados al panel.');
}

// ── Cálculo de datos financieros (FASE 16.2) ─────────────────

function _pgnCalcResumenFinanciero_() {
  Logger.log('  _pgnCalcResumenFinanciero_()');

  var result = {
    facturadoMes:   0,
    cobradoMes:     0,
    pendienteCobro: 0,
    gastosMes:      0,
    beneficioMes:   0,
    ivaRepercutido: 0,
    ivaSoportado:   0,
    ivaNeto:        0,
    logs: []
  };

  // ── FACTURAS ─────────────────────────────────────────────
  var facSh = pgObtenerHoja_(PGN_SHEET_ALIASES.FACTURAS);
  if (!facSh) {
    result.logs.push('⚠ FACTURAS no encontrada — mostrando 0');
    Logger.log('    ⚠ FACTURAS: hoja no encontrada');
  } else {
    var facData = facSh.getDataRange().getValues();
    if (facData.length < 2) {
      result.logs.push('ℹ FACTURAS vacía');
    } else {
      var fi  = _pgnHeaderIdx_(facData[0]);
      var cFechaEm  = _pgnGetCol_(fi, ['fecha_emision', 'fecha emision', 'fechaemision']);
      var cEstado   = _pgnGetCol_(fi, ['estado']);
      var cTotal    = _pgnGetCol_(fi, ['total']);
      var cIva      = _pgnGetCol_(fi, ['iva_importe', 'iva importe', 'ivaimporte']);
      var cFechaPag = _pgnGetCol_(fi, ['fecha_pago', 'fecha pago', 'fechapago']);

      if (cEstado === -1) result.logs.push('⚠ FACTURAS: columna Estado no encontrada');
      if (cTotal  === -1) result.logs.push('⚠ FACTURAS: columna Total no encontrada');

      var nFac = 0;
      for (var r = 1; r < facData.length; r++) {
        var row = facData[r];
        if (!row[0] && !row[1]) continue;  // fila vacía

        var fechaEm   = cFechaEm  >= 0 ? row[cFechaEm]  : null;
        var estado    = cEstado   >= 0 ? String(row[cEstado]  || '').trim() : '';
        var total     = cTotal    >= 0 ? parseFloat(row[cTotal]   || 0) || 0 : 0;
        var ivaImp    = cIva      >= 0 ? parseFloat(row[cIva]    || 0) || 0 : 0;
        var fechaPag  = cFechaPag >= 0 ? row[cFechaPag] : null;

        var esAnulada = /anulad/i.test(estado);
        var esPagado  = /pagad|cobrad/i.test(estado);

        if (_pgnEsMesActual_(fechaEm) && !esAnulada) {
          result.facturadoMes   += total;
          result.ivaRepercutido += ivaImp;
        }
        if (_pgnEsMesActual_(fechaPag) && esPagado) {
          result.cobradoMes += total;
        }
        if (!esPagado && !esAnulada) {
          result.pendienteCobro += total;
        }
        nFac++;
      }
      result.logs.push('✅ FACTURAS: ' + nFac + ' registros procesados');
      Logger.log('    ✅ FACTURAS: ' + nFac + ' registros');
    }
  }

  // ── GASTOS ───────────────────────────────────────────────
  var gasSh = pgObtenerHoja_(PGN_SHEET_ALIASES.GASTOS);
  if (!gasSh) {
    result.logs.push('⚠ GASTOS no encontrada — mostrando 0');
    Logger.log('    ⚠ GASTOS: hoja no encontrada');
  } else {
    var gasData = gasSh.getDataRange().getValues();
    if (gasData.length < 2) {
      result.logs.push('ℹ GASTOS vacía');
    } else {
      var gi = _pgnHeaderIdx_(gasData[0]);
      // Fecha: preferir Fecha_Gasto, fallback Fecha_Alta
      var cFechaGas = _pgnGetCol_(gi, ['fecha_gasto', 'fecha gasto', 'fechagasto',
                                       'fecha_alta',  'fecha alta',  'fechaalta']);
      var cGTotal   = _pgnGetCol_(gi, ['total']);
      var cGIva     = _pgnGetCol_(gi, ['iva_importe', 'iva importe', 'ivaimporte']);

      if (cGTotal === -1) result.logs.push('⚠ GASTOS: columna Total no encontrada');

      var nGas = 0;
      for (var g = 1; g < gasData.length; g++) {
        var grow = gasData[g];
        if (!grow[0] && !grow[1]) continue;

        var fechaGas = cFechaGas >= 0 ? grow[cFechaGas] : null;
        var gTotal   = cGTotal   >= 0 ? parseFloat(grow[cGTotal] || 0) || 0 : 0;
        var gIva     = cGIva     >= 0 ? parseFloat(grow[cGIva]  || 0) || 0 : 0;

        if (_pgnEsMesActual_(fechaGas)) {
          result.gastosMes    += gTotal;
          result.ivaSoportado += gIva;
        }
        nGas++;
      }
      result.logs.push('✅ GASTOS: ' + nGas + ' registros procesados');
      Logger.log('    ✅ GASTOS: ' + nGas + ' registros');
    }
  }

  // ── Derivados ────────────────────────────────────────────
  result.beneficioMes = result.facturadoMes - result.gastosMes;
  result.ivaNeto      = result.ivaRepercutido - result.ivaSoportado;

  Logger.log('    facturado=' + result.facturadoMes.toFixed(2) +
             ' cobrado='    + result.cobradoMes.toFixed(2) +
             ' pendiente='  + result.pendienteCobro.toFixed(2) +
             ' gastos='     + result.gastosMes.toFixed(2));
  return result;
}

// ============================================================
// FASE 16.3 — Bloque FACTURAS
// ============================================================

function _pgnCalcFacturas_() {
  Logger.log('  _pgnCalcFacturas_()');

  var result = {
    nEmitidosMes:    0,
    nPendientes:     0,
    nPagadas:        0,
    nVencidas:       0,
    nProximasVencer: 0,   // vencen en ≤7 días (para alertas)
    totalPendiente:  0,
    nMas30d:         0,
    ultimas5:        [],
    logs:            []
  };

  var facSh = pgObtenerHoja_(PGN_SHEET_ALIASES.FACTURAS);
  if (!facSh) {
    result.logs.push('⚠ FACTURAS no encontrada');
    Logger.log('    ⚠ FACTURAS: hoja no encontrada');
    return result;
  }

  var data = facSh.getDataRange().getValues();
  if (data.length < 2) {
    result.logs.push('ℹ FACTURAS vacía');
    return result;
  }

  var tz  = Session.getScriptTimeZone();
  var hoy = new Date();
  var fi  = _pgnHeaderIdx_(data[0]);

  var cId        = _pgnGetCol_(fi, ['id_factura','id factura','idfactura','numero','id']);
  var cFechaEm   = _pgnGetCol_(fi, ['fecha_emision','fecha emision','fechaemision']);
  var cCliente   = _pgnGetCol_(fi, ['nombre_cliente','nombre cliente','cliente','nomcliente']);
  var cEstado    = _pgnGetCol_(fi, ['estado']);
  var cTotal     = _pgnGetCol_(fi, ['total']);
  var cFechaVenc = _pgnGetCol_(fi, ['fecha_vencimiento','fecha vencimiento','vencimiento','fechavencimiento']);
  var cFechaPago = _pgnGetCol_(fi, ['fecha_pago','fecha pago','fechapago']);

  if (cEstado === -1) result.logs.push('⚠ FACTURAS: columna Estado no encontrada');
  if (cTotal  === -1) result.logs.push('⚠ FACTURAS: columna Total no encontrada');

  var allFacs = [];

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (!row[0] && !row[1]) continue;

    var id        = cId        >= 0 ? String(row[cId]        || '').trim() : '—';
    var fechaEm   = cFechaEm   >= 0 ? row[cFechaEm]   : null;
    var cliente   = cCliente   >= 0 ? String(row[cCliente]   || '').trim() : '—';
    var estado    = cEstado    >= 0 ? String(row[cEstado]    || '').trim() : '';
    var total     = cTotal     >= 0 ? parseFloat(row[cTotal]  || 0) || 0  : 0;
    var fechaVenc = cFechaVenc >= 0 ? row[cFechaVenc]  : null;

    var esAnulada = /anulad/i.test(estado);
    var esPagado  = /pagad|cobrad/i.test(estado);

    // Fecha de vencimiento efectiva: campo Fecha_Vencimiento o Fecha_Emision + 30d
    var fvEfectiva = null;
    if (fechaVenc instanceof Date && !isNaN(fechaVenc.getTime())) {
      fvEfectiva = fechaVenc;
    } else if (fechaEm instanceof Date && !isNaN(fechaEm.getTime())) {
      fvEfectiva = new Date(fechaEm.getTime() + 30 * 24 * 3600 * 1000);
    }

    var esVencida      = !esPagado && !esAnulada && fvEfectiva && fvEfectiva < hoy;
    var diasParaVencer = fvEfectiva ? (fvEfectiva.getTime() - hoy.getTime()) / 86400000 : Infinity;
    var esProxima      = !esPagado && !esAnulada && !esVencida
                       && diasParaVencer >= 0 && diasParaVencer <= 7;
    var mas30d         = !esPagado && !esAnulada &&
                        (fechaEm instanceof Date && !isNaN(fechaEm.getTime())) &&
                        ((hoy - fechaEm) / (1000 * 3600 * 24) > 30);

    // Contadores
    if (_pgnEsMesActual_(fechaEm) && !esAnulada) result.nEmitidosMes++;
    if (!esPagado && !esAnulada)  { result.nPendientes++; result.totalPendiente += total; }
    if (esPagado)                   result.nPagadas++;
    if (esVencida)                  result.nVencidas++;
    if (esProxima)                  result.nProximasVencer++;
    if (mas30d)                     result.nMas30d++;

    // Guardar para la lista de últimas/atención (excluye anuladas)
    if (!esAnulada) {
      var fechaStr = (fechaEm instanceof Date && !isNaN(fechaEm.getTime()))
        ? Utilities.formatDate(fechaEm, tz, 'dd/MM/yy') : '—';
      var cli20 = cliente.length > 20 ? cliente.substring(0, 18) + '…' : cliente;
      allFacs.push({
        id:       id,
        fecha:    fechaStr,
        fechaMs:  (fechaEm instanceof Date && !isNaN(fechaEm.getTime())) ? fechaEm.getTime() : 0,
        cliente:  cli20,
        total:    total,
        estado:   estado,
        esPagada: esPagado,
        vencida:  esVencida,
        mas30d:   mas30d
      });
    }
  }

  // Ordenar: vencidas → +30d → recientes (dentro de cada grupo, más nuevo primero)
  allFacs.sort(function(a, b) {
    var sA = a.vencida ? 2 : (a.mas30d ? 1 : 0);
    var sB = b.vencida ? 2 : (b.mas30d ? 1 : 0);
    if (sA !== sB) return sB - sA;
    return b.fechaMs - a.fechaMs;
  });
  result.ultimas5 = allFacs.slice(0, 5);

  var nProc = data.length - 1;
  result.logs.push('✅ FACTURAS: ' + nProc + ' reg — emitidas=' + result.nEmitidosMes +
    ' pend=' + result.nPendientes + ' venc=' + result.nVencidas + ' +30d=' + result.nMas30d);
  Logger.log('    ' + result.logs[result.logs.length - 1]);
  return result;
}

// ── Volcar bloque FACTURAS al panel ──────────────────────────

function _pgnActualizarBloqueFacturas_(sh, datos) {
  var R = PGN_ROWS;
  var C = PGN_C;

  // ── 5 KPI de FACTURAS (left half, valor en cols 5-7) ─────
  var kpiDef = [
    { row: R.FAC_GAS_START,     val: datos.nEmitidosMes,
      fmt: '0',         fg: C.AZUL_CORP },
    { row: R.FAC_GAS_START + 1, val: datos.nPendientes,
      fmt: '0',         fg: datos.nPendientes > 0 ? C.ROJO : C.VERDE },
    { row: R.FAC_GAS_START + 2, val: datos.nPagadas,
      fmt: '0',         fg: C.VERDE },
    { row: R.FAC_GAS_START + 3, val: datos.nVencidas,
      fmt: '0',         fg: datos.nVencidas > 0 ? C.ROJO : C.VERDE },
    { row: R.FAC_GAS_START + 4, val: datos.totalPendiente,
      fmt: PGN_EUR_FMT, fg: datos.totalPendiente > 0 ? C.ROJO : C.VERDE }
  ];
  kpiDef.forEach(function(k) {
    sh.getRange(k.row, 5, 1, 3).merge()
      .setValue(k.val).setNumberFormat(k.fmt).setFontColor(k.fg)
      .setFontSize(12).setFontWeight('bold')
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
  });

  // ── Cabecera dinámica: muestra +30d si los hay ───────────
  var nAten = datos.nVencidas + datos.nMas30d;
  var hdrTxt, hdrFg;
  if (nAten > 0) {
    hdrTxt = '  ⚠ Facturas con atención (' + nAten + ' con +30d / vencidas)';
    hdrFg  = C.ROJO;
  } else {
    hdrTxt = '  Últimas 5 facturas emitidas';
    hdrFg  = C.AZUL_CORP;
  }
  sh.getRange(R.FAC_GAS_ULT, 1, 1, 7).merge()
    .setValue(hdrTxt).setBackground(C.AZUL_LIGHT).setFontColor(hdrFg)
    .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('left');

  // ── Lista de hasta 5 facturas (prioritizadas por urgencia) ─
  for (var i = 0; i < 5; i++) {
    var rRow = R.FAC_GAS_ULT + 1 + i;
    var bg   = i % 2 === 0 ? C.BLANCO : C.GRIS_BG;
    var fac  = datos.ultimas5[i];

    if (!fac) {
      sh.getRange(rRow, 1, 1, 7).merge()
        .setValue('—').setBackground(bg).setFontColor(C.GRIS)
        .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
      continue;
    }

    var ico = fac.vencida ? '🔴' : (fac.mas30d ? '🟠' : (fac.esPagada ? '🟢' : '🔵'));
    var fg  = fac.vencida ? C.ROJO : (fac.mas30d ? C.NARANJA : (fac.esPagada ? C.VERDE : C.AZUL_CORP));
    var txt = ico + '  ' + fac.id
            + '   ' + fac.fecha
            + '   ' + fac.cliente
            + '   ' + _pgnFmtEur_(fac.total)
            + '   ' + fac.estado;

    sh.getRange(rRow, 1, 1, 7).merge()
      .setValue(txt).setBackground(bg).setFontColor(fg)
      .setFontSize(10).setFontWeight('normal')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.setRowHeight(rRow, 24);
  }

  Logger.log('  ✓ Bloque FACTURAS actualizado.');
}

// ── Formatea un número como moneda española ───────────────────
// Produce "1.234,56 €" sin depender de setNumberFormat
function _pgnFmtEur_(val) {
  var n    = parseFloat(val) || 0;
  var sign = n < 0 ? '-' : '';
  var abs  = Math.abs(n).toFixed(2).split('.');
  var int  = abs[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return sign + int + ',' + abs[1] + ' €';
}

// ============================================================
// FASE 16.4 — Bloque GASTOS
// ============================================================

function _pgnCalcGastos_() {
  Logger.log('  _pgnCalcGastos_()');

  var result = {
    totalMes:        0,
    nGastos:         0,
    nPendientesPago: 0,
    nSinRevisar:     0,
    nOcrFallido:     0,
    nSinTicket:      0,
    ultimos5:        [],
    logs:            []
  };

  var gasSh = pgObtenerHoja_(PGN_SHEET_ALIASES.GASTOS);
  if (!gasSh) {
    result.logs.push('⚠ GASTOS no encontrada');
    Logger.log('    ⚠ GASTOS: hoja no encontrada');
    return result;
  }

  var data = gasSh.getDataRange().getValues();
  if (data.length < 2) {
    result.logs.push('ℹ GASTOS vacía');
    return result;
  }

  var tz = Session.getScriptTimeZone();
  var gi = _pgnHeaderIdx_(data[0]);

  var cId       = _pgnGetCol_(gi, ['id_gasto','id gasto','idgasto','id']);
  var cFechaGas = _pgnGetCol_(gi, ['fecha_gasto','fecha gasto','fechagasto']);
  var cFechaAlt = _pgnGetCol_(gi, ['fecha_alta','fecha alta','fechaalta','timestamp']);
  var cProv     = _pgnGetCol_(gi, ['proveedor']);
  var cConc     = _pgnGetCol_(gi, ['concepto']);
  var cTotal    = _pgnGetCol_(gi, ['total']);
  var cEstPago  = _pgnGetCol_(gi, ['estado_pago','estado pago','estadopago']);
  var cEstProc  = _pgnGetCol_(gi, ['estado_procesamiento','estado procesamiento','estadoprocesamiento']);
  var cRevisado = _pgnGetCol_(gi, ['revisado']);
  var cFotoUrl  = _pgnGetCol_(gi, ['foto_url','foto url','fotourl','foto']);
  var cArchId   = _pgnGetCol_(gi, ['archivo_id','archivo id','archivoid']);

  if (cTotal  === -1) result.logs.push('⚠ GASTOS: columna Total no encontrada');

  var allGas = [];

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (!row[0] && !row[1]) continue;

    // Fecha: preferir Fecha_Gasto, fallback Fecha_Alta
    var fechaGas = null;
    if (cFechaGas >= 0 && row[cFechaGas] instanceof Date && !isNaN(row[cFechaGas].getTime())) {
      fechaGas = row[cFechaGas];
    } else if (cFechaAlt >= 0 && row[cFechaAlt] instanceof Date && !isNaN(row[cFechaAlt].getTime())) {
      fechaGas = row[cFechaAlt];
    }

    var id      = cId      >= 0 ? String(row[cId]      || '').trim() : '—';
    var prov    = cProv    >= 0 ? String(row[cProv]    || '').trim() : '';
    var conc    = cConc    >= 0 ? String(row[cConc]    || '').trim() : '';
    var total   = cTotal   >= 0 ? parseFloat(row[cTotal] || 0) || 0  : 0;
    var estPago = cEstPago >= 0 ? String(row[cEstPago]  || '').trim() : '';
    var estProc = cEstProc >= 0 ? String(row[cEstProc]  || '').trim() : '';
    var revisado = cRevisado >= 0 ? row[cRevisado] : null;
    var fotoUrl  = cFotoUrl  >= 0 ? String(row[cFotoUrl] || '').trim() : '';
    var archId   = cArchId   >= 0 ? String(row[cArchId]  || '').trim() : '';

    var esPagado     = /pagad/i.test(estPago);
    var esOcrFallido = /ocr.?fallido/i.test(estProc);
    var esRevisar    = /^revisar$/i.test(estProc);

    // Sin revisar: Revisado=FALSE, o Estado_Procesamiento=Revisar/OCR fallido
    var esNoRevisado = (revisado === false) ||
                       (String(revisado || '').toLowerCase() === 'false');
    var esSinRevisar = esNoRevisado || esOcrFallido || esRevisar;

    // Sin ticket/foto: Foto_URL vacío Y Archivo_ID vacío
    var esSinTicket = !fotoUrl && !archId;

    // Acumular
    if (_pgnEsMesActual_(fechaGas)) result.totalMes += total;
    result.nGastos++;
    if (!esPagado)    result.nPendientesPago++;
    if (esSinRevisar) result.nSinRevisar++;
    if (esOcrFallido) result.nOcrFallido++;
    if (esSinTicket)  result.nSinTicket++;

    // Guardar para lista
    var fechaMs  = fechaGas ? fechaGas.getTime() : 0;
    var fechaStr = fechaGas ? Utilities.formatDate(fechaGas, tz, 'dd/MM/yy') : '—';
    var desc     = (prov || conc || id);
    var desc20   = desc.length > 20 ? desc.substring(0, 18) + '…' : desc;

    allGas.push({
      id:          id,
      fecha:       fechaStr,
      fechaMs:     fechaMs,
      desc:        desc20,
      total:       total,
      estPago:     estPago,
      esPagado:    esPagado,
      sinRevisar:  esSinRevisar,
      ocrFallido:  esOcrFallido,
      sinTicket:   esSinTicket
    });
  }

  // Ordenar: OCR fallido → sin revisar → sin ticket → más reciente
  allGas.sort(function(a, b) {
    var sA = a.ocrFallido ? 3 : (a.sinRevisar ? 2 : (a.sinTicket ? 1 : 0));
    var sB = b.ocrFallido ? 3 : (b.sinRevisar ? 2 : (b.sinTicket ? 1 : 0));
    if (sA !== sB) return sB - sA;
    return b.fechaMs - a.fechaMs;
  });
  result.ultimos5 = allGas.slice(0, 5);

  result.logs.push('✅ GASTOS: ' + result.nGastos + ' reg — totalMes=' +
    result.totalMes.toFixed(2) + ' pend_pago=' + result.nPendientesPago +
    ' sin_rev=' + result.nSinRevisar + ' ocr=' + result.nOcrFallido +
    ' sin_ticket=' + result.nSinTicket);
  Logger.log('    ' + result.logs[result.logs.length - 1]);
  return result;
}

// ── Volcar bloque GASTOS al panel (right half, cols 8-14) ────

function _pgnActualizarBloqueGastos_(sh, datos) {
  var R = PGN_ROWS;
  var C = PGN_C;

  // ── 5 KPI de GASTOS (valor en cols 12-14) ────────────────
  var kpiDef = [
    { row: R.FAC_GAS_START,     val: datos.totalMes,
      fmt: PGN_EUR_FMT, fg: C.NARANJA },
    { row: R.FAC_GAS_START + 1, val: datos.nGastos,
      fmt: '0',         fg: C.AZUL_CORP },
    { row: R.FAC_GAS_START + 2, val: datos.nPendientesPago,
      fmt: '0',         fg: datos.nPendientesPago > 0 ? C.ROJO : C.VERDE },
    { row: R.FAC_GAS_START + 3, val: datos.nSinRevisar,
      fmt: '0',         fg: datos.nSinRevisar > 0 ? C.NARANJA : C.VERDE },
    { row: R.FAC_GAS_START + 4, val: datos.nOcrFallido,
      fmt: '0',         fg: datos.nOcrFallido > 0 ? C.ROJO : C.VERDE }
  ];
  kpiDef.forEach(function(k) {
    sh.getRange(k.row, 12, 1, 3).merge()
      .setValue(k.val).setNumberFormat(k.fmt).setFontColor(k.fg)
      .setFontSize(12).setFontWeight('bold')
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
  });

  // ── Cabecera dinámica: muestra OCR + sin-ticket si los hay ─
  var alertas = [];
  if (datos.nOcrFallido > 0) alertas.push(datos.nOcrFallido + ' OCR fallido');
  if (datos.nSinTicket  > 0) alertas.push(datos.nSinTicket  + ' sin ticket/foto');

  var hdrTxt, hdrFg;
  if (alertas.length > 0) {
    hdrTxt = '  ⚠ Últimos gastos — ' + alertas.join(', ');
    hdrFg  = C.ROJO;
  } else {
    hdrTxt = '  Últimos 5 gastos registrados';
    hdrFg  = C.NARANJA;
  }
  sh.getRange(R.FAC_GAS_ULT, 8, 1, 7).merge()
    .setValue(hdrTxt).setBackground(C.NARANJA_BG).setFontColor(hdrFg)
    .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('left');

  // ── Lista de hasta 5 gastos (prioritizados por urgencia) ──
  for (var i = 0; i < 5; i++) {
    var rRow = R.FAC_GAS_ULT + 1 + i;
    var bg   = i % 2 === 0 ? C.BLANCO : C.GRIS_BG;
    var gas  = datos.ultimos5[i];

    if (!gas) {
      sh.getRange(rRow, 8, 1, 7).merge()
        .setValue('—').setBackground(bg).setFontColor(C.GRIS)
        .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
      continue;
    }

    var ico = gas.ocrFallido ? '🔴'
            : (gas.sinRevisar ? '🟠'
            : (gas.sinTicket  ? '🟡'
            : (gas.esPagado   ? '🟢' : '🔵')));
    var fg  = gas.ocrFallido ? C.ROJO
            : (gas.sinRevisar ? C.NARANJA
            : (gas.esPagado   ? C.VERDE : C.AZUL_CORP));
    var txt = ico + '  ' + gas.fecha
            + '   ' + gas.desc
            + '   ' + _pgnFmtEur_(gas.total)
            + '   ' + gas.estPago;

    sh.getRange(rRow, 8, 1, 7).merge()
      .setValue(txt).setBackground(bg).setFontColor(fg)
      .setFontSize(10).setFontWeight('normal')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.setRowHeight(rRow, 24);
  }

  Logger.log('  ✓ Bloque GASTOS actualizado.');
}

// ============================================================
// FASE 16.5 — Bloque TRABAJOS
// ============================================================

function _pgnCalcTrabajos_() {
  Logger.log('  _pgnCalcTrabajos_()');

  var result = {
    nRealizadosMes: 0,
    nPendientes:    0,
    nCompletados:   0,
    horasMes:       0,
    nTotal:         0,
    ultimos5:       [],
    hojaUsada:      null,
    colsDetectadas: [],
    colsFaltantes:  [],
    logs:           []
  };

  var trabSh = pgObtenerHoja_(PGN_SHEET_ALIASES.TRABAJOS);
  if (!trabSh) {
    result.logs.push('⚠ TRABAJOS no encontrada');
    Logger.log('    ⚠ TRABAJOS: hoja no encontrada');
    return result;
  }
  result.hojaUsada = trabSh.getName();

  var data = trabSh.getDataRange().getValues();
  if (data.length < 2) {
    result.logs.push('ℹ TRABAJOS vacía');
    return result;
  }

  var tz = Session.getScriptTimeZone();
  var ti = _pgnHeaderIdx_(data[0]);

  var cFecha    = _pgnGetCol_(ti, ['fecha_inicio','fecha inicio','fechainicio','fecha_trabajo','fecha trabajo','fecha_creacion','fecha creacion','fecha','timestamp']);
  var cFechaFin = _pgnGetCol_(ti, ['fecha_fin_estimada','fecha fin estimada','fecha_fin','fechafin']);
  var cEstado   = _pgnGetCol_(ti, ['estado','status','estado_trabajo']);
  var cCliente  = _pgnGetCol_(ti, ['nombre_cliente','nomcliente','id_cliente','cliente']);
  var cOperario = _pgnGetCol_(ti, ['nombre_operario','operario','id_operario']);
  var cHoras    = _pgnGetCol_(ti, ['horas_trabajadas','horas trabajadas','horas','hours']);
  var cHoraIni  = _pgnGetCol_(ti, ['hora_inicio','hora inicio','horainicio']);
  var cHoraFin  = _pgnGetCol_(ti, ['hora_fin','hora fin','horafin','hora_final']);
  // Empleado 2 (opcional)
  var cOperario2 = _pgnGetCol_(ti, ['operario_2','operario 2','operario2','nombre_operario_2']);
  var cHoraIni2  = _pgnGetCol_(ti, ['hora_inicio_2','hora inicio 2','horainicio2','hora_inicio2']);
  var cHoraFin2  = _pgnGetCol_(ti, ['hora_fin_2','hora fin 2','horafin2','hora_fin2']);
  var cDesc     = _pgnGetCol_(ti, ['descripcion','descripción','description','desc']);
  var cTotal    = _pgnGetCol_(ti, ['precio_total','precio total','total','importe']);

  var colMap = { 'Fecha': cFecha, 'Estado': cEstado, 'Cliente': cCliente, 'Operario': cOperario, 'Horas': cHoras };
  Object.keys(colMap).forEach(function(name) {
    if (colMap[name] !== -1) result.colsDetectadas.push(name);
    else result.colsFaltantes.push(name);
  });

  if (cFecha  === -1) result.logs.push('⚠ TRABAJOS: col Fecha no encontrada');
  if (cEstado === -1) result.logs.push('⚠ TRABAJOS: col Estado no encontrada');

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var filas = [];
  for (var i = 1; i < data.length; i++) {
    var row   = data[i];
    var idTrab = row[0];
    if (!idTrab) continue;
    result.nTotal++;

    var rawFecha = cFecha !== -1 ? row[cFecha] : null;
    var fecha    = rawFecha instanceof Date ? rawFecha : (rawFecha ? new Date(rawFecha) : null);
    if (fecha && isNaN(fecha.getTime())) fecha = null;

    var rawFechaFin = cFechaFin !== -1 ? row[cFechaFin] : null;
    var fechaFin    = rawFechaFin instanceof Date ? rawFechaFin : (rawFechaFin ? new Date(rawFechaFin) : null);
    if (fechaFin && isNaN(fechaFin.getTime())) fechaFin = null;

    var estado   = cEstado   !== -1 ? String(row[cEstado]).trim()   : '';
    var cliente  = cCliente  !== -1 ? String(row[cCliente]).trim()  : '—';
    var desc     = cDesc     !== -1 ? String(row[cDesc]).trim()     : '—';
    var total    = cTotal    !== -1 ? (parseFloat(row[cTotal]) || 0) : 0;

    var esCompletado = /complet/i.test(estado);
    var esCancelado  = /cancel/i.test(estado);
    var esPendiente  = !esCompletado && !esCancelado;

    if (esPendiente)  result.nPendientes++;
    if (esCompletado) result.nCompletados++;

    var esMesActual = fecha && _pgnEsMesActual_(fecha, tz);
    if (esMesActual) result.nRealizadosMes++;

    if (esMesActual) {
      var horas = 0;
      if (cHoras !== -1 && row[cHoras] !== '' && row[cHoras] !== null && row[cHoras] !== undefined) {
        horas = parseFloat(row[cHoras]) || 0;
      } else {
        // Empleado 1. _calcHoras_ admite Date, número fraccionario y "HH:MM".
        if (cHoraIni !== -1 && cHoraFin !== -1) {
          horas += _calcHoras_(row[cHoraIni], row[cHoraFin]);
        }
        // Empleado 2 (solo si tiene operario asignado).
        if (cOperario2 !== -1 && String(row[cOperario2] || '').trim() &&
            cHoraIni2 !== -1 && cHoraFin2 !== -1) {
          horas += _calcHoras_(row[cHoraIni2], row[cHoraFin2]);
        }
      }
      result.horasMes += horas;
    }

    var vencido  = esPendiente && fechaFin && fechaFin < today;
    var semaforo = esCompletado ? '🟢' : esCancelado ? '⚫' : vencido ? '🔴' : '🟠';

    filas.push({
      fecha:        fecha,
      fechaStr:     fecha ? Utilities.formatDate(fecha, tz, 'dd/MM/yy') : '—',
      estado:       estado || '—',
      cliente:      cliente.substring(0, 25),
      desc:         desc.substring(0, 30),
      total:        total,
      semaforo:     semaforo,
      esCompletado: esCompletado,
      esCancelado:  esCancelado,
      vencido:      vencido
    });
  }

  filas.sort(function(a, b) {
    var priA = a.vencido ? 0 : (a.esCancelado || a.esCompletado ? 2 : 1);
    var priB = b.vencido ? 0 : (b.esCancelado || b.esCompletado ? 2 : 1);
    if (priA !== priB) return priA - priB;
    var ta = a.fecha ? a.fecha.getTime() : 0;
    var tb = b.fecha ? b.fecha.getTime() : 0;
    return tb - ta;
  });

  result.ultimos5  = filas.slice(0, 5);
  result.horasMes  = Math.round(result.horasMes * 10) / 10;

  result.logs.push(
    'TRABAJOS: ' + result.nTotal + ' total, ' +
    result.nPendientes + ' pend, ' +
    result.nCompletados + ' compl, ' +
    result.horasMes + 'h mes'
  );
  Logger.log('    TRABAJOS: ' + result.nTotal + ' total, ' +
    result.nRealizadosMes + ' este mes, ' +
    result.nPendientes + ' pendientes, ' +
    result.horasMes + 'h');

  return result;
}

function _pgnActualizarBloqueTrabajos_(sh, datos) {
  var R = PGN_ROWS;
  var C = PGN_C;

  // ── KPI values izquierda: cols 5-7, filas 28-31 ─────────────
  var kpiDef = [
    { row: R.TRAB_CLI_START + 0, val: datos.nRealizadosMes, fmt: '0',   fg: C.VERDE_DARK },
    { row: R.TRAB_CLI_START + 1, val: datos.nPendientes,    fmt: '0',   fg: datos.nPendientes > 0 ? C.NARANJA : C.VERDE },
    { row: R.TRAB_CLI_START + 2, val: datos.nCompletados,   fmt: '0',   fg: C.VERDE      },
    { row: R.TRAB_CLI_START + 3, val: datos.horasMes,       fmt: '0.0', fg: C.AZUL_CORP  }
  ];
  kpiDef.forEach(function(k) {
    sh.getRange(k.row, 5, 1, 3).merge()
      .setValue(k.val).setNumberFormat(k.fmt).setFontColor(k.fg)
      .setFontSize(12).setFontWeight('bold')
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
  });

  // ── KDEST card 3 — Trabajos pendientes (fila 6, cols 8-9) ──
  sh.getRange(R.KDEST_VAL, 8, 1, 2).merge()
    .setValue(datos.nPendientes).setNumberFormat('0')
    .setFontColor(datos.nPendientes > 0 ? C.NARANJA : C.VERDE)
    .setFontSize(26).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── Cabecera dinámica (fila TRAB_CLI_ULT, cols 1-7) ─────────
  var nVenc = datos.ultimos5.filter(function(t) { return t.vencido; }).length;
  var hdrTxt, hdrFg, hdrBg;
  if (nVenc > 0) {
    hdrTxt = '  ⚠ ' + nVenc + ' trabajo(s) vencido(s)';
    hdrFg  = C.ROJO;
    hdrBg  = C.ROJO_BG;
  } else if (datos.nPendientes > 0) {
    hdrTxt = '  ⚠ ' + datos.nPendientes + ' pendiente(s)';
    hdrFg  = C.NARANJA;
    hdrBg  = '#d1e7dd';
  } else {
    hdrTxt = '  ✅ Últimos trabajos';
    hdrFg  = C.VERDE_DARK;
    hdrBg  = '#d1e7dd';
  }
  sh.getRange(R.TRAB_CLI_ULT, 1, 1, 7).merge()
    .setValue(hdrTxt).setBackground(hdrBg).setFontColor(hdrFg)
    .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('left');

  // ── Lista de hasta 5 trabajos (filas 33-37, cols 1-7) ───────
  for (var i = 0; i < 5; i++) {
    var rRow = R.TRAB_CLI_ULT + 1 + i;
    var bg   = i % 2 === 0 ? C.BLANCO : C.GRIS_BG;
    var trab = datos.ultimos5[i];

    if (!trab) {
      sh.getRange(rRow, 1, 1, 7).merge()
        .setValue('—').setBackground(bg).setFontColor(C.GRIS)
        .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
      continue;
    }

    var fg = trab.vencido      ? C.ROJO
           : trab.esCompletado ? C.VERDE
           : trab.esCancelado  ? C.GRIS
           : C.NARANJA;
    var txt = trab.semaforo + '  ' + trab.fechaStr
            + '   ' + trab.cliente
            + '   ' + trab.desc
            + (trab.total > 0 ? '   ' + _pgnFmtEur_(trab.total) : '');

    sh.getRange(rRow, 1, 1, 7).merge()
      .setValue(txt).setBackground(bg).setFontColor(fg)
      .setFontSize(10).setFontWeight('normal')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.setRowHeight(rRow, 24);
  }

  Logger.log('  ✓ Bloque TRABAJOS actualizado.');
}

// ============================================================
// FASE 16.6 — Bloque CLIENTES
// ============================================================

function _pgnCalcClientes_() {
  Logger.log('  _pgnCalcClientes_()');

  var result = {
    nActivos:         0,
    nConPendientes:   0,
    nSinActividad:    0,
    facturacionTotal: 0,
    top5:             [],
    hojaUsada:        null,
    colsDetectadas:   [],
    colsFaltantes:    [],
    logs:             []
  };

  var hoy   = new Date();
  var hace60 = new Date(hoy.getTime() - 60 * 24 * 3600000);

  // ── CLIENTES — construir mapa de clientes activos ─────────
  var cliSh = pgObtenerHoja_(PGN_SHEET_ALIASES.CLIENTES);
  result.hojaUsada = cliSh ? cliSh.getName() : null;
  if (!cliSh) result.logs.push('⚠ CLIENTES no encontrada');

  // activosNorm: {normName → true}  (solo activos para sin-actividad)
  // todosNorm:   {normName → true}  (todos, fallback)
  var activosNorm  = {};
  var todosNorm    = {};
  var nDescartados = 0;
  var tieneColActivo = false;

  if (cliSh) {
    var cliData = cliSh.getDataRange().getValues();
    if (cliData.length > 1) {
      var ci         = _pgnHeaderIdx_(cliData[0]);
      var cNombre    = _pgnGetCol_(ci, ['nombre']);
      var cApellidos = _pgnGetCol_(ci, ['apellidos']);
      var cEmpresa   = _pgnGetCol_(ci, ['empresa']);
      var cActivo    = _pgnGetCol_(ci, ['activo']);
      tieneColActivo = cActivo !== -1;

      var colDet = [], colMiss = [];
      [['Nombre',cNombre],['Apellidos',cApellidos],['Empresa',cEmpresa],['Activo',cActivo]]
        .forEach(function(p){ if(p[1]!==-1) colDet.push(p[0]); else colMiss.push(p[0]); });
      result.colsDetectadas = colDet;
      result.colsFaltantes  = colMiss;

      for (var i = 1; i < cliData.length; i++) {
        var rowCli = cliData[i];
        if (!rowCli[0]) continue;

        var nombreCli = '';
        if (cEmpresa !== -1 && String(rowCli[cEmpresa]).trim()) {
          nombreCli = String(rowCli[cEmpresa]).trim();
        } else {
          var n = cNombre    !== -1 ? String(rowCli[cNombre]).trim()    : '';
          var a = cApellidos !== -1 ? String(rowCli[cApellidos]).trim() : '';
          nombreCli = (n + ' ' + a).trim();
        }
        if (!nombreCli) { nDescartados++; continue; }

        var norm = nombreCli.toUpperCase();
        todosNorm[norm] = true;

        var esActivo = true;
        if (tieneColActivo) {
          var actVal = String(rowCli[cActivo]).trim().toUpperCase();
          esActivo = actVal === 'TRUE' || actVal === 'SÍ' || actVal === 'SI'
                  || actVal === '1'    || actVal === 'VERDADERO';
        }
        if (esActivo) {
          activosNorm[norm] = true;
          result.nActivos++;
        } else {
          nDescartados++;
        }
      }
    }
  }

  // ── FACTURAS — agrupación por cliente normalizado ─────────
  var facSh = pgObtenerHoja_(PGN_SHEET_ALIASES.FACTURAS);
  // {normName → {original: str, total: number}}
  var facPorCli       = {};
  var pendientesCli   = {};  // normName → true
  var facRecenteNorm  = {};  // normName → true (últimos 60d)

  if (facSh) {
    var facData = facSh.getDataRange().getValues();
    if (facData.length > 1) {
      var fi      = _pgnHeaderIdx_(facData[0]);
      var fNom    = _pgnGetCol_(fi, ['nombre_cliente','nomcliente','cliente']);
      var fTotal  = _pgnGetCol_(fi, ['total']);
      var fEstado = _pgnGetCol_(fi, ['estado']);
      var fFecha  = _pgnGetCol_(fi, ['fecha_emision','fecha emision','fecha_alta','fecha']);

      for (var j = 1; j < facData.length; j++) {
        var rowFac = facData[j];
        if (!rowFac[0]) continue;

        var nomCli = fNom !== -1 ? String(rowFac[fNom]).trim() : '';
        if (!nomCli) continue;
        var normCli = nomCli.toUpperCase();

        var total   = fTotal  !== -1 ? (parseFloat(rowFac[fTotal])  || 0) : 0;
        var estado  = fEstado !== -1 ? String(rowFac[fEstado]).trim().toLowerCase() : '';
        var rawFec  = fFecha  !== -1 ? rowFac[fFecha] : null;
        var fechaFac = rawFec instanceof Date ? rawFec : (rawFec ? new Date(rawFec) : null);
        if (fechaFac && isNaN(fechaFac.getTime())) fechaFac = null;

        var esCancelada = /cancel/i.test(estado);
        if (!esCancelada) {
          if (!facPorCli[normCli]) facPorCli[normCli] = { original: nomCli, total: 0 };
          facPorCli[normCli].total += total;
          result.facturacionTotal  += total;
        }

        var esPagada = /paga|cobra/i.test(estado);
        if (!esPagada && !esCancelada && estado !== '') {
          pendientesCli[normCli] = true;
        }

        if (fechaFac && fechaFac >= hace60) facRecenteNorm[normCli] = true;
      }
    }
  } else {
    result.logs.push('⚠ FACTURAS no encontrada — pendientes y top5 serán 0');
  }

  // ── TRABAJOS — actividad reciente ─────────────────────────
  var trabSh2 = pgObtenerHoja_(PGN_SHEET_ALIASES.TRABAJOS);
  var trabRecenteNorm = {};  // normName → true (últimos 60d)

  if (trabSh2) {
    var trabData2 = trabSh2.getDataRange().getValues();
    if (trabData2.length > 1) {
      var ti2     = _pgnHeaderIdx_(trabData2[0]);
      var tNom    = _pgnGetCol_(ti2, ['nombre_cliente','nomcliente','id_cliente','cliente']);
      var tFecha  = _pgnGetCol_(ti2, ['fecha_inicio','fecha inicio','fechainicio','fecha_trabajo','fecha trabajo','fecha','timestamp']);

      for (var k2 = 1; k2 < trabData2.length; k2++) {
        var rowT = trabData2[k2];
        if (!rowT[0]) continue;
        var nomT  = tNom !== -1 ? String(rowT[tNom]).trim() : '';
        if (!nomT) continue;
        var normT  = nomT.toUpperCase();
        var rawFecT = tFecha !== -1 ? rowT[tFecha] : null;
        var fechaT  = rawFecT instanceof Date ? rawFecT : (rawFecT ? new Date(rawFecT) : null);
        if (fechaT && isNaN(fechaT.getTime())) fechaT = null;
        if (fechaT && fechaT >= hace60) trabRecenteNorm[normT] = true;
      }
    }
  }

  // ── Calcular métricas finales ─────────────────────────────
  result.nConPendientes = Object.keys(pendientesCli).length;

  // Sin actividad: basado en activos (o todos si no hay col Activo)
  var baseNorm = result.nActivos > 0 ? activosNorm : todosNorm;
  Object.keys(baseNorm).forEach(function(norm) {
    if (!facRecenteNorm[norm] && !trabRecenteNorm[norm]) result.nSinActividad++;
  });

  // Top 5 por facturación (orden desc)
  var clientesArr = Object.keys(facPorCli).map(function(norm) {
    return {
      nombre:         facPorCli[norm].original,
      total:          facPorCli[norm].total,
      tienePendiente: !!pendientesCli[norm],
      sinActividad:   !facRecenteNorm[norm] && !trabRecenteNorm[norm]
    };
  });
  clientesArr.sort(function(a, b) { return b.total - a.total; });

  result.top5 = clientesArr.slice(0, 5).map(function(c) {
    return {
      nombre:         c.nombre.substring(0, 28),
      total:          c.total,
      semaforo:       c.tienePendiente ? '🔴' : (c.sinActividad ? '🟠' : '🟢'),
      tienePendiente: c.tienePendiente,
      sinActividad:   c.sinActividad
    };
  });

  // ── Logs ─────────────────────────────────────────────────
  var nAgrupados = Object.keys(facPorCli).length;
  result.logs.push(
    'CLIENTES: ' + result.nActivos + ' activos, ' +
    result.nConPendientes + ' con pend, ' +
    result.nSinActividad + ' sin activ'
  );
  result.logs.push('CLIENTES: ' + nAgrupados + ' agrupados de FACTURAS');
  if (nDescartados > 0) result.logs.push('CLIENTES: ' + nDescartados + ' descartados (sin nombre/inactivos)');

  Logger.log('    CLIENTES: ' + result.nActivos + ' activos, ' +
    result.nConPendientes + ' con pend, ' +
    result.nSinActividad + ' sin activ 60d, ' +
    nAgrupados + ' en FACTURAS');

  return result;
}

function _pgnActualizarBloqueClientes_(sh, datos) {
  var R = PGN_ROWS;
  var C = PGN_C;

  // ── KPI values derecha: cols 12-14, filas 28-31 ─────────────
  var kpiDef = [
    { row: R.TRAB_CLI_START + 0, val: datos.nActivos,
      fmt: '0',        fg: C.GRIS_DARK },
    { row: R.TRAB_CLI_START + 1, val: datos.nConPendientes,
      fmt: '0',        fg: datos.nConPendientes > 0 ? C.ROJO   : C.VERDE },
    { row: R.TRAB_CLI_START + 2, val: datos.nSinActividad,
      fmt: '0',        fg: datos.nSinActividad  > 0 ? C.GRIS   : C.VERDE },
    { row: R.TRAB_CLI_START + 3, val: datos.facturacionTotal,
      fmt: PGN_EUR_FMT, fg: C.VERDE }
  ];
  kpiDef.forEach(function(k) {
    sh.getRange(k.row, 12, 1, 3).merge()
      .setValue(k.val).setNumberFormat(k.fmt).setFontColor(k.fg)
      .setFontSize(12).setFontWeight('bold')
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
  });

  // ── Cabecera dinámica (fila TRAB_CLI_ULT, cols 8-14) ────────
  var hdrTxt, hdrFg, hdrBg;
  if (datos.nConPendientes > 0) {
    hdrTxt = '  ⚠ ' + datos.nConPendientes + ' cliente(s) con facturas pendientes';
    hdrFg  = C.ROJO;
    hdrBg  = C.ROJO_BG;
  } else {
    hdrTxt = '  ✅ Todos los clientes al día';
    hdrFg  = C.VERDE_DARK;
    hdrBg  = '#e9ecef';
  }
  sh.getRange(R.TRAB_CLI_ULT, 8, 1, 7).merge()
    .setValue(hdrTxt).setBackground(hdrBg).setFontColor(hdrFg)
    .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('left');

  // ── Top 5 clientes (filas 33-37, cols 8-14) ─────────────────
  for (var i = 0; i < 5; i++) {
    var rRow = R.TRAB_CLI_ULT + 1 + i;
    var bg   = i % 2 === 0 ? C.BLANCO : C.GRIS_BG;
    var cli  = datos.top5[i];

    if (!cli) {
      sh.getRange(rRow, 8, 1, 7).merge()
        .setValue('—').setBackground(bg).setFontColor(C.GRIS)
        .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
      continue;
    }

    var fg = cli.tienePendiente ? C.ROJO
           : cli.sinActividad   ? C.GRIS
           : C.VERDE_DARK;
    var txt = cli.semaforo + '  ' + cli.nombre
            + '   ' + _pgnFmtEur_(cli.total);

    sh.getRange(rRow, 8, 1, 7).merge()
      .setValue(txt).setBackground(bg).setFontColor(fg)
      .setFontSize(10).setFontWeight('normal')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.setRowHeight(rRow, 24);
  }

  Logger.log('  ✓ Bloque CLIENTES actualizado.');
}

// ============================================================
// FASE 16.7 — Bloque PRESUPUESTOS / PROFORMAS
// ============================================================

function _pgnCalcPresupuestos_() {
  Logger.log('  _pgnCalcPresupuestos_()');

  var result = {
    nPendientes:       0,
    nAceptados:        0,
    nRechazados:       0,
    nPresupPend15d:    0,   // pendientes con más de 15 días sin respuesta (para alertas)
    nProformasPend:    0,
    ultimos3:          [],
    hojaPresup:        null,
    hojaProforma:      null,
    colsPresup:        [],
    colsMissPresup:    [],
    colsProforma:      [],
    colsMissProforma:  [],
    logs:              []
  };

  var tz       = Session.getScriptTimeZone();
  var allItems = [];   // mezclamos presup + proformas para ordenar
  var haceQuince = new Date(new Date().getTime() - 15 * 24 * 3600000);

  // ── PRESUPUESTOS ──────────────────────────────────────────
  var prepSh = pgObtenerHoja_(PGN_SHEET_ALIASES.PRESUPUESTOS);
  result.hojaPresup = prepSh ? prepSh.getName() : null;
  if (!prepSh) {
    result.logs.push('⚠ PRESUPUESTOS no encontrada');
    Logger.log('    ⚠ PRESUPUESTOS: hoja no encontrada');
  } else {
    var prepData = prepSh.getDataRange().getValues();
    if (prepData.length > 1) {
      var pi = _pgnHeaderIdx_(prepData[0]);

      var pEst = _pgnGetCol_(pi, ['estado','estado_presupuesto','estadopresupuesto']);
      var pFec = _pgnGetCol_(pi, ['fecha_emision','fecha emision','fecha_creacion','fecha creacion','fecha']);
      var pCli = _pgnGetCol_(pi, ['nombre_cliente','nomcliente','cliente']);
      var pTot = _pgnGetCol_(pi, ['total','importe','total_presupuesto']);

      var pDet = [], pMiss = [];
      [['Estado',pEst],['Fecha',pFec],['Cliente',pCli],['Total',pTot]]
        .forEach(function(x){ if(x[1]!==-1) pDet.push(x[0]); else pMiss.push(x[0]); });
      result.colsPresup    = pDet;
      result.colsMissPresup = pMiss;
      if (pEst === -1) result.logs.push('⚠ PRESUPUESTOS: col Estado no encontrada');

      for (var i = 1; i < prepData.length; i++) {
        var row = prepData[i];
        if (!row[0]) continue;

        var est = pEst !== -1 ? String(row[pEst]).trim() : '';
        var estL = est.toLowerCase();
        var cli = pCli !== -1 ? String(row[pCli]).trim() : '—';
        var tot = pTot !== -1 ? (parseFloat(row[pTot]) || 0) : 0;
        var rawF = pFec !== -1 ? row[pFec] : null;
        var fec  = rawF instanceof Date ? rawF : (rawF ? new Date(rawF) : null);
        if (fec && isNaN(fec.getTime())) fec = null;

        var esPend  = /pendiente|enviado/i.test(estL);
        var esAcep  = /aceptado/i.test(estL);
        var esRech  = /rechazado/i.test(estL);

        if (esPend) result.nPendientes++;
        if (esAcep) result.nAceptados++;
        if (esRech) result.nRechazados++;
        if (esPend && fec && fec < haceQuince) result.nPresupPend15d++;

        allItems.push({
          tipo:        'PRE',
          fecha:       fec,
          fechaStr:    fec ? Utilities.formatDate(fec, tz, 'dd/MM/yy') : '—',
          cliente:     cli.substring(0, 25),
          total:       tot,
          estado:      est || '—',
          semaforo:    esRech ? '🔴' : esPend ? '🟠' : esAcep ? '🟢' : '🔵',
          esPendiente: esPend,
          esAceptado:  esAcep,
          esRechazado: esRech
        });
      }
    } else {
      result.logs.push('ℹ PRESUPUESTOS vacía');
    }
  }

  // ── PROFORMAS ─────────────────────────────────────────────
  var profSh = pgObtenerHoja_(PGN_SHEET_ALIASES.PROFORMAS);
  result.hojaProforma = profSh ? profSh.getName() : null;
  if (!profSh) {
    result.logs.push('⚠ PROFORMAS no encontrada');
    Logger.log('    ⚠ PROFORMAS: hoja no encontrada');
  } else {
    var profData = profSh.getDataRange().getValues();
    if (profData.length > 1) {
      var qi = _pgnHeaderIdx_(profData[0]);

      var qEst = _pgnGetCol_(qi, ['estado','estado_proforma','estadoproforma','estado_cobro']);
      var qFec = _pgnGetCol_(qi, ['fecha_emision','fecha emision','fecha_creacion','fecha creacion','fecha']);
      var qCli = _pgnGetCol_(qi, ['nombre_cliente','nomcliente','cliente']);
      var qTot = _pgnGetCol_(qi, ['total','importe','total_proforma']);

      var qDet = [], qMiss = [];
      [['Estado',qEst],['Fecha',qFec],['Cliente',qCli],['Total',qTot]]
        .forEach(function(x){ if(x[1]!==-1) qDet.push(x[0]); else qMiss.push(x[0]); });
      result.colsProforma    = qDet;
      result.colsMissProforma = qMiss;
      if (qEst === -1) result.logs.push('⚠ PROFORMAS: col Estado no encontrada');

      for (var j = 1; j < profData.length; j++) {
        var rowP = profData[j];
        if (!rowP[0]) continue;

        var estP  = qEst !== -1 ? String(rowP[qEst]).trim() : '';
        var estPL = estP.toLowerCase();
        var cliP  = qCli !== -1 ? String(rowP[qCli]).trim() : '—';
        var totP  = qTot !== -1 ? (parseFloat(rowP[qTot]) || 0) : 0;
        var rawFP = qFec !== -1 ? rowP[qFec] : null;
        var fecP  = rawFP instanceof Date ? rawFP : (rawFP ? new Date(rawFP) : null);
        if (fecP && isNaN(fecP.getTime())) fecP = null;

        var esPagP  = /paga|cobra/i.test(estPL);
        var esAnulP = /anul/i.test(estPL);
        var esPendP = !esPagP && !esAnulP;

        if (esPendP) result.nProformasPend++;

        allItems.push({
          tipo:        'PRF',
          fecha:       fecP,
          fechaStr:    fecP ? Utilities.formatDate(fecP, tz, 'dd/MM/yy') : '—',
          cliente:     cliP.substring(0, 25),
          total:       totP,
          estado:      estP || '—',
          semaforo:    esAnulP ? '🔴' : esPendP ? '🟠' : '🟢',
          esPendiente: esPendP,
          esAceptado:  esPagP,
          esRechazado: esAnulP
        });
      }
    } else {
      result.logs.push('ℹ PROFORMAS vacía');
    }
  }

  // ── Top 3 más recientes (mixto) ───────────────────────────
  allItems.sort(function(a, b) {
    var ta = a.fecha ? a.fecha.getTime() : 0;
    var tb = b.fecha ? b.fecha.getTime() : 0;
    return tb - ta;
  });
  result.ultimos3 = allItems.slice(0, 3);

  result.logs.push(
    'PRESP: ' + result.nPendientes + ' pend, ' +
    result.nAceptados + ' acept, ' + result.nRechazados + ' rechaz'
  );
  result.logs.push('PROFORMAS: ' + result.nProformasPend + ' pend pago');

  Logger.log('    PRESP: ' + result.nPendientes + ' pend, ' +
    result.nAceptados + ' acept, ' + result.nRechazados + ' rechaz' +
    ' | PROFORMAS: ' + result.nProformasPend + ' pend pago');

  return result;
}

function _pgnActualizarBloquePresupuestos_(sh, datos) {
  var R  = PGN_ROWS;
  var C  = PGN_C;
  var nc = PGN_NC;

  // ── KPI values: cols 9-14 para filas 40-43 ──────────────────
  var kpiDef = [
    { row: R.PRESP_START + 0, val: datos.nPendientes,
      fg: datos.nPendientes    > 0 ? C.NARANJA   : C.VERDE },
    { row: R.PRESP_START + 1, val: datos.nAceptados,
      fg: C.VERDE },
    { row: R.PRESP_START + 2, val: datos.nRechazados,
      fg: datos.nRechazados    > 0 ? C.ROJO      : C.VERDE },
    { row: R.PRESP_START + 3, val: datos.nProformasPend,
      fg: datos.nProformasPend > 0 ? C.AZUL_CORP : C.VERDE }
  ];
  kpiDef.forEach(function(k) {
    sh.getRange(k.row, 9, 1, 6).merge()
      .setValue(k.val).setNumberFormat('0').setFontColor(k.fg)
      .setFontSize(13).setFontWeight('bold')
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
  });

  // ── Cabecera dinámica (fila PRESP_ULT, ancho completo) ──────
  var hdrTxt, hdrFg, hdrBg;
  if (datos.nRechazados > 0) {
    hdrTxt = '  ⚠ ' + datos.nRechazados + ' presupuesto(s) rechazado(s)';
    hdrFg  = C.ROJO;
    hdrBg  = C.ROJO_BG;
  } else if (datos.nPendientes > 0) {
    hdrTxt = '  ⚠ ' + datos.nPendientes + ' presupuesto(s) pendiente(s) de respuesta';
    hdrFg  = C.MORADO;
    hdrBg  = C.MORADO_BG;
  } else {
    hdrTxt = '  ✅ Todos los presupuestos respondidos';
    hdrFg  = C.VERDE_DARK;
    hdrBg  = C.VERDE_BG;
  }
  sh.getRange(R.PRESP_ULT, 1, 1, nc).merge()
    .setValue(hdrTxt).setBackground(hdrBg).setFontColor(hdrFg)
    .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('left');

  // ── Lista mixta: 3 más recientes (cols 1-14) ────────────────
  for (var i = 0; i < 3; i++) {
    var rRow = R.PRESP_ULT + 1 + i;
    var bg   = i % 2 === 0 ? C.BLANCO : C.GRIS_BG;
    var item = datos.ultimos3[i];

    if (!item) {
      sh.getRange(rRow, 1, 1, nc).merge()
        .setValue('—').setBackground(bg).setFontColor(C.GRIS)
        .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
      continue;
    }

    var fg  = item.esRechazado ? C.ROJO
            : item.esAceptado  ? C.VERDE
            : item.esPendiente ? C.MORADO
            : C.AZUL_CORP;
    var lbl = item.tipo === 'PRF' ? '[PROF]' : '[PRES]';
    var txt = item.semaforo + '  ' + lbl
            + '  ' + item.fechaStr
            + '   ' + item.cliente
            + '   ' + _pgnFmtEur_(item.total)
            + '   ' + item.estado;

    sh.getRange(rRow, 1, 1, nc).merge()
      .setValue(txt).setBackground(bg).setFontColor(fg)
      .setFontSize(10).setFontWeight('normal')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.setRowHeight(rRow, 24);
  }

  Logger.log('  ✓ Bloque PRESUPUESTOS/PROFORMAS actualizado.');
}

// ============================================================
// FASE 16.8 — Bloque ALERTAS
// ============================================================

// Recibe los datos ya calculados por los bloques anteriores.
// Solo lee adicionalmente RECORDATORIOS_LOG.
function _pgnCalcAlertas_(datFac, datGas, datTrab, datPresp) {
  Logger.log('  _pgnCalcAlertas_()');

  var result = {
    nFacVencidas:     datFac   ? (datFac.nVencidas        || 0) : 0,
    nFacProximas:     datFac   ? (datFac.nProximasVencer   || 0) : 0,
    nGasSinRevisar:   datGas   ? (datGas.nSinRevisar       || 0) : 0,
    nGasOcr:          datGas   ? (datGas.nOcrFallido       || 0) : 0,
    nGasSinTicket:    datGas   ? (datGas.nSinTicket        || 0) : 0,
    nTrabPendientes:  datTrab  ? (datTrab.nPendientes       || 0) : 0,
    nPrespPend15d:    datPresp ? (datPresp.nPresupPend15d   || 0) : 0,
    nRecordatorios:   0,
    recordatoriosTxt: 'No disponible',
    totalCriticas:    0,
    logs:             []
  };

  // Críticas = las que requieren acción inmediata para el KDEST card
  result.totalCriticas = result.nFacVencidas + result.nGasOcr
                       + result.nTrabPendientes + result.nPrespPend15d;

  // ── RECORDATORIOS_LOG (lectura ligera: solo cuenta filas) ──
  var recSh = pgObtenerHoja_(PGN_SHEET_ALIASES.RECORDATORIOS_LOG);
  if (recSh) {
    result.nRecordatorios   = Math.max(0, recSh.getLastRow() - 1);
    result.recordatoriosTxt = String(result.nRecordatorios);
  } else {
    result.logs.push('ℹ RECORDATORIOS_LOG no encontrada');
  }

  result.logs.push(
    'ALERTAS: ' + result.totalCriticas + ' críticas — ' +
    'facVenc=' + result.nFacVencidas   + ' ocr='      + result.nGasOcr +
    ' trabPend=' + result.nTrabPendientes + ' presp15d=' + result.nPrespPend15d
  );
  Logger.log('    ALERTAS totalCríticas=' + result.totalCriticas +
    ' (venc=' + result.nFacVencidas + ' ocr=' + result.nGasOcr +
    ' trab=' + result.nTrabPendientes + ' presp=' + result.nPrespPend15d + ')');

  return result;
}

function _pgnActualizarBloqueAlertas_(sh, alertas) {
  var R  = PGN_ROWS;
  var C  = PGN_C;
  var nc = PGN_NC;

  // ── KDEST card 4: Alertas activas (fila 6, cols 11-12) ──────
  var nCrit  = alertas.totalCriticas;
  var cardFg = nCrit > 5 ? C.ROJO : nCrit > 0 ? C.NARANJA : C.VERDE;
  sh.getRange(R.KDEST_VAL, 11, 1, 2).merge()
    .setValue(nCrit).setNumberFormat('0').setFontColor(cardFg)
    .setFontSize(26).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── Estado del sistema (fila 2, cols 6-9) ───────────────────
  _pgnSetStatus_(sh, nCrit > 5 ? 'urgente' : nCrit > 0 ? 'revisar' : 'ok');

  // ── Cabecera dinámica de la sección (fila 49) ────────────────
  var secTxt, secBg;
  if (nCrit > 5) {
    secTxt = '🚨  ALERTAS — 🔴 ' + nCrit + ' alertas requieren atención urgente';
    secBg  = C.ROJO;
  } else if (nCrit > 0) {
    secTxt = '🚨  ALERTAS — 🟠 ' + nCrit + ' alerta(s) pendiente(s) de revisión';
    secBg  = C.NARANJA;
  } else {
    secTxt = '🚨  ALERTAS — 🟢 Sin alertas críticas activas';
    secBg  = C.VERDE;
  }
  sh.getRange(R.SEC_ALERTAS, 1, 1, nc).merge()
    .setValue(secTxt).setBackground(secBg).setFontColor(C.BLANCO)
    .setFontSize(12).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');

  // ── 8 filas de alertas (filas 50-57) ────────────────────────
  // Escribe label + color strip + valor para cada fila.
  // sev: r=rojo, n=naranja, m=morado, g=gris
  var sevStrip = { r: C.ROJO_BG, n: C.NARANJA_BG, m: C.MORADO_BG, g: C.GRIS_BG };

  var defs = [
    { lbl: 'Facturas vencidas',
      val: alertas.nFacVencidas,    sev: 'r',
      sem: alertas.nFacVencidas    > 0 ? '🔴' : '🟢' },
    { lbl: 'Facturas próximas a vencer (≤7 días)',
      val: alertas.nFacProximas,    sev: 'n',
      sem: alertas.nFacProximas    > 0 ? '🟠' : '🟢' },
    { lbl: 'Gastos sin revisar',
      val: alertas.nGasSinRevisar,  sev: 'n',
      sem: alertas.nGasSinRevisar  > 0 ? '🟠' : '🟢' },
    { lbl: 'OCR fallidos',
      val: alertas.nGasOcr,         sev: 'r',
      sem: alertas.nGasOcr         > 0 ? '🔴' : '🟢' },
    { lbl: 'Gastos sin ticket / foto',
      val: alertas.nGasSinTicket,   sev: 'n',
      sem: alertas.nGasSinTicket   > 0 ? '🟠' : '🟢' },
    { lbl: 'Trabajos pendientes',
      val: alertas.nTrabPendientes, sev: 'n',
      sem: alertas.nTrabPendientes > 0 ? '🟠' : '🟢' },
    { lbl: 'Presupuestos sin respuesta (>15 días)',
      val: alertas.nPrespPend15d,   sev: 'm',
      sem: alertas.nPrespPend15d   > 0 ? '🟠' : '🟢' },
    { lbl: 'Recordatorios enviados',
      val: alertas.recordatoriosTxt, sev: 'g',
      sem: '⚪', isText: true }
  ];

  defs.forEach(function(def, a) {
    var ra    = R.ALERTAS_START + a;
    var bgRow = a % 2 === 0 ? '#fff8f8' : C.BLANCO;

    // color strip (col 1)
    sh.getRange(ra, 1).setBackground(sevStrip[def.sev]);

    // label (cols 2-10)
    sh.getRange(ra, 2, 1, 9).merge()
      .setValue('  ' + def.lbl)
      .setBackground(bgRow).setFontColor(C.NEGRO).setFontSize(10)
      .setHorizontalAlignment('left').setVerticalAlignment('middle');

    // valor (cols 11-14): color por severidad si > 0, verde si = 0
    var valTxt = def.isText ? String(def.val) : String(def.val);
    var fgVal;
    if (def.isText) {
      fgVal = C.GRIS;
    } else {
      var n = parseInt(def.val, 10) || 0;
      fgVal = n > 0
        ? (def.sev === 'r' ? C.ROJO : def.sev === 'm' ? C.MORADO : C.NARANJA)
        : C.VERDE;
    }
    sh.getRange(ra, 11, 1, 4).merge()
      .setValue(def.sem + '  ' + valTxt)
      .setBackground(bgRow).setFontColor(fgVal)
      .setFontSize(12).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');

    sh.setRowHeight(ra, 28);
  });

  Logger.log('  ✓ Bloque ALERTAS actualizado. Críticas=' + nCrit);
}

// ============================================================
// FASE 16.10 — Helper: inspección de una hoja
// ============================================================

// Analiza una hoja: alias, nº registros, cabeceras duplicadas,
// filas vacías y detección de columnas. No modifica datos.
function _pgnDiagHoja_(aliases, colDefs) {
  var r = {
    found: false, hojaUsada: null,
    nTotal: 0, nVacias: 0,
    dupHeaders: [], colsOk: [], colsMiss: []
  };

  var sh = pgObtenerHoja_(aliases);
  if (!sh) return r;

  r.found     = true;
  r.hojaUsada = sh.getName();
  var lastRow = sh.getLastRow();
  r.nTotal    = Math.max(0, lastRow - 1);
  if (lastRow < 1) return r;

  var lastCol   = sh.getLastColumn();
  var headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  // Cabeceras duplicadas
  var hCounts = {};
  headerRow.forEach(function(h) {
    var k = String(h).trim().toLowerCase();
    if (k) hCounts[k] = (hCounts[k] || 0) + 1;
  });
  Object.keys(hCounts).forEach(function(k) {
    if (hCounts[k] > 1) r.dupHeaders.push('"' + k + '"×' + hCounts[k]);
  });

  // Detección de columnas esperadas
  var hi = _pgnHeaderIdx_(headerRow);
  (colDefs || []).forEach(function(c) {
    if (_pgnGetCol_(hi, c.al) !== -1) r.colsOk.push(c.n);
    else r.colsMiss.push(c.n);
  });

  // Filas con primera columna vacía (sin ID)
  if (r.nTotal > 0) {
    var col1 = sh.getRange(2, 1, r.nTotal, 1).getValues();
    col1.forEach(function(row) {
      if (!String(row[0]).trim()) r.nVacias++;
    });
  }

  return r;
}

// ============================================================
// _diagnosticarPanelGeneral_ — FASE 16.10 reescritura completa
// ============================================================

function _diagnosticarPanelGeneral_() {
  var tz  = Session.getScriptTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  Logger.log('=== diagnosticarPanelGeneral() ' + now + ' ===');

  var ss    = getTargetSS_();
  var warns = [];   // todas las advertencias acumuladas
  var log   = [];   // detalle completo para Logger

  // ── 1. Ejecutar TODOS los cálculos UNA SOLA VEZ ─────────────
  var dFin   = _pgnCalcResumenFinanciero_();
  var dFac   = _pgnCalcFacturas_();
  var dGas   = _pgnCalcGastos_();
  var dTrab  = _pgnCalcTrabajos_();
  var dCli   = _pgnCalcClientes_();
  var dPresp = _pgnCalcPresupuestos_();
  var dAler  = _pgnCalcAlertas_(dFac, dGas, dTrab, dPresp);

  // ── 2. Definir columnas esperadas por hoja ───────────────────
  var CD = {
    FACTURAS: [
      {n:'ID',      al:['id_factura','id factura','idfactura','numero','id']},
      {n:'Fecha',   al:['fecha_emision','fecha emision','fechaemision']},
      {n:'Cliente', al:['nombre_cliente','nombre cliente','cliente','nomcliente']},
      {n:'Estado',  al:['estado']},
      {n:'Total',   al:['total']},
      {n:'Vencto',  al:['fecha_vencimiento','fecha vencimiento','vencimiento']}
    ],
    GASTOS: [
      {n:'ID',      al:['id_gasto','id gasto','idgasto','id']},
      {n:'Fecha',   al:['fecha_gasto','fecha gasto','fechagasto']},
      {n:'Total',   al:['total']},
      {n:'EstPago', al:['estado_pago','estado pago','estadopago']},
      {n:'Revisado',al:['revisado']},
      {n:'EstProc', al:['estado_procesamiento','estado procesamiento','estadoprocesamiento']}
    ],
    TRABAJOS: [
      {n:'Fecha',   al:['fecha_inicio','fecha inicio','fechainicio','fecha_trabajo','fecha trabajo','fecha','timestamp']},
      {n:'Estado',  al:['estado','status','estado_trabajo']},
      {n:'Cliente', al:['nombre_cliente','nomcliente','id_cliente','cliente']},
      {n:'Operario',al:['nombre_operario','operario','id_operario']},
      {n:'Horas',   al:['horas_trabajadas','horas trabajadas','horas','hours']}
    ],
    CLIENTES: [
      {n:'Nombre',  al:['nombre']},
      {n:'Apellidos',al:['apellidos']},
      {n:'Empresa', al:['empresa']},
      {n:'Activo',  al:['activo']}
    ],
    PRESUPUESTOS: [
      {n:'Estado',  al:['estado','estado_presupuesto','estadopresupuesto']},
      {n:'Fecha',   al:['fecha_emision','fecha emision','fecha_creacion','fecha creacion','fecha']},
      {n:'Cliente', al:['nombre_cliente','nomcliente','cliente']},
      {n:'Total',   al:['total','importe','total_presupuesto']}
    ],
    PROFORMAS: [
      {n:'Estado',  al:['estado','estado_proforma','estadoproforma','estado_cobro']},
      {n:'Fecha',   al:['fecha_emision','fecha emision','fecha_creacion','fecha creacion','fecha']},
      {n:'Cliente', al:['nombre_cliente','nomcliente','cliente']},
      {n:'Total',   al:['total','importe','total_proforma']}
    ]
  };

  var hojaSpecs = [
    {key:'CLIENTES',          colDefs: CD.CLIENTES},
    {key:'TRABAJOS',          colDefs: CD.TRABAJOS},
    {key:'FACTURAS',          colDefs: CD.FACTURAS},
    {key:'GASTOS',            colDefs: CD.GASTOS},
    {key:'PROFORMAS',         colDefs: CD.PROFORMAS},
    {key:'PRESUPUESTOS',      colDefs: CD.PRESUPUESTOS},
    {key:'INFORMES',          colDefs: []},
    {key:'RECORDATORIOS_LOG', colDefs: []},
    {key:'ALERTAS_LOG',       colDefs: []}
  ];

  // ── 3. Inspeccionar cada hoja ────────────────────────────────
  var shR = {};
  hojaSpecs.forEach(function(spec) {
    shR[spec.key] = _pgnDiagHoja_(
      PGN_SHEET_ALIASES[spec.key] || [spec.key],
      spec.colDefs
    );
  });

  // ── 4. Construir output ──────────────────────────────────────
  var a = [];
  a.push('=== DIAGNÓSTICO PANEL_GENERAL ===');
  a.push(now);
  a.push('');

  // 4a. Inventario de hojas
  a.push('📋 HOJAS DEL SISTEMA:');
  hojaSpecs.forEach(function(spec) {
    var d = shR[spec.key];
    if (d.found) {
      var alias = d.hojaUsada !== spec.key ? ' → "' + d.hojaUsada + '"' : '';
      var ext = [];
      if (d.nVacias > 0)          ext.push(d.nVacias + ' filas vacías');
      if (d.dupHeaders.length > 0) {
        ext.push('⚠ dup: ' + d.dupHeaders.join(', '));
        warns.push(spec.key + ': cabeceras duplicadas — ' + d.dupHeaders.join(', '));
      }
      var extStr = ext.length ? '  [' + ext.join(' · ') + ']' : '';
      a.push('  ✅ ' + spec.key + alias + ' — ' + d.nTotal + ' reg' + extStr);
    } else {
      a.push('  ⚠ ' + spec.key + ' — no encontrada');
      warns.push(spec.key + ' no encontrada');
    }
  });
  var panelSh = ss.getSheetByName(PGN_SHEET);
  if (panelSh) {
    a.push('  ✅ PANEL_GENERAL — ' + panelSh.getLastRow() + ' filas usadas');
  } else {
    a.push('  ❌ PANEL_GENERAL — no existe (ejecuta "Crear completo")');
    warns.push('PANEL_GENERAL no existe — ejecutar "Crear completo"');
  }

  // 4b. Columnas con problemas
  var colProblems = [];
  hojaSpecs.forEach(function(spec) {
    var d = shR[spec.key];
    if (d.found && d.colsMiss.length > 0) {
      colProblems.push(spec.key + ': ⚠ ' + d.colsMiss.join(', '));
      d.colsMiss.forEach(function(c) {
        warns.push('Col "' + c + '" no detectada en ' + spec.key);
      });
    }
  });
  if (colProblems.length > 0) {
    a.push('');
    a.push('🔍 COLUMNAS NO DETECTADAS:');
    colProblems.forEach(function(p) { a.push('  ' + p); });
  }

  // 4c. Valores actuales por módulo
  a.push('');
  a.push('📊 MÓDULOS — VALORES ACTUALES:');
  a.push('  💰 Financiero:  fact=' + _pgnFmtEur_(dFin.facturadoMes) +
         '  pend.cobro=' + _pgnFmtEur_(dFin.pendienteCobro) +
         '  gas=' + _pgnFmtEur_(dFin.gastosMes) +
         '  benef=' + _pgnFmtEur_(dFin.beneficioMes));
  a.push('  🧾 Facturas:  emit=' + dFac.nEmitidosMes +
         '  pend=' + dFac.nPendientes +
         '  venc=' + dFac.nVencidas +
         '  prox7d=' + dFac.nProximasVencer);
  a.push('  💸 Gastos:  ' + _pgnFmtEur_(dGas.totalMes) +
         '  sinRev=' + dGas.nSinRevisar +
         '  ocr=' + dGas.nOcrFallido +
         '  sinTick=' + dGas.nSinTicket);
  a.push('  🔧 Trabajos:  pend=' + dTrab.nPendientes +
         '  compl=' + dTrab.nCompletados +
         '  h.mes=' + dTrab.horasMes +
         '  realizados=' + dTrab.nRealizadosMes);
  a.push('  👥 Clientes:  activos=' + dCli.nActivos +
         '  conPend=' + dCli.nConPendientes +
         '  sinActiv=' + dCli.nSinActividad +
         (dCli.top5.length > 0 ? '  top=' + _pgnFmtEur_(dCli.top5[0].total) : ''));
  a.push('  📋 Presp:  pend=' + dPresp.nPendientes +
         '  acept=' + dPresp.nAceptados +
         '  rechaz=' + dPresp.nRechazados +
         '  (+15d=' + dPresp.nPresupPend15d + ')' +
         ' | Prof.pend=' + dPresp.nProformasPend);
  a.push('  🚨 Alertas:  críticas=' + dAler.totalCriticas +
         ' (venc=' + dAler.nFacVencidas +
         ' ocr=' + dAler.nGasOcr +
         ' trab=' + dAler.nTrabPendientes +
         ' presp=' + dAler.nPrespPend15d + ')' +
         '  Record.=' + dAler.recordatoriosTxt);
  a.push('  ⚡ Accesos rápidos:  8 checkboxes configurados');

  // 4d. Advertencias operativas
  if (dFac.nVencidas       > 0) warns.push(dFac.nVencidas       + ' factura(s) vencida(s)');
  if (dFac.nProximasVencer > 0) warns.push(dFac.nProximasVencer + ' factura(s) próximas a vencer (≤7d)');
  if (dGas.nOcrFallido     > 0) warns.push(dGas.nOcrFallido     + ' gasto(s) con OCR fallido');
  if (dGas.nSinRevisar     > 0) warns.push(dGas.nSinRevisar     + ' gasto(s) sin revisar');
  if (dGas.nSinTicket      > 0) warns.push(dGas.nSinTicket      + ' gasto(s) sin ticket/foto');
  if (dPresp.nRechazados   > 0) warns.push(dPresp.nRechazados   + ' presupuesto(s) rechazado(s)');
  if (dPresp.nPresupPend15d> 0) warns.push(dPresp.nPresupPend15d+ ' presupuesto(s) sin respuesta >15d');

  a.push('');
  if (warns.length > 0) {
    a.push('⚠ ADVERTENCIAS (' + warns.length + '):');
    warns.forEach(function(w) { a.push('  · ' + w); });
  } else {
    a.push('✅ Sin advertencias');
  }

  // 4e. Estado final
  a.push('');
  var resumen = (warns.length === 0)
    ? '✅ PANEL_GENERAL listo'
    : '⚠ PANEL_GENERAL con advertencias — ' + warns.length + ' incidencia(s)';
  a.push(resumen);

  // ── 5. Logger: detalle completo ──────────────────────────────
  log.push(a.join('\n'));
  log.push('');
  log.push('── DETALLE COLUMNAS POR HOJA ──');
  hojaSpecs.forEach(function(spec) {
    var d = shR[spec.key];
    if (!d.found) { log.push(spec.key + ': NO ENCONTRADA'); return; }
    if (spec.colDefs.length === 0) {
      log.push(spec.key + ': ' + d.nTotal + ' reg (sin check columnas)'); return;
    }
    log.push(spec.key + ': OK=[' + d.colsOk.join(',') + ']' +
             (d.colsMiss.length > 0 ? '  MISS=[' + d.colsMiss.join(',') + ']' : '') +
             (d.dupHeaders.length > 0 ? '  DUP=[' + d.dupHeaders.join(',') + ']' : '') +
             '  vacías=' + d.nVacias);
  });
  log.push('');
  log.push('── LOGS DE MÓDULOS ──');
  [dFin, dFac, dGas, dTrab, dCli, dPresp, dAler].forEach(function(d) {
    if (d && d.logs) d.logs.forEach(function(l) { log.push('  ' + l); });
  });

  var logTxt = log.join('\n');
  Logger.log(logTxt);
  logEvent('INFO', 'diagnosticarPanelGeneral', resumen, warns.slice(0,5).join(' | '));

  SpreadsheetApp.getUi().alert('🔍 Diagnóstico PANEL_GENERAL\n\n' + a.join('\n'));
}

// ============================================================
// onEditPanelGeneral_  —  gestiona checkboxes de accesos rápidos
// ============================================================

function onEditPanelGeneral_(e) {
  if (!e || !e.range) return;
  if (e.range.getSheet().getName() !== PGN_SHEET) return;
  if (e.value !== 'TRUE') return;

  e.range.setValue(false);  // reset inmediato

  var accion = _pgnGetAccion_(e.range.getRow(), e.range.getColumn());
  if (accion) _pgEjecutarAccionRapida_(accion);
}

function _pgnGetAccion_(row, col) {
  var B = PGN_BTNS;
  if (row === B.ACTUALIZAR[0]       && col === B.ACTUALIZAR[1])       return 'ACTUALIZAR';
  if (row === B.IR_PANEL_CLIENTE[0] && col === B.IR_PANEL_CLIENTE[1]) return 'IR_PANEL_CLIENTE';
  if (row === B.IR_PANEL_GASTOS[0]  && col === B.IR_PANEL_GASTOS[1])  return 'IR_PANEL_GASTOS';
  if (row === B.IR_PANEL_FAC[0]     && col === B.IR_PANEL_FAC[1])     return 'IR_PANEL_FAC';
  if (row === B.INF_GASTOS[0]       && col === B.INF_GASTOS[1])       return 'INF_GASTOS';
  if (row === B.IVA_TRIMESTRAL[0]   && col === B.IVA_TRIMESTRAL[1])   return 'IVA_TRIMESTRAL';
  if (row === B.VER_RECORD[0]       && col === B.VER_RECORD[1])       return 'VER_RECORDATORIOS';
  if (row === B.VER_ALERTAS[0]      && col === B.VER_ALERTAS[1])      return 'VER_ALERTAS';
  return null;
}

function _pgEjecutarAccionRapida_(accion) {
  var ui = SpreadsheetApp.getUi();
  switch (accion) {
    case 'ACTUALIZAR':
      _actualizarPanelGeneral_();
      break;
    case 'IR_PANEL_CLIENTE':
      _pgIrAHoja_('PANEL_CLIENTE');
      break;
    case 'IR_PANEL_GASTOS':
      _pgIrAHoja_('PANEL_GASTOS');
      break;
    case 'IR_PANEL_FAC':
      ui.alert('ℹ️ PANEL_FACTURAS no disponible aún. Se implementará en una fase posterior.');
      break;
    case 'INF_GASTOS':
      try { generarInformeGastos(); } catch (e) {
        ui.alert('⚠ Error al generar informe de gastos:\n' + e.message);
      }
      break;
    case 'IVA_TRIMESTRAL':
      try { generarIVATrimestral(); } catch (e) {
        ui.alert('⚠ Error al generar IVA trimestral:\n' + e.message);
      }
      break;
    case 'VER_RECORDATORIOS':
      _pgIrAHoja_('RECORDATORIOS_LOG');
      break;
    case 'VER_ALERTAS':
      _pgIrAHoja_('ALERTAS_LOG');
      break;
  }
}

function _pgIrAHoja_(nombreHoja) {
  var ss = getTargetSS_();
  var sh = ss.getSheetByName(nombreHoja);
  if (!sh) {
    SpreadsheetApp.getUi().alert(
      '⚠ La hoja "' + nombreHoja + '" no existe en este libro.'
    );
    return;
  }
  ss.setActiveSheet(sh);
}

// ============================================================
// FUNCIONES AUXILIARES
// ============================================================

// Devuelve la primera hoja encontrada entre varias alternativas de nombre.
// (Alias: pgObtenerHoja_ es también pública para uso desde otros módulos)

function _pgnSecHdr_(sh, row, nc, texto, bg, fg) {
  sh.getRange(row, 1, 1, nc).merge()
    .setValue('  ' + texto)
    .setBackground(bg || PGN_C.AZUL_CORP).setFontColor(fg || PGN_C.BLANCO)
    .setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(row, 36);
}

function _pgnSplitHdr_(sh, row, nc, lbl1, bg1, fg1, lbl2, bg2, fg2) {
  sh.getRange(row, 1, 1, 7).merge()
    .setValue(lbl1).setBackground(bg1).setFontColor(fg1)
    .setFontSize(12).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(row, 8, 1, nc - 7).merge()
    .setValue(lbl2).setBackground(bg2).setFontColor(fg2)
    .setFontSize(12).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(row, 34);
}

function _pgnSepRow_(sh, row, nc) {
  sh.getRange(row, 1, 1, nc).merge().setValue('').setBackground(PGN_C.SEP_BG);
  sh.setRowHeight(row, 8);
}

// 4 KPI cards en columnas B-C | E-F | H-I | K-L
// valFontSize: tamaño de la fuente del valor (por defecto 18)
// valRowHeight: altura de la fila del valor (por defecto 48)
function _pgnKpiRow_(sh, lblRow, valRow, kpis, valFontSize, valRowHeight) {
  var colStarts = [2, 5, 8, 11];
  var C   = PGN_C;
  var vfs = valFontSize  || 18;
  var vrh = valRowHeight || 48;
  var bStyle = SpreadsheetApp.BorderStyle.SOLID;
  var bColor = '#dadce0';

  // Aplicar fondo neutro a columnas separadoras (1,4,7,10,13,14)
  [1, 4, 7, 10, 13, 14].forEach(function(c) {
    sh.getRange(lblRow, c).setBackground(C.BLANCO);
    sh.getRange(valRow, c).setBackground(C.BLANCO);
  });

  kpis.forEach(function(kpi, i) {
    var col = colStarts[i];
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

    // Borde sutil alrededor de cada tarjeta
    lblR.setBorder(true, true, false, true, null, null, bColor, bStyle);
    valR.setBorder(false, true, true, true, null, null, bColor, bStyle);
  });

  sh.setRowHeight(lblRow, 20);
  sh.setRowHeight(valRow, vrh);
}

// KPI split: izq cols 1-4 (label) + 5-7 (valor) | der cols 8-11 (label) + 12-14 (valor)
function _pgnSplitKpi_(sh, row, bg, kpiL, kpiR) {
  sh.getRange(row, 1, 1, 4).merge()
    .setValue('  ' + kpiL[0])
    .setBackground(bg).setFontColor(PGN_C.NEGRO).setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(row, 5, 1, 3).merge()
    .setValue(kpiL[1])
    .setBackground(bg).setFontColor(kpiL[2]).setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('right').setVerticalAlignment('middle');
  sh.getRange(row, 8, 1, 4).merge()
    .setValue('  ' + kpiR[0])
    .setBackground(bg).setFontColor(PGN_C.NEGRO).setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(row, 12, 1, 3).merge()
    .setValue(kpiR[1])
    .setBackground(bg).setFontColor(kpiR[2]).setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('right').setVerticalAlignment('middle');
  sh.setRowHeight(row, 28);
}

// Actualiza el indicador de estado del sistema (fila 2, cols 6-9)
function _pgnSetStatus_(sh, estado) {
  var mapa = {
    ok:      { txt: '🟢 Sistema OK',       bg: '#e6f4ea', fg: PGN_C.VERDE   },
    revisar: { txt: '🟠 Revisar',           bg: '#fef7e0', fg: PGN_C.NARANJA },
    urgente: { txt: '🔴 Atención urgente', bg: '#fad0cd', fg: PGN_C.ROJO }
  };
  var m = mapa[estado] || mapa.ok;
  sh.getRange(PGN_ROWS.SUBTITULO, 6, 1, 4).merge()
    .setValue(m.txt).setBackground(m.bg).setFontColor(m.fg)
    .setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
}

// Construye un índice {header_lowercase: col_index} desde la primera fila de datos
function _pgnHeaderIdx_(headerRow) {
  var idx = {};
  headerRow.forEach(function(h, i) {
    if (h) idx[String(h).trim().toLowerCase()] = i;
  });
  return idx;
}

// Devuelve el índice de columna para la primera variante encontrada, o -1
function _pgnGetCol_(idx, variants) {
  for (var i = 0; i < variants.length; i++) {
    if (idx.hasOwnProperty(variants[i])) return idx[variants[i]];
  }
  return -1;
}

// Devuelve true si la fecha pertenece al mes y año actuales
function _pgnEsMesActual_(fecha) {
  if (!fecha || !(fecha instanceof Date) || isNaN(fecha.getTime())) return false;
  var now = new Date();
  return fecha.getFullYear() === now.getFullYear() &&
         fecha.getMonth()    === now.getMonth();
}

function _pgnGetEmpresaNombre_() {
  try { return getConfig('empresa_nombre') || 'Gestión Álvaro'; }
  catch (_) { return 'Gestión Álvaro'; }
}

function _pgnMesActual_() {
  var meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
               'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var d = new Date();
  return meses[d.getMonth()] + ' ' + d.getFullYear();
}
