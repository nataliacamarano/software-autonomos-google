// ================================================================
// Panel.gs — Interfaz visual PANEL (Documentos + Presupuestos)
// ================================================================

var P_ = {
  // LEFT: DOCUMENTOS (etiqueta B=2, valor C=3)
  CLIENTE:           [5,  3],
  FECHA_INI:         [6,  3],
  FECHA_FIN:         [7,  3],
  TIPO:              [8,  3],
  PRECIO_HORA:       [9,  3],
  BOTON_BUSCAR:      [39, 3],   // C39 — "Buscar trabajos" checkbox
  JOBS_OK:           [39, 10],  // J39 — flag: trabajos cargados
  BOTON_PREVIEW:     [74, 3],   // C74 — tras tabla trabajos (filas 40-72)
  PREVIEW_OK:        [74, 10],  // col J — flag oculto
  PRV_CLIENTE:       [77, 3],
  PRV_PERIODO:       [78, 3],
  PRV_TRABAJOS:      [79, 3],
  PRV_HORAS:         [80, 3],
  PRV_MATS_TRAB:     [81, 3],
  PRV_MATS_OTROS:    [82, 3],
  PRV_MATS_TOTAL:    [83, 3],
  PRV_BASE:          [84, 3],
  PRV_IVA:           [85, 3],
  PRV_TOTAL:         [86, 3],
  BOTON_GENERAR:     [88, 3],
  ESTADO:            [90, 3],
  URL:               [91, 3],
  PRV_DETALLE:       [87, 3],   // C87 — listado detallado de trabajos (texto multilínea)

  // RIGHT: PRESUPUESTOS (etiqueta E=5, valor F=6 / tabla cols E-H)
  PRE_CLIENTE:       [5,  6],
  PRE_FECHA_EMI:     [6,  6],
  PRE_VALIDEZ:       [7,  6],
  PRE_IVA:           [73, 6],   // F73 — IVA derecho (desplazado +33 por tabla trabajos)
  BOTON_PRE_PREV:    [74, 8],   // H74 — checkbox
  PRE_PREVIEW_OK:    [74, 11],  // col K — flag oculto
  PRE_PRV_CLIENTE:   [77, 6],
  PRE_PRV_LINEAS:    [78, 6],
  PRE_PRV_SUBTOTAL:  [79, 6],
  PRE_PRV_IVA:       [80, 6],
  PRE_PRV_TOTAL:     [81, 6],
  PRE_PRV_FECHA_EMI: [82, 6],
  PRE_PRV_CADUCIDAD: [83, 6],
  BOTON_PRE_GEN:     [88, 8],   // H88 — checkbox
  PRE_ESTADO:        [90, 6],
  PRE_URL:           [91, 6],
};

// Tabla de trabajos disponibles para facturar (filas 40-71, cols B-I)
var JOB_TABLE = {
  SEC_ROW:      40,   // fila cabecera de sección
  HDR_ROW:      41,   // fila de encabezados de columna
  START_ROW:    42,   // primera fila de datos
  MAX_ROWS:     30,   // filas 42-71
  COL_CHECK:    2,    // B — checkbox "Seleccionar"
  COL_INFO:     3,    // C-D merged — ID · Fecha · Operario
  COL_TRAB:     5,    // E — Trabajo realizado / Incidencias
  COL_HORAS:    6,    // F — Horario y horas
  COL_INFORME:  7,    // G — ID del informe (o "—")
  COL_PROFORMA: 8,    // H — ID de la proforma (o "—")
  COL_FACTURA:  9,    // I — ID de la factura (o "—")
  COL_DATA:     10    // J (oculta) — JSON con datos del trabajo
};

// Tabla de líneas libres del presupuesto
var PRE_TABLE = {
  START_ROW:    9,
  ROWS_VISIBLE: 10,  // visibles al inicio (filas 9-18)
  ROWS:         30,  // total pre-allocadas (10 visibles + 20 ocultas, filas 9-38)
  COL_DESC:     5,
  COL_CANT:     6,
  COL_PU:       7,
  COL_IMP:      8
};

var PRE_BTN_ADD = [39, 8];  // checkbox "➕ Añadir concepto" — H39

// Tabla de materiales adicionales (izquierdo, filas 11-18, cols B=2 y C=3)
var MAT_TABLE = {
  START_ROW: 11,  // primera fila de entrada
  ROWS:       8,  // filas 11-18 (siempre visibles)
  COL_DESC:   2,  // col B — descripción
  COL_IMP:    3   // col C — importe €
};

// ── SECCIÓN: FACTURA DESDE PRESUPUESTO (filas 95-130) ────────────
var PF_ = {
  CLIENTE:       [98,  3],   // C98 — dropdown clientes con presupuestos Aceptados
  PRE_ID:        [99,  3],   // C99 — dropdown presupuestos filtrado por cliente+Aceptado
  INFO_CLIENTE:  [98,  10],  // J98 — cliente cargado (fuera de vista, uso interno)
  INFO_TOTAL:    [100, 3],   // C100 — total presup. read-only
  INFO_FECHA:    [101, 3],   // C101 — fecha emisión read-only
  BOTON_CARGAR:  [103, 8],   // H103 — checkbox
  CARGAR_OK:     [103, 10],  // J103 — flag: ID cargado
  BOTON_PREVIEW: [117, 8],   // H117 — checkbox
  PREVIEW_OK:    [117, 11],  // K117 — flag (distinto de K74 que es PRE_PREVIEW_OK)
  PRV_CLIENTE:   [120, 3],
  PRV_LINEAS:    [121, 3],
  PRV_IVA:       [122, 3],
  PRV_PRE_REF:   [120, 6],
  PRV_SUBTOTAL:  [121, 6],
  PRV_TOTAL:     [122, 6],
  BOTON_GENERAR:   [125, 8],   // H125 — checkbox
  ESTADO:          [127, 3],
  URL:             [128, 3],
  BASE_PRESUP_RAW: [100, 10],  // J100 — subtotal bruto del presupuesto (sin IVA), fuera de vista
  TIPO_DOC:        [102, 3],   // C102 — Tipo de documento: Factura / Proforma
};

var PFP_TABLE = {
  START_ROW:  106,
  ROWS:       10,
  COL_DESC:   5, COL_CANT: 6, COL_PU: 7, COL_IMP: 8,
  COL_ORIGEN: 9   // I73:I82 — marca de origen (PRESUPUESTO / vacío=extra), columna oculta
};

// ================================================================
// REPARACIÓN — FÓRMULAS IMPORTE PRESUPUESTO
// ================================================================

/**
 * Aplica en batch la fórmula Importe = Cant × P.U. a toda la columna H
 * de la tabla de presupuesto (filas START_ROW … START_ROW+ROWS-1).
 * Seguro de ejecutar varias veces — no toca datos ni formatos.
 * @param {Sheet} sh  hoja PANEL
 */
function _applyPreImpFormulas_(sh) {
  var formulas = [];
  for (var r = PRE_TABLE.START_ROW; r < PRE_TABLE.START_ROW + PRE_TABLE.ROWS; r++) {
    formulas.push(['=IF(OR(F' + r + '="",G' + r + '=""),"",ROUND(F' + r + '*G' + r + ',2))']);
  }
  sh.getRange(PRE_TABLE.START_ROW, PRE_TABLE.COL_IMP, PRE_TABLE.ROWS, 1).setFormulas(formulas);
}

/**
 * Aplica en batch la fórmula Importe = Cant × P.U. a la columna H
 * de la tabla de FACTURA DESDE PRESUPUESTO (filas 73-82).
 * Se llama desde panelCargarPresupuesto y desde el layout.
 * @param {Sheet} sh  hoja PANEL
 */
function _applyPFPImpFormulas_(sh) {
  var formulas = [];
  for (var r = PFP_TABLE.START_ROW; r < PFP_TABLE.START_ROW + PFP_TABLE.ROWS; r++) {
    formulas.push(['=IF(OR(F' + r + '="",G' + r + '=""),"",ROUND(F' + r + '*G' + r + ',2))']);
  }
  sh.getRange(PFP_TABLE.START_ROW, PFP_TABLE.COL_IMP, PFP_TABLE.ROWS, 1).setFormulas(formulas);
}

/**
 * Wrapper de menú: repara las fórmulas de Importe sin reconstruir el panel.
 * Usar cuando la columna Importe aparece vacía en la primera fila.
 */
function repairPreImpFormulas() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('PANEL');
  if (!sh) { SpreadsheetApp.getUi().alert('No se encontró la hoja PANEL.'); return; }
  _applyPreImpFormulas_(sh);
  _applyPFPImpFormulas_(sh);
  sh.setColumnWidth(10, 2);  // col J — auxiliares
  sh.setColumnWidth(11, 2);  // col K — flags
  SpreadsheetApp.getUi().alert(
    '✅  Reparación aplicada:\n\n' +
    '· Fórmulas Importe tabla Presupuesto (H9:H38)\n' +
    '· Fórmulas Importe tabla Factura/Pres. (H106:H115)\n' +
    '· Columnas J y K ocultadas (datos auxiliares)'
  );
}

// ================================================================
// TEST — PRESUPUESTO DE PRUEBA
// ================================================================

/**
 * Crea un presupuesto de prueba con 3 líneas fijas y lo marca como Aceptado.
 * Pide el nombre del cliente por diálogo (debe existir o no importa para la prueba).
 * Ejecutar desde: Menú ⚙️ Sistema → 🧪 Crear presupuesto de prueba
 */
function testCrearPresupuestoPrueba() {
  var ui  = SpreadsheetApp.getUi();
  var res = ui.prompt(
    'Presupuesto de prueba',
    'Introduce el nombre del cliente (el mismo que tienes en el desplegable C65):',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var cliente = res.getResponseText().trim();
  if (!cliente) { ui.alert('Cliente vacío — cancelado.'); return; }

  var lineItems = [
    { desc: 'Cortar pasto',   cant: 1, pu: 20,  imp: 20  },
    { desc: 'Horas empleado', cant: 2, pu: 60,  imp: 120 },
    { desc: 'Cambiar foco',   cant: 1, pu: 10,  imp: 10  }
  ];

  try {
    var result = generatePresupuestoPDF({
      cliente:      cliente,
      lineItems:    lineItems,
      conIva:       true,
      fechaEmision: new Date(),
      validezDias:  30
    });
    updatePresupuestoEstado(result.id, 'Aceptado');
    ui.alert(
      '✅  Presupuesto creado\n\n' +
      'ID: ' + result.id + '\n' +
      'Cliente: ' + cliente + '\n' +
      'Líneas: cortar pasto 20 € · horas empleado 120 € · cambiar foco 10 €\n' +
      'Base: 150 €  |  IVA 21%: 31,50 €  |  Total: 181,50 €\n' +
      'Estado: Aceptado\n\n' +
      'Siguiente paso:\n' +
      '⑤ Refrescar dropdowns Factura/Pres.\n' +
      'Luego selecciona cliente "' + cliente + '" y presupuesto ' + result.id
    );
  } catch (ex) {
    ui.alert('❌ Error: ' + ex.message);
    logEvent('ERROR', 'testCrearPresupuestoPrueba', ex.message, ex.stack || '');
  }
}

// ================================================================
// TEST — CREAR HOJAS PROFORMAS E INFORMES
// ================================================================

/**
 * Crea las hojas PROFORMAS e INFORMES de forma aislada, sin reconstruir el panel.
 * Ejecutar desde: menú o directamente en el editor de Apps Script.
 * Muestra el resultado en un alert y en los logs.
 */
function testCrearHojasDocumentos() {
  var resultados = [];

  Logger.log('[testCrearHojasDocumentos] INICIO');

  try {
    var shP = _ensureProformasSheet_();
    var msgP = '✅ PROFORMAS — GID=' + shP.getSheetId() + '  filas=' + shP.getLastRow();
    Logger.log(msgP);
    resultados.push(msgP);
  } catch (eP) {
    var msgPErr = '❌ PROFORMAS — ' + eP.message;
    Logger.log(msgPErr + '\n' + (eP.stack || ''));
    resultados.push(msgPErr);
  }

  try {
    var shI = _ensureInformesSheet_();
    var msgI = '✅ INFORMES — GID=' + shI.getSheetId() + '  filas=' + shI.getLastRow();
    Logger.log(msgI);
    resultados.push(msgI);
  } catch (eI) {
    var msgIErr = '❌ INFORMES — ' + eI.message;
    Logger.log(msgIErr + '\n' + (eI.stack || ''));
    resultados.push(msgIErr);
  }

  Logger.log('[testCrearHojasDocumentos] FIN');

  try {
    SpreadsheetApp.getUi().alert('Resultado:\n\n' + resultados.join('\n'));
  } catch (_) {
    Logger.log('Alert no disponible (ejecución desde editor): ' + resultados.join(' | '));
  }
}

// ================================================================
// MENÚ Y TRIGGER
// ================================================================

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('⚙️ Sistema')
      .addItem('① Crear / Reconfigurar PANEL',        'setupPanel')
      .addItem('② Instalar trigger (1ª vez)',          'installPanelTrigger')
      .addItem('③ Actualizar lista de clientes',       'panelRefreshClientes')
      .addItem('④ Instalar módulo Presupuestos',       'installPresupuestosModule')
      .addItem('⑤ Refrescar dropdowns Factura/Pres.', 'panelRefrescarDropdownsFP')
      .addItem('⑥ Añadir columna Lineas_JSON',         'menuEnsureLineasJson')
      .addItem('⑦ PANEL_CLIENTE — Fase 1 (estructura)',   'setupPanelClienteFase1')
      .addItem('⑧ PANEL_CLIENTE — Fase 2 (resumen)',     'setupPanelClienteFase2')
      .addItem('⑨ PANEL_CLIENTE — Fase 3 (secciones)',   'setupPanelClienteFase3')
      .addItem('⑩ PANEL_CLIENTE — Fase 4 (validaciones)','setupPanelClienteFase4')
      .addItem('⑪ PANEL_CLIENTE — Actualizar datos',     'refreshPanelCliente')
      .addItem('⑫ PANEL_CLIENTE — Crear completo',       'setupPanelClienteCompleto')
      .addItem('⑬ PANEL_CLIENTE — Exportar historial PDF','exportarHistorialClientePDF')
      .addItem('⑭ PANEL_CLIENTE — Diagnosticar duplicados DOC_TRABAJOS','diagnosticarDocTrabajosDuplicados')
      .addSeparator()
      .addItem('🧹 Limpiar clientes duplicados',        'limpiarClientesDuplicados')
      .addSeparator()
      .addItem('⚡ Generar documento (manual)',         'panelGenerarDocumento')
      .addSeparator()
      .addItem('🧪 Crear presupuesto de prueba (3 líneas)', 'testCrearPresupuestoPrueba')
      .addSeparator()
      .addItem('🔧 Reparar fórmulas Importe (presup.)', 'repairPreImpFormulas')
      .addItem('🔍 Diagnóstico: PRESUPUESTOS Estado',  'verificarEstadoPresupuestos')
      .addItem('🔧 Reparar filas PRESUPUESTOS',        'runRepairPresupuestosRows')
      .addItem('🔍 Diagnóstico: Panel Factura/Pres.',  'verificarFacturaDesdePresupuestoPanel')
      .addToUi();
  } catch (e) {}
}

function installPanelTrigger() {
  var ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'onPanelEdit'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onPanelEdit').forSpreadsheet(ss).onEdit().create();
  SpreadsheetApp.getUi().alert(
    '✅ Trigger instalado.\nEl botón GENERAR ya funciona automáticamente.'
  );
}

/**
 * Reemplaza el trigger onOpen existente (posiblemente apuntando a una version
 * antigua del script) por uno nuevo que corre desde HEAD.
 * Ejecutar UNA VEZ desde el editor de Apps Script (no desde el menu).
 */
function reinstallOnOpenTrigger() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) {
    Logger.log('ERROR: SpreadsheetApp.getActive() es null. Ejecuta desde el menu del Spreadsheet.');
    return;
  }

  // Listar todos los triggers onOpen antes de borrar
  var all = ScriptApp.getProjectTriggers();
  Logger.log('Triggers actuales: ' + all.length);
  var deleted = 0;
  all.forEach(function(t) {
    Logger.log('  handler=' + t.getHandlerFunction() +
               '  tipo=' + t.getEventType() +
               '  ss=' + t.getTriggerSourceId());
    if (t.getHandlerFunction() === 'onOpen') {
      ScriptApp.deleteTrigger(t);
      deleted++;
      Logger.log('  → BORRADO');
    }
  });
  Logger.log('Triggers onOpen borrados: ' + deleted);

  // Crear trigger nuevo apuntando a HEAD
  ScriptApp.newTrigger('onOpen').forSpreadsheet(ss).onOpen().create();
  Logger.log('Nuevo trigger onOpen creado (HEAD).');

  SpreadsheetApp.getUi().alert(
    'Trigger onOpen reinstalado.\n\n' +
    'Borrados: ' + deleted + ' trigger(s) antiguo(s).\n' +
    'Creado: 1 trigger nuevo (HEAD).\n\n' +
    'Recarga el spreadsheet. El menu deberia mostrar setupPanelCliente en el item 7.'
  );
}

// Trigger instalable: gestiona checkboxes e invalida previews
function onPanelEdit(e) {
  if (!e) return;
  var sheetName = e.range.getSheet().getName();

  // ── PANEL_CLIENTE ─────────────────────────────────────────────
  if (sheetName === 'PANEL_CLIENTE') {
    var pcRow = e.range.getRow();
    var pcCol = e.range.getColumn();
    // Checkbox ACTUALIZAR (C8)
    if (e.value === 'TRUE' && pcRow === PC_.BOTON[0] && pcCol === PC_.BOTON[1]) {
      e.range.setValue(false);
      refreshPanelCliente();
      return;
    }
    // Checkbox ✏ Editar datos de pago — valor personalizado 'EDITAR_PAGO' al marcarse
    if (String(e.value || '') === 'EDITAR_PAGO' && pcCol === 3 && pcRow >= PC_.DATA_START) {
      e.range.setValue('');
      try {
        editarDatosPagoDesdePanel();
      } catch (err) {
        SpreadsheetApp.getUi().alert('Error al abrir el formulario de pago:\n' + err.message);
      }
      return;
    }
    // Checkbox ✓ Cobrada — toggle estado de cobro en FACTURAS/PROFORMAS (col G = 7)
    // El ID del documento está guardado en la nota de la celda checkbox (no en col A)
    if ((e.value === 'TRUE' || e.value === 'FALSE') && pcCol === 7) {
      var cellNote = e.range.getNote().trim();
      if (cellNote.indexOf('FAC:') === 0) {
        actualizarEstadoCobro(cellNote.substring(4), 'Factura',  e.value === 'TRUE', pcRow);
        return;
      }
      if (cellNote.indexOf('PRO:') === 0) {
        actualizarEstadoCobro(cellNote.substring(4), 'Proforma', e.value === 'TRUE', pcRow);
        return;
      }
    }
    // Checkbox 🔍 Ver detalle trabajo (col J = 10)
    // El ID del trabajo está guardado en la nota de la celda checkbox (no en col A)
    if (e.value === 'TRUE' && pcCol === 10) {
      e.range.setValue(false);
      var tId  = e.range.getNote().trim();
      var docA = String(e.range.getSheet().getRange(pcRow, 9).getValue() || '').trim(); // col I = Factura
      if (tId) mostrarDetalleTrabajoPC(tId, docA);
      return;
    }
    return;
  }

  if (sheetName !== 'PANEL') return;
  var row = e.range.getRow();
  var col = e.range.getColumn();
  var sh  = e.range.getSheet();

  // Inputs DOCUMENTOS (col C=3, filas 5-9) → invalida preview
  if (col === 3 && row >= 5 && row <= 9) {
    sh.getRange(P_.PREVIEW_OK[0], P_.PREVIEW_OK[1]).setValue('');
    // Si cambia cliente o fechas → también invalida la lista de trabajos cargados
    if (row <= 7) sh.getRange(P_.JOBS_OK[0], P_.JOBS_OK[1]).setValue('');
    return;
  }
  // Tabla materiales adicionales (filas 11-18, cols B=2 o C=3) → invalida preview
  if (row >= MAT_TABLE.START_ROW && row < MAT_TABLE.START_ROW + MAT_TABLE.ROWS &&
      (col === MAT_TABLE.COL_DESC || col === MAT_TABLE.COL_IMP)) {
    sh.getRange(P_.PREVIEW_OK[0], P_.PREVIEW_OK[1]).setValue('');
    return;
  }
  // Checkboxes tabla trabajos (filas 100-129, col B) → invalida preview
  var jobEnd = JOB_TABLE.START_ROW + JOB_TABLE.MAX_ROWS - 1;
  if (row >= JOB_TABLE.START_ROW && row <= jobEnd && col === JOB_TABLE.COL_CHECK) {
    sh.getRange(P_.PREVIEW_OK[0], P_.PREVIEW_OK[1]).setValue('');
    return;
  }

  // Inputs PRESUPUESTOS → invalida su preview
  var tEnd = PRE_TABLE.START_ROW + PRE_TABLE.ROWS - 1;  // 38
  var invalidaPre = (
    (row === 5  && col === 6)  ||
    (row === 6  && col === 6)  ||
    (row === 7  && col === 6)  ||
    (row >= PRE_TABLE.START_ROW && row <= tEnd && col >= 5 && col <= 7) ||
    (row === P_.PRE_IVA[0] && col === P_.PRE_IVA[1])
  );
  if (invalidaPre) {
    sh.getRange(P_.PRE_PREVIEW_OK[0], P_.PRE_PREVIEW_OK[1]).setValue('');
    return;
  }

  // Sección FACTURA DESDE PRESUPUESTO — invalidaciones
  if (row === PF_.CLIENTE[0] && col === PF_.CLIENTE[1]) {
    // Cliente cambiado → limpiar todo lo que depende de él
    sh.getRange(PF_.PRE_ID[0],       PF_.PRE_ID[1]).clearContent();
    sh.getRange(PF_.INFO_CLIENTE[0],  PF_.INFO_CLIENTE[1]).setValue('');
    sh.getRange(PF_.INFO_TOTAL[0],    PF_.INFO_TOTAL[1]).setValue('');
    sh.getRange(PF_.INFO_FECHA[0],    PF_.INFO_FECHA[1]).setValue('');
    sh.getRange(PF_.CARGAR_OK[0],     PF_.CARGAR_OK[1]).setValue('');
    sh.getRange(PF_.PREVIEW_OK[0],    PF_.PREVIEW_OK[1]).setValue('');
    _fpClearPreview_(sh);
    _fpEstado_(sh, '', '#3c4043');
    var nuevoCliente = String(e.value || '').trim();
    _fpRefreshPresupuestoDropdown_(sh, nuevoCliente);
    return;
  }
  if (row === PF_.PRE_ID[0] && col === PF_.PRE_ID[1]) {
    sh.getRange(PF_.CARGAR_OK[0],  PF_.CARGAR_OK[1]).setValue('');
    sh.getRange(PF_.PREVIEW_OK[0], PF_.PREVIEW_OK[1]).setValue('');
    return;
  }
  var pfEnd = PFP_TABLE.START_ROW + PFP_TABLE.ROWS - 1;  // 82
  if (row >= PFP_TABLE.START_ROW && row <= pfEnd && col >= 5 && col <= 7) {
    sh.getRange(PF_.PREVIEW_OK[0], PF_.PREVIEW_OK[1]).setValue('');
    return;
  }

  if (e.value !== 'TRUE') return;

  if (row === P_.BOTON_BUSCAR[0]   && col === P_.BOTON_BUSCAR[1])   { e.range.setValue(false); panelBuscarTrabajos();                   return; }
  if (row === P_.BOTON_PREVIEW[0]  && col === P_.BOTON_PREVIEW[1])  { e.range.setValue(false); panelVistaPrevia();                      return; }
  if (row === P_.BOTON_GENERAR[0]  && col === P_.BOTON_GENERAR[1])  { e.range.setValue(false); panelGenerarDocumento();                 return; }
  if (row === P_.BOTON_PRE_PREV[0] && col === P_.BOTON_PRE_PREV[1]) { e.range.setValue(false); panelVistaPreviaPresupuesto();           return; }
  if (row === P_.BOTON_PRE_GEN[0]  && col === P_.BOTON_PRE_GEN[1])  { e.range.setValue(false); panelGenerarPresupuesto();               return; }
  if (row === PRE_BTN_ADD[0]       && col === PRE_BTN_ADD[1])        { e.range.setValue(false); panelAgregarConcepto();                  return; }
  if (row === PF_.BOTON_CARGAR[0]  && col === PF_.BOTON_CARGAR[1])  { e.range.setValue(false); panelCargarPresupuesto();                return; }
  if (row === PF_.BOTON_PREVIEW[0] && col === PF_.BOTON_PREVIEW[1]) { e.range.setValue(false); panelVistaPreviaFacturaPresupuesto();    return; }
  if (row === PF_.BOTON_GENERAR[0] && col === PF_.BOTON_GENERAR[1]) { e.range.setValue(false); panelGenerarFacturaPresupuesto();        return; }
}

// ================================================================
// SIMPLE onEdit — garantiza que BUSCAR TRABAJOS funciona sin
// necesidad de instalar el trigger instalable.
// Solo cubre operaciones que usan únicamente SpreadsheetApp.
// Los botones VISTA PREVIA, GENERAR, etc. usan DriveApp/DocumentApp
// y requieren el trigger instalable para evitar efectos duplicados.
// ================================================================

function onEdit(e) {
  if (!e) return;
  var sh = e.range.getSheet();
  if (sh.getName() !== 'PANEL') return;

  var row = e.range.getRow();
  var col = e.range.getColumn();

  // Invalidar preview cuando cambian inputs de Documentos (col C, filas 5-9)
  if (col === 3 && row >= 5 && row <= 9) {
    sh.getRange(P_.PREVIEW_OK[0], P_.PREVIEW_OK[1]).setValue('');
    if (row <= 7) sh.getRange(P_.JOBS_OK[0], P_.JOBS_OK[1]).setValue('');
    return;
  }
  // Invalidar preview cuando cambia un checkbox de la tabla de trabajos
  var jobEnd = JOB_TABLE.START_ROW + JOB_TABLE.MAX_ROWS - 1;
  if (row >= JOB_TABLE.START_ROW && row <= jobEnd && col === JOB_TABLE.COL_CHECK) {
    sh.getRange(P_.PREVIEW_OK[0], P_.PREVIEW_OK[1]).setValue('');
    return;
  }

  if (e.value !== 'TRUE') return;

  // BUSCAR TRABAJOS — solo SpreadsheetApp: seguro con trigger simple
  if (row === P_.BOTON_BUSCAR[0] && col === P_.BOTON_BUSCAR[1]) {
    e.range.setValue(false);
    panelBuscarTrabajos();
    return;
  }
  // VISTA PREVIA (documentos) — solo SpreadsheetApp: seguro con trigger simple
  if (row === P_.BOTON_PREVIEW[0] && col === P_.BOTON_PREVIEW[1]) {
    e.range.setValue(false);
    panelVistaPrevia();
    return;
  }
}

// ================================================================
// SETUP DE LA HOJA PANEL
// ================================================================

function setupPanel() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('PANEL');
  if (!sh) sh = ss.insertSheet('PANEL', 0);
  else {
    sh.clear();
    var _full = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
    _full.breakApart();
    _full.clearDataValidations();
  }

  // ── Columnas ──────────────────────────────────────────────────
  sh.setColumnWidth(1,  10);
  sh.setColumnWidth(2,  185);
  sh.setColumnWidth(3,  235);
  sh.setColumnWidth(4,  25);
  sh.setColumnWidth(5,  155);
  sh.setColumnWidth(6,  70);
  sh.setColumnWidth(7,  90);   // G — Informe
  sh.setColumnWidth(8,  90);   // H — Proforma
  sh.setColumnWidth(9,  90);   // I — Factura
  sh.setColumnWidth(10, 2);    // J — JSON datos (oculta)
  sh.setColumnWidth(11, 1);

  // ── Alturas de filas ──────────────────────────────────────────
  var r;
  sh.setRowHeight(1, 15);
  sh.setRowHeight(2, 54);
  sh.setRowHeight(3, 28);
  sh.setRowHeight(4, 14);
  sh.setRowHeight(5, 42);
  sh.setRowHeight(6, 40);
  sh.setRowHeight(7, 40);
  sh.setRowHeight(8, 34);
  for (r = 9;  r <= 18; r++) sh.setRowHeight(r, 38);  // tabla visible
  for (r = 19; r <= 38; r++) sh.setRowHeight(r, 2);   // tabla oculta
  sh.setRowHeight(39, 36);  // Añadir concepto / BUSCAR TRABAJOS
  // Filas 40-72: tabla TRABAJOS DISPONIBLES (construida por _buildJobsSection_)
  // Filas 73-93: continuación panel tras la tabla de trabajos
  sh.setRowHeight(73, 40);  // IVA (derecho)
  sh.setRowHeight(74, 52);  // Vista Previa
  sh.setRowHeight(75, 10);
  sh.setRowHeight(76, 32);  // cabecera resumen
  for (r = 77; r <= 86; r++) sh.setRowHeight(r, 30);
  sh.setRowHeight(87, 30);  // detalle trabajos
  sh.setRowHeight(88, 52);  // Generar
  sh.setRowHeight(89, 14);
  sh.setRowHeight(90, 36);  // Estado
  sh.setRowHeight(91, 36);  // URL
  sh.setRowHeight(92, 14);
  sh.setRowHeight(93, 28);  // hint

  // ── Divisor central (col D, filas 2-92) ──────────────────────
  sh.getRange(2, 4, 91, 1).setBackground('#dadce0');

  // ── Títulos (fila 2) ─────────────────────────────────────────
  sh.getRange(2, 2, 1, 2).merge()
    .setValue('📄  FACTURAS Y DOCUMENTOS')
    .setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#1a73e8').setFontColor('#ffffff');
  sh.getRange(2, 5, 1, 4).merge()
    .setValue('📋  PRESUPUESTOS')
    .setFontSize(18).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#34a853').setFontColor('#ffffff');

  // ── Subtítulos (fila 3) ───────────────────────────────────────
  sh.getRange(3, 2, 1, 2).merge()
    .setFormula('="Facturas: "&IFERROR(MAX(0,COUNTA(FACTURAS!A:A)-1),0)&"  ·  Proformas: "&IFERROR(MAX(0,COUNTA(PROFORMAS!A:A)-1),0)&"  ·  Informes: "&IFERROR(MAX(0,COUNTA(INFORMES!A:A)-1),0)')
    .setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#e8f0fe').setFontColor('#1a73e8');
  sh.getRange(3, 5, 1, 4).merge()
    .setValue('Creación de Presupuestos')
    .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#e6f4ea').setFontColor('#34a853');

  // ── Fila 5: Cliente ───────────────────────────────────────────
  _lbl_(sh, 5, 2, '👤  Cliente');
  _val_(sh, 5, 3);
  _lbl_(sh, 5, 5, '👤  Cliente');
  sh.getRange(5, 6, 1, 3).merge()
    .setBackground('#ffffff').setVerticalAlignment('middle').setFontSize(11)
    .setBorder(true, true, true, true, false, false, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);

  // ── Fila 6: Fecha Inicio / Fecha de emisión ───────────────────
  _lbl_(sh, 6, 2, '📅  Fecha Inicio');
  _val_(sh, 6, 3);
  _lbl_(sh, 6, 5, '📅  Fecha de emisión');
  sh.getRange(6, 6, 1, 3).merge()
    .setBackground('#fff8e1').setVerticalAlignment('middle').setFontSize(11)
    .setBorder(true, true, true, true, false, false, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);

  // ── Fila 7: Fecha Fin / Validez ───────────────────────────────
  _lbl_(sh, 7, 2, '📅  Fecha Fin');
  _val_(sh, 7, 3);
  _lbl_(sh, 7, 5, '⏳  Validez');
  sh.getRange(7, 6, 1, 3).merge()
    .setBackground('#ffffff').setVerticalAlignment('middle').setFontSize(11)
    .setBorder(true, true, true, true, false, false, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);

  // ── Fila 8: Tipo / Cabecera tabla ─────────────────────────────
  _lbl_(sh, 8, 2, '📄  Tipo de documento');
  _val_(sh, 8, 3);
  var hdrCells = [[5,'Descripción'],[6,'Cant.'],[7,'P.U. €'],[8,'Importe']];
  hdrCells.forEach(function(hc) {
    sh.getRange(8, hc[0])
      .setValue(hc[1]).setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBackground('#34a853');
  });
  sh.getRange(8, 5, 1, 4).setBorder(
    true, true, true, true, true, false, '#1e7e34', SpreadsheetApp.BorderStyle.SOLID
  );

  // ── Fila 9: Precio/hora ───────────────────────────────────────
  _lbl_(sh, 9, 2, '💶  Precio/hora €');
  _val_(sh, 9, 3);

  // ── Fila 10: Cabecera tabla materiales adicionales ────────────
  sh.getRange(10, 2)
    .setValue('🔧  Descripción material')
    .setFontWeight('bold').setFontSize(9).setFontColor('#ffffff')
    .setBackground('#1a73e8').setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(10, 3)
    .setValue('Importe €')
    .setFontWeight('bold').setFontSize(9).setFontColor('#ffffff')
    .setBackground('#1a73e8').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(10, 2, 1, 2).setBorder(
    true, true, true, true, true, false, '#0d47a1', SpreadsheetApp.BorderStyle.SOLID
  );

  // ── Filas 11-38: vacío left (fondo base) ─────────────────────
  for (r = 11; r <= 38; r++) {
    sh.getRange(r, 2).setBackground('#f8f9fa');
    sh.getRange(r, 3).setBackground('#ffffff');
  }

  // ── Filas 11-18: tabla materiales adicionales ─────────────────
  for (r = MAT_TABLE.START_ROW; r < MAT_TABLE.START_ROW + MAT_TABLE.ROWS; r++) {
    sh.getRange(r, MAT_TABLE.COL_DESC)
      .setBackground('#ffffff').setVerticalAlignment('middle').setFontSize(10)
      .setBorder(true, true, true, true, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, MAT_TABLE.COL_IMP)
      .setBackground('#e8f0fe').setVerticalAlignment('middle').setFontSize(10)
      .setHorizontalAlignment('right').setNumberFormat('#,##0.00')
      .setBorder(true, true, true, true, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);
  }

  // ── Tabla right: filas 9-38 (30 líneas) ──────────────────────
  for (r = PRE_TABLE.START_ROW; r <= PRE_TABLE.START_ROW + PRE_TABLE.ROWS - 1; r++) {
    sh.getRange(r, PRE_TABLE.COL_DESC)
      .setBackground('#ffffff').setVerticalAlignment('middle').setFontSize(10)
      .setBorder(true, true, true, true, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, PRE_TABLE.COL_CANT)
      .setBackground('#ffffff').setVerticalAlignment('middle').setFontSize(10)
      .setHorizontalAlignment('right').setNumberFormat('#,##0.##')
      .setBorder(true, true, true, true, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, PRE_TABLE.COL_PU)
      .setBackground('#ffffff').setVerticalAlignment('middle').setFontSize(10)
      .setHorizontalAlignment('right').setNumberFormat('#,##0.00')
      .setBorder(true, true, true, true, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, PRE_TABLE.COL_IMP)
      .setBackground('#f0faf3').setVerticalAlignment('middle').setFontSize(10)
      .setHorizontalAlignment('right').setNumberFormat('#,##0.00').setFontColor('#1e7e34')
      .setBorder(true, true, true, true, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);
  }
  // Fórmulas de importe en batch — garantiza que H9 (primera fila) no quede vacía
  _applyPreImpFormulas_(sh);

  // ── Fila 39: "Buscar trabajos" (left) + "Añadir concepto" (right) ──
  sh.getRange(39, 2)
    .setValue('🔍  BUSCAR TRABAJOS')
    .setFontSize(11).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#f9ab00').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(39, 3).insertCheckboxes().setValue(false)
    .setBackground('#e37400').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(39, 2, 1, 2).setBorder(
    true, true, true, true, true, false, '#b06000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  sh.getRange(39, 5, 1, 3).merge()
    .setValue('➕  Añadir concepto al presupuesto')
    .setFontSize(11).setFontWeight('bold').setFontColor('#34a853')
    .setBackground('#e6f4ea').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(39, 8).insertCheckboxes().setValue(false)
    .setBackground('#c8e6c9').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(39, 5, 1, 4).setBorder(
    true, true, true, true, true, false, '#34a853', SpreadsheetApp.BorderStyle.SOLID
  );

  // ── Fila 73: IVA right (desplazada +33 por tabla trabajos) ───
  sh.getRange(73, 2).setBackground('#f8f9fa');
  sh.getRange(73, 3).setBackground('#ffffff');
  _lbl_(sh, 73, 5, '💰  IVA');
  sh.getRange(73, 6, 1, 3).merge()
    .setBackground('#ffffff').setVerticalAlignment('middle').setFontSize(11)
    .setBorder(true, true, true, true, false, false, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);

  // ── Fila 74: Botones VISTA PREVIA ────────────────────────────
  sh.getRange(74, 2)
    .setValue('🔍  VISTA PREVIA')
    .setFontSize(13).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#1a73e8').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(74, 3).insertCheckboxes().setValue(false)
    .setBackground('#1765cc').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(74, 2, 1, 2).setBorder(
    true, true, true, true, true, false, '#0d47a1', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  sh.getRange(74, 5, 1, 3).merge()
    .setValue('🔍  VISTA PREVIA')
    .setFontSize(13).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#34a853').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(74, 8).insertCheckboxes().setValue(false)
    .setBackground('#2d9140').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(74, 5, 1, 4).setBorder(
    true, true, true, true, true, false, '#1e7e34', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  // ── Fila 76: Cabecera RESUMEN ─────────────────────────────────
  sh.getRange(76, 2, 1, 2).merge()
    .setValue('— RESUMEN PREVIO —')
    .setFontSize(10).setFontWeight('bold').setFontColor('#5f6368')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#f1f3f4');
  sh.getRange(76, 5, 1, 4).merge()
    .setValue('— RESUMEN PREVIO —')
    .setFontSize(10).setFontWeight('bold').setFontColor('#5f6368')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#f1f3f4');

  // ── Filas 77-86: RESUMEN LEFT (10 filas) ─────────────────────
  var leftPrv = [
    [77,'Cliente'],[78,'Periodo'],[79,'Trabajos'],[80,'Horas totales'],
    [81,'Materiales (trabajos)'],[82,'Materiales adicionales'],[83,'Total materiales adic.'],
    [84,'Base imponible'],[85,'IVA / Régimen'],[86,'Total estimado'],
  ];
  leftPrv.forEach(function(c) {
    sh.getRange(c[0], 2)
      .setValue(c[1]).setFontSize(10).setFontColor('#5f6368')
      .setVerticalAlignment('middle').setBackground('#f8f9fa');
    sh.getRange(c[0], 3)
      .setFontSize(10).setVerticalAlignment('middle').setBackground('#ffffff')
      .setBorder(false, false, true, false, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);
  });
  // Fila 82 (Materiales adicionales) admite texto multilínea
  sh.getRange(P_.PRV_MATS_OTROS[0], P_.PRV_MATS_OTROS[1])
    .setWrap(true).setVerticalAlignment('top').setFontSize(9);

  // ── Filas 77-83: RESUMEN RIGHT (7 filas) ─────────────────────
  var rightPrv = [
    [77,'Cliente'],[78,'Nº líneas'],[79,'Subtotal'],
    [80,'IVA'],[81,'Total estimado'],[82,'Fecha emisión'],[83,'Válido hasta'],
  ];
  rightPrv.forEach(function(c) {
    sh.getRange(c[0], 5)
      .setValue(c[1]).setFontSize(10).setFontColor('#5f6368')
      .setVerticalAlignment('middle').setBackground('#f8f9fa');
    sh.getRange(c[0], 6, 1, 3).merge()
      .setFontSize(10).setVerticalAlignment('middle').setBackground('#ffffff')
      .setBorder(false, false, true, false, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);
  });
  for (r = 84; r <= 86; r++) {
    sh.getRange(r, 5).setBackground('#f8f9fa');
    sh.getRange(r, 6, 1, 3).merge().setBackground('#ffffff');
  }

  // ── Fila 87: Detalle de trabajos (expande al hacer vista previa) ──
  sh.getRange(87, 2)
    .setValue('Detalle trabajos').setFontSize(10).setFontColor('#5f6368')
    .setVerticalAlignment('top').setBackground('#f8f9fa');
  sh.getRange(87, 3)
    .setFontSize(9).setVerticalAlignment('top').setBackground('#ffffff').setWrap(true)
    .setBorder(false, false, true, false, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);

  // ── Fila 88: Botones GENERAR ──────────────────────────────────
  sh.getRange(88, 2)
    .setValue('🚀  GENERAR DOCUMENTO')
    .setFontSize(13).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#1a73e8').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(88, 3).insertCheckboxes().setValue(false)
    .setBackground('#1765cc').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(88, 2, 1, 2).setBorder(
    true, true, true, true, true, false, '#0d47a1', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  sh.getRange(88, 5, 1, 3).merge()
    .setValue('📋  GENERAR PRESUPUESTO')
    .setFontSize(13).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#34a853').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(88, 8).insertCheckboxes().setValue(false)
    .setBackground('#2d9140').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(88, 5, 1, 4).setBorder(
    true, true, true, true, true, false, '#1e7e34', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  // ── Filas 90-91: Estado y URL ─────────────────────────────────
  sh.getRange(90, 2).setValue('Estado').setFontWeight('bold').setFontSize(10)
    .setVerticalAlignment('middle').setBackground('#f8f9fa');
  sh.getRange(90, 3).setValue('En espera…').setFontColor('#9aa0a6').setFontSize(11)
    .setVerticalAlignment('middle');
  sh.getRange(91, 2).setValue('Documento PDF').setFontWeight('bold').setFontSize(10)
    .setVerticalAlignment('middle').setBackground('#f8f9fa');
  sh.getRange(91, 3).setFontSize(10).setFontColor('#1a73e8').setVerticalAlignment('middle');

  sh.getRange(90, 5).setValue('Estado').setFontWeight('bold').setFontSize(10)
    .setVerticalAlignment('middle').setBackground('#f8f9fa');
  sh.getRange(90, 6, 1, 3).merge()
    .setValue('En espera…').setFontColor('#9aa0a6').setFontSize(11)
    .setVerticalAlignment('middle');
  sh.getRange(91, 5).setValue('PDF Presupuesto').setFontWeight('bold').setFontSize(10)
    .setVerticalAlignment('middle').setBackground('#f8f9fa');
  sh.getRange(91, 6, 1, 3).merge()
    .setFontSize(10).setFontColor('#1a73e8').setVerticalAlignment('middle');

  // ── Fila 93: Hints ────────────────────────────────────────────
  sh.getRange(93, 2, 1, 2).merge()
    .setValue('💡 Genera siempre la vista previa antes de emitir el documento')
    .setFontSize(9).setFontColor('#9aa0a6')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(93, 5, 1, 4).merge()
    .setValue('💡 Genera siempre la vista previa antes de emitir el presupuesto')
    .setFontSize(9).setFontColor('#9aa0a6')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── Validaciones LEFT ─────────────────────────────────────────
  var clienteNames = _panelGetClienteNames_(ss);
  var clienteVal   = clienteNames.length > 0
    ? SpreadsheetApp.newDataValidation().requireValueInList(clienteNames, true).setAllowInvalid(false).build()
    : null;
  if (clienteVal) sh.getRange(5, 3).setDataValidation(clienteVal);

  var dateVal = SpreadsheetApp.newDataValidation()
    .requireDateOnOrAfter(new Date('1900-01-01')).setAllowInvalid(false).build();
  sh.getRange(6, 3).setDataValidation(dateVal).setNumberFormat('dd/MM/yyyy').setBackground('#e8f0fe');
  sh.getRange(7, 3).setDataValidation(dateVal).setNumberFormat('dd/MM/yyyy').setBackground('#e8f0fe');
  sh.getRange(8, 3).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Informe', 'Factura', 'Proforma'], true).setAllowInvalid(false).build()
  );
  sh.getRange(9, 3).setValue(0).setNumberFormat('#,##0.00');

  // ── Validaciones RIGHT ────────────────────────────────────────
  if (clienteVal) sh.getRange(5, 6).setDataValidation(clienteVal);

  sh.getRange(6, 6)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireDateOnOrAfter(new Date('2000-01-01')).setAllowInvalid(false).build()
    ).setNumberFormat('dd/MM/yyyy').setValue(new Date()).setBackground('#fff8e1');

  sh.getRange(7, 6)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['7 días', '15 días', '30 días', '60 días', '90 días'], true)
        .setAllowInvalid(false).build()
    ).setValue('30 días');

  sh.getRange(73, 6)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Sí', 'No'], true).setAllowInvalid(false).build()
    ).setValue('Sí');

  // ── Sección FACTURA DESDE PRESUPUESTO ─────────────────────────
  Logger.log('[setupPanel] llamando _buildPanelFacturaPresupuesto_...');
  _buildPanelFacturaPresupuesto_(sh);
  Logger.log('[setupPanel] _buildPanelFacturaPresupuesto_ completada.');

  // ── Crear hojas PROFORMAS e INFORMES si no existen ────────────
  Logger.log('[setupPanel] asegurando hojas PROFORMAS e INFORMES...');
  try {
    _ensureProformasSheet_();
    Logger.log('[setupPanel] PROFORMAS OK.');
  } catch (eP) {
    Logger.log('[setupPanel] ERROR creando PROFORMAS: ' + eP.message + '\n' + (eP.stack || ''));
  }
  try {
    _ensureInformesSheet_();
    Logger.log('[setupPanel] INFORMES OK.');
  } catch (eI) {
    Logger.log('[setupPanel] ERROR creando INFORMES: ' + eI.message + '\n' + (eI.stack || ''));
  }
  Logger.log('[setupPanel] hojas auxiliares procesadas.');

  _buildJobsSection_(sh);

  SpreadsheetApp.getUi().alert(
    '✅ PANEL listo.\n\n' +
    '▸ Izquierdo: FACTURAS Y DOCUMENTOS\n' +
    '  · Tabla de materiales adicionales (8 filas, fila 10-18)\n' +
    '  · Botón "🔍 BUSCAR TRABAJOS" para cargar trabajos del período\n' +
    '  · Selecciona manualmente qué trabajos incluir en la factura\n\n' +
    '▸ Derecho:   PRESUPUESTOS\n' +
    '  · 10 líneas de conceptos visibles\n' +
    '  · Hasta 30 líneas con "➕ Añadir concepto"\n\n' +
    'Hojas creadas/verificadas: PROFORMAS, INFORMES\n\n' +
    'Si el módulo Presupuestos no está instalado:\n' +
    '⚙️ Sistema → "④ Instalar módulo Presupuestos"'
  );
}

// ================================================================
// AÑADIR CONCEPTO (revela siguiente fila oculta)
// ================================================================

function panelAgregarConcepto() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('PANEL');
  if (!sh) return;

  var hiddenStart = PRE_TABLE.START_ROW + PRE_TABLE.ROWS_VISIBLE; // 19
  var hiddenEnd   = PRE_TABLE.START_ROW + PRE_TABLE.ROWS - 1;     // 38
  var firstHidden = -1;
  for (var r = hiddenStart; r <= hiddenEnd; r++) {
    if (sh.getRowHeight(r) <= 2) { firstHidden = r; break; }
  }

  if (firstHidden === -1) {
    SpreadsheetApp.getUi().alert('✅ Ya se usan todas las líneas disponibles (30 máximo).');
    return;
  }

  sh.setRowHeight(firstHidden, 38);
  SpreadsheetApp.flush();
}

// ================================================================
// VISTA PREVIA — DOCUMENTOS
// ================================================================

function panelVistaPrevia() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('PANEL');
  if (!sh) return;

  var cliente    = String(sh.getRange(P_.CLIENTE[0],     P_.CLIENTE[1]).getValue()     || '').trim();
  var fechaIni   = sh.getRange(P_.FECHA_INI[0],   P_.FECHA_INI[1]).getDisplayValue().trim();
  var fechaFin   = sh.getRange(P_.FECHA_FIN[0],   P_.FECHA_FIN[1]).getDisplayValue().trim();
  var tipo       = String(sh.getRange(P_.TIPO[0],         P_.TIPO[1]).getValue()         || '').trim();
  var precioHora = Number(sh.getRange(P_.PRECIO_HORA[0], P_.PRECIO_HORA[1]).getValue()) || 0;
  var matsItems  = _getMaterialesAdicionales_(sh);
  var otrosMats  = matsItems.reduce(function(s, m) { return s + m.imp; }, 0);

  var err = [];
  if (!cliente)                                                                          err.push('Selecciona un cliente');
  if (!_parseReportDate_(fechaIni))                                                      err.push('Fecha Inicio inválida (DD/MM/YYYY)');
  if (!_parseReportDate_(fechaFin))                                                      err.push('Fecha Fin inválida (DD/MM/YYYY)');
  if (!tipo)                                                                             err.push('Selecciona tipo de documento');
  if ((tipo === 'Factura' || tipo === 'Proforma') && precioHora <= 0)                   err.push('Precio/hora debe ser mayor que 0');

  var fi = _parseReportDate_(fechaIni);
  var ff = _parseReportDate_(fechaFin);
  if (fi && ff && fi > ff) err.push('Fecha Inicio es posterior a Fecha Fin');

  _panelClearPreview_(sh);
  sh.getRange(P_.PREVIEW_OK[0], P_.PREVIEW_OK[1]).setValue('');

  if (err.length) {
    _panelEstado_(sh, '❌  ' + err.join('  ·  '), '#ea4335');
    return;
  }

  // Verificar que los trabajos han sido cargados mediante "Buscar trabajos"
  var jobsOk = String(sh.getRange(P_.JOBS_OK[0], P_.JOBS_OK[1]).getValue()).trim();
  if (jobsOk !== 'OK') {
    _panelEstado_(sh, '⚠️  Pulsa primero "🔍 BUSCAR TRABAJOS" para cargar los trabajos del período.', '#f9ab00');
    return;
  }
  var trabajos = _getSelectedJobs_(sh);
  if (trabajos.length === 0) {
    _panelEstado_(sh, '⚠️  No hay trabajos seleccionados. Marca al menos uno en la tabla de trabajos.', '#f9ab00');
    return;
  }

  _panelEstado_(sh, '⏳  Calculando vista previa…', '#f9ab00');
  SpreadsheetApp.flush();

  try {
    var totalHoras = 0;
    var totalMats  = 0;
    for (var j = 0; j < trabajos.length; j++) {
      trabajos[j].horas = _calcHoras_(trabajos[j].horaInicio, trabajos[j].horaFin);
      totalHoras += trabajos[j].horas;
      totalMats  += trabajos[j].materiales;
    }

    var totalMatsTotal = totalMats + otrosMats;
    var baseImponible  = totalHoras * precioHora + totalMatsTotal;
    var ivaPct         = parseFloat(getConfig('iva_porcentaje') || '21');
    var ivaImporte     = (tipo === 'Factura') ? baseImponible * ivaPct / 100 : 0;
    var total          = baseImponible + ivaImporte;
    var esFactura      = (tipo === 'Factura' || tipo === 'Proforma');

    sh.getRange(P_.PRV_CLIENTE[0],    P_.PRV_CLIENTE[1]).setValue(cliente);
    sh.getRange(P_.PRV_PERIODO[0],    P_.PRV_PERIODO[1]).setValue(fechaIni + ' → ' + fechaFin);
    sh.getRange(P_.PRV_TRABAJOS[0],   P_.PRV_TRABAJOS[1]).setValue(trabajos.length + ' trabajo(s)');
    sh.getRange(P_.PRV_HORAS[0],      P_.PRV_HORAS[1]).setValue(totalHoras.toFixed(2) + ' h');
    sh.getRange(P_.PRV_MATS_TRAB[0],  P_.PRV_MATS_TRAB[1]).setValue(esFactura ? formatCurrency(totalMats) : '—');
    if (esFactura && matsItems.length > 0) {
      var matsLines = matsItems.map(function(m) {
        return '• ' + m.desc + ' ........ ' + formatCurrency(m.imp);
      });
      sh.getRange(P_.PRV_MATS_OTROS[0], P_.PRV_MATS_OTROS[1])
        .setValue(matsLines.join('\n'))
        .setWrap(true).setFontSize(9).setVerticalAlignment('top');
      sh.setRowHeight(P_.PRV_MATS_OTROS[0], Math.max(30, matsItems.length * 18 + 10));
    } else {
      sh.getRange(P_.PRV_MATS_OTROS[0], P_.PRV_MATS_OTROS[1])
        .setValue(esFactura ? '(ninguno)' : '—').setFontSize(10).setVerticalAlignment('middle');
      sh.setRowHeight(P_.PRV_MATS_OTROS[0], 30);
    }
    sh.getRange(P_.PRV_MATS_TOTAL[0], P_.PRV_MATS_TOTAL[1]).setValue(esFactura ? formatCurrency(totalMatsTotal) : '—');
    sh.getRange(P_.PRV_BASE[0],       P_.PRV_BASE[1]).setValue(esFactura ? formatCurrency(baseImponible) : '—');
    sh.getRange(P_.PRV_IVA[0],        P_.PRV_IVA[1]).setValue(
      tipo === 'Factura'  ? 'IVA ' + ivaPct + '%:  ' + formatCurrency(ivaImporte) :
      tipo === 'Proforma' ? 'Sin IVA (Proforma)' :
                            'Sin IVA (Informe)'
    );
    sh.getRange(P_.PRV_TOTAL[0], P_.PRV_TOTAL[1])
      .setValue(esFactura ? formatCurrency(total) : '—')
      .setFontWeight(esFactura ? 'bold' : 'normal');

    // Detalle de trabajos (fila 54)
    if (trabajos.length > 0) {
      var detalleLines = [];
      for (var di = 0; di < trabajos.length; di++) {
        var dt = trabajos[di];
        var horasStr = dt.horas > 0 ? dt.horas.toFixed(2).replace('.', ',') + ' h' : '0,00 h';
        var incStr   = dt.incidencias ? dt.incidencias.substring(0, 60) : 'Sin incidencias';
        detalleLines.push(
          formatDate(dt.fechaTrabajo) + '  |  ' +
          (dt.trabajoRealizado.substring(0, 60) || '—') + '  |  ' +
          horasStr + '  |  ' + incStr
        );
      }
      sh.getRange(P_.PRV_DETALLE[0], P_.PRV_DETALLE[1])
        .setValue(detalleLines.join('\n'))
        .setWrap(true).setFontSize(9).setVerticalAlignment('top');
      sh.setRowHeight(P_.PRV_DETALLE[0], Math.max(30, detalleLines.length * 15 + 10));
    }

    sh.getRange(P_.PREVIEW_OK[0], P_.PREVIEW_OK[1]).setValue('OK');
    _panelEstado_(sh, '✅  Vista previa lista. Revisa y pulsa GENERAR DOCUMENTO.', '#34a853');

  } catch (ex) {
    _panelEstado_(sh, '❌  Error: ' + ex.message, '#ea4335');
    logEvent('ERROR', 'panelVistaPrevia', ex.message, ex.stack || '');
  }
}

function _panelClearPreview_(sh) {
  var cells = [
    P_.PRV_CLIENTE, P_.PRV_PERIODO, P_.PRV_TRABAJOS, P_.PRV_HORAS,
    P_.PRV_MATS_TRAB, P_.PRV_MATS_OTROS, P_.PRV_MATS_TOTAL,
    P_.PRV_BASE, P_.PRV_IVA, P_.PRV_TOTAL, P_.PRV_DETALLE
  ];
  for (var i = 0; i < cells.length; i++) {
    sh.getRange(cells[i][0], cells[i][1]).setValue('').setFontWeight('normal');
  }
  sh.setRowHeight(P_.PRV_DETALLE[0], 30);
  sh.setRowHeight(P_.PRV_MATS_OTROS[0], 30);
  sh.getRange(P_.PRV_MATS_OTROS[0], P_.PRV_MATS_OTROS[1])
    .setFontSize(10).setVerticalAlignment('middle');
}

// ================================================================
// VISTA PREVIA — PRESUPUESTOS
// ================================================================

function panelVistaPreviaPresupuesto() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('PANEL');
  if (!sh) return;

  var cliente     = String(sh.getRange(P_.PRE_CLIENTE[0],   P_.PRE_CLIENTE[1]).getValue()   || '').trim();
  var fechaEmiVal = sh.getRange(P_.PRE_FECHA_EMI[0], P_.PRE_FECHA_EMI[1]).getValue();
  var validezStr  = String(sh.getRange(P_.PRE_VALIDEZ[0],   P_.PRE_VALIDEZ[1]).getValue()   || '').trim();
  var conIva      = String(sh.getRange(P_.PRE_IVA[0],       P_.PRE_IVA[1]).getValue()).trim() === 'Sí';
  var items       = _getPanelPresupuestoItems_(sh);

  var err = [];
  if (!cliente)                                                         err.push('Selecciona un cliente');
  if (!(fechaEmiVal instanceof Date) || isNaN(fechaEmiVal.getTime()))   err.push('Selecciona una fecha de emisión');
  if (!validezStr)                                                      err.push('Selecciona la validez');
  if (!items.length)                                                    err.push('Introduce al menos una línea en la tabla');

  _panelClearPresupuestoPreview_(sh);
  sh.getRange(P_.PRE_PREVIEW_OK[0], P_.PRE_PREVIEW_OK[1]).setValue('');

  if (err.length) {
    _panelPreEstado_(sh, '❌  ' + err.join('  ·  '), '#ea4335');
    return;
  }

  _panelPreEstado_(sh, '⏳  Calculando…', '#f9ab00');
  SpreadsheetApp.flush();

  try {
    var tz          = Session.getScriptTimeZone();
    var validezDias = parseInt(validezStr, 10) || 30;
    var fechaCaduc  = new Date(fechaEmiVal.getTime());
    fechaCaduc.setDate(fechaCaduc.getDate() + validezDias);

    var ivaPct   = parseFloat(getConfig('iva_porcentaje') || '21');
    var subtotal = 0;
    for (var i = 0; i < items.length; i++) subtotal += items[i].imp;
    subtotal     = Math.round(subtotal * 100) / 100;
    var ivaImp   = conIva ? Math.round(subtotal * ivaPct) / 100 : 0;
    var total    = subtotal + ivaImp;

    sh.getRange(P_.PRE_PRV_CLIENTE[0],   P_.PRE_PRV_CLIENTE[1]).setValue(cliente);
    sh.getRange(P_.PRE_PRV_LINEAS[0],    P_.PRE_PRV_LINEAS[1]).setValue(items.length + ' línea(s)');
    sh.getRange(P_.PRE_PRV_SUBTOTAL[0],  P_.PRE_PRV_SUBTOTAL[1]).setValue(formatCurrency(subtotal));
    sh.getRange(P_.PRE_PRV_IVA[0],       P_.PRE_PRV_IVA[1]).setValue(
      conIva ? 'IVA ' + ivaPct + '%:  ' + formatCurrency(ivaImp) : 'Sin IVA'
    );
    sh.getRange(P_.PRE_PRV_TOTAL[0],     P_.PRE_PRV_TOTAL[1])
      .setValue(formatCurrency(total)).setFontWeight('bold');
    sh.getRange(P_.PRE_PRV_FECHA_EMI[0], P_.PRE_PRV_FECHA_EMI[1])
      .setValue(Utilities.formatDate(fechaEmiVal, tz, 'dd/MM/yyyy'));
    sh.getRange(P_.PRE_PRV_CADUCIDAD[0], P_.PRE_PRV_CADUCIDAD[1])
      .setValue(Utilities.formatDate(fechaCaduc, tz, 'dd/MM/yyyy'));

    sh.getRange(P_.PRE_PREVIEW_OK[0], P_.PRE_PREVIEW_OK[1]).setValue('OK');
    _panelPreEstado_(sh, '✅  Vista previa lista. Revisa y pulsa GENERAR PRESUPUESTO.', '#34a853');

  } catch (ex) {
    _panelPreEstado_(sh, '❌  Error: ' + ex.message, '#ea4335');
    logEvent('ERROR', 'panelVistaPreviaPresupuesto', ex.message, ex.stack || '');
  }
}

function _panelClearPresupuestoPreview_(sh) {
  var cells = [
    P_.PRE_PRV_CLIENTE,   P_.PRE_PRV_LINEAS,
    P_.PRE_PRV_SUBTOTAL,  P_.PRE_PRV_IVA,     P_.PRE_PRV_TOTAL,
    P_.PRE_PRV_FECHA_EMI, P_.PRE_PRV_CADUCIDAD
  ];
  for (var i = 0; i < cells.length; i++) {
    sh.getRange(cells[i][0], cells[i][1]).setValue('').setFontWeight('normal');
  }
}

// ================================================================
// REFRESCAR CLIENTES
// ================================================================

function panelRefreshClientes() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('PANEL');
  if (!sh) { SpreadsheetApp.getUi().alert('Primero crea el PANEL.'); return; }

  var nombres = _panelGetClienteNames_(ss);
  if (!nombres.length) {
    SpreadsheetApp.getUi().alert('⚠️ No se encontraron clientes activos en la hoja CLIENTES.');
    return;
  }

  var val = SpreadsheetApp.newDataValidation()
    .requireValueInList(nombres, true).setAllowInvalid(false).build();
  sh.getRange(P_.CLIENTE[0],     P_.CLIENTE[1]).setDataValidation(val);
  sh.getRange(P_.PRE_CLIENTE[0], P_.PRE_CLIENTE[1]).setDataValidation(val);
  SpreadsheetApp.getUi().alert('✅ Lista de clientes actualizada (' + nombres.length + ' clientes).');
}

// ================================================================
// GENERAR DOCUMENTO DEFINITIVO
// ================================================================

function panelGenerarDocumento() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('PANEL');
  if (!sh) { Logger.log('Hoja PANEL no encontrada'); return; }

  var previewOk = String(sh.getRange(P_.PREVIEW_OK[0], P_.PREVIEW_OK[1]).getValue()).trim();
  if (previewOk !== 'OK') {
    _panelEstado_(sh, '⚠️  Genera primero la vista previa antes de emitir el documento.', '#f9ab00');
    return;
  }

  var cliente        = String(sh.getRange(P_.CLIENTE[0],     P_.CLIENTE[1]).getValue()     || '').trim();
  var fechaIni       = sh.getRange(P_.FECHA_INI[0],   P_.FECHA_INI[1]).getDisplayValue().trim();
  var fechaFin       = sh.getRange(P_.FECHA_FIN[0],   P_.FECHA_FIN[1]).getDisplayValue().trim();
  var tipo           = String(sh.getRange(P_.TIPO[0],         P_.TIPO[1]).getValue()         || '').trim();
  var precioHora     = Number(sh.getRange(P_.PRECIO_HORA[0], P_.PRECIO_HORA[1]).getValue()) || 0;
  var matsItems      = _getMaterialesAdicionales_(sh);
  var selectedJobIds = _getSelectedJobs_(sh).map(function(j) { return j.id; });

  _panelEstado_(sh, '⏳  Procesando…', '#f9ab00');
  sh.getRange(P_.URL[0], P_.URL[1]).setValue('');
  SpreadsheetApp.flush();

  try {
    var url = tipo === 'Informe'
      ? _panelRunInforme_(cliente, fechaIni, fechaFin, selectedJobIds)
      : _panelRunFactura_(cliente, fechaIni, fechaFin, precioHora, tipo, matsItems, selectedJobIds);

    if (url) {
      sh.getRange(P_.URL[0], P_.URL[1]).setValue(url).setFontColor('#1a73e8');
      _panelEstado_(sh, '✅  ' + tipo + ' generada correctamente.', '#34a853');
      sh.getRange(P_.PREVIEW_OK[0], P_.PREVIEW_OK[1]).setValue('');
    } else {
      _panelEstado_(sh, '⚠️  No se encontraron trabajos para ese período.', '#f9ab00');
    }
  } catch (ex) {
    _panelEstado_(sh, '❌  Error: ' + ex.message, '#ea4335');
    logEvent('ERROR', 'panelGenerarDocumento', ex.message, ex.stack || '');
  }
}

// ================================================================
// GENERAR PRESUPUESTO
// ================================================================

function panelGenerarPresupuesto() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('PANEL');
  if (!sh) return;

  var previewOk = String(sh.getRange(P_.PRE_PREVIEW_OK[0], P_.PRE_PREVIEW_OK[1]).getValue()).trim();
  if (previewOk !== 'OK') {
    _panelPreEstado_(sh, '⚠️  Genera primero la vista previa antes de emitir el presupuesto.', '#f9ab00');
    return;
  }

  var cliente     = String(sh.getRange(P_.PRE_CLIENTE[0],   P_.PRE_CLIENTE[1]).getValue()   || '').trim();
  var fechaEmiVal = sh.getRange(P_.PRE_FECHA_EMI[0], P_.PRE_FECHA_EMI[1]).getValue();
  var validezStr  = String(sh.getRange(P_.PRE_VALIDEZ[0],   P_.PRE_VALIDEZ[1]).getValue()   || '').trim();
  var conIva      = String(sh.getRange(P_.PRE_IVA[0],       P_.PRE_IVA[1]).getValue()).trim() === 'Sí';
  var items       = _getPanelPresupuestoItems_(sh);

  var fechaEmision = (fechaEmiVal instanceof Date && !isNaN(fechaEmiVal.getTime()))
    ? fechaEmiVal : new Date();
  var validezDias  = parseInt(validezStr, 10) || 30;

  _panelPreEstado_(sh, '⏳  Generando presupuesto…', '#f9ab00');
  sh.getRange(P_.PRE_URL[0], P_.PRE_URL[1]).setValue('');
  SpreadsheetApp.flush();

  try {
    var result = generatePresupuestoPDF({
      cliente:      cliente,
      lineItems:    items,
      conIva:       conIva,
      fechaEmision: fechaEmision,
      validezDias:  validezDias
    });

    sh.getRange(P_.PRE_URL[0], P_.PRE_URL[1]).setValue(result.url).setFontColor('#1a73e8');
    _panelPreEstado_(sh, '✅  ' + result.id + ' generado correctamente.', '#34a853');
    sh.getRange(P_.PRE_PREVIEW_OK[0], P_.PRE_PREVIEW_OK[1]).setValue('');
    // Refresco automático — el nuevo presupuesto aparece en el dropdown de Factura sin acción manual
    _fpRefreshClienteDropdown_(sh);

  } catch (ex) {
    _panelPreEstado_(sh, '❌  Error: ' + ex.message, '#ea4335');
    logEvent('ERROR', 'panelGenerarPresupuesto', ex.message, ex.stack || '');
  }
}

// ================================================================
// WRAPPERS DOCUMENTOS
// ================================================================

function _panelRunInforme_(cliente, fechaIni, fechaFin, selectedJobIds) {
  setConfig('cliente_reporte',      cliente);
  setConfig('fecha_inicio_reporte', fechaIni);
  setConfig('fecha_fin_reporte',    fechaFin);
  setConfig('ids_trabajos_reporte', Array.isArray(selectedJobIds) && selectedJobIds.length
    ? selectedJobIds.join(',') : '');

  var infSheet   = _ensureInformesSheet_();
  var rowsBefore = infSheet.getLastRow();

  generateClientReportPDF();

  var newRow = infSheet.getLastRow();
  if (newRow > rowsBefore) {
    return String(infSheet.getRange(newRow, 9).getValue() || '') || null;  // URL_PDF = col 9
  }
  return null;
}

function _panelRunFactura_(cliente, fechaIni, fechaFin, precioHora, tipo, matsItems, selectedJobIds) {
  var matsArray       = Array.isArray(matsItems) ? matsItems : [];
  var otrosMateriales = matsArray.reduce(function(s, m) { return s + (Number(m.imp) || 0); }, 0);
  setConfig('cliente_factura',        cliente);
  setConfig('fecha_inicio_factura',   fechaIni);
  setConfig('fecha_fin_factura',      fechaFin);
  setConfig('precio_hora',            String(precioHora));
  setConfig('otros_materiales',       String(otrosMateriales));
  setConfig('materiales_adicionales', JSON.stringify(matsArray));

  var tipoEfectivo = tipo || 'Factura';
  var targetSheet, urlCol;
  if (tipoEfectivo === 'Proforma') {
    targetSheet = _ensureProformasSheet_();
    urlCol = 15;  // URL_PDF en PROFORMAS = col 15
  } else {
    targetSheet = getOrCreateSheet('FACTURAS');
    urlCol = 19;  // URL_PDF en FACTURAS = col 19
  }
  var rowsBefore = targetSheet.getLastRow();

  generateInvoicePDF(tipoEfectivo, selectedJobIds);

  var newRow = targetSheet.getLastRow();
  if (newRow > rowsBefore) {
    return String(targetSheet.getRange(newRow, urlCol).getValue() || '');
  }
  return null;
}

// ================================================================
// HELPERS PRIVADOS
// ================================================================

function _getMaterialesAdicionales_(sh) {
  var data = sh.getRange(MAT_TABLE.START_ROW, MAT_TABLE.COL_DESC, MAT_TABLE.ROWS, 2).getValues();
  var items = [];
  for (var i = 0; i < data.length; i++) {
    var desc = String(data[i][0] || '').trim();
    var imp  = Number(data[i][1]) || 0;
    if (!desc && imp === 0) continue;
    items.push({ desc: desc, imp: imp });
  }
  return items;
}

function _getPanelPresupuestoItems_(sh) {
  var items = [];
  // Lectura batch de las 30 filas (cols E, F, G)
  var data = sh.getRange(PRE_TABLE.START_ROW, PRE_TABLE.COL_DESC, PRE_TABLE.ROWS, 3).getValues();
  for (var i = 0; i < data.length; i++) {
    var desc = String(data[i][0] || '').trim();
    var cant = Number(data[i][1]) || 0;
    var pu   = Number(data[i][2]) || 0;
    if (!desc && cant === 0 && pu === 0) continue;
    items.push({ desc: desc, cant: cant, pu: pu, imp: Math.round(cant * pu * 100) / 100 });
  }
  return items;
}

function _panelGetClienteNames_(ss) {
  try {
    var sh = (ss || SpreadsheetApp.getActive()).getSheetByName('CLIENTES');
    if (!sh || sh.getLastRow() < 2) return [];
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, COL_CLI.ACTIVO).getValues();
    var seen = {}, res = [];
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][COL_CLI.ACTIVO - 1]).toLowerCase() !== 'true') continue;
      var nombre    = String(data[i][COL_CLI.NOMBRE    - 1] || '').trim();
      var apellidos = String(data[i][COL_CLI.APELLIDOS - 1] || '').trim();
      var empresa   = String(data[i][COL_CLI.EMPRESA   - 1] || '').trim();
      var label     = (nombre + ' ' + apellidos).trim();
      if (empresa) label += ' (' + empresa + ')';
      if (!label) continue;
      var key = _normalizeClientLabel_(label);
      if (!seen[key]) { seen[key] = true; res.push(label); }
    }
    return res.sort();
  } catch (e) {
    return [];
  }
}

function _panelEstado_(sh, msg, color) {
  sh.getRange(P_.ESTADO[0], P_.ESTADO[1]).setValue(msg).setFontColor(color || '#3c4043');
}

function _panelPreEstado_(sh, msg, color) {
  sh.getRange(P_.PRE_ESTADO[0], P_.PRE_ESTADO[1]).setValue(msg).setFontColor(color || '#3c4043');
}

function _lbl_(sh, row, col, text) {
  sh.getRange(row, col)
    .setValue(text).setFontWeight('bold').setFontSize(11)
    .setVerticalAlignment('middle').setBackground('#f8f9fa').setFontColor('#3c4043');
}

function _val_(sh, row, col) {
  sh.getRange(row, col)
    .setBackground('#ffffff').setVerticalAlignment('middle').setFontSize(11)
    .setBorder(true, true, true, true, false, false, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
}

// ================================================================
// SECCIÓN FACTURA DESDE PRESUPUESTO — LAYOUT
// ================================================================

function _buildPanelFacturaPresupuesto_(sh) {
  var r;

  // ── Alturas de fila ───────────────────────────────────────────
  sh.setRowHeight(94, 14);
  sh.setRowHeight(95, 8);
  sh.setRowHeight(96, 54);
  sh.setRowHeight(97, 24);
  sh.setRowHeight(98, 42);
  sh.setRowHeight(99, 42);
  sh.setRowHeight(100, 34);
  sh.setRowHeight(101, 34);
  sh.setRowHeight(102, 38);  // Tipo de documento (Factura / Proforma)
  sh.setRowHeight(103, 52);
  sh.setRowHeight(104, 10);
  sh.setRowHeight(105, 34);
  for (r = 106; r <= 115; r++) sh.setRowHeight(r, 38);
  sh.setRowHeight(116, 40);
  sh.setRowHeight(117, 52);
  sh.setRowHeight(118, 10);
  sh.setRowHeight(119, 32);
  for (r = 120; r <= 123; r++) sh.setRowHeight(r, 30);
  sh.setRowHeight(124, 14);
  sh.setRowHeight(125, 52);
  sh.setRowHeight(126, 14);
  sh.setRowHeight(127, 36);
  sh.setRowHeight(128, 36);
  sh.setRowHeight(129, 14);
  sh.setRowHeight(130, 28);

  // ── Línea separadora (fila 95) ────────────────────────────────
  sh.getRange(95, 2, 1, 7).setBackground('#bdc1c6');

  // ── Cabecera (fila 96) ────────────────────────────────────────
  sh.getRange(96, 2, 1, 7).merge()
    .setValue('🧾  FACTURA DESDE PRESUPUESTO')
    .setFontSize(18).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#e37400').setFontColor('#ffffff');

  // ── Subtítulo (fila 97) ───────────────────────────────────────
  sh.getRange(97, 2, 1, 7).merge()
    .setValue('Convierte un presupuesto aceptado en factura')
    .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#fef7e0').setFontColor('#e37400');

  // ── Fila 98: selector de cliente ─────────────────────────────
  _lbl_(sh, 98, 2, '👤  Cliente');
  sh.getRange(98, 3, 1, 6).merge()
    .setBackground('#ffffff').setVerticalAlignment('middle').setFontSize(11)
    .setBorder(true, true, true, true, false, false, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
  Logger.log('[_buildPanelFacturaPresupuesto_] llamando _fpRefreshClienteDropdown_...');
  _fpRefreshClienteDropdown_(sh);
  Logger.log('[_buildPanelFacturaPresupuesto_] _fpRefreshClienteDropdown_ completada.');

  // ── Fila 99: selector de presupuesto (filtrado por cliente) ──
  _lbl_(sh, 99, 2, '📋  Presupuesto');
  sh.getRange(99, 3, 1, 6).merge()
    .setBackground('#ffffff').setVerticalAlignment('middle').setFontSize(11)
    .setBorder(true, true, true, true, false, false, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
  // El dropdown de presupuestos se rellena cuando el usuario elige un cliente

  // ── Filas 100-101: info read-only (se rellenan al cargar) ────
  var infoRows = [[100,'📊  Base presup.'],[101,'📅  Fecha emisión']];
  infoRows.forEach(function(c) {
    _lbl_(sh, c[0], 2, c[1]);
    sh.getRange(c[0], 3, 1, 6).merge()
      .setBackground('#f8f9fa').setVerticalAlignment('middle').setFontSize(11)
      .setFontColor('#5f6368')
      .setBorder(true, true, true, true, false, false, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
  });
  // Fila 102: Tipo de documento (Factura / Proforma)
  _lbl_(sh, 102, 2, '📄  Tipo de documento');
  sh.getRange(102, 3, 1, 6).merge()
    .setBackground('#ffffff').setVerticalAlignment('middle').setFontSize(11)
    .setBorder(true, true, true, true, false, false, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(PF_.TIPO_DOC[0], PF_.TIPO_DOC[1])
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Factura', 'Proforma'], true).setAllowInvalid(false).build()
    ).setValue('Factura');

  // ── Fila 103: botón CARGAR ────────────────────────────────────
  sh.getRange(103, 2, 1, 6).merge()
    .setValue('⬇️  CARGAR PRESUPUESTO EN TABLA')
    .setFontSize(13).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#e37400').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(103, 8).clearContent().clearDataValidations();
  sh.getRange(103, 8).insertCheckboxes().setValue(false)
    .setBackground('#c75000').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(103, 2, 1, 7).setBorder(
    true, true, true, true, true, false, '#b06000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  // ── Fila 105: cabecera tabla ──────────────────────────────────
  sh.getRange(105, 2, 1, 1)
    .setValue('📌  Líneas de factura')
    .setFontSize(10).setFontWeight('bold').setFontColor('#e37400')
    .setVerticalAlignment('middle').setBackground('#f8f9fa');
  sh.getRange(105, 3).setBackground('#f8f9fa');
  var hdrCols = [[5,'Descripción'],[6,'Cant.'],[7,'P.U. €'],[8,'Importe']];
  hdrCols.forEach(function(hc) {
    sh.getRange(105, hc[0])
      .setValue(hc[1]).setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBackground('#e37400');
  });
  sh.getRange(105, 5, 1, 4).setBorder(
    true, true, true, true, true, false, '#b06000', SpreadsheetApp.BorderStyle.SOLID
  );

  // Cols I, J, K — ocultar: contienen datos auxiliares/flags no visibles para el usuario
  sh.setColumnWidth(PFP_TABLE.COL_ORIGEN, 2);  // col I — origen (PRESUPUESTO / extra)
  sh.setColumnWidth(10, 2);                     // col J — INFO_CLIENTE, BASE_PRESUP_RAW, CARGAR_OK
  sh.setColumnWidth(11, 2);                     // col K — PREVIEW_OK (docs + presupuesto)

  // ── Filas 106-115: tabla editable ─────────────────────────────
  for (r = PFP_TABLE.START_ROW; r < PFP_TABLE.START_ROW + PFP_TABLE.ROWS; r++) {
    sh.getRange(r, 2).setBackground('#f8f9fa');
    sh.getRange(r, 3).setBackground('#ffffff');
    // Cols E-G: fondo blanco, texto negro (reset de colores de carga anterior)
    sh.getRange(r, 5)
      .setBackground('#ffffff').setFontColor('#000000')
      .setVerticalAlignment('middle').setFontSize(10)
      .setBorder(true, true, true, true, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, 6)
      .setBackground('#ffffff').setFontColor('#000000')
      .setVerticalAlignment('middle').setFontSize(10)
      .setHorizontalAlignment('right').setNumberFormat('#,##0.##')
      .setBorder(true, true, true, true, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, 7)
      .setBackground('#ffffff').setFontColor('#000000')
      .setVerticalAlignment('middle').setFontSize(10)
      .setHorizontalAlignment('right').setNumberFormat('#,##0.00')
      .setBorder(true, true, true, true, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, 8)
      .setBackground('#fef7e0').setVerticalAlignment('middle').setFontSize(10)
      .setHorizontalAlignment('right').setNumberFormat('#,##0.00').setFontColor('#b06000')
      .setBorder(true, true, true, true, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);
    // Col I: origen oculto — limpiar y formatear como texto
    sh.getRange(r, PFP_TABLE.COL_ORIGEN).clearContent().setNumberFormat('@');
  }
  _applyPFPImpFormulas_(sh);  // batch — garantiza H106:H115 completo desde la primera fila

  // ── Fila 116: IVA ─────────────────────────────────────────────
  sh.getRange(116, 2).setBackground('#f8f9fa');
  sh.getRange(116, 3).setBackground('#ffffff');
  _lbl_(sh, 116, 5, '💰  IVA');
  sh.getRange(116, 6, 1, 3).merge()
    .setValue('Factura: 21%  ·  Proforma: sin IVA')
    .setBackground('#f8f9fa').setVerticalAlignment('middle').setFontSize(9).setFontColor('#5f6368')
    .setBorder(true, true, true, true, false, false, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);

  // ── Fila 117: botón VISTA PREVIA ──────────────────────────────
  sh.getRange(117, 2, 1, 6).merge()
    .setValue('🔍  VISTA PREVIA')
    .setFontSize(13).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#e37400').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(117, 8).clearContent().clearDataValidations();
  sh.getRange(117, 8).insertCheckboxes().setValue(false)
    .setBackground('#c75000').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(117, 2, 1, 7).setBorder(
    true, true, true, true, true, false, '#b06000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  // ── Fila 119: cabecera RESUMEN ────────────────────────────────
  sh.getRange(119, 2, 1, 7).merge()
    .setValue('— RESUMEN PREVIO —')
    .setFontSize(10).setFontWeight('bold').setFontColor('#5f6368')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#f1f3f4');

  // ── Filas 120-123: resumen (2 columnas) ───────────────────────
  var prvDef = [
    [120, 'Cliente',   'Presupuesto ref.'],
    [121, 'Nº líneas', 'Subtotal'],
    [122, 'IVA 21%',   'Total factura'],
    [123, '',          ''],
  ];
  prvDef.forEach(function(c) {
    if (c[1]) {
      sh.getRange(c[0], 2).setValue(c[1]).setFontSize(10).setFontColor('#5f6368')
        .setVerticalAlignment('middle').setBackground('#f8f9fa');
    } else {
      sh.getRange(c[0], 2).setBackground('#f8f9fa');
    }
    sh.getRange(c[0], 3).setFontSize(10).setVerticalAlignment('middle').setBackground('#ffffff')
      .setBorder(false, false, true, false, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);
    if (c[2]) {
      sh.getRange(c[0], 5).setValue(c[2]).setFontSize(10).setFontColor('#5f6368')
        .setVerticalAlignment('middle').setBackground('#f8f9fa');
    } else {
      sh.getRange(c[0], 5).setBackground('#f8f9fa');
    }
    sh.getRange(c[0], 6, 1, 3).merge()
      .setFontSize(10).setVerticalAlignment('middle').setBackground('#ffffff')
      .setBorder(false, false, true, false, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);
  });

  // ── Fila 125: botón GENERAR FACTURA ──────────────────────────
  sh.getRange(125, 2, 1, 6).merge()
    .setValue('🧾  GENERAR DOCUMENTO')
    .setFontSize(13).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#e37400').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(125, 8).clearContent().clearDataValidations();
  sh.getRange(125, 8).insertCheckboxes().setValue(false)
    .setBackground('#c75000').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(125, 2, 1, 7).setBorder(
    true, true, true, true, true, false, '#b06000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  // ── Filas 127-128: Estado y URL ───────────────────────────────
  sh.getRange(127, 2).setValue('Estado').setFontWeight('bold').setFontSize(10)
    .setVerticalAlignment('middle').setBackground('#f8f9fa');
  sh.getRange(127, 3, 1, 6).merge()
    .setValue('En espera…').setFontColor('#9aa0a6').setFontSize(11).setVerticalAlignment('middle');
  sh.getRange(128, 2).setValue('PDF Factura').setFontWeight('bold').setFontSize(10)
    .setVerticalAlignment('middle').setBackground('#f8f9fa');
  sh.getRange(128, 3, 1, 6).merge()
    .setFontSize(10).setFontColor('#1a73e8').setVerticalAlignment('middle');

  // ── Fila 130: hint ────────────────────────────────────────────
  sh.getRange(130, 2, 1, 7).merge()
    .setValue('💡 Carga el presupuesto, edita si es necesario, genera vista previa y emite la factura')
    .setFontSize(9).setFontColor('#9aa0a6')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
}

/**
 * Rellena el dropdown C65 con los clientes únicos que tienen
 * al menos un presupuesto con Estado = 'Aceptado'.
 * Detecta columnas por nombre de cabecera (tolerante a estructuras distintas).
 */
function _fpRefreshClienteDropdown_(sh) {
  Logger.log('══════════════════════════════════════════════');
  Logger.log('[_fpRefreshClienteDropdown_] INICIO');

  try {
    // ── 1. Hoja PRESUPUESTOS ───────────────────────────────────
    var ss       = SpreadsheetApp.getActive();
    var preSheet = ss.getSheetByName('PRESUPUESTOS');
    Logger.log('  SS id       : ' + ss.getId());
    Logger.log('  preSheet    : ' + (preSheet ? '"' + preSheet.getName() + '" lastRow=' + preSheet.getLastRow() + ' lastCol=' + preSheet.getLastColumn() : 'NULL — hoja no encontrada'));
    if (!preSheet || preSheet.getLastRow() < 2) {
      Logger.log('  → SALIDA: sin hoja o sin filas de datos.');
      return;
    }

    // ── 2. Cabeceras reales ────────────────────────────────────
    var hdrs       = preSheet.getRange(1, 1, 1, preSheet.getLastColumn()).getValues()[0];
    var colCliente = -1, colEstado = -1;
    Logger.log('  Cabeceras detectadas:');
    for (var h = 0; h < hdrs.length; h++) {
      var hn = String(hdrs[h]).trim();
      Logger.log('    col ' + (h + 1) + ': "' + hn + '"');
      if (hn === 'Cliente') colCliente = h + 1;
      if (hn === 'Estado')  colEstado  = h + 1;
    }
    Logger.log('  colCliente=' + colCliente + '  colEstado=' + colEstado);
    if (colCliente < 1 || colEstado < 1) {
      Logger.log('  → SALIDA: columna "Cliente" o "Estado" no encontrada en cabeceras.');
      return;
    }

    // ── 3. Leer filas ─────────────────────────────────────────
    var colMax = Math.max(colCliente, colEstado);
    var rows   = preSheet.getRange(2, 1, preSheet.getLastRow() - 1, colMax).getValues();
    Logger.log('  Filas de datos leídas: ' + rows.length);

    var seen = {}, clientes = [];
    for (var i = 0; i < rows.length; i++) {
      var cli    = String(rows[i][colCliente - 1] || '').trim();
      var estado = String(rows[i][colEstado  - 1] || '').trim();
      // Log de caracteres para detectar espacios/invisibles
      var estadoHex = estado.split('').map(function(c) { return c.charCodeAt(0); }).join(',');
      Logger.log('  fila ' + (i + 2) + ': cli="' + cli + '" | estado="' + estado + '" | charCodes=[' + estadoHex + '] | esAceptado=' + (estado === 'Aceptado'));
      if (cli && estado === 'Aceptado' && !seen[cli]) {
        seen[cli] = true;
        clientes.push(cli);
      }
    }

    // ── 4. Resultado ──────────────────────────────────────────
    Logger.log('  Clientes aceptados encontrados: ' + (clientes.length ? clientes.join(', ') : '(ninguno)'));
    if (!clientes.length) {
      Logger.log('  → SALIDA: ningún presupuesto con Estado="Aceptado". Dropdown C65 no se aplica.');
      return;
    }

    // ── 5. Aplicar validación ─────────────────────────────────
    var targetRow = PF_.CLIENTE[0], targetCol = PF_.CLIENTE[1];
    Logger.log('  PF_.CLIENTE = [' + targetRow + ', ' + targetCol + '] → celda ' + _col2letter_(targetCol) + targetRow);
    Logger.log('  Aplicando dropdown en C' + targetRow);
    sh.getRange(targetRow, targetCol).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(clientes, true).setAllowInvalid(false).build()
    );
    Logger.log('  setDataValidation() ejecutado sobre ' + _col2letter_(targetCol) + targetRow + ' ✓');

  } catch (e) {
    Logger.log('  ❌ ERROR: ' + e.message);
    Logger.log('  Stack: ' + e.stack);
  }
  Logger.log('[_fpRefreshClienteDropdown_] FIN');
  Logger.log('══════════════════════════════════════════════');
}

function _col2letter_(col) {
  var s = '';
  while (col > 0) {
    col--;
    s = String.fromCharCode(65 + (col % 26)) + s;
    col = Math.floor(col / 26);
  }
  return s;
}

/**
 * Rellena el dropdown C66 con los IDs de presupuestos del cliente dado
 * que tengan Estado = 'Aceptado'. Si no hay ninguno, borra la validación.
 * Detecta columnas por nombre de cabecera (tolerante a estructuras distintas).
 */
function _fpRefreshPresupuestoDropdown_(sh, cliente) {
  Logger.log('[_fpRefreshPresupuestoDropdown_] INICIO cliente="' + cliente + '"');
  var cell = sh.getRange(PF_.PRE_ID[0], PF_.PRE_ID[1]);
  if (!cliente) { cell.clearDataValidations(); return; }
  try {
    var preSheet = SpreadsheetApp.getActive().getSheetByName('PRESUPUESTOS');
    Logger.log('  preSheet: ' + (preSheet ? 'OK lastRow=' + preSheet.getLastRow() : 'NULL'));
    if (!preSheet || preSheet.getLastRow() < 2) { cell.clearDataValidations(); return; }

    // Detectar columnas por cabecera real
    var hdrs   = preSheet.getRange(1, 1, 1, preSheet.getLastColumn()).getValues()[0];
    var colId = -1, colCliente = -1, colEstado = -1;
    for (var h = 0; h < hdrs.length; h++) {
      var hn = String(hdrs[h]).trim();
      if (hn === 'ID_Presupuesto') colId      = h + 1;
      if (hn === 'Cliente')        colCliente  = h + 1;
      if (hn === 'Estado')         colEstado   = h + 1;
    }
    Logger.log('  colId=' + colId + '  colCliente=' + colCliente + '  colEstado=' + colEstado);
    if (colId < 1 || colCliente < 1 || colEstado < 1) { cell.clearDataValidations(); return; }

    var colMax = Math.max(colId, colCliente, colEstado);
    var rows   = preSheet.getRange(2, 1, preSheet.getLastRow() - 1, colMax).getValues();
    var ids    = [];
    for (var i = 0; i < rows.length; i++) {
      var id     = String(rows[i][colId      - 1] || '').trim();
      var cli    = String(rows[i][colCliente  - 1] || '').trim();
      var estado = String(rows[i][colEstado   - 1] || '').trim();
      if (id && cli === cliente && estado === 'Aceptado') ids.push(id);
    }
    Logger.log('  Presupuestos aceptados para "' + cliente + '": ' + ids.length + ' → ' + JSON.stringify(ids));
    if (!ids.length) { cell.clearDataValidations(); return; }
    cell.setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(ids, true).setAllowInvalid(false).build()
    );
    Logger.log('  Validación C66 aplicada OK.');
  } catch (e) {
    Logger.log('  ❌ ERROR: ' + e.message);
    cell.clearDataValidations();
  }
  Logger.log('[_fpRefreshPresupuestoDropdown_] FIN');
}

// ================================================================
// SECCIÓN FACTURA DESDE PRESUPUESTO — LÓGICA
// ================================================================

function panelCargarPresupuesto() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('PANEL');
  if (!sh) return;

  var cliente = String(sh.getRange(PF_.CLIENTE[0], PF_.CLIENTE[1]).getValue() || '').trim();
  if (!cliente) {
    _fpEstado_(sh, '⚠️  Selecciona primero un cliente.', '#f9ab00');
    return;
  }

  var preId = String(sh.getRange(PF_.PRE_ID[0], PF_.PRE_ID[1]).getValue() || '').trim();
  if (!preId) {
    _fpEstado_(sh, '⚠️  Selecciona un presupuesto Aceptado del desplegable.', '#f9ab00');
    return;
  }

  _fpEstado_(sh, '⏳  Cargando ' + preId + '…', '#f9ab00');
  SpreadsheetApp.flush();

  try {
    var data = getPresupuestoParaFacturar(preId);
    if (!data) {
      _fpEstado_(sh, '❌  Presupuesto ' + preId + ' no encontrado.', '#ea4335');
      return;
    }

    // Info read-only
    var tz       = Session.getScriptTimeZone();
    var fechaStr = (data.fechaEmision instanceof Date && !isNaN(data.fechaEmision.getTime()))
      ? Utilities.formatDate(data.fechaEmision, tz, 'dd/MM/yyyy') : '—';
    var _liRaw_ = data.lineItems || [];
    var baseRaw = 0;
    for (var _bi_ = 0; _bi_ < _liRaw_.length; _bi_++) {
      var _bLi_ = _liRaw_[_bi_];
      baseRaw += (_bLi_.cant > 0 && _bLi_.pu > 0)
        ? Math.round(_bLi_.cant * _bLi_.pu * 100) / 100
        : (Number(_bLi_.imp) || 0);
    }
    baseRaw = _liRaw_.length ? Math.round(baseRaw * 100) / 100 : (data.subtotal || data.total || 0);
    sh.getRange(PF_.INFO_CLIENTE[0],    PF_.INFO_CLIENTE[1]).setValue(data.cliente);
    sh.getRange(PF_.INFO_TOTAL[0],      PF_.INFO_TOTAL[1]).setValue(formatCurrency(baseRaw));
    sh.getRange(PF_.INFO_FECHA[0],      PF_.INFO_FECHA[1]).setValue(fechaStr);
    sh.getRange(PF_.BASE_PRESUP_RAW[0], PF_.BASE_PRESUP_RAW[1]).setValue(baseRaw);

    // Líneas del presupuesto (conceptos + cant + P.U., sin IVA)
    var originalItems = data.lineItems || [];
    var items = originalItems.length > PFP_TABLE.ROWS
      ? originalItems.slice(0, PFP_TABLE.ROWS)
      : originalItems;

    var msg, color;
    if (!originalItems.length) {
      // Sin líneas: no limpiar tabla ni borrar entradas manuales
      msg   = '⚠️  ' + preId + ' — sin líneas recuperables. Ejecuta "🔧 Reparar filas PRESUPUESTOS" y vuelve a cargar.';
      color = '#ea4335';
    } else {
      // Volcar todas las líneas en batch (desc, cant, pu) + marcar origen en col I
      var tableData   = [];
      var origenData  = [];
      for (var i = 0; i < PFP_TABLE.ROWS; i++) {
        if (i < items.length) {
          tableData.push([
            items[i].desc || '',
            items[i].cant > 0 ? items[i].cant : '',
            items[i].pu   > 0 ? items[i].pu   : ''
          ]);
          origenData.push(['PRESUPUESTO']);
        } else {
          tableData.push(['', '', '']);
          origenData.push(['']);
        }
      }
      sh.getRange(PFP_TABLE.START_ROW, PFP_TABLE.COL_DESC,   PFP_TABLE.ROWS, 3).setValues(tableData);
      sh.getRange(PFP_TABLE.START_ROW, PFP_TABLE.COL_ORIGEN,  PFP_TABLE.ROWS, 1).setValues(origenData);
      _applyPFPImpFormulas_(sh);  // garantiza fórmula en H73 (primera fila) y resto

      // Colorear: líneas del presupuesto = azul claro / texto azul
      //           filas extra vacías     = blanco / texto negro
      sh.getRange(PFP_TABLE.START_ROW, PFP_TABLE.COL_DESC, items.length, 4)
        .setBackground('#e8f0fe').setFontColor('#1a73e8');
      if (items.length < PFP_TABLE.ROWS) {
        sh.getRange(PFP_TABLE.START_ROW + items.length, PFP_TABLE.COL_DESC,
                    PFP_TABLE.ROWS - items.length, 4)
          .setBackground('#ffffff').setFontColor('#000000');
      }

      if (originalItems.length > PFP_TABLE.ROWS) {
        msg   = '⚠️  ' + preId + ' — truncado a ' + PFP_TABLE.ROWS + ' líneas (presupuesto tiene ' + originalItems.length + '). Añade extras debajo y pulsa VISTA PREVIA.';
        color = '#f9ab00';
      } else {
        msg   = '✅  ' + preId + ' — ' + items.length + ' línea(s) en azul. Añade extras en negro debajo y pulsa VISTA PREVIA.';
        color = '#34a853';
      }
    }

    // Flags
    sh.getRange(PF_.CARGAR_OK[0],  PF_.CARGAR_OK[1]).setValue(preId);
    sh.getRange(PF_.PREVIEW_OK[0], PF_.PREVIEW_OK[1]).setValue('');
    _fpClearPreview_(sh);
    _fpEstado_(sh, msg, color);

  } catch (ex) {
    _fpEstado_(sh, '❌  Error: ' + ex.message, '#ea4335');
    logEvent('ERROR', 'panelCargarPresupuesto', ex.message, ex.stack || '');
  }
}

function panelVistaPreviaFacturaPresupuesto() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('PANEL');
  if (!sh) return;

  var cargarOk = String(sh.getRange(PF_.CARGAR_OK[0], PF_.CARGAR_OK[1]).getValue()).trim();
  if (!cargarOk) {
    _fpEstado_(sh, '⚠️  Carga primero un presupuesto antes de hacer la vista previa.', '#f9ab00');
    return;
  }

  var items = _getPanelFPItems_(sh);
  if (!items.length) {
    _fpEstado_(sh, '⚠️  La tabla no tiene líneas con contenido.', '#f9ab00');
    return;
  }

  _fpClearPreview_(sh);
  sh.getRange(PF_.PREVIEW_OK[0], PF_.PREVIEW_OK[1]).setValue('');
  _fpEstado_(sh, '⏳  Calculando…', '#f9ab00');
  SpreadsheetApp.flush();

  try {
    var tipo     = String(sh.getRange(PF_.TIPO_DOC[0], PF_.TIPO_DOC[1]).getValue() || 'Factura').trim();
    var ivaPct   = parseFloat(getConfig('iva_porcentaje') || '21');
    var subtotal = 0;
    for (var i = 0; i < items.length; i++) subtotal += items[i].imp;
    subtotal    = Math.round(subtotal * 100) / 100;
    var ivaImp  = tipo === 'Proforma' ? 0 : Math.round(subtotal * ivaPct) / 100;
    var total   = subtotal + ivaImp;
    var cliente = String(sh.getRange(PF_.INFO_CLIENTE[0], PF_.INFO_CLIENTE[1]).getValue() || '').trim();

    sh.getRange(PF_.PRV_CLIENTE[0],  PF_.PRV_CLIENTE[1]).setValue(cliente);
    sh.getRange(PF_.PRV_LINEAS[0],   PF_.PRV_LINEAS[1]).setValue(items.length + ' línea(s)');
    sh.getRange(PF_.PRV_IVA[0],      PF_.PRV_IVA[1]).setValue(
      tipo === 'Proforma' ? 'Sin IVA (Proforma)' : 'IVA ' + ivaPct + '%:  ' + formatCurrency(ivaImp)
    );
    sh.getRange(PF_.PRV_PRE_REF[0],  PF_.PRV_PRE_REF[1]).setValue(cargarOk);
    sh.getRange(PF_.PRV_SUBTOTAL[0], PF_.PRV_SUBTOTAL[1]).setValue(formatCurrency(subtotal));
    sh.getRange(PF_.PRV_TOTAL[0],    PF_.PRV_TOTAL[1]).setValue(formatCurrency(total)).setFontWeight('bold');

    // Validación: aviso solo si el subtotal es menor que la base del presupuesto (líneas eliminadas)
    var baseRaw = Number(sh.getRange(PF_.BASE_PRESUP_RAW[0], PF_.BASE_PRESUP_RAW[1]).getValue() || 0);
    var warnMsg = '';
    if (baseRaw > 0 && subtotal < baseRaw - 0.02) {
      warnMsg = '  ⚠️  El subtotal de la tabla (' + formatCurrency(subtotal) +
                ') es menor que la base del presupuesto (' + formatCurrency(baseRaw) + '). Puede que falten líneas.';
    }

    sh.getRange(PF_.PREVIEW_OK[0], PF_.PREVIEW_OK[1]).setValue('OK');
    _fpEstado_(sh,
      '✅  Vista previa lista. Revisa y pulsa GENERAR ' + (tipo === 'Proforma' ? 'PROFORMA' : 'FACTURA') + '.' + warnMsg,
      warnMsg ? '#f9ab00' : '#34a853');

  } catch (ex) {
    _fpEstado_(sh, '❌  Error: ' + ex.message, '#ea4335');
    logEvent('ERROR', 'panelVistaPreviaFacturaPresupuesto', ex.message, ex.stack || '');
  }
}

function panelGenerarFacturaPresupuesto() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('PANEL');
  if (!sh) return;

  var previewOk = String(sh.getRange(PF_.PREVIEW_OK[0], PF_.PREVIEW_OK[1]).getValue()).trim();
  if (previewOk !== 'OK') {
    _fpEstado_(sh, '⚠️  Genera primero la vista previa antes de emitir la factura.', '#f9ab00');
    return;
  }

  var preId   = String(sh.getRange(PF_.CARGAR_OK[0],  PF_.CARGAR_OK[1]).getValue()).trim();
  var cliente = String(sh.getRange(PF_.INFO_CLIENTE[0], PF_.INFO_CLIENTE[1]).getValue() || '').trim();
  var tipo    = String(sh.getRange(PF_.TIPO_DOC[0],    PF_.TIPO_DOC[1]).getValue()    || 'Factura').trim();
  var items   = _getPanelFPItems_(sh);

  _fpEstado_(sh, '⏳  Generando ' + (tipo === 'Proforma' ? 'proforma' : 'factura') + '…', '#f9ab00');
  sh.getRange(PF_.URL[0], PF_.URL[1]).setValue('');
  SpreadsheetApp.flush();

  try {
    var result = generateInvoiceFromPresupuestoPDF({
      presupuestoId: preId,
      cliente:       cliente,
      lineItems:     items,
      tipo:          tipo
    });

    sh.getRange(PF_.URL[0], PF_.URL[1]).setValue(result.url).setFontColor('#1a73e8');
    _fpEstado_(sh, '✅  ' + result.id + ' emitida correctamente.', '#34a853');
    sh.getRange(PF_.PREVIEW_OK[0], PF_.PREVIEW_OK[1]).setValue('');
    sh.getRange(PF_.CARGAR_OK[0],  PF_.CARGAR_OK[1]).setValue('');

  } catch (ex) {
    _fpEstado_(sh, '❌  Error: ' + ex.message, '#ea4335');
    logEvent('ERROR', 'panelGenerarFacturaPresupuesto', ex.message, ex.stack || '');
  }
}

function _getPanelFPItems_(sh) {
  var items = [];
  // Lee 4 cols: desc(E), cant(F), pu(G), imp(H — fórmula o valor manual directo)
  var data = sh.getRange(PFP_TABLE.START_ROW, PFP_TABLE.COL_DESC, PFP_TABLE.ROWS, 4).getValues();
  for (var i = 0; i < data.length; i++) {
    var desc = String(data[i][0] || '').trim();
    var cant = Number(data[i][1]) || 0;
    var pu   = Number(data[i][2]) || 0;
    var impH = Number(data[i][3]) || 0; // valor de col H (fórmula calculada o entrada directa)
    if (!desc && cant === 0 && pu === 0 && impH === 0) continue;
    // Usar cant×pu si ambos disponibles; si no, usar el valor de col H como fallback
    var imp = (cant > 0 && pu > 0)
      ? Math.round(cant * pu * 100) / 100
      : impH;
    items.push({ desc: desc, cant: cant, pu: pu, imp: imp });
  }
  return items;
}

function _fpClearPreview_(sh) {
  var cells = [
    PF_.PRV_CLIENTE, PF_.PRV_LINEAS, PF_.PRV_IVA,
    PF_.PRV_PRE_REF, PF_.PRV_SUBTOTAL, PF_.PRV_TOTAL
  ];
  for (var i = 0; i < cells.length; i++) {
    sh.getRange(cells[i][0], cells[i][1]).setValue('').setFontWeight('normal');
  }
}

function _fpEstado_(sh, msg, color) {
  sh.getRange(PF_.ESTADO[0], PF_.ESTADO[1]).setValue(msg).setFontColor(color || '#3c4043');
}

// ================================================================
// DIAGNÓSTICO Y MANTENIMIENTO
// ================================================================

/**
 * Fuerza el refresco de los dropdowns de la sección Factura desde Presupuesto
 * sin reconstruir el panel completo. Útil cuando se han cambiado estados
 * de presupuestos y el panel ya está abierto.
 */
function panelRefrescarDropdownsFP() {
  var sh = SpreadsheetApp.getActive().getSheetByName('PANEL');
  if (!sh) {
    SpreadsheetApp.getUi().alert('❌ No existe la hoja PANEL. Ejecuta primero ① Crear/Reconfigurar PANEL.');
    return;
  }
  _fpRefreshClienteDropdown_(sh);
  var clienteActual = String(sh.getRange(PF_.CLIENTE[0], PF_.CLIENTE[1]).getValue() || '').trim();
  if (clienteActual) {
    _fpRefreshPresupuestoDropdown_(sh, clienteActual);
  } else {
    sh.getRange(PF_.PRE_ID[0], PF_.PRE_ID[1]).clearDataValidations();
  }
  SpreadsheetApp.getUi().alert('✅ Dropdowns refrescados.\n\nCliente activo: ' + (clienteActual || '(ninguno)'));
}

/**
 * Diagnóstico completo de la hoja PRESUPUESTOS:
 * estructura, Estado, validaciones y datos Aceptados.
 * Resultados visibles en: Ver → Registros de ejecución.
 */
function verificarEstadoPresupuestos() {
  Logger.log('============================================================');
  Logger.log('verificarEstadoPresupuestos()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('============================================================');

  // ── 1. Identificar spreadsheets ──────────────────────────────
  var ssActive = SpreadsheetApp.getActiveSpreadsheet();
  var ssTarget = getTargetSS_();
  Logger.log('\n[0] Spreadsheets:');
  Logger.log('  getActive() ID  : ' + ssActive.getId());
  Logger.log('  getTargetSS_()  : ' + ssTarget.getId());
  Logger.log('  Son iguales     : ' + (ssActive.getId() === ssTarget.getId()));

  // ── 2. Buscar hoja PRESUPUESTOS ──────────────────────────────
  Logger.log('\n[1] Hoja PRESUPUESTOS:');
  var sh = ssActive.getSheetByName('PRESUPUESTOS');
  if (!sh) {
    Logger.log('  ❌ NO encontrada en getActive().');
    sh = ssTarget.getSheetByName('PRESUPUESTOS');
    if (!sh) { Logger.log('  ❌ Tampoco en getTargetSS_().'); return; }
    Logger.log('  ⚠️  Encontrada en getTargetSS_() pero no en getActive(). Las funciones de Panel usan getActive().');
  } else {
    Logger.log('  ✓ Encontrada en getActive(). Filas totales: ' + sh.getMaxRows());
  }

  // ── 3. Cabeceras reales ──────────────────────────────────────
  Logger.log('\n[2] Cabeceras (fila 1):');
  var lastCol   = Math.max(sh.getLastColumn(), 1);
  var hdrs      = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var colEstado = -1;
  for (var c = 0; c < hdrs.length; c++) {
    Logger.log('  Col ' + (c + 1) + ': "' + hdrs[c] + '"');
    if (String(hdrs[c]).trim() === 'Estado') colEstado = c + 1;
  }
  Logger.log('  → Columna Estado detectada: ' + (colEstado > 0 ? colEstado : '❌ NO ENCONTRADA'));
  Logger.log('  → COL_PRE.ESTADO constante : ' + COL_PRE.ESTADO);
  if (colEstado > 0 && colEstado !== COL_PRE.ESTADO) {
    Logger.log('  ⚠️  DESAJUSTE: columna real (' + colEstado + ') ≠ constante (' + COL_PRE.ESTADO + ')');
  }

  // ── 4. Validación de datos en columna Estado (por nombre real) ──
  var cm        = _buildPreColMap_(sh);
  var colEstadoReal = cm['Estado'] || COL_PRE.ESTADO;
  var colClienteReal = cm['Cliente'] || COL_PRE.CLIENTE;
  var colIdReal     = cm['ID_Presupuesto'] || COL_PRE.ID;
  Logger.log('\n[3] Validación en columna Estado (col real=' + colEstadoReal + ', constante=' + COL_PRE.ESTADO + '):');
  if (colEstadoReal !== COL_PRE.ESTADO) {
    Logger.log('  ⚠️  DESAJUSTE detectado — la hoja tiene Estado en col ' + colEstadoReal +
               ' pero COL_PRE.ESTADO es ' + COL_PRE.ESTADO +
               '. Ejecuta "🔧 Reparar filas PRESUPUESTOS" desde el menú.');
  }
  var val = sh.getRange(2, colEstadoReal).getDataValidation();
  if (!val) {
    Logger.log('  ❌ Sin validación en fila 2.');
  } else {
    Logger.log('  ✓ Tipo: ' + val.getCriteriaType());
    Logger.log('  ✓ Valores: ' + JSON.stringify(val.getCriteriaValues()));
    Logger.log('  ✓ Permite inválido: ' + val.getAllowInvalid());
  }

  // ── 5. Datos ─────────────────────────────────────────────────
  Logger.log('\n[4] Presupuestos (usando columnas reales de la hoja):');
  var lastRow = sh.getLastRow();
  if (lastRow < 2) {
    Logger.log('  Sin datos (solo cabecera o vacío).');
    return;
  }
  var numCols = Math.max(colEstadoReal, colClienteReal, colIdReal);
  var rows    = sh.getRange(2, 1, lastRow - 1, numCols).getValues();
  var total   = 0, aceptados = 0;
  var clientesAceptados = {};
  for (var i = 0; i < rows.length; i++) {
    var id     = String(rows[i][colIdReal      - 1] || '').trim();
    var cli    = String(rows[i][colClienteReal  - 1] || '').trim();
    var estado = String(rows[i][colEstadoReal   - 1] || '').trim();
    if (!id) continue;
    total++;
    var esValido = (ESTADOS_PRESUPUESTO.indexOf(estado) !== -1);
    Logger.log('  ' + id + ' | ' + cli + ' | Estado="' + estado + '"' + (esValido ? '' : ' ⚠️ INVÁLIDO'));
    if (estado === 'Aceptado') {
      aceptados++;
      clientesAceptados[cli] = true;
    }
  }
  Logger.log('\n  Total presupuestos : ' + total);
  Logger.log('  Con Estado=Aceptado: ' + aceptados);
  Logger.log('  Clientes Aceptados : ' + Object.keys(clientesAceptados).join(', '));

  Logger.log('\n============================================================');
  Logger.log('FIN verificarEstadoPresupuestos');
  Logger.log('============================================================');
}

/**
 * Wrapper con feedback UI para repairPresupuestosRows().
 * Aparece en el menú como "🔧 Reparar filas PRESUPUESTOS".
 */
function runRepairPresupuestosRows() {
  try {
    var result = repairPresupuestosRows();
    var msg = 'Reparación completada:\n' +
              '  • Filas reparadas : ' + result.fixed  + '\n' +
              '  • Ya correctas    : ' + result.skipped + '\n' +
              (result.errors.length ? '  • Errores        : ' + result.errors.join('; ') : '');
    SpreadsheetApp.getUi().alert('🔧 Reparar PRESUPUESTOS', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error en reparación: ' + e.message);
  }
}

/**
 * Diagnóstico de la sección Factura desde Presupuesto en el PANEL:
 * verifica coordenadas, validaciones y estado de los dropdowns.
 * Resultados visibles en: Ver → Registros de ejecución.
 */
function verificarFacturaDesdePresupuestoPanel() {
  Logger.log('============================================================');
  Logger.log('verificarFacturaDesdePresupuestoPanel()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('============================================================');

  // ── 1. Spreadsheets ──────────────────────────────────────────
  var ssActive = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('\n[0] getActive() ID: ' + ssActive.getId());

  // ── 2. Hoja PANEL ────────────────────────────────────────────
  Logger.log('\n[1] Hoja PANEL:');
  var sh = ssActive.getSheetByName('PANEL');
  if (!sh) { Logger.log('  ❌ NO existe.'); return; }
  Logger.log('  ✓ Existe. Filas: ' + sh.getMaxRows() + ', Cols: ' + sh.getMaxColumns());

  // ── 3. Constantes PF_ ────────────────────────────────────────
  Logger.log('\n[2] Constantes PF_:');
  Logger.log('  PF_.CLIENTE       : fila ' + PF_.CLIENTE[0]       + ', col ' + PF_.CLIENTE[1]       + ' → C' + PF_.CLIENTE[0]);
  Logger.log('  PF_.PRE_ID        : fila ' + PF_.PRE_ID[0]        + ', col ' + PF_.PRE_ID[1]        + ' → C' + PF_.PRE_ID[0]);
  Logger.log('  PF_.INFO_CLIENTE  : fila ' + PF_.INFO_CLIENTE[0]  + ', col ' + PF_.INFO_CLIENTE[1]);
  Logger.log('  PF_.BOTON_CARGAR  : fila ' + PF_.BOTON_CARGAR[0]  + ', col ' + PF_.BOTON_CARGAR[1]);
  Logger.log('  PF_.CARGAR_OK     : fila ' + PF_.CARGAR_OK[0]     + ', col ' + PF_.CARGAR_OK[1]);
  Logger.log('  PF_.BOTON_PREVIEW : fila ' + PF_.BOTON_PREVIEW[0] + ', col ' + PF_.BOTON_PREVIEW[1]);
  Logger.log('  PF_.PREVIEW_OK    : fila ' + PF_.PREVIEW_OK[0]    + ', col ' + PF_.PREVIEW_OK[1]);
  Logger.log('  PF_.BOTON_GENERAR : fila ' + PF_.BOTON_GENERAR[0] + ', col ' + PF_.BOTON_GENERAR[1]);
  Logger.log('  PF_.ESTADO        : fila ' + PF_.ESTADO[0]        + ', col ' + PF_.ESTADO[1]);

  // ── 4. Celda Cliente (C65) ────────────────────────────────────
  Logger.log('\n[3] Celda Cliente (' + PF_.CLIENTE[0] + ',' + PF_.CLIENTE[1] + '):');
  _logCeldaDiag_(sh, PF_.CLIENTE[0], PF_.CLIENTE[1]);

  // ── 5. Celda Presupuesto (C66) ────────────────────────────────
  Logger.log('\n[4] Celda Presupuesto (' + PF_.PRE_ID[0] + ',' + PF_.PRE_ID[1] + '):');
  _logCeldaDiag_(sh, PF_.PRE_ID[0], PF_.PRE_ID[1]);

  // ── 6. Checkbox CARGAR (H70) ─────────────────────────────────
  Logger.log('\n[5] Checkbox CARGAR (' + PF_.BOTON_CARGAR[0] + ',' + PF_.BOTON_CARGAR[1] + '):');
  _logCeldaDiag_(sh, PF_.BOTON_CARGAR[0], PF_.BOTON_CARGAR[1]);

  // ── 7. Flags ─────────────────────────────────────────────────
  Logger.log('\n[6] Flags:');
  Logger.log('  CARGAR_OK  (' + PF_.CARGAR_OK[0]  + ',' + PF_.CARGAR_OK[1]  + '): "' +
             sh.getRange(PF_.CARGAR_OK[0],  PF_.CARGAR_OK[1]).getValue()  + '"');
  Logger.log('  PREVIEW_OK (' + PF_.PREVIEW_OK[0] + ',' + PF_.PREVIEW_OK[1] + '): "' +
             sh.getRange(PF_.PREVIEW_OK[0], PF_.PREVIEW_OK[1]).getValue() + '"');

  // ── 8. Alturas de filas clave ─────────────────────────────────
  Logger.log('\n[7] Alturas filas 63-70:');
  for (var r = 63; r <= 70; r++) {
    Logger.log('  Fila ' + r + ': ' + sh.getRowHeight(r) + ' px');
  }

  Logger.log('\n============================================================');
  Logger.log('FIN verificarFacturaDesdePresupuestoPanel');
  Logger.log('============================================================');
}

function _logCeldaDiag_(sh, row, col) {
  var cell = sh.getRange(row, col);
  var val  = cell.getValue();
  Logger.log('  Valor actual : "' + val + '"');
  Logger.log('  Alto de fila : ' + sh.getRowHeight(row) + ' px');
  var dv = cell.getDataValidation();
  if (!dv) {
    Logger.log('  Validación   : ❌ ninguna');
  } else {
    var cv = dv.getCriteriaValues();
    Logger.log('  Validación   : ✓ tipo=' + dv.getCriteriaType());
    Logger.log('  Valores list : ' + (cv && cv[0] ? JSON.stringify(cv[0]).substring(0, 200) : '(vacío)'));
  }
}

// ================================================================
// TABLA TRABAJOS DISPONIBLES PARA FACTURAR (filas 40-72)
// ================================================================

/**
 * Construye la sección "TRABAJOS DISPONIBLES PARA FACTURAR" en la hoja PANEL.
 * Llamado desde setupPanel(). Crea cabecera, encabezados y 30 filas de datos
 * inicialmente colapsadas (altura 2px). panelBuscarTrabajos() las expande.
 */
function _buildJobsSection_(sh) {
  var END_ROW  = JOB_TABLE.START_ROW + JOB_TABLE.MAX_ROWS - 1;  // 129
  var HINT_ROW = END_ROW + 1;                                    // 130
  var r;

  // ── Alturas ────────────────────────────────────────────────────
  sh.setRowHeight(JOB_TABLE.SEC_ROW, 44);   // 98 — cabecera sección
  sh.setRowHeight(JOB_TABLE.HDR_ROW, 30);   // 99 — encabezados columna
  for (r = JOB_TABLE.START_ROW; r <= END_ROW; r++) sh.setRowHeight(r, 2);
  sh.setRowHeight(HINT_ROW, 20);

  // ── Cabecera sección (fila 40) ─────────────────────────────────
  sh.getRange(JOB_TABLE.SEC_ROW, 2, 1, 8).merge()
    .setValue('📋  TRABAJOS DISPONIBLES PARA FACTURAR')
    .setFontSize(14).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#1a73e8').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, false, true, false, false, '#0d47a1', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // ── Encabezados columna (fila 41) ─────────────────────────────
  sh.getRange(JOB_TABLE.HDR_ROW, 2)
    .setValue('Sel.').setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#1a73e8');
  sh.getRange(JOB_TABLE.HDR_ROW, JOB_TABLE.COL_INFO, 1, 2).merge()
    .setValue('ID · Fecha · Operario').setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
    .setHorizontalAlignment('left').setVerticalAlignment('middle').setBackground('#1a73e8');
  sh.getRange(JOB_TABLE.HDR_ROW, JOB_TABLE.COL_TRAB)
    .setValue('Trabajo realizado / Incidencias').setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
    .setHorizontalAlignment('left').setVerticalAlignment('middle').setBackground('#1a73e8');
  sh.getRange(JOB_TABLE.HDR_ROW, JOB_TABLE.COL_HORAS)
    .setValue('Horas').setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#1a73e8');
  sh.getRange(JOB_TABLE.HDR_ROW, JOB_TABLE.COL_INFORME)
    .setValue('Informe').setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#1565c0');
  sh.getRange(JOB_TABLE.HDR_ROW, JOB_TABLE.COL_PROFORMA)
    .setValue('Proforma').setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#e65100');
  sh.getRange(JOB_TABLE.HDR_ROW, JOB_TABLE.COL_FACTURA)
    .setValue('Factura').setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#2e7d32');
  sh.getRange(JOB_TABLE.HDR_ROW, 2, 1, 8)
    .setBorder(true, true, true, true, false, false, '#0d47a1', SpreadsheetApp.BorderStyle.SOLID);

  // ── Filas de datos (preformateadas, inicialmente colapsadas) ───
  for (r = JOB_TABLE.START_ROW; r <= END_ROW; r++) {
    sh.getRange(r, JOB_TABLE.COL_CHECK)
      .setBackground('#ffffff').setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(r, JOB_TABLE.COL_INFO, 1, 2).merge()
      .setBackground('#ffffff').setFontSize(8).setVerticalAlignment('middle').setWrap(true)
      .setBorder(false, true, false, false, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, JOB_TABLE.COL_TRAB)
      .setBackground('#ffffff').setFontSize(8).setVerticalAlignment('middle').setWrap(true)
      .setBorder(false, true, false, false, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, JOB_TABLE.COL_HORAS)
      .setBackground('#ffffff').setFontSize(9).setVerticalAlignment('middle')
      .setHorizontalAlignment('center')
      .setBorder(false, true, false, false, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, JOB_TABLE.COL_INFORME)
      .setBackground('#f8f9fa').setFontSize(8).setVerticalAlignment('middle')
      .setHorizontalAlignment('center').setWrap(true)
      .setBorder(false, true, false, false, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, JOB_TABLE.COL_PROFORMA)
      .setBackground('#f8f9fa').setFontSize(8).setVerticalAlignment('middle')
      .setHorizontalAlignment('center').setWrap(true)
      .setBorder(false, true, false, false, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, JOB_TABLE.COL_FACTURA)
      .setBackground('#f8f9fa').setFontSize(8).setVerticalAlignment('middle')
      .setHorizontalAlignment('center').setWrap(true)
      .setBorder(false, true, false, true, false, false, '#e8eaed', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, JOB_TABLE.COL_DATA)
      .setFontSize(1).setFontColor('#ffffff').setBackground('#ffffff');
  }

  // ── Hint (fila 130) ───────────────────────────────────────────
  sh.getRange(HINT_ROW, 2, 1, 7).merge()
    .setValue('💡 Pulsa "🔍 BUSCAR TRABAJOS" para cargar los trabajos del período')
    .setFontSize(9).setFontColor('#9aa0a6')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
}

// ================================================================
// BUSCAR TRABAJOS — carga y muestra la tabla de trabajos del período
// ================================================================

function panelBuscarTrabajos() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('PANEL');
  if (!sh) return;

  var cliente  = String(sh.getRange(P_.CLIENTE[0],   P_.CLIENTE[1]).getValue()   || '').trim();
  var fechaIni = sh.getRange(P_.FECHA_INI[0], P_.FECHA_INI[1]).getDisplayValue().trim();
  var fechaFin = sh.getRange(P_.FECHA_FIN[0], P_.FECHA_FIN[1]).getDisplayValue().trim();

  sh.getRange(P_.JOBS_OK[0],    P_.JOBS_OK[1]).setValue('');
  sh.getRange(P_.PREVIEW_OK[0], P_.PREVIEW_OK[1]).setValue('');
  _panelEstado_(sh, '⏳  Buscando trabajos…', '#f9ab00');
  SpreadsheetApp.flush();

  if (!cliente) {
    _panelEstado_(sh, '❌  Selecciona un cliente antes de buscar.', '#ea4335');
    return;
  }
  if (!fechaIni || !fechaFin) {
    _panelEstado_(sh, '❌  Rellena Fecha Inicio y Fecha Fin antes de buscar.', '#ea4335');
    return;
  }
  var fi = _parseReportDate_(fechaIni);
  var ff = _parseReportDate_(fechaFin);
  if (!fi || !ff) {
    _panelEstado_(sh, '❌  Fechas inválidas. Usa formato DD/MM/YYYY.', '#ea4335');
    return;
  }
  ff.setHours(23, 59, 59, 999);

  // ── DIAGNÓSTICO ───────────────────────────────────────────────
  Logger.log('================================================');
  Logger.log('panelBuscarTrabajos() — DIAGNÓSTICO');
  Logger.log('================================================');
  Logger.log('[PARAMS] cliente       : "' + cliente + '"');
  Logger.log('[PARAMS] cliente.length: ' + cliente.length);
  Logger.log('[PARAMS] cliente charCodes: ' + cliente.split('').map(function(c){return c.charCodeAt(0);}).join(','));
  Logger.log('[PARAMS] fechaIni      : "' + fechaIni + '"');
  Logger.log('[PARAMS] fechaFin      : "' + fechaFin + '"');
  Logger.log('[PARAMS] fi (parsed)   : ' + (fi ? fi.toISOString() : 'null'));
  Logger.log('[PARAMS] ff (parsed)   : ' + (ff ? ff.toISOString() : 'null'));

  try {
    // ── Leer TRABAJOS sin pasar por PropertiesService ─────────────
    var shTra = ss.getSheetByName('TRABAJOS');
    if (!shTra) {
      _panelEstado_(sh, '❌  No existe la hoja TRABAJOS.', '#ea4335');
      return;
    }

    var lastRow = shTra.getLastRow();
    Logger.log('[TRABAJOS] lastRow: ' + lastRow + '  lastColumn: ' + shTra.getLastColumn());
    var trabajos = [];

    if (lastRow > 1) {
      var numCols = Math.min(shTra.getLastColumn(), COL_TRA.MATERIALES);
      var data    = shTra.getRange(2, 1, lastRow - 1, numCols).getValues();
      Logger.log('[TRABAJOS] filas leídas: ' + data.length + '  columnas leídas: ' + numCols);
      var rechazadosCliente = 0;
      var rechazadosFecha   = 0;
      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var rawId     = String(row[COL_TRA.ID       - 1] || '').trim();
        var rawCliente = String(row[COL_TRA.CLIENTE  - 1] || '').trim();
        var rawFecha  = row[COL_TRA.FECHA_TRABAJO - 1];

        if (rawCliente.toLowerCase() !== cliente.toLowerCase()) {
          if (rawId) {
            Logger.log('[FILTRO-CLIENTE] RECHAZADO fila ' + (i + 2) + ' ID=' + rawId +
                       ' | rawCliente="' + rawCliente + '"' +
                       ' | length=' + rawCliente.length +
                       ' | charCodes=' + rawCliente.split('').map(function(c){return c.charCodeAt(0);}).join(','));
          }
          rechazadosCliente++;
          continue;
        }

        var ft = _parseTrabajosDate_(rawFecha);
        if (!ft || ft < fi || ft > ff) {
          Logger.log('[FILTRO-FECHA] RECHAZADO fila ' + (i + 2) + ' ID=' + rawId +
                     ' | rawFecha=' + rawFecha + ' (tipo:' + typeof rawFecha + ')' +
                     ' | ft=' + (ft ? ft.toISOString() : 'null') +
                     ' | fi=' + fi.toISOString() + ' | ff=' + ff.toISOString());
          rechazadosFecha++;
          continue;
        }

        Logger.log('[FILTRO-OK] ACEPTADO fila ' + (i + 2) + ' ID=' + rawId +
                   ' | fecha=' + ft.toISOString());
        trabajos.push({
          id:               String(row[COL_TRA.ID               - 1] || '').trim(),
          fechaTrabajo:     ft,
          operario:         numCols >= COL_TRA.OPERARIO          ? String(row[COL_TRA.OPERARIO          - 1] || '').trim() : '',
          horaInicio:       numCols >= COL_TRA.HORA_INICIO       ? row[COL_TRA.HORA_INICIO - 1]                            : '',
          horaFin:          numCols >= COL_TRA.HORA_FIN          ? row[COL_TRA.HORA_FIN    - 1]                            : '',
          trabajoRealizado: numCols >= COL_TRA.TRABAJO_REALIZADO ? String(row[COL_TRA.TRABAJO_REALIZADO - 1] || '').trim() : '',
          incidencias:      numCols >= COL_TRA.INCIDENCIAS       ? String(row[COL_TRA.INCIDENCIAS       - 1] || '').trim() : '',
          materiales:       numCols >= COL_TRA.MATERIALES        ? Number(row[COL_TRA.MATERIALES        - 1] || 0)         : 0
        });
      }
      Logger.log('================================================');
      Logger.log('[RESUMEN] Filas totales leídas  : ' + data.length);
      Logger.log('[RESUMEN] Rechazados por cliente: ' + rechazadosCliente);
      Logger.log('[RESUMEN] Rechazados por fecha  : ' + rechazadosFecha);
      Logger.log('[RESUMEN] Aceptados             : ' + trabajos.length);
      Logger.log('================================================');
    }

    // Mostrar progreso intermedio para confirmar que la función está corriendo
    _panelEstado_(sh, '⏳  ' + trabajos.length + ' trabajo(s) encontrado(s). Dibujando tabla…', '#f9ab00');
    SpreadsheetApp.flush();

    // ── Estado documental (FACTURAS, PROFORMAS, INFORMES) ─────────
    var statusMap = _buildJobStatusMap_();

    // ── Actualizar cabecera para 3 columnas de documentos ─────────
    var END_ROW = JOB_TABLE.START_ROW + JOB_TABLE.MAX_ROWS - 1;
    sh.getRange(JOB_TABLE.SEC_ROW, 2, 1, 8).breakApart().merge()
      .setValue('📋  TRABAJOS DISPONIBLES PARA FACTURAR')
      .setFontSize(14).setFontWeight('bold').setFontColor('#ffffff')
      .setBackground('#1a73e8').setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(JOB_TABLE.HDR_ROW, JOB_TABLE.COL_INFORME, 1, 2).breakApart();
    sh.getRange(JOB_TABLE.HDR_ROW, JOB_TABLE.COL_INFORME)
      .setValue('Informe').setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#1565c0');
    sh.getRange(JOB_TABLE.HDR_ROW, JOB_TABLE.COL_PROFORMA)
      .setValue('Proforma').setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#e65100');
    sh.getRange(JOB_TABLE.HDR_ROW, JOB_TABLE.COL_FACTURA)
      .setValue('Factura').setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#2e7d32');
    // Ajustar anchos de columna (aplica también a hojas creadas con versión anterior)
    sh.setColumnWidth(JOB_TABLE.COL_INFORME,  90);
    sh.setColumnWidth(JOB_TABLE.COL_PROFORMA, 90);
    sh.setColumnWidth(JOB_TABLE.COL_FACTURA,  90);
    sh.setColumnWidth(JOB_TABLE.COL_DATA,      2);

    // ── Limpiar tabla (batch sobre cols B-J) ──────────────────────
    sh.getRange(JOB_TABLE.START_ROW, JOB_TABLE.COL_INFORME, JOB_TABLE.MAX_ROWS, 2).breakApart();
    sh.getRange(JOB_TABLE.START_ROW, 2, JOB_TABLE.MAX_ROWS, 9)
      .clearContent().clearDataValidations().setBackground('#ffffff').setFontColor('#000000');
    for (var rc = JOB_TABLE.START_ROW; rc <= END_ROW; rc++) sh.setRowHeight(rc, 2);

    // ── Rellenar tabla fila a fila ────────────────────────────────
    var tz      = Session.getScriptTimeZone();
    var maxShow = Math.min(trabajos.length, JOB_TABLE.MAX_ROWS);
    for (var j = 0; j < maxShow; j++) {
      var t     = trabajos[j];
      var dr    = JOB_TABLE.START_ROW + j;
      var horas = _calcHoras_(t.horaInicio, t.horaFin);
      var st    = statusMap[t.id] || { informe: '', proforma: '', factura: '' };

      var infoText  = t.id + ' · ' + Utilities.formatDate(t.fechaTrabajo, tz, 'dd/MM/yyyy') +
                      (t.operario ? ' · ' + t.operario : '');
      var trabText  = (t.trabajoRealizado ? t.trabajoRealizado.substring(0, 100) : '—') +
                      (t.incidencias ? ' | ⚠ ' + t.incidencias.substring(0, 60) : '');
      var horasText = horas > 0 ? horas.toFixed(2).replace('.', ',') + ' h' : '0,00 h';
      var checked   = !st.factura; // no pre-marcar si ya tiene factura

      var rowBg = st.factura ? '#f5f5f5' : '#ffffff';

      sh.setRowHeight(dr, 44);
      sh.getRange(dr, 2).insertCheckboxes().setValue(checked);
      sh.getRange(dr, 3).setValue(infoText).setBackground(rowBg).setWrap(true).setFontSize(8).setVerticalAlignment('middle');
      sh.getRange(dr, 5).setValue(trabText).setBackground(rowBg).setWrap(true).setFontSize(8).setVerticalAlignment('middle');
      sh.getRange(dr, 6).setValue(horasText).setBackground(rowBg).setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
      // G — Informe (azul)
      sh.getRange(dr, JOB_TABLE.COL_INFORME)
        .setValue(st.informe || '—')
        .setBackground(st.informe ? '#e8f0fe' : '#f8f9fa')
        .setFontColor(st.informe ? '#1565c0' : '#9aa0a6')
        .setFontWeight(st.informe ? 'bold' : 'normal')
        .setFontSize(8).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
      // H — Proforma (naranja)
      sh.getRange(dr, JOB_TABLE.COL_PROFORMA)
        .setValue(st.proforma || '—')
        .setBackground(st.proforma ? '#fef3cd' : '#f8f9fa')
        .setFontColor(st.proforma ? '#e65100' : '#9aa0a6')
        .setFontWeight(st.proforma ? 'bold' : 'normal')
        .setFontSize(8).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
      // I — Factura (verde)
      sh.getRange(dr, JOB_TABLE.COL_FACTURA)
        .setValue(st.factura || '—')
        .setBackground(st.factura ? '#e6f4ea' : '#f8f9fa')
        .setFontColor(st.factura ? '#2e7d32' : '#9aa0a6')
        .setFontWeight(st.factura ? 'bold' : 'normal')
        .setFontSize(8).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
      // J — JSON datos (oculta)
      sh.getRange(dr, JOB_TABLE.COL_DATA).setValue(JSON.stringify({
        id: t.id,
        fechaTrabajo:     Utilities.formatDate(t.fechaTrabajo, tz, 'dd/MM/yyyy'),
        operario:         t.operario,
        horaInicio:       _horaToStr_(t.horaInicio),
        horaFin:          _horaToStr_(t.horaFin),
        trabajoRealizado: t.trabajoRealizado,
        incidencias:      t.incidencias,
        materiales:       t.materiales,
        horas:            horas
      }));
    }

    SpreadsheetApp.flush();

    sh.getRange(P_.JOBS_OK[0], P_.JOBS_OK[1]).setValue('OK');

    if (trabajos.length === 0) {
      _panelEstado_(sh, '⚠️  No hay trabajos para "' + cliente + '" en ' + fechaIni + ' → ' + fechaFin + '.', '#f9ab00');
    } else {
      var fin = trabajos.length > JOB_TABLE.MAX_ROWS ? ' (mostrando ' + JOB_TABLE.MAX_ROWS + ')' : '';
      _panelEstado_(sh, '✅  ' + trabajos.length + ' trabajo(s) cargado(s)' + fin + '. Baja a la tabla, marca los que incluir y pulsa VISTA PREVIA.', '#34a853');
      // Desplazar la vista a la tabla de trabajos
      sh.setActiveCell(sh.getRange(JOB_TABLE.SEC_ROW, 2));
    }

  } catch (ex) {
    _panelEstado_(sh, '❌  Error: ' + ex.message, '#ea4335');
    logEvent('ERROR', 'panelBuscarTrabajos', ex.message, ex.stack || '');
  }
}

/**
 * Lee FACTURAS (col 10), PROFORMAS (col 10) e INFORMES (col 11 = IDs_Trabajos)
 * y devuelve un mapa jobId → { informe, proforma, factura }.
 * Cada campo contiene el ID del documento o '' si no existe.
 */
function _buildJobStatusMap_() {
  var map = {};
  var ss  = SpreadsheetApp.getActive();

  var _indexSheet_ = function(shName, field, idsCol) {
    var sh = ss.getSheetByName(shName);
    if (!sh || sh.getLastRow() < 2) return;
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, idsCol).getValues();
    for (var i = 0; i < data.length; i++) {
      var docId  = String(data[i][0] || '').trim();
      var idsRaw = String(data[i][idsCol - 1] || '').trim();
      if (!docId || !idsRaw) continue;
      var ids = idsRaw.split(',');
      for (var j = 0; j < ids.length; j++) {
        var jobId = ids[j].trim();
        if (!jobId) continue;
        if (!map[jobId]) map[jobId] = { informe: '', proforma: '', factura: '' };
        if (!map[jobId][field]) map[jobId][field] = docId; // primer documento gana
      }
    }
  };

  _indexSheet_('FACTURAS',  'factura',  10);
  _indexSheet_('PROFORMAS', 'proforma', 10);
  _indexSheet_('INFORMES',  'informe',  11);
  return map;
}

/**
 * Convierte un valor de hora (Date, número fraccionario o string) a "HH:MM".
 * Devuelve '' si no hay valor.
 */
function _horaToStr_(val) {
  if (!val && val !== 0) return '';
  var tz = Session.getScriptTimeZone();
  if (val instanceof Date) {
    return isNaN(val.getTime()) ? '' : Utilities.formatDate(val, tz, 'HH:mm');
  }
  if (typeof val === 'number') {
    var totalMin = Math.round(val * 24 * 60);
    var h = Math.floor(totalMin / 60);
    var m = totalMin % 60;
    return (h < 10 ? '0' : '') + h + ':' + (m < 10 ? '0' : '') + m;
  }
  return String(val || '');
}

/**
 * Lee la tabla de trabajos (filas 100-129) y devuelve los que tienen
 * el checkbox marcado como TRUE. Usa lecturas en batch para eficiencia.
 */
function _getSelectedJobs_(sh) {
  var jobs  = [];
  var START = JOB_TABLE.START_ROW;
  var MAX   = JOB_TABLE.MAX_ROWS;

  var checks = sh.getRange(START, JOB_TABLE.COL_CHECK, MAX, 1).getValues();
  var jsons  = sh.getRange(START, JOB_TABLE.COL_DATA,  MAX, 1).getValues();

  for (var i = 0; i < MAX; i++) {
    if (checks[i][0] !== true) continue;
    var jsonStr = String(jsons[i][0] || '').trim();
    if (!jsonStr) continue;
    try {
      var obj = JSON.parse(jsonStr);
      if (typeof obj.fechaTrabajo === 'string') {
        obj.fechaTrabajo = _parseReportDate_(obj.fechaTrabajo) || new Date();
      }
      jobs.push(obj);
    } catch (e) {
      Logger.log('_getSelectedJobs_: JSON parse error fila ' + (START + i) + ': ' + e.message);
    }
  }
  return jobs;
}
