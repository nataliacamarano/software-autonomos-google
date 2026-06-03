// ============================================================
// Dashboard.gs — Validación de datos para Looker Studio
// ============================================================
//
// SOLO LECTURA — no crea hojas, no modifica datos, no toca formularios.
//
// Ejecutar: validateDashboardData()
// ============================================================

/**
 * Verifica que todas las hojas tienen los campos necesarios para Looker Studio.
 * Genera en el log: diagnóstico del spreadsheet, estado de columnas,
 * métricas y dimensiones recomendadas, paneles sugeridos y pasos de conexión.
 */
function validateDashboardData() {
  Logger.log('================================================');
  Logger.log('validateDashboardData()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('VERIFICACIÓN PARA LOOKER STUDIO — SOLO LECTURA');
  Logger.log('================================================');

  var ss    = getTargetSS_();
  var allOk = true;

  // ── DIAGNÓSTICO: spreadsheet y hojas disponibles ──────────────
  Logger.log('\n[DIAGNÓSTICO] Spreadsheet:');
  Logger.log('  Nombre: ' + ss.getName());
  Logger.log('  ID:     ' + ss.getId());
  var allSheets = ss.getSheets();
  Logger.log('  Hojas encontradas (' + allSheets.length + '):');
  allSheets.forEach(function(sh) {
    Logger.log('    · "' + sh.getName() + '"  (filas: ' + sh.getLastRow() + ')');
  });

  // ── DIAGNÓSTICO: cabeceras reales de FACTURAS ─────────────────
  Logger.log('\n[DIAGNÓSTICO] Cabeceras reales de FACTURAS:');
  var facSheet = ss.getSheetByName('FACTURAS');
  if (facSheet && facSheet.getLastColumn() > 0) {
    var facHdrs = facSheet.getRange(1, 1, 1, facSheet.getLastColumn()).getValues()[0];
    facHdrs.forEach(function(h, i) {
      Logger.log('  col ' + (i + 1) + ': "' + h + '"');
    });
  } else {
    Logger.log('  ❌ Hoja FACTURAS vacía o no encontrada.');
  }

  // ── Helper: comprueba existencia + columnas requeridas ────────
  function auditSheet(sheetName, required) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log('  ❌ Hoja "' + sheetName + '" no encontrada en este spreadsheet.');
      allOk = false;
      return { rows: 0, found: [], missing: required.slice() };
    }
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var headers = lastCol > 0
      ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
          return String(h).trim();
        })
      : [];
    var found   = [];
    var missing = [];
    required.forEach(function(col) {
      var colLower = col.toLowerCase().trim();
      var hit = headers.some(function(h) {
        return h.toLowerCase().indexOf(colLower) !== -1;
      });
      if (hit) { found.push(col); } else { missing.push(col); }
      Logger.log('    ' + (hit ? '✓' : '❌') + ' ' + col);
    });
    if (missing.length) allOk = false;
    return { rows: Math.max(0, lastRow - 1), found: found, missing: missing };
  }

  // ── 1. CLIENTES ───────────────────────────────────────────────
  Logger.log('\n[1] CLIENTES — columnas necesarias:');
  var rCli = auditSheet('CLIENTES', [
    'ID_Cliente', 'Nombre', 'Apellidos', 'Empresa', 'Email', 'Activo'
  ]);
  Logger.log('  Total registros: ' + rCli.rows);

  // ── 2. OPERARIOS ──────────────────────────────────────────────
  Logger.log('\n[2] OPERARIOS — columnas necesarias:');
  var rOpe = auditSheet('OPERARIOS', [
    'ID_Operario', 'Nombre', 'Apellidos', 'Especialidad', 'Activo'
  ]);
  Logger.log('  Total registros: ' + rOpe.rows);

  // ── DIAGNÓSTICO: cabeceras reales de TRABAJOS ────────────────
  Logger.log('\n[DIAGNÓSTICO] Cabeceras reales de TRABAJOS:');
  var traSheet = ss.getSheetByName('TRABAJOS');
  if (traSheet && traSheet.getLastColumn() > 0) {
    var traHdrs = traSheet.getRange(1, 1, 1, traSheet.getLastColumn()).getValues()[0];
    traHdrs.forEach(function(h, i) {
      Logger.log('  col ' + (i + 1) + ': "' + h + '"');
    });
  } else {
    Logger.log('  ❌ Hoja TRABAJOS vacía o no encontrada.');
  }

  // ── 3. TRABAJOS ───────────────────────────────────────────────
  // Columnas reales (HEADERS_TRABAJOS_V2):
  //   ID, Fecha Alta, Cliente, Operario, Fecha Trabajo,
  //   Hora Inicio, Hora Fin, Trabajo Realizado, Incidencias,
  //   Estado, Fotos, Ubicación
  Logger.log('\n[3] TRABAJOS — columnas necesarias:');
  var rTra = auditSheet('TRABAJOS', [
    'Cliente', 'Operario', 'Estado',
    'Fecha Trabajo', 'Hora Inicio', 'Hora Fin', 'Trabajo Realizado'
  ]);
  Logger.log('  Total registros: ' + rTra.rows);
  if (rTra.missing.length === 0) {
    Logger.log('  ✓ Schema HEADERS_TRABAJOS_V2 compatible con Looker Studio.');
  } else {
    Logger.log('  ❌ Columnas faltantes: ' + rTra.missing.join(', '));
    Logger.log('  Compara con las cabeceras reales mostradas en [DIAGNÓSTICO] arriba.');
  }

  // ── 4. FACTURAS ───────────────────────────────────────────────
  Logger.log('\n[4] FACTURAS — columnas necesarias:');
  var rFac = auditSheet('FACTURAS', [
    'ID_Factura', 'Nombre_Cliente', 'Total',
    'Estado Cobro', 'Fecha_Emision', 'Base_Imponible', 'IVA_Importe'
  ]);
  Logger.log('  Total registros: ' + rFac.rows);

  // ── 5. GASTOS ─────────────────────────────────────────────────
  Logger.log('\n[5] GASTOS — columnas necesarias:');
  var rGas = auditSheet('GASTOS', [
    'Fecha', 'Total', 'IVA', 'Proveedor', 'Concepto', 'Estado Pago'
  ]);
  Logger.log('  Total registros: ' + rGas.rows);

  // ── 6. Informe Looker Studio ──────────────────────────────────
  Logger.log('\n================================================');
  Logger.log('[6] CONFIGURACIÓN RECOMENDADA PARA LOOKER STUDIO');
  Logger.log('================================================');

  Logger.log('\n  FUENTES DE DATOS (una por hoja):');
  Logger.log('    Looker Studio > Añadir datos > Google Sheets');
  Logger.log('    Selecciona este spreadsheet y elige la hoja.');
  Logger.log('    Activa "Usar primera fila como encabezado".');
  Logger.log('    Repite para TRABAJOS, FACTURAS y GASTOS.');

  Logger.log('\n  ── HOJA: TRABAJOS ────────────────────────────');
  Logger.log('  Dimensiones recomendadas:');
  Logger.log('    • Cliente            → filtrar y agrupar por cliente');
  Logger.log('    • Operario           → rendimiento por operario');
  Logger.log('    • Estado             → Pendiente / En curso / Completado');
  Logger.log('    • Fecha Trabajo      → eje temporal (semana / mes / año)');
  Logger.log('    • Hora Inicio        → franja horaria de inicio');
  Logger.log('    • Hora Fin           → franja horaria de fin');
  Logger.log('  Métricas recomendadas:');
  Logger.log('    • ID                 → COUNT → nº de trabajos realizados');
  Logger.log('  Campo calculado sugerido en Looker Studio:');
  Logger.log('    • Duración = HOUR(Hora Fin) − HOUR(Hora Inicio)  → horas por trabajo');

  Logger.log('\n  ── HOJA: FACTURAS ────────────────────────────');
  Logger.log('  Dimensiones recomendadas:');
  Logger.log('    • Nombre_Cliente     → agrupación por cliente');
  Logger.log('    • Estado Cobro       → Pendiente / Cobrado / Parcial');
  Logger.log('    • Fecha_Emision      → eje temporal');
  Logger.log('    • Serie              → agrupación por serie de factura');
  Logger.log('  Métricas recomendadas:');
  Logger.log('    • Base_Imponible     → SUM  → base total facturada');
  Logger.log('    • IVA_Importe        → SUM  → IVA repercutido (Out)');
  Logger.log('    • Total              → SUM  → total facturado');
  Logger.log('    • ID_Factura         → COUNT → nº de facturas emitidas');

  Logger.log('\n  ── HOJA: GASTOS ──────────────────────────────');
  Logger.log('  Dimensiones recomendadas:');
  Logger.log('    • Proveedor          → agrupación por proveedor');
  Logger.log('    • Concepto           → categoría de gasto');
  Logger.log('    • Estado Pago        → Pendiente / Pagado');
  Logger.log('    • Fecha              → eje temporal');
  Logger.log('  Métricas recomendadas:');
  Logger.log('    • Base Imponible     → SUM  → base total de gastos');
  Logger.log('    • IVA                → SUM  → IVA soportado (In)');
  Logger.log('    • Total              → SUM  → gasto total');
  Logger.log('    • ID Gasto           → COUNT → nº de gastos');

  Logger.log('\n  ── HOJA: CLIENTES ────────────────────────────');
  Logger.log('  Dimensiones recomendadas:');
  Logger.log('    • Nombre + Apellidos → identificación del cliente');
  Logger.log('    • Empresa            → agrupación empresarial');
  Logger.log('    • Activo             → filtro Sí/No');
  Logger.log('  Métricas recomendadas:');
  Logger.log('    • ID_Cliente         → COUNT → total clientes activos');

  Logger.log('\n  ── PANELES SUGERIDOS ─────────────────────────');
  Logger.log('');
  Logger.log('  Panel 1 — Resumen ejecutivo:');
  Logger.log('    KPI: trabajos este mes (TRABAJOS, filtro Fecha Trabajo = mes actual)');
  Logger.log('    KPI: total facturado   (FACTURAS.Total SUM)');
  Logger.log('    KPI: total cobrado     (FACTURAS.Total WHERE Estado Cobro = Cobrado)');
  Logger.log('    KPI: pendiente cobro   (FACTURAS.Total WHERE Estado Cobro = Pendiente)');
  Logger.log('    Gráfico líneas: ingresos por mes  (Fecha_Emision × Total)');
  Logger.log('    Gráfico barras: trabajos por estado');
  Logger.log('');
  Logger.log('  Panel 2 — Operarios:');
  Logger.log('    Tabla: Operario × COUNT(trabajos) × campo calculado Duración');
  Logger.log('    Gráfico barras: trabajos por operario por mes');
  Logger.log('    Filtro de fecha: Fecha Trabajo');
  Logger.log('');
  Logger.log('  Panel 3 — Clientes:');
  Logger.log('    Tabla: Cliente × COUNT(trabajos) × SUM(Total FACTURAS)');
  Logger.log('    Gráfico tarta: distribución de trabajos por cliente');
  Logger.log('    Filtro: Estado Cobro = Pendiente → deuda viva por cliente');
  Logger.log('');
  Logger.log('  Panel 4 — IVA / Fiscalidad (trimestral):');
  Logger.log('    KPI: IVA Out = SUM(FACTURAS.IVA_Importe)');
  Logger.log('    KPI: IVA In  = SUM(GASTOS.IVA)');
  Logger.log('    KPI: IVA Neto = IVA Out − IVA In');
  Logger.log('    Campo calculado: IVA_Neto = IVA_Importe (FACTURAS) − IVA (GASTOS)');
  Logger.log('    Filtro: trimestre (Q1/Q2/Q3/Q4) sobre Fecha_Emision / Fecha');
  Logger.log('    Tabla: mes × IVA Out × IVA In × Neto');

  Logger.log('\n  ── PASOS PARA CONECTAR ───────────────────────');
  Logger.log('  1. Ir a https://lookerstudio.google.com');
  Logger.log('  2. Crear informe > Añadir datos > Google Sheets');
  Logger.log('  3. Seleccionar este spreadsheet → hoja TRABAJOS');
  Logger.log('  4. Activar "Usar primera fila como encabezado" → Añadir');
  Logger.log('  5. Repetir para FACTURAS y GASTOS (fuentes adicionales)');
  Logger.log('  6. En el panel, combinar fuentes si necesitas cruzar datos');
  Logger.log('     (ej. TRABAJOS + FACTURAS por Nombre_Cliente)');

  // ── Resumen ───────────────────────────────────────────────────
  Logger.log('\n================================================');
  Logger.log('RESULTADO:');
  if (allOk) {
    Logger.log('  ✓ TODAS LAS HOJAS LISTAS PARA LOOKER STUDIO');
    Logger.log('  Conecta las fuentes en https://lookerstudio.google.com');
  } else {
    Logger.log('  ❌ HAY PROBLEMAS — revisa los ❌ y el DIAGNÓSTICO anteriores');
  }
  Logger.log('================================================');
}
