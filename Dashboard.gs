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

// ============================================================
// FASE 10 — Dashboard visual en hoja DASHBOARD
// ============================================================

// ── Entrada pública ──────────────────────────────────────────

function generarDashboard() {
  Logger.log('=== generarDashboard() ===');
  var ss   = getTargetSS_();
  var tz   = Session.getScriptTimeZone();
  var now  = new Date();
  var anyo = now.getFullYear();

  var datos = _db_recopilar_(ss, anyo, tz);

  var sh = ss.getSheetByName('DASHBOARD');
  if (sh) { sh.clearContents(); sh.clearFormats(); sh.clearDataValidations(); }
  else    { sh = ss.insertSheet('DASHBOARD'); }

  _db_render_(sh, datos, anyo, tz, now);
  ss.setActiveSheet(sh);

  logEvent('INFO', 'generarDashboard', 'Dashboard generado año=' + anyo, '');
  Logger.log('=== generarDashboard() OK ===');
}

// ── Recopilación de datos ────────────────────────────────────

function _db_recopilar_(ss, anyo, tz) {
  var d = {
    // Facturación
    facTotal: 0, facBase: 0, facIva: 0,
    facCobradoImporte: 0, facPendienteImporte: 0,
    facCobradas: 0, facPendientes: 0, facAnuladas: 0, nFac: 0,
    ultFac: [],
    // Gastos
    gasTotal: 0, gasBase: 0, gasIva: 0,
    gasPendienteImporte: 0, gasPendientePago: 0,
    gasRevisar: 0, gasOCRFallido: 0, gasSinTicket: 0, nGas: 0,
    ultGas: [],
    // Trabajos
    traPendientes: 0, traEnCurso: 0, traCompletados: 0, traFacturados: 0,
    traHoras: 0, nTra: 0,
    // Clientes / Operarios
    nClientes: 0, nOperarios: 0,
    // IVA anual (de IVA_TRIMESTRAL)
    ivaNeto: null, ivaResultado: null
  };

  var desde = new Date(anyo, 0, 1, 0, 0, 0);
  var hasta = new Date(anyo, 11, 31, 23, 59, 59);

  // ── FACTURAS ──────────────────────────────────────────────────
  try {
    var shF = ss.getSheetByName('FACTURAS');
    if (shF && shF.getLastRow() > 1) {
      var facRows = shF.getRange(2, 1, shF.getLastRow() - 1, N_COLS_FAC).getValues();
      var facPool = [];

      facRows.forEach(function(row) {
        if (!String(row[0]).trim()) return;
        var fEm    = row[CFAC.FECHA_EMISION - 1];
        if (!(fEm instanceof Date) || isNaN(fEm.getTime())) return;
        if (fEm < desde || fEm > hasta) return;

        var estado = String(row[CFAC.ESTADO - 1]).trim();
        d.nFac++;

        if (estado.toLowerCase() === 'anulada') { d.facAnuladas++; return; }

        var base  = parseFloat(row[CFAC.BASE_IMPONIBLE - 1]) || 0;
        var iva   = parseFloat(row[CFAC.IVA_IMPORTE    - 1]) || 0;
        var total = parseFloat(row[CFAC.TOTAL           - 1]) || 0;
        d.facBase  += base;
        d.facIva   += iva;
        d.facTotal += total;

        if (/cobrad|pagad/i.test(estado)) {
          d.facCobradas++;
          d.facCobradoImporte += total;
        } else {
          d.facPendientes++;
          d.facPendienteImporte += total;
        }

        facPool.push({ fecha: fEm, id: String(row[0]).trim(),
          cliente:  String(row[5]).trim(),
          concepto: String(row[10]).trim(),
          total:    total, estado: estado });
      });

      // Últimas 5 por fecha desc
      d.ultFac = facPool.sort(function(a, b) { return b.fecha - a.fecha; }).slice(0, 5);
    }
  } catch (e) { logEvent('WARN', '_db_recopilar_', 'FACTURAS: ' + e.message, ''); }

  // ── GASTOS ────────────────────────────────────────────────────
  try {
    var shG = ss.getSheetByName('GASTOS');
    if (shG && shG.getLastRow() > 1) {
      var gasRows = shG.getRange(2, 1, shG.getLastRow() - 1, HEADERS_GASTOS.length).getValues();
      var gasPool = [];

      gasRows.forEach(function(row) {
        if (String(row[COL_GAS.ID_GASTO - 1]).trim().indexOf('GAS-') !== 0) return;
        var fGas = row[COL_GAS.FECHA_GASTO - 1];
        var fRef = (fGas instanceof Date && !isNaN(fGas.getTime())) ? fGas : row[COL_GAS.FECHA_ALTA - 1];
        if (!(fRef instanceof Date) || isNaN(fRef.getTime())) return;
        if (fRef < desde || fRef > hasta) return;

        var base  = parseFloat(row[COL_GAS.BASE_IMPONIBLE - 1]) || 0;
        var iva   = parseFloat(row[COL_GAS.IVA_IMPORTE    - 1]) || 0;
        var total = parseFloat(row[COL_GAS.TOTAL           - 1]) || 0;
        d.gasBase  += base;
        d.gasIva   += iva;
        d.gasTotal += total;
        d.nGas++;

        var ep = String(row[COL_GAS.ESTADO_PAGO           - 1]).trim();
        var es = String(row[COL_GAS.ESTADO_PROCESAMIENTO  - 1]).trim();
        if (!/^pagado$/i.test(ep))    { d.gasPendientePago++;  d.gasPendienteImporte += total; }
        if (es === 'OCR fallido')       d.gasOCRFallido++;
        else if (es === 'Revisar')      d.gasRevisar++;
        if (!String(row[COL_GAS.FOTO_URL   - 1]).trim() &&
            !String(row[COL_GAS.ARCHIVO_ID - 1]).trim()) d.gasSinTicket++;

        gasPool.push({ fecha: fRef,
          id:        String(row[COL_GAS.ID_GASTO   - 1]).trim(),
          tipo:      String(row[COL_GAS.TIPO_GASTO  - 1]).trim(),
          proveedor: String(row[COL_GAS.PROVEEDOR   - 1]).trim() || '—',
          total:     total,
          estado:    ep });
      });

      d.ultGas = gasPool.sort(function(a, b) { return b.fecha - a.fecha; }).slice(0, 5);
    }
  } catch (e) { logEvent('WARN', '_db_recopilar_', 'GASTOS: ' + e.message, ''); }

  // ── TRABAJOS ──────────────────────────────────────────────────
  try {
    var shT = ss.getSheetByName('TRABAJOS');
    if (shT && shT.getLastRow() > 1) {
      var nColsT  = Math.min(shT.getLastColumn(), 19);
      var tHdrs   = shT.getRange(1, 1, 1, nColsT).getValues()[0]
                       .map(function(h) { return String(h).trim().toLowerCase(); });
      var cEst    = tHdrs.indexOf('estado');
      var cHoras  = tHdrs.indexOf('horas_trabajadas');
      var traRows = shT.getRange(2, 1, shT.getLastRow() - 1, nColsT).getValues();

      traRows.forEach(function(row) {
        if (!String(row[0]).trim()) return;
        d.nTra++;
        if (cEst >= 0) {
          var est = String(row[cEst]).trim().toLowerCase();
          if      (est === 'pendiente')               d.traPendientes++;
          else if (est === 'en curso')                d.traEnCurso++;
          else if (est.indexOf('complet') !== -1)     d.traCompletados++;
          else if (est === 'facturado')               d.traFacturados++;
        }
        if (cHoras >= 0) d.traHoras += parseFloat(row[cHoras]) || 0;
      });
    }
  } catch (e) { logEvent('WARN', '_db_recopilar_', 'TRABAJOS: ' + e.message, ''); }

  // ── CLIENTES ─────────────────────────────────────────────────
  try {
    var shC = ss.getSheetByName('CLIENTES');
    if (shC && shC.getLastRow() > 1) {
      shC.getRange(2, 10, shC.getLastRow() - 1, 1).getValues()
        .forEach(function(row) {
          var v = row[0];
          if (v === true || /^(sí|si|1|true)$/i.test(String(v))) d.nClientes++;
        });
    }
  } catch (e) { logEvent('WARN', '_db_recopilar_', 'CLIENTES: ' + e.message, ''); }

  // ── OPERARIOS ────────────────────────────────────────────────
  try {
    var shO = ss.getSheetByName('OPERARIOS');
    if (shO && shO.getLastRow() > 1) {
      shO.getRange(2, 8, shO.getLastRow() - 1, 1).getValues()
        .forEach(function(row) {
          var v = row[0];
          if (v === true || /^(sí|si|1|true)$/i.test(String(v))) d.nOperarios++;
        });
    }
  } catch (e) { logEvent('WARN', '_db_recopilar_', 'OPERARIOS: ' + e.message, ''); }

  // ── IVA_TRIMESTRAL ────────────────────────────────────────────
  try {
    var shI = ss.getSheetByName('IVA_TRIMESTRAL');
    if (shI && shI.getLastRow() > 1) {
      var ivaRows = shI.getRange(2, 1, shI.getLastRow() - 1, HEADERS_IVA_TRIM.length).getValues();
      var acum = 0, cnt = 0;
      ivaRows.forEach(function(row) {
        if (Number(row[CIVA.ANYO - 1]) === anyo) { acum += parseFloat(row[CIVA.IVA_NETO - 1]) || 0; cnt++; }
      });
      if (cnt > 0) {
        d.ivaNeto      = Math.round(acum * 100) / 100;
        d.ivaResultado = acum > 0.005 ? 'A pagar' : acum < -0.005 ? 'A compensar' : 'Sin diferencia';
      }
    }
  } catch (e) { logEvent('WARN', '_db_recopilar_', 'IVA_TRIMESTRAL: ' + e.message, ''); }

  // ── Redondeos finales ─────────────────────────────────────────
  ['facTotal','facBase','facIva','facCobradoImporte','facPendienteImporte',
   'gasTotal','gasBase','gasIva','gasPendienteImporte'].forEach(function(k) {
     d[k] = Math.round(d[k] * 100) / 100;
   });
  d.beneficio = Math.round((d.facTotal - d.gasTotal) * 100) / 100;
  d.traHoras  = Math.round(d.traHoras * 10) / 10;

  Logger.log('[_db_recopilar_] fac=' + d.nFac + ' gas=' + d.nGas
             + ' tra=' + d.nTra + ' cli=' + d.nClientes + ' ope=' + d.nOperarios);
  return d;
}

// ── Renderizado ──────────────────────────────────────────────

function _db_render_(sh, d, anyo, tz, now) {
  var NC      = 6;
  var r       = 1;
  var empresa = getConfig('empresa_nombre') || 'Álvaro Gestión';
  var EUR     = '#,##0.00 €';
  var NUM     = '#,##0';

  // Anchos: label(185) valor(115) × 3
  [185, 115, 185, 115, 185, 115].forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });

  // ── helpers ─────────────────────────────────────────────────
  function titulo(val, bg, fg, size) {
    sh.getRange(r, 1, 1, NC).merge()
      .setValue(val).setBackground(bg).setFontColor(fg || '#ffffff')
      .setFontWeight('bold').setFontSize(size || 12)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  }

  function seccion(val, bg, fg) {
    titulo(val, bg || '#e8f0fe', fg || '#1a73e8', 11);
    sh.setRowHeight(r, 28); r++;
  }

  function spacer(h) { sh.setRowHeight(r, h || 8); r++; }

  function kpiRow(items, bg) {
    if (bg) sh.getRange(r, 1, 1, NC).setBackground(bg);
    sh.setRowHeight(r, 32);
    items.forEach(function(it) {
      // it = [labelCol, label, val, fmt, fgColor]
      var lc = it[0], vc = it[0] + 1;
      sh.getRange(r, lc).setValue(it[1]).setFontSize(9)
        .setFontColor('#5f6368').setFontWeight('bold')
        .setVerticalAlignment('bottom');
      var cell = sh.getRange(r, vc).setFontSize(14).setFontWeight('bold')
                    .setVerticalAlignment('bottom');
      if (it[2] === null || it[2] === '') {
        cell.setValue('');
      } else if (typeof it[2] === 'number' && it[3]) {
        cell.setValue(it[2]).setNumberFormat(it[3]);
      } else {
        cell.setValue(it[2]);
      }
      if (it[4]) cell.setFontColor(it[4]);
    });
    r++;
  }

  function tableHdr(cols) {
    sh.getRange(r, 1, 1, NC)
      .setValues([cols])
      .setBackground('#3c4043').setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center');
    sh.setRowHeight(r, 22); r++;
  }

  function tableRow(vals, bg) {
    sh.getRange(r, 1, 1, NC).setBackground(bg || '#ffffff').setFontSize(10);
    vals.forEach(function(v, i) {
      var cell = sh.getRange(r, i + 1);
      if (v && typeof v === 'object' && v.v !== undefined) {
        cell.setValue(v.v);
        if (v.fmt)   cell.setNumberFormat(v.fmt);
        if (v.fg)    cell.setFontColor(v.fg);
        if (v.bold)  cell.setFontWeight('bold');
        if (v.align) cell.setHorizontalAlignment(v.align);
      } else {
        cell.setValue(v === null || v === undefined ? '' : v);
      }
    });
    r++;
  }

  // ── CABECERA ─────────────────────────────────────────────────
  titulo('🏠  DASHBOARD — ' + empresa.toUpperCase(), '#1a73e8', '#ffffff', 16);
  sh.setRowHeight(r, 46); r++;

  titulo('Año ' + anyo + '   ·   Actualizado: '
    + Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm'), '#f8f9fa', '#5f6368', 10);
  sh.setRowHeight(r, 20); r++;
  spacer(6);

  // ── FACTURACIÓN ───────────────────────────────────────────────
  seccion('💰  FACTURACIÓN — ' + anyo);
  kpiRow([
    [1, 'Total facturado',     d.facTotal,            EUR, '#137333'],
    [3, 'Cobrado',             d.facCobradoImporte,   EUR, '#137333'],
    [5, 'Pendiente de cobro',  d.facPendienteImporte, EUR,
        d.facPendienteImporte > 0 ? '#f29900' : '#202124']
  ], '#f8f9fa');
  kpiRow([
    [1, 'Base imponible',      d.facBase,        EUR],
    [3, 'IVA repercutido',     d.facIva,         EUR],
    [5, 'Anuladas',            d.facAnuladas,    NUM]
  ]);
  kpiRow([
    [1, 'Nº facturas',         d.nFac,           NUM],
    [3, 'Cobradas',            d.facCobradas,    NUM, '#137333'],
    [5, 'Pendientes',          d.facPendientes,  NUM,
        d.facPendientes > 0 ? '#f29900' : '#202124']
  ], '#f8f9fa');
  spacer(6);

  // ── GASTOS ────────────────────────────────────────────────────
  seccion('🧾  GASTOS — ' + anyo);
  kpiRow([
    [1, 'Total gastos',       d.gasTotal,            EUR, '#c5221f'],
    [3, 'Pendiente de pago',  d.gasPendienteImporte, EUR,
        d.gasPendienteImporte > 0 ? '#f29900' : '#202124'],
    [5, 'Nº gastos',          d.nGas,                NUM]
  ], '#f8f9fa');
  kpiRow([
    [1, 'Base imponible', d.gasBase,          EUR],
    [3, 'IVA soportado',  d.gasIva,           EUR],
    [5, 'Pend. de pago',  d.gasPendientePago, NUM,
        d.gasPendientePago > 0 ? '#f29900' : '#202124']
  ]);
  kpiRow([
    [1, 'Por revisar',   d.gasRevisar,    NUM, d.gasRevisar    > 0 ? '#f29900' : '#202124'],
    [3, 'OCR fallido',   d.gasOCRFallido, NUM, d.gasOCRFallido > 0 ? '#c5221f' : '#202124'],
    [5, 'Sin ticket',    d.gasSinTicket,  NUM, d.gasSinTicket  > 0 ? '#c5221f' : '#202124']
  ], '#f8f9fa');
  spacer(6);

  // ── RESULTADO ESTIMADO ────────────────────────────────────────
  seccion('📈  RESULTADO ESTIMADO — ' + anyo);
  var bColor = d.beneficio >= 0 ? '#137333' : '#c5221f';
  kpiRow([
    [1, 'Beneficio estimado', d.beneficio,  EUR, bColor],
    [3, 'Ingresos brutos',    d.facTotal,   EUR],
    [5, 'Gastos totales',     d.gasTotal,   EUR]
  ], '#f8f9fa');

  if (d.ivaNeto !== null) {
    var ivaFg = d.ivaResultado === 'A pagar' ? '#c5221f'
              : d.ivaResultado === 'A compensar' ? '#137333' : '#5f6368';
    kpiRow([
      [1, 'IVA neto acumulado', d.ivaNeto,       EUR, ivaFg],
      [3, 'Resultado IVA',      d.ivaResultado,  null, ivaFg],
      [5, '', '']
    ]);
  }
  spacer(6);

  // ── TRABAJOS ─────────────────────────────────────────────────
  seccion('🔧  TRABAJOS');
  kpiRow([
    [1, 'Pendientes',   d.traPendientes,  NUM, d.traPendientes > 0 ? '#f29900' : '#202124'],
    [3, 'En curso',     d.traEnCurso,     NUM, d.traEnCurso    > 0 ? '#1a73e8' : '#202124'],
    [5, 'Completados',  d.traCompletados, NUM]
  ], '#f8f9fa');
  kpiRow([
    [1, 'Facturados',    d.traFacturados, NUM],
    [3, 'Total',         d.nTra,          NUM],
    [5, 'Horas totales', d.traHoras,      '#,##0.0 h']
  ]);
  spacer(6);

  // ── CLIENTES Y OPERARIOS ─────────────────────────────────────
  seccion('👥  CLIENTES Y OPERARIOS');
  kpiRow([
    [1, 'Clientes activos',  d.nClientes,  NUM],
    [3, 'Operarios activos', d.nOperarios, NUM],
    [5, '', '']
  ], '#f8f9fa');
  spacer(6);

  // ── ÚLTIMAS 5 FACTURAS ────────────────────────────────────────
  if (d.ultFac.length > 0) {
    seccion('📄  ÚLTIMAS FACTURAS', '#202124', '#ffffff');
    tableHdr(['ID Factura', 'Fecha', 'Cliente', 'Concepto', 'Total', 'Estado']);

    d.ultFac.forEach(function(f, i) {
      var bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
      var ec = /cobrad|pagad/i.test(f.estado) ? '#137333' : '#f29900';
      tableRow([
        f.id,
        Utilities.formatDate(f.fecha, tz, 'dd/MM/yyyy'),
        f.cliente,
        f.concepto || '—',
        { v: f.total, fmt: EUR, align: 'right' },
        { v: f.estado, fg: ec, align: 'center' }
      ], bg);
    });
    spacer(6);
  }

  // ── ÚLTIMOS 5 GASTOS ─────────────────────────────────────────
  if (d.ultGas.length > 0) {
    seccion('🧾  ÚLTIMOS GASTOS', '#202124', '#ffffff');
    tableHdr(['ID Gasto', 'Fecha', 'Tipo', 'Proveedor', 'Total', 'Pago']);

    d.ultGas.forEach(function(g, i) {
      var bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
      var pc = /^pagado$/i.test(g.estado) ? '#137333' : '#c5221f';
      tableRow([
        g.id,
        Utilities.formatDate(g.fecha, tz, 'dd/MM/yyyy'),
        g.tipo,
        g.proveedor,
        { v: g.total, fmt: EUR, align: 'right' },
        { v: g.estado, fg: pc, align: 'center' }
      ], bg);
    });
    spacer(6);
  }

  // ── PIE ───────────────────────────────────────────────────────
  sh.getRange(r, 1, 1, NC).merge()
    .setValue('Beneficio estimado = ingresos − gastos (excluye IRPF). '
              + 'IVA neto requiere haber generado IVA_TRIMESTRAL. '
              + 'Actualizar: ⚙️ Sistema → 🏠 Actualizar dashboard.')
    .setFontColor('#9aa0a6').setFontSize(9).setHorizontalAlignment('center');
  sh.setRowHeight(r, 28);

  sh.setFrozenRows(2);
  SpreadsheetApp.flush();
  Logger.log('[_db_render_] ' + r + ' filas totales.');
}
