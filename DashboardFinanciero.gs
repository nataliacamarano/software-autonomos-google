// ============================================================
// DashboardFinanciero.gs — FASE 11: Dashboard financiero
// ============================================================
// Sheet: DASHBOARD_FINANCIERO
// Secciones:
//   1. Flujo mensual  (ingresos, gastos, beneficio, IVA por mes)
//   2. IVA trimestral (T1-T4 + total año)
//   3. Aging facturas (días pendientes de cobro)
//   4. Top clientes   (por facturación del año)
//   5. Gastos por categoría
// ============================================================

// ── Entrada pública ──────────────────────────────────────────

function generarDashboardFinanciero() {
  Logger.log('=== generarDashboardFinanciero() ===');
  var ss   = getTargetSS_();
  var tz   = Session.getScriptTimeZone();
  var now  = new Date();
  var anyo = now.getFullYear();

  var datos = _dfi_recopilar_(ss, anyo, tz, now);

  var sh = ss.getSheetByName('DASHBOARD_FINANCIERO');
  if (sh) { sh.clearContents(); sh.clearFormats(); sh.clearDataValidations(); }
  else    { sh = ss.insertSheet('DASHBOARD_FINANCIERO'); }

  _dfi_render_(sh, datos, anyo, tz, now);
  ss.setActiveSheet(sh);

  logEvent('INFO', 'generarDashboardFinanciero',
    'Dashboard financiero generado año=' + anyo, '');
  Logger.log('=== generarDashboardFinanciero() OK ===');
}

// ── Recopilación de datos ────────────────────────────────────

function _dfi_recopilar_(ss, anyo, tz, now) {
  var desde = new Date(anyo, 0,  1,  0,  0,  0);
  var hasta = new Date(anyo, 11, 31, 23, 59, 59);

  // ── Flujo mensual (12 posiciones) ────────────────────────────
  var meses = Array.apply(null, Array(12)).map(function(_, i) {
    return { mes: i, ingresos: 0, gastos: 0, ivaOut: 0, ivaIn: 0,
             nFac: 0, nGas: 0 };
  });

  // ── Aging facturas pendientes (todas, sin límite de año) ──────
  var agingBuckets = [
    { label: 'Al corriente',  min: 0,  max: 29,   n: 0, importe: 0 },
    { label: 'Atención',      min: 30, max: 59,   n: 0, importe: 0 },
    { label: 'Vencidas',      min: 60, max: 89,   n: 0, importe: 0 },
    { label: 'Crítico',       min: 90, max: 99999, n: 0, importe: 0 }
  ];

  // ── Top clientes (año actual) ─────────────────────────────────
  var cliMap = {};

  // ── FACTURAS ─────────────────────────────────────────────────
  try {
    var shF   = ss.getSheetByName('FACTURAS');
    var lastF = shF ? shF.getLastRow() : 0;
    if (lastF > 1) {
      var facData = shF.getRange(2, 1, lastF - 1, N_COLS_FAC).getValues();
      facData.forEach(function(row) {
        if (!String(row[0]).trim()) return;
        var estado = String(row[CFAC.ESTADO - 1]).trim();
        if (/^anulada$/i.test(estado)) return;

        var fEm   = row[CFAC.FECHA_EMISION - 1];
        if (!(fEm instanceof Date) || isNaN(fEm.getTime())) return;

        var total = parseFloat(row[CFAC.TOTAL           - 1]) || 0;
        var iva   = parseFloat(row[CFAC.IVA_IMPORTE     - 1]) || 0;
        var cobrada = /cobrad|pagad/i.test(estado);

        // Flujo mensual — solo año actual
        if (fEm >= desde && fEm <= hasta) {
          var m = fEm.getMonth();
          meses[m].ingresos += total;
          meses[m].ivaOut   += iva;
          meses[m].nFac++;

          // Top clientes
          var cli = String(row[5]).trim() || 'Sin nombre';
          if (!cliMap[cli]) cliMap[cli] = { n: 0, total: 0, cobrado: 0, pendiente: 0 };
          cliMap[cli].n++;
          cliMap[cli].total += total;
          if (cobrada) cliMap[cli].cobrado   += total;
          else         cliMap[cli].pendiente += total;
        }

        // Aging — todas las no cobradas, cualquier año
        if (!cobrada) {
          var dias = Math.floor((now - fEm) / 86400000);
          for (var a = 0; a < agingBuckets.length; a++) {
            if (dias >= agingBuckets[a].min && dias <= agingBuckets[a].max) {
              agingBuckets[a].n++;
              agingBuckets[a].importe += total;
              break;
            }
          }
        }
      });
    }
  } catch (e) { logEvent('WARN', '_dfi_recopilar_', 'FACTURAS: ' + e.message, ''); }

  // ── GASTOS ────────────────────────────────────────────────────
  var tipoMap = {};
  try {
    var shG   = ss.getSheetByName('GASTOS');
    var lastG = shG ? shG.getLastRow() : 0;
    if (lastG > 1) {
      var gasData = shG.getRange(2, 1, lastG - 1, HEADERS_GASTOS.length).getValues();
      gasData.forEach(function(row) {
        if (String(row[COL_GAS.ID_GASTO - 1]).trim().indexOf('GAS-') !== 0) return;
        var fGas = row[COL_GAS.FECHA_GASTO - 1];
        var fRef = (fGas instanceof Date && !isNaN(fGas.getTime()))
                   ? fGas : row[COL_GAS.FECHA_ALTA - 1];
        if (!(fRef instanceof Date) || isNaN(fRef.getTime())) return;
        if (fRef < desde || fRef > hasta) return;

        var total = parseFloat(row[COL_GAS.TOTAL           - 1]) || 0;
        var base  = parseFloat(row[COL_GAS.BASE_IMPONIBLE  - 1]) || 0;
        var iva   = parseFloat(row[COL_GAS.IVA_IMPORTE     - 1]) || 0;
        var tipo  = String(row[COL_GAS.TIPO_GASTO          - 1]).trim() || 'Sin tipo';

        var m = fRef.getMonth();
        meses[m].gastos += total;
        meses[m].ivaIn  += iva;
        meses[m].nGas++;

        if (!tipoMap[tipo]) tipoMap[tipo] = { n: 0, base: 0, iva: 0, total: 0 };
        tipoMap[tipo].n++;
        tipoMap[tipo].base  += base;
        tipoMap[tipo].iva   += iva;
        tipoMap[tipo].total += total;
      });
    }
  } catch (e) { logEvent('WARN', '_dfi_recopilar_', 'GASTOS: ' + e.message, ''); }

  // ── IVA trimestral (cálculo fresco) ──────────────────────────
  var ivaTrims = [];
  for (var t = 1; t <= 4; t++) {
    try { ivaTrims.push(_calcularTrimestre_(anyo, t)); }
    catch (e) {
      logEvent('WARN', '_dfi_recopilar_', 'IVA T' + t + ': ' + e.message, '');
      ivaTrims.push(null);
    }
  }

  // ── Post-proceso: meses ───────────────────────────────────────
  meses.forEach(function(m) {
    m.ingresos  = _rnd2_(m.ingresos);
    m.gastos    = _rnd2_(m.gastos);
    m.ivaOut    = _rnd2_(m.ivaOut);
    m.ivaIn     = _rnd2_(m.ivaIn);
    m.beneficio = _rnd2_(m.ingresos - m.gastos);
    m.ivaNeto   = _rnd2_(m.ivaOut   - m.ivaIn);
    m.margen    = m.ingresos > 0.01 ? m.beneficio / m.ingresos : null;
  });

  // Totales anuales (suma de meses con datos)
  var totM = meses.reduce(function(acc, m) {
    acc.ingresos += m.ingresos; acc.gastos  += m.gastos;
    acc.ivaOut   += m.ivaOut;  acc.ivaIn   += m.ivaIn;
    acc.nFac     += m.nFac;    acc.nGas    += m.nGas;
    return acc;
  }, { ingresos: 0, gastos: 0, ivaOut: 0, ivaIn: 0, nFac: 0, nGas: 0 });
  totM.beneficio = _rnd2_(totM.ingresos - totM.gastos);
  totM.ivaNeto   = _rnd2_(totM.ivaOut   - totM.ivaIn);
  totM.margen    = totM.ingresos > 0.01 ? totM.beneficio / totM.ingresos : null;

  // Aging totales
  agingBuckets.forEach(function(a) { a.importe = _rnd2_(a.importe); });
  var totAging = agingBuckets.reduce(function(acc, a) {
    acc.n += a.n; acc.importe = _rnd2_(acc.importe + a.importe); return acc;
  }, { n: 0, importe: 0 });

  // Top clientes (ordenados por total desc)
  var topCli = Object.keys(cliMap).map(function(k) {
    var c = cliMap[k];
    return { nombre: k, n: c.n,
             total:     _rnd2_(c.total),
             cobrado:   _rnd2_(c.cobrado),
             pendiente: _rnd2_(c.pendiente) };
  }).sort(function(a, b) { return b.total - a.total; }).slice(0, 8);
  var totCliTotal = _rnd2_(topCli.reduce(function(s, c) { return s + c.total; }, 0));

  // Tipos de gasto (ordenados por total desc)
  var topTipos = Object.keys(tipoMap).map(function(k) {
    var t = tipoMap[k];
    return { tipo: k, n: t.n,
             base:  _rnd2_(t.base),
             iva:   _rnd2_(t.iva),
             total: _rnd2_(t.total) };
  }).sort(function(a, b) { return b.total - a.total; });
  var totTiposTotal = _rnd2_(topTipos.reduce(function(s, t) { return s + t.total; }, 0));

  Logger.log('[_dfi_recopilar_] meses con datos='
    + meses.filter(function(m) { return m.ingresos || m.gastos; }).length
    + ' clientes=' + topCli.length
    + ' tiposGasto=' + topTipos.length);

  return {
    meses: meses, totM: totM,
    aging: agingBuckets, totAging: totAging,
    topCli: topCli, totCliTotal: totCliTotal,
    topTipos: topTipos, totTiposTotal: totTiposTotal,
    ivaTrims: ivaTrims
  };
}

// ── Renderizado ──────────────────────────────────────────────

function _dfi_render_(sh, d, anyo, tz, now) {
  var NC      = 8;
  var r       = 1;
  var EUR     = '#,##0.00 €';
  var NUM     = '#,##0';
  var empresa = getConfig('empresa_nombre') || 'Álvaro Gestión';
  var MESES_CORTO = ['Ene','Feb','Mar','Abr','May','Jun',
                     'Jul','Ago','Sep','Oct','Nov','Dic'];

  // Anchos de columna
  [145, 88, 88, 88, 82, 88, 88, 100].forEach(function(w, i) {
    sh.setColumnWidth(i + 1, w);
  });

  // ── Helpers de renderizado ─────────────────────────────────

  function bloque(val, bg, fg) {
    sh.getRange(r, 1, 1, NC).merge()
      .setValue(val)
      .setBackground(bg || '#e8f0fe').setFontColor(fg || '#1a73e8')
      .setFontWeight('bold').setFontSize(11)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.setRowHeight(r, 28); r++;
  }

  function cabecera(cols) {
    sh.getRange(r, 1, 1, NC)
      .setValues([cols.slice(0, NC)])
      .setBackground('#202124').setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center');
    sh.setRowHeight(r, 22); r++;
  }

  // Escribe una fila de datos. Cada celda puede ser:
  //   - valor primitivo (escrito directamente)
  //   - objeto { v, fmt, fg, bold, align }
  function fila(celdas, bg, negrita) {
    var rowRng = sh.getRange(r, 1, 1, NC);
    if (bg)      rowRng.setBackground(bg);
    if (negrita) rowRng.setFontWeight('bold');
    celdas.forEach(function(c, i) {
      if (i >= NC) return;
      var cell = sh.getRange(r, i + 1).setFontSize(10);
      if (c === null || c === undefined || c === '') { cell.setValue(''); return; }
      if (typeof c === 'object' && c !== null) {
        cell.setValue(c.v === undefined ? '' : c.v);
        if (c.fmt)   cell.setNumberFormat(c.fmt);
        if (c.fg)    cell.setFontColor(c.fg);
        if (c.bold)  cell.setFontWeight('bold');
        if (c.align) cell.setHorizontalAlignment(c.align);
      } else {
        cell.setValue(c);
      }
    });
    r++;
  }

  function sep(h) { sh.setRowHeight(r, h || 8); r++; }

  // Helpers de celda tipados
  function eur(v, fg)  { return { v: v, fmt: EUR, fg: fg || null, align: 'right' }; }
  function num(v, fg)  { return { v: v, fmt: NUM, fg: fg || null, align: 'right' }; }
  function pct(v, fg)  { return { v: v, fmt: '0.0%', fg: fg || null, align: 'right' }; }
  function cnt(v, fg)  { return { v: v, align: 'center', fg: fg || null }; }

  function colorBen(v) { return v >= 0 ? '#137333' : '#c5221f'; }
  function colorIva(v) { return v >  0 ? '#c5221f' : v < 0 ? '#137333' : '#5f6368'; }
  function colorRes(s) {
    return s === 'A pagar' ? '#c5221f' : s === 'A compensar' ? '#137333' : '#5f6368';
  }

  // ── CABECERA ──────────────────────────────────────────────────
  sh.getRange(r, 1, 1, NC).merge()
    .setValue('💹  DASHBOARD FINANCIERO — ' + empresa.toUpperCase())
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(16)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(r, 46); r++;

  sh.getRange(r, 1, 1, NC).merge()
    .setValue('Año ' + anyo + '   ·   Actualizado: '
      + Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm'))
    .setBackground('#f8f9fa').setFontColor('#5f6368').setFontSize(10)
    .setHorizontalAlignment('center');
  sh.setRowHeight(r, 20); r++;
  sep(8);

  // ── 1. FLUJO MENSUAL ──────────────────────────────────────────
  bloque('📅  FLUJO MENSUAL — ' + anyo);
  cabecera(['Mes', 'Ingresos', 'Gastos', 'Beneficio', 'Margen',
            'IVA OUT', 'IVA IN', 'IVA Neto']);

  var hayMeses = false;
  d.meses.forEach(function(m, i) {
    if (m.ingresos === 0 && m.gastos === 0) return;
    hayMeses = true;
    var bg  = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
    var cb  = colorBen(m.beneficio);
    var cin = colorIva(m.ivaNeto);
    fila([
      MESES_CORTO[m.mes],
      eur(m.ingresos,  '#137333'),
      eur(m.gastos,    '#c5221f'),
      eur(m.beneficio, cb),
      m.margen !== null ? pct(m.margen, cb) : cnt('—'),
      eur(m.ivaOut),
      eur(m.ivaIn),
      eur(m.ivaNeto, cin)
    ], bg);
  });

  if (!hayMeses) {
    sh.getRange(r, 1, 1, NC).merge()
      .setValue('Sin datos de facturación o gastos para ' + anyo)
      .setFontColor('#9aa0a6').setFontSize(10).setHorizontalAlignment('center');
    r++;
  }

  var tm = d.totM;
  fila([
    'TOTAL',
    eur(tm.ingresos,  '#137333'),
    eur(tm.gastos,    '#c5221f'),
    eur(tm.beneficio, colorBen(tm.beneficio)),
    tm.margen !== null ? pct(tm.margen, colorBen(tm.beneficio)) : cnt('—'),
    eur(tm.ivaOut),
    eur(tm.ivaIn),
    eur(tm.ivaNeto,   colorIva(tm.ivaNeto))
  ], '#e8f0fe', true);
  sep(8);

  // ── 2. IVA TRIMESTRAL ─────────────────────────────────────────
  bloque('🧮  IVA TRIMESTRAL — ' + anyo);
  cabecera(['Trimestre', 'Período', 'IVA OUT', 'IVA IN', 'IVA Neto',
            'Resultado', 'Nº Fac', 'Nº Gas']);

  var ivaAnual = 0, ivaOutAcum = 0, ivaInAcum = 0, nFacAcum = 0, nGasAcum = 0;
  d.ivaTrims.forEach(function(t, i) {
    if (!t) { fila(['T' + (i+1), '—', '', '', '', '', '', ''], '#ffffff'); return; }
    var bg  = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
    var cin = colorIva(t.ivaNeto);
    var res = t.resultado;
    var per = Utilities.formatDate(t.desde, tz, 'dd/MM')
            + ' – '
            + Utilities.formatDate(t.hasta, tz, 'dd/MM');
    fila([
      'T' + t.trimestre,
      per,
      eur(t.ivaOut),
      eur(t.ivaIn),
      eur(t.ivaNeto, cin),
      cnt(res, colorRes(res)),
      num(t.nFac),
      num(t.nGas)
    ], bg);
    ivaAnual   += t.ivaNeto;
    ivaOutAcum += t.ivaOut;
    ivaInAcum  += t.ivaIn;
    nFacAcum   += t.nFac;
    nGasAcum   += t.nGas;
  });

  ivaAnual   = _rnd2_(ivaAnual);
  ivaOutAcum = _rnd2_(ivaOutAcum);
  ivaInAcum  = _rnd2_(ivaInAcum);
  var resAnual = ivaAnual > 0.005 ? 'A pagar' : ivaAnual < -0.005 ? 'A compensar' : 'Sin diferencia';
  fila([
    'TOTAL', 'Año ' + anyo,
    eur(ivaOutAcum),
    eur(ivaInAcum),
    eur(ivaAnual, colorIva(ivaAnual)),
    cnt(resAnual, colorRes(resAnual)),
    num(nFacAcum),
    num(nGasAcum)
  ], '#e8f0fe', true);
  sep(8);

  // ── 3. AGING FACTURAS ─────────────────────────────────────────
  bloque('📋  FACTURAS PENDIENTES DE COBRO — ANTIGÜEDAD', '#fce8b2', '#b06000');
  cabecera(['Antigüedad', 'Días', 'Nº facturas', 'Importe',
            '% del pendiente', '', '', '']);

  var agingColores = ['#137333', '#f29900', '#e37400', '#c5221f'];
  var agingDias    = ['0 – 29 d', '30 – 59 d', '60 – 89 d', '≥ 90 d'];
  var totalPend    = d.totAging.importe;

  var hayAging = false;
  d.aging.forEach(function(a, i) {
    if (a.n === 0) return;
    hayAging = true;
    var bg  = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
    var fg  = agingColores[i];
    var p   = totalPend > 0 ? a.importe / totalPend : 0;
    fila([
      cnt(a.label, fg),
      cnt(agingDias[i], fg),
      num(a.n),
      eur(a.importe, fg),
      pct(p),
      '', '', ''
    ], bg);
  });

  if (!hayAging) {
    sh.getRange(r, 1, 1, NC).merge()
      .setValue('✅ No hay facturas pendientes de cobro')
      .setBackground('#e6f4ea').setFontColor('#137333').setFontSize(10)
      .setHorizontalAlignment('center');
    r++;
  } else {
    fila([
      cnt('TOTAL PENDIENTE'),
      '',
      num(d.totAging.n),
      eur(totalPend, '#c5221f'),
      pct(1),
      '', '', ''
    ], '#fce8b2', true);
  }
  sep(8);

  // ── 4. TOP CLIENTES ───────────────────────────────────────────
  if (d.topCli.length > 0) {
    bloque('🏆  TOP CLIENTES — ' + anyo);
    cabecera(['Cliente', 'Nº Facturas', 'Total facturado', 'Cobrado',
              'Pendiente', '% del total', '', '']);

    d.topCli.forEach(function(c, i) {
      var bg  = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
      var p   = d.totCliTotal > 0 ? c.total / d.totCliTotal : 0;
      var fgP = c.pendiente > 0 ? '#f29900' : '#202124';
      fila([
        c.nombre,
        num(c.n),
        eur(c.total),
        eur(c.cobrado,   c.cobrado   > 0 ? '#137333' : null),
        eur(c.pendiente, fgP),
        pct(p),
        '', ''
      ], bg);
    });
    sep(8);
  }

  // ── 5. GASTOS POR CATEGORÍA ───────────────────────────────────
  if (d.topTipos.length > 0) {
    bloque('🧾  GASTOS POR CATEGORÍA — ' + anyo);
    cabecera(['Categoría', 'Nº', 'Base imponible', 'IVA',
              'Total', '% del total', '', '']);

    var totT = d.totTiposTotal;
    var totN = 0, totBase = 0, totIva = 0;
    d.topTipos.forEach(function(t, i) {
      var bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
      var p  = totT > 0 ? t.total / totT : 0;
      fila([
        t.tipo,
        num(t.n),
        eur(t.base),
        eur(t.iva),
        eur(t.total, '#c5221f'),
        pct(p),
        '', ''
      ], bg);
      totN += t.n; totBase += t.base; totIva += t.iva;
    });
    fila([
      'TOTAL',
      num(totN),
      eur(_rnd2_(totBase)),
      eur(_rnd2_(totIva)),
      eur(totT, '#c5221f'),
      pct(1),
      '', ''
    ], '#f8e8e8', true);
    sep(8);
  }

  // ── PIE ───────────────────────────────────────────────────────
  sh.getRange(r, 1, 1, NC).merge()
    .setValue('IVA neto > 0 = "A pagar" (rojo). '
              + 'Aging: días desde Fecha_Emision en facturas no cobradas. '
              + 'Actualizar: ⚙️ Sistema → 💹 Dashboard financiero.')
    .setFontColor('#9aa0a6').setFontSize(9)
    .setHorizontalAlignment('center');
  sh.setRowHeight(r, 28);

  sh.setFrozenRows(2);
  SpreadsheetApp.flush();
  Logger.log('[_dfi_render_] ' + r + ' filas totales.');
}

// ── Utilidad privada ─────────────────────────────────────────

function _rnd2_(v) { return Math.round(v * 100) / 100; }
