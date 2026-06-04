// ================================================================
// InformeMensual.gs — Informe mensual automático al gestor
// FASE 13
// ================================================================

// ----------------------------------------------------------------
// Entradas públicas
// ----------------------------------------------------------------

/**
 * Llamada por trigger diario a las 8h.
 * Solo ejecuta el informe el día 1 de cada mes.
 */
function enviarInformeMensual() {
  if (new Date().getDate() !== 1) return;
  _infMensual_(-1);  // Mes anterior
}

/**
 * Llamada desde el menú: pregunta qué mes generar.
 */
function generarInformeMensual() {
  var ui   = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    '📊 Informe mensual',
    'Offset de mes (0 = actual, -1 = anterior, -2 = hace 2 meses):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var offset = parseInt(resp.getResponseText(), 10);
  if (isNaN(offset)) offset = -1;
  _infMensual_(offset);
}

// ----------------------------------------------------------------
// Lógica principal
// ----------------------------------------------------------------

function _infMensual_(mesOffset) {
  var ss           = getTargetSS_();
  var tz           = ss.getSpreadsheetTimeZone();
  var hoy          = new Date();
  var emailGestor  = getConfig('email_gestor')  || '';
  var empresa      = getConfig('empresa_nombre') || 'Empresa';
  var emailEmpresa = getConfig('empresa_email')  || Session.getActiveUser().getEmail();

  if (!emailGestor) {
    _infAlerta_('⚠️ email_gestor no configurado en CONFIG.\nEl informe no puede enviarse.');
    logEvent('WARN', '_infMensual_', 'email_gestor vacío', '');
    return;
  }

  var target   = new Date(hoy.getFullYear(), hoy.getMonth() + mesOffset, 1);
  var mesInicio = new Date(target.getFullYear(), target.getMonth(),     1);
  var mesFin    = new Date(target.getFullYear(), target.getMonth() + 1, 0, 23, 59, 59);
  var mesStr    = Utilities.formatDate(mesInicio, tz, 'MMMM yyyy');

  var d        = _infRecopilar_(ss, mesInicio, mesFin, tz);
  var htmlBody = _infRenderHtml_(d, mesStr, empresa, emailEmpresa, tz);
  var subject  = '📊 Informe mensual ' + mesStr + ' — ' + empresa;

  GmailApp.sendEmail(emailGestor, subject, '', {
    htmlBody: htmlBody,
    name:     empresa,
    replyTo:  emailEmpresa
  });

  logEvent('INFO', '_infMensual_', 'Informe enviado a ' + emailGestor + ' (' + mesStr + ')', '');
  _infAlerta_('✅ Informe de ' + mesStr + ' enviado a ' + emailGestor);
}

// ----------------------------------------------------------------
// Recopilación de datos
// ----------------------------------------------------------------

function _infRecopilar_(ss, inicio, fin, tz) {
  var d = {
    facEmitidas: 0,  facBase: 0,    facIVA: 0,    facTotal: 0,
    facCobradas: 0,  facTotalCobrado: 0,
    facPendientes: 0, facTotalPendiente: 0,
    gastos: 0,       gasBase: 0,   gasIVA: 0,    gasTotal: 0,
    resultado: 0,
    topClientes:  {},
    topGastos:    {}
  };

  // ── FACTURAS ──────────────────────────────────────────────────
  var facSheet = ss.getSheetByName(SHEET_NAMES.FACTURAS);
  if (facSheet && facSheet.getLastRow() > 1) {
    var facLastCol = facSheet.getLastColumn();
    var facHdrs    = facSheet.getRange(1, 1, 1, facLastCol).getValues()[0];
    var facData    = facSheet.getRange(2, 1, facSheet.getLastRow() - 1, facLastCol).getValues();

    // Columnas dinámicas del módulo de cobros
    var colEstCobro = -1, colFecCobro = -1;
    for (var h = 0; h < facHdrs.length; h++) {
      var hh = String(facHdrs[h]).trim();
      if (hh === 'Estado Cobro') colEstCobro = h;
      if (hh === 'Fecha Cobro')  colFecCobro = h;
    }

    for (var i = 0; i < facData.length; i++) {
      var r       = facData[i];
      var idFac   = String(r[CREC.ID_FACTURA - 1] || '').trim();
      if (!idFac) continue;

      var estado  = String(r[CREC.ESTADO - 1] || '').toLowerCase();
      if (/anulad|cancel/i.test(estado)) continue;

      var fecEmi  = r[CREC.FECHA_EMISION - 1];
      var base    = parseFloat(r[CFAC.BASE_IMPONIBLE - 1] || 0);
      var ivaImp  = parseFloat(r[CFAC.IVA_IMPORTE - 1]    || 0);
      var total   = parseFloat(r[CFAC.TOTAL - 1]           || 0);
      var nombre  = String(r[CREC.NOMBRE_CLIENTE - 1] || '').trim();

      var fecCobro   = colFecCobro   >= 0 ? r[colFecCobro]   : r[CREC.FECHA_PAGO - 1];
      var estCobro   = colEstCobro   >= 0 ? String(r[colEstCobro] || '').trim() : '';

      // Emitidas este mes
      if (fecEmi instanceof Date && fecEmi >= inicio && fecEmi <= fin) {
        d.facEmitidas++;
        d.facBase  += base;
        d.facIVA   += ivaImp;
        d.facTotal += total;
        if (nombre) d.topClientes[nombre] = (d.topClientes[nombre] || 0) + total;
      }

      // Cobradas este mes
      if (fecCobro instanceof Date && fecCobro >= inicio && fecCobro <= fin) {
        d.facCobradas++;
        d.facTotalCobrado += total;
      }

      // Pendientes de cobro (global, sin filtro de mes)
      var cobrada = /cobrad|pagad/i.test(estCobro) || (fecCobro instanceof Date);
      if (!cobrada) {
        d.facPendientes++;
        d.facTotalPendiente += total;
      }
    }
  }

  // ── GASTOS ────────────────────────────────────────────────────
  var gasSheet = ss.getSheetByName('GASTOS');
  if (gasSheet && gasSheet.getLastRow() > 1) {
    var gasData = gasSheet.getRange(2, 1, gasSheet.getLastRow() - 1, COL_GAS.TOTAL).getValues();
    for (var j = 0; j < gasData.length; j++) {
      var g      = gasData[j];
      var fecGas = g[COL_GAS.FECHA_GASTO - 1] instanceof Date
        ? g[COL_GAS.FECHA_GASTO - 1]
        : g[COL_GAS.FECHA_ALTA - 1];
      if (!(fecGas instanceof Date) || fecGas < inicio || fecGas > fin) continue;

      var gBase  = parseFloat(g[COL_GAS.BASE_IMPONIBLE - 1] || 0);
      var gIVA   = parseFloat(g[COL_GAS.IVA_IMPORTE - 1]    || 0);
      var gTotal = parseFloat(g[COL_GAS.TOTAL - 1]           || 0);
      var tipo   = String(g[COL_GAS.TIPO_GASTO - 1] || '').trim() || 'Sin clasificar';

      d.gastos++;
      d.gasBase  += gBase;
      d.gasIVA   += gIVA;
      d.gasTotal += gTotal;
      d.topGastos[tipo] = (d.topGastos[tipo] || 0) + gTotal;
    }
  }

  d.resultado = _rnd2_(d.facTotal - d.gasTotal);
  d.facBase   = _rnd2_(d.facBase);
  d.facIVA    = _rnd2_(d.facIVA);
  d.facTotal  = _rnd2_(d.facTotal);
  d.gasBase   = _rnd2_(d.gasBase);
  d.gasIVA    = _rnd2_(d.gasIVA);
  d.gasTotal  = _rnd2_(d.gasTotal);
  d.facTotalCobrado     = _rnd2_(d.facTotalCobrado);
  d.facTotalPendiente   = _rnd2_(d.facTotalPendiente);
  return d;
}

// ----------------------------------------------------------------
// Renderizado HTML
// ----------------------------------------------------------------

function _infRenderHtml_(d, mesStr, empresa, emailEmpresa, tz) {
  var resColor = d.resultado >= 0 ? '#27ae60' : '#e74c3c';
  var now = Utilities.formatDate(new Date(), tz || 'Europe/Madrid', 'dd/MM/yyyy HH:mm');

  function eur(v) { return _rnd2_(v).toFixed(2).replace('.', ',') + ' €'; }

  // Top clientes (máx 6, ordenados por total desc)
  var clientesHtml = '';
  var clientesArr = Object.keys(d.topClientes).map(function(k) {
    return { nombre: k, total: d.topClientes[k] };
  }).sort(function(a, b) { return b.total - a.total; }).slice(0, 6);
  if (clientesArr.length > 0) {
    clientesHtml = '<tr><td colspan="2" style="padding:14px 16px 4px;font-size:13px;font-weight:bold;color:#2c3e50;">Facturación por cliente</td></tr>';
    clientesArr.forEach(function(c, idx) {
      var bg = idx % 2 === 0 ? '#fff' : '#f5f5f5';
      clientesHtml += '<tr style="background:' + bg + ';">'
        + '<td style="padding:6px 16px;font-size:13px;color:#333;">' + _infEsc_(c.nombre) + '</td>'
        + '<td style="padding:6px 16px;font-size:13px;color:#333;text-align:right;">' + eur(c.total) + '</td>'
        + '</tr>';
    });
  }

  // Top tipos de gasto (máx 6)
  var gastosHtml = '';
  var gastosArr = Object.keys(d.topGastos).map(function(k) {
    return { tipo: k, total: d.topGastos[k] };
  }).sort(function(a, b) { return b.total - a.total; }).slice(0, 6);
  if (gastosArr.length > 0) {
    gastosHtml = '<tr><td colspan="2" style="padding:14px 16px 4px;font-size:13px;font-weight:bold;color:#2c3e50;">Gastos por categoría</td></tr>';
    gastosArr.forEach(function(g, idx) {
      var bg = idx % 2 === 0 ? '#fff' : '#f5f5f5';
      gastosHtml += '<tr style="background:' + bg + ';">'
        + '<td style="padding:6px 16px;font-size:13px;color:#333;">' + _infEsc_(g.tipo) + '</td>'
        + '<td style="padding:6px 16px;font-size:13px;color:#333;text-align:right;">' + eur(g.total) + '</td>'
        + '</tr>';
    });
  }

  return [
    '<div style="font-family:Arial,sans-serif;max-width:640px;margin:0 auto;background:#f4f6f9;',
    '  border:1px solid #ddd;border-radius:8px;overflow:hidden;">',

    // Header
    '  <div style="background:#1a1a2e;color:#fff;padding:22px 28px;">',
    '    <h2 style="margin:0;font-size:22px;">' + _infEsc_(empresa) + '</h2>',
    '    <p style="margin:6px 0 0;font-size:15px;opacity:0.85;">📊 Informe mensual — ' + mesStr + '</p>',
    '  </div>',

    '  <div style="padding:24px;">',

    // KPIs
    '    <table style="width:100%;border-collapse:collapse;margin-bottom:20px;">',
    '      <tr>',
    _infKpi_('Facturado', eur(d.facTotal),    '#2980b9'),
    _infKpi_('Cobrado',   eur(d.facTotalCobrado), '#27ae60'),
    _infKpi_('Gastos',    eur(d.gasTotal),    '#e67e22'),
    _infKpi_('Resultado', eur(d.resultado),   resColor),
    '      </tr>',
    '    </table>',

    // Facturas
    '    <table style="width:100%;border-collapse:collapse;background:#fff;border:1px solid #ddd;',
    '      border-radius:4px;margin-bottom:16px;">',
    '      <tr style="background:#2980b9;color:#fff;">',
    '        <th colspan="2" style="padding:10px 16px;text-align:left;font-size:13px;">Facturas del mes</th>',
    '      </tr>',
    _infFila_('Emitidas', d.facEmitidas + ' — base: ' + eur(d.facBase) + ' · IVA: ' + eur(d.facIVA) + ' · total: ' + eur(d.facTotal), false),
    _infFila_('Cobradas en el mes',  d.facCobradas + ' — ' + eur(d.facTotalCobrado), true),
    _infFila_('Pendientes de cobro (acumulado)', d.facPendientes + ' — ' + eur(d.facTotalPendiente), false),
    clientesHtml,
    '    </table>',

    // Gastos
    '    <table style="width:100%;border-collapse:collapse;background:#fff;border:1px solid #ddd;',
    '      border-radius:4px;margin-bottom:16px;">',
    '      <tr style="background:#e67e22;color:#fff;">',
    '        <th colspan="2" style="padding:10px 16px;text-align:left;font-size:13px;">Gastos del mes</th>',
    '      </tr>',
    _infFila_('Total gastos', d.gastos + ' — base: ' + eur(d.gasBase) + ' · IVA: ' + eur(d.gasIVA) + ' · total: ' + eur(d.gasTotal), false),
    gastosHtml,
    '    </table>',

    // Resultado
    '    <table style="width:100%;border-collapse:collapse;background:#fff;border:1px solid #ddd;border-radius:4px;margin-bottom:16px;">',
    '      <tr style="background:' + (d.resultado >= 0 ? '#27ae60' : '#e74c3c') + ';color:#fff;">',
    '        <th colspan="2" style="padding:10px 16px;text-align:left;font-size:13px;">Resultado estimado del mes</th>',
    '      </tr>',
    _infFila_('Facturado − Gastos', eur(d.facTotal) + ' − ' + eur(d.gasTotal) + ' = ' + eur(d.resultado), false),
    '      <tr><td colspan="2" style="padding:6px 16px;font-size:11px;color:#aaa;">* Resultado antes de IRPF e impuestos no reflejados en gastos.</td></tr>',
    '    </table>',

    '  </div>',

    // Footer
    '  <div style="background:#e8e8e8;padding:12px 24px;text-align:center;">',
    '    <p style="margin:0;color:#888;font-size:11px;">',
    _infEsc_(empresa) + ' &middot; ' + _infEsc_(emailEmpresa) + ' &middot; Generado ' + now,
    '    </p>',
    '  </div>',
    '</div>'
  ].join('\n');
}

function _infKpi_(label, valor, color) {
  return '        <td style="width:25%;padding:0 4px;">'
    + '<div style="background:#fff;border:1px solid #ddd;border-top:3px solid ' + color + ';'
    + 'border-radius:4px;padding:12px;text-align:center;">'
    + '<div style="font-size:11px;color:#888;margin-bottom:4px;">' + label + '</div>'
    + '<div style="font-size:16px;font-weight:bold;color:' + color + ';">' + valor + '</div>'
    + '</div></td>';
}

function _infFila_(label, valor, alt) {
  var bg = alt ? '#f5f5f5' : '#fff';
  return '      <tr style="background:' + bg + ';">'
    + '<td style="padding:8px 16px;font-size:13px;color:#555;font-weight:bold;width:40%;">' + label + ':</td>'
    + '<td style="padding:8px 16px;font-size:13px;color:#333;">' + valor + '</td>'
    + '</tr>';
}

// ----------------------------------------------------------------
// Trigger
// ----------------------------------------------------------------

function instalarTriggerMensual() {
  var ya = ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === 'enviarInformeMensual';
  });
  if (!ya) {
    ScriptApp.newTrigger('enviarInformeMensual')
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .create();
    logEvent('INFO', 'instalarTriggerMensual', 'Trigger diario instalado (ejecuta el día 1 de cada mes)', '');
    SpreadsheetApp.getUi().alert(
      '✅ Trigger mensual instalado.\n\nEl informe se enviará automáticamente el día 1 de cada mes a las 8h.'
    );
  } else {
    SpreadsheetApp.getUi().alert('ℹ️ El trigger ya estaba instalado.');
  }
}

function desinstalarTriggerMensual() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'enviarInformeMensual') {
      ScriptApp.deleteTrigger(t);
      n++;
    }
  });
  logEvent('INFO', 'desinstalarTriggerMensual', 'Triggers eliminados: ' + n, '');
  SpreadsheetApp.getUi().alert(n > 0 ? '✅ Trigger mensual desinstalado.' : 'ℹ️ No había trigger activo.');
}

// ----------------------------------------------------------------
// Utilidades internas
// ----------------------------------------------------------------

function _infEsc_(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function _infAlerta_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (_) {}
}
