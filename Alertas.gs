// ================================================================
// Alertas.gs — Alertas internas diarias de operaciones pendientes
// FASE 14
// ================================================================
// Detecta y notifica por email a empresa_email:
//   1. Trabajos vencidos (Fecha_Trabajo pasada y no completados)
//   2. Facturas emitidas sin enviar al cliente pasados N días
//   3. Gastos sin clasificar (Tipo_Gasto vacío) más antiguos de N días

// ----------------------------------------------------------------
// Entrada pública
// ----------------------------------------------------------------

function enviarAlertasDiarias() {
  var activo = getConfig('alert_activo');
  if (activo === '0' || activo === 'false') {
    logEvent('INFO', 'enviarAlertasDiarias', 'Alertas desactivadas (alert_activo=0)', '');
    return;
  }

  var ss           = getTargetSS_();
  var hoy          = new Date();
  var empresa      = getConfig('empresa_nombre')  || 'Empresa';
  var emailDest    = getConfig('alert_email_destino') || getConfig('empresa_email') || Session.getActiveUser().getEmail();
  var diasFacSin   = _alertCfgInt_('alert_dias_factura_sin_enviar', 3);
  var diasGasSin   = _alertCfgInt_('alert_dias_gasto_sin_tipo',     5);

  var alertas = _alertasRecopilar_(ss, hoy, diasFacSin, diasGasSin);

  var total = alertas.trabajos.length + alertas.facturas.length + alertas.gastos.length;
  if (total === 0) {
    logEvent('INFO', 'enviarAlertasDiarias', 'Sin alertas — no se envía email', '');
    _alertUi_('✅ Sin alertas pendientes. No se envía email.');
    return;
  }

  var htmlBody = _alertasRenderHtml_(alertas, empresa, hoy);
  var subject  = '⚠️ ' + total + ' alerta(s) pendiente(s) — ' + empresa;

  GmailApp.sendEmail(emailDest, subject, '', {
    htmlBody: htmlBody,
    name:     empresa
  });

  logEvent('INFO', 'enviarAlertasDiarias',
    'Email enviado a ' + emailDest + ' — trabajos:' + alertas.trabajos.length
    + ' facturas:' + alertas.facturas.length + ' gastos:' + alertas.gastos.length, '');
  _alertUi_('✅ Alertas enviadas a ' + emailDest + '\n\n'
    + '• Trabajos vencidos: '       + alertas.trabajos.length + '\n'
    + '• Facturas sin enviar: '     + alertas.facturas.length + '\n'
    + '• Gastos sin clasificar: '   + alertas.gastos.length);
}

// ----------------------------------------------------------------
// Recopilación de alertas
// ----------------------------------------------------------------

function _alertasRecopilar_(ss, hoy, diasFacSin, diasGasSin) {
  var alertas = { trabajos: [], facturas: [], gastos: [] };
  var tz      = ss.getSpreadsheetTimeZone();

  // ── 1. Trabajos vencidos ──────────────────────────────────────
  var traSheet = ss.getSheetByName(SHEET_NAMES.TRABAJOS);
  if (traSheet && traSheet.getLastRow() > 1) {
    var traData = traSheet.getRange(2, 1, traSheet.getLastRow() - 1, COL_TRA.ESTADO).getValues();
    var estadosOk = { 'completado': true, 'cancelado': true };

    for (var i = 0; i < traData.length; i++) {
      var t       = traData[i];
      var idTra   = String(t[COL_TRA.ID - 1]             || '').trim();
      var fecTra  = t[COL_TRA.FECHA_TRABAJO - 1];
      var estadoT = String(t[COL_TRA.ESTADO - 1]         || '').toLowerCase().trim();
      var cliente = String(t[COL_TRA.CLIENTE - 1]        || '').trim();
      var trabajo = String(t[COL_TRA.TRABAJO_REALIZADO - 1] || '').trim();

      if (!idTra) continue;
      if (estadosOk[estadoT]) continue;
      if (!(fecTra instanceof Date)) continue;

      // Pasada y no completada
      var diasRetraso = Math.floor((hoy.getTime() - fecTra.getTime()) / 86400000);
      if (diasRetraso > 0) {
        alertas.trabajos.push({
          id:      idTra,
          cliente: cliente,
          trabajo: trabajo.substring(0, 60) || '(sin descripción)',
          estado:  estadoT,
          fecha:   Utilities.formatDate(fecTra, tz, 'dd/MM/yyyy'),
          dias:    diasRetraso
        });
      }
    }
  }

  // ── 2. Facturas emitidas sin enviar ───────────────────────────
  var facSheet = ss.getSheetByName(SHEET_NAMES.FACTURAS);
  if (facSheet && facSheet.getLastRow() > 1) {
    var facData = facSheet.getRange(2, 1, facSheet.getLastRow() - 1, N_COLS_FAC).getValues();
    // Col 20 = Fecha_Envio (base-1: index 19)
    var IDX_FECHA_ENVIO = 19;

    for (var j = 0; j < facData.length; j++) {
      var f       = facData[j];
      var idFac   = String(f[CREC.ID_FACTURA - 1] || '').trim();
      var fecEmi  = f[CREC.FECHA_EMISION - 1];
      var estadoF = String(f[CREC.ESTADO - 1]     || '').toLowerCase().trim();
      var fecEnv  = f[IDX_FECHA_ENVIO];
      var nombre  = String(f[CREC.NOMBRE_CLIENTE - 1] || '').trim();
      var total   = parseFloat(f[CFAC.TOTAL - 1] || 0);

      if (!idFac) continue;
      if (/anulad|cancel|cobrad|pagad/i.test(estadoF)) continue;
      if (fecEnv instanceof Date) continue; // Ya fue enviada

      if (!(fecEmi instanceof Date)) continue;
      var diasEmi = Math.floor((hoy.getTime() - fecEmi.getTime()) / 86400000);
      if (diasEmi < diasFacSin) continue;

      alertas.facturas.push({
        id:      idFac,
        cliente: nombre,
        total:   total,
        fecha:   Utilities.formatDate(fecEmi, tz, 'dd/MM/yyyy'),
        dias:    diasEmi
      });
    }
  }

  // ── 3. Gastos sin clasificar ──────────────────────────────────
  var gasSheet = ss.getSheetByName('GASTOS');
  if (gasSheet && gasSheet.getLastRow() > 1) {
    var gasData = gasSheet.getRange(2, 1, gasSheet.getLastRow() - 1, COL_GAS.TOTAL).getValues();

    for (var k = 0; k < gasData.length; k++) {
      var g      = gasData[k];
      var idGas  = String(g[COL_GAS.ID_GASTO - 1]  || '').trim();
      var tipo   = String(g[COL_GAS.TIPO_GASTO - 1] || '').trim();
      var fecAlt = g[COL_GAS.FECHA_ALTA - 1];
      var total2 = parseFloat(g[COL_GAS.TOTAL - 1]  || 0);
      var conc   = String(g[COL_GAS.CONCEPTO - 1]   || '').trim();

      if (!idGas) continue;
      if (tipo)   continue; // Ya clasificado

      if (!(fecAlt instanceof Date)) continue;
      var diasAlt = Math.floor((hoy.getTime() - fecAlt.getTime()) / 86400000);
      if (diasAlt < diasGasSin) continue;

      alertas.gastos.push({
        id:      idGas,
        concepto: conc.substring(0, 60) || '(sin concepto)',
        total:   total2,
        dias:    diasAlt
      });
    }
  }

  return alertas;
}

// ----------------------------------------------------------------
// Renderizado HTML
// ----------------------------------------------------------------

function _alertasRenderHtml_(a, empresa, hoy) {
  var tz    = 'Europe/Madrid';
  var fecha = Utilities.formatDate(hoy, tz, 'dd/MM/yyyy HH:mm');

  function sec(titulo, color, filas) {
    if (!filas.length) return '';
    var rows = filas.map(function(f, idx) {
      return '<tr style="background:' + (idx % 2 ? '#f5f5f5' : '#fff') + ';">'
        + '<td style="padding:7px 14px;font-size:12px;color:#333;">' + f + '</td>'
        + '</tr>';
    }).join('');
    return '<table style="width:100%;border-collapse:collapse;background:#fff;border:1px solid #ddd;'
      + 'border-radius:4px;margin-bottom:14px;">'
      + '<tr style="background:' + color + ';color:#fff;">'
      + '<th style="padding:9px 14px;text-align:left;font-size:13px;">'
      + titulo + ' <span style="font-weight:normal;opacity:0.85;">(' + filas.length + ')</span></th>'
      + '</tr>'
      + rows + '</table>';
  }

  var trabaRen = a.trabajos.map(function(t) {
    return '<strong>' + _altEsc_(t.id) + '</strong> · '
      + _altEsc_(t.cliente) + ' — ' + _altEsc_(t.trabajo)
      + ' · fecha: ' + t.fecha
      + ' · <span style="color:#e74c3c;font-weight:bold;">' + t.dias + ' día(s) de retraso</span>';
  });

  function eur(v) { return _rnd2_(v).toFixed(2).replace('.', ',') + ' €'; }

  var facRen = a.facturas.map(function(f) {
    return '<strong>' + _altEsc_(f.id) + '</strong> · '
      + _altEsc_(f.cliente) + ' — ' + eur(f.total)
      + ' · emitida: ' + f.fecha
      + ' · <span style="color:#e67e22;font-weight:bold;">sin enviar hace ' + f.dias + ' días</span>';
  });

  var gasRen = a.gastos.map(function(g) {
    return '<strong>' + _altEsc_(g.id) + '</strong> · '
      + _altEsc_(g.concepto) + ' — ' + eur(g.total)
      + ' · <span style="color:#8e44ad;font-weight:bold;">sin tipo hace ' + g.dias + ' días</span>';
  });

  var total = a.trabajos.length + a.facturas.length + a.gastos.length;

  return [
    '<div style="font-family:Arial,sans-serif;max-width:640px;margin:0 auto;background:#f4f6f9;',
    '  border:1px solid #ddd;border-radius:8px;overflow:hidden;">',
    '  <div style="background:#c0392b;color:#fff;padding:20px 26px;">',
    '    <h2 style="margin:0;font-size:20px;">⚠️ ' + _altEsc_(empresa) + '</h2>',
    '    <p style="margin:5px 0 0;font-size:14px;opacity:0.9;">' + total + ' alerta(s) pendiente(s) — ' + fecha + '</p>',
    '  </div>',
    '  <div style="padding:22px;">',
    sec('🔧 Trabajos vencidos',       '#c0392b', trabaRen),
    sec('📄 Facturas sin enviar',     '#e67e22', facRen),
    sec('💸 Gastos sin clasificar',   '#8e44ad', gasRen),
    '    <p style="color:#999;font-size:11px;margin-top:12px;">',
    'Para desactivar estas alertas, pon alert_activo=0 en la hoja CONFIG.',
    '    </p>',
    '  </div>',
    '</div>'
  ].join('\n');
}

// ----------------------------------------------------------------
// Trigger
// ----------------------------------------------------------------

function instalarTriggerAlertas() {
  _alertEnsureConfig_();
  var ya = ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === 'enviarAlertasDiarias';
  });
  if (!ya) {
    ScriptApp.newTrigger('enviarAlertasDiarias')
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .nearMinute(30)
      .create();
    logEvent('INFO', 'instalarTriggerAlertas', 'Trigger diario a las 8:30h instalado', '');
    SpreadsheetApp.getUi().alert(
      '✅ Trigger de alertas instalado.\n\nSe comprobará cada día a las 8:30h.\n'
      + 'Solo se envía email si hay alertas pendientes.'
    );
  } else {
    SpreadsheetApp.getUi().alert('ℹ️ El trigger ya estaba instalado.');
  }
}

function desinstalarTriggerAlertas() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'enviarAlertasDiarias') { ScriptApp.deleteTrigger(t); n++; }
  });
  logEvent('INFO', 'desinstalarTriggerAlertas', 'Triggers eliminados: ' + n, '');
  SpreadsheetApp.getUi().alert(n > 0 ? '✅ Trigger desinstalado.' : 'ℹ️ No había trigger activo.');
}

// ----------------------------------------------------------------
// Utilidades internas
// ----------------------------------------------------------------

function _alertEnsureConfig_() {
  var defaults = [
    ['alert_activo',                   '1'],
    ['alert_dias_factura_sin_enviar',  '3'],
    ['alert_dias_gasto_sin_tipo',      '5'],
    ['alert_email_destino',            '']
  ];
  defaults.forEach(function(p) {
    if (getConfig(p[0]) === null) setConfig(p[0], p[1]);
  });
}

function _alertCfgInt_(key, def) {
  var v = parseInt(getConfig(key), 10);
  return isNaN(v) ? def : v;
}

function _altEsc_(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function _alertUi_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (_) {}
}
