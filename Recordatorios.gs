// ================================================================
// Recordatorios.gs — Recordatorios automáticos de facturas pendientes
// FASE 12
// ================================================================

// Columnas de FACTURAS usadas (1-indexed, alineadas con SHEET_HEADERS.FACTURAS)
var CREC = {
  ID_FACTURA:     1,
  FECHA_EMISION:  2,
  NOMBRE_CLIENTE: 6,
  EMAIL_CLIENTE:  7,
  CONCEPTO:       11,
  TOTAL:          17,
  ESTADO:         18,
  FECHA_PAGO:     21
};

var HEADERS_REC_LOG = [
  'ID_Recordatorio', 'Fecha_Envio', 'ID_Factura', 'Nombre_Cliente',
  'Email', 'Total', 'Fecha_Emision', 'Dias_Pendiente',
  'Tipo', 'Descripcion_Tipo', 'Resultado'
];

// ----------------------------------------------------------------
// Entrada principal — llamada por trigger diario y por menú
// ----------------------------------------------------------------

function enviarRecordatoriosPendientes() {
  var ss  = getTargetSS_();
  var hoy = new Date();

  // CONFIG
  var d1  = _recCfgInt_('rec_dias_1',              30);
  var d2  = _recCfgInt_('rec_dias_2',              60);
  var d3  = _recCfgInt_('rec_dias_3',              90);
  var dRC = _recCfgInt_('rec_dias_reenvio_critico', 30);

  var activo = getConfig('rec_activo');
  if (activo === '0' || activo === 'false') {
    logEvent('INFO', 'enviarRecordatoriosPendientes', 'Sistema desactivado en CONFIG (rec_activo=0)', '');
    _recAlerta_('ℹ️ Los recordatorios están desactivados.\nCambia rec_activo a 1 en la hoja CONFIG para activarlos.');
    return;
  }

  var empresa      = getConfig('empresa_nombre') || 'Empresa';
  var emailEmpresa = getConfig('empresa_email')  || Session.getActiveUser().getEmail();
  var datosPago    = getConfig('rec_datos_pago') || '';

  var logSheet = _rec_getOrCreateLogSheet_(ss);

  var facSheet = ss.getSheetByName(SHEET_NAMES.FACTURAS);
  if (!facSheet) {
    logEvent('WARN', 'enviarRecordatoriosPendientes', 'Hoja FACTURAS no encontrada', '');
    return;
  }

  var lastRow = facSheet.getLastRow();
  if (lastRow < 2) {
    _recAlerta_('No hay facturas en la hoja FACTURAS.');
    return;
  }

  var data = facSheet.getRange(2, 1, lastRow - 1, N_COLS_FAC).getValues();

  var enviados = 0, omitidos = 0, errores = 0;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];

    var idFac    = String(row[CREC.ID_FACTURA - 1]     || '').trim();
    var fecEmi   = row[CREC.FECHA_EMISION - 1];
    var nombre   = String(row[CREC.NOMBRE_CLIENTE - 1] || '').trim();
    var email    = String(row[CREC.EMAIL_CLIENTE - 1]  || '').trim();
    var concepto = String(row[CREC.CONCEPTO - 1]       || '').trim();
    var total    = row[CREC.TOTAL - 1] || 0;
    var estado   = String(row[CREC.ESTADO - 1]         || '').trim().toLowerCase();
    var fecPago  = row[CREC.FECHA_PAGO - 1];

    if (!idFac) continue;

    // Excluir anuladas o ya cobradas
    if (/anulad|cancel/i.test(estado)) continue;
    if (fecPago && fecPago !== '') continue;

    // Email necesario para enviar
    if (!email || email.indexOf('@') === -1) continue;

    // Fecha emisión válida
    if (!fecEmi || !(fecEmi instanceof Date)) continue;

    var dias = Math.floor((hoy.getTime() - fecEmi.getTime()) / 86400000);
    if (dias < d1) continue;

    // Determinar tipo
    var tipo, descTipo;
    if (dias >= d3) {
      tipo     = 3;
      descTipo = 'Aviso final (≥' + d3 + ' días)';
    } else if (dias >= d2) {
      tipo     = 2;
      descTipo = 'Segundo aviso (' + d2 + '–' + (d3 - 1) + ' días)';
    } else {
      tipo     = 1;
      descTipo = 'Primer aviso (' + d1 + '–' + (d2 - 1) + ' días)';
    }

    // Tipos 1 y 2: una sola vez por factura.
    // Tipo 3: repetir cada dRC días mientras siga impagada.
    var limiteDias = (tipo === 3) ? dRC : 9999;
    if (_rec_yaEnviado_(logSheet, idFac, tipo, limiteDias)) {
      omitidos++;
      continue;
    }

    var fac = {
      id: idFac, fechaEmision: fecEmi, nombre: nombre, email: email,
      concepto: concepto, total: total, dias: dias,
      tipo: tipo, descTipo: descTipo,
      empresa: empresa, emailEmpresa: emailEmpresa, datosPago: datosPago
    };

    var resultado;
    try {
      _rec_enviarEmail_(fac);
      resultado = 'Enviado';
      enviados++;
    } catch (e) {
      resultado = 'Error: ' + e.message;
      errores++;
      logEvent('ERROR', 'enviarRecordatoriosPendientes',
        'Error enviando a ' + email + ' (' + idFac + '): ' + e.message, '');
    }

    _rec_logEnvio_(logSheet, fac, resultado);
  }

  var resumen = 'Enviados: ' + enviados + ' | Omitidos: ' + omitidos + ' | Errores: ' + errores;
  logEvent('INFO', 'enviarRecordatoriosPendientes', resumen, '');
  _recAlerta_('✅ Recordatorios procesados\n\n' + resumen);
}

// ----------------------------------------------------------------
// Lógica de deduplicación
// ----------------------------------------------------------------

/**
 * Devuelve true si ya existe un envío exitoso para idFac+tipo dentro de los
 * últimos `limiteDias` días (usar 9999 para "nunca repetir").
 */
function _rec_yaEnviado_(logSheet, idFac, tipo, limiteDias) {
  var last = logSheet.getLastRow();
  if (last < 2) return false;

  var data = logSheet.getRange(2, 1, last - 1, HEADERS_REC_LOG.length).getValues();
  var hoy  = new Date();

  // Recorrer en orden inverso (más reciente primero)
  for (var i = data.length - 1; i >= 0; i--) {
    if (String(data[i][2] || '').trim() !== idFac)    continue; // ID_Factura
    if (data[i][8] !== tipo)                           continue; // Tipo
    if (!/^Enviado/i.test(String(data[i][10] || ''))) continue; // Resultado

    if (limiteDias >= 9999) return true;

    var fEnvio = data[i][1];
    if (fEnvio instanceof Date) {
      var diasDesde = (hoy.getTime() - fEnvio.getTime()) / 86400000;
      if (diasDesde < limiteDias) return true;
    }
  }
  return false;
}

// ----------------------------------------------------------------
// Envío de email
// ----------------------------------------------------------------

function _rec_enviarEmail_(fac) {
  var subject;
  if (fac.tipo === 3) {
    subject = '⚠️ Último aviso de pago — Factura ' + fac.id + ' — ' + fac.empresa;
  } else if (fac.tipo === 2) {
    subject = 'Recordatorio de pago — Factura ' + fac.id + ' — ' + fac.empresa;
  } else {
    subject = 'Aviso de pago pendiente — Factura ' + fac.id + ' — ' + fac.empresa;
  }

  var tz = 'Europe/Madrid';
  var fechaEmiStr = (fac.fechaEmision instanceof Date)
    ? Utilities.formatDate(fac.fechaEmision, tz, 'dd/MM/yyyy')
    : String(fac.fechaEmision);

  var totalNum = typeof fac.total === 'number' ? fac.total : parseFloat(fac.total) || 0;
  var totalStr = totalNum.toFixed(2).replace('.', ',') + ' €';

  var colorHeader = fac.tipo === 3 ? '#c0392b' : fac.tipo === 2 ? '#e67e22' : '#2980b9';

  var textoIntro;
  if (fac.tipo === 3) {
    textoIntro = 'Le comunicamos que la factura indicada a continuación lleva <strong>'
      + fac.dias + ' días</strong> pendiente de pago. Le rogamos que proceda al abono con carácter '
      + '<strong>urgente</strong> para evitar acciones adicionales.';
  } else if (fac.tipo === 2) {
    textoIntro = 'Le recordamos que la factura indicada a continuación se encuentra pendiente de pago '
      + 'desde hace <strong>' + fac.dias + ' días</strong>. Agradeceríamos que gestionara el pago a la mayor brevedad posible.';
  } else {
    textoIntro = 'Le enviamos este aviso porque la factura indicada a continuación lleva <strong>'
      + fac.dias + ' días</strong> pendiente de pago. Si ya ha realizado el pago, por favor ignore este mensaje.';
  }

  var pagoCells = fac.datosPago
    ? '<tr><td style="padding:8px 16px;color:#555;font-size:13px;font-weight:bold;width:40%;">Datos de pago:</td>'
      + '<td style="padding:8px 16px;color:#333;font-size:13px;white-space:pre-line;">' + fac.datosPago + '</td></tr>'
    : '';

  var htmlBody = [
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#f9f9f9;',
    '  border:1px solid #ddd;border-radius:6px;overflow:hidden;">',

    '  <div style="background:' + colorHeader + ';color:#fff;padding:20px 24px;">',
    '    <h2 style="margin:0;font-size:20px;">' + _recEsc_(fac.empresa) + '</h2>',
    '    <p style="margin:4px 0 0;font-size:14px;opacity:0.9;">',
    (fac.tipo === 3 ? '&#9888; Último aviso de pago' : '&#128276; Recordatorio de pago'),
    '    </p>',
    '  </div>',

    '  <div style="padding:24px;">',
    '    <p style="color:#333;font-size:14px;line-height:1.6;">',
    '      Estimado/a <strong>' + _recEsc_(fac.nombre) + '</strong>,',
    '    </p>',
    '    <p style="color:#555;font-size:14px;line-height:1.6;">' + textoIntro + '</p>',

    '    <table style="width:100%;border-collapse:collapse;margin:20px 0;background:#fff;',
    '      border:1px solid #ddd;border-radius:4px;">',
    '      <tr style="background:' + colorHeader + ';color:#fff;">',
    '        <th colspan="2" style="padding:10px 16px;text-align:left;font-size:13px;">',
    '          Detalle de la factura',
    '        </th>',
    '      </tr>',
    '      <tr>',
    '        <td style="padding:8px 16px;color:#555;font-size:13px;font-weight:bold;width:40%;">Nº Factura:</td>',
    '        <td style="padding:8px 16px;color:#333;font-size:13px;">' + _recEsc_(fac.id) + '</td>',
    '      </tr>',
    '      <tr style="background:#f5f5f5;">',
    '        <td style="padding:8px 16px;color:#555;font-size:13px;font-weight:bold;">Fecha de emisión:</td>',
    '        <td style="padding:8px 16px;color:#333;font-size:13px;">' + fechaEmiStr + '</td>',
    '      </tr>',
    '      <tr>',
    '        <td style="padding:8px 16px;color:#555;font-size:13px;font-weight:bold;">Concepto:</td>',
    '        <td style="padding:8px 16px;color:#333;font-size:13px;">' + _recEsc_(fac.concepto) + '</td>',
    '      </tr>',
    '      <tr style="background:#f5f5f5;">',
    '        <td style="padding:8px 16px;color:#333;font-size:14px;font-weight:bold;">IMPORTE TOTAL:</td>',
    '        <td style="padding:8px 16px;color:' + colorHeader + ';font-size:16px;font-weight:bold;">' + totalStr + '</td>',
    '      </tr>',
    pagoCells,
    '    </table>',

    '    <p style="color:#777;font-size:13px;line-height:1.5;">',
    '      Si tiene cualquier consulta, no dude en ponerse en contacto con nosotros',
    '      respondiendo a este correo.',
    '    </p>',
    '  </div>',

    '  <div style="background:#f0f0f0;padding:12px 24px;text-align:center;">',
    '    <p style="margin:0;color:#999;font-size:11px;">',
    _recEsc_(fac.empresa) + ' &middot; ' + _recEsc_(fac.emailEmpresa),
    '    </p>',
    '  </div>',
    '</div>'
  ].join('\n');

  GmailApp.sendEmail(fac.email, subject, '', {
    htmlBody:  htmlBody,
    name:      fac.empresa,
    replyTo:   fac.emailEmpresa
  });
}

// ----------------------------------------------------------------
// Log sheet
// ----------------------------------------------------------------

function _rec_logEnvio_(logSheet, fac, resultado) {
  var idRec = 'REC-' + Utilities.formatDate(new Date(), 'UTC', 'yyyyMMddHHmmss') + '-' + fac.id;
  logSheet.appendRow([
    idRec, new Date(),
    fac.id, fac.nombre, fac.email,
    fac.total, fac.fechaEmision, fac.dias,
    fac.tipo, fac.descTipo,
    resultado
  ]);
}

function _rec_getOrCreateLogSheet_(ss) {
  var nombre = 'RECORDATORIOS_LOG';
  var sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, 1, HEADERS_REC_LOG.length)
      .setValues([HEADERS_REC_LOG])
      .setBackground('#1a1a2e')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    var widths = [180, 130, 120, 160, 200, 80, 110, 90, 50, 200, 200];
    widths.forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });
    sh.setFrozenRows(1);
    logEvent('INFO', '_rec_getOrCreateLogSheet_', 'Hoja RECORDATORIOS_LOG creada', '');
  }
  return sh;
}

// ----------------------------------------------------------------
// Trigger
// ----------------------------------------------------------------

function instalarTriggerRecordatorios() {
  _rec_ensureConfig_();

  var ya = ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === 'enviarRecordatoriosPendientes';
  });

  if (!ya) {
    ScriptApp.newTrigger('enviarRecordatoriosPendientes')
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .create();
    logEvent('INFO', 'instalarTriggerRecordatorios', 'Trigger diario a las 9h instalado', '');
    SpreadsheetApp.getUi().alert(
      '✅ Trigger instalado.\n\nLos recordatorios se comprobarán automáticamente cada día a las 9h.'
    );
  } else {
    SpreadsheetApp.getUi().alert('ℹ️ El trigger ya estaba instalado.');
  }
}

function desinstalarTriggerRecordatorios() {
  var borrados = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'enviarRecordatoriosPendientes') {
      ScriptApp.deleteTrigger(t);
      borrados++;
    }
  });
  logEvent('INFO', 'desinstalarTriggerRecordatorios', 'Triggers eliminados: ' + borrados, '');
  SpreadsheetApp.getUi().alert(borrados > 0
    ? '✅ Trigger de recordatorios desinstalado.'
    : 'ℹ️ No se encontró ningún trigger de recordatorios activo.'
  );
}

// ----------------------------------------------------------------
// Menú — acciones secundarias
// ----------------------------------------------------------------

function verLogRecordatorios() {
  var sh = _rec_getOrCreateLogSheet_(getTargetSS_());
  sh.activate();
}

function configurarRecordatorios() {
  _rec_ensureConfig_();
  var d1  = getConfig('rec_dias_1')              || '30';
  var d2  = getConfig('rec_dias_2')              || '60';
  var d3  = getConfig('rec_dias_3')              || '90';
  var dRC = getConfig('rec_dias_reenvio_critico') || '30';
  var act = getConfig('rec_activo')              || '1';
  var dp  = getConfig('rec_datos_pago')          || '(no configurado)';

  var msg = [
    'Configuración actual de recordatorios:',
    '',
    '• Primer aviso:         ' + d1 + ' días desde emisión',
    '• Segundo aviso:        ' + d2 + ' días desde emisión',
    '• Aviso final:          ' + d3 + ' días desde emisión',
    '• Reintentar aviso final cada: ' + dRC + ' días',
    '• Sistema activo:       ' + (act !== '0' ? 'SÍ' : 'NO'),
    '• Datos de pago:        ' + dp,
    '',
    'Para cambiar estos valores, edita la hoja CONFIG.',
    'Claves:',
    '  rec_dias_1 | rec_dias_2 | rec_dias_3',
    '  rec_dias_reenvio_critico',
    '  rec_activo  (1 = activo, 0 = inactivo)',
    '  rec_datos_pago  (texto con IBAN / datos bancarios)'
  ].join('\n');

  SpreadsheetApp.getUi().alert('⏰ Recordatorios — Configuración', msg,
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// ----------------------------------------------------------------
// Utilidades internas
// ----------------------------------------------------------------

function _rec_ensureConfig_() {
  var defaults = [
    ['rec_dias_1',              '30'],
    ['rec_dias_2',              '60'],
    ['rec_dias_3',              '90'],
    ['rec_dias_reenvio_critico','30'],
    ['rec_activo',              '1'],
    ['rec_datos_pago',          '']
  ];
  defaults.forEach(function(pair) {
    if (getConfig(pair[0]) === null) setConfig(pair[0], pair[1]);
  });
}

function _recCfgInt_(key, def) {
  var v = parseInt(getConfig(key), 10);
  return isNaN(v) ? def : v;
}

/** Escapa caracteres HTML para evitar inyección en el cuerpo del email. */
function _recEsc_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Muestra alerta UI sin lanzar excepción cuando se llama desde un trigger. */
function _recAlerta_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (_) {}
}
