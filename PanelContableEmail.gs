// ============================================================
// PanelContableEmail.gs — FASE 17.6 + 17.7
// Email al gestor y registro de envíos
// ============================================================

var PCN_LOG_SHEET = 'ENVIOS_CONTABLE_LOG';

// ============================================================
// FASE 17.6 — Enviar email al gestor
// ============================================================

/** Wrapper público para menú. */
function enviarEmailGestor() {
  _pcnEnviarEmailGestor_();
}

/**
 * Envía un email HTML al gestor con el resumen del periodo.
 * Muestra diálogo de confirmación antes de enviar (regla obligatoria).
 */
function _pcnEnviarEmailGestor_() {
  var ss = getTargetSS_();
  var sh = ss.getSheetByName(PCN_SHEET);
  if (!sh) {
    SpreadsheetApp.getUi().alert('⚠ PANEL_CONTABLE no existe.');
    return;
  }

  // Leer configuración
  var emailGestor  = String(
    sh.getRange(PCN_CFG.EMAIL_GESTOR[0],  PCN_CFG.EMAIL_GESTOR[1] ).getValue() || ''
  ).trim();
  var nombreGestor = String(
    sh.getRange(PCN_CFG.NOMBRE_GESTOR[0], PCN_CFG.NOMBRE_GESTOR[1]).getValue() || ''
  ).trim();

  if (!emailGestor) {
    SpreadsheetApp.getUi().alert(
      '⚠ Email del gestor no configurado.\n\n' +
      'Introduce el email en la fila 7 del PANEL_CONTABLE e inténtalo de nuevo.'
    );
    return;
  }

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailGestor)) {
    SpreadsheetApp.getUi().alert(
      '⚠ Email del gestor no válido:\n  ' + emailGestor +
      '\n\nCorrígelo en la fila 7 e inténtalo de nuevo.'
    );
    return;
  }

  // Datos del periodo y KPIs
  var tipo       = _pcnLeerTipoPeriodo_(sh);
  var periodo    = _pcnPeriodoEtiqueta_(tipo);
  var empresa    = _pcnGetEmpresaNombre_();
  var folderUrl  = String(sh.getRange(PCN_ROWS.DRIVE_ROW, 11).getValue() || '').trim();
  if (folderUrl === '—') folderUrl = '';

  var kpis = _pcnActualizarResumen_(ss, sh);

  var asunto = '[' + empresa + '] Contabilidad ' + periodo + ' — Documentos preparados';

  // ── Confirmación ──────────────────────────────────────────────
  var ui = SpreadsheetApp.getUi();
  var previewLines = [
    'Para:    ' + emailGestor + (nombreGestor ? ' (' + nombreGestor + ')' : ''),
    'Asunto:  ' + asunto,
    '',
    'Periodo: ' + periodo,
    'Facturas: ' + kpis.facCount + '  |  Total: ' + _pcnFmtEur_(kpis.facTotal),
    'IVA rep.: ' + _pcnFmtEur_(kpis.ivaRep),
    'Gastos:   ' + kpis.gasCount + '  |  Total: ' + _pcnFmtEur_(kpis.gasTotal),
    'IVA neto: ' + _pcnFmtEur_(kpis.ivaNeto)
  ];
  if (folderUrl) {
    previewLines.push('');
    previewLines.push('Se incluye enlace a la carpeta Drive.');
  }

  var respuesta = ui.alert(
    '✉️ Confirmar envío al gestor',
    previewLines.join('\n') + '\n\n¿Enviar este email ahora?',
    ui.ButtonSet.YES_NO
  );

  if (respuesta !== ui.Button.YES) {
    logEvent('INFO', 'enviarEmailGestor', 'Envío cancelado por el usuario.', periodo);
    return;
  }

  // ── Envío ─────────────────────────────────────────────────────
  try {
    var htmlBody  = _pcnComponerBodyHTML_(empresa, periodo, kpis, folderUrl, nombreGestor);
    var plainBody = _pcnComponerBodyTexto_(empresa, periodo, kpis, folderUrl, nombreGestor);

    GmailApp.sendEmail(emailGestor, asunto, plainBody, {
      htmlBody: htmlBody,
      name:     empresa
    });

    // Registrar envío (FASE 17.7)
    var logData = {
      fecha:      new Date(),
      periodo:    periodo,
      gestor:     nombreGestor,
      email:      emailGestor,
      estado:     '✅ Enviado',
      carpetaUrl: folderUrl,
      notas:      kpis.facCount + ' fac · ' + kpis.gasCount + ' gas · IVA ' + _pcnFmtEur_(kpis.ivaNeto)
    };
    _pcnRegistrarEnvio_(ss, sh, logData);

    SpreadsheetApp.flush();
    logEvent('INFO', 'enviarEmailGestor', 'Enviado a ' + emailGestor, periodo);
    ui.alert('✅ Email enviado correctamente a:\n' + emailGestor);

  } catch (e) {
    Logger.log('ERROR _pcnEnviarEmailGestor_: ' + e.message);
    logEvent('ERROR', 'enviarEmailGestor', e.message, emailGestor);
    ui.alert('❌ Error al enviar el email:\n\n' + e.message);
  }
}

// ── Composición del email ─────────────────────────────────────

function _pcnComponerBodyHTML_(empresa, periodo, kpis, folderUrl, gestor) {
  var saludoNombre = gestor || 'gestor/a';
  var ivaColor = kpis.ivaNeto >= 0 ? '#c5221f' : '#137333';
  var folderBlock = folderUrl
    ? '<p style="margin:16px 0">📁 <a href="' + folderUrl + '" style="color:#1a73e8">' +
      'Acceder a la carpeta de documentos en Google Drive</a></p>'
    : '';

  function tr(label, value, bold) {
    var fw = bold ? 'font-weight:bold;' : '';
    return '<tr><td style="padding:8px 12px;border:1px solid #e0e0e0;' + fw + '">' +
      label + '</td><td style="padding:8px 12px;text-align:right;border:1px solid #e0e0e0;' +
      fw + '">' + value + '</td></tr>';
  }

  return '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;' +
    'color:#202124;margin:0;padding:0">' +
    '<div style="max-width:600px;margin:0 auto">' +
    '<div style="background:#4a148c;color:#fff;padding:24px 20px;' +
      'border-radius:8px 8px 0 0">' +
      '<h2 style="margin:0 0 4px">📦 Contabilidad ' + periodo + '</h2>' +
      '<p style="margin:0;opacity:0.8;font-size:14px">' + empresa + '</p>' +
    '</div>' +
    '<div style="border:1px solid #e0e0e0;border-top:none;padding:24px 20px;' +
      'border-radius:0 0 8px 8px;background:#fff">' +
      '<p>Estimado/a <b>' + saludoNombre + '</b>,</p>' +
      '<p>Le adjunto la documentación contable correspondiente al periodo ' +
        '<b>' + periodo + '</b>:</p>' +
      '<table style="width:100%;border-collapse:collapse;font-size:14px;margin:16px 0">' +
        '<thead><tr style="background:#f1f3f4">' +
          '<th style="padding:8px 12px;text-align:left;border:1px solid #e0e0e0;' +
            'font-size:12px">Concepto</th>' +
          '<th style="padding:8px 12px;text-align:right;border:1px solid #e0e0e0;' +
            'font-size:12px">Valor</th>' +
        '</tr></thead><tbody>' +
        tr('Facturas emitidas', kpis.facCount, false) +
        tr('Total facturado', _pcnFmtEur_(kpis.facTotal), false) +
        tr('IVA repercutido', _pcnFmtEur_(kpis.ivaRep), false) +
        tr('Pendiente de cobro', _pcnFmtEur_(kpis.pendiente), false) +
        '<tr><td colspan="2" style="border:1px solid #e0e0e0;' +
          'padding:2px;background:#f8f9fa"></td></tr>' +
        tr('Gastos / tickets', kpis.gasCount, false) +
        tr('Total gastos', _pcnFmtEur_(kpis.gasTotal), false) +
        tr('IVA soportado', _pcnFmtEur_(kpis.ivaSop), false) +
        '<tr style="background:#fce8e6">' +
          '<td style="padding:10px 12px;font-weight:bold;border:1px solid #e0e0e0">' +
            'IVA a declarar</td>' +
          '<td style="padding:10px 12px;text-align:right;font-weight:bold;' +
            'color:' + ivaColor + ';border:1px solid #e0e0e0">' +
            _pcnFmtEur_(kpis.ivaNeto) + '</td></tr>' +
      '</tbody></table>' +
      folderBlock +
      '<p style="color:#5f6368;font-size:12px;margin-top:24px">' +
        'Este email ha sido generado automáticamente por el sistema de gestión de ' +
        empresa + '.</p>' +
    '</div></div></body></html>';
}

function _pcnComponerBodyTexto_(empresa, periodo, kpis, folderUrl, gestor) {
  var lines = [
    'Estimado/a ' + (gestor || 'gestor/a') + ',',
    '',
    'Documentación contable — ' + periodo,
    '',
    'FACTURAS:',
    '  Nº facturas:      ' + kpis.facCount,
    '  Total facturado:  ' + _pcnFmtEur_(kpis.facTotal),
    '  IVA repercutido:  ' + _pcnFmtEur_(kpis.ivaRep),
    '  Pendiente cobro:  ' + _pcnFmtEur_(kpis.pendiente),
    '',
    'GASTOS:',
    '  Nº gastos:        ' + kpis.gasCount,
    '  Total gastos:     ' + _pcnFmtEur_(kpis.gasTotal),
    '  IVA soportado:    ' + _pcnFmtEur_(kpis.ivaSop),
    '',
    'IVA A DECLARAR: ' + _pcnFmtEur_(kpis.ivaNeto)
  ];
  if (folderUrl) {
    lines.push('');
    lines.push('Carpeta de documentos: ' + folderUrl);
  }
  lines.push('');
  lines.push(empresa);
  return lines.join('\n');
}

// ============================================================
// FASE 17.7 — Registro de envíos
// ============================================================

/**
 * Escribe la entrada en ENVIOS_CONTABLE_LOG y actualiza
 * la sección Email y el historial del PANEL_CONTABLE.
 */
function _pcnRegistrarEnvio_(ss, sh, logData) {
  var logSheet = _pcnEnsureLogSheet_(ss);

  // Nueva fila en el log
  logSheet.appendRow([
    logData.fecha,
    logData.periodo,
    logData.gestor,
    logData.email,
    logData.estado,
    logData.carpetaUrl,
    logData.notas
  ]);
  logSheet.getRange(logSheet.getLastRow(), 1).setNumberFormat('dd/MM/yyyy HH:mm');

  // Actualizar sección "EMAIL AL GESTOR" en el panel (fila EMAIL_ROW)
  var tz       = Session.getScriptTimeZone();
  var fmtFecha = Utilities.formatDate(logData.fecha, tz, 'dd/MM/yyyy HH:mm');
  sh.getRange(PCN_ROWS.EMAIL_ROW, 4, 1, 4).merge()
    .setValue(fmtFecha)
    .setBackground(PCN_C.VERDE_BG).setFontColor(PCN_C.VERDE_DARK)
    .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(PCN_ROWS.EMAIL_ROW, 11, 1, 4).merge()
    .setValue(logData.estado)
    .setBackground(PCN_C.VERDE_BG).setFontColor(PCN_C.VERDE_DARK)
    .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');

  // Actualizar historial visual (filas 30-34)
  _pcnActualizarHistorialPanel_(ss, sh, tz);
}

/**
 * Rellena las 5 filas del historial (HIST_START…HIST_START+4)
 * con las últimas entradas del log, más reciente primero.
 */
function _pcnActualizarHistorialPanel_(ss, sh, tz) {
  var logSheet = ss.getSheetByName(PCN_LOG_SHEET);
  if (!logSheet || logSheet.getLastRow() < 2) return;

  tz = tz || Session.getScriptTimeZone();

  var lastRow  = logSheet.getLastRow();
  var startLog = Math.max(2, lastRow - 4);
  var numRows  = lastRow - startLog + 1;
  var data     = logSheet.getRange(startLog, 1, numRows, 7).getValues().reverse();

  var R = PCN_ROWS;

  for (var i = 0; i < 5; i++) {
    var panelRow = R.HIST_START + i;
    var bgh = i % 2 === 0 ? PCN_C.BLANCO : PCN_C.GRIS_BG;

    sh.getRange(panelRow, 1, 1, PCN_NC).breakApart();

    if (i < data.length) {
      var e       = data[i];
      var fStr    = e[0] instanceof Date
        ? Utilities.formatDate(e[0], tz, 'dd/MM/yyyy HH:mm')
        : String(e[0] || '');
      var lineText = fStr + ' | ' + e[1] + ' | ' + e[3] + ' | ' + e[4];
      sh.getRange(panelRow, 1, 1, PCN_NC).merge()
        .setValue(lineText)
        .setBackground(bgh).setFontColor(PCN_C.GRIS_DARK)
        .setFontSize(9).setHorizontalAlignment('left').setVerticalAlignment('middle');
    } else {
      sh.getRange(panelRow, 1, 1, PCN_NC).merge()
        .setValue('—')
        .setBackground(bgh).setFontColor(PCN_C.GRIS)
        .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
    }
    sh.setRowHeight(panelRow, 24);
  }
}

/**
 * Crea ENVIOS_CONTABLE_LOG si no existe.
 * Devuelve la hoja de log.
 */
function _pcnEnsureLogSheet_(ss) {
  var sh = ss.getSheetByName(PCN_LOG_SHEET);
  if (sh) return sh;

  sh = ss.insertSheet(PCN_LOG_SHEET);
  var headers = ['Fecha', 'Periodo', 'Gestor', 'Email', 'Estado', 'URL Carpeta', 'Notas'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#4a148c').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  [80, 100, 140, 200, 120, 300, 250].forEach(function(w, i) {
    sh.setColumnWidth(i + 1, w);
  });
  sh.getRange(1, 1).setNumberFormat('dd/MM/yyyy HH:mm');

  logEvent('INFO', 'ensureLogSheet', PCN_LOG_SHEET + ' creado.', '');
  return sh;
}
