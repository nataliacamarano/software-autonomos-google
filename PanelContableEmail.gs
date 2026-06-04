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
 * Envía el email al gestor.
 * Muestra diálogo de confirmación con:
 *   - Nº PDFs de facturas incluidos / sin PDF
 *   - Nº adjuntos de gastos incluidos / sin adjunto
 *   - Documentos no localizados (errores)
 *   - URL de la carpeta Drive
 * Requiere confirmación explícita antes de enviar.
 */
function _pcnEnviarEmailGestor_() {
  var ss = getTargetSS_();
  var sh = ss.getSheetByName(PCN_SHEET);
  if (!sh) {
    SpreadsheetApp.getUi().alert('⚠ PANEL_CONTABLE no existe.');
    return;
  }

  var emailGestor  = String(
    sh.getRange(PCN_CFG.EMAIL_GESTOR[0],  PCN_CFG.EMAIL_GESTOR[1]).getValue() || ''
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
  var tipo      = _pcnLeerTipoPeriodo_(sh);
  var periodo   = _pcnPeriodoEtiqueta_(tipo);
  var empresa   = _pcnGetEmpresaNombre_();
  var kpis      = _pcnActualizarResumen_(ss, sh);
  var folderUrl = String(sh.getRange(PCN_ROWS.DRIVE_ROW, 11).getValue() || '').trim();
  if (folderUrl === '—') folderUrl = '';

  // Leer estadísticas del último paquete generado
  var stats = _pcnLeerStatsUltimoPaquete_();
  var statsPeriodoOk = stats && stats.periodo === periodo;

  // FASE 18.1: Bloquear si no se generó el paquete para este periodo
  if (!statsPeriodoOk) {
    SpreadsheetApp.getUi().alert(
      '⚠ No se puede enviar el email.\n\n' +
      'No se ha generado el paquete contable para el periodo:\n  ' + periodo + '\n\n' +
      'Ejecuta primero:  📦 Generar paquete contable\n' +
      '(Menú ⚙️ Sistema o botón en el panel)'
    );
    return;
  }

  // Bloquear si la carpeta está vacía (0 docs y 0 resúmenes)
  var totalCopiados = (stats.facPdfOk || 0) + (stats.gasAdjOk || 0);
  var totalSeleccionados = (stats.facSelected || stats.facTotal || 0) +
                           (stats.gasSelected || stats.gasTotal || 0);
  if (totalCopiados === 0 && totalSeleccionados === 0) {
    SpreadsheetApp.getUi().alert(
      '⚠ No se puede enviar el email.\n\n' +
      'El paquete contable está vacío:\n' +
      '  · 0 facturas seleccionadas\n' +
      '  · 0 gastos seleccionados\n\n' +
      'Actualiza el panel, revisa la selección de documentos\n' +
      'y vuelve a generar el paquete.'
    );
    return;
  }

  var asunto = '[' + empresa + '] Contabilidad ' + periodo + ' — Documentación preparada';

  // ── Diálogo de confirmación ───────────────────────────────────
  var ui = SpreadsheetApp.getUi();

  var lineas = [
    'Para:    ' + emailGestor + (nombreGestor ? ' (' + nombreGestor + ')' : ''),
    'Asunto:  ' + asunto,
    ''
  ];

  var noLoc = (stats.facPdfError || 0) + (stats.gasAdjError || 0);
  var facSel = stats.facSelected !== undefined ? stats.facSelected : stats.facTotal;
  var gasSel = stats.gasSelected !== undefined ? stats.gasSelected : stats.gasTotal;

  lineas.push('📄 Facturas:  ' + facSel + ' seleccionadas / ' +
    stats.facPdfOk + ' PDFs copiados' +
    (stats.facPdfSinUrl > 0 ? ' · ' + stats.facPdfSinUrl + ' sin PDF' : ''));
  lineas.push('🧾 Gastos:    ' + gasSel + ' seleccionados / ' +
    stats.gasAdjOk + ' adjuntos copiados' +
    (stats.gasSinAdj > 0 ? ' · ' + stats.gasSinAdj + ' sin adjunto' : ''));
  if (noLoc > 0) {
    lineas.push('⚠ No localizados: ' + noLoc + ' archivo(s)');
  }
  lineas.push('');
  lineas.push('IVA rep.: ' + _pcnFmtEur_(kpis.ivaRep) +
    '  Total fac: ' + _pcnFmtEur_(kpis.facTotal));
  lineas.push('IVA neto: ' + _pcnFmtEur_(kpis.ivaNeto) +
    '  Total gas: ' + _pcnFmtEur_(kpis.gasTotal));

  if (folderUrl) {
    lineas.push('');
    lineas.push('📁 Carpeta Drive adjunta al email.');
  }

  var respuesta = ui.alert(
    '✉️ Confirmar envío al gestor',
    lineas.join('\n') + '\n\n¿Enviar ahora?',
    ui.ButtonSet.YES_NO
  );

  if (respuesta !== ui.Button.YES) {
    logEvent('INFO', 'enviarEmailGestor', 'Envío cancelado.', periodo);
    return;
  }

  // ── Envío ─────────────────────────────────────────────────────
  try {
    var htmlBody  = _pcnComponerBodyHTML_(
      empresa, periodo, kpis, folderUrl, nombreGestor, stats, statsPeriodoOk
    );
    var plainBody = _pcnComponerBodyTexto_(
      empresa, periodo, kpis, folderUrl, nombreGestor, stats, statsPeriodoOk
    );

    GmailApp.sendEmail(emailGestor, asunto, plainBody, {
      htmlBody: htmlBody,
      name:     empresa
    });

    var logData = {
      fecha:      new Date(),
      periodo:    periodo,
      gestor:     nombreGestor,
      email:      emailGestor,
      estado:     '✅ Enviado',
      carpetaUrl: folderUrl,
      notas:      kpis.facCount + ' fac · ' + kpis.gasCount + ' gas · IVA ' +
                  _pcnFmtEur_(kpis.ivaNeto)
    };
    _pcnRegistrarEnvio_(ss, sh, logData);

    // FASE 18.3: marcar en historial de paquetes
    _pcnMarcarEnviado_(ss, periodo, emailGestor);

    SpreadsheetApp.flush();
    logEvent('INFO', 'enviarEmailGestor', 'Enviado a ' + emailGestor, periodo);
    ui.alert('✅ Email enviado correctamente a:\n' + emailGestor);

  } catch (e) {
    Logger.log('ERROR _pcnEnviarEmailGestor_: ' + e.message);
    logEvent('ERROR', 'enviarEmailGestor', e.message, emailGestor);
    ui.alert('❌ Error al enviar el email:\n\n' + e.message);
  }
}

// ── Composición HTML ──────────────────────────────────────────

function _pcnComponerBodyHTML_(empresa, periodo, kpis, folderUrl, gestor,
                                stats, statsPeriodoOk) {
  var saludoNombre = gestor || 'gestor/a';
  var ivaColor = kpis.ivaNeto >= 0 ? '#c5221f' : '#137333';

  // Bloque principal de carpeta Drive (siempre visible si hay URL)
  var folderBlock = folderUrl
    ? '<div style="background:#e8f0fe;border-left:4px solid #1a73e8;' +
      'padding:14px 16px;margin:20px 0;border-radius:0 4px 4px 0">' +
      '<p style="margin:0 0 6px;font-weight:bold;color:#1a73e8">' +
        '📁 Carpeta con la documentación del periodo</p>' +
      '<a href="' + folderUrl + '" style="color:#1a73e8;word-break:break-all">' +
        folderUrl + '</a>' +
      '<p style="margin:6px 0 0;font-size:12px;color:#5f6368">' +
        'Incluye PDFs de facturas, adjuntos de gastos, resumen IVA y CSV.</p>' +
      '</div>'
    : '';

  // Bloque de documentos incluidos (solo si stats disponibles)
  var docsBlock = '';
  if (statsPeriodoOk && stats) {
    var noLoc = (stats.facPdfError || 0) + (stats.gasAdjError || 0);
    docsBlock =
      '<table style="width:100%;border-collapse:collapse;font-size:13px;margin:12px 0">' +
      '<thead><tr style="background:#f1f3f4">' +
        '<th style="padding:7px 10px;text-align:left;border:1px solid #e0e0e0;font-size:12px">Tipo</th>' +
        '<th style="padding:7px 10px;text-align:center;border:1px solid #e0e0e0;font-size:12px">Incluidos</th>' +
        '<th style="padding:7px 10px;text-align:center;border:1px solid #e0e0e0;font-size:12px">Sin archivo</th>' +
      '</tr></thead><tbody>' +
      '<tr><td style="padding:7px 10px;border:1px solid #e0e0e0">📄 PDFs facturas</td>' +
        '<td style="padding:7px 10px;text-align:center;border:1px solid #e0e0e0;' +
          'font-weight:bold;color:#137333">' + stats.facPdfOk + ' / ' + stats.facTotal + '</td>' +
        '<td style="padding:7px 10px;text-align:center;border:1px solid #e0e0e0;color:#5f6368">' +
          stats.facPdfSinUrl + '</td></tr>' +
      '<tr style="background:#fafafa"><td style="padding:7px 10px;border:1px solid #e0e0e0">🧾 Adjuntos gastos</td>' +
        '<td style="padding:7px 10px;text-align:center;border:1px solid #e0e0e0;' +
          'font-weight:bold;color:#137333">' + stats.gasAdjOk + ' / ' + stats.gasTotal + '</td>' +
        '<td style="padding:7px 10px;text-align:center;border:1px solid #e0e0e0;color:#5f6368">' +
          stats.gasSinAdj + '</td></tr>' +
      (noLoc > 0
        ? '<tr><td colspan="3" style="padding:7px 10px;border:1px solid #e0e0e0;' +
            'color:#c5221f">⚠ ' + noLoc + ' archivo(s) no localizado(s) — ' +
            'verifique en la carpeta Drive.</td></tr>'
        : '') +
      '</tbody></table>';
  }

  function tr(label, value, bold) {
    var fw = bold ? 'font-weight:bold;' : '';
    return '<tr><td style="padding:7px 12px;border:1px solid #e0e0e0;' + fw + '">' +
      label + '</td><td style="padding:7px 12px;text-align:right;border:1px solid #e0e0e0;' +
      fw + '">' + value + '</td></tr>';
  }

  return '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;' +
    'color:#202124;margin:0;padding:0">' +
    '<div style="max-width:620px;margin:0 auto">' +

    // Cabecera
    '<div style="background:#4a148c;color:#fff;padding:24px 20px;' +
      'border-radius:8px 8px 0 0">' +
      '<h2 style="margin:0 0 4px">📦 Contabilidad ' + periodo + '</h2>' +
      '<p style="margin:0;opacity:0.8;font-size:14px">' + empresa + '</p>' +
    '</div>' +

    '<div style="border:1px solid #e0e0e0;border-top:none;padding:24px 20px;' +
      'border-radius:0 0 8px 8px;background:#fff">' +

      '<p>Estimado/a <b>' + saludoNombre + '</b>,</p>' +
      '<p>Le adjunto la documentación contable del periodo <b>' + periodo + '</b>, ' +
        'lista para su revisión y presentación.</p>' +

      // Enlace Drive prominente
      folderBlock +

      // Documentos incluidos
      (docsBlock
        ? '<p style="font-weight:bold;margin:16px 0 4px">Documentos incluidos:</p>' + docsBlock
        : '') +

      // Tabla de KPIs
      '<p style="font-weight:bold;margin:16px 0 4px">Resumen del periodo:</p>' +
      '<table style="width:100%;border-collapse:collapse;font-size:13px;margin:0 0 16px">' +
        '<thead><tr style="background:#f1f3f4">' +
          '<th style="padding:7px 12px;text-align:left;border:1px solid #e0e0e0;' +
            'font-size:12px">Concepto</th>' +
          '<th style="padding:7px 12px;text-align:right;border:1px solid #e0e0e0;' +
            'font-size:12px">Importe</th>' +
        '</tr></thead><tbody>' +
        tr('Facturas emitidas', kpis.facCount, false) +
        tr('Total facturado', _pcnFmtEur_(kpis.facTotal), false) +
        tr('IVA repercutido', _pcnFmtEur_(kpis.ivaRep), false) +
        tr('Pendiente de cobro', _pcnFmtEur_(kpis.pendiente), false) +
        '<tr><td colspan="2" style="border:1px solid #e0e0e0;padding:3px;background:#f8f9fa"></td></tr>' +
        tr('Gastos / tickets', kpis.gasCount, false) +
        tr('Total gastos', _pcnFmtEur_(kpis.gasTotal), false) +
        tr('IVA soportado', _pcnFmtEur_(kpis.ivaSop), false) +
        '<tr style="background:' + (kpis.ivaNeto >= 0 ? '#fce8e6' : '#e6f4ea') + '">' +
          '<td style="padding:10px 12px;font-weight:bold;border:1px solid #e0e0e0">' +
            'IVA a declarar</td>' +
          '<td style="padding:10px 12px;text-align:right;font-weight:bold;' +
            'color:' + ivaColor + ';border:1px solid #e0e0e0">' +
            _pcnFmtEur_(kpis.ivaNeto) + '</td></tr>' +
      '</tbody></table>' +

      '<p style="color:#5f6368;font-size:12px;margin-top:20px;border-top:1px solid #e0e0e0;' +
        'padding-top:12px">' +
        'Generado automáticamente por el sistema de gestión de ' + empresa + '. ' +
        'Para cualquier consulta, responda a este email.</p>' +

    '</div></div></body></html>';
}

// ── Composición texto plano ───────────────────────────────────

function _pcnComponerBodyTexto_(empresa, periodo, kpis, folderUrl, gestor,
                                 stats, statsPeriodoOk) {
  var lines = [
    'Estimado/a ' + (gestor || 'gestor/a') + ',',
    '',
    'Documentación contable — ' + periodo,
    ''
  ];

  if (folderUrl) {
    lines.push('📁 CARPETA CON LA DOCUMENTACIÓN:');
    lines.push(folderUrl);
    lines.push('');
  }

  if (statsPeriodoOk && stats) {
    lines.push('DOCUMENTOS INCLUIDOS:');
    lines.push('  PDFs facturas: ' + stats.facPdfOk + ' / ' + stats.facTotal +
      (stats.facPdfSinUrl > 0 ? ' (' + stats.facPdfSinUrl + ' sin PDF)' : ''));
    lines.push('  Fotos/gastos:  ' + stats.gasAdjOk + ' / ' + stats.gasTotal +
      (stats.gasSinAdj > 0 ? ' (' + stats.gasSinAdj + ' sin adjunto)' : ''));
    var noLoc = (stats.facPdfError || 0) + (stats.gasAdjError || 0);
    if (noLoc > 0) lines.push('  ⚠ No localizados: ' + noLoc);
    lines.push('');
  }

  lines.push('RESUMEN DEL PERIODO:');
  lines.push('  Facturas:      ' + kpis.facCount + '  · Total: ' + _pcnFmtEur_(kpis.facTotal));
  lines.push('  IVA rep.:      ' + _pcnFmtEur_(kpis.ivaRep));
  lines.push('  Pendiente:     ' + _pcnFmtEur_(kpis.pendiente));
  lines.push('  Gastos:        ' + kpis.gasCount + '  · Total: ' + _pcnFmtEur_(kpis.gasTotal));
  lines.push('  IVA sop.:      ' + _pcnFmtEur_(kpis.ivaSop));
  lines.push('  IVA A DECLARAR: ' + _pcnFmtEur_(kpis.ivaNeto));
  lines.push('');
  lines.push(empresa);
  return lines.join('\n');
}

// ============================================================
// FASE 17.7 — Registro de envíos
// ============================================================

function _pcnRegistrarEnvio_(ss, sh, logData) {
  var logSheet = _pcnEnsureLogSheet_(ss);

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

  _pcnActualizarHistorialPanel_(ss, sh, tz);
}

function _pcnActualizarHistorialPanel_(ss, sh, tz) {
  var logSheet = ss.getSheetByName(PCN_LOG_SHEET);
  if (!logSheet || logSheet.getLastRow() < 2) return;

  tz = tz || Session.getScriptTimeZone();

  var lastRow  = logSheet.getLastRow();
  var startLog = Math.max(2, lastRow - 4);
  var numRows  = lastRow - startLog + 1;
  var data     = logSheet.getRange(startLog, 1, numRows, 7).getValues().reverse();

  for (var i = 0; i < 5; i++) {
    var panelRow = PCN_ROWS.HIST_START + i;
    var bgh = i % 2 === 0 ? PCN_C.BLANCO : PCN_C.GRIS_BG;

    sh.getRange(panelRow, 1, 1, PCN_NC).breakApart();

    if (i < data.length) {
      var e    = data[i];
      var fStr = e[0] instanceof Date
        ? Utilities.formatDate(e[0], tz, 'dd/MM/yyyy HH:mm') : String(e[0] || '');
      sh.getRange(panelRow, 1, 1, PCN_NC).merge()
        .setValue(fStr + ' | ' + e[1] + ' | ' + e[3] + ' | ' + e[4])
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
