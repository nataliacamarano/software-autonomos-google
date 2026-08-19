// ============================================================
// LogoEmpresa.gs — Logo corporativo en facturas/presupuestos/proformas
// ============================================================
//
// _cabecera_empresa_(body, nombre, cif, tel, email, dir, docId)
//   Tabla 1×2 sin bordes (eliminados vía Docs REST API):
//   · Columna izquierda: logo empresa (~240pt)
//   · Columna derecha:   info empresa alineada a la derecha
//   · Línea horizontal gris fina debajo
//   · Fallback texto si el logo falla (nunca duplica datos)

var EL_ = {
  CONFIG_KEY: 'empresa_logo_id',
  LOGO_ID:    '1Qbwi6neyzdsF85Ohp2cKW1EuUlaxccyc',
  LOGO_W:     240   // ≈ 85 mm a 72 dpi
};

// ────────────────────────────────────────────────────────────────
// HELPER PRINCIPAL — llamado desde los 3 generadores de PDF
// ────────────────────────────────────────────────────────────────

function _cabecera_empresa_(body, nombre, cif, tel, email, dir, docId) {
  var logoId = getConfig(EL_.CONFIG_KEY) || EL_.LOGO_ID;
  var logoOk = false;

  // Reducir espacio superior (el doc nuevo tiene 1 párrafo vacío inicial)
  var firstChild = body.getChild(0);
  if (firstChild && firstChild.getType() === DocumentApp.ElementType.PARAGRAPH) {
    var topAttrs = {};
    topAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 0;
    topAttrs[DocumentApp.Attribute.SPACING_AFTER]  = 0;
    firstChild.setAttributes(topAttrs);
  }

  if (logoId) {
    try {
      var logoBlob = DriveApp.getFileById(logoId).getBlob();

      // ── Tabla 1×2: logo | info empresa ───────────────────────
      var hdrTbl = body.appendTable([['', '']]);

      // Bordes blancos como fallback visual (serán eliminados vía API)
      var wb = {};
      wb[DocumentApp.Attribute.BORDER_COLOR] = '#ffffff';
      wb[DocumentApp.Attribute.BORDER_WIDTH] = 0;
      hdrTbl.setAttributes(wb);
      var leftCell  = hdrTbl.getRow(0).getCell(0);
      var rightCell = hdrTbl.getRow(0).getCell(1);
      leftCell.setAttributes(wb);
      rightCell.setAttributes(wb);

      // ── Celda izquierda: logo ────────────────────────────────
      leftCell.setVerticalAlignment(DocumentApp.VerticalAlignment.MIDDLE);
      var leftPara = leftCell.getChild(0).asParagraph();
      leftPara.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      var img = leftPara.appendInlineImage(logoBlob);
      var ow  = img.getWidth();
      var oh  = img.getHeight();
      if (ow > 0) img.setWidth(EL_.LOGO_W).setHeight(Math.round(oh * EL_.LOGO_W / ow));

      // ── Celda derecha: info empresa ──────────────────────────
      rightCell.setVerticalAlignment(DocumentApp.VerticalAlignment.MIDDLE);

      var rp0 = rightCell.getChild(0).asParagraph();
      rp0.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      rp0.editAsText().setText(nombre).setFontSize(12).setBold(true).setForegroundColor('#1a1a1a');
      var cmpAttrs = {};
      cmpAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 0;
      cmpAttrs[DocumentApp.Attribute.SPACING_AFTER]  = 0;
      rp0.setAttributes(cmpAttrs);

      var rLines = [];
      if (cif)   rLines.push('NIF/CIF: ' + cif);
      if (dir)   rLines.push(dir);
      if (email) rLines.push(email);
      if (tel)   rLines.push(tel);

      rLines.forEach(function(line) {
        var rp = rightCell.appendParagraph(line);
        rp.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
        rp.editAsText().setFontSize(9).setBold(false).setForegroundColor('#555555');
        rp.setAttributes(cmpAttrs);
      });

      // ── Eliminar bordes vía Docs REST API ────────────────────
      if (docId) _el_removeTableBorders_(docId);

      logoOk = true;
      Logger.log('[LOGO] OK — ' + EL_.LOGO_W + 'pt');

    } catch (e) {
      Logger.log('[LOGO] Error: ' + e.message);
      try { logEvent('WARN', '_cabecera_empresa_', 'Logo no cargado: ' + e.message, logoId); } catch (_) {}
    }
  }

  // ── Fallback: solo texto (nunca duplica) ─────────────────────
  if (!logoOk) {
    body.appendParagraph(nombre)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .editAsText().setFontSize(16).setBold(true);
    var infoEmp = [];
    if (cif)   infoEmp.push('NIF/CIF: ' + cif);
    if (tel)   infoEmp.push('Tel: ' + tel);
    if (email) infoEmp.push(email);
    if (infoEmp.length > 0)
      body.appendParagraph(infoEmp.join('   ·   ')).editAsText().setFontSize(9);
    if (dir)
      body.appendParagraph(dir).editAsText().setFontSize(9);
  }

  // ── Línea separadora fina ─────────────────────────────────────
  body.appendHorizontalRule();
  return logoOk;
}

// ────────────────────────────────────────────────────────────────
// DOCS REST API — elimina bordes de la primera tabla del documento
// ────────────────────────────────────────────────────────────────

function _el_removeTableBorders_(docId) {
  try {
    var token = ScriptApp.getOAuthToken();

    // Leer JSON del documento para localizar la primera tabla
    var getResp = UrlFetchApp.fetch(
      'https://docs.googleapis.com/v1/documents/' + docId,
      { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
    );
    if (getResp.getResponseCode() !== 200) {
      Logger.log('[LOGO] Docs API GET error: ' + getResp.getResponseCode());
      return;
    }

    var docJson      = JSON.parse(getResp.getContentText());
    var bodyContent  = docJson.body.content;
    var tableStartIdx = null;
    var tblRows = 1, tblCols = 2;

    for (var i = 0; i < bodyContent.length; i++) {
      if (bodyContent[i].table) {
        tableStartIdx = bodyContent[i].startIndex;
        tblRows       = bodyContent[i].table.rows;
        tblCols       = bodyContent[i].table.columns;
        break;
      }
    }

    if (tableStartIdx === null) {
      Logger.log('[LOGO] No table found in document JSON');
      return;
    }

    // Border a 0pt color blanco
    var noBorder = {
      color:     { color: { rgbColor: { red: 1, green: 1, blue: 1 } } },
      dashStyle: 'SOLID',
      width:     { magnitude: 0, unit: 'PT' }
    };

    var requests = [];

    for (var r = 0; r < tblRows; r++) {
      for (var c = 0; c < tblCols; c++) {
        requests.push({
          updateTableCellStyle: {
            tableRange: {
              tableCellLocation: {
                tableStartLocation: { index: tableStartIdx },
                rowIndex:    r,
                columnIndex: c
              },
              rowSpan: 1, columnSpan: 1
            },
            tableCellStyle: {
              borderTop:    noBorder,
              borderBottom: noBorder,
              borderLeft:   noBorder,
              borderRight:  noBorder,
              contentAlignment: 'MIDDLE'
            },
            fields: 'borderTop,borderBottom,borderLeft,borderRight,contentAlignment'
          }
        });
      }
    }

    var patchResp = UrlFetchApp.fetch(
      'https://docs.googleapis.com/v1/documents/' + docId + ':batchUpdate',
      {
        method:           'POST',
        headers:          { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload:          JSON.stringify({ requests: requests }),
        muteHttpExceptions: true
      }
    );

    if (patchResp.getResponseCode() === 200) {
      Logger.log('[LOGO] Borders removed — table index ' + tableStartIdx);
    } else {
      Logger.log('[LOGO] batchUpdate error: ' + patchResp.getContentText());
    }

  } catch (e) {
    Logger.log('[LOGO] _el_removeTableBorders_ error: ' + e.message);
  }
}

// ────────────────────────────────────────────────────────────────
// CONFIGURAR / DIAGNÓSTICO
// ────────────────────────────────────────────────────────────────

function configurarLogoEmpresa() {
  setConfig(EL_.CONFIG_KEY, EL_.LOGO_ID);
  SpreadsheetApp.getUi().alert('✅ Logo configurado en CONFIG.\nID: ' + EL_.LOGO_ID, SpreadsheetApp.getUi().ButtonSet.OK);
}

function diagnosticarLogoEmpresa() {
  var ss  = SpreadsheetApp.getActive();
  var ui  = SpreadsheetApp.getUi();
  var tz  = Session.getScriptTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');

  var logoId = getConfig(EL_.CONFIG_KEY) || '';
  if (!logoId) { setConfig(EL_.CONFIG_KEY, EL_.LOGO_ID); logoId = EL_.LOGO_ID; }

  var filas = [
    ['CONCEPTO', 'ESTADO', 'DETALLE'],
    ['Fecha', now, ''], ['', '', ''],
    ['LOGO', '', ''],
    ['empresa_logo_id', logoId ? '✅' : '❌', logoId],
  ];

  var logoAcceso = '❌', logoDim = '';
  try {
    var b  = DriveApp.getFileById(logoId).getBlob();
    var td = DocumentApp.create('__logo_chk__');
    var ti = td.getBody().appendImage(b);
    logoDim = ti.getWidth() + ' × ' + ti.getHeight() + ' pt';
    DriveApp.getFileById(td.getId()).setTrashed(true);
    logoAcceso = '✅ Accesible';
  } catch (e) { logoAcceso = '❌ ' + e.message; }

  filas.push(['Logo accesible', logoAcceso, logoDim]);
  filas.push(['Técnica', 'Tabla 1×2, bordes 0pt vía Docs REST API', '']);
  filas.push(['', '', '']);
  filas.push(['DOCUMENTOS PDF', '', '']);
  filas.push(['Facturas',     '✅', 'generateInvoicePDF']);
  filas.push(['Presupuestos', '✅', 'generatePresupuestoPDF']);
  filas.push(['Proformas',    '✅', 'generateInvoiceFromPresupuestoPDF']);
  filas.push(['', '', '']);
  filas.push(['CONFIG EMPRESA', '', '']);
  [['empresa_nombre','Nombre'],['empresa_cif','NIF/CIF'],
   ['empresa_direccion','Dirección'],['empresa_email','Email'],
   ['empresa_telefono','Teléfono']].forEach(function(kv) {
    var v = getConfig(kv[0]) || '';
    filas.push([kv[1], v ? '✅ ' + v : '⚠️ Vacío', kv[0]]);
  });

  var HOJA = 'DIAGNOSTICO_LOGO_EMPRESA';
  var sh   = ss.getSheetByName(HOJA);
  if (sh)  sh.clearContents(); else sh = ss.insertSheet(HOJA);
  sh.getRange(1, 1, filas.length, 3).setValues(filas);
  sh.getRange(1, 1, 1, 3).setBackground('#1a73e8').setFontColor('#fff').setFontWeight('bold');
  [4, 9, 13].forEach(function(r) {
    if (r <= filas.length) sh.getRange(r, 1, 1, 3).setBackground('#e8f0fe').setFontWeight('bold');
  });
  sh.setColumnWidth(1, 200).setColumnWidth(2, 260).setColumnWidth(3, 300);
  sh.getRange(1, 1, filas.length, 3).setVerticalAlignment('top').setWrap(true);
  ss.setActiveSheet(sh);

  ui.alert('🖼️ Diagnóstico Logo', 'Logo: ' + logoAcceso + (logoDim ? '\nDim: ' + logoDim : ''), ui.ButtonSet.OK);
}
