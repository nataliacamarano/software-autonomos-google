// ============================================================
// LogoEmpresa.gs — Logo corporativo en facturas/presupuestos/proformas
// ============================================================
//
// _cabecera_empresa_(body, nombre, cif, tel, email, dir)
//   Reemplaza el bloque "Cabecera empresa" + appendHorizontalRule()
//   en generateInvoicePDF, generatePresupuestoPDF y
//   generateInvoiceFromPresupuestoPDF.
//
// Diseño: logo (140pt) izquierda · info empresa derecha · sin bordes.
// Fallback texto si el logo no puede cargarse.

var EL_ = {
  CONFIG_KEY: 'empresa_logo_id',
  LOGO_ID:    '1Qbwi6neyzdsF85Ohp2cKW1EuUlaxccyc',
  LOGO_W:     220   // ancho en puntos (≈ 220 px a 72 dpi)
};

// ────────────────────────────────────────────────────────────────
// HELPER PRINCIPAL — llamado desde los 3 generadores de PDF
// ────────────────────────────────────────────────────────────────

function _cabecera_empresa_(body, nombre, cif, tel, email, dir) {
  var logoId = getConfig(EL_.CONFIG_KEY) || EL_.LOGO_ID;
  var logoOk = false;

  if (logoId) {
    try {
      var logoBlob = DriveApp.getFileById(logoId).getBlob();

      // Tabla 1 fila × 2 columnas: logo | info empresa
      var hdrTbl = body.appendTable([['', '']]);

      // Bordes completamente invisibles: ancho 0 + color blanco en tabla y celdas
      var tblAttrs = {};
      tblAttrs[DocumentApp.Attribute.BORDER_WIDTH] = 0;
      tblAttrs[DocumentApp.Attribute.BORDER_COLOR] = '#ffffff';
      hdrTbl.setAttributes(tblAttrs);

      var cellAttrs = {};
      cellAttrs[DocumentApp.Attribute.BORDER_WIDTH] = 0;
      cellAttrs[DocumentApp.Attribute.BORDER_COLOR] = '#ffffff';

      var leftCell  = hdrTbl.getRow(0).getCell(0);
      var rightCell = hdrTbl.getRow(0).getCell(1);
      leftCell.setAttributes(cellAttrs);
      rightCell.setAttributes(cellAttrs);

      // ── Celda izquierda: logo ────────────────────────────────
      var leftPara = leftCell.getChild(0).asParagraph();
      leftPara.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      var img = leftPara.appendInlineImage(logoBlob);
      var ow  = img.getWidth();
      var oh  = img.getHeight();
      if (ow > 0) img.setWidth(EL_.LOGO_W).setHeight(Math.round(oh * EL_.LOGO_W / ow));

      // ── Celda derecha: info empresa ──────────────────────────
      rightCell.setVerticalAlignment(DocumentApp.VerticalAlignment.MIDDLE);
      var rp0 = rightCell.getChild(0).asParagraph();
      rp0.editAsText().setText(nombre).setFontSize(13).setBold(true);
      rp0.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

      var rLines = [];
      if (cif)   rLines.push('NIF/CIF: ' + cif);
      if (dir)   rLines.push(dir);
      if (email) rLines.push(email);
      if (tel)   rLines.push(tel);
      rLines.forEach(function(line) {
        rightCell.appendParagraph(line)
          .editAsText().setFontSize(9).setForegroundColor('#555555').setBold(false);
        rightCell.getChild(rightCell.getNumChildren() - 1).asParagraph()
          .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      });

      logoOk = true;
      Logger.log('[LOGO] Insertado OK — ' + logoId);

    } catch (e) {
      Logger.log('[LOGO] Error: ' + e.message);
      try { logEvent('WARN', '_cabecera_empresa_', 'Logo no cargado: ' + e.message, logoId); } catch (_) {}
    }
  }

  // Fallback: cabecera solo texto (igual que antes)
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

  body.appendHorizontalRule();
  return logoOk;
}

// ────────────────────────────────────────────────────────────────
// CONFIGURAR — escribe el ID en CONFIG
// ────────────────────────────────────────────────────────────────

function configurarLogoEmpresa() {
  setConfig(EL_.CONFIG_KEY, EL_.LOGO_ID);
  SpreadsheetApp.getUi().alert(
    '✅ Logo configurado',
    'ID guardado en CONFIG:\n' + EL_.LOGO_ID + '\n\n' +
    'Se insertará automáticamente en:\n' +
    '  · Facturas nuevas\n' +
    '  · Presupuestos nuevos\n' +
    '  · Proformas nuevas',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ────────────────────────────────────────────────────────────────
// DIAGNÓSTICO — crea hoja DIAGNOSTICO_LOGO_EMPRESA
// ────────────────────────────────────────────────────────────────

function diagnosticarLogoEmpresa() {
  var ss  = SpreadsheetApp.getActive();
  var ui  = SpreadsheetApp.getUi();
  var tz  = Session.getScriptTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');

  // Asegurar que el ID está en CONFIG
  var logoId = getConfig(EL_.CONFIG_KEY) || '';
  if (!logoId) {
    setConfig(EL_.CONFIG_KEY, EL_.LOGO_ID);
    logoId = EL_.LOGO_ID;
  }

  var filas = [
    ['CONCEPTO', 'ESTADO', 'DETALLE'],
    ['Fecha diagnóstico', now, ''],
    ['', '', ''],
    ['LOGO', '', ''],
    ['empresa_logo_id en CONFIG', logoId ? '✅ Configurado' : '❌ Vacío', logoId],
  ];

  // Verificar acceso al logo en Drive
  var logoAcceso = '❌ No accesible';
  var logoDim    = '';
  try {
    var logoFile = DriveApp.getFileById(logoId);
    logoAcceso = '✅ Encontrado en Drive';
    var b      = logoFile.getBlob();
    var tmpDoc = DocumentApp.create('__logo_check_tmp__');
    var tmpImg = tmpDoc.getBody().appendImage(b);
    logoDim = tmpImg.getWidth() + ' × ' + tmpImg.getHeight() + ' pt';
    DriveApp.getFileById(tmpDoc.getId()).setTrashed(true);
  } catch (e) {
    logoAcceso = '❌ ' + e.message;
  }

  filas.push(['Logo accesible', logoAcceso, logoDim || '']);
  filas.push(['', '', '']);

  // Estado por tipo de documento
  filas.push(['DOCUMENTOS PDF', '', '']);
  filas.push(['PDF Facturas',      '✅ Integrado', 'generateInvoicePDF — Facturas.gs']);
  filas.push(['PDF Presupuestos',  '✅ Integrado', 'generatePresupuestoPDF — Presupuestos.gs']);
  filas.push(['PDF Proformas',     '✅ Integrado', 'generateInvoiceFromPresupuestoPDF — FacturasPresupuesto.gs']);
  filas.push(['', '', '']);

  // Config empresa
  filas.push(['CONFIG EMPRESA', '', '']);
  var empKeys = [
    ['empresa_nombre',    'Nombre / razón social'],
    ['empresa_cif',       'NIF / CIF'],
    ['empresa_direccion', 'Dirección'],
    ['empresa_email',     'Email'],
    ['empresa_telefono',  'Teléfono']
  ];
  empKeys.forEach(function(kv) {
    var v = getConfig(kv[0]) || '';
    filas.push([kv[1], v ? '✅ ' + v : '⚠️ Vacío', kv[0]]);
  });

  // Crear / limpiar hoja
  var HOJA = 'DIAGNOSTICO_LOGO_EMPRESA';
  var sh   = ss.getSheetByName(HOJA);
  if (sh)  sh.clearContents(); else sh = ss.insertSheet(HOJA);

  sh.getRange(1, 1, filas.length, 3).setValues(filas);

  // Formato
  sh.getRange(1, 1, 1, 3)
    .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  [4, 8, 12].forEach(function(r) {
    if (r <= filas.length)
      sh.getRange(r, 1, 1, 3).setBackground('#e8f0fe').setFontWeight('bold');
  });
  sh.setColumnWidth(1, 220).setColumnWidth(2, 240).setColumnWidth(3, 300);
  sh.getRange(1, 1, filas.length, 3).setVerticalAlignment('top').setWrap(true);

  ss.setActiveSheet(sh);
  logEvent('INFO', 'diagnosticarLogoEmpresa', logoAcceso, logoId);

  ui.alert('🖼️ Diagnóstico Logo',
    'Logo: ' + logoAcceso + '\n' +
    (logoDim ? 'Dimensiones: ' + logoDim + '\n' : '') +
    '\nIntegrado en: Facturas · Presupuestos · Proformas' +
    '\n\nVer hoja DIAGNOSTICO_LOGO_EMPRESA',
    ui.ButtonSet.OK);
}
