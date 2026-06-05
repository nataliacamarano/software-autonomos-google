// ============================================================
// Informes.gs — Generación de informes PDF por cliente y periodo
// ============================================================

/**
 * Lee de CONFIG las claves:
 *   cliente_reporte        → nombre exacto del cliente (tal como aparece en TRABAJOS)
 *   fecha_inicio_reporte   → DD/MM/YYYY
 *   fecha_fin_reporte      → DD/MM/YYYY
 *
 * Filtra TRABAJOS por cliente y rango de fechas, calcula totales,
 * genera un Google Doc, lo convierte a PDF, lo guarda en la carpeta
 * Drive de facturas/reportes y deja registro en LOGS.
 */
function generateClientReportPDF() {
  Logger.log('================================================');
  Logger.log('generateClientReportPDF()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  // ── 1. Leer parámetros de CONFIG ─────────────────────────────
  var cliente     = getConfig('cliente_reporte');
  var fechaIniStr = getConfig('fecha_inicio_reporte');
  var fechaFinStr = getConfig('fecha_fin_reporte');

  Logger.log('\n[1] Parámetros de CONFIG:');
  Logger.log('  cliente_reporte:      ' + (cliente     || '❌ NO CONFIGURADO'));
  Logger.log('  fecha_inicio_reporte: ' + (fechaIniStr || '❌ NO CONFIGURADO'));
  Logger.log('  fecha_fin_reporte:    ' + (fechaFinStr || '❌ NO CONFIGURADO'));

  if (!cliente || !fechaIniStr || !fechaFinStr) {
    Logger.log('\nABORTANDO. Rellena las 3 claves en CONFIG y vuelve a ejecutar.');
    return;
  }

  var fechaIni = _parseReportDate_(fechaIniStr);
  var fechaFin = _parseReportDate_(fechaFinStr);
  if (!fechaIni || !fechaFin) {
    Logger.log('ERROR: Formato de fecha incorrecto. Usa DD/MM/YYYY.');
    return;
  }
  fechaFin.setHours(23, 59, 59, 999); // incluir todo el día final

  // ── 2. Filtrar TRABAJOS ───────────────────────────────────────
  Logger.log('\n[2] Filtrando TRABAJOS...');
  var sheet   = getSheet('TRABAJOS');
  var lastRow = sheet.getLastRow();
  var trabajos = [];

  if (lastRow > 1) {
    var numCols = Math.min(18, sheet.getLastColumn());
    var data    = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    var seenIds = {};
    for (var i = 0; i < data.length; i++) {
      var row = data[i];

      var clienteFila = String(row[COL_TRA.CLIENTE - 1] || '').trim();
      if (clienteFila.toLowerCase() !== cliente.toLowerCase()) continue;

      var traId = String(row[COL_TRA.ID - 1] || '').trim();
      if (traId && seenIds[traId]) {
        Logger.log('  ⚠️ ID duplicado saltado: ' + traId);
        continue;
      }
      if (traId) seenIds[traId] = true;

      var fechaTrabajo = _parseTrabajosDate_(row[COL_TRA.FECHA_TRABAJO - 1]);
      if (!fechaTrabajo) continue;
      if (fechaTrabajo < fechaIni || fechaTrabajo > fechaFin) continue;

      trabajos.push({
        id:               String(row[COL_TRA.ID               - 1] || ''),
        fechaTrabajo:     fechaTrabajo,
        operario:         String(row[COL_TRA.OPERARIO          - 1] || '').trim(),
        horaInicio:       row[COL_TRA.HORA_INICIO       - 1],
        horaFin:          row[COL_TRA.HORA_FIN          - 1],
        trabajoRealizado: String(row[COL_TRA.TRABAJO_REALIZADO - 1] || '').trim(),
        incidencias:      String(row[COL_TRA.INCIDENCIAS       - 1] || '').trim(),
        estado:           String(row[COL_TRA.ESTADO            - 1] || '').trim(),
        fotos:            String(row[COL_TRA.FOTOS             - 1] || '').trim(),
        foto1:            numCols >= COL_TRA.FOTO_1 ? String(row[COL_TRA.FOTO_1 - 1] || '').trim() : '',
        foto2:            numCols >= COL_TRA.FOTO_2 ? String(row[COL_TRA.FOTO_2 - 1] || '').trim() : '',
        foto3:            numCols >= COL_TRA.FOTO_3 ? String(row[COL_TRA.FOTO_3 - 1] || '').trim() : '',
        foto4:            numCols >= COL_TRA.FOTO_4 ? String(row[COL_TRA.FOTO_4 - 1] || '').trim() : '',
        foto5:            numCols >= COL_TRA.FOTO_5 ? String(row[COL_TRA.FOTO_5 - 1] || '').trim() : ''
      });
    }
  }

  Logger.log('  Trabajos encontrados: ' + trabajos.length);

  // Filtrar por IDs seleccionados (cuando el usuario marcó trabajos concretos en la tabla)
  var selectedIdsStr = getConfig('ids_trabajos_reporte') || '';
  if (selectedIdsStr) {
    var selSet = {};
    selectedIdsStr.split(',').forEach(function(id) { var t = id.trim(); if (t) selSet[t] = true; });
    trabajos = trabajos.filter(function(t) { return !!selSet[String(t.id).trim()]; });
    Logger.log('  Filtrado por IDs seleccionados (' + Object.keys(selSet).length + ') → ' + trabajos.length + ' trabajos');
  }

  if (trabajos.length === 0) {
    Logger.log('  Sin resultados para ese cliente y rango. PDF no generado.');
    logEvent('WARN', 'generateClientReportPDF',
             'Sin trabajos: "' + cliente + '" (' + fechaIniStr + ' → ' + fechaFinStr + ')', '');
    return;
  }

  // ── 3. Calcular totales ───────────────────────────────────────
  Logger.log('\n[3] Calculando totales...');
  var totalHoras = 0;
  for (var j = 0; j < trabajos.length; j++) {
    var h = _calcHoras_(trabajos[j].horaInicio, trabajos[j].horaFin);
    trabajos[j].horas = h;
    totalHoras += h;
  }
  Logger.log('  Total trabajos: ' + trabajos.length);
  Logger.log('  Total horas:    ' + totalHoras.toFixed(2));

  // ── 4. Crear Google Doc ───────────────────────────────────────
  Logger.log('\n[4] Creando Google Doc...');
  var empresa = getConfig('empresa_nombre') || 'Mi Empresa';
  var cliInfPago = null; try { cliInfPago = _getClienteByLabel_(cliente); } catch(_) {}
  var titulo  = 'Informe_' + _sanitizeFilename_(cliente) +
                '_' + fechaIniStr.replace(/\//g, '-') +
                '_a_' + fechaFinStr.replace(/\//g, '-');

  var doc  = DocumentApp.create(titulo);
  var body = doc.getBody();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);

  // Cabecera del documento
  var pEmpresa = body.appendParagraph(empresa);
  pEmpresa.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  pEmpresa.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  var pTitulo = body.appendParagraph('Informe de Trabajos — ' + cliente);
  pTitulo.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  pTitulo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('Periodo: ' + fechaIniStr + ' — ' + fechaFinStr)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph(
    'Generado el ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
  ).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendHorizontalRule();

  // Resumen
  body.appendParagraph('RESUMEN').setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph('Total de trabajos: ' + trabajos.length);
  body.appendParagraph('Total de horas:    ' + totalHoras.toFixed(2) + ' h');
  if (cliInfPago && cliInfPago.metodosPago && cliInfPago.metodosPago.length > 0) {
    body.appendParagraph('Método de pago:    ' + cliInfPago.metodosPago.map(function(m) { return m.tipo; }).join(' / '));
  }

  body.appendHorizontalRule();

  // Detalle de trabajos — tabla principal
  body.appendParagraph('DETALLE DE TRABAJOS').setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph('');

  var tableData = [['Fecha', 'Trabajo realizado', 'Horas', 'Incidencias', 'Estado']];
  for (var k = 0; k < trabajos.length; k++) {
    var t = trabajos[k];
    tableData.push([
      formatDate(t.fechaTrabajo),
      t.trabajoRealizado ? t.trabajoRealizado.substring(0, 80) : '—',
      t.horas > 0 ? t.horas.toFixed(2).replace('.', ',') + ' h' : '0,00 h',
      t.incidencias ? t.incidencias.substring(0, 80) : 'Sin incidencias',
      t.estado || '—'
    ]);
  }

  var tbl = body.appendTable(tableData);
  for (var col = 0; col < 5; col++) {
    var hdrCell = tbl.getRow(0).getCell(col);
    hdrCell.setBackgroundColor('#34a853');
    hdrCell.editAsText().setBold(true).setForegroundColor('#ffffff').setFontSize(9);
  }
  for (var tr = 1; tr < tbl.getNumRows(); tr++) {
    tbl.getRow(tr).editAsText().setFontSize(9);
  }

  // Registro fotográfico (solo si algún trabajo tiene fotos)
  var conFotos = false;
  for (var kf = 0; kf < trabajos.length; kf++) {
    var tf = trabajos[kf];
    if (tf.foto1 || tf.foto2 || tf.foto3 || tf.foto4 || tf.foto5 || tf.fotos) { conFotos = true; break; }
  }

  if (conFotos) {
    body.appendParagraph('');
    body.appendHorizontalRule();
    body.appendParagraph('REGISTRO FOTOGRÁFICO').setHeading(DocumentApp.ParagraphHeading.HEADING3);

    for (var k2 = 0; k2 < trabajos.length; k2++) {
      var t2 = trabajos[k2];
      var fotoList = [t2.foto1, t2.foto2, t2.foto3, t2.foto4, t2.foto5].filter(function(u) { return u; });
      if (fotoList.length === 0 && !t2.fotos) continue;

      body.appendParagraph(
        formatDate(t2.fechaTrabajo) + (t2.trabajoRealizado ? ' — ' + t2.trabajoRealizado.substring(0, 60) : '')
      ).setHeading(DocumentApp.ParagraphHeading.HEADING4);

      if (fotoList.length > 0) {
        for (var fi = 0; fi < fotoList.length; fi++) {
          var fotoIds = _extractDriveFileIds_(fotoList[fi]);
          if (fotoIds.length === 0) {
            body.appendParagraph('  [' + (fi + 1) + '] Foto no accesible');
            continue;
          }
          try {
            var fotoFile = DriveApp.getFileById(fotoIds[0]);
            var blob     = fotoFile.getBlob();
            var img      = body.appendImage(blob);
            var imgW = img.getWidth();
            var imgH = img.getHeight();
            if (imgW > 250) img.setWidth(250).setHeight(Math.round(imgH * 250 / imgW));
          } catch (imgErr) {
            body.appendParagraph('  [' + (fi + 1) + '] Foto no accesible');
          }
        }
      } else if (t2.fotos) {
        body.appendParagraph('Fotos: ' + t2.fotos);
      }
    }
  }

  doc.saveAndClose();
  Logger.log('  Doc creado: "' + titulo + '"  ID=' + doc.getId());

  // ── 5. Convertir a PDF ────────────────────────────────────────
  Logger.log('\n[5] Convirtiendo a PDF...');
  var docFile = DriveApp.getFileById(doc.getId());
  var pdfBlob = docFile.getAs('application/pdf');
  pdfBlob.setName(titulo + '.pdf');

  // ── 6. Guardar PDF en carpeta Drive ──────────────────────────
  Logger.log('\n[6] Guardando PDF en Drive...');
  var carpetaId = getConfig('carpeta_facturas_id');
  var carpeta;

  if (carpetaId) {
    try {
      carpeta = DriveApp.getFolderById(carpetaId);
      Logger.log('  Carpeta: "' + carpeta.getName() + '"');
    } catch (e) {
      Logger.log('  ⚠️ carpeta_facturas_id no accesible. Guardando en raíz Drive.');
      carpeta = DriveApp.getRootFolder();
    }
  } else {
    Logger.log('  ⚠️ carpeta_facturas_id no configurado. Guardando en raíz Drive.');
    carpeta = DriveApp.getRootFolder();
  }

  var pdfFile = carpeta.createFile(pdfBlob);
  var pdfUrl  = pdfFile.getUrl();
  Logger.log('  ✓ PDF guardado: ' + pdfUrl);

  // Eliminar Doc temporal
  docFile.setTrashed(true);
  Logger.log('  ✓ Doc temporal eliminado');

  // ── 7. Registrar en INFORMES ─────────────────────────────────
  Logger.log('\n[7] Registrando en INFORMES...');
  var infNum    = _nextInformeNumber_();
  var infSheet  = _ensureInformesSheet_();
  var cliInfData = _getClienteByLabel_(cliente);
  infSheet.appendRow([
    infNum.id,                                                      // ID_Informe
    new Date(),                                                     // Fecha_Emision
    cliInfData ? cliInfData.id : '',                               // ID_Cliente
    cliente,                                                        // Nombre_Cliente
    fechaIniStr,                                                    // Periodo_Desde
    fechaFinStr,                                                    // Periodo_Hasta
    trabajos.length,                                                // Trabajos_Contados
    totalHoras,                                                     // Horas_Totales
    pdfUrl,                                                         // URL_PDF
    'Emitido',                                                      // Estado
    trabajos.map(function(t) { return t.id; }).filter(Boolean).join(', ')  // IDs_Trabajos
  ]);
  Logger.log('  ✓ Registrado en INFORMES como ' + infNum.id);

  // ── 8. Registrar en LOGS ──────────────────────────────────────
  Logger.log('\n[8] Registrando en LOGS...');
  logEvent(
    'INFO',
    'generateClientReportPDF',
    'Informe PDF | Cliente: ' + cliente +
      ' | Periodo: ' + fechaIniStr + ' → ' + fechaFinStr +
      ' | Trabajos: ' + trabajos.length +
      ' | Horas: ' + totalHoras.toFixed(2) + ' h',
    pdfUrl
  );

  Logger.log('\n================================================');
  Logger.log('INFORME GENERADO');
  Logger.log('  Cliente:       ' + cliente);
  Logger.log('  Periodo:       ' + fechaIniStr + ' → ' + fechaFinStr);
  Logger.log('  Trabajos:      ' + trabajos.length);
  Logger.log('  Horas totales: ' + totalHoras.toFixed(2) + ' h');
  Logger.log('  URL PDF:       ' + pdfUrl);
  Logger.log('================================================');
}

// ─────────────────────────────────────────────────────────────

/**
 * Valida que el módulo de informes puede ejecutarse correctamente.
 * Comprueba hoja TRABAJOS, hoja LOGS, carpeta Drive y claves de CONFIG.
 */
function validateReportModule() {
  Logger.log('================================================');
  Logger.log('validateReportModule()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var moduleOk = true;

  // 1. Hoja TRABAJOS
  Logger.log('\n[1] Hoja TRABAJOS:');
  try {
    var sheet = getSheet('TRABAJOS');
    Logger.log('  ✓ GID=' + sheet.getSheetId() +
               '  filas=' + Math.max(0, sheet.getLastRow() - 1));
  } catch (e) {
    Logger.log('  ❌ NO EXISTE: ' + e.message);
    moduleOk = false;
  }

  // 2. Hoja LOGS
  Logger.log('\n[2] Hoja LOGS:');
  try {
    var logs = getSheet('LOGS');
    Logger.log('  ✓ GID=' + logs.getSheetId());
  } catch (e) {
    Logger.log('  ❌ NO EXISTE: ' + e.message);
    moduleOk = false;
  }

  // 3. Carpeta Drive
  Logger.log('\n[3] Carpeta Drive (carpeta_facturas_id):');
  var carpetaId = getConfig('carpeta_facturas_id');
  var driveOk   = false;
  if (carpetaId) {
    try {
      var carpeta = DriveApp.getFolderById(carpetaId);
      Logger.log('  ✓ "' + carpeta.getName() + '"  ID=' + carpetaId);
      driveOk = true;
    } catch (e) {
      Logger.log('  ⚠️ ID configurado pero no accesible: ' + e.message);
      Logger.log('     El PDF se guardará en la raíz de Drive.');
    }
  } else {
    Logger.log('  ⚠️ No configurado. El PDF se guardará en la raíz de Drive.');
    Logger.log('     Para configurar: añade la clave "carpeta_facturas_id" en CONFIG.');
  }

  // 4. Claves del informe en CONFIG
  Logger.log('\n[4] Claves del informe en CONFIG:');
  var cliente     = getConfig('cliente_reporte');
  var fechaIniStr = getConfig('fecha_inicio_reporte');
  var fechaFinStr = getConfig('fecha_fin_reporte');

  Logger.log('  cliente_reporte:      ' + (cliente     || '(vacío)'));
  Logger.log('  fecha_inicio_reporte: ' + (fechaIniStr || '(vacío) — formato DD/MM/YYYY'));
  Logger.log('  fecha_fin_reporte:    ' + (fechaFinStr || '(vacío) — formato DD/MM/YYYY'));

  var clavesOk = !!(cliente && fechaIniStr && fechaFinStr);
  if (clavesOk) {
    var fi = _parseReportDate_(fechaIniStr);
    var ff = _parseReportDate_(fechaFinStr);
    if (!fi || !ff) {
      Logger.log('  ❌ Formato de fecha incorrecto. Usa DD/MM/YYYY.');
      clavesOk = false;
    } else if (fi > ff) {
      Logger.log('  ⚠️ fecha_inicio es posterior a fecha_fin.');
    } else {
      Logger.log('  ✓ Fechas válidas.');
    }
  } else {
    Logger.log('  ⚠️ Rellena las 3 claves antes de ejecutar generateClientReportPDF().');
  }

  // 5. Acceso a DocumentApp (crea y destruye un Doc de prueba)
  Logger.log('\n[5] Acceso a DocumentApp y DriveApp:');
  try {
    var testDoc = DocumentApp.create('_validateReportModule_test_');
    testDoc.saveAndClose();
    DriveApp.getFileById(testDoc.getId()).setTrashed(true);
    Logger.log('  ✓ DocumentApp y DriveApp accesibles');
  } catch (e) {
    Logger.log('  ❌ ' + e.message);
    moduleOk = false;
  }

  // Resumen
  Logger.log('\n================================================');
  Logger.log('RESULTADO:');
  Logger.log('  TRABAJOS:    ' + (moduleOk    ? '✓' : '❌'));
  Logger.log('  LOGS:        ✓');
  Logger.log('  Drive:       ' + (driveOk     ? '✓' : '⚠️  usará raíz'));
  Logger.log('  CONFIG:      ' + (clavesOk    ? '✓' : '⚠️  claves pendientes'));
  Logger.log('  DocApp:      ' + (moduleOk    ? '✓' : '❌'));
  Logger.log('');
  if (moduleOk && clavesOk) {
    Logger.log('MÓDULO LISTO. Ejecuta generateClientReportPDF().');
  } else if (moduleOk) {
    Logger.log('Infraestructura OK. Rellena las claves en CONFIG y ejecuta generateClientReportPDF().');
  } else {
    Logger.log('HAY PROBLEMAS — revisa los puntos con ❌.');
  }
  Logger.log('================================================');
}

// ── Funciones privadas ────────────────────────────────────────

/**
 * Parsea "DD/MM/YYYY" → Date a medianoche. Devuelve null si no válido.
 */
function _parseReportDate_(str) {
  if (!str) return null;
  var parts = String(str).trim().split('/');
  if (parts.length !== 3) return null;
  var d = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10) - 1;
  var y = parseInt(parts[2], 10);
  if (isNaN(d) || isNaN(m) || isNaN(y)) return null;
  var date = new Date(y, m, d, 0, 0, 0, 0);
  return isNaN(date.getTime()) ? null : date;
}

/**
 * Normaliza el valor de Fecha Trabajo (Date nativa de Sheets o string) → Date.
 */
function _parseTrabajosDate_(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  var d = _parseReportDate_(String(val).trim());
  if (d) return d;
  var generic = new Date(val);
  return isNaN(generic.getTime()) ? null : generic;
}

/**
 * Calcula horas entre dos valores de hora (Date de Sheets o string HH:MM).
 * Devuelve 0 si no puede calcular.
 */
function _calcHoras_(inicio, fin) {
  if (!inicio || !fin) return 0;
  var i = _parseHHMM_(inicio);
  var f = _parseHHMM_(fin);
  if (i === null || f === null) return 0;
  var diff = (f - i) / 60;
  return diff > 0 ? diff : 0;
}

/**
 * Convierte un valor de hora (Date, número fraccionario de día, o string HH:MM)
 * a minutos desde medianoche. Devuelve null si no puede parsear.
 */
function _parseHHMM_(val) {
  if (!val && val !== 0) return null;
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null;
    return val.getHours() * 60 + val.getMinutes();
  }
  if (typeof val === 'number') {
    return Math.round(val * 24 * 60);
  }
  var str   = String(val).trim();
  var match = str.match(/^(\d{1,2}):(\d{2})/);
  if (match) return parseInt(match[1], 10) * 60 + parseInt(match[2], 10);
  var d = new Date(str);
  if (!isNaN(d.getTime())) return d.getHours() * 60 + d.getMinutes();
  return null;
}

/**
 * Formatea un valor de hora para el PDF.
 * Soporta Date, número fraccionario de día, "HH:MM", y strings de fecha completos.
 * Devuelve '' si no hay valor o no puede parsear.
 */
function _formatHoraValue_(val) {
  if (!val && val !== 0) return '';
  var tz = Session.getScriptTimeZone();
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return Utilities.formatDate(val, tz, 'HH:mm');
  }
  if (typeof val === 'number') {
    var totalMin = Math.round(val * 24 * 60);
    return ('0' + Math.floor(totalMin / 60) % 24).slice(-2) + ':' + ('0' + totalMin % 60).slice(-2);
  }
  var str = String(val).trim();
  if (!str) return '';
  if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(str)) return str.substring(0, 5);
  var d = new Date(str);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, 'HH:mm');
  return str;
}

/**
 * Limpia un string para usarlo en nombres de archivo Drive.
 */
function _sanitizeFilename_(str) {
  return String(str || '').replace(/[^a-zA-Z0-9À-ÿÀ-ɏ_-]/g, '_').replace(/_+/g, '_');
}

/**
 * Devuelve el siguiente número de informe del año en curso.
 * Formato: INF-YYYY-NNN. Lee de la hoja INFORMES.
 */
function _nextInformeNumber_() {
  var year   = String(new Date().getFullYear());
  var serie  = 'INF';
  var prefix = serie + '-' + year + '-';
  var sheet  = _ensureInformesSheet_();
  var lastRow = sheet.getLastRow();
  var maxNum  = 0;
  if (lastRow > 1) {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    for (var i = 0; i < ids.length; i++) {
      var idStr = String(ids[i]);
      if (idStr.indexOf(prefix) === 0) {
        var n = parseInt(idStr.split('-').pop(), 10);
        if (!isNaN(n) && n > maxNum) maxNum = n;
      }
    }
  }
  var numero = maxNum + 1;
  return {
    id:     prefix + String(numero).padStart(3, '0'),
    serie:  serie,
    numero: numero
  };
}

/**
 * Devuelve la hoja INFORMES, creándola con cabecera si no existe.
 */
function _ensureInformesSheet_() {
  var ss    = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('INFORMES');
  if (!sheet) {
    sheet = ss.insertSheet('INFORMES');
    var hdrs = SHEET_HEADERS.INFORMES;
    sheet.appendRow(hdrs);
    sheet.getRange(1, 1, 1, hdrs.length)
      .setFontWeight('bold').setBackground('#0f9d58').setFontColor('#ffffff');
  } else if (sheet.getLastRow() === 0) {
    var hdrs2 = SHEET_HEADERS.INFORMES;
    sheet.appendRow(hdrs2);
    sheet.getRange(1, 1, 1, hdrs2.length)
      .setFontWeight('bold').setBackground('#0f9d58').setFontColor('#ffffff');
  } else {
    // Migración: añadir columna IDs_Trabajos si aún no existe
    var existingHdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var hasIdsCol = existingHdrs.some(function(h) { return String(h).trim() === 'IDs_Trabajos'; });
    if (!hasIdsCol) {
      var newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue('IDs_Trabajos')
        .setFontWeight('bold').setBackground('#0f9d58').setFontColor('#ffffff');
    }
  }
  return sheet;
}
