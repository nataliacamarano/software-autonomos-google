// ============================================================
// PanelContableExport.gs — FASE 17.4 + 17.5
// Reúne documentos y genera resúmenes/CSV en carpeta Drive
// ============================================================

// Índices de columna (0-based) para campos de documento
var PCN_IDX_FAC_NOMBRE_CLI = 5;   // FACTURAS col 6  — Nombre_Cliente
var PCN_IDX_FAC_URL_PDF    = 18;  // FACTURAS col 19 — URL_PDF
var PCN_IDX_GAS_PROVEEDOR  = 4;   // GASTOS   col 5  — Proveedor
var PCN_IDX_GAS_FOTO_URL   = 15;  // GASTOS   col 16 — Foto_URL
var PCN_IDX_GAS_ARCHIVO_ID = 16;  // GASTOS   col 17 — Archivo_ID

// Clave para las estadísticas del último paquete
var PCN_PROP_STATS = 'pcn_stats_ultimo_paquete';

// ============================================================
// Orquestador principal
// ============================================================

/** Wrapper público para menú. */
function generarPaqueteContable() {
  _pcnGenerarPaqueteContable_();
}

/**
 * Genera el paquete contable completo:
 *   1. Actualiza KPIs
 *   2. Prepara carpeta Drive
 *   3. Copia PDFs de facturas → Facturas Emitidas/
 *   4. Copia adjuntos/fotos de gastos → Gastos/
 *   5. Crea hojas-copia de datos (facturas, gastos, opcionales)
 *   6. Genera Resumen contable, IVA y CSV
 *   7. Guarda estadísticas y actualiza el panel
 */
function _pcnGenerarPaqueteContable_() {
  var ss = getTargetSS_();
  var sh = ss.getSheetByName(PCN_SHEET);
  if (!sh) {
    SpreadsheetApp.getUi().alert(
      '⚠ PANEL_CONTABLE no existe.\n' +
      'Ejecuta: ⚙️ Sistema → 📦 PANEL_CONTABLE — Crear completo'
    );
    return;
  }

  var t = function(msg) {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, '📦 Paquete contable', 8);
  };

  try {
    // 1. KPIs
    t('1/7 — Actualizando KPIs...');
    var kpis    = _pcnActualizarResumen_(ss, sh);
    var tipo    = _pcnLeerTipoPeriodo_(sh);
    var periodo = _pcnPeriodoEtiqueta_(tipo);
    var rango   = _pcnObtenerRango_(sh);
    var empresa = _pcnGetEmpresaNombre_();

    // 2. Carpeta Drive
    t('2/7 — Preparando carpeta Drive...');
    var perFolder  = _pcnResolvePeriodFolder_(sh);
    var folderUrl  = perFolder.getUrl();
    var urlActual  = String(sh.getRange(PCN_ROWS.DRIVE_ROW, 11).getValue() || '').trim();
    if (!urlActual || urlActual === '—') {
      sh.getRange(PCN_ROWS.DRIVE_ROW, 11, 1, 4).merge()
        .setValue(folderUrl)
        .setBackground(PCN_C.AZUL_LIGHT).setFontColor(PCN_C.AZUL_CORP)
        .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
    }

    // 3+4. Leer datos del periodo
    var facData = _pcnGetFacturasData_(ss, rango);
    var gasData = _pcnGetGastosData_(ss, rango);

    // FASE 18.1: Filtrar por selección del usuario
    var sel = _pcnGetDocumentosSeleccionados_(sh, facData, gasData);

    if (sel.stale) {
      var ui = SpreadsheetApp.getUi();
      var cont = ui.alert(
        '⚠ El selector de documentos no está actualizado para ' + periodo + '.\n\n' +
        'Se incluirán TODOS los documentos del periodo:\n' +
        '  · ' + facData.rows.length + ' factura(s)\n' +
        '  · ' + gasData.rows.length + ' gasto(s)\n\n' +
        'Para seleccionar individualmente pulsa Cancelar,\n' +
        'ejecuta "Actualizar" y desmarca los que no quieras.\n\n' +
        '¿Continuar con todos?',
        ui.ButtonSet.YES_NO);
      if (cont !== ui.Button.YES) return;
    }

    var facFilt = sel.facDataFilt;
    var gasFilt = sel.gasDataFilt;

    // 3. Copiar PDFs de facturas seleccionadas
    t('3/7 — Copiando PDFs de facturas (' + sel.facSelCount + '/' +
      sel.facTotal + ')...');
    var facFolder   = _pcnGetOrCreateFolder_(perFolder, 'Facturas Emitidas');
    var facDocStats = _pcnCopiarPDFsFacturas_(facFilt, facFolder);

    // 4. Copiar adjuntos de gastos seleccionados
    t('4/7 — Copiando adjuntos de gastos (' + sel.gasSelCount + '/' +
      sel.gasTotal + ')...');
    var gasFolder   = _pcnGetOrCreateFolder_(perFolder, 'Gastos');
    var gasDocStats = _pcnCopiarAdjuntosGastos_(gasFilt, gasFolder);

    // 5. Hojas-copia de datos (solo seleccionados)
    t('5/7 — Creando hojas de datos...');
    _pcnReunirDocumentos_(ss, sh, perFolder, periodo, facFilt, gasFilt,
                          facFolder, gasFolder);

    // 6. Resúmenes y CSV (solo seleccionados)
    t('6/7 — Generando resúmenes e IVA...');
    _pcnGenerarResumenes_(sh, perFolder, periodo, empresa, kpis, facFilt, gasFilt);

    // 7. Guardar estadísticas y actualizar panel
    t('7/7 — Finalizando...');
    var stats = {
      periodo:      periodo,
      ts:           Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
      facTotal:     sel.facTotal,
      facSelected:  sel.facSelCount,
      facPdfOk:     facDocStats.ok,
      facPdfSinUrl: facDocStats.sinUrl,
      facPdfError:  facDocStats.errores,
      gasTotal:     sel.gasTotal,
      gasSelected:  sel.gasSelCount,
      gasAdjOk:     gasDocStats.ok,
      gasSinAdj:    gasDocStats.sinAdj,
      gasAdjError:  gasDocStats.errores,
      folderUrl:    folderUrl
    };
    PropertiesService.getScriptProperties().setProperty(
      PCN_PROP_STATS, JSON.stringify(stats)
    );

    var estadoTexto = '✅ Paquete listo · ' + periodo +
      ' · PDFs: ' + facDocStats.ok + '/' + sel.facSelCount + ' fac' +
      ' · adj: ' + gasDocStats.ok + '/' + sel.gasSelCount + ' gas';
    _pcnActualizarEstadoDrive_(sh, estadoTexto);
    SpreadsheetApp.flush();

    logEvent('INFO', 'generarPaqueteContable', 'Paquete: ' + periodo,
      'fac=' + facDocStats.ok + '/' + sel.facSelCount +
      '(' + sel.facTotal + ') gas=' + gasDocStats.ok + '/' + sel.gasSelCount +
      '(' + sel.gasTotal + ')');

    // Resumen final
    var noLoc = facDocStats.errores + gasDocStats.errores;
    SpreadsheetApp.getUi().alert(
      '✅ Paquete contable listo.\n\n' +
      '  Periodo:             ' + periodo + '\n' +
      '  Facturas detectadas: ' + sel.facTotal + '\n' +
      '  Facturas incluidas:  ' + sel.facSelCount + '\n' +
      '  PDFs copiados:       ' + facDocStats.ok +
        (facDocStats.sinUrl > 0 ? ' · ' + facDocStats.sinUrl + ' sin PDF' : '') + '\n' +
      '  Gastos detectados:   ' + sel.gasTotal + '\n' +
      '  Gastos incluidos:    ' + sel.gasSelCount + '\n' +
      '  Adjuntos copiados:   ' + gasDocStats.ok +
        (gasDocStats.sinAdj > 0 ? ' · ' + gasDocStats.sinAdj + ' sin adjunto' : '') + '\n' +
      (noLoc > 0 ? '  ⚠ No localizados:    ' + noLoc + ' archivo(s)\n' : '') +
      '  IVA neto:            ' + _pcnFmtEur_(kpis.ivaNeto) + '\n\n' +
      'Carpeta Drive preparada.\nUsa ☑ Enviar email al gestor para finalizar.'
    );

  } catch (e) {
    Logger.log('ERROR _pcnGenerarPaqueteContable_: ' + e.message + '\n' + e.stack);
    logEvent('ERROR', 'generarPaqueteContable', e.message, '');
    SpreadsheetApp.getUi().alert(
      '❌ Error al generar el paquete:\n\n' + e.message +
      '\n\nRevisa el log: Ver → Registros de ejecución.'
    );
  }
}

// ── Helpers de carpeta ────────────────────────────────────────

function _pcnResolvePeriodFolder_(sh) {
  var tipo    = _pcnLeerTipoPeriodo_(sh);
  var anio    = String(new Date().getFullYear());
  var periodo = _pcnPeriodoEtiqueta_(tipo);
  var root    = DriveApp.getRootFolder();
  var cont    = _pcnGetOrCreateFolder_(root, PCN_DRIVE_ROOT);
  var year    = _pcnGetOrCreateFolder_(cont, anio);
  return _pcnGetOrCreateFolder_(year, periodo);
}

function _pcnActualizarEstadoDrive_(sh, texto) {
  sh.getRange(PCN_ROWS.DRIVE_ROW, 4, 1, 4).merge()
    .setValue(texto)
    .setBackground(PCN_C.VERDE_BG).setFontColor(PCN_C.VERDE_DARK)
    .setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
}

// ============================================================
// COPIA DE DOCUMENTOS REALES
// ============================================================

/**
 * Copia el PDF de cada factura del periodo a destFolder.
 * Lee URL_PDF (col 19, idx 18) de cada fila.
 * Devuelve { ok, sinUrl, errores, listaErr }.
 */
function _pcnCopiarPDFsFacturas_(facData, destFolder) {
  var ok = 0, sinUrl = 0, errores = 0, listaErr = [];

  facData.rows.forEach(function(r) {
    var id      = String(r[0] || '').trim();
    var cliente = String(r.length > PCN_IDX_FAC_NOMBRE_CLI
      ? r[PCN_IDX_FAC_NOMBRE_CLI] : '').trim();
    var urlPdf  = String(r.length > PCN_IDX_FAC_URL_PDF
      ? r[PCN_IDX_FAC_URL_PDF] : '').trim();

    if (!urlPdf || urlPdf === '—') { sinUrl++; return; }

    var fileId = _pcnExtraerFileId_(urlPdf);
    if (!fileId) { sinUrl++; return; }

    var nombre = _pcnSanitizarNombre_(id +
      (cliente ? ' - ' + cliente.substring(0, 28) : ''));
    var copiado = _pcnCopiarArchivoADrive_(fileId, destFolder, nombre);
    if (copiado) { ok++; } else { errores++; listaErr.push(id); }
  });

  return { ok: ok, sinUrl: sinUrl, errores: errores, listaErr: listaErr };
}

/**
 * Copia el adjunto (foto/PDF) de cada gasto del periodo a destFolder.
 * Prefiere Archivo_ID (col 17, idx 16); si vacío usa Foto_URL (col 16, idx 15).
 * Devuelve { ok, sinAdj, errores, listaErr }.
 */
function _pcnCopiarAdjuntosGastos_(gasData, destFolder) {
  var ok = 0, sinAdj = 0, errores = 0, listaErr = [];
  var tz = Session.getScriptTimeZone();

  gasData.rows.forEach(function(r) {
    var id         = String(r[0] || '').trim();
    var proveedor  = String(r.length > PCN_IDX_GAS_PROVEEDOR
      ? r[PCN_IDX_GAS_PROVEEDOR] : '').trim();
    var fecha      = r[3]; // Fecha_Gasto idx 3
    var archivoId  = String(r.length > PCN_IDX_GAS_ARCHIVO_ID
      ? r[PCN_IDX_GAS_ARCHIVO_ID] : '').trim();
    var fotoUrl    = String(r.length > PCN_IDX_GAS_FOTO_URL
      ? r[PCN_IDX_GAS_FOTO_URL] : '').trim();

    var fileId = archivoId || _pcnExtraerFileId_(fotoUrl);
    if (!fileId) { sinAdj++; return; }

    var fechaStr = (fecha instanceof Date && !isNaN(fecha.getTime()))
      ? Utilities.formatDate(fecha, tz, 'yyyy-MM-dd') : '';
    var nombre = _pcnSanitizarNombre_(id +
      (proveedor ? ' - ' + proveedor.substring(0, 22) : '') +
      (fechaStr ? ' ' + fechaStr : ''));

    var copiado = _pcnCopiarArchivoADrive_(fileId, destFolder, nombre);
    if (copiado) { ok++; } else { errores++; listaErr.push(id); }
  });

  return { ok: ok, sinAdj: sinAdj, errores: errores, listaErr: listaErr };
}

/**
 * Copia un archivo Drive a la carpeta destino con el nombre dado.
 * Si ya existe un archivo con ese nombre (mismo base), lo omite.
 * Añade automáticamente la extensión del archivo original.
 * Devuelve true si el archivo quedó en destFolder, null si falló.
 */
function _pcnCopiarArchivoADrive_(fileId, destFolder, nombreBase) {
  try {
    var file = DriveApp.getFileById(fileId);

    // Detectar extensión a partir del nombre original
    var origName = file.getName();
    var dotIdx   = origName.lastIndexOf('.');
    var ext      = (dotIdx > 0 && dotIdx < origName.length - 1)
      ? origName.substring(dotIdx)   // incluye el punto
      : '';
    var nombreFinal = nombreBase + ext;

    // Saltar si ya existe (re-runs son idempotentes)
    var existing = destFolder.getFilesByName(nombreFinal);
    if (existing.hasNext()) return true;

    file.makeCopy(nombreFinal, destFolder);
    return true;
  } catch (e) {
    Logger.log('WARN _pcnCopiarArchivoADrive_ [' + fileId + ']: ' + e.message);
    return null;
  }
}

/**
 * Extrae el ID de un archivo Drive a partir de su URL.
 * Soporta los formatos más comunes de Google Drive.
 */
function _pcnExtraerFileId_(url) {
  if (!url) return null;
  var s = String(url).trim();
  if (!s || s === '—') return null;

  // https://drive.google.com/file/d/ID/view
  var m = s.match(/\/file\/d\/([a-zA-Z0-9_-]{15,})/);
  if (m) return m[1];

  // https://drive.google.com/open?id=ID  o  ...&id=ID
  m = s.match(/[?&]id=([a-zA-Z0-9_-]{15,})/);
  if (m) return m[1];

  // https://docs.google.com/*/d/ID/
  m = s.match(/\/d\/([a-zA-Z0-9_-]{15,})\//);
  if (m) return m[1];

  // ID puro (25+ chars alfanuméricos)
  if (/^[a-zA-Z0-9_-]{25,}$/.test(s)) return s;

  return null;
}

/** Elimina caracteres no válidos para nombres de archivo en Drive. */
function _pcnSanitizarNombre_(nombre) {
  return nombre.replace(/[\/\\:*?"<>|]/g, '-').trim();
}

// ============================================================
// FASE 17.4 — Hojas-copia de datos (spreadsheets)
// ============================================================

/**
 * Crea copias en spreadsheet de Facturas, Gastos y opcionales.
 * Las subcarpetas Facturas Emitidas/ y Gastos/ ya existen (se pasan).
 */
function _pcnReunirDocumentos_(ss, sh, perFolder, periodo, facData, gasData,
                                facFolder, gasFolder) {
  facFolder = facFolder || _pcnGetOrCreateFolder_(perFolder, 'Facturas Emitidas');
  gasFolder = gasFolder || _pcnGetOrCreateFolder_(perFolder, 'Gastos');

  // Hojas de datos (todos los campos, útiles como índice para el gestor)
  var chkFac = sh.getRange(PCN_DOCS.FACTURAS[0], PCN_DOCS.FACTURAS[1]).getValue();
  if (chkFac !== false && facData.headers.length > 0) {
    _pcnCrearSheetEnDrive_(facFolder, 'Índice_Facturas_' + periodo,
      facData.headers, facData.rows);
  }

  var chkGas = sh.getRange(PCN_DOCS.GASTOS[0], PCN_DOCS.GASTOS[1]).getValue();
  if (chkGas !== false && gasData.headers.length > 0) {
    _pcnCrearSheetEnDrive_(gasFolder, 'Índice_Gastos_' + periodo,
      gasData.headers, gasData.rows);
  }

  // Proformas (sin filtro de fecha)
  if (sh.getRange(PCN_DOCS.PROFORMAS[0], PCN_DOCS.PROFORMAS[1]).getValue()) {
    var profData = _pcnGetAllSheetData_(ss, 'PROFORMAS');
    if (profData.rows.length > 0) {
      _pcnCrearSheetEnDrive_(_pcnGetOrCreateFolder_(perFolder, 'Proformas'),
        'Proformas_' + periodo, profData.headers, profData.rows);
    }
  }

  // Presupuestos (sin filtro de fecha)
  if (sh.getRange(PCN_DOCS.PRESUPUESTOS[0], PCN_DOCS.PRESUPUESTOS[1]).getValue()) {
    var presData = _pcnGetAllSheetData_(ss, 'PRESUPUESTOS');
    if (presData.rows.length > 0) {
      _pcnCrearSheetEnDrive_(_pcnGetOrCreateFolder_(perFolder, 'Presupuestos'),
        'Presupuestos_' + periodo, presData.headers, presData.rows);
    }
  }
}

// ── Lectores de datos ─────────────────────────────────────────

function _pcnGetFacturasData_(ss, rango) {
  var src = ss.getSheetByName('FACTURAS');
  if (!src || src.getLastRow() < 1) return { headers: [], rows: [] };
  var lastCol = src.getLastColumn();
  if (lastCol < 1) return { headers: [], rows: [] };

  var headers = src.getRange(1, 1, 1, lastCol).getValues()[0];
  if (src.getLastRow() < 2) return { headers: headers, rows: [] };

  var all  = src.getRange(2, 1, src.getLastRow() - 1, lastCol).getValues();
  var rows = all.filter(function(r) {
    if (!String(r[0]).trim()) return false;
    var estado = lastCol > 17 ? String(r[17]).trim() : '';
    if (estado === 'Cancelada') return false;
    return _pcnEnRango_(r[1], rango);
  });

  return { headers: headers, rows: rows };
}

function _pcnGetGastosData_(ss, rango) {
  var src = ss.getSheetByName('GASTOS');
  if (!src || src.getLastRow() < 1) return { headers: [], rows: [] };
  var lastCol = src.getLastColumn();
  if (lastCol < 1) return { headers: [], rows: [] };

  var headers = src.getRange(1, 1, 1, lastCol).getValues()[0];
  if (src.getLastRow() < 2) return { headers: headers, rows: [] };

  var all  = src.getRange(2, 1, src.getLastRow() - 1, lastCol).getValues();
  var rows = all.filter(function(r) {
    if (!String(r[0]).trim()) return false;
    return _pcnEnRango_(r[3], rango);
  });

  return { headers: headers, rows: rows };
}

function _pcnGetAllSheetData_(ss, sheetName) {
  var src = ss.getSheetByName(sheetName);
  if (!src || src.getLastRow() < 1) return { headers: [], rows: [] };
  var lastCol = src.getLastColumn();
  if (lastCol < 1) return { headers: [], rows: [] };

  var headers = src.getRange(1, 1, 1, lastCol).getValues()[0];
  if (src.getLastRow() < 2) return { headers: headers, rows: [] };

  var rows = src.getRange(2, 1, src.getLastRow() - 1, lastCol).getValues()
    .filter(function(r) { return String(r[0]).trim() !== ''; });

  return { headers: headers, rows: rows };
}

// ── Crear Google Sheet en carpeta Drive ───────────────────────

function _pcnCrearSheetEnDrive_(folder, nombre, headers, rows) {
  var existing = folder.getFilesByName(nombre);
  while (existing.hasNext()) { existing.next().setTrashed(true); }

  var newSS = SpreadsheetApp.create(nombre);
  var sheet = newSS.getActiveSheet();
  sheet.setName('Datos');

  var nc = headers.length;
  if (nc > 0) {
    sheet.getRange(1, 1, 1, nc).setValues([headers])
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  if (rows.length > 0 && nc > 0) {
    sheet.getRange(2, 1, rows.length, nc).setValues(rows);
  }
  if (nc > 0) sheet.autoResizeColumns(1, Math.min(nc, 26));

  var file = DriveApp.getFileById(newSS.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  return newSS.getUrl();
}

// ============================================================
// FASE 17.5 — Resúmenes y CSV
// ============================================================

function _pcnGenerarResumenes_(sh, perFolder, periodo, empresa, kpis, facData, gasData) {
  var resFolder = _pcnGetOrCreateFolder_(perFolder, 'Resumen');
  var ivaFolder = _pcnGetOrCreateFolder_(perFolder, 'IVA');

  _pcnCrearResumenContable_(resFolder, 'Resumen_' + periodo, periodo, empresa, kpis);

  if (sh.getRange(PCN_DOCS.RESUMEN_IVA[0], PCN_DOCS.RESUMEN_IVA[1]).getValue() !== false) {
    _pcnCrearResumenIVA_(ivaFolder, 'IVA_' + periodo, periodo, empresa, kpis,
                         facData, gasData);
  }

  if (sh.getRange(PCN_DOCS.CSV_FAC[0], PCN_DOCS.CSV_FAC[1]).getValue()) {
    _pcnCrearCSVEnDrive_(resFolder, 'Facturas_' + periodo + '.csv',
      facData.headers, facData.rows);
  }
  if (sh.getRange(PCN_DOCS.CSV_GAS[0], PCN_DOCS.CSV_GAS[1]).getValue()) {
    _pcnCrearCSVEnDrive_(resFolder, 'Gastos_' + periodo + '.csv',
      gasData.headers, gasData.rows);
  }
}

// ── Resumen contable ──────────────────────────────────────────

function _pcnCrearResumenContable_(folder, nombre, periodo, empresa, kpis) {
  var existing = folder.getFilesByName(nombre);
  while (existing.hasNext()) { existing.next().setTrashed(true); }

  var newSS = SpreadsheetApp.create(nombre);
  var sh    = newSS.getActiveSheet();
  sh.setName('Resumen');

  var tz  = Session.getScriptTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');

  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 160);

  sh.getRange(1, 1, 1, 2).merge()
    .setValue('📦 RESUMEN CONTABLE — ' + periodo)
    .setBackground('#4a148c').setFontColor('#ffffff')
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 48);

  var info = [
    ['Empresa',              empresa],
    ['Periodo',              periodo],
    ['Generado',             now],
    ['',                     ''],
    ['📄 FACTURAS EMITIDAS', ''],
    ['Nº facturas',          kpis.facCount],
    ['Total facturado',      kpis.facTotal],
    ['IVA repercutido',      kpis.ivaRep],
    ['Pendiente de cobro',   kpis.pendiente],
    ['',                     ''],
    ['💸 GASTOS / TICKETS',  ''],
    ['Nº gastos',            kpis.gasCount],
    ['Total gastos',         kpis.gasTotal],
    ['IVA soportado',        kpis.ivaSop],
    ['',                     ''],
    ['🧾 LIQUIDACIÓN IVA',   ''],
    ['IVA repercutido (+)',   kpis.ivaRep],
    ['IVA soportado (−)',     kpis.ivaSop],
    ['IVA a declarar',        kpis.ivaNeto]
  ];

  info.forEach(function(row, i) {
    sh.getRange(2 + i, 1).setValue(row[0]);
    sh.getRange(2 + i, 2).setValue(row[1]);
  });

  [6, 12, 17].forEach(function(r) {
    sh.getRange(r, 1, 1, 2).merge()
      .setBackground('#e8f0fe').setFontWeight('bold').setFontSize(11);
  });

  [8, 9, 10, 14, 15, 18, 19, 20].forEach(function(r) {
    var v = sh.getRange(r, 2).getValue();
    if (typeof v === 'number') sh.getRange(r, 2).setNumberFormat('#,##0.00 "€"');
  });

  var bgIva = kpis.ivaNeto >= 0 ? '#fce8e6' : '#e6f4ea';
  sh.getRange(20, 1, 1, 2).setBackground(bgIva).setFontWeight('bold');

  sh.autoResizeColumns(1, 2);

  var file = DriveApp.getFileById(newSS.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  SpreadsheetApp.flush();
}

// ── Resumen IVA ───────────────────────────────────────────────

function _pcnCrearResumenIVA_(folder, nombre, periodo, empresa, kpis, facData, gasData) {
  var existing = folder.getFilesByName(nombre);
  while (existing.hasNext()) { existing.next().setTrashed(true); }

  var newSS = SpreadsheetApp.create(nombre);
  var sh    = newSS.getActiveSheet();
  sh.setName('IVA');

  var tz = Session.getScriptTimeZone();
  [120, 100, 140, 130, 130, 100].forEach(function(w, i) {
    sh.setColumnWidth(i + 1, w);
  });

  sh.getRange(1, 1, 1, 6).merge()
    .setValue('🧾 RESUMEN IVA — ' + periodo + ' — ' + empresa)
    .setBackground('#4a148c').setFontColor('#ffffff')
    .setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 40);

  var row = 2;

  // Facturas
  sh.getRange(row, 1, 1, 6).merge()
    .setValue('  📄 IVA REPERCUTIDO — FACTURAS EMITIDAS')
    .setBackground('#e8f0fe').setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(row, 28); row++;

  sh.getRange(row, 1, 1, 6).setValues([
    ['Nº Factura', 'Fecha', 'Base Imponible', 'IVA (€)', 'Total', 'Estado']
  ]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  row++;

  var facIvaRows = facData.rows.map(function(r) {
    var total = parseFloat(r.length > 16 ? r[16] : 0) || 0;
    var iva   = parseFloat(r.length > 13 ? r[13] : 0) || 0;
    var base  = total - iva;
    var fecha = r[1] instanceof Date
      ? Utilities.formatDate(r[1], tz, 'dd/MM/yyyy') : String(r[1] || '');
    return [r[0], fecha, base, iva, total, r.length > 17 ? r[17] : ''];
  });

  if (facIvaRows.length > 0) {
    sh.getRange(row, 1, facIvaRows.length, 6).setValues(facIvaRows);
    sh.getRange(row, 3, facIvaRows.length, 3).setNumberFormat('#,##0.00 "€"');
    row += facIvaRows.length;
  }

  sh.getRange(row, 1, 1, 3).merge()
    .setValue('TOTAL IVA REPERCUTIDO')
    .setFontWeight('bold').setBackground('#e8f0fe');
  sh.getRange(row, 4).setValue(kpis.ivaRep)
    .setNumberFormat('#,##0.00 "€"').setFontWeight('bold').setBackground('#e8f0fe');
  row += 2;

  // Gastos
  sh.getRange(row, 1, 1, 5).merge()
    .setValue('  💸 IVA SOPORTADO — GASTOS / TICKETS')
    .setBackground('#fef7e0').setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(row, 28); row++;

  sh.getRange(row, 1, 1, 5).setValues([
    ['Nº Gasto', 'Fecha', 'Base Imponible', 'IVA (€)', 'Total']
  ]).setFontWeight('bold').setBackground('#e37400').setFontColor('#ffffff');
  row++;

  var gasIvaRows = gasData.rows.map(function(r) {
    var total = parseFloat(r.length > 12 ? r[12] : 0) || 0;
    var iva   = parseFloat(r.length > 11 ? r[11] : 0) || 0;
    var base  = total - iva;
    var fecha = r[3] instanceof Date
      ? Utilities.formatDate(r[3], tz, 'dd/MM/yyyy') : String(r[3] || '');
    return [r[0], fecha, base, iva, total];
  });

  if (gasIvaRows.length > 0) {
    sh.getRange(row, 1, gasIvaRows.length, 5).setValues(gasIvaRows);
    sh.getRange(row, 3, gasIvaRows.length, 3).setNumberFormat('#,##0.00 "€"');
    row += gasIvaRows.length;
  }

  sh.getRange(row, 1, 1, 3).merge()
    .setValue('TOTAL IVA SOPORTADO')
    .setFontWeight('bold').setBackground('#fef7e0');
  sh.getRange(row, 4).setValue(kpis.ivaSop)
    .setNumberFormat('#,##0.00 "€"').setFontWeight('bold').setBackground('#fef7e0');
  row += 2;

  // Liquidación
  sh.getRange(row, 1, 1, 3).merge()
    .setValue('🧾 LIQUIDACIÓN IVA')
    .setBackground('#4a148c').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11);
  row++;

  [['IVA Repercutido (+)', kpis.ivaRep, '#f8f9fa'],
   ['IVA Soportado (−)',   kpis.ivaSop, '#f8f9fa'],
   ['IVA A DECLARAR',      kpis.ivaNeto, kpis.ivaNeto >= 0 ? '#fce8e6' : '#e6f4ea']
  ].forEach(function(lr, i) {
    sh.getRange(row, 1).setValue(lr[0])
      .setBackground(lr[2]).setFontWeight(i === 2 ? 'bold' : 'normal');
    sh.getRange(row, 3).setValue(lr[1])
      .setNumberFormat('#,##0.00 "€"')
      .setBackground(lr[2]).setFontWeight(i === 2 ? 'bold' : 'normal');
    row++;
  });

  var file = DriveApp.getFileById(newSS.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  SpreadsheetApp.flush();
}

// ── CSV ───────────────────────────────────────────────────────

function _pcnCrearCSVEnDrive_(folder, nombre, headers, rows) {
  var existing = folder.getFilesByName(nombre);
  while (existing.hasNext()) { existing.next().setTrashed(true); }

  var tz    = Session.getScriptTimeZone();
  var lines = [];
  if (headers.length > 0) lines.push(_pcnRowToCSV_(headers, tz));
  rows.forEach(function(r) { lines.push(_pcnRowToCSV_(r, tz)); });

  var blob = Utilities.newBlob(
    '﻿' + lines.join('\r\n'),
    'text/csv; charset=utf-8',
    nombre
  );
  folder.createFile(blob);
}

function _pcnRowToCSV_(row, tz) {
  return row.map(function(cell) {
    var s;
    if (cell instanceof Date) {
      s = Utilities.formatDate(cell, tz, 'dd/MM/yyyy');
    } else if (cell === null || cell === undefined) {
      s = '';
    } else {
      s = String(cell);
    }
    if (s.indexOf(',') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1) {
      s = '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  }).join(',');
}

// ============================================================
// Leer estadísticas del último paquete generado
// ============================================================

/**
 * Devuelve el objeto stats del último paquete, o null si no existe.
 * El objeto tiene: periodo, ts, facTotal, facPdfOk, facPdfSinUrl,
 * facPdfError, gasTotal, gasAdjOk, gasSinAdj, gasAdjError, folderUrl.
 */
function _pcnLeerStatsUltimoPaquete_() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(PCN_PROP_STATS);
    return raw ? JSON.parse(raw) : null;
  } catch (_) {
    return null;
  }
}
