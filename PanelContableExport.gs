// ============================================================
// PanelContableExport.gs — FASE 17.4 + 17.5
// Reúne documentos y genera resúmenes/CSV en carpeta Drive
// ============================================================

// ============================================================
// FASE 17.4 — Orquestador del paquete contable
// ============================================================

/** Wrapper público para menú. */
function generarPaqueteContable() {
  _pcnGenerarPaqueteContable_();
}

/**
 * Orquestador principal: actualiza KPIs → prepara carpeta →
 * reúne docs (17.4) → genera resúmenes y CSV (17.5).
 * Llamado desde el botón "📊 Generar resumen contable".
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
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, '📦 Paquete contable', 6);
  };

  try {
    // 1. KPIs actualizados
    t('1/5 — Actualizando KPIs del periodo...');
    var kpis    = _pcnActualizarResumen_(ss, sh);
    var tipo    = _pcnLeerTipoPeriodo_(sh);
    var periodo = _pcnPeriodoEtiqueta_(tipo);
    var rango   = _pcnObtenerRango_(sh);
    var empresa = _pcnGetEmpresaNombre_();

    // 2. Carpeta Drive (crear si no existe, asegurar URL en panel)
    t('2/5 — Preparando carpeta Drive...');
    var perFolder = _pcnResolvePeriodFolder_(sh);
    var urlActual = String(sh.getRange(PCN_ROWS.DRIVE_ROW, 11).getValue() || '').trim();
    if (!urlActual || urlActual === '—') {
      sh.getRange(PCN_ROWS.DRIVE_ROW, 11, 1, 4).merge()
        .setValue(perFolder.getUrl())
        .setBackground(PCN_C.AZUL_LIGHT).setFontColor(PCN_C.AZUL_CORP)
        .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
    }

    // 3. Leer datos una sola vez (reutilizado en 17.4 y 17.5)
    var facData = _pcnGetFacturasData_(ss, rango);
    var gasData = _pcnGetGastosData_(ss, rango);

    // 4. Reunir documentos (FASE 17.4)
    t('3/5 — Reuniendo documentos en Drive...');
    var docsResult = _pcnReunirDocumentos_(ss, sh, perFolder, periodo, facData, gasData);

    // 5. Resúmenes y CSV (FASE 17.5)
    t('4/5 — Generando resúmenes e IVA...');
    _pcnGenerarResumenes_(sh, perFolder, periodo, empresa, kpis, facData, gasData);

    // 6. Finalizar estado en panel
    t('5/5 — Finalizando...');
    _pcnActualizarEstadoDrive_(sh,
      '✅ Paquete listo · ' + periodo +
      ' · ' + docsResult.facCount + ' fac · ' + docsResult.gasCount + ' gas');
    SpreadsheetApp.flush();

    logEvent('INFO', 'generarPaqueteContable',
      'Paquete: ' + periodo,
      docsResult.facCount + ' fac, ' + docsResult.gasCount + ' gas');

    SpreadsheetApp.getUi().alert(
      '✅ Paquete contable generado.\n\n' +
      '  Periodo:  ' + periodo + '\n' +
      '  Facturas: ' + docsResult.facCount + '\n' +
      '  Gastos:   ' + docsResult.gasCount + '\n' +
      '  IVA neto: ' + _pcnFmtEur_(kpis.ivaNeto) + '\n\n' +
      'Abre la carpeta Drive para verificar los archivos.\n' +
      'Cuando esté listo, usa ☑ Enviar email al gestor.'
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

/**
 * Devuelve la carpeta del periodo, creándola si no existe.
 * Ruta: Contabilidad / AÑO / PERIODO
 */
function _pcnResolvePeriodFolder_(sh) {
  var tipo    = _pcnLeerTipoPeriodo_(sh);
  var anio    = String(new Date().getFullYear());
  var periodo = _pcnPeriodoEtiqueta_(tipo);
  var root    = DriveApp.getRootFolder();
  var cont    = _pcnGetOrCreateFolder_(root, PCN_DRIVE_ROOT);
  var year    = _pcnGetOrCreateFolder_(cont, anio);
  return _pcnGetOrCreateFolder_(year, periodo);
}

/** Actualiza la celda Estado en la sección CARPETA DRIVE. */
function _pcnActualizarEstadoDrive_(sh, texto) {
  sh.getRange(PCN_ROWS.DRIVE_ROW, 4, 1, 4).merge()
    .setValue(texto)
    .setBackground(PCN_C.VERDE_BG).setFontColor(PCN_C.VERDE_DARK)
    .setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
}

// ============================================================
// FASE 17.4 — Reunir documentos
// ============================================================

/**
 * Crea copias de Facturas y Gastos (y opcionales) en las subcarpetas Drive.
 * Regla: NO modifica los datos originales, solo crea copias.
 * Devuelve { facCount, gasCount }.
 */
function _pcnReunirDocumentos_(ss, sh, perFolder, periodo, facData, gasData) {
  // Facturas Emitidas
  var chkFac = sh.getRange(PCN_DOCS.FACTURAS[0], PCN_DOCS.FACTURAS[1]).getValue();
  if (chkFac !== false) {
    var facFolder = _pcnGetOrCreateFolder_(perFolder, 'Facturas Emitidas');
    if (facData.headers.length > 0) {
      _pcnCrearSheetEnDrive_(facFolder, 'Facturas_' + periodo, facData.headers, facData.rows);
    }
  }

  // Gastos
  var chkGas = sh.getRange(PCN_DOCS.GASTOS[0], PCN_DOCS.GASTOS[1]).getValue();
  if (chkGas !== false) {
    var gasFolder = _pcnGetOrCreateFolder_(perFolder, 'Gastos');
    if (gasData.headers.length > 0) {
      _pcnCrearSheetEnDrive_(gasFolder, 'Gastos_' + periodo, gasData.headers, gasData.rows);
    }
  }

  // Proformas (sin filtro de fecha)
  var chkProf = sh.getRange(PCN_DOCS.PROFORMAS[0], PCN_DOCS.PROFORMAS[1]).getValue();
  if (chkProf) {
    var profData = _pcnGetAllSheetData_(ss, 'PROFORMAS');
    if (profData.rows.length > 0) {
      var profFolder = _pcnGetOrCreateFolder_(perFolder, 'Proformas');
      _pcnCrearSheetEnDrive_(profFolder, 'Proformas_' + periodo, profData.headers, profData.rows);
    }
  }

  // Presupuestos (sin filtro de fecha)
  var chkPres = sh.getRange(PCN_DOCS.PRESUPUESTOS[0], PCN_DOCS.PRESUPUESTOS[1]).getValue();
  if (chkPres) {
    var presData = _pcnGetAllSheetData_(ss, 'PRESUPUESTOS');
    if (presData.rows.length > 0) {
      var presFolder = _pcnGetOrCreateFolder_(perFolder, 'Presupuestos');
      _pcnCrearSheetEnDrive_(presFolder, 'Presupuestos_' + periodo, presData.headers, presData.rows);
    }
  }

  return { facCount: facData.rows.length, gasCount: gasData.rows.length };
}

// ── Lectores de datos ─────────────────────────────────────────

/** Lee FACTURAS filtradas por rango. Devuelve { headers, rows }. */
function _pcnGetFacturasData_(ss, rango) {
  var src = ss.getSheetByName('FACTURAS');
  if (!src || src.getLastRow() < 1) return { headers: [], rows: [] };
  var lastCol = src.getLastColumn();
  if (lastCol < 1)                   return { headers: [], rows: [] };

  var headers = src.getRange(1, 1, 1, lastCol).getValues()[0];
  if (src.getLastRow() < 2)          return { headers: headers, rows: [] };

  var all  = src.getRange(2, 1, src.getLastRow() - 1, lastCol).getValues();
  var rows = all.filter(function(r) {
    if (!String(r[0]).trim()) return false;
    var estado = lastCol > 17 ? String(r[17]).trim() : '';
    if (estado === 'Cancelada') return false;
    return _pcnEnRango_(r[1], rango);
  });

  return { headers: headers, rows: rows };
}

/** Lee GASTOS filtrados por rango. Devuelve { headers, rows }. */
function _pcnGetGastosData_(ss, rango) {
  var src = ss.getSheetByName('GASTOS');
  if (!src || src.getLastRow() < 1) return { headers: [], rows: [] };
  var lastCol = src.getLastColumn();
  if (lastCol < 1)                   return { headers: [], rows: [] };

  var headers = src.getRange(1, 1, 1, lastCol).getValues()[0];
  if (src.getLastRow() < 2)          return { headers: headers, rows: [] };

  var all  = src.getRange(2, 1, src.getLastRow() - 1, lastCol).getValues();
  var rows = all.filter(function(r) {
    if (!String(r[0]).trim()) return false;
    return _pcnEnRango_(r[3], rango);   // col 4 = Fecha_Gasto
  });

  return { headers: headers, rows: rows };
}

/** Lee todos los registros de una hoja (sin filtro de fecha). */
function _pcnGetAllSheetData_(ss, sheetName) {
  var src = ss.getSheetByName(sheetName);
  if (!src || src.getLastRow() < 1) return { headers: [], rows: [] };
  var lastCol = src.getLastColumn();
  if (lastCol < 1)                   return { headers: [], rows: [] };

  var headers = src.getRange(1, 1, 1, lastCol).getValues()[0];
  if (src.getLastRow() < 2)          return { headers: headers, rows: [] };

  var rows = src.getRange(2, 1, src.getLastRow() - 1, lastCol).getValues()
    .filter(function(r) { return String(r[0]).trim() !== ''; });

  return { headers: headers, rows: rows };
}

// ── Crear Google Sheet en carpeta Drive ───────────────────────

/**
 * Crea un nuevo Google Sheet en la carpeta indicada con los datos dados.
 * Si ya existe un archivo con ese nombre, lo traslada a la papelera.
 * Devuelve la URL del nuevo spreadsheet.
 */
function _pcnCrearSheetEnDrive_(folder, nombre, headers, rows) {
  // Limpiar versión anterior
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

  if (nc > 0) {
    sheet.autoResizeColumns(1, Math.min(nc, 26));
  }

  // Mover a la carpeta destino
  var file = DriveApp.getFileById(newSS.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  return newSS.getUrl();
}

// ============================================================
// FASE 17.5 — Resúmenes contables y CSV
// ============================================================

/**
 * Genera resumen contable, resumen IVA (si está marcado) y CSV.
 */
function _pcnGenerarResumenes_(sh, perFolder, periodo, empresa, kpis, facData, gasData) {
  var resFolder = _pcnGetOrCreateFolder_(perFolder, 'Resumen');
  var ivaFolder = _pcnGetOrCreateFolder_(perFolder, 'IVA');

  // Resumen contable (siempre)
  _pcnCrearResumenContable_(resFolder, 'Resumen_' + periodo, periodo, empresa, kpis);

  // Resumen IVA (si el checkbox RESUMEN_IVA está marcado)
  var chkIva = sh.getRange(PCN_DOCS.RESUMEN_IVA[0], PCN_DOCS.RESUMEN_IVA[1]).getValue();
  if (chkIva !== false) {
    _pcnCrearResumenIVA_(ivaFolder, 'IVA_' + periodo, periodo, empresa, kpis, facData, gasData);
  }

  // CSV facturas
  var chkCsvFac = sh.getRange(PCN_DOCS.CSV_FAC[0], PCN_DOCS.CSV_FAC[1]).getValue();
  if (chkCsvFac) {
    _pcnCrearCSVEnDrive_(resFolder, 'Facturas_' + periodo + '.csv',
      facData.headers, facData.rows);
  }

  // CSV gastos
  var chkCsvGas = sh.getRange(PCN_DOCS.CSV_GAS[0], PCN_DOCS.CSV_GAS[1]).getValue();
  if (chkCsvGas) {
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

  // Fila 1: Título
  sh.getRange(1, 1, 1, 2).merge()
    .setValue('📦 RESUMEN CONTABLE — ' + periodo)
    .setBackground('#4a148c').setFontColor('#ffffff')
    .setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 48);

  // Datos del resumen [label, value]
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

  // startRow=2 → índice i → fila (2 + i)
  info.forEach(function(row, i) {
    sh.getRange(2 + i, 1).setValue(row[0]);
    sh.getRange(2 + i, 2).setValue(row[1]);
  });

  // Cabeceras de sección (índices 4, 10, 15 → filas 6, 12, 17)
  [6, 12, 17].forEach(function(r) {
    sh.getRange(r, 1, 1, 2).merge()
      .setBackground('#e8f0fe').setFontWeight('bold').setFontSize(11);
  });

  // Formato EUR (filas 8,9,10, 14,15, 18,19,20)
  [8, 9, 10, 14, 15, 18, 19, 20].forEach(function(r) {
    var v = sh.getRange(r, 2).getValue();
    if (typeof v === 'number') {
      sh.getRange(r, 2).setNumberFormat('#,##0.00 "€"');
    }
  });

  // Fila 20 = IVA a declarar — resaltar
  var bgIva = kpis.ivaNeto >= 0 ? '#fce8e6' : '#e6f4ea';
  sh.getRange(20, 1, 1, 2)
    .setBackground(bgIva).setFontWeight('bold');

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

  var tz  = Session.getScriptTimeZone();
  [120, 100, 140, 130, 130, 100].forEach(function(w, i) {
    sh.setColumnWidth(i + 1, w);
  });

  // Título
  sh.getRange(1, 1, 1, 6).merge()
    .setValue('🧾 RESUMEN IVA — ' + periodo + ' — ' + empresa)
    .setBackground('#4a148c').setFontColor('#ffffff')
    .setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 40);

  var row = 2;

  // ── Sección FACTURAS ──────────────────────────────────────────
  sh.getRange(row, 1, 1, 6).merge()
    .setValue('  📄 IVA REPERCUTIDO — FACTURAS EMITIDAS')
    .setBackground('#e8f0fe').setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(row, 28);
  row++;

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

  // Total IVA repercutido
  sh.getRange(row, 1, 1, 3).merge()
    .setValue('TOTAL IVA REPERCUTIDO')
    .setFontWeight('bold').setBackground('#e8f0fe');
  sh.getRange(row, 4).setValue(kpis.ivaRep)
    .setNumberFormat('#,##0.00 "€"').setFontWeight('bold').setBackground('#e8f0fe');
  row += 2;

  // ── Sección GASTOS ────────────────────────────────────────────
  sh.getRange(row, 1, 1, 5).merge()
    .setValue('  💸 IVA SOPORTADO — GASTOS / TICKETS')
    .setBackground('#fef7e0').setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(row, 28);
  row++;

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

  // ── Liquidación ───────────────────────────────────────────────
  sh.getRange(row, 1, 1, 3).merge()
    .setValue('🧾 LIQUIDACIÓN IVA')
    .setBackground('#4a148c').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11);
  row++;

  var liq = [
    ['IVA Repercutido (+)', kpis.ivaRep, '#f8f9fa'],
    ['IVA Soportado (−)',   kpis.ivaSop, '#f8f9fa'],
    ['IVA A DECLARAR',      kpis.ivaNeto, kpis.ivaNeto >= 0 ? '#fce8e6' : '#e6f4ea']
  ];
  liq.forEach(function(lr, i) {
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

/** Crea un archivo CSV en la carpeta Drive especificada. */
function _pcnCrearCSVEnDrive_(folder, nombre, headers, rows) {
  var existing = folder.getFilesByName(nombre);
  while (existing.hasNext()) { existing.next().setTrashed(true); }

  var tz    = Session.getScriptTimeZone();
  var lines = [];
  if (headers.length > 0) lines.push(_pcnRowToCSV_(headers, tz));
  rows.forEach(function(r) { lines.push(_pcnRowToCSV_(r, tz)); });

  var blob = Utilities.newBlob(
    '﻿' + lines.join('\r\n'),   // BOM UTF-8 para compatibilidad Excel
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
