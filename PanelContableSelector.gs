// ============================================================
// PanelContableSelector.gs — FASE 18.1
// Selector interactivo de documentos contables por periodo
// ============================================================

// Primera fila dinámica del selector (las 40 del panel fijo quedan intactas)
var PCN_SEL_SEP_ROW = 41;

// Clave PropertiesService con la geometría del selector activo
var PCN_PROP_SEL = 'pcn_selector_layout';

// Columnas del selector (1-based, respetan la rejilla de 14 cols del panel)
var PCN_SEL_COL = {
  CHECK:  1,   // Checkbox Incluir / Excluir
  ID:     2,   // Nº Factura / ID Gasto
  TOTAL:  3,   // Importe €
  // col 4 = espaciador estrecho (15 px)
  FECHA:  5,   // Fecha del documento
  ESTADO: 6,   // Estado (Pagada / Estado_Pago)
  // col 7 = espaciador
  NOMBRE: 8,   // Cliente / Proveedor
  EXTRA:  9,   // vacío (fac) / Concepto (gas)
  // col 10 = espaciador
  ARCH:   11,  // Archivo label (merged 11-12)
  // col 13 = espaciador
  ESTDOC: 14   // ✅ ⚠ ❌
};

// Índices de columna (0-based) del selector — complementan los de PanelContableExport.gs
var PCN_IDX_FAC_FECHA_S      = 1;
var PCN_IDX_FAC_TOTAL_S      = 16;
var PCN_IDX_FAC_ESTADO_S     = 17;
var PCN_IDX_GAS_CONCEPTO_S   = 7;
var PCN_IDX_GAS_TOTAL_S      = 12;
var PCN_IDX_GAS_ESTPAGO_S    = 14;

// ============================================================
// RENDER PRINCIPAL
// ============================================================

/**
 * Borra las filas 41+ y regenera el selector de documentos
 * con checkboxes y estado de archivo para el periodo actual.
 * Guarda la geometría en PropertiesService.
 * Llamar desde: Actualizar, Preparar carpeta.
 */
function _pcnRenderizarSelector_(ss, sh) {
  var tipo    = _pcnLeerTipoPeriodo_(sh);
  var periodo = _pcnPeriodoEtiqueta_(tipo);
  var rango   = _pcnObtenerRango_(sh);
  var tz      = Session.getScriptTimeZone();
  var nc      = PCN_NC;

  var facData = _pcnGetFacturasData_(ss, rango);
  var gasData = _pcnGetGastosData_(ss, rango);

  _pcnLimpiarSelector_(sh);

  var row = PCN_SEL_SEP_ROW;

  // Separador visual
  _pcnSepRow_(sh, row, nc);
  row++;

  // Cabecera de la sección selector
  sh.getRange(row, 1, 1, nc).merge()
    .setValue('  📋  DOCUMENTOS DEL PERIODO — ' + periodo +
      '     ·  ☑ Incluir / ☐ Excluir  →  pulsa "Generar paquete" para aplicar')
    .setBackground(PCN_C.MORADO_DARK).setFontColor(PCN_C.BLANCO)
    .setFontSize(12).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(row, 34);
  row++;

  // ── Sección FACTURAS ──────────────────────────────────────────
  var facSectionRow = row;
  row = _pcnSelRenderFac_(sh, facData, row, tz);
  var facDataStart = facData.rows.length > 0 ? facSectionRow + 2 : -1;

  // Separador entre secciones
  _pcnSepRow_(sh, row, nc);
  row++;

  // ── Sección GASTOS ────────────────────────────────────────────
  var gasSectionRow = row;
  row = _pcnSelRenderGas_(sh, gasData, row, tz);
  var gasDataStart = gasData.rows.length > 0 ? gasSectionRow + 2 : -1;

  // Nota al pie
  sh.getRange(row, 1, 1, nc).merge()
    .setValue(
      '  💡  Se excluyen por defecto: facturas anuladas/borrador · gastos sin importe · ' +
      'gastos sin ticket.  Marca manualmente los que quieras incluir.')
    .setBackground(PCN_C.NARANJA_BG).setFontColor(PCN_C.NARANJA)
    .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(row, 22);

  // Guardar geometría del selector
  PropertiesService.getScriptProperties().setProperty(PCN_PROP_SEL, JSON.stringify({
    periodo:      periodo,
    facDataStart: facDataStart,
    facCount:     facData.rows.length,
    gasDataStart: gasDataStart,
    gasCount:     gasData.rows.length
  }));

  SpreadsheetApp.flush();
  Logger.log('  ✓ Selector: ' + facData.rows.length + ' fac / ' +
    gasData.rows.length + ' gas — ' + periodo);
}

/** Borra las filas del selector (41 en adelante). */
function _pcnLimpiarSelector_(sh) {
  var last = sh.getLastRow();
  if (last < PCN_SEL_SEP_ROW) return;
  var n = last - PCN_SEL_SEP_ROW + 1;
  sh.getRange(PCN_SEL_SEP_ROW, 1, n, PCN_NC).breakApart();
  sh.deleteRows(PCN_SEL_SEP_ROW, n);
}

// ============================================================
// SECCIÓN FACTURAS
// ============================================================

function _pcnSelRenderFac_(sh, facData, startRow, tz) {
  var nc  = PCN_NC;
  var row = startRow;

  sh.getRange(row, 1, 1, nc).merge()
    .setValue('  📄  FACTURAS DEL PERIODO (' + facData.rows.length + ')')
    .setBackground(PCN_C.AZUL_CORP).setFontColor(PCN_C.BLANCO)
    .setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(row, 28);
  row++;

  if (facData.rows.length === 0) {
    sh.getRange(row, 1, 1, nc).merge()
      .setValue('  Sin facturas en el periodo seleccionado.')
      .setBackground(PCN_C.GRIS_BG).setFontColor(PCN_C.GRIS).setFontSize(10)
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.setRowHeight(row, 26);
    return row + 1;
  }

  _pcnSelColHeaders_(sh, row,
    'Nº Factura', 'Total', 'Fecha', 'Estado', 'Cliente', '', 'Archivo PDF');
  row++;

  for (var i = 0; i < facData.rows.length; i++) {
    var r   = facData.rows[i];
    var id  = String(r[0] || '').trim();
    var fv  = r[PCN_IDX_FAC_FECHA_S];
    var fch = (fv instanceof Date && !isNaN(fv))
              ? Utilities.formatDate(fv, tz, 'dd/MM/yyyy') : String(fv || '');
    var cli = String(r.length > PCN_IDX_FAC_NOMBRE_CLI
              ? (r[PCN_IDX_FAC_NOMBRE_CLI] || '') : '').trim();
    var tot = typeof r[PCN_IDX_FAC_TOTAL_S] === 'number' ? r[PCN_IDX_FAC_TOTAL_S] : 0;
    var est = String(r[PCN_IDX_FAC_ESTADO_S] || '').trim();
    var fs  = _pcnFilesStatusFac_(r);
    var chk = _pcnDefaultCheckFac_(r);
    var bg  = (i % 2 === 0) ? PCN_C.BLANCO : PCN_C.AZUL_LIGHT;

    _pcnSelWriteRow_(sh, row, bg, chk, id, tot, fch, est, cli, '', fs);
    row++;
  }

  return row;
}

// ============================================================
// SECCIÓN GASTOS
// ============================================================

function _pcnSelRenderGas_(sh, gasData, startRow, tz) {
  var nc  = PCN_NC;
  var row = startRow;

  sh.getRange(row, 1, 1, nc).merge()
    .setValue('  🧾  GASTOS / TICKETS DEL PERIODO (' + gasData.rows.length + ')')
    .setBackground(PCN_C.VERDE).setFontColor(PCN_C.BLANCO)
    .setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(row, 28);
  row++;

  if (gasData.rows.length === 0) {
    sh.getRange(row, 1, 1, nc).merge()
      .setValue('  Sin gastos en el periodo seleccionado.')
      .setBackground(PCN_C.GRIS_BG).setFontColor(PCN_C.GRIS).setFontSize(10)
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    sh.setRowHeight(row, 26);
    return row + 1;
  }

  _pcnSelColHeaders_(sh, row,
    'ID Gasto', 'Total', 'Fecha', 'Estado Pago', 'Proveedor', 'Concepto', 'Ticket/Foto');
  row++;

  for (var i = 0; i < gasData.rows.length; i++) {
    var r    = gasData.rows[i];
    var id   = String(r[0] || '').trim();
    var fv   = r[3]; // Fecha_Gasto (idx 3)
    var fch  = (fv instanceof Date && !isNaN(fv))
               ? Utilities.formatDate(fv, tz, 'dd/MM/yyyy') : String(fv || '');
    var prov = String(r.length > PCN_IDX_GAS_PROVEEDOR
               ? (r[PCN_IDX_GAS_PROVEEDOR] || '') : '').trim();
    var conc = String(r.length > PCN_IDX_GAS_CONCEPTO_S
               ? (r[PCN_IDX_GAS_CONCEPTO_S] || '') : '').trim();
    var tot  = typeof r[PCN_IDX_GAS_TOTAL_S] === 'number' ? r[PCN_IDX_GAS_TOTAL_S] : 0;
    var est  = String(r.length > PCN_IDX_GAS_ESTPAGO_S
               ? (r[PCN_IDX_GAS_ESTPAGO_S] || '') : '').trim();
    var fs   = _pcnFilesStatusGas_(r);
    var chk  = _pcnDefaultCheckGas_(r);
    var bg   = (i % 2 === 0) ? PCN_C.BLANCO : PCN_C.VERDE_BG;

    _pcnSelWriteRow_(sh, row, bg, chk, id, tot, fch, est, prov, conc, fs);
    row++;
  }

  return row;
}

// ============================================================
// HELPERS DE RENDERIZADO
// ============================================================

function _pcnSelColHeaders_(sh, row, lblId, lblTot, lblFch, lblEst,
                             lblNom, lblExtra, lblArch) {
  var nc = PCN_NC;
  var bg = PCN_C.GRIS_BG;
  var fg = PCN_C.GRIS_DARK;
  var fs = 9;

  sh.getRange(row, 1, 1, nc).setBackground(bg);

  function h(col, txt) {
    sh.getRange(row, col)
      .setValue(txt).setBackground(bg).setFontColor(fg)
      .setFontSize(fs).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  }
  h(PCN_SEL_COL.CHECK,  '☑');
  h(PCN_SEL_COL.ID,     lblId);
  h(PCN_SEL_COL.TOTAL,  lblTot);
  h(PCN_SEL_COL.FECHA,  lblFch);
  h(PCN_SEL_COL.ESTADO, lblEst);
  h(PCN_SEL_COL.NOMBRE, lblNom);
  h(PCN_SEL_COL.EXTRA,  lblExtra);
  h(PCN_SEL_COL.ESTDOC, '📋');

  sh.getRange(row, PCN_SEL_COL.ARCH, 1, 2).merge()
    .setValue(lblArch).setBackground(bg).setFontColor(fg)
    .setFontSize(fs).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sh.setRowHeight(row, 20);
}

function _pcnSelWriteRow_(sh, row, bg, chk, id, tot, fecha, estado,
                           nombre, extra, fs) {
  sh.setRowHeight(row, 22);

  sh.getRange(row, PCN_SEL_COL.CHECK)
    .insertCheckboxes().setValue(chk)
    .setBackground(bg).setHorizontalAlignment('center');

  sh.getRange(row, PCN_SEL_COL.ID)
    .setValue(id)
    .setBackground(bg).setFontSize(10).setFontWeight('bold')
    .setFontColor(PCN_C.MORADO).setHorizontalAlignment('left');

  sh.getRange(row, PCN_SEL_COL.TOTAL)
    .setValue(tot)
    .setBackground(bg).setFontSize(10).setFontColor(PCN_C.NEGRO)
    .setNumberFormat(PCN_EUR_FMT).setHorizontalAlignment('right');

  sh.getRange(row, 4).setBackground(PCN_C.SEP_BG).setValue('');

  sh.getRange(row, PCN_SEL_COL.FECHA)
    .setValue(fecha)
    .setBackground(bg).setFontSize(10).setFontColor(PCN_C.GRIS)
    .setHorizontalAlignment('center');

  sh.getRange(row, PCN_SEL_COL.ESTADO)
    .setValue(estado)
    .setBackground(bg).setFontSize(9).setFontColor(PCN_C.GRIS)
    .setHorizontalAlignment('center');

  sh.getRange(row, 7).setBackground(PCN_C.SEP_BG).setValue('');

  sh.getRange(row, PCN_SEL_COL.NOMBRE)
    .setValue(nombre)
    .setBackground(bg).setFontSize(10).setFontColor(PCN_C.NEGRO)
    .setHorizontalAlignment('left');

  sh.getRange(row, PCN_SEL_COL.EXTRA)
    .setValue(extra)
    .setBackground(bg).setFontSize(9).setFontColor(PCN_C.GRIS)
    .setHorizontalAlignment('left');

  sh.getRange(row, 10).setBackground(PCN_C.SEP_BG).setValue('');

  sh.getRange(row, PCN_SEL_COL.ARCH, 1, 2).merge()
    .setValue(fs.label)
    .setBackground(bg).setFontSize(9)
    .setFontColor(fs.hasFile ? PCN_C.VERDE : PCN_C.NARANJA)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');

  sh.getRange(row, 13).setBackground(PCN_C.SEP_BG).setValue('');

  sh.getRange(row, PCN_SEL_COL.ESTDOC)
    .setValue(fs.emoji)
    .setBackground(bg).setFontSize(12)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
}

// ============================================================
// DETECCIÓN DE ARCHIVOS
// ============================================================

function _pcnFilesStatusFac_(r) {
  var url = String(r.length > PCN_IDX_FAC_URL_PDF
    ? (r[PCN_IDX_FAC_URL_PDF] || '') : '').trim();
  if (!url || url === '—')    return { label: 'Sin PDF',           emoji: '⚠', hasFile: false };
  if (_pcnExtraerFileId_(url)) return { label: 'PDF en Drive',      emoji: '✅', hasFile: true  };
  return                                { label: 'URL sin ID Drive', emoji: '⚠', hasFile: false };
}

function _pcnFilesStatusGas_(r) {
  var archId = String(r.length > PCN_IDX_GAS_ARCHIVO_ID
    ? (r[PCN_IDX_GAS_ARCHIVO_ID] || '') : '').trim();
  if (archId && archId !== '—')
    return { label: 'Archivo en Drive', emoji: '✅', hasFile: true };

  var fUrl = String(r.length > PCN_IDX_GAS_FOTO_URL
    ? (r[PCN_IDX_GAS_FOTO_URL] || '') : '').trim();
  if (fUrl && fUrl !== '—') {
    if (_pcnExtraerFileId_(fUrl)) return { label: 'Foto en Drive',      emoji: '✅', hasFile: true  };
    return                               { label: 'URL sin ID Drive',   emoji: '⚠', hasFile: false };
  }
  return { label: 'Sin ticket/foto', emoji: '⚠', hasFile: false };
}

// ============================================================
// REGLAS DE SELECCIÓN POR DEFECTO
// ============================================================

function _pcnDefaultCheckFac_(r) {
  var est = String(r[PCN_IDX_FAC_ESTADO_S] || '').trim().toUpperCase();
  if (!est || est === 'ANULADA' || est === 'BORRADOR') return false;
  return true;
}

function _pcnDefaultCheckGas_(r) {
  // Pre-marcar solo si tiene archivo; los demás se muestran sin marcar para revisión
  return _pcnFilesStatusGas_(r).hasFile;
}

// ============================================================
// LEER SELECCIÓN ACTIVA
// ============================================================

/**
 * Lee los checkboxes del selector y devuelve los datos filtrados.
 * Si el selector no existe o es de otro periodo, devuelve todos los datos
 * con hasSelector=false (y stale=true si hay desfase de periodo).
 */
function _pcnGetDocumentosSeleccionados_(sh, facData, gasData) {
  var none = {
    facDataFilt: facData,  gasDataFilt: gasData,
    hasSelector: false,    stale:       false,
    facTotal:    facData.rows.length, facSelCount: facData.rows.length,
    gasTotal:    gasData.rows.length, gasSelCount: gasData.rows.length
  };

  var raw = PropertiesService.getScriptProperties().getProperty(PCN_PROP_SEL);
  if (!raw) return none;

  var layout;
  try { layout = JSON.parse(raw); } catch (e) { return none; }

  var tipo    = _pcnLeerTipoPeriodo_(sh);
  var periodo = _pcnPeriodoEtiqueta_(tipo);
  if (layout.periodo !== periodo) {
    return {
      facDataFilt: facData,  gasDataFilt: gasData,
      hasSelector: false,    stale: true,
      facTotal:    facData.rows.length, facSelCount: facData.rows.length,
      gasTotal:    gasData.rows.length, gasSelCount: gasData.rows.length
    };
  }

  // Leer checkboxes + IDs de facturas (cols CHECK y ID juntas)
  var facSel = {};
  if (layout.facCount > 0 && layout.facDataStart > 0) {
    var fv = sh.getRange(layout.facDataStart, PCN_SEL_COL.CHECK,
                         layout.facCount, 2).getValues();
    for (var i = 0; i < fv.length; i++) {
      if (fv[i][0] === true) facSel[String(fv[i][1]).trim()] = true;
    }
  }

  // Leer checkboxes + IDs de gastos
  var gasSel = {};
  if (layout.gasCount > 0 && layout.gasDataStart > 0) {
    var gv = sh.getRange(layout.gasDataStart, PCN_SEL_COL.CHECK,
                         layout.gasCount, 2).getValues();
    for (var j = 0; j < gv.length; j++) {
      if (gv[j][0] === true) gasSel[String(gv[j][1]).trim()] = true;
    }
  }

  var facFilt = facData.rows.filter(function(r) {
    return facSel[String(r[0] || '').trim()] === true;
  });
  var gasFilt = gasData.rows.filter(function(r) {
    return gasSel[String(r[0] || '').trim()] === true;
  });

  return {
    facDataFilt: { headers: facData.headers, rows: facFilt },
    gasDataFilt: { headers: gasData.headers, rows: gasFilt },
    hasSelector: true,  stale: false,
    facTotal:    facData.rows.length, facSelCount: facFilt.length,
    gasTotal:    gasData.rows.length, gasSelCount: gasFilt.length
  };
}
