// ============================================================
// DocTrabajos.gs — Relación explícita entre documentos y trabajos
// ============================================================
//
// Hoja DOC_TRABAJOS — columnas:
//   1: ID_Documento
//   2: Tipo_Documento   (Factura / Proforma)
//   3: ID_Trabajo
//   4: Fecha_Trabajo
//   5: Cliente
//   6: Trabajo_Realizado
//
// Funciones:
//   _ensureDocTrabajosSheet_()           — crea la hoja si no existe
//   registrarDocTrabajos(id, tipo, trabajos, cli) — escribe filas
//   _getDocTrabajosMapByDoc_(clienteLabel)        — lee y devuelve mapa
// ============================================================

/**
 * OBSOLETO — DOC_TRABAJOS ya no se crea.
 * La relación documento→trabajos se guarda en col IDs_Trabajos de FACTURAS/PROFORMAS.
 * Se mantiene para poder leer la hoja si ya existía (migración).
 */
function _ensureDocTrabajosSheet_() {
  return SpreadsheetApp.getActive().getSheetByName('DOC_TRABAJOS') || null;
}

/**
 * OBSOLETO — no-op. Los IDs de trabajos ya se guardan en col 10 (IDs_Trabajos)
 * de FACTURAS y PROFORMAS directamente en generateInvoicePDF().
 */
function registrarDocTrabajos(docId, tipo, trabajos, cliNombre) {
  // no-op: la información ya está en FACTURAS/PROFORMAS col IDs_Trabajos
}

/**
 * Migración única: lee DOC_TRABAJOS (si existe) y rellena la columna IDs_Trabajos
 * en FACTURAS/PROFORMAS para los documentos que la tengan vacía.
 * SOLO LECTURA sobre DOC_TRABAJOS. No borra ni modifica esa hoja.
 * Ejecutar una sola vez desde el editor de Apps Script.
 */
function migrarDocTrabajosAFacturas() {
  var ss = SpreadsheetApp.getActive();
  var tz = Session.getScriptTimeZone();
  Logger.log('================================================');
  Logger.log('migrarDocTrabajosAFacturas()');
  Logger.log(Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var shDT = ss.getSheetByName('DOC_TRABAJOS');
  if (!shDT || shDT.getLastRow() < 2) {
    Logger.log('DOC_TRABAJOS vacía o no existe. Nada que migrar.');
    return;
  }

  // Construir mapa docId → [jobId, jobId, ...]
  var dtData = shDT.getRange(2, 1, shDT.getLastRow() - 1, 3).getValues();
  var docMap = {};
  for (var i = 0; i < dtData.length; i++) {
    var docId = String(dtData[i][0] || '').trim();
    var jobId = String(dtData[i][2] || '').trim();
    if (!docId || !jobId) continue;
    if (!docMap[docId]) docMap[docId] = [];
    if (docMap[docId].indexOf(jobId) < 0) docMap[docId].push(jobId);
  }
  Logger.log('Documentos en DOC_TRABAJOS: ' + Object.keys(docMap).length);

  var _backfill_ = function(shName) {
    var sh = ss.getSheetByName(shName);
    if (!sh || sh.getLastRow() < 2) return;
    var hdrs   = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var colIds = -1;
    for (var h = 0; h < hdrs.length; h++) {
      if (String(hdrs[h]).trim() === 'IDs_Trabajos') { colIds = h + 1; break; }
    }
    if (colIds < 0) { Logger.log('  ' + shName + ': columna IDs_Trabajos no encontrada.'); return; }
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, colIds).getValues();
    var actualizados = 0;
    for (var r = 0; r < data.length; r++) {
      var docId  = String(data[r][0] || '').trim();
      var actual = String(data[r][colIds - 1] || '').trim();
      if (!docId || actual) continue; // ya tiene IDs_Trabajos, no tocar
      if (!docMap[docId]) continue;   // no hay info en DOC_TRABAJOS para este doc
      var valor = docMap[docId].join(', ');
      sh.getRange(r + 2, colIds).setValue(valor);
      Logger.log('  ' + shName + ' ' + docId + ' → IDs_Trabajos="' + valor + '"');
      actualizados++;
    }
    Logger.log('  ' + shName + ': ' + actualizados + ' fila(s) actualizadas.');
  };

  _backfill_('FACTURAS');
  _backfill_('PROFORMAS');

  Logger.log('\n✓ Migración completada.');
  Logger.log('Ahora puedes borrar la hoja DOC_TRABAJOS del spreadsheet.');
  Logger.log('================================================');
}

/**
 * Lee DOC_TRABAJOS completo y devuelve un mapa: docId → [{id, fecha, trabajo}]
 * Filtra por clienteLabel (case-insensitive) para no mezclar datos entre clientes.
 *
 * @param  {string} clienteLabel — etiqueta del dropdown del cliente
 * @returns {Object} mapa: docId → array de {id:string, fecha:string, trabajo:string}
 */
function _getDocTrabajosMapByDoc_(clienteLabel) {
  var sh = SpreadsheetApp.getActive().getSheetByName('DOC_TRABAJOS');
  if (!sh || sh.getLastRow() < 2) return {};
  var data   = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  var map    = {};
  var cliLow = String(clienteLabel || '').toLowerCase().trim();
  for (var i = 0; i < data.length; i++) {
    var row   = data[i];
    var docId = String(row[0] || '').trim();
    if (!docId) continue;
    if (cliLow) {
      var cli = String(row[4] || '').toLowerCase().trim();
      if (cli !== cliLow) continue;
    }
    if (!map[docId]) map[docId] = [];
    map[docId].push({
      id:      String(row[2] || '').trim(),
      fecha:   String(row[3] || '').trim(),
      trabajo: String(row[5] || '').trim()
    });
  }
  return map;
}

// ================================================================
// DIAGNÓSTICO — detecta duplicados y problemas de integridad
// ================================================================

/**
 * Revisa DOC_TRABAJOS, TRABAJOS, FACTURAS y PROFORMAS en busca de:
 *   · Trabajos con ID_Trabajo en más de un documento (duplicados reales)
 *   · Documentos sin ningún trabajo en DOC_TRABAJOS
 *   · Trabajos sin facturar (informativo, NO es un error)
 *
 * Muestra el resultado en Logger y en un alert de UI.
 * Ejecutar manualmente desde el menú o Apps Script editor.
 */
function diagnosticarDocTrabajosDuplicados() {
  var ui  = SpreadsheetApp.getUi();
  var ss  = SpreadsheetApp.getActive();
  var shDT  = ss.getSheetByName('DOC_TRABAJOS');
  var shTra = ss.getSheetByName('TRABAJOS');
  var shFac = ss.getSheetByName('FACTURAS');
  var shPro = ss.getSheetByName('PROFORMAS');

  var lines = ['DIAGNÓSTICO DOC_TRABAJOS\n'];

  // ── 1. Leer DOC_TRABAJOS ─────────────────────────────────────
  var dtData = [];
  if (shDT && shDT.getLastRow() > 1) {
    dtData = shDT.getRange(2, 1, shDT.getLastRow() - 1, 6).getValues();
  }
  if (dtData.length === 0) {
    lines.push('⚠ DOC_TRABAJOS vacío o no existe. Genera al menos una factura/proforma para poblarla.');
    Logger.log(lines.join('\n'));
    ui.alert('Diagnóstico DOC_TRABAJOS', lines.join('\n'), ui.ButtonSet.OK);
    return;
  }

  // mapa trabajoId → [{docId, tipo}]
  var trabajoDocsMap = {};
  // set docId → true (docs que tienen al menos un trabajo)
  var docsConTrabajos = {};

  for (var i = 0; i < dtData.length; i++) {
    var row   = dtData[i];
    var docId = String(row[0] || '').trim();
    var tipo  = String(row[1] || '').trim();
    var tId   = String(row[2] || '').trim();
    if (!docId || !tId) continue;
    if (!trabajoDocsMap[tId]) trabajoDocsMap[tId] = [];
    trabajoDocsMap[tId].push({ docId: docId, tipo: tipo });
    docsConTrabajos[docId] = true;
  }

  // ── 2. Duplicados: mismo trabajo en >1 documento ──────────────
  var duplicados = [];
  for (var tid in trabajoDocsMap) {
    if (trabajoDocsMap[tid].length > 1) {
      var docList = trabajoDocsMap[tid].map(function(x) {
        return x.docId + ' (' + x.tipo + ')';
      }).join(', ');
      duplicados.push('   ' + tid + ' → ' + docList);
    }
  }

  if (duplicados.length === 0) {
    lines.push('✓ Sin trabajos duplicados en DOC_TRABAJOS.\n');
  } else {
    lines.push('❌ TRABAJOS DUPLICADOS EN VARIOS DOCUMENTOS (' + duplicados.length + '):');
    for (var di = 0; di < duplicados.length; di++) lines.push(duplicados[di]);
    lines.push('   → Acción: eliminar manualmente las filas duplicadas de la hoja DOC_TRABAJOS.\n');
  }

  // ── 3. Documentos sin trabajos en DOC_TRABAJOS ────────────────
  var docsHuerfanos = [];
  var _checkSheet_ = function(sh, tipo) {
    if (!sh || sh.getLastRow() < 2) return;
    var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (var r = 0; r < ids.length; r++) {
      var id = String(ids[r][0] || '').trim();
      if (id && !docsConTrabajos[id]) docsHuerfanos.push(id + ' (' + tipo + ')');
    }
  };
  _checkSheet_(shFac, 'Factura');
  _checkSheet_(shPro, 'Proforma');

  if (docsHuerfanos.length === 0) {
    lines.push('✓ Todos los documentos tienen entradas en DOC_TRABAJOS.\n');
  } else {
    lines.push('⚠ DOCUMENTOS SIN TRABAJOS EN DOC_TRABAJOS (' + docsHuerfanos.length + '):');
    for (var hi = 0; hi < docsHuerfanos.length; hi++) lines.push('   ' + docsHuerfanos[hi]);
    lines.push('   → Normal si fueron generados desde presupuesto (no tienen ID_Trabajo).\n');
  }

  // ── 4. Trabajos sin facturar (informativo) ────────────────────
  var sinFacturar = [];
  if (shTra && shTra.getLastRow() > 1) {
    var traIds = shTra.getRange(2, 1, shTra.getLastRow() - 1, 1).getValues();
    for (var t = 0; t < traIds.length; t++) {
      var tId2 = String(traIds[t][0] || '').trim();
      if (tId2 && !trabajoDocsMap[tId2]) sinFacturar.push(tId2);
    }
  }

  if (sinFacturar.length === 0) {
    lines.push('✓ Todos los trabajos están facturados.\n');
  } else {
    lines.push('ℹ TRABAJOS SIN FACTURAR (' + sinFacturar.length + ') — no es un error:');
    var mostrar = Math.min(sinFacturar.length, 15);
    for (var sf = 0; sf < mostrar; sf++) lines.push('   ' + sinFacturar[sf]);
    if (sinFacturar.length > mostrar) lines.push('   ... y ' + (sinFacturar.length - mostrar) + ' más.');
    lines.push('   → Aparecerán como "Sin facturar" en PANEL_CLIENTE.\n');
  }

  lines.push('FIN DIAGNÓSTICO');

  var report = lines.join('\n');
  Logger.log(report);
  ui.alert('Diagnóstico DOC_TRABAJOS', report, ui.ButtonSet.OK);
}
