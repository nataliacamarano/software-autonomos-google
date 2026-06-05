// ============================================================
// DriveScan185.gs — FASE 18.5: Diagnóstico Drive (solo lectura)
// Escanea Mi Unidad, audita IDs almacenados, genera informe.
// NO modifica, mueve, borra ni renombra ningún archivo o carpeta.
// ============================================================

var DS185_SHEET     = 'DRIVE_SCAN_185';
var DS185_MAX_FILES = 200;   // límite por búsqueda para evitar timeout
var DS185_MAX_ROWS  = 150;   // límite de filas auditadas en secciones 2 y 3

// ─── Punto de entrada público ────────────────────────────────

function diagnosticarMigracionDrive185() {
  _ds185_run_();
}

// ─── Runner principal ────────────────────────────────────────

function _ds185_run_() {
  var ss = getTargetSS_();
  var ui = SpreadsheetApp.getUi();

  // Preparar hoja de resultados
  var sh = ss.getSheetByName(DS185_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DS185_SHEET);
  } else {
    sh.clearContents();
    sh.clearFormats();
  }

  var rows = [];
  var now  = new Date();

  _ds185_hdr_(rows, 'FASE 18.5 — DIAGNÓSTICO DE MIGRACIÓN DRIVE (solo lectura)', now);

  // Precarga de hojas y sets de referencia
  var gasSheet = ss.getSheetByName('GASTOS');
  var facSheet = ss.getSheetByName('FACTURAS');
  var gasIdSet = _ds185_buildIdSet_(gasSheet, 16);      // col 17 = ARCHIVO_ID
  var facIdSet = _ds185_buildFacUrlSet_(facSheet, 18);  // col 19 = URL_PDF

  // Secciones del informe
  _ds185_sec1_ids_(ss, rows);
  _ds185_sec2_gastos_(ss, gasSheet, gasIdSet, rows);
  _ds185_sec3_facturas_(ss, facSheet, facIdSet, rows);
  _ds185_sec4_carpetas_(ss, gasIdSet, facIdSet, rows);
  _ds185_sec5_huerfanos_(ss, gasIdSet, facIdSet, rows);
  _ds185_sec6_legacy_(ss, rows);

  // Escribir a la hoja (9 columnas)
  var NCOLS = 9;
  var data = rows.map(function(r) {
    var out = [];
    for (var i = 0; i < NCOLS; i++) out.push(r[i] !== undefined ? r[i] : '');
    return out;
  });
  if (data.length > 0) {
    sh.getRange(1, 1, data.length, NCOLS).setValues(data);
  }

  // Formato básico
  sh.setFrozenRows(1);
  try { sh.autoResizeColumns(1, NCOLS); } catch(e) {}

  var msg = '✅ Diagnóstico completado\n\nHoja: ' + DS185_SHEET +
            '\n' + data.length + ' filas generadas.\n\nNO se ha modificado nada.';
  ui.alert(msg);
  logEvent('INFO', 'diagnosticarMigracionDrive185', 'Diagnóstico completado', data.length + ' filas');
}

// ─────────────────────────────────────────────────────────────
// SECCIÓN 1 — IDs almacenados en configuración
// ─────────────────────────────────────────────────────────────

function _ds185_sec1_ids_(ss, rows) {
  _ds185_sep_(rows, '1. IDs ALMACENADOS EN CONFIGURACIÓN (CONFIG + DRIVE_CONFIG)');
  rows.push(['Fuente', 'Clave', 'Valor ID', 'Tipo', 'Accesible', 'Nombre en Drive', 'Carpeta padre', 'Observaciones', '']);

  // CONFIG (hoja clave-valor usada por módulos legacy)
  var configKeys = [
    'carpeta_facturas_id',
    'carpeta_trabajos_id',
    'gastos_folder_id'
  ];
  configKeys.forEach(function(key) {
    _ds185_auditId_(rows, 'CONFIG', key, String(getConfig(key) || ''));
  });

  // DRIVE_CONFIG (hoja gestionada por DriveManager.gs)
  var dmKeys = [
    'ROOT_FOLDER_ID',
    'SISTEMA_FOLDER_ID',
    'CLIENTES_FOLDER_ID',
    'FACTURAS_FOLDER_ID',
    'GASTOS_FOLDER_ID',
    'CONTABILIDAD_FOLDER_ID'
  ];
  dmKeys.forEach(function(key) {
    var val = _dmGetConfig_(ss, key);
    _ds185_auditId_(rows, 'DRIVE_CONFIG', key, String(val || '').trim());
  });

  // Detectar solapamientos (mismo ID en ambas fuentes)
  var configVals = configKeys.map(function(k) { return String(getConfig(k) || '').trim(); }).filter(Boolean);
  var dmVals     = dmKeys.map(function(k) { return String(_dmGetConfig_(ss, k) || '').trim(); }).filter(Boolean);
  var overlap    = configVals.filter(function(v) { return dmVals.indexOf(v) >= 0; });
  if (overlap.length > 0) {
    rows.push(['⚠ SOLAPAMIENTO', 'Mismos IDs en CONFIG y DRIVE_CONFIG:', overlap.join(', '), '', '', '', '', 'Dos claves apuntan al mismo objeto Drive', '']);
  }

  rows.push(['']);
}

// ─────────────────────────────────────────────────────────────
// SECCIÓN 2 — GASTOS.ARCHIVO_ID
// ─────────────────────────────────────────────────────────────

function _ds185_sec2_gastos_(ss, gasSheet, gasIdSet, rows) {
  _ds185_sep_(rows, '2. TICKETS DE GASTOS — hoja GASTOS, columna ARCHIVO_ID (col 17)');
  rows.push(['Fila', 'ID_Gasto', 'Descripción', 'ARCHIVO_ID', 'Accesible', 'Nombre archivo', 'Carpeta actual', 'Carpeta destino', 'Acción']);

  if (!gasSheet) {
    rows.push(['—', '—', 'Hoja GASTOS no encontrada', '', '', '', '', '', '']);
    rows.push(['']);
    return;
  }

  var gasTargetId = String(_dmGetConfig_(ss, 'GASTOS_FOLDER_ID') || '').trim();
  var gasTargetNm = gasTargetId ? _ds185_folderName_(gasTargetId) : '(sin configurar)';
  var data = gasSheet.getDataRange().getValues();

  var ok = 0, roto = 0, vacio = 0, yaDentro = 0, fuera = 0;
  var auditados = 0;

  for (var i = 1; i < data.length && auditados < DS185_MAX_ROWS; i++) {
    var row       = data[i];
    var archivoId = String(row[16] || '').trim();
    var idGasto   = String(row[0]  || '').trim();
    var descrip   = String(row[3]  || '').trim().substring(0, 55);
    if (!idGasto && !archivoId) continue;
    if (!archivoId) { vacio++; continue; }

    auditados++;
    var info = _ds185_fileInfo_(archivoId);
    if (!info.ok) {
      rows.push([i + 1, idGasto, descrip, archivoId, '❌ NO ACCESIBLE', '—', '—', '—', 'Verificar manualmente']);
      roto++;
      continue;
    }

    var enDestino = gasTargetId && info.parentId === gasTargetId;
    var accion    = enDestino ? '✅ Ya en destino'
                  : gasTargetId ? '⚠ Mover a ' + gasTargetNm
                  : '— Sin destino configurado';
    if (enDestino) yaDentro++; else fuera++;

    rows.push([i + 1, idGasto, descrip, archivoId, '✅ SÍ', info.name, info.parentName, gasTargetNm, accion]);
    ok++;
  }

  if (auditados >= DS185_MAX_ROWS) {
    rows.push(['⚠ Límite ' + DS185_MAX_ROWS + ' filas auditadas — hay más registros en GASTOS', '', '', '', '', '', '', '', '']);
  }

  rows.push(['', '', 'RESUMEN GASTOS:', '',
    ok + ' accesibles', yaDentro + ' ya en destino', fuera + ' fuera de destino',
    roto + ' no accesibles', vacio + ' sin ARCHIVO_ID']);
  rows.push(['']);
}

// ─────────────────────────────────────────────────────────────
// SECCIÓN 3 — FACTURAS.URL_PDF
// ─────────────────────────────────────────────────────────────

function _ds185_sec3_facturas_(ss, facSheet, facIdSet, rows) {
  _ds185_sep_(rows, '3. PDFs DE FACTURAS — hoja FACTURAS, columna URL_PDF (col 19)');
  rows.push(['Fila', 'ID_Factura', 'Cliente', 'URL_PDF (extracto)', 'Accesible', 'Nombre archivo', 'Carpeta actual', 'Carpeta destino', 'Acción']);

  if (!facSheet) {
    rows.push(['—', '—', 'Hoja FACTURAS no encontrada', '', '', '', '', '', '']);
    rows.push(['']);
    return;
  }

  var facTargetId = String(_dmGetConfig_(ss, 'FACTURAS_FOLDER_ID') || '').trim();
  var facTargetNm = facTargetId ? _ds185_folderName_(facTargetId) : '(sin configurar)';
  var data = facSheet.getDataRange().getValues();

  var ok = 0, roto = 0, vacio = 0, yaDentro = 0, fuera = 0;
  var auditados = 0;

  for (var i = 1; i < data.length && auditados < DS185_MAX_ROWS; i++) {
    var row    = data[i];
    var url    = String(row[18] || '').trim();  // col 19
    var idFac  = String(row[0]  || '').trim();
    var cliente = String(row[2] || '').trim().substring(0, 35);
    if (!idFac && !url) continue;
    if (!url) { vacio++; continue; }

    auditados++;
    var fileId = _ds185_extractId_(url);
    if (!fileId) {
      rows.push([i + 1, idFac, cliente, url.substring(0, 60), '⚠ Sin ID extraíble', '—', '—', facTargetNm, 'Revisar URL']);
      continue;
    }

    var info = _ds185_fileInfo_(fileId);
    if (!info.ok) {
      rows.push([i + 1, idFac, cliente, url.substring(0, 60), '❌ NO ACCESIBLE', '—', '—', facTargetNm, 'PDF eliminado o sin acceso']);
      roto++;
      continue;
    }

    var enDestino = facTargetId && info.parentId === facTargetId;
    var accion    = enDestino ? '✅ Ya en destino'
                  : facTargetId ? '⚠ Mover a ' + facTargetNm
                  : '— Sin destino configurado';
    if (enDestino) yaDentro++; else fuera++;

    rows.push([i + 1, idFac, cliente, url.substring(0, 60), '✅ SÍ', info.name, info.parentName, facTargetNm, accion]);
    ok++;
  }

  if (auditados >= DS185_MAX_ROWS) {
    rows.push(['⚠ Límite ' + DS185_MAX_ROWS + ' filas auditadas — hay más registros en FACTURAS', '', '', '', '', '', '', '', '']);
  }

  rows.push(['', '', 'RESUMEN FACTURAS:', '',
    ok + ' accesibles', yaDentro + ' ya en destino', fuera + ' fuera de destino',
    roto + ' no accesibles', vacio + ' sin URL']);
  rows.push(['']);
}

// ─────────────────────────────────────────────────────────────
// SECCIÓN 4 — Contenido de carpetas conocidas
// ─────────────────────────────────────────────────────────────

function _ds185_sec4_carpetas_(ss, gasIdSet, facIdSet, rows) {
  _ds185_sep_(rows, '4. CONTENIDO DE CARPETAS CONOCIDAS (máx ' + DS185_MAX_FILES + ' archivos por carpeta)');
  rows.push(['Carpeta (fuente)', 'Nombre archivo', 'ID archivo', 'Tipo', 'Tamaño KB', 'Modificado', 'Referenciado en datos', 'Duplicado en otra carpeta', '']);

  var foldersToScan = [
    { label: 'CONFIG:gastos_folder_id',             id: getConfig('gastos_folder_id') },
    { label: 'CONFIG:carpeta_facturas_id',          id: getConfig('carpeta_facturas_id') },
    { label: 'CONFIG:carpeta_trabajos_id',          id: getConfig('carpeta_trabajos_id') },
    { label: 'DRIVE_CONFIG:GASTOS_FOLDER_ID',       id: _dmGetConfig_(ss, 'GASTOS_FOLDER_ID') },
    { label: 'DRIVE_CONFIG:FACTURAS_FOLDER_ID',     id: _dmGetConfig_(ss, 'FACTURAS_FOLDER_ID') },
    { label: 'DRIVE_CONFIG:CONTABILIDAD_FOLDER_ID', id: _dmGetConfig_(ss, 'CONTABILIDAD_FOLDER_ID') }
  ];

  var seenFolderIds = {};  // folder dedup
  var allFileIds    = {};  // para detectar duplicados entre carpetas

  foldersToScan.forEach(function(f) {
    var fid = String(f.id || '').trim();
    if (!fid) {
      rows.push([f.label, '(no configurado)', '', '', '', '', '', '', '']);
      return;
    }
    if (seenFolderIds[fid]) {
      rows.push([f.label, '↕ Misma carpeta que ' + seenFolderIds[fid], fid, '', '', '', '', '', '']);
      return;
    }
    seenFolderIds[fid] = f.label;

    try {
      var folder = DriveApp.getFolderById(fid);
      var files  = folder.getFiles();
      var count  = 0;

      while (files.hasNext() && count < DS185_MAX_FILES) {
        var file   = files.next();
        var fileId = file.getId();
        var refSt  = gasIdSet[fileId] ? '✅ GASTOS'
                   : facIdSet[fileId] ? '✅ FACTURAS'
                   : '⚠ No referenciado en datos';
        var dupNote = allFileIds[fileId] ? '⚠ Ya listado en ' + allFileIds[fileId] : '';
        allFileIds[fileId] = f.label;

        rows.push([
          f.label,
          file.getName(),
          fileId,
          _ds185_mimeShort_(file.getMimeType()),
          Math.round(file.getSize() / 1024),
          Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          refSt,
          dupNote,
          ''
        ]);
        count++;
      }

      if (count === 0) {
        rows.push([f.label, '(carpeta vacía — ' + folder.getName() + ')', fid, '', '', '', '', '', '']);
      } else if (count >= DS185_MAX_FILES) {
        rows.push([f.label, '⚠ Límite ' + DS185_MAX_FILES + ' alcanzado — hay más archivos', '', '', '', '', '', '', '']);
      }
    } catch(e) {
      rows.push([f.label, '❌ Carpeta no accesible: ' + e.message, fid, '', '', '', '', '', '']);
    }
  });

  rows.push(['']);
}

// ─────────────────────────────────────────────────────────────
// SECCIÓN 5 — Archivos fuera de carpetas conocidas (huérfanos)
// ─────────────────────────────────────────────────────────────

function _ds185_sec5_huerfanos_(ss, gasIdSet, facIdSet, rows) {
  _ds185_sep_(rows, '5. BÚSQUEDA EN MI UNIDAD — PDFs y CSVs fuera de carpetas conocidas (máx ' + DS185_MAX_FILES + ' por tipo)');
  rows.push(['Nombre', 'ID', 'Tipo', 'Carpeta actual', 'Modificado', 'Referenciado en datos', 'Carpeta destino sugerida', 'Acción', '']);

  var knownIds = _ds185_knownFolderIds_(ss);
  var allRefIds = {};
  Object.keys(gasIdSet).forEach(function(k) { allRefIds[k] = 'GASTOS'; });
  Object.keys(facIdSet).forEach(function(k) { allRefIds[k] = 'FACTURAS'; });

  var facTargetNm = _ds185_folderName_(_dmGetConfig_(ss, 'FACTURAS_FOLDER_ID'));
  var gasTargetNm = _ds185_folderName_(_dmGetConfig_(ss, 'GASTOS_FOLDER_ID'));
  var cntTargetNm = _ds185_folderName_(_dmGetConfig_(ss, 'CONTABILIDAD_FOLDER_ID'));

  var queries = [
    { label: 'PDF',             q: "mimeType='application/pdf' and trashed=false" },
    { label: 'CSV',             q: "mimeType='text/csv' and trashed=false" },
    { label: 'Nombre: Factura', q: "title contains 'Factura' and not mimeType='application/vnd.google-apps.folder' and trashed=false" },
    { label: 'Nombre: Gasto/Ticket', q: "(title contains 'Gasto' or title contains 'Ticket' or title contains 'ticket') and not mimeType='application/vnd.google-apps.folder' and trashed=false" }
  ];

  var seenIds = {};

  queries.forEach(function(qDef) {
    rows.push(['--- ' + qDef.label + ' ---', '', '', '', '', '', '', '', '']);
    try {
      var iter  = DriveApp.searchFiles(qDef.q);
      var count = 0;
      var outOfKnown = 0;

      while (iter.hasNext() && count < DS185_MAX_FILES) {
        var file   = iter.next();
        var fileId = file.getId();
        count++;
        if (seenIds[fileId]) continue;
        seenIds[fileId] = true;

        var parentName = '', parentId = '';
        try {
          if (file.getParents().hasNext()) {
            var p  = file.getParents().next();
            parentName = p.getName();
            parentId   = p.getId();
          }
        } catch(e) {}

        var inKnown  = knownIds.indexOf(parentId) >= 0;
        var refSt    = allRefIds[fileId] ? '✅ ' + allRefIds[fileId] : '⚠ No referenciado';
        var destino  = '';
        var accion   = '';

        if (inKnown) {
          destino = parentName;
          accion  = '✅ Ya en carpeta conocida';
        } else {
          outOfKnown++;
          var mime = file.getMimeType();
          if (mime === 'application/pdf' && file.getName().match(/[Ff]actura|FAC-|FRA-/)) {
            destino = facTargetNm || '?';
            accion  = '⚠ Candidato → FACTURAS_FOLDER';
          } else if (mime === 'application/pdf' && file.getName().match(/[Gg]asto|[Tt]icket/)) {
            destino = gasTargetNm || '?';
            accion  = '⚠ Candidato → GASTOS_FOLDER';
          } else if (mime === 'application/pdf') {
            destino = '?';
            accion  = '⚠ Revisar manualmente';
          } else if (mime === 'text/csv') {
            destino = cntTargetNm || '?';
            accion  = '⚠ Candidato → CONTABILIDAD_FOLDER';
          } else {
            destino = '?';
            accion  = '⚠ Revisar manualmente';
          }
        }

        rows.push([
          file.getName(), fileId, qDef.label, parentName,
          Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          refSt, destino, accion, ''
        ]);
      }

      if (count === 0) rows.push(['(sin resultados para este tipo)', '', '', '', '', '', '', '', '']);
      if (count >= DS185_MAX_FILES) rows.push(['⚠ Límite ' + DS185_MAX_FILES + ' alcanzado — puede haber más', '', '', '', '', '', '', '', '']);
    } catch(e) {
      rows.push(['❌ Error en búsqueda: ' + e.message, '', '', '', '', '', '', '', '']);
    }
  });

  rows.push(['']);
}

// ─────────────────────────────────────────────────────────────
// SECCIÓN 6 — Carpetas legacy
// ─────────────────────────────────────────────────────────────

function _ds185_sec6_legacy_(ss, rows) {
  _ds185_sep_(rows, '6. CARPETAS LEGACY EN MI UNIDAD (búsqueda por nombre)');
  rows.push(['Nombre', 'ID', 'Carpeta padre', 'Archivos dentro', 'Modificado', 'Estado', 'Acción sugerida', '', '']);

  var knownIds = _ds185_knownFolderIds_(ss);

  // Patrones que indican carpetas del sistema antiguo o nombrado diferente al estándar ALVARO_EMPRESA
  var legacyQueries = [
    "title contains 'Álvaro Gestión' and mimeType='application/vnd.google-apps.folder' and trashed=false",
    "title contains 'Alvaro Gestion' and mimeType='application/vnd.google-apps.folder' and trashed=false",
    "title contains 'Tickets y Facturas' and mimeType='application/vnd.google-apps.folder' and trashed=false",
    "title contains 'Paquete_Contable' and mimeType='application/vnd.google-apps.folder' and trashed=false",
    "title contains 'CONTABILIDAD_' and mimeType='application/vnd.google-apps.folder' and trashed=false",
    "(title = 'FACTURAS' or title = 'GASTOS' or title = 'TRABAJOS') and mimeType='application/vnd.google-apps.folder' and trashed=false"
  ];

  var seen = {};

  legacyQueries.forEach(function(q) {
    try {
      var iter = DriveApp.searchFiles(q);
      while (iter.hasNext()) {
        var f   = iter.next();
        var fid = f.getId();
        if (seen[fid]) return;
        seen[fid] = true;

        var parentName = '';
        try { if (f.getParents().hasNext()) parentName = f.getParents().next().getName(); } catch(e) {}

        var fileCount = 0;
        try { var fi = f.getFiles(); while (fi.hasNext()) { fi.next(); fileCount++; if (fileCount > 999) break; } } catch(e) {}

        var isKnown = knownIds.indexOf(fid) >= 0;
        var estado  = isKnown ? '🔗 En uso (referenciado en CONFIG/DRIVE_CONFIG)'
                              : '❓ No referenciado en ninguna config';
        var accion  = isKnown ? '✅ Conservar sin mover'
                              : (fileCount > 0 ? '⚠ Tiene ' + fileCount + ' archivos — verificar antes de archivar'
                                               : '⚠ Vacía — candidata a eliminar tras confirmación');

        rows.push([
          f.getName(), fid, parentName, fileCount,
          Utilities.formatDate(f.getLastUpdated(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          estado, accion, '', ''
        ]);
      }
    } catch(e) {
      rows.push(['❌ Error en búsqueda legacy: ' + e.message, '', '', '', '', '', '', '', '']);
    }
  });

  if (Object.keys(seen).length === 0) {
    rows.push(['(ninguna carpeta legacy detectada con los patrones buscados)', '', '', '', '', '', '', '', '']);
  }

  rows.push(['']);
}

// ─────────────────────────────────────────────────────────────
// Helpers internos
// ─────────────────────────────────────────────────────────────

function _ds185_hdr_(rows, title, now) {
  rows.push([title, '', '', '', '', '', '', '', '']);
  rows.push(['Generado: ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'), '', '', '', '', '', '', '', '']);
  rows.push(['DIAGNÓSTICO DE SOLO LECTURA — No se ha modificado ningún archivo ni carpeta.', '', '', '', '', '', '', '', '']);
  rows.push(['']);
}

function _ds185_sep_(rows, title) {
  rows.push(['=== ' + title + ' ===', '', '', '', '', '', '', '', '']);
}

/** Audita un ID de carpeta/archivo y añade fila al informe. */
function _ds185_auditId_(rows, fuente, key, val) {
  if (!val) {
    rows.push([fuente, key, '(no configurado)', 'carpeta', '—', '—', '—', 'No existe en CONFIG', '']);
    return;
  }
  // Intentar como carpeta primero, luego como archivo
  try {
    var folder = DriveApp.getFolderById(val);
    var pNm = '';
    try { if (folder.getParents().hasNext()) pNm = folder.getParents().next().getName(); } catch(e) {}
    rows.push([fuente, key, val, 'carpeta', '✅ SÍ', folder.getName(), pNm, '', '']);
  } catch(e) {
    try {
      var file = DriveApp.getFileById(val);
      var pNm = '';
      try { if (file.getParents().hasNext()) pNm = file.getParents().next().getName(); } catch(e2) {}
      rows.push([fuente, key, val, 'archivo', '✅ SÍ', file.getName(), pNm, '', '']);
    } catch(e2) {
      rows.push([fuente, key, val, 'carpeta', '❌ NO ACCESIBLE', '—', '—', e.message.substring(0, 80), '']);
    }
  }
}

/** Obtiene nombre e ID del padre de un archivo. Devuelve {ok, name, parentName, parentId}. */
function _ds185_fileInfo_(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var pNm = '', pId = '';
    try {
      if (file.getParents().hasNext()) {
        var p = file.getParents().next();
        pNm = p.getName();
        pId = p.getId();
      }
    } catch(e) {}
    return { ok: true, name: file.getName(), parentName: pNm, parentId: pId };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

/** Extrae el ID de Drive de una URL. Devuelve null si no puede. */
function _ds185_extractId_(url) {
  if (!url) return null;
  var m;
  m = url.match(/\/d\/([a-zA-Z0-9_-]{20,})/);       if (m) return m[1];
  m = url.match(/folders\/([a-zA-Z0-9_-]{20,})/);   if (m) return m[1];
  m = url.match(/[?&]id=([a-zA-Z0-9_-]{20,})/);     if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{20,}$/.test(url.trim())) return url.trim();
  return null;
}

/** Construye un set de IDs de archivo desde una columna de la hoja. */
function _ds185_buildIdSet_(sheet, colIdx) {
  var set = {};
  if (!sheet) return set;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][colIdx] || '').trim();
    if (id) set[id] = true;
  }
  return set;
}

/** Construye un set de file IDs extraídos de URLs en una columna de la hoja. */
function _ds185_buildFacUrlSet_(sheet, colIdx) {
  var set = {};
  if (!sheet) return set;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var id = _ds185_extractId_(String(data[i][colIdx] || '').trim());
    if (id) set[id] = true;
  }
  return set;
}

/** Array de todos los folder IDs conocidos (CONFIG + DRIVE_CONFIG). */
function _ds185_knownFolderIds_(ss) {
  var ids = [];
  var push = function(v) { var s = String(v || '').trim(); if (s) ids.push(s); };
  push(getConfig('gastos_folder_id'));
  push(getConfig('carpeta_facturas_id'));
  push(getConfig('carpeta_trabajos_id'));
  push(_dmGetConfig_(ss, 'ROOT_FOLDER_ID'));
  push(_dmGetConfig_(ss, 'SISTEMA_FOLDER_ID'));
  push(_dmGetConfig_(ss, 'CLIENTES_FOLDER_ID'));
  push(_dmGetConfig_(ss, 'FACTURAS_FOLDER_ID'));
  push(_dmGetConfig_(ss, 'GASTOS_FOLDER_ID'));
  push(_dmGetConfig_(ss, 'CONTABILIDAD_FOLDER_ID'));
  return ids;
}

/** Obtiene el nombre de una carpeta por ID o '' si no accesible. */
function _ds185_folderName_(folderId) {
  if (!folderId) return '';
  try { return DriveApp.getFolderById(String(folderId).trim()).getName(); } catch(e) { return '?'; }
}

/** Abrevia el MIME type a una etiqueta corta. */
function _ds185_mimeShort_(mime) {
  var map = {
    'application/pdf':                          'PDF',
    'text/csv':                                 'CSV',
    'image/jpeg':                               'JPG',
    'image/png':                                'PNG',
    'image/heic':                               'HEIC',
    'application/vnd.google-apps.document':     'GDoc',
    'application/vnd.google-apps.spreadsheet':  'GSheet',
    'application/vnd.google-apps.folder':       'Carpeta'
  };
  return map[mime] || mime;
}
