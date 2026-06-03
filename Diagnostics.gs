// ============================================================
// Diagnostics.gs — Diagnóstico de solo lectura
// NINGUNA función modifica, crea, archiva ni borra datos.
// Sin const a nivel de archivo para máxima compatibilidad GAS V8.
// ============================================================

/**
 * Lista todas las hojas con nombre, GID, cabeceras y filas.
 * Clasifica por cabeceras marcando explícitamente que es inferencia.
 * Al final da instrucciones para la prueba manual definitiva.
 * SOLO LECTURA — no modifica nada.
 */
function generatePreciseResponseSheetReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  // Hojas del sistema — siempre mantener
  var SYSTEM = ['CONFIG','CLIENTES','OPERARIOS','TRABAJOS','FACTURAS',
                'PENDIENTES','COMPLETADOS','LOGS'];

  // Protegidas por decisión explícita del usuario
  var PROTECTED = ['ARCHIVO_Trabajo de formulario 1', 'ARCHIVO_Facturacion'];

  Logger.log('================================================');
  Logger.log('INFORME PRECISO DE HOJAS - SOLO LECTURA');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('Total hojas: ' + sheets.length);
  Logger.log('================================================');

  // ── A) Sistema ────────────────────────────────────────────
  Logger.log('\nA) MANTENER SIEMPRE (sistema)');
  Logger.log('-----------------------------------------------');
  for (var i = 0; i < sheets.length; i++) {
    var s = sheets[i];
    if (SYSTEM.indexOf(s.getName()) !== -1) {
      Logger.log('  OK  ' + s.getName() +
                 '  GID=' + s.getSheetId() +
                 '  filas=' + Math.max(0, s.getLastRow() - 1));
    }
  }

  // ── B) Protegidas ─────────────────────────────────────────
  Logger.log('\nB) NO BORRAR (decision explicita usuario)');
  Logger.log('-----------------------------------------------');
  for (var i = 0; i < sheets.length; i++) {
    var s = sheets[i];
    if (PROTECTED.indexOf(s.getName()) !== -1) {
      Logger.log('  PROT  ' + s.getName() +
                 '  GID=' + s.getSheetId() +
                 '  filas=' + Math.max(0, s.getLastRow() - 1));
    }
  }

  // ── C) Resto de hojas ─────────────────────────────────────
  Logger.log('\nC) RESTO DE HOJAS (clasificacion por cabeceras)');
  Logger.log('  AVISO: "Posible" = inferencia, NO confirmado.');
  Logger.log('-----------------------------------------------');

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var name = sheet.getName();

    if (SYSTEM.indexOf(name) !== -1) continue;
    if (PROTECTED.indexOf(name) !== -1) continue;

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var dataRows = Math.max(0, lastRow - 1);
    var headers = [];

    if (lastRow > 0 && lastCol > 0) {
      var raw = sheet.getRange(1, 1, 1, Math.min(lastCol, 15)).getValues()[0];
      for (var j = 0; j < raw.length; j++) {
        var h = String(raw[j]).trim();
        if (h !== '') headers.push(h);
      }
    }

    Logger.log('\n  Hoja: "' + name + '"');
    Logger.log('    GID:      ' + sheet.getSheetId());
    Logger.log('    Filas:    ' + dataRows);
    Logger.log('    Cabeceras: ' + (headers.length > 0 ? headers.join(' | ') : '(vacia)'));
    Logger.log('    Tipo:     ' + inferirTipoHoja_(name, headers));
  }

  // ── D) Formularios ────────────────────────────────────────
  Logger.log('\nD) FORMULARIOS EN CONFIG');
  Logger.log('-----------------------------------------------');

  var formDefs = [
    ['form_clientes_id',  'form_clientes_url',  'Clientes' ],
    ['form_operarios_id', 'form_operarios_url', 'Operarios']
  ];

  for (var k = 0; k < formDefs.length; k++) {
    var idKey  = formDefs[k][0];
    var urlKey = formDefs[k][1];
    var label  = formDefs[k][2];
    var formId = getConfig(idKey);

    Logger.log('\n  Formulario ' + label + ':');

    if (!formId) {
      Logger.log('    NO configurado en CONFIG (' + idKey + ')');
      continue;
    }

    Logger.log('    ID: ' + formId);
    Logger.log('    URL guardada: ' + (getConfig(urlKey) || '(no guardada)'));

    try {
      var form   = FormApp.openById(formId);
      var destId = form.getDestinationId();
      Logger.log('    Titulo: ' + form.getTitle());

      if (!destId) {
        Logger.log('    Destino: SIN DESTINO (desvinculado)');
      } else if (destId === ss.getId()) {
        Logger.log('    Destino: este spreadsheet OK');
      } else {
        Logger.log('    Destino: OTRO spreadsheet = ' + destId);
      }

      var items = form.getItems();
      var preguntas = [];
      for (var p = 0; p < items.length; p++) {
        preguntas.push(items[p].getTitle());
      }
      Logger.log('    Preguntas: ' + preguntas.join(' | '));

    } catch (e) {
      Logger.log('    ERROR al abrir formulario: ' + e.message);
    }
  }

  // ── E) Instrucciones prueba manual ───────────────────────
  Logger.log('\n================================================');
  Logger.log('E) PRUEBA MANUAL - ejecutar despues');
  Logger.log('================================================');
  Logger.log('Paso 1. Abre el formulario de CLIENTES (URL en D).');
  Logger.log('        Nombre: CLIENTE_TEST_LIMPIEZA');
  Logger.log('        Apellidos: TEST  Email: test@test.com');
  Logger.log('        Telefono: 000000000  Direccion: TEST');
  Logger.log('        Envia el formulario.');
  Logger.log('');
  Logger.log('Paso 2. Abre el formulario de OPERARIOS (URL en D).');
  Logger.log('        Nombre: OPERARIO_TEST_LIMPIEZA');
  Logger.log('        Apellidos: TEST  Email: test@test.com');
  Logger.log('        Telefono: 000000000  Especialidad: Otro');
  Logger.log('        Envia el formulario.');
  Logger.log('');
  Logger.log('Paso 3. Espera 15 segundos.');
  Logger.log('Paso 4. Ejecuta findTestResponses()');
  Logger.log('================================================');
  Logger.log('NINGUNA MODIFICACION REALIZADA.');
  Logger.log('================================================');
}

// ─────────────────────────────────────────────────────────────

/**
 * Busca CLIENTE_TEST_LIMPIEZA y OPERARIO_TEST_LIMPIEZA en todas las hojas.
 * Revela exactamente en que hoja (nombre + GID) y fila cayo cada respuesta.
 * SOLO LECTURA — no modifica nada.
 */
function findTestResponses() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var TEST_C = 'CLIENTE_TEST_LIMPIEZA';
  var TEST_O = 'OPERARIO_TEST_LIMPIEZA';

  Logger.log('================================================');
  Logger.log('BUSQUEDA DE RESPUESTAS DE PRUEBA - SOLO LECTURA');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('Buscando: "' + TEST_C + '"');
  Logger.log('          "' + TEST_O + '"');
  Logger.log('Hojas analizadas: ' + sheets.length);
  Logger.log('================================================');

  var hitC = [];
  var hitO = [];

  for (var i = 0; i < sheets.length; i++) {
    var sheet   = sheets[i];
    var name    = sheet.getName();
    var gid     = sheet.getSheetId();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow < 1 || lastCol < 1) continue;

    var data = sheet.getRange(1, 1, lastRow, Math.min(lastCol, 20)).getValues();

    for (var r = 0; r < data.length; r++) {
      var rowStr = data[r].join('|').toLowerCase();

      if (rowStr.indexOf(TEST_C.toLowerCase()) !== -1) {
        var cells = [];
        for (var c = 0; c < data[r].length; c++) {
          if (String(data[r][c]).trim() !== '') cells.push(String(data[r][c]).trim());
        }
        hitC.push({ hoja: name, gid: gid, fila: r + 1, preview: cells.slice(0, 5).join(' | ') });
      }

      if (rowStr.indexOf(TEST_O.toLowerCase()) !== -1) {
        var cells = [];
        for (var c = 0; c < data[r].length; c++) {
          if (String(data[r][c]).trim() !== '') cells.push(String(data[r][c]).trim());
        }
        hitO.push({ hoja: name, gid: gid, fila: r + 1, preview: cells.slice(0, 5).join(' | ') });
      }
    }
  }

  // ── Resultado CLIENTE ─────────────────────────────────────
  Logger.log('\nRESULTADO "' + TEST_C + '":');
  if (hitC.length === 0) {
    Logger.log('  NO ENCONTRADO.');
    Logger.log('  Posibles causas: formulario no enviado, o han pasado menos de 15 s.');
  } else {
    for (var i = 0; i < hitC.length; i++) {
      Logger.log('  >> Hoja: "' + hitC[i].hoja + '"  GID=' + hitC[i].gid + '  Fila=' + hitC[i].fila);
      Logger.log('     Contenido: ' + hitC[i].preview);
    }
  }

  // ── Resultado OPERARIO ────────────────────────────────────
  Logger.log('\nRESULTADO "' + TEST_O + '":');
  if (hitO.length === 0) {
    Logger.log('  NO ENCONTRADO.');
    Logger.log('  Posibles causas: formulario no enviado, o han pasado menos de 15 s.');
  } else {
    for (var i = 0; i < hitO.length; i++) {
      Logger.log('  >> Hoja: "' + hitO[i].hoja + '"  GID=' + hitO[i].gid + '  Fila=' + hitO[i].fila);
      Logger.log('     Contenido: ' + hitO[i].preview);
    }
  }

  // ── Conclusion ────────────────────────────────────────────
  Logger.log('\n================================================');
  if (hitC.length > 0 && hitO.length > 0) {
    Logger.log('CONCLUSION:');
    Logger.log('  Form Clientes  -> hoja "' + hitC[0].hoja + '" (GID=' + hitC[0].gid + ')');
    Logger.log('  Form Operarios -> hoja "' + hitO[0].hoja + '" (GID=' + hitO[0].gid + ')');
    if (hitC[0].hoja === hitO[0].hoja) {
      Logger.log('  ALERTA: ambos formularios escriben en la MISMA hoja.');
    } else {
      Logger.log('  OK: cada formulario escribe en una hoja diferente.');
    }
  } else {
    Logger.log('CONCLUSION: datos insuficientes. Envia los formularios y reintenta.');
  }
  Logger.log('NINGUNA MODIFICACION REALIZADA.');
  Logger.log('================================================');
}

// ─────────────────────────────────────────────────────────────

/**
 * Diagnóstico de fotos en RESP_TRABAJOS → carpeta_trabajos_id.
 * SOLO LECTURA — no modifica, no copia, no borra nada.
 *
 * Muestra para cada fila de RESP_TRABAJOS:
 *   1. Contenido exacto de la columna Fotos
 *   2. Enlaces Drive detectados
 *   3. IDs de archivo extraídos
 *   4. Si DriveApp.getFileById() puede abrirlos
 * Y al final:
 *   5. carpeta_trabajos_id y si Drive puede abrirla
 */
function debugFotosTrabajos() {
  Logger.log('================================================');
  Logger.log('debugFotosTrabajos() — SOLO LECTURA');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var ss = getTargetSS_();

  // [1] carpeta_trabajos_id
  Logger.log('\n[1] carpeta_trabajos_id:');
  var carpetaId = getConfig('carpeta_trabajos_id');
  if (!carpetaId || carpetaId.trim() === '') {
    Logger.log('  ❌ NO configurado en CONFIG (clave: carpeta_trabajos_id)');
    carpetaId = null;
  } else {
    Logger.log('  Valor: "' + carpetaId + '"');
    try {
      var carpeta = DriveApp.getFolderById(carpetaId);
      Logger.log('  ✓ Drive puede abrir la carpeta: "' + carpeta.getName() + '"');
    } catch (e) {
      Logger.log('  ❌ DriveApp.getFolderById() falla: ' + e.message);
    }
  }

  // [2] Localizar hoja RESP_TRABAJOS
  Logger.log('\n[2] Hoja RESP_TRABAJOS:');
  var respSheet = ss.getSheetByName('RESP_TRABAJOS');
  if (!respSheet) {
    Logger.log('  ❌ Hoja RESP_TRABAJOS no encontrada.');
    Logger.log('  Ejecuta validateTrabajosProcessing() para diagnóstico completo.');
    return;
  }
  var lastRow = respSheet.getLastRow();
  var lastCol = respSheet.getLastColumn();
  Logger.log('  ✓ Encontrada. Filas totales (con cabecera): ' + lastRow + '  Columnas: ' + lastCol);

  if (lastRow < 1 || lastCol < 1) {
    Logger.log('  ❌ Hoja vacía. Sin datos que analizar.');
    return;
  }

  // [3] Localizar TODAS las columnas de Fotos
  Logger.log('\n[3] Cabeceras de RESP_TRABAJOS:');
  var headers = respSheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h).trim(); });
  Logger.log('  ' + headers.join(' | '));

  var fotosColsIdx = [];
  for (var h = 0; h < headers.length; h++) {
    var lh = headers[h].toLowerCase();
    if (lh.indexOf('foto') !== -1 || lh.indexOf('photo') !== -1 || lh === 'imagen' || lh.indexOf('image') !== -1) {
      fotosColsIdx.push(h);
    }
  }

  if (fotosColsIdx.length === 0) {
    Logger.log('\n  ❌ Ninguna columna de Fotos encontrada entre las cabeceras.');
    Logger.log('  Cabeceras disponibles: ' + headers.join(', '));
    Logger.log('  Verifica que el campo del formulario se llame "Fotos" o "Fotos (enlaces Drive)".');
    return;
  }
  Logger.log('\n  Columnas de Fotos detectadas (' + fotosColsIdx.length + '):');
  for (var ci = 0; ci < fotosColsIdx.length; ci++) {
    Logger.log('    [' + (ci + 1) + '] "' + headers[fotosColsIdx[ci]] + '" (posición ' + (fotosColsIdx[ci] + 1) + ')');
  }

  if (lastRow < 2) {
    Logger.log('\n  Sin filas de datos (solo cabecera). Sin fotos que analizar.');
    return;
  }

  // [4] Analizar cada fila — todas las columnas de fotos
  var allData = respSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  Logger.log('\n[4] Análisis fila por fila (' + allData.length + ' respuesta(s)):');

  var totalEnlaces  = 0;
  var totalIds      = 0;
  var totalAbiertos = 0;
  var totalFallidos = 0;

  for (var r = 0; r < allData.length; r++) {
    var rowNum = r + 2;
    Logger.log('\n  --- Fila ' + rowNum + ' ---');

    var rawParts = [];
    for (var ci = 0; ci < fotosColsIdx.length; ci++) {
      var rawCell = allData[r][fotosColsIdx[ci]];
      var rawVal  = (rawCell === null || rawCell === undefined) ? '' : String(rawCell).trim();
      Logger.log('  Col "' + headers[fotosColsIdx[ci]] + '": ' + (rawVal === '' ? '(vacío)' : '"' + rawVal + '"'));
      if (rawVal !== '') rawParts.push(rawVal);
    }

    var rawStr = rawParts.join(', ');
    Logger.log('  Valor combinado: ' + (rawStr === '' ? '(vacío — sin fotos en esta fila)' : '"' + rawStr + '"'));

    if (rawStr === '') continue;

    // Extraer enlaces Drive
    var enlaces = _extractDriveLinks_(rawStr);
    Logger.log('  Enlaces Drive detectados (' + enlaces.length + '):');
    if (enlaces.length === 0) {
      Logger.log('    (ninguno — el valor no contiene URLs de Drive reconocibles)');
    }
    for (var e = 0; e < enlaces.length; e++) {
      Logger.log('    [' + (e + 1) + '] ' + enlaces[e]);
    }
    totalEnlaces += enlaces.length;

    // Extraer IDs de archivo
    var ids = _extractDriveFileIds_(rawStr);
    Logger.log('  IDs de archivo extraídos (' + ids.length + '):');
    if (ids.length === 0) {
      Logger.log('    (ninguno)');
    }
    for (var i = 0; i < ids.length; i++) {
      Logger.log('    [' + (i + 1) + '] ID: "' + ids[i] + '"');
    }
    totalIds += ids.length;

    // Intentar abrir cada archivo
    Logger.log('  Prueba DriveApp.getFileById():');
    if (ids.length === 0) {
      Logger.log('    (sin IDs que probar)');
    }
    for (var j = 0; j < ids.length; j++) {
      try {
        var archivo = DriveApp.getFileById(ids[j]);
        Logger.log('    ✓ ID "' + ids[j] + '" → "' + archivo.getName() + '" (' + archivo.getMimeType() + ')');
        totalAbiertos++;
      } catch (err) {
        Logger.log('    ❌ ID "' + ids[j] + '" → ' + err.message);
        totalFallidos++;
      }
    }
  }

  // [5] Resumen
  Logger.log('\n================================================');
  Logger.log('RESUMEN:');
  Logger.log('  Filas analizadas:         ' + allData.length);
  Logger.log('  Total enlaces detectados: ' + totalEnlaces);
  Logger.log('  Total IDs extraídos:      ' + totalIds);
  Logger.log('  Archivos accesibles:      ' + totalAbiertos);
  Logger.log('  Archivos inaccesibles:    ' + totalFallidos);
  Logger.log('  carpeta_trabajos_id:      ' + (carpetaId || '❌ no configurado'));
  Logger.log('');
  if (totalIds === 0 && totalEnlaces === 0) {
    Logger.log('  ⚠️ No se encontraron enlaces Drive en ninguna respuesta.');
    Logger.log('  Comprueba que el campo Fotos del formulario recibe URLs de Drive.');
  } else if (totalFallidos > 0) {
    Logger.log('  ⚠️ Algunos archivos no son accesibles (permisos o ID incorrecto).');
    Logger.log('  El propietario del script necesita acceso a esos archivos Drive.');
  } else if (totalAbiertos === totalIds && totalIds > 0) {
    Logger.log('  ✓ Todos los archivos son accesibles.');
    if (!carpetaId) {
      Logger.log('  ❌ Falta carpeta_trabajos_id en CONFIG → las fotos no se copiarán.');
    }
  }
  Logger.log('NINGUNA MODIFICACIÓN REALIZADA.');
  Logger.log('================================================');
}

/**
 * Extrae todas las URLs de Drive de un string (separadas por coma, espacio o salto de línea).
 */
function _extractDriveLinks_(str) {
  var links = [];
  var re = /https:\/\/drive\.google\.com\/[^\s,\n"<>]+/gi;
  var match;
  while ((match = re.exec(str)) !== null) {
    links.push(match[0].replace(/[.,;)]+$/, ''));
  }
  return links;
}

/**
 * Extrae IDs de archivo Drive de un string.
 * Soporta:
 *   /file/d/ID/
 *   open?id=ID
 *   uc?id=ID
 *   id=ID  (genérico)
 *   IDs directos (alfanumérico 25-44 chars sin espacios, si no hay URLs)
 */
function _extractDriveFileIds_(str) {
  var ids = [];
  var seen = {};

  var patterns = [
    /\/file\/d\/([a-zA-Z0-9_\-]{10,})/g,
    /[?&]id=([a-zA-Z0-9_\-]{10,})/g,
    /\/folders\/([a-zA-Z0-9_\-]{10,})/g
  ];

  for (var p = 0; p < patterns.length; p++) {
    var re = patterns[p];
    var match;
    while ((match = re.exec(str)) !== null) {
      var id = match[1];
      if (!seen[id]) { seen[id] = true; ids.push(id); }
    }
  }
  return ids;
}

// ─────────────────────────────────────────────────────────────

/**
 * Valida el estado de las columnas de fotos en TRABAJOS.
 * SOLO LECTURA — no modifica nada.
 *
 * Comprueba:
 *   1. Columnas Foto 1-5 y Carpeta Fotos existen
 *   2. carpeta_trabajos_id accesible en Drive
 *   3. Trabajos con Fotos tienen URLs individuales en Foto 1
 *   4. Archivos de Foto 1 accesibles con DriveApp
 */
function validateTrabajoPhotos() {
  Logger.log('================================================');
  Logger.log('validateTrabajoPhotos() — SOLO LECTURA');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  // [1] Columnas en TRABAJOS
  Logger.log('\n[1] Columnas de fotos en TRABAJOS:');
  var trabSheet;
  try { trabSheet = getSheet('TRABAJOS'); } catch (e) {
    Logger.log('  ❌ Hoja TRABAJOS no encontrada: ' + e.message);
    return;
  }

  var lastCol = trabSheet.getLastColumn();
  var headers = lastCol > 0
    ? trabSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); })
    : [];

  var expected   = ['Foto 1', 'Foto 2', 'Foto 3', 'Foto 4', 'Foto 5', 'Carpeta Fotos'];
  var colMap     = {};
  var allPresent = true;

  for (var i = 0; i < expected.length; i++) {
    var found = -1;
    for (var j = 0; j < headers.length; j++) {
      if (headers[j] === expected[i]) { found = j; break; }
    }
    if (found === -1) {
      Logger.log('  ❌ "' + expected[i] + '" — NO EXISTE');
      allPresent = false;
    } else {
      Logger.log('  ✓ "' + expected[i] + '" — columna ' + (found + 1));
      colMap[expected[i]] = found;
    }
  }

  if (!allPresent) {
    Logger.log('  → Ejecuta repairExistingTrabajoPhotos() para crear las columnas.');
  }

  // [2] carpeta_trabajos_id
  Logger.log('\n[2] carpeta_trabajos_id:');
  var carpetaRaizId = getConfig('carpeta_trabajos_id');
  if (!carpetaRaizId || carpetaRaizId.trim() === '') {
    Logger.log('  ❌ NO configurado en CONFIG (clave: carpeta_trabajos_id).');
    Logger.log('  Las carpetas por trabajo no se crearán hasta configurarlo.');
  } else {
    Logger.log('  Valor: "' + carpetaRaizId + '"');
    try {
      var raiz = DriveApp.getFolderById(carpetaRaizId);
      Logger.log('  ✓ Accesible: "' + raiz.getName() + '"');
    } catch (e) {
      Logger.log('  ❌ No accesible: ' + e.message);
    }
  }

  // [3] Revisión fila a fila
  Logger.log('\n[3] Revisión fila a fila:');
  var lastRow = trabSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('  Sin filas de datos.');
    Logger.log('NINGUNA MODIFICACIÓN REALIZADA.');
    Logger.log('================================================');
    return;
  }

  var data        = trabSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var fotosIdx    = COL_TRA.FOTOS - 1;
  var foto1Idx    = colMap['Foto 1'];
  var carpetaIdx  = colMap['Carpeta Fotos'];

  var conFotosSrc = 0;
  var conFoto1    = 0;
  var sinFoto1    = 0;
  var conCarpeta  = 0;
  var fileOk      = 0;
  var fileFail    = 0;

  for (var r = 0; r < data.length; r++) {
    var row        = data[r];
    var workId     = String(row[0] || '').trim();
    var fotosVal   = String(row[fotosIdx] || '').trim();
    var foto1Val   = foto1Idx   !== undefined ? String(row[foto1Idx]   || '').trim() : '';
    var carpetaVal = carpetaIdx !== undefined ? String(row[carpetaIdx] || '').trim() : '';

    if (fotosVal)   conFotosSrc++;
    if (foto1Val)   conFoto1++;
    if (carpetaVal) conCarpeta++;

    if (fotosVal && !foto1Val) {
      Logger.log('  ⚠️ ' + workId + ': tiene Fotos pero Foto 1 vacía → necesita repairExistingTrabajoPhotos()');
      sinFoto1++;
    }

    // Verificar acceso a Foto 1
    if (foto1Val) {
      var ids = _extractDriveFileIds_(foto1Val);
      for (var fi = 0; fi < ids.length; fi++) {
        try {
          var f = DriveApp.getFileById(ids[fi]);
          Logger.log('  ✓ ' + workId + ' Foto 1 → "' + f.getName() + '"');
          fileOk++;
        } catch (e) {
          Logger.log('  ❌ ' + workId + ' Foto 1 ID "' + ids[fi] + '": ' + e.message);
          fileFail++;
        }
      }
    }

    // Verificar carpeta
    if (carpetaVal) {
      var carpIds = carpetaVal.match(/\/folders\/([a-zA-Z0-9_\-]{10,})/);
      if (carpIds) {
        try { DriveApp.getFolderById(carpIds[1]); }
        catch (e) { Logger.log('  ❌ ' + workId + ' Carpeta no accesible: ' + e.message); }
      }
    }
  }

  // [4] Resumen
  Logger.log('\n================================================');
  Logger.log('RESUMEN:');
  Logger.log('  Total trabajos:              ' + data.length);
  Logger.log('  Con Fotos (resumen):         ' + conFotosSrc);
  Logger.log('  Con Foto 1 (URL individual): ' + conFoto1);
  Logger.log('  Con Carpeta Fotos:           ' + conCarpeta);
  Logger.log('  Necesitan repair:            ' + sinFoto1);
  Logger.log('  Archivos accesibles:         ' + fileOk);
  Logger.log('  Archivos inaccesibles:       ' + fileFail);
  Logger.log('');
  if (!allPresent || sinFoto1 > 0) {
    Logger.log('ACCIÓN: Ejecuta repairExistingTrabajoPhotos()');
  } else if (fileFail > 0) {
    Logger.log('ACCIÓN: Revisa permisos Drive del propietario del script.');
  } else if (conFoto1 > 0) {
    Logger.log('✓ Todo correcto.');
  } else {
    Logger.log('  Sin fotos registradas todavía.');
  }
  Logger.log('NINGUNA MODIFICACIÓN REALIZADA.');
  Logger.log('================================================');
}

// ── Utilidad privada ──────────────────────────────────────────

function inferirTipoHoja_(name, headers) {
  if (name.indexOf('ARCHIVO_') === 0) return 'ARCHIVADA';

  var lh = [];
  for (var i = 0; i < headers.length; i++) {
    lh.push(headers[i].toLowerCase());
  }

  var has = function(term) {
    for (var i = 0; i < lh.length; i++) {
      if (lh[i].indexOf(term) !== -1) return true;
    }
    return false;
  };

  if (has('id_cliente'))  return 'SISTEMA — CLIENTES (ID_Cliente en cabecera)';
  if (has('id_operario')) return 'SISTEMA — OPERARIOS (ID_Operario en cabecera)';
  if (has('id_trabajo'))  return 'SISTEMA — TRABAJOS (ID_Trabajo en cabecera)';
  if (has('id_factura'))  return 'SISTEMA — FACTURAS (ID_Factura en cabecera)';

  var tieneTS  = has('marca temporal') || has('timestamp');
  var esCliente  = has('empresa') || has('direcci') || (has('cif') || has('nif'));
  var esOperario = has('especialidad');

  if (/respuestas.de.formulario|form.responses/i.test(name)) {
    if (esCliente)  return 'POSIBLE respuestas CLIENTES — NO confirmado (usar findTestResponses)';
    if (esOperario) return 'POSIBLE respuestas OPERARIOS — NO confirmado (usar findTestResponses)';
    return 'Hoja de respuestas Google — tipo no determinado';
  }

  if (tieneTS) {
    if (esCliente)  return 'POSIBLE respuestas CLIENTES — NO confirmado';
    if (esOperario) return 'POSIBLE respuestas OPERARIOS — NO confirmado';
    return 'Hoja con marca temporal — tipo no determinado';
  }

  if (lh.length === 0) return 'VACIA';
  return 'DESCONOCIDA — primeras cabeceras: [' + lh.slice(0, 4).join(', ') + ']';
}

// ============================================================
// diagnosticaTRA006 — diagnóstico específico de TRA-006
// SOLO LECTURA — no modifica ningún dato.
// Ejecutar desde el editor de Apps Script.
// ============================================================
function diagnosticaTRA006() {
  var TARGET_ID = 'TRA-006';
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var tz  = Session.getScriptTimeZone();

  Logger.log('================================================');
  Logger.log('diagnosticaTRA006() — SOLO LECTURA');
  Logger.log(Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  // ── 1. Parámetros del panel (tal como los lee panelBuscarTrabajos) ──
  var sh = ss.getSheetByName('PANEL');
  var clientePanel = '';
  var fechaIniStr  = '';
  var fechaFinStr  = '';
  if (sh) {
    clientePanel = String(sh.getRange(5, 3).getValue() || '').trim();
    fechaIniStr  = sh.getRange(6, 3).getDisplayValue().trim();
    fechaFinStr  = sh.getRange(7, 3).getDisplayValue().trim();
  }
  Logger.log('\n[PANEL] cliente   : "' + clientePanel + '"');
  Logger.log('[PANEL] fechaIni  : "' + fechaIniStr + '"');
  Logger.log('[PANEL] fechaFin  : "' + fechaFinStr + '"');

  var fi = _parseReportDate_(fechaIniStr);
  var ff = _parseReportDate_(fechaFinStr);
  if (ff) ff.setHours(23, 59, 59, 999);
  Logger.log('[PANEL] fi parsed : ' + (fi ? fi.toISOString() : 'null — fecha inicio INVÁLIDA'));
  Logger.log('[PANEL] ff parsed : ' + (ff ? ff.toISOString() : 'null — fecha fin INVÁLIDA'));

  // ── 2. Localizar la fila de TRA-006 en TRABAJOS ──────────────────
  Logger.log('\n================================================');
  Logger.log('[TRABAJOS] Buscando ' + TARGET_ID + '...');
  var shTra = ss.getSheetByName('TRABAJOS');
  if (!shTra) { Logger.log('❌ Hoja TRABAJOS no encontrada.'); return; }

  var lastRow = shTra.getLastRow();
  var lastCol = shTra.getLastColumn();
  Logger.log('[TRABAJOS] lastRow=' + lastRow + '  lastCol=' + lastCol);

  var numCols = Math.min(lastCol, 19); // COL_TRA.MATERIALES = 19
  var data    = shTra.getRange(2, 1, lastRow - 1, numCols).getValues();
  var headers = shTra.getRange(1, 1, 1, numCols).getValues()[0];

  var targetRow = null;
  var targetRowNum = -1;
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === TARGET_ID) {
      targetRow    = data[i];
      targetRowNum = i + 2; // fila en la hoja (base 1, +1 por cabecera)
      break;
    }
  }

  if (!targetRow) {
    Logger.log('❌ ' + TARGET_ID + ' NO ENCONTRADO en la hoja TRABAJOS.');
    Logger.log('   Comprueba que la columna A (ID) contiene exactamente "' + TARGET_ID + '".');
    return;
  }
  Logger.log('✓ Encontrado en fila ' + targetRowNum + ' de la hoja.');

  // ── 3. Volcar la fila completa ────────────────────────────────────
  Logger.log('\n[FILA COMPLETA]');
  for (var c = 0; c < numCols; c++) {
    var val  = targetRow[c];
    var tipo = typeof val;
    var repr = (val instanceof Date)
      ? val.toISOString() + ' (Date en hoja → ' + Utilities.formatDate(val, tz, 'dd/MM/yyyy HH:mm:ss') + ')'
      : '"' + String(val) + '" (' + tipo + ')';
    Logger.log('  col ' + (c + 1) + ' [' + (headers[c] || '?') + ']: ' + repr);
  }

  // ── 4. Campos clave ───────────────────────────────────────────────
  Logger.log('\n[CAMPOS CLAVE]');
  var rawId      = String(targetRow[0] || '').trim();
  var rawCliente = String(targetRow[2] || '').trim(); // COL 3 = índice 2
  var rawFecha   = targetRow[4];                      // COL 5 = índice 4
  var rawOpe     = String(targetRow[3] || '').trim(); // COL 4 = índice 3
  var rawEstado  = String(targetRow[9] || '').trim(); // COL 10 = índice 9

  Logger.log('  ID        : "' + rawId + '"');
  Logger.log('  Cliente   : "' + rawCliente + '"  length=' + rawCliente.length);
  Logger.log('  Cliente charCodes: ' + rawCliente.split('').map(function(c){return c.charCodeAt(0);}).join(','));
  Logger.log('  Operario  : "' + rawOpe + '"');
  Logger.log('  Estado    : "' + rawEstado + '"');
  Logger.log('  Fecha raw : ' + rawFecha + '  tipo=' + typeof rawFecha +
             (rawFecha instanceof Date ? '  ISO=' + rawFecha.toISOString() : ''));

  // ── 5. Filtro 1 — cliente ─────────────────────────────────────────
  Logger.log('\n[FILTRO 1 — CLIENTE]');
  Logger.log('  clientePanel.toLowerCase() : "' + clientePanel.toLowerCase() + '"');
  Logger.log('  rawCliente.toLowerCase()   : "' + rawCliente.toLowerCase() + '"');
  var pasaCliente = (rawCliente.toLowerCase() === clientePanel.toLowerCase());
  Logger.log('  Resultado: ' + (pasaCliente ? '✓ PASA' : '❌ RECHAZADO — los valores no coinciden'));
  if (!pasaCliente) {
    // Mostrar primer carácter diferente
    var a = rawCliente.toLowerCase();
    var b = clientePanel.toLowerCase();
    for (var ci = 0; ci < Math.max(a.length, b.length); ci++) {
      if (a[ci] !== b[ci]) {
        Logger.log('  Primera diferencia en posición ' + ci +
                   ': rawCliente[' + ci + ']=U+' + (a.charCodeAt(ci) || 'undefined').toString(16) +
                   '  panel[' + ci + ']=U+' + (b.charCodeAt(ci) || 'undefined').toString(16));
        break;
      }
    }
  }

  // ── 6. Filtro 2 — fecha ───────────────────────────────────────────
  Logger.log('\n[FILTRO 2 — FECHA]');
  var ft = _parseTrabajosDate_(rawFecha);
  Logger.log('  rawFecha (tipo ' + typeof rawFecha + '): ' + rawFecha);
  Logger.log('  ft parseada : ' + (ft ? ft.toISOString() + ' → ' + Utilities.formatDate(ft, tz, 'dd/MM/yyyy') : 'null — NO PARSEABLE'));
  if (fi && ff) {
    Logger.log('  fi (límite inferior): ' + fi.toISOString() + ' → ' + Utilities.formatDate(fi, tz, 'dd/MM/yyyy'));
    Logger.log('  ff (límite superior): ' + ff.toISOString() + ' → ' + Utilities.formatDate(ff, tz, 'dd/MM/yyyy'));
    var pasaFecha = !(!ft || ft < fi || ft > ff);
    if (!ft)        Logger.log('  Resultado: ❌ RECHAZADO — ft es null (fecha no parseable)');
    else if (ft < fi) Logger.log('  Resultado: ❌ RECHAZADO — ft (' + Utilities.formatDate(ft,tz,'dd/MM/yyyy') + ') < fi (' + Utilities.formatDate(fi,tz,'dd/MM/yyyy') + ')');
    else if (ft > ff) Logger.log('  Resultado: ❌ RECHAZADO — ft (' + Utilities.formatDate(ft,tz,'dd/MM/yyyy') + ') > ff (' + Utilities.formatDate(ff,tz,'dd/MM/yyyy') + ')');
    else              Logger.log('  Resultado: ✓ PASA');
  } else {
    Logger.log('  ⚠️ fi o ff no válidas — filtro de fecha no evaluable');
    var pasaFecha = false;
  }

  // ── 7. ¿Entra en la lista de trabajos? ───────────────────────────
  Logger.log('\n[RESULTADO FILTROS]');
  Logger.log('  Pasa filtro cliente: ' + pasaCliente);
  Logger.log('  Pasa filtro fecha  : ' + pasaFecha);
  Logger.log('  → ' + (pasaCliente && pasaFecha
    ? '✓ ENTRA en trabajos[]'
    : '❌ NO ENTRA en trabajos[] — rechazado en filtro ' + (!pasaCliente ? 'CLIENTE' : 'FECHA')));

  // ── 8. DOC_TRABAJOS — ¿hay referencia a TRA-006? ─────────────────
  Logger.log('\n[DOC_TRABAJOS]');
  var shDT = ss.getSheetByName('DOC_TRABAJOS');
  if (!shDT || shDT.getLastRow() < 2) {
    Logger.log('  DOC_TRABAJOS vacía o no existe — no hay filtro por facturación.');
  } else {
    var dtData = shDT.getRange(2, 1, shDT.getLastRow() - 1, 3).getValues();
    var refs = [];
    for (var di = 0; di < dtData.length; di++) {
      if (String(dtData[di][2] || '').trim() === TARGET_ID) {
        refs.push({ docId: dtData[di][0], tipo: dtData[di][1] });
      }
    }
    if (refs.length === 0) {
      Logger.log('  ' + TARGET_ID + ' NO aparece en DOC_TRABAJOS → estado = "Sin facturar"');
      Logger.log('  DOC_TRABAJOS no filtra filas — solo asigna etiqueta de estado.');
    } else {
      Logger.log('  ' + TARGET_ID + ' aparece ' + refs.length + ' vez/veces en DOC_TRABAJOS:');
      refs.forEach(function(r) { Logger.log('    docId=' + r.docId + '  tipo=' + r.tipo); });
      Logger.log('  ⚠️ IMPORTANTE: panelBuscarTrabajos NO excluye trabajos facturados.');
      Logger.log('  DOC_TRABAJOS solo afecta al color/etiqueta en la tabla, NO a si aparece.');
    }
  }

  // ── 9. Filtros adicionales que NO existen ─────────────────────────
  Logger.log('\n[FILTROS ADICIONALES — VERIFICACIÓN]');
  Logger.log('  Filtro por estado de trabajo (Pendiente/Completado/etc.) : NO EXISTE en panelBuscarTrabajos');
  Logger.log('  Filtro por facturado/proforma                            : NO EXISTE en panelBuscarTrabajos');
  Logger.log('  Filtro por documento previo                              : NO EXISTE en panelBuscarTrabajos');
  Logger.log('  Deduplicación por ID                                     : NO EXISTE en panelBuscarTrabajos');
  Logger.log('  Límite de filas leídas (MAX_ROWS=30)                     : solo limita la VISUALIZACIÓN, no la búsqueda');

  // ── 10. Línea de código responsable ──────────────────────────────
  Logger.log('\n[LÍNEA DE CÓDIGO RESPONSABLE]');
  if (!pasaCliente) {
    Logger.log('  Panel.gs — función panelBuscarTrabajos():');
    Logger.log('    if (String(row[COL_TRA.CLIENTE-1]||"").trim().toLowerCase() !== cliente.toLowerCase()) continue;');
    Logger.log('  Esta línea rechaza la fila porque el nombre del cliente no coincide.');
  } else if (!pasaFecha) {
    Logger.log('  Panel.gs — función panelBuscarTrabajos():');
    Logger.log('    var ft = _parseTrabajosDate_(row[COL_TRA.FECHA_TRABAJO - 1]);');
    Logger.log('    if (!ft || ft < fi || ft > ff) continue;');
    Logger.log('  La fecha almacenada (' + (ft ? Utilities.formatDate(ft,tz,'dd/MM/yyyy') : 'null') + ') queda fuera del rango ' + fechaIniStr + ' → ' + fechaFinStr + '.');
    Logger.log('  Causa probable: la fecha fue grabada en formato MM/DD/YYYY por new Date("DD/MM/YYYY")');
    Logger.log('  Ver createTrabajo() en Trabajos.gs — línea que hace new Date(fechaTrabajo).');
  } else {
    Logger.log('  ✓ TRA-006 PASA los dos filtros de panelBuscarTrabajos.');
    Logger.log('  Si no aparece en pantalla, el problema está en el renderizado de la tabla (MAX_ROWS, rowHeight).');
    Logger.log('  Revisa si hay más de 30 trabajos que cumplen los criterios — MAX_ROWS=30 limita la visualización.');
  }

  Logger.log('\n================================================');
  Logger.log('FIN diagnosticaTRA006()');
  Logger.log('================================================');
}
