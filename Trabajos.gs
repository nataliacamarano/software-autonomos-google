// ============================================================
// Trabajos.gs — CRUD de trabajos (partes de trabajo)
// ============================================================

// Columnas de la hoja TRABAJOS (base 1)
const COL_TRA = {
  ID:                1,
  FECHA_ALTA:        2,
  CLIENTE:           3,
  OPERARIO:          4,
  FECHA_TRABAJO:     5,
  HORA_INICIO:       6,
  HORA_FIN:          7,
  TRABAJO_REALIZADO: 8,
  INCIDENCIAS:       9,
  ESTADO:            10,
  FOTOS:             11,
  UBICACION:         12,
  FOTO_1:            13,
  FOTO_2:            14,
  FOTO_3:            15,
  FOTO_4:            16,
  FOTO_5:            17,
  CARPETA_FOTOS:     18,
  MATERIALES:        19,
  // Empleado 2 (opcional) — columnas añadidas al final para no desplazar las existentes.
  OPERARIO_2:        20,
  HORA_INICIO_2:     21,
  HORA_FIN_2:        22
};

const ESTADOS_TRABAJO = ['Pendiente', 'En curso', 'Completado', 'Cancelado'];

const HEADERS_TRABAJOS_V2 = [
  'ID', 'Fecha Alta', 'Cliente', 'Operario',
  'Fecha Trabajo', 'Hora Inicio', 'Hora Fin',
  'Trabajo Realizado', 'Incidencias', 'Estado', 'Fotos', 'Ubicación',
  'Foto 1', 'Foto 2', 'Foto 3', 'Foto 4', 'Foto 5', 'Carpeta Fotos',
  'Materiales',
  'Operario 2', 'Hora Inicio 2', 'Hora Fin 2'
];

// ---- Lectura ----

function getTrabajosAll() {
  const sheet = getSheet('TRABAJOS');
  return rowsToObjects(sheet.getDataRange().getValues());
}

function getTrabajoById(id) {
  return getTrabajosAll().find(function(t) {
    return String(t['ID']) === String(id);
  }) || null;
}

// ---- Escritura ----

/**
 * Crea un parte de trabajo. Usado por el panel web y el trigger de formulario.
 * Devuelve { success, id } o { success: false, error }.
 */
function createTrabajo(data) {
  try {
    var sheet = getSheet('TRABAJOS');
    _ensureTrabajoPhotoColumns_(sheet);
    _ensureTrabajoMaterialesColumn_(sheet);
    _ensureTrabajoEmpleado2Columns_(sheet);
    var id = data.id || nextId('TRABAJOS', 'TRA');

    // Empleado 2 (opcional). Si no hay operario 2, sus horas se ignoran.
    var operario2   = String(data.operario2   || '').trim();
    var horaInicio2 = String(data.horaInicio2 || '').trim();
    var horaFin2    = String(data.horaFin2    || '').trim();

    // Validación hora fin > hora inicio (no bloquea el alta: avisa y no computará
    // esas horas, igual que _calcHoras_, que devuelve 0 si fin <= inicio).
    _validarHorasTrabajo_(id, String(data.horaInicio || '').trim(), String(data.horaFin || '').trim(), 1);
    if (operario2) _validarHorasTrabajo_(id, horaInicio2, horaFin2, 2);

    // Fecha Trabajo: convertir string a Date.
    // _parseReportDate_ interpreta DD/MM/YYYY (formato del formulario en español).
    // new Date() se usa solo como fallback para ISO (YYYY-MM-DD) u otros formatos
    // no ambiguos; sin este orden, "03/06/2026" se parseaba como marzo (MM/DD/YYYY).
    var fechaTrabajo = data.fechaTrabajo || '';
    if (fechaTrabajo && typeof fechaTrabajo === 'string') {
      var d = _parseReportDate_(fechaTrabajo.trim());
      if (!d) { var tmp = new Date(fechaTrabajo); if (!isNaN(tmp.getTime())) d = tmp; }
      if (d) fechaTrabajo = d;
    }

    sheet.appendRow([
      id,
      new Date(),
      String(data.cliente           || '').trim(),
      String(data.operario          || '').trim(),
      fechaTrabajo,
      String(data.horaInicio        || '').trim(),
      String(data.horaFin           || '').trim(),
      String(data.trabajoRealizado  || '').trim(),
      String(data.incidencias       || '').trim(),
      String(data.estado            || 'Pendiente').trim(),
      String(data.fotos             || '').trim(),
      String(data.ubicacion         || '').trim(),
      String(data.foto1             || '').trim(),
      String(data.foto2             || '').trim(),
      String(data.foto3             || '').trim(),
      String(data.foto4             || '').trim(),
      String(data.foto5             || '').trim(),
      String(data.carpetaFotos      || '').trim(),
      Number(data.materiales        || 0),
      operario2,
      horaInicio2,
      horaFin2
    ]);

    logEvent('INFO', 'createTrabajo', 'Trabajo creado: ' + id, String(data.cliente));
    return { success: true, id: id };
  } catch (e) {
    logEvent('ERROR', 'createTrabajo', e.message, JSON.stringify(data));
    return { success: false, error: e.message };
  }
}

/**
 * Actualiza el estado de un trabajo.
 */
function updateTrabajoEstado(id, nuevoEstado) {
  try {
    const sheet = getSheet('TRABAJOS');
    const row   = findRowById_(sheet, id);
    if (!row) return { success: false, error: 'Trabajo "' + id + '" no encontrado.' };
    sheet.getRange(row, COL_TRA.ESTADO).setValue(nuevoEstado);
    logEvent('INFO', 'updateTrabajoEstado', id + ' → ' + nuevoEstado, '');
    return { success: true };
  } catch (e) {
    logEvent('ERROR', 'updateTrabajoEstado', e.message, id);
    return { success: false, error: e.message };
  }
}

// ---- Handler de formulario ----

/**
 * Procesa namedValues de una respuesta del formulario de trabajos.
 * Llamado desde onFormSubmitHandler() en Triggers.gs.
 */
function onTrabajosFormSubmit(namedValues) {
  var rawFotos  = _combineFotos_(namedValues);
  var workId    = nextId('TRABAJOS', 'TRA');
  var fotosInfo = _procesarFotosParaTrabajo_(rawFotos, workId);

  createTrabajo({
    id:               workId,
    cliente:          (namedValues['Cliente']           || [''])[0],
    operario:         (namedValues['Operario']          || [''])[0],
    fechaTrabajo:     (namedValues['Fecha']             || [''])[0],
    horaInicio:       (namedValues['Hora inicio']       || [''])[0],
    horaFin:          (namedValues['Hora fin']          || [''])[0],
    operario2:        (namedValues['Operario 2']        || [''])[0],
    horaInicio2:      (namedValues['Hora inicio 2']     || [''])[0],
    horaFin2:         (namedValues['Hora fin 2']        || [''])[0],
    trabajoRealizado: (namedValues['Trabajo realizado'] || [''])[0],
    incidencias:      (namedValues['Incidencias']       || [''])[0],
    estado:           (namedValues['Estado']            || ['Pendiente'])[0],
    fotos:            rawFotos,
    ubicacion:        (namedValues['Ubicación']         || [''])[0],
    materiales:       parseFloat(String((namedValues['Materiales (€)'] || ['0'])[0]).replace(',', '.')) || 0,
    foto1:            fotosInfo.foto1,
    foto2:            fotosInfo.foto2,
    foto3:            fotosInfo.foto3,
    foto4:            fotosInfo.foto4,
    foto5:            fotosInfo.foto5,
    carpetaFotos:     fotosInfo.carpetaFotos
  });
}

// ---- Reparación manual de respuestas ----

/**
 * Procesa todas las respuestas del formulario de trabajos aún no volcadas a TRABAJOS.
 * Lee RESP_TRABAJOS o cualquier "Form Responses N" con cabeceras de trabajos.
 * Evita duplicados contando las filas ya procesadas en TRABAJOS:
 *   si TRABAJOS tiene N filas de datos, se saltan las N primeras respuestas del formulario.
 * NO borra datos. NO recrea formularios. NO toca Clientes, Operarios, Facturas ni Gastos.
 *
 * Ejecutar DESPUÉS de validateTrabajosProcessing() para confirmar el estado.
 */
function processExistingTrabajoResponses() {
  Logger.log('================================================');
  Logger.log('processExistingTrabajoResponses()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var ss = getTargetSS_();

  // 1. Localizar hoja de respuestas
  var respSheet = _findTrabajosResponseSheet_(ss);
  if (!respSheet) {
    Logger.log('❌ No se encontró la hoja de respuestas del formulario de trabajos.');
    Logger.log('   Buscadas: RESP_TRABAJOS, Form Responses N con cabeceras de trabajos.');
    Logger.log('   Ejecuta validateTrabajosProcessing() para más detalles.');
    return;
  }
  Logger.log('[1] Hoja de respuestas: "' + respSheet.getName() + '"');

  // 2. Leer todas las filas
  var lastRow = respSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('[2] Sin respuestas que procesar (hoja vacía o solo cabecera).');
    return;
  }
  var allData = respSheet.getDataRange().getValues();
  var headers = allData[0].map(function(h) { return String(h).trim(); });
  Logger.log('[2] Respuestas totales: ' + (lastRow - 1));
  Logger.log('    Cabeceras: ' + headers.join(' | '));

  // 3. Contar filas ya procesadas en TRABAJOS (excluye cabecera)
  var trabajosSheet = getSheet('TRABAJOS');
  var trabajosRows  = Math.max(0, trabajosSheet.getLastRow() - 1);
  Logger.log('[3] Filas ya en TRABAJOS: ' + trabajosRows);

  var totalResp  = lastRow - 1;
  var pendientes = totalResp - trabajosRows;
  Logger.log('    Pendientes: ' + pendientes);

  if (pendientes <= 0) {
    Logger.log('✓ Todas las respuestas ya están procesadas en TRABAJOS. Nada que hacer.');
    return;
  }

  // 4. Procesar las filas pendientes (índice base-0 sobre allData, fila 1 = primera respuesta)
  Logger.log('[4] Procesando ' + pendientes + ' respuesta(s)...');
  var procesados = 0;
  var errores    = 0;

  for (var r = trabajosRows + 1; r <= totalResp; r++) {
    var row = allData[r];
    var nv  = _rowToNamedValues_(headers, row);

    Logger.log('\n    Fila ' + (r + 1) + ':');
    Logger.log('    Cliente:          ' + ((nv['Cliente']           || [''])[0] || '(vacío)'));
    Logger.log('    Operario:         ' + ((nv['Operario']          || [''])[0] || '(vacío)'));
    Logger.log('    Fecha:            ' + ((nv['Fecha']             || [''])[0] || '(vacío)'));
    Logger.log('    Trabajo realizado:' + ((nv['Trabajo realizado'] || [''])[0] || '(vacío)'));

    var rawFotos  = _combineFotos_(nv);
    var workId    = nextId('TRABAJOS', 'TRA');
    var fotosInfo = _procesarFotosParaTrabajo_(rawFotos, workId);
    var nFotos    = [fotosInfo.foto1, fotosInfo.foto2, fotosInfo.foto3, fotosInfo.foto4, fotosInfo.foto5].filter(function(u) { return u; }).length;
    Logger.log('    Fotos procesadas: ' + nFotos);

    var result = createTrabajo({
      id:               workId,
      cliente:          (nv['Cliente']           || [''])[0],
      operario:         (nv['Operario']          || [''])[0],
      fechaTrabajo:     (nv['Fecha']             || [''])[0],
      horaInicio:       (nv['Hora inicio']       || [''])[0],
      horaFin:          (nv['Hora fin']          || [''])[0],
      operario2:        (nv['Operario 2']        || [''])[0],
      horaInicio2:      (nv['Hora inicio 2']     || [''])[0],
      horaFin2:         (nv['Hora fin 2']        || [''])[0],
      trabajoRealizado: (nv['Trabajo realizado'] || [''])[0],
      incidencias:      (nv['Incidencias']       || [''])[0],
      estado:           (nv['Estado']            || ['Pendiente'])[0],
      fotos:            rawFotos,
      ubicacion:        (nv['Ubicación']         || [''])[0],
      foto1:            fotosInfo.foto1,
      foto2:            fotosInfo.foto2,
      foto3:            fotosInfo.foto3,
      foto4:            fotosInfo.foto4,
      foto5:            fotosInfo.foto5,
      carpetaFotos:     fotosInfo.carpetaFotos
    });

    if (result && result.success) {
      Logger.log('    ✓ Creado: ' + result.id);
      procesados++;
    } else {
      Logger.log('    ❌ Error: ' + (result ? result.error : 'desconocido'));
      errores++;
    }
  }

  Logger.log('\n================================================');
  Logger.log('RESULTADO FINAL:');
  Logger.log('  Procesados: ' + procesados);
  Logger.log('  Errores:    ' + errores);
  Logger.log('  TRABAJOS tiene ahora: ' + Math.max(0, getSheet('TRABAJOS').getLastRow() - 1) + ' filas');
  if (errores > 0) {
    Logger.log('  ⚠️ Revisa los errores y verifica los nombres de los campos del formulario.');
  }
  Logger.log('================================================');
}

/**
 * Diagnóstico completo del procesamiento del formulario de trabajos.
 * Comprueba: respuestas existentes en Form Responses 3 / RESP_TRABAJOS,
 * filas procesadas en TRABAJOS, trigger activo y último ID generado.
 * Indica exactamente qué función ejecutar a continuación.
 */
function validateTrabajosProcessing() {
  Logger.log('================================================');
  Logger.log('validateTrabajosProcessing()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var ss = getTargetSS_();

  // 1. Hoja de respuestas del formulario
  Logger.log('\n[1] Hoja de respuestas del formulario de trabajos:');
  var respSheet = _findTrabajosResponseSheet_(ss);
  if (respSheet) {
    var numResp = Math.max(0, respSheet.getLastRow() - 1);
    Logger.log('  ✓ Encontrada: "' + respSheet.getName() + '"');
    Logger.log('  Respuestas: ' + numResp);
    if (respSheet.getName() !== 'RESP_TRABAJOS') {
      Logger.log('  ⚠️ El nombre debe ser RESP_TRABAJOS → ejecuta fixTrabajosResponseSheet()');
    }
    if (respSheet.getLastColumn() > 0) {
      var hdrs = respSheet.getRange(1, 1, 1, respSheet.getLastColumn()).getValues()[0]
        .map(function(h) { return String(h).trim(); });
      Logger.log('  Cabeceras: ' + hdrs.join(' | '));
    }
  } else {
    Logger.log('  ❌ No encontrada (ni RESP_TRABAJOS ni Form Responses N con campos de trabajo)');
  }

  // 2. Hoja TRABAJOS
  Logger.log('\n[2] Hoja TRABAJOS:');
  var nTrab = 0;
  try {
    var trabajosSheet = getSheet('TRABAJOS');
    nTrab = Math.max(0, trabajosSheet.getLastRow() - 1);
    Logger.log('  ✓ Existe. Filas de datos: ' + nTrab);
    if (nTrab > 0) {
      var lastId = trabajosSheet.getRange(trabajosSheet.getLastRow(), 1).getValue();
      Logger.log('  Último ID: ' + lastId);
    } else {
      Logger.log('  Sin datos — próximo ID será TRA-001');
    }
  } catch (e) {
    Logger.log('  ❌ ' + e.message);
  }

  // 3. Trigger onFormSubmit
  Logger.log('\n[3] Trigger onFormSubmitHandler:');
  var triggers     = ScriptApp.getProjectTriggers();
  var activosTrig  = triggers.filter(function(t) {
    return t.getHandlerFunction() === 'onFormSubmitHandler' &&
           t.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT;
  });
  if (activosTrig.length > 0) {
    Logger.log('  ✓ Activo (' + activosTrig.length + ' instancia(s))');
  } else {
    Logger.log('  ❌ No instalado → ejecuta setupTriggers()');
  }

  // 4. Resumen y acciones recomendadas
  var nResp      = respSheet ? Math.max(0, respSheet.getLastRow() - 1) : 0;
  var pendientes = nResp - nTrab;
  var todoOk     = true;

  Logger.log('\n================================================');
  Logger.log('RESUMEN:');
  Logger.log('  Respuestas en formulario:  ' + nResp);
  Logger.log('  Filas en TRABAJOS:         ' + nTrab);
  Logger.log('  Pendientes de procesar:    ' + Math.max(0, pendientes));
  Logger.log('\nACCIONES NECESARIAS (en orden):');

  if (!respSheet) {
    Logger.log('  1. Vincula el formulario al spreadsheet.');
    Logger.log('     Formulario → Respuestas → ⋮ → Seleccionar destino de respuestas.');
    todoOk = false;
  } else if (respSheet.getName() !== 'RESP_TRABAJOS') {
    Logger.log('  1. Ejecuta: fixTrabajosResponseSheet()');
    todoOk = false;
  }
  if (activosTrig.length === 0) {
    Logger.log('  2. Ejecuta: setupTriggers()');
    todoOk = false;
  }
  if (pendientes > 0) {
    Logger.log('  3. Ejecuta: processExistingTrabajoResponses()');
    todoOk = false;
  }
  if (todoOk) {
    Logger.log('  ✓ Todo correcto. El sistema está funcionando.');
  }
  Logger.log('================================================');
}

// ---- Funciones privadas de soporte ----

/**
 * Localiza la hoja de respuestas del formulario de trabajos en el spreadsheet.
 * Prioridad: RESP_TRABAJOS → Form Responses N con cabeceras de trabajos.
 */
function _findTrabajosResponseSheet_(ss) {
  var s = ss.getSheetByName('RESP_TRABAJOS');
  if (s) return s;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (/form.responses|respuestas.de.formulario/i.test(name) &&
        _sheetHasTrabajosHeaders_(sheets[i])) {
      return sheets[i];
    }
  }
  return null;
}

/**
 * Devuelve true si la hoja tiene al menos 2 cabeceras típicas del formulario de trabajos.
 */
function _sheetHasTrabajosHeaders_(sheet) {
  if (sheet.getLastColumn() < 1 || sheet.getLastRow() < 1) return false;
  var hs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function(h) { return String(h).toLowerCase().trim(); });
  var campos = ['cliente', 'operario', 'trabajo realizado', 'hora inicio'];
  return campos.filter(function(c) {
    return hs.some(function(h) { return h.indexOf(c) !== -1; });
  }).length >= 2;
}

/**
 * Convierte una fila de datos en namedValues compatible con onTrabajosFormSubmit.
 * Cada valor se envuelve en array para mantener la misma interfaz que e.namedValues.
 */
function _rowToNamedValues_(headers, row) {
  var nv = {};
  for (var i = 0; i < headers.length; i++) {
    nv[headers[i]] = [(row[i] !== undefined && row[i] !== null) ? String(row[i]) : ''];
  }
  return nv;
}

/**
 * Combina los valores de TODAS las claves de namedValues que contengan "foto".
 * Devuelve un string con todos los valores separados por coma.
 * Usado en onTrabajosFormSubmit y processExistingTrabajoResponses.
 */
function _combineFotos_(namedValues) {
  var parts = [];
  for (var key in namedValues) {
    if (key.toLowerCase().indexOf('foto') === -1) continue;
    var vals = namedValues[key] || [];
    for (var i = 0; i < vals.length; i++) {
      var v = String(vals[i]).trim();
      if (v) parts.push(v);
    }
  }
  return parts.join(', ');
}

/**
 * Extrae IDs de Drive del texto combinado de fotos y devuelve hasta 5 URLs limpias
 * en formato https://drive.google.com/file/d/ID/view
 */
function extractDrivePhotoLinks_(rawFotos) {
  if (!rawFotos) return [];
  var ids  = _extractDriveFileIds_(rawFotos);
  var urls = [];
  for (var i = 0; i < Math.min(ids.length, 5); i++) {
    urls.push('https://drive.google.com/file/d/' + ids[i] + '/view');
  }
  return urls;
}

/**
 * Asegura que TRABAJOS tiene Foto 1-5 y Carpeta Fotos en la cabecera.
 * Añade las que falten al final. No modifica filas de datos.
 */
function _ensureTrabajoPhotoColumns_(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); })
    : [];
  var needed = ['Foto 1', 'Foto 2', 'Foto 3', 'Foto 4', 'Foto 5', 'Carpeta Fotos'];
  for (var i = 0; i < needed.length; i++) {
    var exists = false;
    for (var j = 0; j < headers.length; j++) {
      if (headers[j] === needed[i]) { exists = true; break; }
    }
    if (!exists) {
      lastCol++;
      sheet.getRange(1, lastCol).setValue(needed[i]);
      headers.push(needed[i]);
    }
  }
}

/**
 * Asegura que TRABAJOS tiene la columna Materiales (col 19).
 * Añade la cabecera si falta. No modifica filas de datos.
 */
function _ensureTrabajoMaterialesColumn_(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); })
    : [];
  if (headers.indexOf('Materiales') === -1) {
    sheet.getRange(1, lastCol + 1).setValue('Materiales');
  }
}

/**
 * Asegura que TRABAJOS tiene las columnas del segundo empleado (cols 20-22):
 * 'Operario 2', 'Hora Inicio 2', 'Hora Fin 2'.
 * Añade las cabeceras que falten al final. No modifica filas de datos, de modo
 * que los trabajos antiguos quedan con esas celdas vacías (= sin empleado 2).
 */
function _ensureTrabajoEmpleado2Columns_(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); })
    : [];
  var needed = ['Operario 2', 'Hora Inicio 2', 'Hora Fin 2'];
  for (var i = 0; i < needed.length; i++) {
    if (headers.indexOf(needed[i]) === -1) {
      lastCol++;
      sheet.getRange(1, lastCol).setValue(needed[i]);
      headers.push(needed[i]);
    }
  }
}

/**
 * Valida que, para una pareja inicio/fin rellena, la hora fin sea posterior a la
 * de inicio. Si no lo es, registra un aviso en LOGS (no bloquea el alta: esas
 * horas computarán como 0, coherente con _calcHoras_). Las parejas vacías o
 * parciales se ignoran. `num` indica el empleado (1 o 2) para el mensaje.
 */
function _validarHorasTrabajo_(id, horaInicio, horaFin, num) {
  if (!horaInicio || !horaFin) return true;
  var horas = _calcHoras_(horaInicio, horaFin);
  if (horas <= 0) {
    logEvent('WARN', 'createTrabajo',
      'Trabajo ' + id + ': hora fin del empleado ' + num + ' no es posterior a la de inicio',
      'inicio=' + horaInicio + ' fin=' + horaFin);
    return false;
  }
  return true;
}

/**
 * Obtiene o crea la carpeta Drive: carpetaRaizId / año / workId.
 * Devuelve el Folder, o null si la raíz no es accesible.
 */
function _getOrCreateTrabajoFolder_(workId, year, carpetaRaizId) {
  if (!carpetaRaizId) return null;
  try {
    var raiz       = DriveApp.getFolderById(carpetaRaizId);
    var yearStr    = String(year);
    var yearIter   = raiz.getFoldersByName(yearStr);
    var yearFolder = yearIter.hasNext() ? yearIter.next() : raiz.createFolder(yearStr);
    var workIter   = yearFolder.getFoldersByName(workId);
    return workIter.hasNext() ? workIter.next() : yearFolder.createFolder(workId);
  } catch (e) {
    Logger.log('  ⚠️ _getOrCreateTrabajoFolder_: ' + e.message);
    return null;
  }
}

/**
 * Copia un archivo Drive a la carpeta del trabajo si no existe ya allí.
 * Devuelve la URL del archivo en la carpeta, o la URL /file/d/ID/view si falla.
 */
function _copiarFotoATrabajoFolder_(fileId, folder) {
  try {
    var file      = DriveApp.getFileById(fileId);
    var nombre    = file.getName();
    var existIter = folder.getFilesByName(nombre);
    if (existIter.hasNext()) return existIter.next().getUrl();
    return file.makeCopy(nombre, folder).getUrl();
  } catch (e) {
    Logger.log('  ⚠️ _copiarFotoATrabajoFolder_ ID=' + fileId + ': ' + e.message);
    return 'https://drive.google.com/file/d/' + fileId + '/view';
  }
}

/**
 * Procesa el texto de fotos para un trabajo:
 *   extrae IDs, crea carpeta carpeta_trabajos_id/año/workId, copia archivos.
 * Devuelve {foto1..foto5, carpetaFotos}. Nunca lanza excepciones.
 */
function _procesarFotosParaTrabajo_(rawFotos, workId) {
  var res = { foto1: '', foto2: '', foto3: '', foto4: '', foto5: '', carpetaFotos: '' };
  if (!rawFotos || !rawFotos.trim()) return res;

  var ids           = _extractDriveFileIds_(rawFotos);
  var carpetaRaizId = getConfig('carpeta_trabajos_id');
  var year          = new Date().getFullYear();
  var folder        = carpetaRaizId ? _getOrCreateTrabajoFolder_(workId, year, carpetaRaizId) : null;

  if (folder) res.carpetaFotos = folder.getUrl();

  var keys = ['foto1', 'foto2', 'foto3', 'foto4', 'foto5'];
  for (var i = 0; i < Math.min(ids.length, 5); i++) {
    res[keys[i]] = folder
      ? _copiarFotoATrabajoFolder_(ids[i], folder)
      : 'https://drive.google.com/file/d/' + ids[i] + '/view';
  }
  return res;
}

/**
 * Lee todas las columnas con "foto" en la cabecera de una fila y une sus valores.
 */
function _combineFotosDesdeRow_(headers, row) {
  var parts = [];
  for (var h = 0; h < headers.length; h++) {
    if (headers[h].toLowerCase().indexOf('foto') === -1) continue;
    var v = (row[h] !== undefined && row[h] !== null) ? String(row[h]).trim() : '';
    if (v) parts.push(v);
  }
  return parts.join(', ');
}

/**
 * Construye una clave de emparejamiento normalizada para un trabajo.
 * Combina: cliente | operario | fecha | hora inicio | hora fin | trabajo realizado.
 */
function _buildMatchKey_(cliente, operario, fecha, horaIni, horaFin, trabajo) {
  var tz = Session.getScriptTimeZone();

  var fechaStr = '';
  if (fecha instanceof Date && !isNaN(fecha)) {
    fechaStr = Utilities.formatDate(fecha, tz, 'yyyy-MM-dd');
  } else if (fecha) {
    var fd = new Date(String(fecha));
    fechaStr = !isNaN(fd.getTime())
      ? Utilities.formatDate(fd, tz, 'yyyy-MM-dd')
      : String(fecha).trim().substring(0, 10);
  }

  function _normHora_(v) {
    if (!v) return '';
    if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz, 'HH:mm');
    var s = String(v).trim();
    var m = s.match(/^(\d{1,2}):(\d{2})/);
    if (m) return ('0' + m[1]).slice(-2) + ':' + m[2];
    var d = new Date(s);
    return (!isNaN(d.getTime())) ? Utilities.formatDate(d, tz, 'HH:mm') : s.substring(0, 5);
  }

  return [
    String(cliente  || '').trim().toLowerCase(),
    String(operario || '').trim().toLowerCase(),
    fechaStr,
    _normHora_(horaIni),
    _normHora_(horaFin),
    String(trabajo  || '').trim().toLowerCase().substring(0, 100).replace(/\s+/g, ' ')
  ].join('||');
}

/**
 * Encuentra el índice base-0 de la primera cabecera que contiene algún término
 * de `include` y ninguno de `exclude`. Devuelve -1 si no se encuentra.
 */
function _findColIdx_(headers, include, exclude) {
  exclude = exclude || [];
  for (var j = 0; j < headers.length; j++) {
    var lh = headers[j].toLowerCase();
    var ok = false;
    for (var a = 0; a < include.length; a++) { if (lh.indexOf(include[a]) !== -1) { ok = true; break; } }
    if (!ok) continue;
    var bad = false;
    for (var b = 0; b < exclude.length; b++) { if (lh.indexOf(exclude[b]) !== -1) { bad = true; break; } }
    if (!bad) return j;
  }
  return -1;
}

// ---- Reparación de fotos en trabajos existentes ----

/**
 * Vacía Foto 1-5 y Carpeta Fotos en TRABAJOS para permitir un nuevo repair.
 * NO modifica la columna Fotos (resumen) ni ningún otro dato.
 * Ejecutar antes de repairExistingTrabajoPhotos() si las fotos estaban mal asignadas.
 */
function clearTrabajoPhotoColumns() {
  Logger.log('================================================');
  Logger.log('clearTrabajoPhotoColumns()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var sheet   = getSheet('TRABAJOS');
  _ensureTrabajoPhotoColumns_(sheet);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('Sin filas de datos. Nada que limpiar.');
    return;
  }

  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h).trim(); });

  var toClear   = ['Foto 1', 'Foto 2', 'Foto 3', 'Foto 4', 'Foto 5', 'Carpeta Fotos'];
  var limpiadas = 0;

  for (var c = 0; c < headers.length; c++) {
    if (toClear.indexOf(headers[c]) !== -1) {
      sheet.getRange(2, c + 1, lastRow - 1, 1).clearContent();
      Logger.log('  "' + headers[c] + '" (col ' + (c + 1) + '): ' + (lastRow - 1) + ' celdas limpiadas');
      limpiadas++;
    }
  }

  Logger.log('\nTotal columnas limpiadas: ' + limpiadas + '/' + toClear.length);
  Logger.log('Columna Fotos (resumen) NO modificada.');
  Logger.log('================================================');
}

/**
 * Empareja cada trabajo de TRABAJOS con su respuesta en RESP_TRABAJOS usando
 * cliente + operario + fecha + hora inicio + hora fin + trabajo realizado.
 * Si hay más de una coincidencia, NO asigna fotos y deja log de conflicto.
 * Copia las fotos a carpeta_trabajos_id/año/ID_TRABAJO.
 * NO duplica archivos. NO borra datos.
 * Ejecutar DESPUÉS de clearTrabajoPhotoColumns() si se requiere reasignación.
 */
function repairExistingTrabajoPhotos() {
  Logger.log('================================================');
  Logger.log('repairExistingTrabajoPhotos()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var ss = getTargetSS_();

  // [1] RESP_TRABAJOS — construir mapa clave → filas
  Logger.log('\n[1] Hoja RESP_TRABAJOS:');
  var respSheet = ss.getSheetByName('RESP_TRABAJOS');
  if (!respSheet) { Logger.log('  ❌ No encontrada. Abortando.'); return; }
  var respLastRow = respSheet.getLastRow();
  if (respLastRow < 2) { Logger.log('  Sin respuestas. Nada que reparar.'); return; }

  var respData    = respSheet.getDataRange().getValues();
  var respHeaders = respData[0].map(function(h) { return String(h).trim(); });
  Logger.log('  Filas de datos: ' + (respLastRow - 1));

  var rcCliente  = _findColIdx_(respHeaders, ['cliente']);
  var rcOperario = _findColIdx_(respHeaders, ['operario']);
  var rcFecha    = _findColIdx_(respHeaders, ['fecha'], ['marca temporal', 'timestamp']);
  var rcHoraIni  = _findColIdx_(respHeaders, ['hora inicio', 'hora de inicio']);
  var rcHoraFin  = _findColIdx_(respHeaders, ['hora fin', 'hora de fin']);
  var rcTrabajo  = _findColIdx_(respHeaders, ['trabajo realizado']);

  Logger.log('  Columnas: cliente=' + (rcCliente+1) + ' operario=' + (rcOperario+1) +
             ' fecha=' + (rcFecha+1) + ' horaIni=' + (rcHoraIni+1) +
             ' horaFin=' + (rcHoraFin+1) + ' trabajo=' + (rcTrabajo+1));

  var respKeyMap = {};
  for (var ri = 1; ri < respData.length; ri++) {
    var rrow = respData[ri];
    var key  = _buildMatchKey_(
      rcCliente  >= 0 ? rrow[rcCliente]  : '',
      rcOperario >= 0 ? rrow[rcOperario] : '',
      rcFecha    >= 0 ? rrow[rcFecha]    : '',
      rcHoraIni  >= 0 ? rrow[rcHoraIni]  : '',
      rcHoraFin  >= 0 ? rrow[rcHoraFin]  : '',
      rcTrabajo  >= 0 ? rrow[rcTrabajo]  : ''
    );
    if (!respKeyMap[key]) respKeyMap[key] = [];
    respKeyMap[key].push(ri);
  }
  Logger.log('  Claves únicas: ' + Object.keys(respKeyMap).length);

  // [2] TRABAJOS
  Logger.log('\n[2] Hoja TRABAJOS:');
  var trabSheet = getSheet('TRABAJOS');
  _ensureTrabajoPhotoColumns_(trabSheet);

  var trabLastCol = trabSheet.getLastColumn();
  var trabLastRow = trabSheet.getLastRow();
  if (trabLastRow < 2) { Logger.log('  Sin trabajos. Nada que reparar.'); return; }

  var trabHeaders = trabSheet.getRange(1, 1, 1, trabLastCol).getValues()[0]
    .map(function(h) { return String(h).trim(); });

  var fotoNombres = ['Foto 1', 'Foto 2', 'Foto 3', 'Foto 4', 'Foto 5', 'Carpeta Fotos'];
  var fotoColMap  = {};
  for (var k = 0; k < fotoNombres.length; k++) {
    for (var c = 0; c < trabHeaders.length; c++) {
      if (trabHeaders[c] === fotoNombres[k]) { fotoColMap[fotoNombres[k]] = c; break; }
    }
  }
  var trabData = trabSheet.getRange(2, 1, trabLastRow - 1, trabLastCol).getValues();
  Logger.log('  Filas de datos: ' + trabData.length);

  // [3] Emparejar y reparar
  Logger.log('\n[3] Emparejando por campos clave...');
  var reparados = 0;
  var yaTienen  = 0;
  var sinMatch  = 0;
  var conflicto = 0;
  var sinFotos  = 0;
  var errores   = 0;

  for (var r = 0; r < trabData.length; r++) {
    var trabRow    = trabData[r];
    var workId     = String(trabRow[0] || '').trim();
    var foto1Idx   = fotoColMap['Foto 1'];
    var carpetaIdx = fotoColMap['Carpeta Fotos'];
    var foto1Val   = foto1Idx   !== undefined ? String(trabRow[foto1Idx]   || '').trim() : '';
    var carpetaVal = carpetaIdx !== undefined ? String(trabRow[carpetaIdx] || '').trim() : '';

    Logger.log('\n  ' + workId + ':');

    if (foto1Val && carpetaVal) {
      Logger.log('    ✓ Ya tiene Foto 1 y Carpeta Fotos. Saltando.');
      yaTienen++;
      continue;
    }

    var trabKey = _buildMatchKey_(
      trabRow[COL_TRA.CLIENTE           - 1],
      trabRow[COL_TRA.OPERARIO          - 1],
      trabRow[COL_TRA.FECHA_TRABAJO     - 1],
      trabRow[COL_TRA.HORA_INICIO       - 1],
      trabRow[COL_TRA.HORA_FIN          - 1],
      trabRow[COL_TRA.TRABAJO_REALIZADO - 1]
    );

    var matches = respKeyMap[trabKey] || [];
    Logger.log('    Coincidencias RESP_TRABAJOS: ' + matches.length);

    if (matches.length === 0) {
      Logger.log('    ⚠️ Sin coincidencia — revisar datos del trabajo.');
      sinMatch++;
      continue;
    }
    if (matches.length > 1) {
      Logger.log('    ⚠️ CONFLICTO: ' + matches.length + ' filas coinciden (RESP filas: ' +
                 matches.join(', ') + '). Fotos NO asignadas.');
      conflicto++;
      continue;
    }

    // Única coincidencia — extraer fotos
    var rawFotos    = _combineFotosDesdeRow_(respHeaders, respData[matches[0]]);
    var fotosActual = String(trabRow[COL_TRA.FOTOS - 1] || '').trim();
    if (fotosActual && rawFotos.indexOf(fotosActual) === -1) {
      rawFotos = rawFotos ? rawFotos + ', ' + fotosActual : fotosActual;
    }
    Logger.log('    Raw fotos: ' + (rawFotos || '(vacío)'));

    if (!rawFotos) {
      Logger.log('    ⚠️ Sin fotos en la respuesta correspondiente.');
      sinFotos++;
      continue;
    }

    var fechaVal      = trabRow[COL_TRA.FECHA_TRABAJO - 1];
    var year          = (fechaVal instanceof Date && !isNaN(fechaVal)) ? fechaVal.getFullYear() : new Date().getFullYear();
    var ids           = _extractDriveFileIds_(rawFotos);
    var carpetaRaizId = getConfig('carpeta_trabajos_id');
    var folder        = carpetaRaizId ? _getOrCreateTrabajoFolder_(workId, year, carpetaRaizId) : null;
    var carpetaUrl    = folder ? folder.getUrl() : '';
    var fotoUrls      = [];

    for (var fi = 0; fi < Math.min(ids.length, 5); fi++) {
      fotoUrls.push(folder
        ? _copiarFotoATrabajoFolder_(ids[fi], folder)
        : 'https://drive.google.com/file/d/' + ids[fi] + '/view');
    }
    Logger.log('    Fotos: ' + fotoUrls.length + '  Carpeta: ' + (carpetaUrl || '(sin carpeta raíz)'));

    try {
      var sheetRow = r + 2;
      var fotoKeys = ['Foto 1', 'Foto 2', 'Foto 3', 'Foto 4', 'Foto 5'];
      for (var fi2 = 0; fi2 < fotoKeys.length; fi2++) {
        if (fotoColMap[fotoKeys[fi2]] !== undefined) {
          trabSheet.getRange(sheetRow, fotoColMap[fotoKeys[fi2]] + 1)
            .setValue(fi2 < fotoUrls.length ? fotoUrls[fi2] : '');
        }
      }
      if (carpetaIdx !== undefined) trabSheet.getRange(sheetRow, carpetaIdx + 1).setValue(carpetaUrl);
      if (!fotosActual && rawFotos) trabSheet.getRange(sheetRow, COL_TRA.FOTOS).setValue(rawFotos);
      Logger.log('    ✓ Reparado.');
      reparados++;
    } catch (e) {
      Logger.log('    ❌ Error al escribir: ' + e.message);
      errores++;
    }
  }

  Logger.log('\n================================================');
  Logger.log('RESULTADO:');
  Logger.log('  Reparados:        ' + reparados);
  Logger.log('  Ya tenían fotos:  ' + yaTienen);
  Logger.log('  Sin coincidencia: ' + sinMatch);
  Logger.log('  Conflictos:       ' + conflicto);
  Logger.log('  Sin fotos:        ' + sinFotos);
  Logger.log('  Errores:          ' + errores);
  if (conflicto > 0) Logger.log('  ⚠️ Conflictos: revisar manualmente en RESP_TRABAJOS.');
  if (errores   > 0) Logger.log('  ⚠️ Errores: revisa permisos Drive y carpeta_trabajos_id.');
  Logger.log('================================================');
}

// ---- Campo Materiales en formulario existente ----

/**
 * Añade el campo "Materiales (€)" al formulario de trabajos ya existente.
 * Lo coloca después de "Fotos" y antes de "Ubicación".
 * Si el campo ya existe, no hace nada.
 * NO recrea el formulario. NO borra respuestas.
 */
function addMaterialesFieldToTrabajosForm() {
  Logger.log('================================================');
  Logger.log('addMaterialesFieldToTrabajosForm()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var formId = getConfig('form_trabajos_id');
  Logger.log('\n[1] form_trabajos_id: ' + (formId || '❌ NO CONFIGURADO'));
  if (!formId) {
    Logger.log('ABORTANDO. Ejecuta installTrabajosModule() primero.');
    return;
  }

  var form;
  try {
    form = FormApp.openById(formId);
    Logger.log('  ✓ Formulario: "' + form.getTitle() + '"');
  } catch (e) {
    Logger.log('  ❌ No se puede abrir: ' + e.message);
    return;
  }

  Logger.log('\n[2] Campos actuales:');
  var items           = form.getItems();
  var materialesIdx   = -1;
  var ubicacionIdx    = -1;
  var fotosIdx        = -1;

  for (var i = 0; i < items.length; i++) {
    var titulo = items[i].getTitle();
    Logger.log('  ' + (i + 1) + '. [' + items[i].getType() + '] "' + titulo + '"');
    if (titulo === 'Materiales (€)')                           materialesIdx = i;
    if (titulo === 'Ubicación')                                ubicacionIdx  = i;
    if (titulo === 'Fotos' || titulo === 'Fotos (enlaces Drive)') fotosIdx   = i;
  }

  if (materialesIdx !== -1) {
    Logger.log('\n✓ El campo "Materiales (€)" ya existe en posición ' + (materialesIdx + 1) + '. Nada que hacer.');
    return;
  }

  // Añadir al final y luego mover a la posición correcta
  Logger.log('\n[3] Añadiendo campo "Materiales (€)"...');
  form.addTextItem()
    .setTitle('Materiales (€)')
    .setHelpText('Coste de materiales empleados en este trabajo (ej: 25.50). Dejar vacío si no hay.')
    .setRequired(false);

  var allItems = form.getItems();
  var newPos   = allItems.length - 1;  // índice base-0 del nuevo campo (está al final)

  // Posición objetivo: antes de Ubicación si existe, o justo después de Fotos
  var targetPos;
  if (ubicacionIdx !== -1) {
    targetPos = ubicacionIdx;  // inserta antes de Ubicación
  } else if (fotosIdx !== -1) {
    targetPos = fotosIdx + 1;  // inserta justo después de Fotos
  } else {
    targetPos = newPos;        // se queda al final
  }

  if (targetPos !== newPos) {
    form.moveItem(newPos, targetPos);
    Logger.log('  ✓ Campo movido a posición ' + (targetPos + 1) + '.');
  } else {
    Logger.log('  ✓ Campo añadido al final (Ubicación no encontrada).');
  }

  Logger.log('\n[4] Campos finales:');
  var finalItems = form.getItems();
  for (var j = 0; j < finalItems.length; j++) {
    Logger.log('  ' + (j + 1) + '. [' + finalItems[j].getType() + '] "' + finalItems[j].getTitle() + '"');
  }

  Logger.log('\n================================================');
  Logger.log('CAMPO AÑADIDO CORRECTAMENTE.');
  Logger.log('  Nombre:      Materiales (€)');
  Logger.log('  Tipo:        TEXT (texto corto)');
  Logger.log('  Obligatorio: No');
  Logger.log('  Backend:     onTrabajosFormSubmit y createTrabajo ya lo leen.');
  Logger.log('  Vacío → 0:   sí, el handler normaliza el valor.');
  Logger.log('  URL: ' + form.getPublishedUrl());
  Logger.log('================================================');
}

// ---- Campos del segundo empleado en formulario existente ----

/**
 * Añade al formulario de trabajos ya existente los campos del segundo empleado
 * (opcional): "Operario 2" (lista), "Hora inicio 2" y "Hora fin 2" (hora).
 * Los coloca juntos, justo después de "Hora fin". Si ya existen, no hace nada.
 * NO recrea el formulario. NO borra respuestas. El backend (onTrabajosFormSubmit,
 * processExistingTrabajoResponses y createTrabajo) ya los lee.
 */
function addEmpleado2FieldsToTrabajosForm() {
  Logger.log('================================================');
  Logger.log('addEmpleado2FieldsToTrabajosForm()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var formId = getConfig('form_trabajos_id');
  Logger.log('\n[1] form_trabajos_id: ' + (formId || '❌ NO CONFIGURADO'));
  if (!formId) {
    Logger.log('ABORTANDO. Ejecuta installTrabajosModule() primero.');
    return;
  }

  var form;
  try {
    form = FormApp.openById(formId);
    Logger.log('  ✓ Formulario: "' + form.getTitle() + '"');
  } catch (e) {
    Logger.log('  ❌ No se puede abrir: ' + e.message);
    return;
  }

  // Localizar "Hora fin" (del empleado 1) y comprobar qué campos ya existen
  var items      = form.getItems();
  var horaFinIdx = -1;
  var existentes = {};
  for (var i = 0; i < items.length; i++) {
    var t = items[i].getTitle();
    if (t === 'Hora fin') horaFinIdx = i;
    if (t === 'Operario 2' || t === 'Hora inicio 2' || t === 'Hora fin 2') existentes[t] = true;
  }

  // Lista de operarios para el desplegable (misma fuente que el formulario original)
  var operariosList = _getTrabajosDropdownOperarios_();

  // Insertar en orden inverso justo después de "Hora fin" para que queden ordenados:
  // Operario 2, Hora inicio 2, Hora fin 2.
  var insertBase = (horaFinIdx !== -1) ? horaFinIdx + 1 : form.getItems().length;
  var añadidos   = 0;

  if (!existentes['Hora fin 2']) {
    form.addTimeItem().setTitle('Hora fin 2').setRequired(false);
    form.moveItem(form.getItems().length - 1, insertBase);
    añadidos++;
  }
  if (!existentes['Hora inicio 2']) {
    form.addTimeItem().setTitle('Hora inicio 2').setRequired(false);
    form.moveItem(form.getItems().length - 1, insertBase);
    añadidos++;
  }
  if (!existentes['Operario 2']) {
    var op2 = form.addListItem()
      .setTitle('Operario 2')
      .setHelpText('Solo si un segundo operario participó en este trabajo.')
      .setRequired(false);
    if (operariosList.length) op2.setChoiceValues(operariosList);
    form.moveItem(form.getItems().length - 1, insertBase);
    añadidos++;
  }

  Logger.log('\n[2] Campos añadidos: ' + añadidos + (añadidos === 0 ? ' (ya existían todos)' : ''));
  Logger.log('\n[3] Campos finales:');
  var finalItems = form.getItems();
  for (var j = 0; j < finalItems.length; j++) {
    Logger.log('  ' + (j + 1) + '. [' + finalItems[j].getType() + '] "' + finalItems[j].getTitle() + '"');
  }
  Logger.log('\n================================================');
  Logger.log('LISTO. URL: ' + form.getPublishedUrl());
  Logger.log('================================================');
}

// ---- Dropdowns del formulario ----

/**
 * Actualiza los desplegables Cliente y Operario del formulario de trabajos
 * con los datos actuales de las hojas CLIENTES y OPERARIOS.
 *
 * Ejecutar cada vez que se añadan clientes u operarios nuevos.
 * NO recrea el formulario. NO borra respuestas existentes.
 */
function refreshTrabajosDropdowns() {
  Logger.log('================================================');
  Logger.log('refreshTrabajosDropdowns()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  // 1. Obtener ID del formulario
  var formId = getConfig('form_trabajos_id');
  Logger.log('\n[1] form_trabajos_id: ' + (formId || '❌ NO CONFIGURADO'));
  if (!formId) {
    Logger.log('ABORTANDO. Ejecuta installTrabajosModule() primero.');
    return;
  }

  // 2. Leer CLIENTES activos
  Logger.log('\n[2] Leyendo CLIENTES activos:');
  var clientesList = _getTrabajosDropdownClientes_();
  Logger.log('  Encontrados: ' + clientesList.length);
  for (var c = 0; c < clientesList.length; c++) Logger.log('  - ' + clientesList[c]);
  if (clientesList.length === 0) {
    Logger.log('  ⚠️ Sin clientes activos — el desplegable quedará vacío.');
  }

  // 3. Leer OPERARIOS activos
  Logger.log('\n[3] Leyendo OPERARIOS activos:');
  var operariosList = _getTrabajosDropdownOperarios_();
  Logger.log('  Encontrados: ' + operariosList.length);
  for (var o = 0; o < operariosList.length; o++) Logger.log('  - ' + operariosList[o]);
  if (operariosList.length === 0) {
    Logger.log('  ⚠️ Sin operarios activos — el desplegable quedará vacío.');
  }

  // Si ambas listas están vacías no tiene sentido continuar
  if (clientesList.length === 0 && operariosList.length === 0) {
    Logger.log('\nABORTANDO. Añade clientes y operarios antes de actualizar el formulario.');
    return;
  }

  // 4. Abrir formulario y actualizar items
  Logger.log('\n[4] Abriendo formulario...');
  var form;
  try {
    form = FormApp.openById(formId);
    Logger.log('  ✓ "' + form.getTitle() + '"');
  } catch (e) {
    Logger.log('  ❌ No se puede abrir el formulario: ' + e.message);
    return;
  }

  var items           = form.getItems();
  var updatedClientes  = false;
  var updatedOperarios = false;

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    if (item.getType() !== FormApp.ItemType.LIST) continue;

    var listItem = item.asListItem();
    var titulo   = item.getTitle();

    if (titulo === 'Cliente' && clientesList.length > 0) {
      var uniqueClientes = [...new Set(clientesList)];
      Logger.log('  → setChoiceValues "Cliente": ' + uniqueClientes.length + ' opciones — ' + JSON.stringify(uniqueClientes));
      listItem.setChoiceValues(uniqueClientes);
      updatedClientes = true;
      Logger.log('  ✓ Desplegable "Cliente" actualizado');
    } else if ((titulo === 'Operario' || titulo === 'Operario 2') && operariosList.length > 0) {
      var uniqueOperarios = [...new Set(operariosList)];
      Logger.log('  → setChoiceValues "' + titulo + '": ' + uniqueOperarios.length + ' opciones — ' + JSON.stringify(uniqueOperarios));
      listItem.setChoiceValues(uniqueOperarios);
      updatedOperarios = true;
      Logger.log('  ✓ Desplegable "' + titulo + '" actualizado');
    }
  }

  if (!updatedClientes)  Logger.log('  ⚠️ Item "Cliente" no encontrado en el formulario.');
  if (!updatedOperarios) Logger.log('  ⚠️ Item "Operario" no encontrado en el formulario.');

  Logger.log('\n================================================');
  Logger.log('RESULTADO:');
  Logger.log('  Cliente:  ' + (updatedClientes  ? '✓ actualizado' : '⚠️ no actualizado'));
  Logger.log('  Operario: ' + (updatedOperarios ? '✓ actualizado' : '⚠️ no actualizado'));
  Logger.log('  URL formulario: ' + form.getPublishedUrl());
  Logger.log('================================================');
}

/**
 * Muestra el estado actual de los desplegables Cliente y Operario
 * del formulario de trabajos, y cuántos registros hay en las hojas.
 */
function validateTrabajosDropdowns() {
  Logger.log('================================================');
  Logger.log('validateTrabajosDropdowns()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  // 1. Clientes en la hoja
  Logger.log('\n[1] Hoja CLIENTES:');
  var clientesList = _getTrabajosDropdownClientes_();
  Logger.log('  Activos disponibles: ' + clientesList.length);
  for (var c = 0; c < clientesList.length; c++) Logger.log('  - ' + clientesList[c]);

  // 2. Operarios en la hoja
  Logger.log('\n[2] Hoja OPERARIOS:');
  var operariosList = _getTrabajosDropdownOperarios_();
  Logger.log('  Activos disponibles: ' + operariosList.length);
  for (var o = 0; o < operariosList.length; o++) Logger.log('  - ' + operariosList[o]);

  // 3. Opciones actuales en el formulario
  Logger.log('\n[3] Opciones actuales en el formulario:');
  var formId = getConfig('form_trabajos_id');
  if (!formId) {
    Logger.log('  ❌ form_trabajos_id no configurado.');
    return;
  }

  var form;
  try {
    form = FormApp.openById(formId);
    Logger.log('  Formulario: "' + form.getTitle() + '"');
  } catch (e) {
    Logger.log('  ❌ No se puede abrir el formulario: ' + e.message);
    return;
  }

  var items = form.getItems();
  var foundClientes  = false;
  var foundOperarios = false;

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    if (item.getType() !== FormApp.ItemType.LIST) continue;

    var titulo   = item.getTitle();
    var choices  = item.asListItem().getChoices();
    var valores  = choices.map(function(ch) { return ch.getValue(); });

    if (titulo === 'Cliente') {
      foundClientes = true;
      Logger.log('\n  Desplegable "Cliente" — ' + valores.length + ' opciones actuales:');
      for (var a = 0; a < valores.length; a++) Logger.log('    · ' + valores[a]);
    } else if (titulo === 'Operario') {
      foundOperarios = true;
      Logger.log('\n  Desplegable "Operario" — ' + valores.length + ' opciones actuales:');
      for (var b = 0; b < valores.length; b++) Logger.log('    · ' + valores[b]);
    }
  }

  if (!foundClientes)  Logger.log('  ⚠️ Item "Cliente" no encontrado en el formulario.');
  if (!foundOperarios) Logger.log('  ⚠️ Item "Operario" no encontrado en el formulario.');

  Logger.log('\n================================================');
  Logger.log('DIAGNÓSTICO:');
  Logger.log('  Clientes en hoja:    ' + clientesList.length);
  Logger.log('  Operarios en hoja:   ' + operariosList.length);
  if (clientesList.length > 0 || operariosList.length > 0) {
    Logger.log('  → Ejecuta refreshTrabajosDropdowns() para sincronizar el formulario.');
  }
  Logger.log('================================================');
}

// ---- Campo Fotos ----

/**
 * Actualiza el campo "Fotos" del formulario de trabajos.
 *
 * Google Forms no permite subida de archivos sin login obligatorio,
 * por lo que se mantiene como campo de texto con instrucciones claras
 * para que el operario suba las fotos a Drive y pegue los enlaces.
 *
 * NO recrea el formulario. NO borra respuestas existentes.
 */
function convertFotosToFileUpload() {
  Logger.log('================================================');
  Logger.log('convertFotosToFileUpload()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var formId = getConfig('form_trabajos_id');
  Logger.log('\n[1] form_trabajos_id: ' + (formId || '❌ NO CONFIGURADO'));
  if (!formId) {
    Logger.log('ABORTANDO. Ejecuta installTrabajosModule() primero.');
    return;
  }

  var form;
  try {
    form = FormApp.openById(formId);
    Logger.log('  ✓ Formulario abierto: "' + form.getTitle() + '"');
  } catch (e) {
    Logger.log('  ❌ No se puede abrir el formulario: ' + e.message);
    return;
  }

  // Buscar el campo Fotos actual
  Logger.log('\n[2] Buscando campo "Fotos"...');
  var items     = form.getItems();
  var fotosItem = null;
  var fotosIdx  = -1;

  for (var i = 0; i < items.length; i++) {
    if (items[i].getTitle() === 'Fotos') {
      fotosItem = items[i];
      fotosIdx  = i;
      break;
    }
  }

  if (!fotosItem) {
    Logger.log('  ❌ Campo "Fotos" no encontrado en el formulario.');
    return;
  }

  Logger.log('  Encontrado en posición ' + (fotosIdx + 1) + ' — tipo: ' + fotosItem.getType());

  // Actualizar título y texto de ayuda del campo existente
  Logger.log('\n[3] Actualizando campo "Fotos"...');

  var HELP_TEXT =
    'Sube las fotos a Google Drive y comparte el enlace de cada una.\n' +
    'Pega los enlaces separados por coma o uno por línea.\n' +
    'Máximo recomendado: 5 fotos.';

  var tipo = fotosItem.getType();

  if (tipo === FormApp.ItemType.PARAGRAPH_TEXT) {
    var pItem = fotosItem.asParagraphTextItem();
    pItem.setTitle('Fotos (enlaces Drive)');
    pItem.setHelpText(HELP_TEXT);
    pItem.setRequired(false);
    Logger.log('  ✓ Texto de ayuda actualizado en campo de texto largo.');
  } else if (tipo === FormApp.ItemType.TEXT) {
    var tItem = fotosItem.asTextItem();
    tItem.setTitle('Fotos (enlaces Drive)');
    tItem.setHelpText(HELP_TEXT);
    tItem.setRequired(false);
    Logger.log('  ✓ Texto de ayuda actualizado en campo de texto.');
  } else {
    Logger.log('  ⚠️ El campo "Fotos" es de tipo ' + tipo + ' — no se modificó.');
  }

  // Estado final de todos los campos
  Logger.log('\n[4] Campos del formulario tras el cambio:');
  var finalItems = form.getItems();
  for (var j = 0; j < finalItems.length; j++) {
    Logger.log('  ' + (j + 1) + '. [' + finalItems[j].getType() + '] ' + finalItems[j].getTitle());
  }

  Logger.log('\n================================================');
  Logger.log('CAMPO FOTOS ACTUALIZADO.');
  Logger.log('NOTA: la subida de archivos directa en Google Forms requiere');
  Logger.log('login obligatorio. Se mantiene campo de texto con instrucciones.');
  Logger.log('Los operarios deben subir fotos a Drive y pegar los enlaces.');
  Logger.log('Ejecuta validateFotosUploadField() para confirmar.');
  Logger.log('================================================');
}

/**
 * Valida el estado actual del campo Fotos en el formulario de trabajos.
 * Muestra tipo, título, texto de ayuda y posición.
 */
function validateFotosUploadField() {
  Logger.log('================================================');
  Logger.log('validateFotosUploadField()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var formId = getConfig('form_trabajos_id');
  Logger.log('\n[1] form_trabajos_id: ' + (formId || '❌ NO CONFIGURADO'));
  if (!formId) return;

  var form;
  try {
    form = FormApp.openById(formId);
    Logger.log('  ✓ "' + form.getTitle() + '"');
    Logger.log('  URL: ' + form.getPublishedUrl());
  } catch (e) {
    Logger.log('  ❌ ' + e.message);
    return;
  }

  Logger.log('\n[2] Todos los campos del formulario:');
  var items    = form.getItems();
  var fotoItem = null;

  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    Logger.log('  ' + (i + 1) + '. [' + it.getType() + '] ' + it.getTitle());
    if (it.getTitle() === 'Fotos' || it.getTitle() === 'Fotos (enlaces Drive)') {
      fotoItem = it;
    }
  }

  Logger.log('\n[3] Campo Fotos:');
  if (!fotoItem) {
    Logger.log('  ❌ No encontrado.');
  } else {
    Logger.log('  Título: "' + fotoItem.getTitle() + '"');
    Logger.log('  Tipo:   ' + fotoItem.getType());
    var helpText = '';
    var tipo = fotoItem.getType();
    if (tipo === FormApp.ItemType.PARAGRAPH_TEXT) {
      helpText = fotoItem.asParagraphTextItem().getHelpText();
    } else if (tipo === FormApp.ItemType.TEXT) {
      helpText = fotoItem.asTextItem().getHelpText();
    }
    Logger.log('  Ayuda:  "' + (helpText || '(sin texto de ayuda)') + '"');
    Logger.log('  Tipo FILE_UPLOAD nativo: NO — requiere login obligatorio en Google Forms.');
    Logger.log('  Solución activa: campo de texto con instrucciones de enlace Drive.');
  }

  Logger.log('\n================================================');
  Logger.log('RESULTADO: ' + (fotoItem ? '✓ Campo Fotos presente y configurado.' : '❌ Campo Fotos no encontrado.'));
  Logger.log('================================================');
}

// ---- Subida nativa de fotos (FILE_UPLOAD) ----

/**
 * Convierte el campo Fotos del formulario de trabajos al tipo nativo FILE_UPLOAD.
 *
 * Requisito de plataforma: activa setRequireLogin(true) en el formulario.
 * Los operarios deben tener cuenta Google para enviar el formulario.
 *
 * El campo se configura para:
 *   - Solo imágenes (FormApp.FileType.IMAGE)
 *   - Hasta 5 archivos por respuesta
 *   - No obligatorio
 *
 * Los archivos se guardan automáticamente en Drive del propietario del formulario.
 * Los URLs de Drive llegan a RESP_TRABAJOS y se propagan a TRABAJOS via trigger.
 *
 * NO recrea el formulario. NO borra respuestas existentes.
 */
function enableNativePhotoUpload() {
  Logger.log('================================================');
  Logger.log('enableNativePhotoUpload()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var formId = getConfig('form_trabajos_id');
  Logger.log('\n[1] form_trabajos_id: ' + (formId || '❌ NO CONFIGURADO'));
  if (!formId) {
    Logger.log('ABORTANDO. Ejecuta installTrabajosModule() primero.');
    return;
  }

  var form;
  try {
    form = FormApp.openById(formId);
    Logger.log('  ✓ Formulario abierto: "' + form.getTitle() + '"');
  } catch (e) {
    Logger.log('  ❌ No se puede abrir: ' + e.message);
    return;
  }

  // 2. Login obligatorio — debe activarse manualmente en Form > Configuración
  Logger.log('\n[2] Login obligatorio:');
  Logger.log('  ⚠️ setRequireLogin() no está soportado via API en este proyecto.');
  Logger.log('  ACCIÓN MANUAL requerida ANTES de usar el formulario:');
  Logger.log('  Form → ⚙️ Configuración → "Limitar a 1 respuesta" O activar');
  Logger.log('  "Requiere inicio de sesión en Google".');

  // 3. Estado actual del formulario
  Logger.log('\n[3] Campos actuales del formulario:');
  var items      = form.getItems();
  var existente  = null;
  var existIdx   = -1;
  var ubicIdx    = -1;

  for (var i = 0; i < items.length; i++) {
    var t = items[i].getTitle();
    Logger.log('  ' + (i + 1) + '. [' + items[i].getType() + '] "' + t + '"');
    if (t.toLowerCase().indexOf('foto') !== -1) { existente = items[i]; existIdx = i; }
    if (t === 'Ubicación') { ubicIdx = i; }
  }

  // 4. Eliminar campo Fotos si ya existe (para no duplicar)
  if (existente) {
    Logger.log('\n[4] Eliminando campo Fotos existente ("' + existente.getTitle() + '")...');
    form.deleteItem(existente);
    Logger.log('  ✓ Eliminado.');
    if (ubicIdx > existIdx) ubicIdx--;  // ajustar índice tras borrado
  } else {
    Logger.log('\n[4] Campo Fotos no existe — se creará nuevo.');
  }

  // 5. Crear campo FILE_UPLOAD
  Logger.log('\n[5] Creando campo FILE_UPLOAD "Fotos"...');
  form.addFileUploadItem()
    .setTitle('Fotos')
    .setAllowedFileTypes([FormApp.FileType.IMAGE])
    .setMaxFileCount(5)
    .setRequired(false);
  Logger.log('  ✓ Campo creado (tipos: IMAGE, máx: 5).');

  // 6. Mover antes de Ubicación (o a posición 9 si no existe Ubicación)
  Logger.log('\n[6] Reposicionando campo...');
  var allItems = form.getItems();
  var lastIdx  = allItems.length - 1;
  var targetIdx = ubicIdx >= 0 ? ubicIdx : Math.min(8, lastIdx);
  if (lastIdx !== targetIdx) {
    form.moveItem(lastIdx, targetIdx);
    Logger.log('  ✓ Movido a posición ' + (targetIdx + 1) + ' (antes de Ubicación).');
  } else {
    Logger.log('  ℹ️ Ya en posición correcta (' + (targetIdx + 1) + ').');
  }

  // 7. Estado final
  Logger.log('\n[7] Campos del formulario tras el cambio:');
  var finalItems = form.getItems();
  for (var j = 0; j < finalItems.length; j++) {
    Logger.log('  ' + (j + 1) + '. [' + finalItems[j].getType() + '] ' + finalItems[j].getTitle());
  }

  Logger.log('\n================================================');
  Logger.log('CAMPO FOTOS → FILE_UPLOAD ACTIVADO.');
  Logger.log('  Login requerido: SÍ (cuenta Google).');
  Logger.log('  Tipos: solo imágenes.');
  Logger.log('  Máximo: 5 fotos por trabajo.');
  Logger.log('  Drive: Google guarda los archivos automáticamente.');
  Logger.log('  RESP_TRABAJOS y TRABAJOS reciben los URLs de Drive.');
  Logger.log('Ejecuta validateNativePhotoUpload() para confirmar.');
  Logger.log('================================================');
}

/**
 * Valida que el campo Fotos del formulario es de tipo FILE_UPLOAD
 * y está correctamente configurado.
 */
function validateNativePhotoUpload() {
  Logger.log('================================================');
  Logger.log('validateNativePhotoUpload()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var formId = getConfig('form_trabajos_id');
  Logger.log('\n[1] form_trabajos_id: ' + (formId || '❌ NO CONFIGURADO'));
  if (!formId) return;

  var form;
  try {
    form = FormApp.openById(formId);
    Logger.log('  ✓ "' + form.getTitle() + '"');
    Logger.log('  URL: ' + form.getPublishedUrl());
    try {
      Logger.log('  Login requerido: ' + (form.requiresLogin() ? '✓ SÍ' : '⚠️ NO — actívalo manualmente en Form > Configuración'));
    } catch (_) {
      Logger.log('  Login requerido: (no comprobable via API — verificar manualmente en Form > Configuración)');
    }
  } catch (e) {
    Logger.log('  ❌ ' + e.message);
    return;
  }

  Logger.log('\n[2] Campos del formulario:');
  var items    = form.getItems();
  var fotoItem = null;

  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    Logger.log('  ' + (i + 1) + '. [' + it.getType() + '] ' + it.getTitle());
    if (it.getTitle() === 'Fotos') fotoItem = it;
  }

  Logger.log('\n[3] Verificación campo Fotos:');
  if (!fotoItem) {
    Logger.log('  ❌ Campo "Fotos" no encontrado.');
    return;
  }

  var tipo = fotoItem.getType();
  var esUpload = (tipo === FormApp.ItemType.FILE_UPLOAD);
  Logger.log('  Tipo:           ' + tipo + (esUpload ? ' ✓' : ' ❌ (esperado FILE_UPLOAD)'));

  if (esUpload) {
    var fu = fotoItem.asFileUploadItem();
    Logger.log('  Tipos archivo:  ' + fu.getAllowedFileTypes().join(', '));
    Logger.log('  Máx. archivos:  ' + fu.getMaxFileCount());
    Logger.log('  Obligatorio:    ' + fu.isRequired());
  }

  Logger.log('\n================================================');
  Logger.log('RESULTADO:');
  Logger.log('  FILE_UPLOAD: ' + (esUpload ? '✓' : '❌'));
  Logger.log('  Login:       verificar manualmente en Form > Configuración.');
  Logger.log(esUpload ? 'CAMPO LISTO. Activa login en Form > Configuración y prueba en móvil.' : 'Revisa los puntos con ❌.');
  Logger.log('================================================');
}
