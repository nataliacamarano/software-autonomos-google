// ============================================================
// Operarios.gs — CRUD de operarios
// ============================================================

// Columnas de la hoja OPERARIOS (base 1)
const COL_OPE = {
  ID:           1,
  TIMESTAMP:    2,
  NOMBRE:       3,
  APELLIDOS:    4,
  EMAIL:        5,
  TELEFONO:     6,
  ESPECIALIDAD: 7,
  ACTIVO:       8
};


// ---- Lectura ----

function getOperarios(soloActivos) {
  const sheet = getSheet(SHEET_NAMES.OPERARIOS);
  const rows = sheet.getDataRange().getValues();
  const todos = rowsToObjects(rows);
  return soloActivos
    ? todos.filter(o => String(o.Activo).toLowerCase() === 'true')
    : todos;
}

function getOperarioById(id) {
  return getOperarios(false).find(o => String(o.ID_Operario) === String(id)) || null;
}

// ---- Escritura ----

/**
 * Crea un operario nuevo. Usado por el panel web y el trigger de formulario.
 * Devuelve { success, id } o { success: false, error }.
 */
function createOperario(data) {
  if (!data.nombre || !data.apellidos || !data.email || !data.telefono) {
    return { success: false, error: 'Faltan campos obligatorios: nombre, apellidos, email, teléfono.' };
  }
  try {
    const sheet = getSheet(SHEET_NAMES.OPERARIOS);
    const id = nextId(SHEET_NAMES.OPERARIOS, 'OPE');
    sheet.appendRow([
      id,
      new Date(),
      String(data.nombre).trim(),
      String(data.apellidos).trim(),
      String(data.email).trim().toLowerCase(),
      String(data.telefono).trim(),
      '',   // col 7 Especialidad — retirada, se deja vacía para preservar estructura
      true
    ]);
    logEvent('INFO', 'createOperario', `Operario creado: ${id}`, String(data.nombre));
    refreshTrabajosDropdowns();
    return { success: true, id };
  } catch (e) {
    logEvent('ERROR', 'createOperario', e.message, JSON.stringify(data));
    return { success: false, error: e.message };
  }
}

/**
 * Actualiza los campos presentes en data para el operario con ese ID.
 */
function updateOperario(id, data) {
  try {
    const sheet = getSheet(SHEET_NAMES.OPERARIOS);
    const row = findRowById_(sheet, id);
    if (!row) return { success: false, error: `Operario "${id}" no encontrado.` };

    const colMap = {
      nombre:    COL_OPE.NOMBRE,
      apellidos: COL_OPE.APELLIDOS,
      email:     COL_OPE.EMAIL,
      telefono:  COL_OPE.TELEFONO
    };

    Object.entries(colMap).forEach(([key, col]) => {
      if (data[key] !== undefined) {
        sheet.getRange(row, col).setValue(String(data[key]).trim());
      }
    });

    logEvent('INFO', 'updateOperario', `Operario actualizado: ${id}`, '');
    refreshTrabajosDropdowns();
    return { success: true };
  } catch (e) {
    logEvent('ERROR', 'updateOperario', e.message, id);
    return { success: false, error: e.message };
  }
}

/**
 * Alterna el estado Activo/Inactivo de un operario.
 */
function toggleOperarioActivo(id) {
  try {
    const sheet = getSheet(SHEET_NAMES.OPERARIOS);
    const row = findRowById_(sheet, id);
    if (!row) return { success: false, error: `Operario "${id}" no encontrado.` };
    const nuevoValor = !sheet.getRange(row, COL_OPE.ACTIVO).getValue();
    sheet.getRange(row, COL_OPE.ACTIVO).setValue(nuevoValor);
    return { success: true, activo: nuevoValor };
  } catch (e) {
    logEvent('ERROR', 'toggleOperarioActivo', e.message, id);
    return { success: false, error: e.message };
  }
}

// ---- Handler de formulario ----

/**
 * Procesa namedValues de una respuesta del formulario de operarios.
 * Llamado desde onFormSubmitHandler() en Triggers.gs.
 */
function onOperarioFormSubmit(namedValues) {
  return createOperario({
    nombre:    (namedValues['Nombre']    || [''])[0],
    apellidos: (namedValues['Apellidos'] || [''])[0],
    email:     (namedValues['Email']     || [''])[0],
    telefono:  (namedValues['Teléfono'] || [''])[0]
  });
}

// ---- Mantenimiento del formulario ----

/**
 * Elimina el campo Especialidad del formulario de operarios existente.
 * NO recrea el formulario. NO borra respuestas ni filas de la hoja.
 */
function removeEspecialidadField() {
  Logger.log('================================================');
  Logger.log('removeEspecialidadField()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var formId = getConfig('form_operarios_id');
  Logger.log('\n[1] form_operarios_id: ' + (formId || '❌ NO CONFIGURADO'));
  if (!formId) {
    Logger.log('ABORTANDO. form_operarios_id no encontrado en CONFIG.');
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

  Logger.log('\n[2] Buscando campo "Especialidad"...');
  var items    = form.getItems();
  var deleted  = false;

  for (var i = 0; i < items.length; i++) {
    if (items[i].getTitle() === 'Especialidad') {
      form.deleteItem(items[i]);
      deleted = true;
      Logger.log('  ✓ Campo "Especialidad" eliminado.');
      break;
    }
  }

  if (!deleted) {
    Logger.log('  ℹ️ Campo "Especialidad" no encontrado — puede que ya se hubiera eliminado.');
  }

  Logger.log('\n[3] Campos restantes en el formulario:');
  var remaining = form.getItems();
  for (var j = 0; j < remaining.length; j++) {
    Logger.log('  ' + (j + 1) + '. [' + remaining[j].getType() + '] ' + remaining[j].getTitle());
  }

  Logger.log('\n================================================');
  Logger.log(deleted ? 'CAMPO ELIMINADO.' : 'Sin cambios (campo no existía).');
  Logger.log('Ejecuta validateOperariosForm() para confirmar el estado final.');
  Logger.log('================================================');
}

/**
 * Valida que el formulario de operarios tiene exactamente los 4 campos
 * correctos y que Especialidad no está presente.
 */
function validateOperariosForm() {
  Logger.log('================================================');
  Logger.log('validateOperariosForm()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var formId = getConfig('form_operarios_id');
  Logger.log('\n[1] form_operarios_id: ' + (formId || '❌ NO CONFIGURADO'));
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

  // Campos esperados (en cualquier orden)
  var esperados     = ['Nombre', 'Apellidos', 'Email', 'Teléfono'];
  var items         = form.getItems();
  var titulos       = items.map(function(it) { return it.getTitle(); });
  var tieneEspecialidad = titulos.indexOf('Especialidad') !== -1;

  Logger.log('\n[2] Campos en el formulario (' + items.length + '):');
  for (var i = 0; i < items.length; i++) {
    Logger.log('  ' + (i + 1) + '. [' + items[i].getType() + '] ' + items[i].getTitle());
  }

  Logger.log('\n[3] Comprobación de campos requeridos:');
  var camposOk = true;
  for (var e = 0; e < esperados.length; e++) {
    var presente = titulos.indexOf(esperados[e]) !== -1;
    Logger.log('  ' + esperados[e] + ': ' + (presente ? '✓' : '❌ NO ENCONTRADO'));
    if (!presente) camposOk = false;
  }

  Logger.log('\n[4] Especialidad ausente: ' + (tieneEspecialidad ? '❌ TODAVÍA PRESENTE — ejecuta removeEspecialidadField()' : '✓'));

  // Hoja OPERARIOS
  Logger.log('\n[5] Hoja OPERARIOS:');
  try {
    var sheet = getSheet('OPERARIOS');
    var filas = Math.max(0, sheet.getLastRow() - 1);
    Logger.log('  ✓ GID=' + sheet.getSheetId() + '  registros=' + filas);
    Logger.log('  Columna 7 (Especialidad): conservada con datos históricos.');
  } catch (err) {
    Logger.log('  ❌ ' + err.message);
  }

  var ok = camposOk && !tieneEspecialidad;
  Logger.log('\n================================================');
  Logger.log('RESULTADO:');
  Logger.log('  Campos requeridos: ' + (camposOk          ? '✓' : '❌'));
  Logger.log('  Especialidad fuera: ' + (!tieneEspecialidad ? '✓' : '❌'));
  Logger.log(ok ? 'FORMULARIO CORRECTO.' : 'Revisa los puntos con ❌.');
  Logger.log('================================================');
}
