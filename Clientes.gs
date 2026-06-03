// ============================================================
// Clientes.gs — CRUD de clientes
// ============================================================

// Columnas de la hoja CLIENTES (base 1)
const COL_CLI = {
  ID:        1,
  TIMESTAMP: 2,
  NOMBRE:    3,
  APELLIDOS: 4,
  EMPRESA:   5,
  EMAIL:     6,
  TELEFONO:  7,
  DIRECCION: 8,
  CIF_NIF:   9,
  ACTIVO:    10
};

// ---- Lectura ----

function getClientes(soloActivos) {
  const sheet = getSheet(SHEET_NAMES.CLIENTES);
  const rows = sheet.getDataRange().getValues();
  const todos = rowsToObjects(rows);
  return soloActivos
    ? todos.filter(c => String(c.Activo).toLowerCase() === 'true')
    : todos;
}

function getClienteById(id) {
  return getClientes(false).find(c => String(c.ID_Cliente) === String(id)) || null;
}

// ---- Escritura ----

/**
 * Crea un cliente nuevo. Usado por el panel web y el trigger de formulario.
 * Devuelve { success, id } o { success: false, error }.
 */
function createCliente(data) {
  if (!data.nombre || !data.apellidos || !data.email || !data.telefono || !data.direccion) {
    return { success: false, error: 'Faltan campos obligatorios: nombre, apellidos, email, teléfono, dirección.' };
  }
  try {
    const sheet = getSheet(SHEET_NAMES.CLIENTES);
    const id = nextId(SHEET_NAMES.CLIENTES, 'CLI');
    sheet.appendRow([
      id,
      new Date(),
      String(data.nombre).trim(),
      String(data.apellidos).trim(),
      String(data.empresa || '').trim(),
      String(data.email).trim().toLowerCase(),
      String(data.telefono).trim(),
      String(data.direccion).trim(),
      String(data.cifNif || '').trim().toUpperCase(),
      true
    ]);
    logEvent('INFO', 'createCliente', `Cliente creado: ${id}`, String(data.nombre));
    refreshTrabajosDropdowns();
    return { success: true, id };
  } catch (e) {
    logEvent('ERROR', 'createCliente', e.message, JSON.stringify(data));
    return { success: false, error: e.message };
  }
}

/**
 * Actualiza los campos presentes en data para el cliente con ese ID.
 */
function updateCliente(id, data) {
  try {
    const sheet = getSheet(SHEET_NAMES.CLIENTES);
    const row = findRowById_(sheet, id);
    if (!row) return { success: false, error: `Cliente "${id}" no encontrado.` };

    const colMap = {
      nombre:    COL_CLI.NOMBRE,
      apellidos: COL_CLI.APELLIDOS,
      empresa:   COL_CLI.EMPRESA,
      email:     COL_CLI.EMAIL,
      telefono:  COL_CLI.TELEFONO,
      direccion: COL_CLI.DIRECCION,
      cifNif:    COL_CLI.CIF_NIF
    };

    Object.entries(colMap).forEach(([key, col]) => {
      if (data[key] !== undefined) {
        sheet.getRange(row, col).setValue(String(data[key]).trim());
      }
    });

    logEvent('INFO', 'updateCliente', `Cliente actualizado: ${id}`, '');
    refreshTrabajosDropdowns();
    return { success: true };
  } catch (e) {
    logEvent('ERROR', 'updateCliente', e.message, id);
    return { success: false, error: e.message };
  }
}

/**
 * Alterna el estado Activo/Inactivo de un cliente.
 */
function toggleClienteActivo(id) {
  try {
    const sheet = getSheet(SHEET_NAMES.CLIENTES);
    const row = findRowById_(sheet, id);
    if (!row) return { success: false, error: `Cliente "${id}" no encontrado.` };
    const nuevoValor = !sheet.getRange(row, COL_CLI.ACTIVO).getValue();
    sheet.getRange(row, COL_CLI.ACTIVO).setValue(nuevoValor);
    return { success: true, activo: nuevoValor };
  } catch (e) {
    logEvent('ERROR', 'toggleClienteActivo', e.message, id);
    return { success: false, error: e.message };
  }
}

// ---- Handler de formulario ----

/**
 * Procesa namedValues de una respuesta del formulario de clientes.
 * Llamado desde onFormSubmitHandler() en Triggers.gs.
 */
function onClienteFormSubmit(namedValues) {
  return createCliente({
    nombre:    (namedValues['Nombre']     || [''])[0],
    apellidos: (namedValues['Apellidos']  || [''])[0],
    empresa:   (namedValues['Empresa']    || [''])[0],
    email:     (namedValues['Email']      || [''])[0],
    telefono:  (namedValues['Teléfono']   || [''])[0],
    direccion: (namedValues['Dirección']  || [''])[0],
    cifNif:    (namedValues['CIF / NIF']  || [''])[0]
  });
}
