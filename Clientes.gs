// ============================================================
// Clientes.gs — CRUD de clientes
// ============================================================

// Columnas de la hoja CLIENTES (base 1)
const COL_CLI = {
  ID:           1,
  TIMESTAMP:    2,
  NOMBRE:       3,
  APELLIDOS:    4,
  EMPRESA:      5,
  EMAIL:        6,
  TELEFONO:     7,
  DIRECCION:    8,
  CIF_NIF:      9,
  ACTIVO:       10,
  METODOS_PAGO: 11   // JSON array de métodos de pago
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
      true,
      JSON.stringify(Array.isArray(data.metodosPago) ? data.metodosPago : [])
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
      nombre:      COL_CLI.NOMBRE,
      apellidos:   COL_CLI.APELLIDOS,
      empresa:     COL_CLI.EMPRESA,
      email:       COL_CLI.EMAIL,
      telefono:    COL_CLI.TELEFONO,
      direccion:   COL_CLI.DIRECCION,
      cifNif:      COL_CLI.CIF_NIF
    };

    Object.entries(colMap).forEach(([key, col]) => {
      if (data[key] !== undefined) {
        sheet.getRange(row, col).setValue(String(data[key]).trim());
      }
    });

    // Métodos de pago (JSON array)
    if (data.metodosPago !== undefined) {
      var val = Array.isArray(data.metodosPago)
        ? JSON.stringify(data.metodosPago)
        : String(data.metodosPago);
      sheet.getRange(row, COL_CLI.METODOS_PAGO).setValue(val);
    }

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

/**
 * Abre el modal HTML de edición de datos de pago (desde PANEL_CLIENTE).
 */
function editarDatosPagoDesdePanel() {
  var ui = SpreadsheetApp.getUi();
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName('PANEL_CLIENTE');
    if (!sh) { ui.alert('PANEL_CLIENTE no encontrado.'); return; }

    var cliente = String(sh.getRange(PC_.CLIENTE[0], PC_.CLIENTE[1]).getValue() || '').trim();
    if (!cliente) { ui.alert('Selecciona un cliente primero en el desplegable.'); return; }

    var cliData = _getClienteByLabel_(cliente);
    if (!cliData) { ui.alert('Cliente no encontrado en la hoja CLIENTES:\n"' + cliente + '"'); return; }

    var template = HtmlService.createTemplateFromFile('EditarPago');
    template.clienteLabel  = cliente;
    template.clienteNombre = (cliData.nombre || cliente).substring(0, 35);
    template.metodosJson   = JSON.stringify(cliData.metodosPago || []);

    var titulo = 'Datos de pago' + (cliente.length > 28 ? '' : ' — ' + cliente);
    ui.showModalDialog(
      template.evaluate().setWidth(440).setHeight(560),
      titulo
    );
  } catch (err) {
    ui.alert('Error al abrir datos de pago:\n' + err.message);
    logEvent('ERROR', 'editarDatosPagoDesdePanel', err.message, err.stack || '');
  }
}

// ================================================================
// NORMALIZACIÓN Y DEDUPLICACIÓN DE CLIENTES
// ================================================================

/**
 * Normaliza una etiqueta de cliente para comparación:
 * trim → lowercase → quitar tildes → colapsar espacios.
 */
function _normalizeClientLabel_(s) {
  return String(s || '').trim().toLowerCase()
    .replace(/[áàäâã]/g, 'a').replace(/[éèëê]/g, 'e')
    .replace(/[íìïî]/g, 'i').replace(/[óòöôõ]/g, 'o')
    .replace(/[úùüû]/g, 'u').replace(/ñ/g, 'n').replace(/ç/g, 'c')
    .replace(/\s+/g, ' ').trim();
}

/**
 * Revisa la hoja CLIENTES, elimina filas duplicadas (conserva la primera aparición
 * de cada nombre normalizado) y refresca todos los dropdowns.
 * Llamado desde: Menú ⚙️ Sistema → 🧹 Limpiar clientes duplicados.
 */
function limpiarClientesDuplicados() {
  var ui = SpreadsheetApp.getUi();
  try {
    var sheet   = getSheet(SHEET_NAMES.CLIENTES);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) { ui.alert('No hay clientes registrados.'); return; }

    var data     = sheet.getRange(2, 1, lastRow - 1, COL_CLI.METODOS_PAGO).getValues();
    var seen     = {};
    var toDelete = [];

    for (var i = 0; i < data.length; i++) {
      var nombre    = String(data[i][COL_CLI.NOMBRE    - 1] || '').trim();
      var apellidos = String(data[i][COL_CLI.APELLIDOS - 1] || '').trim();
      var empresa   = String(data[i][COL_CLI.EMPRESA   - 1] || '').trim();
      var label     = (nombre + ' ' + apellidos).trim();
      if (empresa) label += ' (' + empresa + ')';
      if (!label.replace(/[\s()]/g, '')) continue;  // fila completamente vacía

      var key = _normalizeClientLabel_(label);
      if (!seen[key]) {
        seen[key] = true;
      } else {
        toDelete.push(i + 2);  // índice de fila en la hoja (fila 1 = cabecera)
      }
    }

    if (toDelete.length === 0) {
      ui.alert('✅ No se encontraron clientes duplicados.');
      return;
    }

    // Eliminar de abajo a arriba para que los índices no se desplacen
    toDelete.sort(function(a, b) { return b - a; });
    for (var d = 0; d < toDelete.length; d++) {
      sheet.deleteRow(toDelete[d]);
    }

    _refreshAllClienteDropdowns_();

    ui.alert(
      '✅ Clientes duplicados eliminados y dropdowns actualizados.\n\n' +
      'Filas eliminadas: ' + toDelete.length
    );
  } catch (e) {
    ui.alert('Error en limpiarClientesDuplicados:\n' + e.message);
    logEvent('ERROR', 'limpiarClientesDuplicados', e.message, '');
  }
}

/**
 * Refresca los dropdowns de cliente en PANEL_FACTURAS, PANEL_CLIENTE y formulario de trabajos.
 */
function _refreshAllClienteDropdowns_() {
  var ss    = SpreadsheetApp.getActive();
  var names = _panelGetClienteNames_(ss);

  // PANEL_FACTURAS — dropdowns de cliente (izquierdo y presupuestos)
  var panelSh = ss.getSheetByName('PANEL_FACTURAS');
  if (panelSh && names.length > 0) {
    var val = SpreadsheetApp.newDataValidation()
      .requireValueInList(names, true).setAllowInvalid(false).build();
    panelSh.getRange(P_.CLIENTE[0],     P_.CLIENTE[1]).setDataValidation(val);
    panelSh.getRange(P_.PRE_CLIENTE[0], P_.PRE_CLIENTE[1]).setDataValidation(val);
  }

  // PANEL_CLIENTE — dropdown de cliente
  var pcSh = ss.getSheetByName('PANEL_CLIENTE');
  if (pcSh && names.length > 0) {
    pcSh.getRange(PC_.CLIENTE[0], PC_.CLIENTE[1]).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(names, true).setAllowInvalid(false).build()
    );
  }

  // Formulario de trabajos
  try { refreshTrabajosDropdowns(); } catch (_) {}
}

/**
 * Guarda los métodos de pago desde el modal HTML.
 * Llamado via google.script.run desde EditarPago.html.
 */
function guardarMetodosPago(clienteLabel, metodosJson) {
  var cliData = _getClienteByLabel_(clienteLabel);
  if (!cliData) throw new Error('Cliente no encontrado: ' + clienteLabel);

  var metodos;
  try { metodos = JSON.parse(metodosJson); } catch (_) { metodos = []; }
  if (!Array.isArray(metodos)) metodos = [];

  var result = updateCliente(cliData.id, { metodosPago: metodos });
  if (!result.success) throw new Error(result.error);

  refreshPanelCliente();
  var pcSh = SpreadsheetApp.getActive().getSheetByName('PANEL_CLIENTE');
  if (pcSh) _pcEstado_(pcSh, '✅  Datos de pago guardados para ' + clienteLabel, '#ffffff', '#34a853');
}
