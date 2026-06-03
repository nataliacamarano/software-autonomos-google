// ============================================================
// Clientes.gs — CRUD de clientes
// ============================================================

// Columnas de la hoja CLIENTES (base 1)
const COL_CLI = {
  ID:                 1,
  TIMESTAMP:          2,
  NOMBRE:             3,
  APELLIDOS:          4,
  EMPRESA:            5,
  EMAIL:              6,
  TELEFONO:           7,
  DIRECCION:          8,
  CIF_NIF:            9,
  ACTIVO:             10,
  METODO_PAGO:        11,
  TITULAR_CUENTA:     12,
  BANCO:              13,
  IBAN:               14,
  SWIFT_BIC:          15,
  PAYPAL_EMAIL:       16,
  OBSERVACIONES_PAGO: 17
};

const METODOS_PAGO = [
  'Transferencia bancaria', 'Bizum', 'Efectivo',
  'Tarjeta', 'PayPal', 'Domiciliación bancaria', 'Otro'
];

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
      String(data.metodoPago        || '').trim(),
      String(data.titularCuenta     || '').trim(),
      String(data.banco             || '').trim(),
      String(data.iban              || '').trim().toUpperCase(),
      String(data.swiftBic          || '').trim().toUpperCase(),
      String(data.paypalEmail       || '').trim().toLowerCase(),
      String(data.observacionesPago || '').trim()
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
      nombre:             COL_CLI.NOMBRE,
      apellidos:          COL_CLI.APELLIDOS,
      empresa:            COL_CLI.EMPRESA,
      email:              COL_CLI.EMAIL,
      telefono:           COL_CLI.TELEFONO,
      direccion:          COL_CLI.DIRECCION,
      cifNif:             COL_CLI.CIF_NIF,
      metodoPago:         COL_CLI.METODO_PAGO,
      titularCuenta:      COL_CLI.TITULAR_CUENTA,
      banco:              COL_CLI.BANCO,
      iban:               COL_CLI.IBAN,
      swiftBic:           COL_CLI.SWIFT_BIC,
      paypalEmail:        COL_CLI.PAYPAL_EMAIL,
      observacionesPago:  COL_CLI.OBSERVACIONES_PAGO
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

/**
 * Diálogo de edición de datos de pago lanzado desde PANEL_CLIENTE.
 * Usa prompts secuenciales según el método elegido.
 */
function editarDatosPagoDesdePanel() {
  var ui  = SpreadsheetApp.getUi();
  var sh  = SpreadsheetApp.getActive().getSheetByName('PANEL_CLIENTE');
  if (!sh) { ui.alert('PANEL_CLIENTE no encontrado.'); return; }

  var cliente = String(sh.getRange(PC_.CLIENTE[0], PC_.CLIENTE[1]).getValue() || '').trim();
  if (!cliente) { ui.alert('Selecciona un cliente primero.'); return; }

  var cliData = _getClienteByLabel_(cliente);
  if (!cliData) { ui.alert('Cliente no encontrado en la hoja CLIENTES.'); return; }

  // ── Paso 1: método de pago ────────────────────────────────────
  var r1 = ui.prompt(
    'Datos de pago — ' + cliente,
    'Método de pago\n\nOpciones: ' + METODOS_PAGO.join(' · ') +
    '\n\nActual: ' + (cliData.metodoPago || '(sin definir)'),
    ui.ButtonSet.OK_CANCEL
  );
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var nuevoPago = { metodoPago: r1.getResponseText().trim() };
  var mp = nuevoPago.metodoPago.toLowerCase();

  // ── Paso 2: campos según método ──────────────────────────────
  if (mp === 'transferencia bancaria' || mp === 'domiciliación bancaria') {
    var prompts = [
      ['titularCuenta', 'Titular de la cuenta',   cliData.titularCuenta     || ''],
      ['banco',         'Banco',                   cliData.banco             || ''],
      ['iban',          'IBAN',                    cliData.iban              || ''],
      ['swiftBic',      'SWIFT / BIC (opcional)',  cliData.swiftBic          || '']
    ];
    for (var pi = 0; pi < prompts.length; pi++) {
      var p = prompts[pi];
      var resp = ui.prompt(p[1], 'Actual: ' + (p[2] || '(vacío)'), ui.ButtonSet.OK_CANCEL);
      if (resp.getSelectedButton() !== ui.Button.OK) return;
      nuevoPago[p[0]] = resp.getResponseText().trim();
    }
  } else if (mp === 'paypal') {
    var rPP = ui.prompt('Email de PayPal', 'Actual: ' + (cliData.paypalEmail || '(vacío)'), ui.ButtonSet.OK_CANCEL);
    if (rPP.getSelectedButton() !== ui.Button.OK) return;
    nuevoPago.paypalEmail = rPP.getResponseText().trim().toLowerCase();
  } else {
    var rObs = ui.prompt('Observaciones de pago', 'Actual: ' + (cliData.observacionesPago || '(vacío)'), ui.ButtonSet.OK_CANCEL);
    if (rObs.getSelectedButton() !== ui.Button.OK) return;
    nuevoPago.observacionesPago = rObs.getResponseText().trim();
  }

  var result = updateCliente(cliData.id, nuevoPago);
  if (result.success) {
    refreshPanelCliente();
    _pcEstado_(sh, '✅  Datos de pago guardados para ' + cliente, '#ffffff', '#34a853');
  } else {
    ui.alert('Error al guardar: ' + result.error);
  }
}
