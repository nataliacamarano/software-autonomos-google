// ============================================================
// Triggers.gs — Instalación de triggers y creación de formularios
// ============================================================

/**
 * Instala el trigger onFormSubmit del spreadsheet.
 * Si ya existe, no crea un duplicado.
 * Ejecutar UNA VEZ desde el editor tras setup().
 */
function setupTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const yaExiste = ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === 'onFormSubmitHandler' &&
           t.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT;
  });

  if (!yaExiste) {
    ScriptApp.newTrigger('onFormSubmitHandler')
      .forSpreadsheet(ss)
      .onFormSubmit()
      .create();
    Logger.log('✓ Trigger onFormSubmit instalado');
  } else {
    Logger.log('ℹ️ Trigger onFormSubmit ya estaba instalado');
  }
}

/**
 * Renombra la hoja de respuestas del formulario de trabajos a RESP_TRABAJOS.
 * Necesario para que onFormSubmitHandler enrute las respuestas correctamente.
 * Busca automáticamente cualquier "Form Responses N" con cabeceras de trabajos.
 * NO borra respuestas existentes.
 */
function fixTrabajosResponseSheet() {
  Logger.log('================================================');
  Logger.log('fixTrabajosResponseSheet()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var ss = getTargetSS_();

  if (ss.getSheetByName('RESP_TRABAJOS')) {
    Logger.log('✓ La hoja RESP_TRABAJOS ya existe. Sin cambios.');
    return;
  }

  var sheets = ss.getSheets();
  Logger.log('Hojas en el spreadsheet:');
  sheets.forEach(function(s) { Logger.log('  - "' + s.getName() + '"'); });

  var found = null;
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (/form.responses|respuestas.de.formulario/i.test(name) &&
        _sheetHasTrabajosHeaders_(sheets[i])) {
      found = sheets[i];
      break;
    }
  }

  if (!found) {
    Logger.log('❌ No se encontró ninguna hoja de respuestas con cabeceras de trabajos.');
    Logger.log('   Comprueba que el formulario está vinculado al spreadsheet y tiene al menos');
    Logger.log('   los campos: Cliente, Operario, Hora inicio, Trabajo realizado.');
    return;
  }

  var oldName = found.getName();
  found.setName('RESP_TRABAJOS');
  Logger.log('✓ "' + oldName + '" → "RESP_TRABAJOS"');
  Logger.log('  El trigger onFormSubmitHandler procesará las nuevas respuestas automáticamente.');
  Logger.log('  Si el trigger no está instalado, ejecuta setupTriggers().');
}

/**
 * Elimina todos los triggers del proyecto y los reinstala desde cero.
 */
function resetTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  Logger.log('✓ Todos los triggers eliminados');
  setupTriggers();
}

/**
 * Router principal de respuestas de formularios.
 * Detecta el formulario origen por el nombre de la hoja de respuestas.
 * RESP_CLIENTES  → onClienteFormSubmit  + refreshTrabajosDropdowns
 * RESP_OPERARIOS → onOperarioFormSubmit + refreshTrabajosDropdowns
 * RESP_TRABAJOS  → onTrabajosFormSubmit
 */
function onFormSubmitHandler(e) {
  try {
    const sheetName = e.range.getSheet().getName();
    const nv = e.namedValues;
    logEvent('INFO', 'onFormSubmitHandler', 'Respuesta recibida — hoja: "' + sheetName + '"', '');

    if (sheetName === 'RESP_CLIENTES') {
      const res = onClienteFormSubmit(nv);
      if (res && res.success) {
        logEvent('INFO', 'onFormSubmitHandler', 'Alta cliente OK: ' + res.id, '');
      } else {
        logEvent('ERROR', 'onFormSubmitHandler', 'Alta cliente FALLIDA: ' + (res && res.error), '');
      }
      refreshTrabajosDropdowns();
      logEvent('INFO', 'onFormSubmitHandler', 'refreshTrabajosDropdowns completado tras alta de cliente', '');

    } else if (sheetName === 'RESP_OPERARIOS') {
      const res = onOperarioFormSubmit(nv);
      if (res && res.success) {
        logEvent('INFO', 'onFormSubmitHandler', 'Alta operario OK: ' + res.id, '');
      } else {
        logEvent('ERROR', 'onFormSubmitHandler', 'Alta operario FALLIDA: ' + (res && res.error), '');
      }
      refreshTrabajosDropdowns();
      logEvent('INFO', 'onFormSubmitHandler', 'refreshTrabajosDropdowns completado tras alta de operario', '');

    } else if (sheetName === 'RESP_TRABAJOS') {
      onTrabajosFormSubmitSafe(nv);  // versión idempotente — evita duplicados por doble trigger

    } else if (sheetName === 'RESP_GASTOS') {
      const res = onGastoFormSubmit(nv);
      if (res && res.success) {
        logEvent('INFO', 'onFormSubmitHandler', 'Gasto registrado OK: ' + res.id, '');
      } else {
        logEvent('ERROR', 'onFormSubmitHandler', 'Gasto FALLIDO: ' + (res && res.error), '');
      }

    } else {
      logEvent('WARN', 'onFormSubmitHandler',
               'Hoja de respuestas no reconocida: "' + sheetName + '"', '');
    }
  } catch (err) {
    logEvent('ERROR', 'onFormSubmitHandler', err.message, err.stack);
  }
}

// ── Creación de formularios ───────────────────────────────────

/**
 * Crea el formulario de Clientes y el de Operarios, los vincula al spreadsheet
 * y garantiza que cada uno escribe en su hoja correcta.
 * Idempotente: salta lo que ya esté correctamente configurado.
 * Ejecutar tras setupTriggers().
 */
function setupForms() {
  Logger.log('=== CONFIGURANDO FORMULARIOS ===');
  createClienteForm_();
  createOperarioForm_();
  Logger.log('=== FORMULARIOS LISTOS ===');
  Logger.log('URLs en CONFIG → form_clientes_url | form_operarios_url');
}

/**
 * Crea o repara el formulario de Clientes.
 *
 * Lógica:
 *  - Si form_clientes_id existe Y RESP_CLIENTES tiene cabeceras correctas → OK.
 *  - Si form_clientes_id existe Y RESP_CLIENTES es incorrecta o no existe → rebind.
 *  - Si no hay form_clientes_id → crea formulario nuevo y vincula.
 */
function createClienteForm_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formId = getConfig('form_clientes_id');
  const respSheet = ss.getSheetByName('RESP_CLIENTES');

  // Todo correcto: formulario existe y hoja tiene cabeceras de cliente
  if (formId && respSheet && esHojaRespuestasClientes_(respSheet)) {
    Logger.log('ℹ️ Formulario Clientes ya está correctamente configurado. Sin cambios.');
    return;
  }

  // Formulario existe pero la hoja es incorrecta o falta: revincular
  if (formId) {
    Logger.log('⚠️ Formulario Clientes existe pero la hoja de respuestas es incorrecta. Revinculando...');
    rebindFormSafely_(formId, 'RESP_CLIENTES');
    return;
  }

  // Crear formulario nuevo
  Logger.log('Creando formulario Clientes nuevo...');
  const form = FormApp.create('Alta de Cliente — Álvaro Gestión');
  form.setDescription('Registra un nuevo cliente en el sistema.');
  form.setCollectEmail(false);
  form.setAllowResponseEdits(false);
  form.addTextItem().setTitle('Nombre').setRequired(true);
  form.addTextItem().setTitle('Apellidos').setRequired(true);
  form.addTextItem().setTitle('Empresa').setRequired(false);
  form.addTextItem().setTitle('Email').setRequired(true);
  form.addTextItem().setTitle('Teléfono').setRequired(true);
  form.addParagraphTextItem().setTitle('Dirección').setRequired(true);
  form.addTextItem().setTitle('CIF / NIF').setRequired(false);

  setConfig('form_clientes_id',  form.getId());
  setConfig('form_clientes_url', form.getPublishedUrl());

  const snapshot = ss.getSheets().map(function(s) { return s.getName(); });
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  renameNewSheet_(ss, snapshot, 'RESP_CLIENTES');

  Logger.log('✅ Formulario Clientes creado → RESP_CLIENTES');
  Logger.log('URL: ' + getConfig('form_clientes_url'));
}

/**
 * Crea o repara el formulario de Operarios.
 * Misma lógica que createClienteForm_.
 */
function createOperarioForm_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formId = getConfig('form_operarios_id');
  const respSheet = ss.getSheetByName('RESP_OPERARIOS');

  // Todo correcto: formulario existe y hoja tiene cabeceras de operario
  if (formId && respSheet && esHojaRespuestasOperarios_(respSheet)) {
    Logger.log('ℹ️ Formulario Operarios ya está correctamente configurado. Sin cambios.');
    return;
  }

  // Formulario existe pero la hoja es incorrecta o falta: revincular
  if (formId) {
    Logger.log('⚠️ Formulario Operarios existe pero la hoja de respuestas es incorrecta. Revinculando...');
    rebindFormSafely_(formId, 'RESP_OPERARIOS');
    return;
  }

  // Crear formulario nuevo
  Logger.log('Creando formulario Operarios nuevo...');
  const form = FormApp.create('Alta de Operario — Álvaro Gestión');
  form.setDescription('Registra un nuevo operario en el sistema.');
  form.setCollectEmail(false);
  form.setAllowResponseEdits(false);
  form.addTextItem().setTitle('Nombre').setRequired(true);
  form.addTextItem().setTitle('Apellidos').setRequired(true);
  form.addTextItem().setTitle('Email').setRequired(true);
  form.addTextItem().setTitle('Teléfono').setRequired(true);
  setConfig('form_operarios_id',  form.getId());
  setConfig('form_operarios_url', form.getPublishedUrl());

  const snapshot = ss.getSheets().map(function(s) { return s.getName(); });
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  renameNewSheet_(ss, snapshot, 'RESP_OPERARIOS');

  Logger.log('✅ Formulario Operarios creado → RESP_OPERARIOS');
  Logger.log('URL: ' + getConfig('form_operarios_url'));
}

// ── Validadores de cabecera ───────────────────────────────────

/**
 * True si la hoja tiene al menos 3 de los 6 campos del formulario de Clientes.
 */
function esHojaRespuestasOperarios_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (sheet.getLastRow() === 0 || lastCol === 0) return false;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h).trim().toLowerCase(); });

  const camposClave = ['nombre', 'apellidos', 'email', 'teléfono'];
  const coincidencias = camposClave.filter(function(c) {
    return headers.some(function(h) { return h.includes(c); });
  });

  Logger.log('  "' + sheet.getName() + '": ' + coincidencias.length + '/4 campos operario — [' + headers.join(', ') + ']');
  return coincidencias.length >= 3;
}

// ── Utilidades privadas ───────────────────────────────────────

/**
 * Espera hasta 12 s a que aparezca una hoja nueva (no en el snapshot)
 * y la renombra. Detección por diferencia de conjuntos, sin depender
 * del nombre que Google asigne.
 */
function renameNewSheet_(ss, snapshot, newName) {
  for (let i = 0; i < 6; i++) {
    Utilities.sleep(2000);
    const current = ss.getSheets();
    const nueva = current.find(function(s) {
      return !snapshot.includes(s.getName()) && s.getName() !== newName;
    });
    if (nueva) {
      nueva.setName(newName);
      Logger.log('✓ Nueva hoja de respuestas renombrada → "' + newName + '"');
      return true;
    }
    if (ss.getSheetByName(newName)) return true;
  }
  logEvent('WARN', 'renameNewSheet_', 'No apareció hoja nueva para "' + newName + '"', '');
  return false;
}

/**
 * Busca cualquier hoja con nombre tipo "Respuestas de formulario N"
 * o "Form Responses N" y la renombra. Fallback para correcciones manuales.
 */
function renameAnyResponseSheet_(ss, newName) {
  const sheets = ss.getSheets();
  for (let i = sheets.length - 1; i >= 0; i--) {
    const name = sheets[i].getName();
    if (/respuestas.de.formulario|form.responses/i.test(name)) {
      sheets[i].setName(newName);
      Logger.log('✓ "' + name + '" → "' + newName + '"');
      return true;
    }
  }
  return false;
}
