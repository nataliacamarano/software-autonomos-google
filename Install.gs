// ============================================================
// LEGACY — NO USAR
// Install.gs contenía las rutinas de instalación inicial del sistema.
// Ya no está conectado al menú ni es llamado por ningún módulo activo.
// Conservado como referencia histórica. Desactivado en FASE 18.4 (2026-06-04).
// ============================================================

// ============================================================
// Install.gs — Backup del sistema antiguo e instalación limpia
// ============================================================

/**
 * Repara la instalación V2 sin crear nada nuevo.
 *
 * Situación: freshInstall creó el spreadsheet y los formularios pero
 * _installWaitAndRename_ no tuvo tiempo de detectar las hojas en el
 * nuevo spreadsheet. Google creó "Form Responses 1" y "Form Responses 2"
 * en lugar de RESP_CLIENTES / RESP_OPERARIOS.
 *
 * Esta función:
 *  1. Fija TARGET_SS_ID al spreadsheet V2 conocido.
 *  2. Abre los formularios ya creados.
 *  3. Identifica cada hoja de respuestas por cabeceras.
 *  4. Renombra: hoja Clientes → RESP_CLIENTES / hoja Operarios → RESP_OPERARIOS.
 *  5. Guarda los 4 valores en CONFIG.
 *
 * NO crea spreadsheet. NO crea formularios. NO crea hojas. NO toca el backup.
 */
function repairFreshInstallV2() {
  // ── IDs conocidos — solo para esta reparación ─────────────────
  var TARGET_SS_ID = 'SPREADSHEET_ID_HERE';
  var FORM_CLI_ID  = 'FORM_CLIENTES_ID_HERE';
  var FORM_OPE_ID  = 'FORM_OPERARIOS_ID_HERE';

  var SYSTEM = [
    'CONFIG','CLIENTES','OPERARIOS','TRABAJOS','FACTURAS',
    'PENDIENTES','COMPLETADOS','LOGS','RESP_CLIENTES','RESP_OPERARIOS'
  ];

  Logger.log('================================================');
  Logger.log('repairFreshInstallV2()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('Sin creacion de nada nuevo. Solo reparacion.');
  Logger.log('================================================');

  // ── 1. Fijar TARGET_SS_ID ─────────────────────────────────────
  PropertiesService.getScriptProperties().setProperty('TARGET_SS_ID', TARGET_SS_ID);
  Logger.log('\n[1] TARGET_SS_ID fijado: ' + TARGET_SS_ID);

  // ── 2. Abrir spreadsheet V2 ───────────────────────────────────
  Logger.log('\n[2] Abriendo spreadsheet V2...');
  var ss;
  try {
    ss = SpreadsheetApp.openById(TARGET_SS_ID);
    Logger.log('  OK: "' + ss.getName() + '"');
  } catch (e) {
    Logger.log('  ERROR: ' + e.message);
    return;
  }

  // ── 3. Inventario de hojas ────────────────────────────────────
  Logger.log('\n[3] Hojas encontradas:');
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    Logger.log('  "' + sheets[i].getName() + '"  GID=' + sheets[i].getSheetId());
  }

  // ── 4. Clasificar por cabeceras ───────────────────────────────
  Logger.log('\n[4] Clasificando hojas por cabeceras:');

  // Campos exclusivos de cada formulario
  var CLI_MARKERS = ['empresa', 'direcci', 'cif'];   // dirección y CIF / NIF
  var OPE_MARKERS = ['especialidad'];

  var sheetCli = null;
  var sheetOpe = null;

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var name  = sheet.getName();
    if (SYSTEM.indexOf(name) !== -1) continue;

    var lastCol = sheet.getLastColumn();
    if (lastCol === 0) {
      Logger.log('  "' + name + '": vacía — omitida');
      continue;
    }

    var raw     = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var headers = raw.map(function(h) { return String(h).trim().toLowerCase(); });
    Logger.log('  "' + name + '": [' + headers.join(' | ') + ']');

    var cliHits = CLI_MARKERS.filter(function(m) {
      return headers.some(function(h) { return h.indexOf(m) !== -1; });
    });
    var opeHits = OPE_MARKERS.filter(function(m) {
      return headers.some(function(h) { return h.indexOf(m) !== -1; });
    });

    Logger.log('    Marcadores Clientes (' + cliHits.length + '/3): ' + cliHits.join(', '));
    Logger.log('    Marcadores Operarios (' + opeHits.length + '/1): ' + opeHits.join(', '));

    if (opeHits.length >= 1 && !sheetOpe) {
      sheetOpe = sheet;
      Logger.log('    → RESP_OPERARIOS');
    } else if (cliHits.length >= 2 && !sheetCli) {
      sheetCli = sheet;
      Logger.log('    → RESP_CLIENTES');
    } else {
      Logger.log('    → No identificada');
    }
  }

  // ── 5. Renombrar ──────────────────────────────────────────────
  Logger.log('\n[5] Renombrando:');
  if (sheetCli) {
    var prevCli = sheetCli.getName();
    sheetCli.setName('RESP_CLIENTES');
    Logger.log('  ✓ "' + prevCli + '" → "RESP_CLIENTES"');
  } else {
    Logger.log('  ❌ Hoja Clientes no identificada');
  }

  if (sheetOpe) {
    var prevOpe = sheetOpe.getName();
    sheetOpe.setName('RESP_OPERARIOS');
    Logger.log('  ✓ "' + prevOpe + '" → "RESP_OPERARIOS"');
  } else {
    Logger.log('  ❌ Hoja Operarios no identificada');
  }

  // ── 6. Guardar IDs y URLs en CONFIG ──────────────────────────
  Logger.log('\n[6] Guardando en CONFIG:');
  try {
    var formCli = FormApp.openById(FORM_CLI_ID);
    setConfig('form_clientes_id',  FORM_CLI_ID);
    setConfig('form_clientes_url', formCli.getPublishedUrl());
    Logger.log('  ✓ form_clientes_id:  ' + FORM_CLI_ID);
    Logger.log('  ✓ form_clientes_url: ' + formCli.getPublishedUrl());
    Logger.log('    Destino form Clientes: ' +
               (formCli.getDestinationId() === TARGET_SS_ID ? '✓ V2' : '⚠️ ' + formCli.getDestinationId()));
  } catch (e) {
    Logger.log('  ❌ Error formulario Clientes: ' + e.message);
  }

  try {
    var formOpe = FormApp.openById(FORM_OPE_ID);
    setConfig('form_operarios_id',  FORM_OPE_ID);
    setConfig('form_operarios_url', formOpe.getPublishedUrl());
    Logger.log('  ✓ form_operarios_id:  ' + FORM_OPE_ID);
    Logger.log('  ✓ form_operarios_url: ' + formOpe.getPublishedUrl());
    Logger.log('    Destino form Operarios: ' +
               (formOpe.getDestinationId() === TARGET_SS_ID ? '✓ V2' : '⚠️ ' + formOpe.getDestinationId()));
  } catch (e) {
    Logger.log('  ❌ Error formulario Operarios: ' + e.message);
  }

  // ── 7. Estado final ───────────────────────────────────────────
  Logger.log('\n[7] Estado final de hojas:');
  var finalSheets = ss.getSheets();
  for (var i = 0; i < finalSheets.length; i++) {
    var fs = finalSheets[i];
    Logger.log('  "' + fs.getName() + '"  GID=' + fs.getSheetId() +
               '  filas=' + Math.max(0, fs.getLastRow() - 1));
  }

  Logger.log('\n================================================');
  Logger.log('REPARACIÓN COMPLETADA.');
  Logger.log('Siguiente paso: validateFreshInstall()');
  Logger.log('================================================');
}

/**
 * PASO 1 (opcional): Renombra el spreadsheet activo a BACKUP
 * y elimina todos los triggers del proyecto.
 *
 * Ejecutar ANTES de freshInstall() para archivar el sistema anterior.
 * NO modifica hojas ni datos.
 */
function backupOldSystem() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var oldName = ss.getName();

  Logger.log('================================================');
  Logger.log('backupOldSystem()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  Logger.log('\n[1] Spreadsheet actual:');
  Logger.log('  Nombre: "' + oldName + '"');
  Logger.log('  ID:     ' + ss.getId());
  Logger.log('  URL:    ' + ss.getUrl());

  ss.rename('DB_Trabajos_Alvaro_BACKUP');
  Logger.log('\n[2] Renombrado → "DB_Trabajos_Alvaro_BACKUP"');

  var triggers = ScriptApp.getProjectTriggers();
  Logger.log('\n[3] Triggers encontrados: ' + triggers.length);
  for (var i = 0; i < triggers.length; i++) {
    Logger.log('  Eliminando: ' + triggers[i].getHandlerFunction() +
               '  (' + triggers[i].getEventType() + ')');
    ScriptApp.deleteTrigger(triggers[i]);
  }
  Logger.log('  ✓ Todos los triggers eliminados');

  Logger.log('\n================================================');
  Logger.log('BACKUP COMPLETADO.');
  Logger.log('El spreadsheet BACKUP queda intacto. No lo modifiques.');
  Logger.log('Siguiente paso: ejecuta freshInstall()');
  Logger.log('================================================');
}

// ─────────────────────────────────────────────────────────────

/**
 * PASO 2: Crea el sistema nuevo desde cero.
 *
 * Crea: spreadsheet DB_Trabajos_Alvaro_V2, 8 hojas del sistema,
 * CONFIG con valores por defecto, carpetas Drive, 2 formularios
 * vinculados a RESP_CLIENTES / RESP_OPERARIOS, trigger onFormSubmit.
 *
 * Guarda TARGET_SS_ID en Script Properties para que todas las
 * funciones del proyecto apunten al nuevo spreadsheet.
 *
 * NO reutiliza nada del sistema anterior.
 */
function freshInstall() {
  Logger.log('================================================');
  Logger.log('freshInstall()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  // ── 1. Crear spreadsheet nuevo ───────────────────────────────
  Logger.log('\n[1] Creando spreadsheet...');
  var newSs   = SpreadsheetApp.create('DB_Trabajos_Alvaro_V2');
  var newSsId = newSs.getId();
  Logger.log('  Nombre: "DB_Trabajos_Alvaro_V2"');
  Logger.log('  ID:     ' + newSsId);
  Logger.log('  URL:    ' + newSs.getUrl());

  // ── 2. Guardar ID en Script Properties ──────────────────────
  PropertiesService.getScriptProperties().setProperty('TARGET_SS_ID', newSsId);
  Logger.log('\n[2] TARGET_SS_ID guardado en Script Properties');
  Logger.log('  A partir de aquí, todas las funciones apuntan al nuevo spreadsheet.');

  // ── 3. Crear hojas del sistema ───────────────────────────────
  Logger.log('\n[3] Creando hojas del sistema:');
  _installCreateSheets_(newSs);

  // ── 4. Inicializar CONFIG ────────────────────────────────────
  Logger.log('\n[4] Inicializando CONFIG...');
  initConfig();
  Logger.log('  ✓ CONFIG inicializada con valores por defecto');

  // ── 5. Carpetas Drive ────────────────────────────────────────
  Logger.log('\n[5] Creando carpetas Drive...');
  createDriveFolders_();

  // ── 6. Formularios ───────────────────────────────────────────
  Logger.log('\n[6] Creando formularios...');
  _installCreateClienteForm_(newSs, newSsId);
  _installCreateOperarioForm_(newSs, newSsId);

  // ── 7. Trigger ───────────────────────────────────────────────
  Logger.log('\n[7] Instalando trigger onFormSubmit...');
  var existingTrg = ScriptApp.getProjectTriggers().filter(function(t) {
    return t.getHandlerFunction() === 'onFormSubmitHandler';
  });
  if (existingTrg.length === 0) {
    ScriptApp.newTrigger('onFormSubmitHandler')
      .forSpreadsheet(newSs)
      .onFormSubmit()
      .create();
    Logger.log('  ✓ Trigger instalado en el nuevo spreadsheet');
  } else {
    Logger.log('  ℹ️ Trigger ya existía (' + existingTrg.length + ')');
  }

  // ── 8. Formato ───────────────────────────────────────────────
  Logger.log('\n[8] Aplicando formato a las hojas...');
  applySheetFormatting_();
  Logger.log('  ✓ Formato aplicado');

  // ── Resumen ──────────────────────────────────────────────────
  Logger.log('\n================================================');
  Logger.log('FRESH INSTALL COMPLETADO');
  Logger.log('  Spreadsheet:        ' + newSs.getUrl());
  Logger.log('  Spreadsheet ID:     ' + newSsId);
  Logger.log('  Form Clientes  ID:  ' + (getConfig('form_clientes_id')  || 'ERROR — revisar paso 6'));
  Logger.log('  Form Clientes  URL: ' + (getConfig('form_clientes_url') || 'ERROR — revisar paso 6'));
  Logger.log('  Form Operarios ID:  ' + (getConfig('form_operarios_id') || 'ERROR — revisar paso 6'));
  Logger.log('  Form Operarios URL: ' + (getConfig('form_operarios_url') || 'ERROR — revisar paso 6'));
  Logger.log('================================================');
  Logger.log('Siguiente paso: validateFreshInstall()');
}

// ─────────────────────────────────────────────────────────────

/**
 * PASO 3: Valida que el sistema nuevo está correctamente instalado.
 *
 * Comprueba: spreadsheet correcto, 10 hojas presentes, sin hojas extra,
 * formularios accesibles y vinculados, trigger activo, CONFIG con datos.
 */
function validateFreshInstall() {
  Logger.log('================================================');
  Logger.log('validateFreshInstall()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  // ── 1. TARGET_SS_ID ──────────────────────────────────────────
  var ssId = PropertiesService.getScriptProperties().getProperty('TARGET_SS_ID');
  Logger.log('\n[1] TARGET_SS_ID: ' + (ssId || 'NO CONFIGURADO ❌'));
  if (!ssId) {
    Logger.log('ABORTANDO. Ejecuta freshInstall() primero.');
    return;
  }

  var ss;
  try {
    ss = SpreadsheetApp.openById(ssId);
  } catch (e) {
    Logger.log('ERROR abriendo spreadsheet: ' + e.message + ' ❌');
    return;
  }

  var nombreCorrecto = ss.getName() === 'DB_Trabajos_Alvaro_V2';
  Logger.log('  Nombre: "' + ss.getName() + '"  ' + (nombreCorrecto ? '✓' : '⚠️ esperado "DB_Trabajos_Alvaro_V2"'));
  Logger.log('  URL:    ' + ss.getUrl());

  // ── 2. Hojas requeridas ───────────────────────────────────────
  Logger.log('\n[2] Hojas requeridas (10):');
  var required = [
    'CONFIG', 'CLIENTES', 'OPERARIOS', 'TRABAJOS', 'FACTURAS',
    'PENDIENTES', 'COMPLETADOS', 'LOGS', 'RESP_CLIENTES', 'RESP_OPERARIOS'
  ];
  var sheetsOk = true;
  for (var i = 0; i < required.length; i++) {
    var s = ss.getSheetByName(required[i]);
    if (s) {
      Logger.log('  ✓ ' + required[i] + '  GID=' + s.getSheetId() +
                 '  filas=' + Math.max(0, s.getLastRow() - 1));
    } else {
      Logger.log('  ❌ ' + required[i] + ': NO EXISTE');
      sheetsOk = false;
    }
  }

  // ── 3. Hojas extra ───────────────────────────────────────────
  Logger.log('\n[3] Hojas extra (deben ser 0):');
  var extras = ss.getSheets().filter(function(sh) {
    return required.indexOf(sh.getName()) === -1;
  });
  if (extras.length === 0) {
    Logger.log('  ✓ Sin hojas extra');
  } else {
    for (var i = 0; i < extras.length; i++) {
      Logger.log('  ⚠️ Extra: "' + extras[i].getName() + '"  GID=' + extras[i].getSheetId());
    }
  }

  // ── 4. IDs de formularios en CONFIG ──────────────────────────
  Logger.log('\n[4] Formularios en CONFIG:');
  var cliId  = getConfig('form_clientes_id');
  var cliUrl = getConfig('form_clientes_url');
  var opeId  = getConfig('form_operarios_id');
  var opeUrl = getConfig('form_operarios_url');
  Logger.log('  form_clientes_id:   ' + (cliId  || '❌ NO CONFIGURADO'));
  Logger.log('  form_clientes_url:  ' + (cliUrl || '❌ NO CONFIGURADO'));
  Logger.log('  form_operarios_id:  ' + (opeId  || '❌ NO CONFIGURADO'));
  Logger.log('  form_operarios_url: ' + (opeUrl || '❌ NO CONFIGURADO'));

  // ── 5. Verificación de formularios ───────────────────────────
  Logger.log('\n[5] Verificación de formularios:');
  _installValidateForm_(cliId, ssId, 'Clientes',  'RESP_CLIENTES',  ss);
  _installValidateForm_(opeId, ssId, 'Operarios', 'RESP_OPERARIOS', ss);

  // ── 6. Trigger ───────────────────────────────────────────────
  Logger.log('\n[6] Trigger onFormSubmitHandler:');
  var trgs = ScriptApp.getProjectTriggers().filter(function(t) {
    return t.getHandlerFunction() === 'onFormSubmitHandler';
  });
  if (trgs.length === 1) {
    Logger.log('  ✓ 1 trigger activo (correcto)');
  } else if (trgs.length === 0) {
    Logger.log('  ❌ Sin trigger — ejecuta freshInstall() o setupTriggers()');
  } else {
    Logger.log('  ⚠️ ' + trgs.length + ' triggers (debería ser exactamente 1)');
  }

  // ── 7. CONFIG empresa ────────────────────────────────────────
  Logger.log('\n[7] CONFIG — datos de empresa (rellenar manualmente):');
  var claves = ['empresa_nombre', 'empresa_cif', 'empresa_email',
                'empresa_telefono', 'iva_porcentaje', 'irpf_porcentaje', 'serie_factura'];
  for (var i = 0; i < claves.length; i++) {
    var v = getConfig(claves[i]);
    Logger.log('  ' + claves[i] + ': ' + (v !== null ? '"' + v + '"' : '(vacío)'));
  }

  // ── Resumen ──────────────────────────────────────────────────
  var formsOk   = !!(cliId && opeId);
  var triggerOk = trgs.length === 1;

  Logger.log('\n================================================');
  Logger.log('RESULTADO DE VALIDACIÓN:');
  Logger.log('  Spreadsheet: ' + (nombreCorrecto ? '✓' : '⚠️'));
  Logger.log('  Hojas (10):  ' + (sheetsOk ? '✓' : '❌'));
  Logger.log('  Extras:      ' + (extras.length === 0 ? '✓' : '⚠️ hay ' + extras.length));
  Logger.log('  Formularios: ' + (formsOk ? '✓' : '❌'));
  Logger.log('  Trigger:     ' + (triggerOk ? '✓' : '❌'));
  Logger.log('================================================');
  if (sheetsOk && formsOk && triggerOk) {
    Logger.log('SISTEMA LISTO.');
    Logger.log('Abre el spreadsheet nuevo y completa CONFIG con los datos de tu empresa.');
    Logger.log('URL: ' + ss.getUrl());
  } else {
    Logger.log('HAY PROBLEMAS. Revisa los puntos marcados con ❌ arriba.');
  }
  Logger.log('================================================');
}

// ─────────────────────────────────────────────────────────────

/**
 * Instala el módulo de Trabajos sobre el spreadsheet V2 activo.
 *
 * Crea: formulario Trabajos con 10 campos (dropdowns de CLIENTES y OPERARIOS),
 * hoja RESP_TRABAJOS vinculada, actualiza cabeceras de TRABAJOS, guarda IDs en CONFIG.
 *
 * El trigger onFormSubmitHandler ya instalado en el spreadsheet enruta
 * automáticamente RESP_TRABAJOS → onTrabajosFormSubmit. No crea trigger nuevo.
 *
 * NO toca: CONFIG, CLIENTES, OPERARIOS, RESP_CLIENTES, RESP_OPERARIOS.
 * NO crea spreadsheet nuevo. NO crea formularios de Clientes/Operarios.
 */
function installTrabajosModule() {
  Logger.log('================================================');
  Logger.log('installTrabajosModule()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var ss   = getTargetSS_();
  var ssId = ss.getId();
  Logger.log('\n[0] Spreadsheet: "' + ss.getName() + '"  ID=' + ssId);

  // ── 1. Actualizar cabeceras de TRABAJOS ───────────────────────
  Logger.log('\n[1] Hoja TRABAJOS:');
  var trabajosSheet = ss.getSheetByName('TRABAJOS');
  if (!trabajosSheet) {
    Logger.log('  ❌ TRABAJOS no existe. Ejecuta freshInstall() primero.');
    return;
  }
  if (trabajosSheet.getLastRow() <= 1) {
    trabajosSheet.clearContents();
    trabajosSheet.appendRow(HEADERS_TRABAJOS_V2);
    applyHeaderStyle_(trabajosSheet, HEADERS_TRABAJOS_V2.length);
    Logger.log('  ✓ Cabeceras escritas: ' + HEADERS_TRABAJOS_V2.join(' | '));
  } else {
    Logger.log('  ⚠️ TRABAJOS tiene datos — cabeceras no modificadas.');
    Logger.log('  Cabeceras actuales: ' +
      trabajosSheet.getRange(1, 1, 1, trabajosSheet.getLastColumn()).getValues()[0].join(' | '));
  }

  // ── 2. Leer listas para dropdowns ────────────────────────────
  Logger.log('\n[2] Leyendo CLIENTES para dropdown:');
  var clientesList = _getTrabajosDropdownClientes_();
  Logger.log('  Opciones: ' + clientesList.length);
  if (clientesList.length === 0) {
    Logger.log('  ⚠️ Sin clientes activos — dropdown tendrá placeholder.');
    clientesList = ['(añade clientes primero)'];
  } else {
    for (var i = 0; i < clientesList.length; i++) {
      Logger.log('  - ' + clientesList[i]);
    }
  }

  Logger.log('\n[3] Leyendo OPERARIOS para dropdown:');
  var operariosList = _getTrabajosDropdownOperarios_();
  Logger.log('  Opciones: ' + operariosList.length);
  if (operariosList.length === 0) {
    Logger.log('  ⚠️ Sin operarios activos — dropdown tendrá placeholder.');
    operariosList = ['(añade operarios primero)'];
  } else {
    for (var i = 0; i < operariosList.length; i++) {
      Logger.log('  - ' + operariosList[i]);
    }
  }

  // ── 3. Crear formulario ───────────────────────────────────────
  Logger.log('\n[4] Creando formulario Trabajos...');
  var form = FormApp.create('Parte de Trabajo — Álvaro Gestión');
  form.setDescription('Registra un trabajo o servicio realizado.');
  form.setCollectEmail(false);
  form.setAllowResponseEdits(false);

  form.addListItem()
    .setTitle('Cliente')
    .setChoiceValues(clientesList)
    .setRequired(true);

  form.addListItem()
    .setTitle('Operario')
    .setChoiceValues(operariosList)
    .setRequired(true);

  form.addDateItem()
    .setTitle('Fecha')
    .setRequired(true);

  form.addTimeItem()
    .setTitle('Hora inicio')
    .setRequired(true);

  form.addTimeItem()
    .setTitle('Hora fin')
    .setRequired(false);

  form.addParagraphTextItem()
    .setTitle('Trabajo realizado')
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('Incidencias')
    .setRequired(false);

  form.addMultipleChoiceItem()
    .setTitle('Estado')
    .setChoiceValues(ESTADOS_TRABAJO)
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('Fotos')
    .setHelpText('Pega los enlaces de las fotos (Google Drive, etc.)')
    .setRequired(false);

  form.addTextItem()
    .setTitle('Materiales (€)')
    .setHelpText('Coste de materiales empleados en este trabajo (ej: 25.50)')
    .setRequired(false);

  form.addTextItem()
    .setTitle('Ubicación')
    .setHelpText('Dirección o descripción del lugar del trabajo')
    .setRequired(false);

  Logger.log('  Formulario creado: "' + form.getTitle() + '"');
  Logger.log('  ID: ' + form.getId());
  Logger.log('  Campos: ' + form.getItems().length);

  // ── 4. Vincular y renombrar hoja de respuestas ────────────────
  Logger.log('\n[5] Vinculando al spreadsheet V2...');
  var snapshot = ss.getSheets().map(function(s) { return s.getName(); });
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ssId);
  Logger.log('  Esperando RESP_TRABAJOS...');
  var ok = _installWaitAndRename_(ss, snapshot, 'RESP_TRABAJOS');
  if (ok) {
    Logger.log('  ✓ RESP_TRABAJOS creada y vinculada');
  } else {
    Logger.log('  ❌ RESP_TRABAJOS no apareció en 16 s.');
    Logger.log('  Acción manual: renombra la nueva hoja a "RESP_TRABAJOS".');
    logEvent('ERROR', 'installTrabajosModule', 'RESP_TRABAJOS no detectada', '');
  }

  // ── 5. Guardar en CONFIG ──────────────────────────────────────
  Logger.log('\n[6] Guardando en CONFIG:');
  setConfig('form_trabajos_id',  form.getId());
  setConfig('form_trabajos_url', form.getPublishedUrl());
  Logger.log('  ✓ form_trabajos_id:  ' + form.getId());
  Logger.log('  ✓ form_trabajos_url: ' + form.getPublishedUrl());

  // ── Resumen ───────────────────────────────────────────────────
  Logger.log('\n================================================');
  Logger.log('MÓDULO TRABAJOS INSTALADO');
  Logger.log('  Form Trabajos ID:  ' + form.getId());
  Logger.log('  Form Trabajos URL: ' + form.getPublishedUrl());
  Logger.log('  IDs auto: TRA-001, TRA-002, ...');
  Logger.log('  NOTA: dropdowns Cliente/Operario son estáticos.');
  Logger.log('  Si añades clientes/operarios: ejecuta updateTrabajosFormDropdowns()');
  Logger.log('================================================');
  Logger.log('Siguiente paso: validateTrabajosModule()');
}

// ─────────────────────────────────────────────────────────────

/**
 * Valida que el módulo de Trabajos está correctamente instalado.
 */
function validateTrabajosModule() {
  Logger.log('================================================');
  Logger.log('validateTrabajosModule()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var ss   = getTargetSS_();
  var ssId = ss.getId();

  // 1. Hoja TRABAJOS
  Logger.log('\n[1] Hoja TRABAJOS:');
  var trabajosSheet = ss.getSheetByName('TRABAJOS');
  if (trabajosSheet) {
    var filas = Math.max(0, trabajosSheet.getLastRow() - 1);
    var hdrRow = trabajosSheet.getLastColumn() > 0
      ? trabajosSheet.getRange(1, 1, 1, trabajosSheet.getLastColumn()).getValues()[0]
      : [];
    Logger.log('  ✓ GID=' + trabajosSheet.getSheetId() + '  filas=' + filas);
    Logger.log('  Cabeceras: ' + hdrRow.join(' | '));
    // Verify expected headers
    var expectedCount = HEADERS_TRABAJOS_V2.filter(function(h) {
      return hdrRow.indexOf(h) !== -1;
    }).length;
    Logger.log('  Coincidencia: ' + expectedCount + '/' + HEADERS_TRABAJOS_V2.length +
               (expectedCount === HEADERS_TRABAJOS_V2.length ? ' ✓' : ' ⚠️'));
  } else {
    Logger.log('  ❌ NO EXISTE');
  }

  // 2. Hoja RESP_TRABAJOS
  Logger.log('\n[2] Hoja RESP_TRABAJOS:');
  var respSheet = ss.getSheetByName('RESP_TRABAJOS');
  if (respSheet) {
    Logger.log('  ✓ GID=' + respSheet.getSheetId() +
               '  filas=' + Math.max(0, respSheet.getLastRow() - 1));
  } else {
    Logger.log('  ❌ NO EXISTE');
  }

  // 3. CONFIG
  Logger.log('\n[3] CONFIG — módulo Trabajos:');
  var formId  = getConfig('form_trabajos_id');
  var formUrl = getConfig('form_trabajos_url');
  Logger.log('  form_trabajos_id:  ' + (formId  || '❌ NO CONFIGURADO'));
  Logger.log('  form_trabajos_url: ' + (formUrl || '❌ NO CONFIGURADO'));

  // 4. Formulario accesible y vinculado
  Logger.log('\n[4] Formulario Trabajos:');
  if (formId) {
    try {
      var form   = FormApp.openById(formId);
      var destId = form.getDestinationId();
      Logger.log('  Título:  "' + form.getTitle() + '"');
      Logger.log('  ID:      ' + formId);
      Logger.log('  Destino: ' + (destId === ssId ? '✓ V2 correcto' : '❌ ' + (destId || 'SIN DESTINO')));
      Logger.log('  URL:     ' + form.getPublishedUrl());
      var items = form.getItems();
      Logger.log('  Campos (' + items.length + '):');
      for (var i = 0; i < items.length; i++) {
        Logger.log('    ' + (i + 1) + '. [' + items[i].getType() + '] ' + items[i].getTitle());
      }
    } catch (e) {
      Logger.log('  ❌ Error: ' + e.message);
    }
  } else {
    Logger.log('  ❌ Sin form_trabajos_id en CONFIG');
  }

  // 5. Trigger
  Logger.log('\n[5] Trigger onFormSubmitHandler:');
  var trgs = ScriptApp.getProjectTriggers().filter(function(t) {
    return t.getHandlerFunction() === 'onFormSubmitHandler';
  });
  Logger.log('  Activos: ' + trgs.length + (trgs.length === 1 ? ' ✓' : ' ⚠️ (debe ser 1)'));

  // 6. Routing RESP_TRABAJOS
  Logger.log('\n[6] Routing RESP_TRABAJOS → onTrabajosFormSubmit:');
  Logger.log('  onTrabajosFormSubmit: ' +
    (typeof onTrabajosFormSubmit === 'function' ? '✓ definida en Trabajos.gs' : '❌ no encontrada'));

  // Resumen
  var ok = !!(trabajosSheet && respSheet && formId && trgs.length === 1);
  Logger.log('\n================================================');
  Logger.log('RESULTADO:');
  Logger.log('  TRABAJOS:          ' + (trabajosSheet          ? '✓' : '❌'));
  Logger.log('  RESP_TRABAJOS:     ' + (respSheet              ? '✓' : '❌'));
  Logger.log('  CONFIG:            ' + (formId                 ? '✓' : '❌'));
  Logger.log('  Trigger:           ' + (trgs.length === 1      ? '✓' : '⚠️'));
  Logger.log('  Handler:           ' +
    (typeof onTrabajosFormSubmit === 'function' ? '✓' : '❌'));
  Logger.log(ok ? 'MÓDULO LISTO. Envía un formulario de prueba.' : 'HAY PROBLEMAS — revisar ❌ arriba.');
  Logger.log('================================================');
}

// ── Funciones privadas de instalación ────────────────────────

function _installCreateSheets_(ss) {
  var systemSheets = [
    ['CONFIG',      SHEET_HEADERS.CONFIG],
    ['CLIENTES',    SHEET_HEADERS.CLIENTES],
    ['OPERARIOS',   SHEET_HEADERS.OPERARIOS],
    ['TRABAJOS',    SHEET_HEADERS.TRABAJOS],
    ['FACTURAS',    SHEET_HEADERS.FACTURAS],
    ['PENDIENTES',  SHEET_HEADERS.PENDIENTES],
    ['COMPLETADOS', SHEET_HEADERS.COMPLETADOS],
    ['LOGS',        SHEET_HEADERS.LOGS]
  ];

  for (var i = 0; i < systemSheets.length; i++) {
    var name    = systemSheets[i][0];
    var headers = systemSheets[i][1];
    var sheet   = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sheet.getLastRow() === 0) sheet.appendRow(headers);
    Logger.log('  ✓ ' + name);
  }

  // Eliminar hoja vacía por defecto de Google Sheets
  var defaults = ['Sheet1', 'Hoja 1', 'Hoja1'];
  for (var d = 0; d < defaults.length; d++) {
    var def = ss.getSheetByName(defaults[d]);
    if (def && ss.getSheets().length > 1) {
      try { ss.deleteSheet(def); } catch (_) {}
    }
  }

  Logger.log('  8 hojas del sistema creadas.');
  Logger.log('  RESP_CLIENTES y RESP_OPERARIOS se crean al vincular formularios (paso 6).');
}

function _installCreateClienteForm_(ss, ssId) {
  Logger.log('  --- Formulario Clientes ---');
  var form = FormApp.create('Alta de Cliente — Álvaro Gestión');
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
  Logger.log('  Formulario creado. ID=' + form.getId());

  var snapshot = ss.getSheets().map(function(s) { return s.getName(); });
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ssId);

  var ok = _installWaitAndRename_(ss, snapshot, 'RESP_CLIENTES');
  if (ok) {
    setConfig('form_clientes_id',  form.getId());
    setConfig('form_clientes_url', form.getPublishedUrl());
    Logger.log('  ✓ RESP_CLIENTES creada y vinculada');
    Logger.log('  ID:  ' + form.getId());
    Logger.log('  URL: ' + form.getPublishedUrl());
  } else {
    Logger.log('  ❌ RESP_CLIENTES no apareció tras 16 s. Revisar manualmente.');
    logEvent('ERROR', '_installCreateClienteForm_', 'Hoja RESP_CLIENTES no detectada', '');
  }
}

function _installCreateOperarioForm_(ss, ssId) {
  Logger.log('  --- Formulario Operarios ---');
  var form = FormApp.create('Alta de Operario — Álvaro Gestión');
  form.setDescription('Registra un nuevo operario en el sistema.');
  form.setCollectEmail(false);
  form.setAllowResponseEdits(false);
  form.addTextItem().setTitle('Nombre').setRequired(true);
  form.addTextItem().setTitle('Apellidos').setRequired(true);
  form.addTextItem().setTitle('Email').setRequired(true);
  form.addTextItem().setTitle('Teléfono').setRequired(true);
  Logger.log('  Formulario creado. ID=' + form.getId());

  var snapshot = ss.getSheets().map(function(s) { return s.getName(); });
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ssId);

  var ok = _installWaitAndRename_(ss, snapshot, 'RESP_OPERARIOS');
  if (ok) {
    setConfig('form_operarios_id',  form.getId());
    setConfig('form_operarios_url', form.getPublishedUrl());
    Logger.log('  ✓ RESP_OPERARIOS creada y vinculada');
    Logger.log('  ID:  ' + form.getId());
    Logger.log('  URL: ' + form.getPublishedUrl());
  } else {
    Logger.log('  ❌ RESP_OPERARIOS no apareció tras 16 s. Revisar manualmente.');
    logEvent('ERROR', '_installCreateOperarioForm_', 'Hoja RESP_OPERARIOS no detectada', '');
  }
}

function _installWaitAndRename_(ss, snapshot, targetName) {
  for (var attempt = 0; attempt < 8; attempt++) {
    Utilities.sleep(2000);
    var current = ss.getSheets();
    for (var i = 0; i < current.length; i++) {
      var s = current[i];
      var n = s.getName();
      if (snapshot.indexOf(n) === -1 && n !== targetName) {
        s.setName(targetName);
        Logger.log('  "' + n + '" → "' + targetName + '"');
        return true;
      }
    }
    if (ss.getSheetByName(targetName)) return true;
  }
  return false;
}

function _getTrabajosDropdownClientes_() {
  try {
    var sheet   = getSheet('CLIENTES');
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    var data    = sheet.getRange(2, 1, lastRow - 1, COL_CLI.ACTIVO).getValues();
    var options = [];
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][COL_CLI.ACTIVO - 1]).toLowerCase() !== 'true') continue;
      var nombre    = String(data[i][COL_CLI.NOMBRE    - 1] || '').trim();
      var apellidos = String(data[i][COL_CLI.APELLIDOS - 1] || '').trim();
      var empresa   = String(data[i][COL_CLI.EMPRESA   - 1] || '').trim();
      var label     = (nombre + ' ' + apellidos).trim();
      if (empresa) label += ' (' + empresa + ')';
      if (label) options.push(label);
    }
    // Eliminar duplicados por nombre normalizado (trim+lowercase+sin tildes)
    var seenKeys = {}, deduped = [];
    options.forEach(function(lbl) {
      var key = _normalizeClientLabel_(lbl);
      if (!seenKeys[key]) { seenKeys[key] = true; deduped.push(lbl); }
    });
    return deduped;
  } catch (e) {
    Logger.log('  Error leyendo CLIENTES: ' + e.message);
    return [];
  }
}

function _getTrabajosDropdownOperarios_() {
  try {
    var sheet   = getSheet('OPERARIOS');
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    var data    = sheet.getRange(2, 1, lastRow - 1, COL_OPE.ACTIVO).getValues();
    var options = [];
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][COL_OPE.ACTIVO - 1]).toLowerCase() !== 'true') continue;
      var nombre    = String(data[i][COL_OPE.NOMBRE    - 1] || '').trim();
      var apellidos = String(data[i][COL_OPE.APELLIDOS - 1] || '').trim();
      var label = (nombre + ' ' + apellidos).trim();
      if (label) options.push(label);
    }
    // Eliminar duplicados
    return [...new Set(options)];
  } catch (e) {
    Logger.log('  Error leyendo OPERARIOS: ' + e.message);
    return [];
  }
}

function _installValidateForm_(formId, ssId, label, expectedSheet, ss) {
  if (!formId) {
    Logger.log('  ' + label + ': ❌ ID no configurado en CONFIG');
    return;
  }
  try {
    var form   = FormApp.openById(formId);
    var destId = form.getDestinationId();
    var destOk = destId === ssId;
    var sheetObj = ss.getSheetByName(expectedSheet);
    Logger.log('  ' + label + ':');
    Logger.log('    ID:      ' + formId);
    Logger.log('    Título:  "' + form.getTitle() + '"');
    Logger.log('    Destino: ' + (destOk
      ? '✓ este spreadsheet'
      : '❌ ' + (destId || 'SIN DESTINO')));
    Logger.log('    Hoja "' + expectedSheet + '": ' + (sheetObj
      ? '✓ GID=' + sheetObj.getSheetId()
      : '❌ NO EXISTE'));
    Logger.log('    URL: ' + form.getPublishedUrl());
  } catch (e) {
    Logger.log('  ' + label + ': ❌ ERROR — ' + e.message);
  }
}

// ─────────────────────────────────────────────────────────────

/**
 * Verifica que los tres formularios abren limpios:
 *   - URLs sin parámetros prefill (?entry.XXXXX=valor)
 *   - URL publicada termina en /viewform sin query string
 *   - Ningún campo tiene valor por defecto o opción preseleccionada
 *   - Los desplegables Cliente/Operario no tienen opción marcada por defecto
 *
 * No modifica nada.
 */
function validateFormsOpenBlank() {
  Logger.log('================================================');
  Logger.log('validateFormsOpenBlank()');
  Logger.log(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  Logger.log('================================================');

  var forms = [
    { label: 'Clientes',  configKey: 'form_clientes_id',  urlKey: 'form_clientes_url'  },
    { label: 'Operarios', configKey: 'form_operarios_id', urlKey: 'form_operarios_url' },
    { label: 'Trabajos',  configKey: 'form_trabajos_id',  urlKey: 'form_trabajos_url'  }
  ];

  var globalOk = true;

  for (var f = 0; f < forms.length; f++) {
    var def    = forms[f];
    var formId = getConfig(def.configKey);
    var urlCfg = getConfig(def.urlKey) || '';

    Logger.log('\n── Formulario ' + def.label + ' ──────────────────────');

    if (!formId) {
      Logger.log('  ❌ ' + def.configKey + ' no configurado en CONFIG.');
      globalOk = false;
      continue;
    }

    // 1. Verificar URL almacenada en CONFIG
    Logger.log('  [URL CONFIG] ' + (urlCfg || '(vacía)'));
    var urlOk = _checkUrlClean_(urlCfg, def.label + ' CONFIG');

    // 2. Abrir formulario y verificar URL publicada
    var form;
    try {
      form = FormApp.openById(formId);
    } catch (e) {
      Logger.log('  ❌ No se puede abrir el formulario: ' + e.message);
      globalOk = false;
      continue;
    }

    var publishedUrl = form.getPublishedUrl();
    Logger.log('  [URL PUBL.]  ' + publishedUrl);
    var pubOk = _checkUrlClean_(publishedUrl, def.label + ' publicada');

    // 3. Revisar campos del formulario
    Logger.log('  [CAMPOS]');
    var items     = form.getItems();
    var camposOk  = true;

    for (var i = 0; i < items.length; i++) {
      var item  = items[i];
      var tipo  = item.getType();
      var titulo = item.getTitle();
      var issues = [];

      // Texto simple
      if (tipo === FormApp.ItemType.TEXT) {
        var def_ = item.asTextItem().getDefaultValue ? item.asTextItem().getDefaultValue() : '';
        if (def_ && String(def_).trim() !== '') issues.push('valor por defecto: "' + def_ + '"');
      }
      // Texto largo
      if (tipo === FormApp.ItemType.PARAGRAPH_TEXT) {
        var def_ = item.asParagraphTextItem().getDefaultValue ? item.asParagraphTextItem().getDefaultValue() : '';
        if (def_ && String(def_).trim() !== '') issues.push('valor por defecto: "' + def_ + '"');
      }
      // Desplegable (LIST)
      if (tipo === FormApp.ItemType.LIST) {
        var choices = item.asListItem().getChoices();
        var presel  = choices.filter(function(c) { return c.isSelected && c.isSelected(); });
        if (presel.length > 0) {
          issues.push('opción preseleccionada: "' + presel.map(function(c) { return c.getValue(); }).join(', ') + '"');
        }
      }
      // Selección múltiple
      if (tipo === FormApp.ItemType.MULTIPLE_CHOICE) {
        var choices = item.asMultipleChoiceItem().getChoices();
        var presel  = choices.filter(function(c) { return c.isSelected && c.isSelected(); });
        if (presel.length > 0) {
          issues.push('opción preseleccionada: "' + presel.map(function(c) { return c.getValue(); }).join(', ') + '"');
        }
      }

      if (issues.length > 0) {
        Logger.log('    ⚠️  [' + tipo + '] "' + titulo + '" → ' + issues.join('; '));
        camposOk  = false;
        globalOk  = false;
      } else {
        Logger.log('    ✓  [' + tipo + '] "' + titulo + '"');
      }
    }

    var formOk = urlOk && pubOk && camposOk;
    Logger.log('  → ' + def.label + ': ' + (formOk ? '✓ LIMPIO' : '⚠️  HAY VALORES O URLs CON PARÁMETROS'));
    if (!formOk) globalOk = false;
  }

  Logger.log('\n================================================');
  Logger.log('RESULTADO GLOBAL: ' + (globalOk ? '✓ TODOS LOS FORMULARIOS ABREN LIMPIOS.' : '⚠️  REVISAR LOS PUNTOS MARCADOS ARRIBA.'));
  Logger.log('================================================');
}

/**
 * Comprueba que una URL no contiene parámetros prefill (entry.XXXXX).
 * Devuelve true si está limpia.
 */
function _checkUrlClean_(url, label) {
  if (!url) {
    Logger.log('    ⚠️  ' + label + ': URL vacía.');
    return false;
  }
  var hasPrefill = url.indexOf('entry.') !== -1;
  var hasUspPP   = url.indexOf('usp=pp_url') !== -1;
  var endsOk     = url.indexOf('/viewform') !== -1;

  if (hasPrefill || hasUspPP) {
    Logger.log('    ❌ ' + label + ': URL contiene parámetros prefill → ' + url);
    return false;
  }
  if (!endsOk) {
    Logger.log('    ⚠️  ' + label + ': URL no contiene /viewform → ' + url);
    return false;
  }
  Logger.log('    ✓  ' + label + ': URL limpia.');
  return true;
}
