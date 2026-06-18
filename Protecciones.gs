// ============================================================
// Protecciones.gs — FASE F: Anti-duplicación, LockService,
// idempotencia y control de concurrencia.
// ============================================================

// ============================================================
// 1. IDEMPOTENCIA — anti-retry de triggers y anti-doble submit
// ============================================================

/**
 * Comprueba si un evento ya fue procesado (para evitar retries de Apps Script).
 * Usa ScriptProperties con TTL de 24 horas.
 *
 * @param {string} eventKey — clave única del evento (hash del contenido)
 * @returns {boolean} true si ya fue procesado
 */
function _prot_esEventoProcesado_(eventKey) {
  if (!eventKey) return false;
  var props = PropertiesService.getScriptProperties();
  var key   = 'EVT_' + _prot_sanitizeKey_(eventKey);
  var valor = props.getProperty(key);
  if (!valor) return false;

  // Verificar TTL (24 horas)
  var tsGuardado = parseInt(valor, 10);
  if (Date.now() - tsGuardado > 86400000) {
    props.deleteProperty(key);
    return false;
  }
  return true;
}

/**
 * Marca un evento como procesado. Llamar JUSTO ANTES de procesar,
 * no después, para prevenir retries paralelos.
 */
function _prot_marcarEventoProcesado_(eventKey) {
  if (!eventKey) return;
  var props = PropertiesService.getScriptProperties();
  var key   = 'EVT_' + _prot_sanitizeKey_(eventKey);
  props.setProperty(key, String(Date.now()));
}

/**
 * Elimina marcas de eventos con más de 24 horas.
 * Llamar periódicamente desde un trigger diario.
 */
function limpiarEventosAntiguos() {
  var props = PropertiesService.getScriptProperties();
  var all   = props.getProperties();
  var n = 0;
  Object.keys(all).forEach(function(key) {
    if (key.indexOf('EVT_') === 0) {
      var ts = parseInt(all[key], 10);
      if (Date.now() - ts > 86400000) {
        props.deleteProperty(key);
        n++;
      }
    }
  });
  if (n > 0) logEvent('INFO', 'limpiarEventosAntiguos', n + ' eventos antiguos eliminados', '');
}

/**
 * Genera una clave de deduplicación a partir de los valores de un formulario.
 * Combina los campos más identificativos del trabajo para detectar envíos dobles.
 */
function _prot_generarHashEvento_(namedValues) {
  var campos = [
    (namedValues['Cliente']          || [''])[0],
    (namedValues['Operario']         || namedValues['Operarios'] || [''])[0],
    (namedValues['Trabajo realizado']|| namedValues['Descripción'] || namedValues['Descripcion'] || [''])[0],
    (namedValues['Hora inicio']      || namedValues['Fecha']     || [''])[0]
  ].join('|').toLowerCase().trim();

  // Hash simple (suma de códigos de carácter)
  var hash = 0;
  for (var i = 0; i < campos.length; i++) {
    hash = ((hash << 5) - hash) + campos.charCodeAt(i);
    hash |= 0; // Convertir a entero 32-bit
  }
  return 'TRA_' + Math.abs(hash).toString(36);
}

function _prot_sanitizeKey_(key) {
  return String(key).replace(/[^a-zA-Z0-9_]/g, '').substring(0, 80);
}

// ============================================================
// 2. VALIDACIÓN DE ID EXISTENTE
// ============================================================

/**
 * Comprueba si un ID ya existe en la columna 1 de una hoja.
 * Usar antes de appendRow() para detectar duplicados.
 *
 * @param {Sheet} sheet
 * @param {string} id
 * @returns {boolean}
 */
function _prot_idExiste_(sheet, id) {
  if (!sheet || !id) return false;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var valores = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return valores.some(function(row) {
    return String(row[0] || '').trim() === String(id).trim();
  });
}

// ============================================================
// 3. NEXTID PROTEGIDO CON LOCKSERVICE
//    (reemplaza la función en Utils.gs al actualizarla)
// ============================================================

/**
 * Versión thread-safe de nextId() con LockService.
 * Llama a esta función en lugar de nextId() para mayor seguridad.
 * Misma firma: nextIdSafe(sheetName, prefix) → string
 */
function nextIdSafe(sheetName, prefix) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // Esperar hasta 15 segundos para obtener el lock
    var sheet   = getSheet(sheetName);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return prefix + '-001';
    var values  = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    var numbers = values
      .filter(function(v) { return String(v).indexOf(prefix + '-') === 0; })
      .map(function(v) { return parseInt(String(v).split('-').pop(), 10); })
      .filter(function(n) { return !isNaN(n); });
    var next = numbers.length ? Math.max.apply(null, numbers) + 1 : 1;
    return prefix + '-' + String(next).padStart(3, '0');
  } catch (e) {
    logEvent('WARN', 'nextIdSafe', 'Lock timeout o error: ' + e.message, sheetName);
    // Fallback: timestamp como ID para garantizar unicidad aunque no sea secuencial
    return prefix + '-T' + String(Date.now()).slice(-6);
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// ============================================================
// 4. WRAPPER SEGURO PARA ALTA DE TRABAJOS (anti-doble submit)
// ============================================================

/**
 * Versión idempotente del procesamiento del formulario de trabajos.
 * Reemplaza la llamada directa a onTrabajosFormSubmit() en Triggers.gs.
 *
 * CÓMO USAR:
 * En onFormSubmitHandler(), sustituir:
 *   onTrabajosFormSubmit(nv);
 * por:
 *   onTrabajosFormSubmitSafe(nv);
 */
function onTrabajosFormSubmitSafe(namedValues) {
  var hashEvento = _prot_generarHashEvento_(namedValues);

  // Comprobar idempotencia
  if (_prot_esEventoProcesado_(hashEvento)) {
    logEvent('WARN', 'onTrabajosFormSubmitSafe',
      'Evento duplicado detectado y descartado: ' + hashEvento,
      JSON.stringify(Object.keys(namedValues)));
    return { success: false, error: 'Evento duplicado — ya procesado', isDuplicate: true };
  }

  // Marcar como procesado ANTES de procesar (previene retry paralelo)
  _prot_marcarEventoProcesado_(hashEvento);

  // Procesar normalmente
  try {
    onTrabajosFormSubmit(namedValues);
    logEvent('INFO', 'onTrabajosFormSubmitSafe', 'Trabajo registrado OK. Hash: ' + hashEvento, '');
    return { success: true };
  } catch (e) {
    // En caso de error, eliminar la marca para permitir reintentos manuales
    try {
      var props = PropertiesService.getScriptProperties();
      props.deleteProperty('EVT_' + _prot_sanitizeKey_(hashEvento));
    } catch (_) {}
    logEvent('ERROR', 'onTrabajosFormSubmitSafe', e.message, e.stack);
    throw e;
  }
}

// ============================================================
// 5. WRAPPER SEGURO PARA ALTA DE CLIENTES
// ============================================================

/**
 * Versión de createCliente() con LockService y validación de duplicados.
 * Usa esta función en lugar de createCliente() para producción.
 */
function createClienteSafe(data) {
  if (!data.nombre || !data.email) {
    return { success: false, error: 'Faltan campos: nombre, email' };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var sheet = getSheet(SHEET_NAMES.CLIENTES);

    // Verificar duplicado por email (campo único de facto)
    var email = String(data.email).trim().toLowerCase();
    if (email) {
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        var emails = sheet.getRange(2, COL_CLI.EMAIL, lastRow - 1, 1).getValues();
        var emailExiste = emails.some(function(r) {
          return String(r[0] || '').trim().toLowerCase() === email;
        });
        if (emailExiste) {
          return { success: false, error: 'Ya existe un cliente con el email: ' + email, isDuplicate: true };
        }
      }
    }

    // Generar ID con lock activo (atómico)
    var id = nextIdSafe(SHEET_NAMES.CLIENTES, 'CLI');

    // Doble verificación: que el ID no existe
    if (_prot_idExiste_(sheet, id)) {
      return { success: false, error: 'ID generado ya existe: ' + id + '. Intenta de nuevo.', isDuplicate: true };
    }

    sheet.appendRow([
      id, new Date(),
      String(data.nombre    || '').trim(),
      String(data.apellidos || '').trim(),
      String(data.empresa   || '').trim(),
      email,
      String(data.telefono  || '').trim(),
      String(data.direccion || '').trim(),
      String(data.cifNif    || '').trim().toUpperCase(),
      true,
      JSON.stringify(Array.isArray(data.metodosPago) ? data.metodosPago : [])
    ]);

    logEvent('INFO', 'createClienteSafe', 'Cliente creado: ' + id, data.nombre);
    return { success: true, id: id };

  } catch (e) {
    logEvent('ERROR', 'createClienteSafe', e.message, JSON.stringify(data));
    return { success: false, error: e.message };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
    try { sincronizarClientesSilencioso(); } catch (_) {}
  }
}

// ============================================================
// 6. VERIFICACIÓN DE ESTADO DE PROTECCIONES
// ============================================================

function verificarEstadoProtecciones() {
  var ui = SpreadsheetApp.getUi();
  var lines = [];
  lines.push('=== ESTADO DE PROTECCIONES ANTI-DUPLICADO ===\n');

  // 1. Triggers instalados
  var triggers = ScriptApp.getProjectTriggers();
  lines.push('TRIGGERS ACTIVOS (' + triggers.length + '):');
  triggers.forEach(function(t) {
    lines.push('  • ' + t.getHandlerFunction() + ' (' + t.getEventType() + ')');
  });

  var tieneEditClientes = triggers.some(function(t) {
    return t.getHandlerFunction() === 'onEditClientesHandler';
  });
  lines.push('\n  onEditClientesHandler: ' + (tieneEditClientes ? '✅ instalado' : '❌ NO instalado'));

  // 2. Eventos almacenados en ScriptProperties
  var props  = PropertiesService.getScriptProperties().getProperties();
  var nEvts  = Object.keys(props).filter(function(k) { return k.indexOf('EVT_') === 0; }).length;
  lines.push('\nEVENTOS DE IDEMPOTENCIA ALMACENADOS: ' + nEvts);

  // 3. LockService (se asume disponible siempre en Apps Script)
  lines.push('\nLockService: ✅ disponible (getScriptLock)');

  // 4. nextIdSafe
  lines.push('nextIdSafe: ✅ disponible en Protecciones.gs');

  // 5. Advertencias
  lines.push('\nADVERTENCIAS:');
  if (!tieneEditClientes) {
    lines.push('  ⚠ onEditClientesHandler no está instalado.');
    lines.push('    Los clientes añadidos manualmente NO sincronizarán los dropdowns.');
    lines.push('    Ir a: Menú → 🔌 Instalar trigger sincronización CLIENTES');
  }

  var onFormSubmitHandler = triggers.some(function(t) {
    return t.getHandlerFunction() === 'onFormSubmitHandler';
  });
  if (!onFormSubmitHandler) {
    lines.push('  ⚠ onFormSubmitHandler no está instalado.');
    lines.push('    Los formularios de clientes y trabajos NO procesarán respuestas.');
  }

  lines.push('\nRECOMENDACIÓN:');
  lines.push('  Para protección completa contra duplicados en createCliente(),');
  lines.push('  sustituir la llamada en onClienteFormSubmit() por createClienteSafe().');
  lines.push('  Para trabajos, sustituir onTrabajosFormSubmit() por onTrabajosFormSubmitSafe().');

  var txt = lines.join('\n');
  Logger.log(txt);
  // Mostrar en partes si es muy largo
  if (txt.length <= 1400) {
    ui.alert('🛡️ Estado de Protecciones\n\n' + txt);
  } else {
    ui.alert('🛡️ Estado de Protecciones (1/2)\n\n' + txt.substring(0, 1400));
    ui.alert('🛡️ Estado de Protecciones (2/2)\n\n' + txt.substring(1400));
  }
}
