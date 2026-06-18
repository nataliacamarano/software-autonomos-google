// ============================================================
// MejorasImplementadas.gs — FASE H: Informe de mejoras
// implementadas en el sistema de gestión.
// ============================================================

/**
 * Genera la hoja MEJORAS_IMPLEMENTADAS con un resumen completo
 * de todos los archivos, funciones, triggers y mejoras añadidas
 * durante las fases de auditoría y limpieza.
 */
function generarInformeMejoras() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = getTargetSS_();
    var sh = ss.getSheetByName('MEJORAS_IMPLEMENTADAS');
    if (sh) { sh.clear(); sh.clearFormats(); }
    else    { sh = ss.insertSheet('MEJORAS_IMPLEMENTADAS'); }

    sh.setColumnWidth(1, 200);
    sh.setColumnWidth(2, 300);
    sh.setColumnWidth(3, 250);
    sh.setColumnWidth(4, 150);
    sh.setColumnWidth(5, 200);

    var r = 1;
    var tz = Session.getScriptTimeZone();
    var ahora = new Date();

    // ---- Título ----
    sh.getRange(r, 1, 1, 5).merge()
      .setValue('📋 MEJORAS IMPLEMENTADAS — Sistema de Gestión Álvaro')
      .setBackground('#1a73e8').setFontColor('#ffffff')
      .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
    sh.setRowHeight(r, 45); r++;

    sh.getRange(r, 1, 1, 5).merge()
      .setValue('Generado el: ' + Utilities.formatDate(ahora, tz, 'dd/MM/yyyy HH:mm:ss'))
      .setBackground('#e8f0fe').setFontColor('#1a73e8').setHorizontalAlignment('center');
    r += 2;

    // ---- SECCIÓN 1: Archivos creados/modificados ----
    r = _mej_seccion_(sh, r, '📁 ARCHIVOS CREADOS / MODIFICADOS');

    var archivos = [
      ['Auditoria.gs',            'CREADO',    'Auditoría completa de datos reales vs. prueba. Genera AUDITORIA_DATOS.',                 'ejecutarAuditoriaDatos()'],
      ['AuditoriaFase2.gs',       'CREADO',    'Informe ejecutivo profundo: semáforo 🟢/🟡/🔴, bugs detectados, plan de limpieza.',      'ejecutarAuditoriaFase2()'],
      ['Backup.gs',               'CREADO',    'Backup completo de 9 hojas con prefijo BKP_{fecha}. Aborta limpieza si falla.',           'ejecutarBackupCompleto()'],
      ['Limpieza.gs',             'CREADO',    'Fases B/C/D: eliminar pruebas, corregir duplicados, integridad referencial. Log completo.', 'ejecutarFase3Completa()'],
      ['SincClientes.gs',         'CREADO',    'Sincronización automática de clientes en todos los dropdowns y formularios.',             'sincronizarClientesSistema()'],
      ['Protecciones.gs',         'CREADO',    'Anti-duplicado: LockService, idempotencia, hash deduplicación, nextIdSafe().',            'verificarEstadoProtecciones()'],
      ['MejorasImplementadas.gs', 'CREADO',    'Este informe. Resumen de todas las mejoras implementadas.',                              'generarInformeMejoras()'],
      ['Panel.gs',                'MODIFICADO','Menú ampliado con 20 nuevas opciones de administración (FASE G).',                       'onOpen()'],
      ['Utils.gs',                'MODIFICADO','nextId() marcada como no thread-safe. Usar nextIdSafe() en producción.',                  'nextId()']
    ];

    r = _mej_tabla_(sh, r, ['Archivo', 'Estado', 'Descripción', 'Función principal'], archivos, ['#e8f5e9','#f0fff4']);
    r++;

    // ---- SECCIÓN 2: Funciones clave añadidas ----
    r = _mej_seccion_(sh, r, '⚙️ FUNCIONES CLAVE AÑADIDAS');

    var funciones = [
      ['nextIdSafe()',                    'Protecciones.gs', 'Versión thread-safe de nextId() con LockService (15s timeout).'],
      ['createClienteSafe()',             'Protecciones.gs', 'Alta de cliente con LockService + validación email duplicado.'],
      ['onTrabajosFormSubmitSafe()',      'Protecciones.gs', 'Wrapper idempotente para formulario de trabajos (anti-retry).'],
      ['_prot_generarHashEvento_()',      'Protecciones.gs', 'Hash de deduplicación de envíos de formulario (Cliente+Operario+Fecha).'],
      ['limpiarEventosAntiguos()',        'Protecciones.gs', 'Limpia tokens de idempotencia con más de 24h en ScriptProperties.'],
      ['sincronizarClientesSistema()',    'SincClientes.gs', 'Actualiza dropdowns en PANEL_FACTURAS, PANEL_CLIENTE, PANEL_CONTABLE, PANEL_GENERAL, PANEL_PENDIENTES + formulario Trabajos.'],
      ['sincronizarClientesSilencioso()', 'SincClientes.gs', 'Versión sin UI para llamadas automáticas (triggers, createClienteSafe).'],
      ['onEditClientesHandler()',         'SincClientes.gs', 'Trigger onEdit: sincroniza al modificar nombre/apellidos/empresa/email/activo en CLIENTES.'],
      ['sincronizarSistemaCompleto()',    'SincClientes.gs', 'Sync clientes + refresca todos los paneles del sistema.'],
      ['ejecutarBackupCompleto()',        'Backup.gs',       'Backup de 9 hojas con prefijo BKP_YYYYMMDD_HHMM. Genera BACKUP_REALIZADO.'],
      ['ejecutarFase3Completa()',         'Limpieza.gs',     'Orquestador: Backup → Fase B → Fase C → Fase D (con confirmación).'],
      ['ejecutarFaseB_LimpiezaPrueba()', 'Limpieza.gs',     'Elimina datos de prueba con verificación de actividad y facturas asociadas.'],
      ['ejecutarFaseC_CorreccionDuplicados()', 'Limpieza.gs','Corrige duplicados exactos. Envía casos complejos a REVISIÓN_MANUAL.'],
      ['ejecutarFaseD_IntegridadReferencial()', 'Limpieza.gs','Verifica referencias entre hojas. Todo problema → REVISIÓN_MANUAL.'],
      ['ejecutarAuditoriaDatos()',        'Auditoria.gs',    'Fase 1: clasificación REAL/PRUEBA/DUDOSO + duplicados + huérfanos.'],
      ['ejecutarAuditoriaFase2()',        'AuditoriaFase2.gs','Fase 2: comparación campo a campo, análisis timestamps, bugs, semáforo.'],
      ['generarInformeMejoras()',         'MejorasImplementadas.gs','Este informe.']
    ];

    r = _mej_tabla_(sh, r, ['Función', 'Archivo', 'Descripción'], funciones, ['#e8f0fe','#f0f4ff']);
    r++;

    // ---- SECCIÓN 3: Triggers ----
    r = _mej_seccion_(sh, r, '⏰ TRIGGERS DISPONIBLES');

    var triggers = [
      ['onEditClientesHandler', 'ON_EDIT', 'CLIENTES',  'Sincroniza dropdowns al editar nombre/apellidos/empresa/email/activo.', 'instalarTriggerSincClientes()'],
      ['limpiarEventosAntiguos','TIME_BASED','Diario',   'Limpia tokens de idempotencia caducados de ScriptProperties.',          'Instalar manualmente en Apps Script'],
      ['onFormSubmitHandler',   'ON_FORM_SUBMIT','Trabajos','Procesa respuestas del formulario de trabajos.',                    'Ya existía — verificar que usa Safe()']
    ];

    r = _mej_tabla_(sh, r, ['Handler', 'Tipo', 'Fuente', 'Descripción', 'Instalación'], triggers, ['#fff3cd','#fffde7']);
    r++;

    // ---- SECCIÓN 4: Hojas generadas por el sistema ----
    r = _mej_seccion_(sh, r, '📊 HOJAS GENERADAS POR EL SISTEMA');

    var hojas = [
      ['AUDITORIA_DATOS',        'Auditoría Fase 1',  'Clasificación REAL/PRUEBA/DUDOSO, duplicados, huérfanos, recomendaciones.'],
      ['AUDITORIA_FASE2',        'Auditoría Fase 2',  'Semáforo 🟢/🟡/🔴 por registro, bugs detectados, plan de 9 pasos.'],
      ['BKP_YYYYMMDD_HHMM_*',   'Backup',            '9 hojas de backup (CLI, TRA, FAC, PRO, PRE, GAS, INF, REC, ALT).'],
      ['BACKUP_REALIZADO',       'Backup',            'Informe del último backup: hojas copiadas, registros, instrucciones restauración.'],
      ['LIMPIEZA_REALIZADA',     'Limpieza',          'Log completo de toda eliminación: qué, cuándo, fila, datos originales.'],
      ['REVISIÓN_MANUAL',        'Limpieza',          'Casos que requieren revisión humana: duplicados complejos, huérfanos, inconsistencias.'],
      ['MEJORAS_IMPLEMENTADAS',  'Este archivo',      'Resumen de todas las mejoras implementadas en el sistema.']
    ];

    r = _mej_tabla_(sh, r, ['Hoja', 'Origen', 'Descripción'], hojas, ['#e8f5e9','#f0fff4']);
    r++;

    // ---- SECCIÓN 5: Protecciones anti-duplicado ----
    r = _mej_seccion_(sh, r, '🛡️ PROTECCIONES ANTI-DUPLICADO IMPLEMENTADAS');

    var prot = [
      ['LockService (getScriptLock)',    'nextIdSafe(), createClienteSafe()',        'Previene condición de carrera en generación de IDs y alta de clientes.'],
      ['Idempotencia con ScriptProperties', 'onTrabajosFormSubmitSafe()',            'Hash único por envío. TTL 24h. Previene retry paralelo de Apps Script.'],
      ['Validación de email único',      'createClienteSafe()',                      'Rechaza alta de cliente si el email ya existe en la hoja CLIENTES.'],
      ['Doble verificación de ID',       'createClienteSafe()',                      '_prot_idExiste_() verifica que el ID generado no exista antes de appendRow().'],
      ['Fallback timestamp en ID',       'nextIdSafe()',                             'Si LockService agota timeout (15s), genera ID con últimos 6 dígitos del timestamp.'],
      ['setAllowInvalid(true) en dropdowns', 'SincClientes.gs',                     'Permite valores existentes que no estén en la lista, sin romper datos previos.'],
      ['Backup obligatorio antes de limpieza', 'ejecutarFase3Completa()',            'Aborta toda la limpieza si el backup falla. No elimina datos sin respaldo.']
    ];

    r = _mej_tabla_(sh, r, ['Protección', 'Dónde', 'Qué previene'], prot, ['#fce8e6','#fff0ed']);
    r++;

    // ---- SECCIÓN 6: Reglas de limpieza aplicadas ----
    r = _mej_seccion_(sh, r, '📏 REGLAS DE LIMPIEZA (INAMOVIBLES)');

    var reglas = [
      ['🔴 NO ELIMINAR registros marcados como NO TOCAR',          'Limpieza.gs'],
      ['🔴 NO ELIMINAR registros con factura asociada (ID_Factura con FAC-)', 'Limpieza.gs'],
      ['🔴 NO ELIMINAR registros con diferencias entre duplicados', 'Limpieza.gs'],
      ['🔴 NO ELIMINAR registros dudosos — van a REVISIÓN_MANUAL',  'Limpieza.gs'],
      ['✅ TODO cambio queda registrado en LIMPIEZA_REALIZADA',     'Limpieza.gs'],
      ['✅ TODO puede revertirse usando el backup',                  'Backup.gs'],
      ['✅ NO CONTINUAR si el backup falla',                        'Limpieza.gs — ejecutarFase3Completa()'],
      ['✅ Eliminación siempre de abajo hacia arriba (evitar desplazamiento de índices)', 'Limpieza.gs']
    ];

    var bgReglas = ['#fce8e6','#fff0ed'];
    sh.getRange(r, 1, 1, 4).setValues([['Regla', '', '', 'Archivo']]).setFontWeight('bold').setBackground('#202124').setFontColor('#ffffff');
    r++;
    reglas.forEach(function(reg, i) {
      var bg = bgReglas[i % 2];
      sh.getRange(r, 1, 1, 3).merge().setValue(reg[0]).setBackground(bg).setWrap(true);
      sh.getRange(r, 4).setValue(reg[1]).setBackground(bg);
      r++;
    });
    r++;

    // ---- SECCIÓN 7: Menú FASE G ----
    r = _mej_seccion_(sh, r, '🍽️ MENÚ ÁLVARO GESTIÓN — OPCIONES AÑADIDAS (FASE G)');

    var menu = [
      ['💾 BACKUP — Crear copia de seguridad completa',   'ejecutarBackupCompleto'],
      ['🧹 LIMPIEZA — Fase 3 completa (A→D)',             'ejecutarFase3Completa'],
      ['🧹 LIMPIEZA B — Solo prueba',                     'ejecutarFaseB_solo'],
      ['🔧 LIMPIEZA C — Solo duplicados exactos',         'ejecutarFaseC_solo'],
      ['🔗 LIMPIEZA D — Solo integridad referencial',     'ejecutarFaseD_solo'],
      ['🔄 CLIENTES — Sincronizar desplegables',          'sincronizarClientesSistema'],
      ['🔄 SISTEMA — Sincronizar todo + paneles',         'sincronizarSistemaCompleto'],
      ['🔌 CLIENTES — Instalar trigger sincronización',   'instalarTriggerSincClientes'],
      ['🔌 CLIENTES — Desinstalar trigger sincronización','desinstalarTriggerSincClientes'],
      ['🛡️ PROTECCIONES — Ver estado',                    'verificarEstadoProtecciones'],
      ['🗑️ EVENTOS — Limpiar idempotencia antigua',       'limpiarEventosAntiguos'],
      ['📋 INFORME — Mejoras implementadas (FASE H)',      'generarInformeMejoras']
    ];

    sh.getRange(r, 1, 1, 2).setValues([['Opción de menú', 'Función']]).setFontWeight('bold').setBackground('#202124').setFontColor('#ffffff');
    r++;
    menu.forEach(function(item, i) {
      var bg = i % 2 === 0 ? '#e8f0fe' : '#f0f4ff';
      sh.getRange(r, 1).setValue(item[0]).setBackground(bg);
      sh.getRange(r, 2).setValue(item[1]).setFontFamily('Courier New').setBackground(bg);
      r++;
    });
    r++;

    // ---- Pie ----
    sh.getRange(r, 1, 1, 5).merge()
      .setValue('⚠️ Para aplicar cambios: Menú → 🧹 LIMPIEZA — Fase 3 completa. Siempre se crea backup previo automáticamente.')
      .setBackground('#fff3cd').setFontColor('#856404').setFontWeight('bold').setWrap(true);

    SpreadsheetApp.flush();
    ss.setActiveSheet(sh);

    logEvent('INFO', 'generarInformeMejoras', 'Hoja MEJORAS_IMPLEMENTADAS generada', '');
    ui.alert('✅ Informe generado en hoja: MEJORAS_IMPLEMENTADAS');

  } catch (e) {
    logEvent('ERROR', 'generarInformeMejoras', e.message, e.stack);
    ui.alert('❌ Error generando informe:\n' + e.message);
  }
}

// ---- Helpers de formato ----

function _mej_seccion_(sh, r, titulo) {
  sh.getRange(r, 1, 1, 5).merge()
    .setValue(titulo)
    .setBackground('#202124').setFontColor('#ffffff')
    .setFontSize(11).setFontWeight('bold');
  sh.setRowHeight(r, 32);
  return r + 1;
}

function _mej_tabla_(sh, r, cabeceras, filas, bgs) {
  sh.getRange(r, 1, 1, cabeceras.length).setValues([cabeceras]).setFontWeight('bold').setBackground('#5f6368').setFontColor('#ffffff');
  r++;
  filas.forEach(function(fila, i) {
    var bg = bgs[i % 2];
    for (var c = 0; c < fila.length; c++) {
      sh.getRange(r, c + 1).setValue(fila[c]).setBackground(bg).setWrap(true);
    }
    r++;
  });
  return r;
}

// ---- Wrappers para menú (Fases individuales) ----

function ejecutarFaseB_solo() {
  ejecutarFaseB_LimpiezaPrueba(false);
}

function ejecutarFaseC_solo() {
  ejecutarFaseC_CorreccionDuplicados(false);
}

function ejecutarFaseD_solo() {
  ejecutarFaseD_IntegridadReferencial(false);
}
