// ============================================================
// SincClientes.gs — FASE E: Sincronización automática de
// clientes creados manualmente en la hoja CLIENTES.
//
// PROBLEMA: si Álvaro añade un cliente directamente en la
// hoja (sin usar el formulario), los dropdowns de PANEL_FACTURAS,
// PANEL_CLIENTE y el formulario de TRABAJOS no se actualizan.
//
// SOLUCIÓN: sincronizarClientesSistema() + trigger onEdit.
// ============================================================

// ---- Función principal de sincronización ----

/**
 * Sincroniza todos los desplegables y validaciones con la lista
 * actual de clientes en la hoja CLIENTES.
 * Llamar: desde menú, desde onEdit, o tras alta manual de cliente.
 */
function sincronizarClientesSistema() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss      = getTargetSS_();
    var nombres = _sinc_obtenerListaClientes_(ss);

    if (nombres.length === 0) {
      ui.alert('⚠️ No hay clientes activos en CLIENTES. No se actualizarán los desplegables.');
      return { success: false, error: 'Sin clientes activos' };
    }

    var nActualizados = _sinc_actualizarDropdownsSheets_(ss, nombres);
    _sinc_actualizarFormularioTrabajos_();

    logEvent('INFO', 'sincronizarClientesSistema',
      nombres.length + ' clientes — ' + nActualizados + ' paneles actualizados', '');

    ui.alert(
      '✅ Sincronización completada',
      'Clientes encontrados: ' + nombres.length + '\n'
      + 'Paneles/dropdowns actualizados: ' + nActualizados + '\n\n'
      + 'Los desplegables en PANEL_FACTURAS, PANEL_CLIENTE,\n'
      + 'PANEL_CONTABLE y el formulario de Trabajos ahora\n'
      + 'reflejan la lista actual de clientes.',
      ui.ButtonSet.OK
    );

    return { success: true, clientes: nombres.length, paneles: nActualizados };

  } catch (e) {
    logEvent('ERROR', 'sincronizarClientesSistema', e.message, e.stack);
    ui.alert('❌ Error en sincronización:\n' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Sincronización silenciosa — misma lógica sin alerts.
 * Llamada desde el trigger onEdit y otras funciones automáticas.
 */
function sincronizarClientesSilencioso() {
  try {
    var ss      = getTargetSS_();
    var nombres = _sinc_obtenerListaClientes_(ss);
    if (nombres.length === 0) return;

    _sinc_actualizarDropdownsSheets_(ss, nombres);
    _sinc_actualizarFormularioTrabajos_();

    logEvent('INFO', 'sincronizarClientesSilencioso',
      nombres.length + ' clientes sincronizados', '');
  } catch (e) {
    logEvent('WARN', 'sincronizarClientesSilencioso', e.message, '');
  }
}

// ---- Lectura de clientes ----

/**
 * Devuelve array de etiquetas de cliente en el formato:
 * "Nombre Apellidos" o "Nombre Apellidos (Empresa)"
 * Solo clientes con Activo = true.
 */
function _sinc_obtenerListaClientes_(ss) {
  var sh = ss.getSheetByName('CLIENTES');
  if (!sh || sh.getLastRow() < 2) return [];

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues();
  var nombres = [];

  data.forEach(function(row) {
    var nombre   = String(row[COL_CLI.NOMBRE    - 1] || '').trim();
    var ape      = String(row[COL_CLI.APELLIDOS - 1] || '').trim();
    var empresa  = String(row[COL_CLI.EMPRESA   - 1] || '').trim();
    var activo   = row[COL_CLI.ACTIVO - 1];
    var estaActivo = (activo === true || String(activo).toLowerCase() === 'true');

    if (!estaActivo) return;

    var label = (nombre + ' ' + ape).trim();
    if (!label && empresa) label = empresa;
    if (!label) return;

    if (empresa && empresa !== label) label += ' (' + empresa + ')';

    if (nombres.indexOf(label) === -1) nombres.push(label);
  });

  return nombres.sort();
}

// ---- Actualización de dropdowns en hojas ----

function _sinc_actualizarDropdownsSheets_(ss, nombres) {
  var nActualizados = 0;
  if (nombres.length === 0) return nActualizados;

  var validacion = SpreadsheetApp.newDataValidation()
    .requireValueInList(nombres, true)
    .setAllowInvalid(true) // permite valores que no están en la lista (no rompe datos existentes)
    .build();

  // --- PANEL_FACTURAS ---
  try {
    var pfSh = ss.getSheetByName('PANEL_FACTURAS');
    if (pfSh) {
      if (typeof P_ !== 'undefined' && P_.CLIENTE) {
        pfSh.getRange(P_.CLIENTE[0], P_.CLIENTE[1]).setDataValidation(validacion);
      }
      if (typeof P_ !== 'undefined' && P_.PRE_CLIENTE) {
        pfSh.getRange(P_.PRE_CLIENTE[0], P_.PRE_CLIENTE[1]).setDataValidation(validacion);
      }
      nActualizados++;
    }
  } catch (_) {}

  // --- PANEL_CLIENTE ---
  try {
    var pcSh = ss.getSheetByName('PANEL_CLIENTE');
    if (pcSh) {
      if (typeof PC_ !== 'undefined' && PC_.CLIENTE) {
        pcSh.getRange(PC_.CLIENTE[0], PC_.CLIENTE[1]).setDataValidation(validacion);
      }
      nActualizados++;
    }
  } catch (_) {}

  // --- PANEL_CONTABLE --- (busca celda de cliente en posición estándar)
  try {
    var pconSh = ss.getSheetByName('PANEL_CONTABLE');
    if (pconSh) {
      // El selector de cliente suele estar en las primeras filas; buscar por contenido
      _sinc_actualizarValidacionEnHoja_(pconSh, validacion, 'Cliente', 20);
      nActualizados++;
    }
  } catch (_) {}

  // --- PANEL_GENERAL ---
  try {
    var pgenSh = ss.getSheetByName('PANEL_GENERAL');
    if (pgenSh) {
      _sinc_actualizarValidacionEnHoja_(pgenSh, validacion, 'Cliente', 30);
      nActualizados++;
    }
  } catch (_) {}

  // --- PANEL_PENDIENTES ---
  try {
    var ppSh = ss.getSheetByName('PANEL_PENDIENTES');
    if (ppSh) {
      _sinc_actualizarValidacionEnHoja_(ppSh, validacion, 'Cliente', 20);
      nActualizados++;
    }
  } catch (_) {}

  return nActualizados;
}

/**
 * Busca en las primeras `maxFilas` filas de una hoja una celda que contenga
 * la palabra `etiqueta` y actualiza la validación de la celda adyacente.
 */
function _sinc_actualizarValidacionEnHoja_(sh, validacion, etiqueta, maxFilas) {
  var data = sh.getRange(1, 1, Math.min(maxFilas, sh.getLastRow()), Math.min(sh.getLastColumn(), 5)).getValues();
  for (var r = 0; r < data.length; r++) {
    for (var c = 0; c < data[r].length; c++) {
      var celda = String(data[r][c] || '').toLowerCase();
      if (celda.indexOf(etiqueta.toLowerCase()) !== -1) {
        // Actualizar la celda a la derecha (la de valor)
        try {
          sh.getRange(r + 1, c + 2).setDataValidation(validacion);
        } catch (_) {}
        return;
      }
    }
  }
}

// ---- Actualización del formulario de Trabajos ----

function _sinc_actualizarFormularioTrabajos_() {
  try {
    refreshTrabajosDropdowns();
  } catch (e) {
    logEvent('WARN', '_sinc_actualizarFormularioTrabajos_',
      'No se pudo actualizar el formulario de Trabajos: ' + e.message, '');
  }
}

// ============================================================
// TRIGGER onEdit para CLIENTES
// ============================================================

/**
 * Handler del trigger onEdit para la hoja CLIENTES.
 * Sólo se dispara cuando se modifica un campo de datos de cliente,
 * no por formato, colores o celdas vacías.
 *
 * Se instala como trigger de tipo ON_EDIT (ver instalarTriggerSincClientes).
 */
function onEditClientesHandler(e) {
  try {
    // Verificar que el evento tiene el contexto esperado
    if (!e || !e.range) return;

    var sh = e.range.getSheet();
    if (!sh || sh.getName() !== 'CLIENTES') return;

    var filaEditada = e.range.getRow();
    var colEditada  = e.range.getColumn();

    // Ignorar la fila de cabecera
    if (filaEditada <= 1) return;

    // Columnas de datos que afectan a los desplegables:
    // NOMBRE(3), APELLIDOS(4), EMPRESA(5), EMAIL(6), ACTIVO(10)
    var colsRelevantes = [
      COL_CLI.NOMBRE, COL_CLI.APELLIDOS, COL_CLI.EMPRESA,
      COL_CLI.EMAIL,  COL_CLI.ACTIVO
    ];

    var esColumnaRelevante = colsRelevantes.some(function(col) {
      return colEditada >= col && colEditada <= col;
    });

    if (!esColumnaRelevante) return;

    // Verificar que el nuevo valor no es vacío (evitar sincronizar por borrado de celda)
    var nuevoValor = String(e.value || '').trim();
    if (!nuevoValor) return;

    // Sincronizar con un pequeño delay para que Sheets termine de escribir
    Utilities.sleep(500);
    sincronizarClientesSilencioso();

  } catch (err) {
    logEvent('WARN', 'onEditClientesHandler', err.message, '');
  }
}

// ---- Instalación y gestión del trigger ----

function instalarTriggerSincClientes() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = getTargetSS_();

    // Eliminar trigger previo si existe para evitar duplicados
    desinstalarTriggerSincClientes(true);

    ScriptApp.newTrigger('onEditClientesHandler')
      .forSpreadsheet(ss)
      .onEdit()
      .create();

    logEvent('INFO', 'instalarTriggerSincClientes', 'Trigger onEdit para CLIENTES instalado', '');
    ui.alert('✅ Trigger instalado.\n\n'
      + 'Ahora, cuando añadas o modifiques un cliente directamente\n'
      + 'en la hoja CLIENTES, los desplegables se actualizarán\n'
      + 'automáticamente en todos los paneles y formularios.');

  } catch (e) {
    logEvent('ERROR', 'instalarTriggerSincClientes', e.message, '');
    ui.alert('❌ Error al instalar trigger:\n' + e.message);
  }
}

function desinstalarTriggerSincClientes(silencioso) {
  var triggers = ScriptApp.getProjectTriggers();
  var nEliminados = 0;
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'onEditClientesHandler') {
      ScriptApp.deleteTrigger(t);
      nEliminados++;
    }
  });
  if (!silencioso) {
    SpreadsheetApp.getUi().alert(
      nEliminados > 0
        ? '✅ Trigger onEdit de CLIENTES desinstalado.'
        : 'ℹ️ No había trigger de CLIENTES instalado.'
    );
  }
  return nEliminados;
}

function verEstadoTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var resumen = triggers.map(function(t) {
    return t.getHandlerFunction() + ' (' + t.getEventType() + ')';
  });
  SpreadsheetApp.getUi().alert(
    '⚙️ Triggers activos (' + triggers.length + '):\n\n'
    + (resumen.length ? resumen.join('\n') : 'Ninguno')
  );
}

// ---- Sincronización del sistema completo ----

/**
 * Sincroniza clientes Y refresca todos los paneles activos.
 */
function sincronizarSistemaCompleto() {
  var ui = SpreadsheetApp.getUi();
  try {
    // 1. Sync clientes
    var ss     = getTargetSS_();
    var nombres = _sinc_obtenerListaClientes_(ss);
    _sinc_actualizarDropdownsSheets_(ss, nombres);
    _sinc_actualizarFormularioTrabajos_();

    // 2. Refrescar paneles si sus funciones existen
    var actualizados = [];

    try { actualizarPanelGeneral();    actualizados.push('PANEL_GENERAL');    } catch (_) {}
    try { refreshPanelCliente();       actualizados.push('PANEL_CLIENTE');     } catch (_) {}
    try { actualizarPanelContable();   actualizados.push('PANEL_CONTABLE');    } catch (_) {}
    try { refreshPanelPendientes();    actualizados.push('PANEL_PENDIENTES');  } catch (_) {}
    try { _actualizarPanelGastos_();   actualizados.push('PANEL_GASTOS');     } catch (_) {}

    logEvent('INFO', 'sincronizarSistemaCompleto',
      'Clientes: ' + nombres.length + ' | Paneles: ' + actualizados.join(', '), '');

    ui.alert(
      '✅ Sistema sincronizado',
      'Clientes: ' + nombres.length + '\n'
      + 'Paneles actualizados:\n' + actualizados.map(function(p){ return '  • ' + p; }).join('\n'),
      ui.ButtonSet.OK
    );
  } catch (e) {
    logEvent('ERROR', 'sincronizarSistemaCompleto', e.message, e.stack);
    ui.alert('❌ Error:\n' + e.message);
  }
}
