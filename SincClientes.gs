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

// ---- Checkboxes en columna Activo ----

/**
 * Inserta checkboxes visuales (TRUE/FALSE) en CLIENTES!J para todas las filas.
 * Usa insertCheckboxes() — el método más directo para mostrar el tick visual.
 * Marca TRUE: Gregor, Pougeoux, Fouss. El resto queda como estaba (normalizado).
 */
function repararCheckboxesActivo() {
  var ui = SpreadsheetApp.getUi();
  var ss = getTargetSS_();

  try {
    var sh = ss.getSheetByName('CLIENTES');
    if (!sh) { ui.alert('❌ Hoja CLIENTES no encontrada.'); return; }

    var lastRow = sh.getLastRow();
    if (lastRow < 2) { ui.alert('⚠️ CLIENTES sin filas de datos.'); return; }
    var nRows = lastRow - 1;

    var REALES = ['gregor', 'pougeoux', 'fouss'];
    function norm(t) {
      return String(t || '').toLowerCase()
        .replace(/[áàä]/g, 'a').replace(/[éèë]/g, 'e')
        .replace(/[íìï]/g, 'i').replace(/[óòö]/g, 'o')
        .replace(/[úùü]/g, 'u').trim();
    }

    var data = sh.getRange(2, 1, nRows, COL_CLI.METODOS_PAGO).getValues();

    // Calcular qué valor debe tener cada fila
    var nuevosValores = data.map(function(row) {
      var texto = norm(row[COL_CLI.NOMBRE - 1]) + ' '
                + norm(row[COL_CLI.APELLIDOS - 1]) + ' '
                + norm(row[COL_CLI.EMPRESA - 1]);
      if (REALES.some(function(k) { return texto.indexOf(k) !== -1; })) return [true];
      var actual = row[COL_CLI.ACTIVO - 1];
      return [(actual === true || String(actual).toLowerCase() === 'true')];
    });

    var rangoJ = sh.getRange(2, COL_CLI.ACTIVO, nRows, 1);

    // 1. Limpiar todo: validaciones y formato anterior
    rangoJ.clearDataValidations();
    rangoJ.clearFormat();

    // 2. Insertar checkboxes visuales (TRUE = marcado, FALSE = desmarcado)
    rangoJ.insertCheckboxes();

    // 3. Escribir los valores booleanos
    rangoJ.setValues(nuevosValores);

    var nActivos = nuevosValores.filter(function(v) { return v[0] === true; }).length;

    logEvent('INFO', 'repararCheckboxesActivo',
      nRows + ' filas | ' + nActivos + ' marcados TRUE', '');

    ui.alert(
      '✅ Checkboxes instalados en CLIENTES!J',
      'Filas procesadas: ' + nRows + '\n'
      + 'Marcados como activos (✓): ' + nActivos + '\n\n'
      + 'La columna Activo ahora muestra checkboxes.\n'
      + 'Gregor, Pougeoux y Fouss están marcados como activos.',
      ui.ButtonSet.OK
    );

  } catch (e) {
    logEvent('ERROR', 'repararCheckboxesActivo', e.message, e.stack);
    ui.alert('❌ Error:\n' + e.message);
  }
}

// ---- Verificación rápida de C5 ----

/**
 * Muestra en un alert qué clientes hay en el desplegable de C5 de PANEL_FACTURAS
 * y si Gregor, Pougeoux y Fouss están presentes.
 */
function verificarC5PanelFacturas() {
  var ui = SpreadsheetApp.getUi();
  var ss = getTargetSS_();

  var pf = ss.getSheetByName('PANEL_FACTURAS');
  if (!pf) { ui.alert('❌ Hoja PANEL_FACTURAS no encontrada.'); return; }

  var dv = pf.getRange(5, 3).getDataValidation();
  if (!dv) {
    ui.alert('⚠️ C5 de PANEL_FACTURAS NO tiene validación de datos.\n'
      + 'Ejecuta primero: ✅ CLIENTES — Reparar columna Activo + desplegables');
    return;
  }

  var vals = dv.getCriteriaValues();
  var lista = (vals && Array.isArray(vals[0])) ? vals[0] : [];

  function norm(t) {
    return String(t || '').toLowerCase()
      .replace(/[áàä]/g, 'a').replace(/[éèë]/g, 'e')
      .replace(/[íìï]/g, 'i').replace(/[óòö]/g, 'o')
      .replace(/[úùü]/g, 'u').trim();
  }

  var checks = [
    { frag: 'gregor',   label: 'Gregor' },
    { frag: 'pougeoux', label: 'Patrick Pougeoux' },
    { frag: 'fouss',    label: 'Jose Fouss' }
  ].map(function(c) {
    var encontrado = lista.find(function(n) { return norm(n).indexOf(c.frag) !== -1; });
    return (encontrado ? '✅ ' : '❌ ') + c.label + (encontrado ? ' → "' + encontrado + '"' : ' → NO ENCONTRADO');
  });

  ui.alert(
    '🔍 Verificación C5 — PANEL_FACTURAS',
    'Clientes en el desplegable (' + lista.length + '):\n'
    + lista.map(function(n) { return '  • ' + n; }).join('\n')
    + '\n\n'
    + 'Verificación clientes reales:\n'
    + checks.map(function(c) { return '  ' + c; }).join('\n'),
    ui.ButtonSet.OK
  );
}

// ---- Diagnóstico completo de CLIENTES ----

/**
 * Muestra en un alert el estado real de la hoja CLIENTES:
 * - Todas las filas con su nombre, apellidos y valor de Activo
 * - Qué clientes pasan el filtro de activos
 * - Qué hay en la validación de C5 de PANEL_FACTURAS
 */
function diagnosticarClientes() {
  var ui = SpreadsheetApp.getUi();
  var ss = getTargetSS_();

  var sh = ss.getSheetByName('CLIENTES');
  if (!sh || sh.getLastRow() < 2) {
    ui.alert('CLIENTES vacía o no encontrada.');
    return;
  }

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, COL_CLI.METODOS_PAGO).getValues();

  var lineas = ['=== CLIENTES (todas las filas) ==='];
  data.forEach(function(row, i) {
    var id    = String(row[COL_CLI.ID        - 1] || '').trim();
    var nom   = String(row[COL_CLI.NOMBRE    - 1] || '').trim();
    var ape   = String(row[COL_CLI.APELLIDOS - 1] || '').trim();
    var emp   = String(row[COL_CLI.EMPRESA   - 1] || '').trim();
    var act   = row[COL_CLI.ACTIVO - 1];
    var actStr = String(act) + ' (' + typeof act + ')';
    var estaActivo = (act === true || String(act).toLowerCase() === 'true');
    lineas.push('Fila ' + (i + 2) + ': [' + id + '] "' + nom + '" "' + ape + '" emp:"' + emp + '" | Activo=' + actStr + ' → ' + (estaActivo ? '✅ ACTIVO' : '❌ inactivo'));
  });

  var nombresActivos = _sinc_obtenerListaClientes_(ss);
  lineas.push('');
  lineas.push('=== Lista que va al desplegable (' + nombresActivos.length + ') ===');
  nombresActivos.forEach(function(n) { lineas.push('  • "' + n + '"'); });

  var pf = ss.getSheetByName('PANEL_FACTURAS');
  lineas.push('');
  lineas.push('=== Validación actual de C5 en PANEL_FACTURAS ===');
  if (pf) {
    var dv = pf.getRange(5, 3).getDataValidation();
    if (dv) {
      var vals = dv.getCriteriaValues();
      var lista = (vals && Array.isArray(vals[0])) ? vals[0] : [];
      lineas.push('Items en dropdown: ' + lista.length);
      lista.forEach(function(n) { lineas.push('  • "' + n + '"'); });
    } else {
      lineas.push('Sin validación de datos en C5');
    }
    lineas.push('Valor actual de C5: "' + pf.getRange(5, 3).getValue() + '"');
  }

  ui.alert('DIAGNÓSTICO CLIENTES', lineas.join('\n'), ui.ButtonSet.OK);
}

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

  // Usar VALUE_IN_RANGE → CLIENTES_ACTIVOS (fuente maestra) si existe.
  // Fallback a VALUE_IN_LIST solo si la hoja maestra no está disponible.
  var validacion = getValidacionClienteMaestra_(ss);
  if (!validacion) {
    validacion = SpreadsheetApp.newDataValidation()
      .requireValueInList(nombres, true)
      .setAllowInvalid(true)
      .build();
  }

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

// ============================================================
// REPARACIÓN COLUMNA ACTIVO + DESPLEGABLES COMPLETOS
// ============================================================

/**
 * 1. Convierte CLIENTES!J en checkboxes (TRUE/FALSE booleano).
 * 2. Marca TRUE a Gregor, Patrick Pougeoux y Jose Fouss.
 *    El resto: normaliza el valor existente (true→true, resto→false).
 * 3. Reconstruye desplegables en PANEL_FACTURAS (C5, F5, C99) y PANEL_CLIENTE (C4).
 * 4. Verifica que los 3 clientes aparecen en C5 de PANEL_FACTURAS.
 * 5. Crea REPARACION_ACTIVO_CLIENTES con el informe.
 *
 * No modifica ninguna otra columna ni ninguna otra hoja de datos.
 */
function repararColumnaActivoClientes() {
  var ui  = SpreadsheetApp.getUi();
  var ss  = getTargetSS_();
  var log = [];

  try {
    var sh = ss.getSheetByName('CLIENTES');
    if (!sh) { ui.alert('❌ Hoja CLIENTES no encontrada.'); return; }

    var lastRow = sh.getLastRow();
    if (lastRow < 2) { ui.alert('⚠️ CLIENTES no tiene filas de datos.'); return; }
    var nRows = lastRow - 1;

    // ── 1. Leer todas las filas de datos ────────────────────────
    var data = sh.getRange(2, 1, nRows, COL_CLI.METODOS_PAGO).getValues();

    // Fragmentos de nombre que identifican a los 3 clientes reales
    var CLIENTES_REALES = ['gregor', 'pougeoux', 'fouss'];

    function _norm(t) {
      return String(t || '').toLowerCase()
        .replace(/[áàä]/g, 'a').replace(/[éèë]/g, 'e')
        .replace(/[íìï]/g, 'i').replace(/[óòö]/g, 'o')
        .replace(/[úùü]/g, 'u').trim();
    }

    function _esReal(row) {
      var texto = _norm(row[COL_CLI.NOMBRE    - 1]) + ' '
                + _norm(row[COL_CLI.APELLIDOS - 1]) + ' '
                + _norm(row[COL_CLI.EMPRESA   - 1]);
      return CLIENTES_REALES.some(function(k) { return texto.indexOf(k) !== -1; });
    }

    // ── 2. Calcular nuevo valor Activo para cada fila ────────────
    var nuevosActivos = data.map(function(row, i) {
      var nombreCompleto = (String(row[COL_CLI.NOMBRE    - 1] || '').trim() + ' '
                         +  String(row[COL_CLI.APELLIDOS - 1] || '').trim()).trim()
                         || String(row[COL_CLI.EMPRESA   - 1] || '').trim()
                         || '(fila ' + (i + 2) + ')';

      if (_esReal(row)) {
        log.push(['ACTIVO', nombreCompleto, 'Fila ' + (i + 2) + ' → TRUE (cliente real)']);
        return [true];
      }
      // No es cliente real: normalizar valor existente
      var v = row[COL_CLI.ACTIVO - 1];
      var eraTrue = (v === true || String(v).toLowerCase() === 'true');
      if (eraTrue) {
        log.push(['CONSERVADO', nombreCompleto, 'Fila ' + (i + 2) + ' → TRUE (ya estaba activo)']);
        return [true];
      }
      log.push(['INACTIVO', nombreCompleto, 'Fila ' + (i + 2) + ' → FALSE']);
      return [false];
    });

    // ── 3. Aplicar checkbox booleano a toda la columna J ─────────
    var rangoJ = sh.getRange(2, COL_CLI.ACTIVO, nRows, 1);
    rangoJ.setDataValidation(null);   // limpiar cualquier validación anterior
    var dvCheck = SpreadsheetApp.newDataValidation()
      .requireCheckbox()              // TRUE cuando marcado, FALSE cuando no
      .build();
    rangoJ.setDataValidation(dvCheck);

    // ── 4. Escribir los valores booleanos ────────────────────────
    rangoJ.setValues(nuevosActivos);

    var nActivos = nuevosActivos.filter(function(v) { return v[0] === true; }).length;
    log.push(['OK', 'Checkbox aplicado a ' + nRows + ' filas en CLIENTES!J',
      nActivos + ' marcados como TRUE']);

    // ── 5. Leer clientes activos resultantes ─────────────────────
    var nombresActivos = _sinc_obtenerListaClientes_(ss);
    log.push(['INFO', 'Clientes activos tras reparación: ' + nombresActivos.length,
      nombresActivos.join(' | ')]);

    if (nombresActivos.length === 0) {
      log.push(['WARN', 'Sin clientes activos — desplegables no se pueden reconstruir', '']);
      _rac_crearInforme_(ss, log, []);
      ui.alert('⚠️ Se aplicaron los checkboxes pero no se encontraron clientes activos.\n'
        + 'Verifica que los nombres en CLIENTES coinciden con Gregor / Pougeoux / Fouss.');
      return;
    }

    // ── 6. Reconstruir desplegables ───────────────────────────────
    var validacionFresca = SpreadsheetApp.newDataValidation()
      .requireValueInList(nombresActivos, true)
      .setAllowInvalid(true)
      .build();

    // PANEL_FACTURAS: C5, F5, C99
    var pf = ss.getSheetByName('PANEL_FACTURAS');
    if (pf) {
      [[5, 3, 'C5 (Cliente principal)'],
       [5, 6, 'F5 (Cliente presupuesto)'],
       [99, 3, 'C99 (Factura desde Presupuesto)']].forEach(function(c) {
        try {
          var rng = pf.getRange(c[0], c[1]);
          rng.setDataValidation(null);
          rng.setDataValidation(validacionFresca);
          log.push(['OK', 'PANEL_FACTURAS ' + c[2], 'Lista: ' + nombresActivos.join(', ')]);
        } catch (e2) {
          log.push(['ERROR', 'PANEL_FACTURAS ' + c[2], e2.message]);
        }
      });
    } else {
      log.push(['WARN', 'PANEL_FACTURAS no encontrada', '']);
    }

    // PANEL_CLIENTE: C4
    var pc = ss.getSheetByName('PANEL_CLIENTE');
    if (pc) {
      try {
        pc.getRange(4, 3).setDataValidation(null);
        pc.getRange(4, 3).setDataValidation(validacionFresca);
        log.push(['OK', 'PANEL_CLIENTE C4', 'Lista: ' + nombresActivos.join(', ')]);
      } catch (e2) {
        log.push(['ERROR', 'PANEL_CLIENTE C4', e2.message]);
      }
    }

    // ── 7. Verificar C5 de PANEL_FACTURAS ────────────────────────
    if (pf) {
      try {
        var dvC5  = pf.getRange(5, 3).getDataValidation();
        var c5V   = dvC5 ? dvC5.getCriteriaValues() : null;
        var c5List = (c5V && Array.isArray(c5V[0])) ? c5V[0] : [];

        log.push(['VERIFICACIÓN', 'C5 PANEL_FACTURAS tiene ' + c5List.length + ' clientes',
          c5List.join(' | ')]);

        CLIENTES_REALES.forEach(function(frag) {
          var encontrado = c5List.find(function(n) { return _norm(n).indexOf(frag) !== -1; });
          log.push(encontrado
            ? ['✓ PASS', '"' + frag + '" encontrado en C5', '"' + encontrado + '"']
            : ['✗ FAIL', '"' + frag + '" NO encontrado en C5', '']);
        });
      } catch (e2) {
        log.push(['WARN', 'No se pudo verificar C5', e2.message]);
      }
    }

    // ── Informe ───────────────────────────────────────────────────
    _rac_crearInforme_(ss, log, nombresActivos);

    logEvent('INFO', 'repararColumnaActivoClientes',
      nActivos + ' activos | ' + nombresActivos.length + ' en lista | desplegables reconstruidos', '');

    ui.alert(
      '✅ Columna Activo reparada y desplegables reconstruidos',
      'Filas procesadas en CLIENTES: ' + nRows + '\n'
      + 'Marcados como TRUE: ' + nActivos + '\n'
      + 'Clientes en desplegables: ' + nombresActivos.length + '\n\n'
      + 'Desplegables actualizados:\n'
      + '  • PANEL_FACTURAS C5, F5, C99\n'
      + '  • PANEL_CLIENTE C4\n\n'
      + 'Ver REPARACION_ACTIVO_CLIENTES para el detalle.',
      ui.ButtonSet.OK
    );

  } catch (e) {
    logEvent('ERROR', 'repararColumnaActivoClientes', e.message, e.stack);
    ui.alert('❌ Error:\n' + e.message);
  }
}

/**
 * Crea o sobreescribe REPARACION_ACTIVO_CLIENTES con el informe.
 */
function _rac_crearInforme_(ss, log, nombresActivos) {
  try {
    var shName = 'REPARACION_ACTIVO_CLIENTES';
    var sh = ss.getSheetByName(shName);
    if (sh) { sh.clearContents(); sh.clearFormats(); }
    else     { sh = ss.insertSheet(shName); }

    sh.getRange(1, 1, 1, 4).setValues([['Estado', 'Descripción', 'Detalle', 'Timestamp']]);
    sh.getRange(1, 1, 1, 4)
      .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
    sh.setFrozenRows(1);

    sh.appendRow(['RESUMEN', 'Clientes activos resultantes: ' + nombresActivos.length,
      nombresActivos.join(' | '), new Date()]);
    sh.appendRow(['', '', '', '']);

    log.forEach(function(e) {
      sh.appendRow([e[0], e[1], e[2] || '', new Date()]);
    });

    var nData = sh.getLastRow() - 1;
    if (nData > 0) {
      var colorMap = {
        'ACTIVO':       '#e6f4ea', 'CONSERVADO': '#e6f4ea', 'OK': '#e6f4ea',
        '✓ PASS':       '#e6f4ea',
        'INACTIVO':     '#f8f9fa', 'INFO':  '#f1f3f4', 'RESUMEN': '#e8f0fe',
        'VERIFICACIÓN': '#fff3e0',
        'WARN':         '#fef9c3',
        'ERROR':        '#fce8e6', '✗ FAIL': '#fce8e6'
      };
      sh.getRange(2, 1, nData, 1).getValues().forEach(function(row, i) {
        var color = colorMap[String(row[0])] || '#ffffff';
        sh.getRange(i + 2, 1, 1, 4).setBackground(color);
      });
    }

    sh.autoResizeColumns(1, 4);
  } catch (e) {
    logEvent('WARN', '_rac_crearInforme_', e.message, '');
  }
}

// ============================================================
// REPARACIÓN DIRIGIDA: CLIENTE EN PANEL_FACTURAS
// ============================================================

/**
 * Escanea TODAS las celdas de PANEL_FACTURAS con validación,
 * identifica cuáles son selectores de cliente, elimina la validación
 * antigua y aplica la lista fresca de CLIENTES (Activo=TRUE).
 *
 * No toca CLIENTES, PANEL_CLIENTE ni ninguna otra hoja de datos.
 * Genera REPARACION_PANEL_FACTURAS_CLIENTE con el informe completo.
 */
function repararClientePanelFacturas() {
  var ui  = SpreadsheetApp.getUi();
  var ss  = getTargetSS_();
  var log = [];

  try {
    // ── 1. Lista fresca de clientes activos ─────────────────────
    var nombresActivos = _sinc_obtenerListaClientes_(ss);
    log.push(['INFO', 'Clientes activos leídos de CLIENTES',
      nombresActivos.length + ' → ' + nombresActivos.join(' | ')]);

    if (nombresActivos.length === 0) {
      ui.alert('⚠️ No hay clientes activos en la hoja CLIENTES.\n'
        + 'Revisa que la columna Activo sea TRUE para al menos un cliente.');
      return;
    }

    // ── 2. Leer validación de PANEL_CLIENTE C4 (referencia que sí funciona) ──
    var pcSh   = ss.getSheetByName('PANEL_CLIENTE');
    var pcLista = [];
    if (pcSh) {
      try {
        var pcDv = pcSh.getRange(4, 3).getDataValidation();
        if (pcDv) {
          var pcVals = pcDv.getCriteriaValues();
          pcLista = (pcVals && Array.isArray(pcVals[0])) ? pcVals[0] : [];
        }
        log.push(['REF', 'PANEL_CLIENTE C4 tiene ' + pcLista.length + ' valores',
          pcLista.join(' | ')]);
      } catch (e) {
        log.push(['WARN', 'No se pudo leer validación de PANEL_CLIENTE C4', e.message]);
      }
    } else {
      log.push(['WARN', 'Hoja PANEL_CLIENTE no encontrada', '']);
    }

    // ── 3. Escanear PANEL_FACTURAS — TODAS las celdas con validación ──
    var pfSh = ss.getSheetByName('PANEL_FACTURAS');
    if (!pfSh) {
      ui.alert('❌ Hoja PANEL_FACTURAS no encontrada.');
      return;
    }

    var scanRows = Math.min(pfSh.getLastRow(), 135);
    var scanCols = Math.min(pfSh.getLastColumn(), 15);
    log.push(['INFO', 'Escaneando PANEL_FACTURAS filas 1-' + scanRows + ', cols 1-' + scanCols, '']);

    var allDv    = pfSh.getRange(1, 1, scanRows, scanCols).getDataValidations();
    var allVals  = pfSh.getRange(1, 1, scanRows, scanCols).getValues();

    // Coordenadas conocidas que SIEMPRE se actualizan
    var coordsConocidas = [
      [5,  3, 'P_.CLIENTE (selector principal)'],
      [5,  6, 'P_.PRE_CLIENTE (presupuesto)'],
      [99, 3, 'PF_.CLIENTE (Factura desde Presupuesto)']
    ];
    var coordSet = {};
    coordsConocidas.forEach(function(c) { coordSet[c[0] + '_' + c[1]] = true; });

    // Construir set de nombres para detección rápida
    var nombresSet = {};
    nombresActivos.forEach(function(n) { nombresSet[n.toLowerCase()] = true; });

    var celdasCliente = [];   // {row, col, tipo, listaAnterior, etiqueta}
    var celdasOtras   = [];

    for (var r = 0; r < allDv.length; r++) {
      for (var c = 0; c < allDv[r].length; c++) {
        var dv = allDv[r][c];
        if (!dv) continue;

        var row  = r + 1;
        var col  = c + 1;
        var key  = row + '_' + col;
        var tipo = dv.getCriteriaType();
        var vals = dv.getCriteriaValues();

        var listaActual = [];
        var infoLista   = '';
        var esRango     = false;

        if (tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
          listaActual = (vals && Array.isArray(vals[0])) ? vals[0] : [];
          infoLista   = listaActual.length + ' items: ' + listaActual.slice(0, 6).join(', ')
                        + (listaActual.length > 6 ? '…' : '');
        } else if (tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
          esRango   = true;
          infoLista = 'RANGO: ' + (vals && vals[0] ? vals[0].getA1Notation() : '?');
        } else {
          infoLista = 'tipo: ' + tipo;
        }

        // Etiqueta en la celda de la izquierda (si existe)
        var etiquetaIzq = (c > 0) ? String(allVals[r][c - 1] || '').toLowerCase().trim() : '';
        var etiquetaArriba = (r > 0) ? String(allVals[r - 1][c] || '').toLowerCase().trim() : '';

        // ¿Es un selector de cliente?
        var esConocida  = coordSet[key] === true;
        var etiquetaEsCliente = etiquetaIzq.indexOf('cliente') !== -1
                             || etiquetaArriba.indexOf('cliente') !== -1;

        // Overlap con lista de clientes activos (para VALUE_IN_LIST)
        var overlap = 0;
        if (listaActual.length > 0) {
          listaActual.forEach(function(v) {
            if (nombresSet[String(v).toLowerCase()]) overlap++;
          });
        }
        var solapaPorcentaje = listaActual.length > 0 ? overlap / listaActual.length : 0;

        var esCliente = esConocida
                     || etiquetaEsCliente
                     || (listaActual.length >= 1 && solapaPorcentaje >= 0.3);

        var info = {
          row:          row,
          col:          col,
          tipo:         String(tipo),
          listaActual:  listaActual,
          infoLista:    infoLista,
          esRango:      esRango,
          esConocida:   esConocida,
          etiqueta:     etiquetaIzq || etiquetaArriba || '',
          overlap:      overlap,
          solapaPct:    Math.round(solapaPorcentaje * 100),
          esCliente:    esCliente
        };

        if (esCliente) {
          celdasCliente.push(info);
          log.push(['DETECTADO', 'Fila ' + row + ' Col ' + col
            + (esConocida ? ' [COORD CONOCIDA]' : '')
            + (etiquetaEsCliente ? ' [ETIQUETA: "' + info.etiqueta + '"]' : '')
            + (solapaPorcentaje > 0 ? ' [OVERLAP ' + info.solapaPct + '%]' : ''),
            infoLista]);
        } else {
          celdasOtras.push(info);
          log.push(['OTRO', 'Fila ' + row + ' Col ' + col, infoLista]);
        }
      }
    }

    log.push(['INFO', 'Celdas detectadas como cliente: ' + celdasCliente.length
      + ' | Otras validaciones: ' + celdasOtras.length, '']);

    // ── 4. Validación fresca desde CLIENTES ─────────────────────
    var validacionFresca = SpreadsheetApp.newDataValidation()
      .requireValueInList(nombresActivos, true)
      .setAllowInvalid(true)
      .build();

    // ── 5 & 6. Limpiar validación antigua y aplicar nueva ────────
    var nReparadas = 0;

    celdasCliente.forEach(function(info) {
      try {
        var rng = pfSh.getRange(info.row, info.col);
        var valorAntes = rng.getValue();
        rng.setDataValidation(null);          // eliminar completamente
        rng.setDataValidation(validacionFresca); // lista fresca
        log.push(['REPARADO',
          'PANEL_FACTURAS fila ' + info.row + ' col ' + info.col,
          'Antes: "' + valorAntes + '" | Lista: '
          + info.listaActual.length + ' items → ' + nombresActivos.length + ' items']);
        nReparadas++;
      } catch (e2) {
        log.push(['ERROR', 'Fila ' + info.row + ' col ' + info.col, e2.message]);
      }
    });

    // Asegurar que las coordenadas conocidas siempre se actualizan,
    // aunque el escáner no las haya detectado como cliente
    coordsConocidas.forEach(function(coord) {
      var yaReparada = celdasCliente.some(function(i) {
        return i.row === coord[0] && i.col === coord[1];
      });
      if (!yaReparada) {
        try {
          var rng = pfSh.getRange(coord[0], coord[1]);
          var val = rng.getValue();
          rng.setDataValidation(null);
          rng.setDataValidation(validacionFresca);
          log.push(['FORZADO', coord[2] + ' fila ' + coord[0] + ' col ' + coord[1]
            + ' (coord conocida, sin validación previa)', 'Valor: "' + val + '"']);
          nReparadas++;
        } catch (e2) {
          log.push(['ERROR', coord[2] + ' [' + coord[0] + ',' + coord[1] + ']', e2.message]);
        }
      }
    });

    // ── 7 & 8. Verificación en log ────────────────────────────────
    log.push(['VERIFICACIÓN', 'Lista nueva (' + nombresActivos.length + ' clientes activos)',
      nombresActivos.join(' | ')]);

    var clientesDiferentes = nombresActivos.filter(function(n) {
      return pcLista.indexOf(n) === -1;
    });
    if (clientesDiferentes.length > 0) {
      log.push(['INFO', 'Clientes en CLIENTES que no estaban en PANEL_CLIENTE',
        clientesDiferentes.join(' | ')]);
    } else {
      log.push(['OK', 'Lista nueva coincide con la de PANEL_CLIENTE', '']);
    }

    // ── Informe ───────────────────────────────────────────────────
    _rpf_crearInforme_(ss, log, nombresActivos, celdasCliente, celdasOtras, nReparadas);

    logEvent('INFO', 'repararClientePanelFacturas',
      nReparadas + ' celdas reparadas | ' + nombresActivos.length + ' clientes', '');

    ui.alert(
      '🧯 Reparación completada — PANEL_FACTURAS',
      'Celdas reparadas: ' + nReparadas + '\n'
      + 'Clientes en la nueva lista: ' + nombresActivos.length + '\n\n'
      + 'La validación anterior fue eliminada completamente.\n'
      + 'La lista nueva se leyó directamente de la hoja CLIENTES.\n\n'
      + 'Abre el desplegable de cliente en PANEL_FACTURAS\n'
      + 'para verificar que aparece el nuevo cliente.\n\n'
      + 'Ver REPARACION_PANEL_FACTURAS_CLIENTE para detalles.',
      ui.ButtonSet.OK
    );

  } catch (e) {
    logEvent('ERROR', 'repararClientePanelFacturas', e.message, e.stack);
    ui.alert('❌ Error en reparación:\n' + e.message);
  }
}

/**
 * Crea o sobreescribe REPARACION_PANEL_FACTURAS_CLIENTE.
 */
function _rpf_crearInforme_(ss, log, nombresActivos, celdasCliente, celdasOtras, nReparadas) {
  try {
    var shName = 'REPARACION_PANEL_FACTURAS_CLIENTE';
    var sh = ss.getSheetByName(shName);
    if (sh) { sh.clearContents(); sh.clearFormats(); }
    else     { sh = ss.insertSheet(shName); }

    sh.getRange(1, 1, 1, 4).setValues([['Estado', 'Descripción', 'Detalle', 'Timestamp']]);
    sh.getRange(1, 1, 1, 4)
      .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
    sh.setFrozenRows(1);

    sh.appendRow(['RESUMEN', 'Celdas reparadas en PANEL_FACTURAS: ' + nReparadas, '', new Date()]);
    sh.appendRow(['RESUMEN', 'Clientes activos en la nueva lista: ' + nombresActivos.length,
      nombresActivos.join(' | '), new Date()]);
    sh.appendRow(['RESUMEN', 'Celdas detectadas como cliente: ' + celdasCliente.length
      + ' | Otras validaciones encontradas: ' + celdasOtras.length, '', new Date()]);
    sh.appendRow(['', '', '', '']);

    log.forEach(function(entry) {
      sh.appendRow([entry[0], entry[1], entry[2] || '', new Date()]);
    });

    var nData = sh.getLastRow() - 1;
    if (nData > 0) {
      var estados = sh.getRange(2, 1, nData, 1).getValues();
      var colorMap = {
        REPARADO: '#e6f4ea', FORZADO: '#e6f4ea', OK: '#e6f4ea',
        ERROR: '#fce8e6', WARN: '#fef9c3',
        DETECTADO: '#e8f0fe', REF: '#e8f0fe',
        OTRO: '#f8f9fa', INFO: '#f1f3f4',
        VERIFICACIÓN: '#fff3e0', RESUMEN: '#e8f0fe'
      };
      estados.forEach(function(row, i) {
        var color = colorMap[String(row[0])] || '#ffffff';
        sh.getRange(i + 2, 1, 1, 4).setBackground(color);
      });
    }

    sh.autoResizeColumns(1, 4);
  } catch (e) {
    logEvent('WARN', '_rpf_crearInforme_', e.message, '');
  }
}

// ============================================================
// RECONSTRUCCIÓN COMPLETA DE DESPLEGABLES
// ============================================================

/**
 * Reconstruye COMPLETAMENTE todos los desplegables de clientes.
 * NO usa listas antiguas. NO usa cache. NO usa validaciones anteriores.
 * Lee directamente de CLIENTES (Activo=TRUE) y aplica validaciones
 * frescas en todos los paneles y hojas conocidos.
 *
 * Genera SYNC_CLIENTES_DIAGNOSTICO con informe detallado.
 */
function reconstruirDesplegablesClientes() {
  var ui  = SpreadsheetApp.getUi();
  var ss  = getTargetSS_();
  var log = [];
  var nActualizados = 0;

  try {
    // 1. Leer clientes frescos — sin cache, directo de CLIENTES
    var nombres = _sinc_obtenerListaClientes_(ss);
    log.push(['INFO', 'Clientes activos encontrados: ' + nombres.length, nombres.join(', ')]);

    if (nombres.length === 0) {
      ui.alert('⚠️ No hay clientes activos en la hoja CLIENTES.\n'
        + 'Los desplegables no pueden reconstruirse sin datos.');
      return;
    }

    // Validación fresca — construida en este momento, sin reutilizar nada anterior
    var validacion = SpreadsheetApp.newDataValidation()
      .requireValueInList(nombres, true)
      .setAllowInvalid(true)
      .build();

    // 2. PANEL_FACTURAS — coordenadas explícitas (todas)
    var pf = ss.getSheetByName('PANEL_FACTURAS');
    if (pf) {
      _rdc_limpiarYAplicar_(pf, 5,  3, validacion, 'PANEL_FACTURAS C5 (Cliente principal)',       log); nActualizados++;
      _rdc_limpiarYAplicar_(pf, 5,  6, validacion, 'PANEL_FACTURAS F5 (Cliente presupuesto)',      log); nActualizados++;
      _rdc_limpiarYAplicar_(pf, 99, 3, validacion, 'PANEL_FACTURAS C99 (Factura desde Presupuesto)', log); nActualizados++;
    } else {
      log.push(['WARN', 'PANEL_FACTURAS no encontrado', '']);
    }

    // 3. PANEL_CLIENTE — coordenada explícita
    var pc = ss.getSheetByName('PANEL_CLIENTE');
    if (pc) {
      _rdc_limpiarYAplicar_(pc, 4, 3, validacion, 'PANEL_CLIENTE C4', log); nActualizados++;
    } else {
      log.push(['WARN', 'PANEL_CLIENTE no encontrado', '']);
    }

    // 4. Paneles con búsqueda dinámica por etiqueta "Cliente"
    ['PANEL_CONTABLE', 'PANEL_GENERAL', 'PANEL_PENDIENTES', 'PANEL_GASTOS'].forEach(function(nombre) {
      var sh = ss.getSheetByName(nombre);
      if (!sh) { log.push(['SKIP', nombre + ' no existe', '']); return; }
      nActualizados += _rdc_buscarYActualizarPorEtiqueta_(sh, validacion, 'Cliente', 40, nombre, log);
    });

    // 5. Hojas de datos — buscar columna "Cliente" y limpiar validaciones
    ['FACTURAS', 'PROFORMAS', 'PRESUPUESTOS', 'TRABAJOS'].forEach(function(nombre) {
      var sh = ss.getSheetByName(nombre);
      if (!sh) { log.push(['SKIP', nombre + ' no existe', '']); return; }
      nActualizados += _rdc_escanearValidacionesCliente_(sh, validacion, nombre, log);
    });

    // 6. Informe
    _rdc_crearDiagnostico_(ss, nombres, log, nActualizados);

    logEvent('INFO', 'reconstruirDesplegablesClientes',
      nombres.length + ' clientes | ' + nActualizados + ' celdas actualizadas', '');

    ui.alert(
      '✅ Desplegables reconstruidos desde cero',
      'Clientes activos cargados: ' + nombres.length + '\n'
      + 'Celdas actualizadas: ' + nActualizados + '\n\n'
      + 'Validaciones anteriores eliminadas completamente.\n'
      + 'Lista reconstruida directamente de la hoja CLIENTES.\n\n'
      + 'Ver hoja SYNC_CLIENTES_DIAGNOSTICO para el detalle.',
      ui.ButtonSet.OK
    );

  } catch (e) {
    logEvent('ERROR', 'reconstruirDesplegablesClientes', e.message, e.stack);
    ui.alert('❌ Error en reconstrucción:\n' + e.message);
  }
}

/**
 * Elimina la validación actual de una celda y aplica la nueva lista fresca.
 */
function _rdc_limpiarYAplicar_(sh, row, col, validacion, descripcion, log) {
  try {
    var rng = sh.getRange(row, col);
    var valorActual = rng.getValue();
    rng.setDataValidation(null);          // eliminar validación anterior
    rng.setDataValidation(validacion);    // aplicar lista fresca
    log.push(['OK', descripcion, 'Valor actual: "' + valorActual + '"']);
  } catch (e) {
    log.push(['ERROR', descripcion, e.message]);
  }
}

/**
 * Busca la etiqueta "Cliente" en las primeras maxFilas de la hoja y actualiza
 * TODAS las ocurrencias (no solo la primera) con validación fresca.
 */
function _rdc_buscarYActualizarPorEtiqueta_(sh, validacion, etiqueta, maxFilas, nombreHoja, log) {
  var nAct = 0;
  try {
    var lastRow = Math.min(maxFilas, sh.getLastRow());
    var lastCol = Math.min(sh.getLastColumn(), 8);
    if (lastRow < 1 || lastCol < 1) return 0;

    var data = sh.getRange(1, 1, lastRow, lastCol).getValues();
    var patron = etiqueta.toLowerCase();

    for (var r = 0; r < data.length; r++) {
      for (var c = 0; c < data[r].length; c++) {
        var celda = String(data[r][c] || '').toLowerCase().trim();
        if (celda === patron || celda.indexOf(patron) !== -1) {
          var targetCol = c + 2;
          try {
            var rng = sh.getRange(r + 1, targetCol);
            var val = rng.getValue();
            rng.setDataValidation(null);
            rng.setDataValidation(validacion);
            log.push(['OK',
              nombreHoja + ' fila ' + (r + 1) + ' col ' + targetCol + ' (etiqueta "' + data[r][c] + '")',
              'Valor: "' + val + '"']);
            nAct++;
          } catch (e2) {
            log.push(['ERROR', nombreHoja + ' fila ' + (r + 1) + ' col ' + targetCol, e2.message]);
          }
        }
      }
    }

    if (nAct === 0) {
      log.push(['SKIP', nombreHoja + ': sin etiqueta "' + etiqueta + '" en primeras ' + maxFilas + ' filas', '']);
    }
  } catch (e) {
    log.push(['ERROR', nombreHoja + ' búsqueda etiqueta', e.message]);
  }
  return nAct;
}

/**
 * Escanea una hoja de datos buscando columna con cabecera "Cliente" y
 * reemplaza las validaciones que encuentre en esa columna (filas 2-10).
 */
function _rdc_escanearValidacionesCliente_(sh, validacion, nombreHoja, log) {
  var nAct = 0;
  try {
    if (sh.getLastRow() < 1) { log.push(['SKIP', nombreHoja + ': hoja vacía', '']); return 0; }

    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var colsCliente = [];
    headers.forEach(function(h, i) {
      var s = String(h || '').toLowerCase().trim();
      if (s === 'cliente' || s === 'nombre cliente' || s === 'nombre_cliente') {
        colsCliente.push(i + 1);
      }
    });

    if (colsCliente.length === 0) {
      log.push(['SKIP', nombreHoja + ': sin columna "Cliente" en cabecera', '']);
      return 0;
    }

    colsCliente.forEach(function(col) {
      var maxRow = Math.min(sh.getLastRow(), 10);
      for (var r = 2; r <= maxRow; r++) {
        try {
          var rng = sh.getRange(r, col);
          var dv  = rng.getDataValidation();
          if (dv) {
            var val = rng.getValue();
            rng.setDataValidation(null);
            rng.setDataValidation(validacion);
            log.push(['OK', nombreHoja + ' fila ' + r + ' col ' + col + ' (col Cliente)', 'Valor: "' + val + '"']);
            nAct++;
          }
        } catch (e2) {
          log.push(['ERROR', nombreHoja + ' fila ' + r + ' col ' + col, e2.message]);
        }
      }
      if (nAct === 0) {
        log.push(['INFO', nombreHoja + ' col ' + col + ': sin validaciones en filas 2-10', '']);
      }
    });

  } catch (e) {
    log.push(['ERROR', nombreHoja + ' escaneo datos', e.message]);
  }
  return nAct;
}

/**
 * Crea o sobreescribe SYNC_CLIENTES_DIAGNOSTICO con el informe completo.
 */
function _rdc_crearDiagnostico_(ss, nombres, log, nActualizados) {
  try {
    var shName = 'SYNC_CLIENTES_DIAGNOSTICO';
    var sh = ss.getSheetByName(shName);
    if (sh) { sh.clearContents(); sh.clearFormats(); }
    else     { sh = ss.insertSheet(shName); }

    // Cabecera
    sh.getRange(1, 1, 1, 4).setValues([['Estado', 'Descripción', 'Detalle', 'Timestamp']]);
    sh.getRange(1, 1, 1, 4).setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
    sh.setFrozenRows(1);

    // Resumen
    sh.appendRow(['RESUMEN', 'Clientes activos cargados: ' + nombres.length,  nombres.join(' | '), new Date()]);
    sh.appendRow(['RESUMEN', 'Celdas de validación actualizadas: ' + nActualizados, '', new Date()]);
    sh.appendRow(['', '', '', '']);

    // Log detallado
    log.forEach(function(entry) {
      sh.appendRow([entry[0], entry[1], entry[2], new Date()]);
    });

    // Color por estado
    var nData = sh.getLastRow() - 1;
    if (nData > 0) {
      var estados = sh.getRange(2, 1, nData, 1).getValues();
      estados.forEach(function(row, i) {
        var color = { OK: '#e6f4ea', ERROR: '#fce8e6', WARN: '#fef9c3',
                      SKIP: '#f8f9fa', INFO: '#f1f3f4', RESUMEN: '#e8f0fe' }[row[0]] || '#ffffff';
        sh.getRange(i + 2, 1, 1, 4).setBackground(color);
      });
    }

    sh.autoResizeColumns(1, 4);
  } catch (e) {
    logEvent('WARN', '_rdc_crearDiagnostico_', 'No se pudo crear diagnóstico: ' + e.message, '');
  }
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
