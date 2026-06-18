// ================================================================
// LimpiezaClientesReales.gs — Limpieza controlada: conservar SOLO
// Gregor, Patrick Pougeoux y Jose Fouss.
//
// FLUJO OBLIGATORIO:
//   1. Backup completo + verificación
//   2. Informe previo (LCR_INFORME_PREVIO)
//   3. Confirmación manual antes de ejecutar
//   4. Limpieza por hoja
//   5. Integridad referencial post-limpieza
//   6. Detección de duplicados de clientes reales
//   7. Hoja LIMPIEZA_CLIENTES_REALES
//
// REGLAS ABSOLUTAS:
//   - No ejecutar si el backup falla
//   - Gastos con importe → REVISION_MANUAL (no eliminar automáticamente)
//   - Duplicados de clientes reales → NO eliminar, listar para revisión
//   - Informes de cliente desconocido → REVISION_MANUAL
// ================================================================

// Fragmentos de nombre que identifican a cada cliente real
var LCR_CLIENTES_REALES = ['gregor', 'pougeoux', 'fouss'];

// ================================================================
// IDENTIFICADOR DE CLIENTE REAL
// ================================================================

function _lcr_esClienteReal_(texto) {
  if (!texto) return false;
  var t = String(texto).toLowerCase()
    .replace(/[áàä]/g, 'a').replace(/[éèë]/g, 'e')
    .replace(/[íìï]/g, 'i').replace(/[óòö]/g, 'o')
    .replace(/[úùü]/g, 'u');
  return LCR_CLIENTES_REALES.some(function(k) { return t.indexOf(k) !== -1; });
}

// ================================================================
// LOG DE OPERACIONES
// ================================================================

function _lcr_nuevoLog_() {
  return {
    conservados: [],  // { hoja, id, nombre }
    eliminados:  [],  // { hoja, fila, id, nombre, motivo }
    revision:    [],  // { hoja, fila, id, tipo, descripcion }
    duplicados:  []   // { clave, registros[] }
  };
}

function _lcr_registrarConservado_(log, hoja, id, nombre) {
  log.conservados.push({ hoja: hoja, id: id, nombre: nombre });
}

function _lcr_registrarEliminado_(log, hoja, fila, id, nombre, motivo) {
  log.eliminados.push({ hoja: hoja, fila: fila, id: id, nombre: nombre, motivo: motivo });
  logEvent('INFO', 'LCR_LIMPIEZA',
    'ELIMINADO — ' + hoja + ' fila ' + fila + ' [' + id + '] ' + nombre, motivo);
}

function _lcr_registrarRevision_(ss, log, hoja, fila, id, tipo, descripcion, recomendacion) {
  var sh = ss.getSheetByName('REVISION_MANUAL');
  if (!sh) {
    sh = ss.insertSheet('REVISION_MANUAL');
    sh.appendRow(['Timestamp','Hoja','Fila','ID','Tipo_Problema','Descripción','Acción_Recomendada','Estado']);
    applyHeaderStyle_(sh, 8);
  }
  sh.appendRow([new Date(), hoja, fila, id, tipo, descripcion, recomendacion, 'PENDIENTE']);
  log.revision.push({ hoja: hoja, fila: fila, id: id, tipo: tipo, descripcion: descripcion });
  logEvent('WARN', 'LCR_LIMPIEZA', 'REVISION_MANUAL — ' + hoja + ' [' + id + ']', tipo);
}

// ================================================================
// ORQUESTADOR PRINCIPAL
// ================================================================

function ejecutarLimpiezaClientesReales() {
  var ui = SpreadsheetApp.getUi();
  var ss = getTargetSS_();

  var conf1 = ui.alert(
    '🧹 LIMPIEZA CONTROLADA — CLIENTES REALES',
    'Se conservarán ÚNICAMENTE:\n'
    + '  • Gregor\n'
    + '  • Patrick Pougeoux\n'
    + '  • Jose Fouss\n\n'
    + 'Pasos que se ejecutarán:\n'
    + '  1. Backup completo + verificación\n'
    + '  2. Informe previo para revisar\n'
    + '  3. Confirmación antes de limpiar\n'
    + '  4. Limpieza por hoja\n'
    + '  5. Integridad referencial\n'
    + '  6. Detección duplicados\n'
    + '  7. Hoja LIMPIEZA_CLIENTES_REALES\n\n'
    + '¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (conf1 !== ui.Button.YES) { ui.alert('Operación cancelada.'); return; }

  try {

    // ── PASO 1: Backup obligatorio ───────────────────────────────
    ui.alert('PASO 1/5', 'Creando backup completo...\nNo cierres la hoja.', ui.ButtonSet.OK);
    var bkpResultado = ejecutarBackupCompleto();
    if (!bkpResultado || !bkpResultado.ok) {
      throw new Error('El backup no se completó correctamente. LIMPIEZA CANCELADA.');
    }
    var prefijoBkp = bkpResultado.prefijo;

    // ── PASO 2: Informe previo ───────────────────────────────────
    ui.alert('PASO 2/5', 'Generando informe previo...', ui.ButtonSet.OK);
    var informePrevio = _lcr_generarInformePrevio_(ss);

    var shPrevio = ss.getSheetByName('LCR_INFORME_PREVIO');
    if (shPrevio) ss.setActiveSheet(shPrevio);

    var conf2 = ui.alert(
      '📋 INFORME PREVIO — REVISA ANTES DE CONTINUAR',
      'Hoja LCR_INFORME_PREVIO generada.\n\n'
      + 'Registros que se eliminarán:\n'
      + _lcr_resumenEliminaciones_(informePrevio)
      + '\n\n⚠️ Esta acción no se puede deshacer automáticamente.\nEl backup está en hojas BKP_*.\n\n'
      + '¿Ejecutar la limpieza ahora?',
      ui.ButtonSet.YES_NO
    );
    if (conf2 !== ui.Button.YES) {
      ui.alert('Limpieza cancelada.\nEl backup está disponible (prefijo: ' + prefijoBkp + ').');
      return;
    }

    // ── PASO 3: Limpieza ─────────────────────────────────────────
    ui.alert('PASO 3/5', 'Ejecutando limpieza...\nEspera, no cierres la hoja.', ui.ButtonSet.OK);
    var log = _lcr_nuevoLog_();

    var idsReales = _lcr_obtenerIdsClientesReales_(ss);

    // Orden: datos dependientes primero, clientes al final
    _lcr_limpiarTrabajos_(ss, log, idsReales);
    _lcr_limpiarFacturas_(ss, log, idsReales);
    _lcr_limpiarProformas_(ss, log, idsReales);
    _lcr_limpiarPresupuestos_(ss, log);
    _lcr_limpiarGastos_(ss, log);
    _lcr_limpiarInformes_(ss, log, idsReales);
    _lcr_limpiarClientes_(ss, log, idsReales);

    SpreadsheetApp.flush();

    // ── PASO 4: Integridad y duplicados ─────────────────────────
    ui.alert('PASO 4/5', 'Verificando integridad referencial y duplicados...', ui.ButtonSet.OK);
    _lcr_verificarIntegridad_(ss, log);
    _lcr_detectarDuplicadosReales_(ss, log);

    // ── PASO 5: Informe final ────────────────────────────────────
    ui.alert('PASO 5/5', 'Generando hoja LIMPIEZA_CLIENTES_REALES...', ui.ButtonSet.OK);
    _lcr_generarHojaResultados_(ss, log, prefijoBkp);

    var shFinal = ss.getSheetByName('LIMPIEZA_CLIENTES_REALES');
    if (shFinal) ss.setActiveSheet(shFinal);

    ui.alert(
      '✅ LIMPIEZA COMPLETADA',
      'Backup guardado con prefijo: ' + prefijoBkp + '\n\n'
      + 'Registros conservados:           ' + log.conservados.length + '\n'
      + 'Registros eliminados:            ' + log.eliminados.length + '\n'
      + 'Enviados a REVISION_MANUAL:      ' + log.revision.length + '\n'
      + 'Duplicados pendientes (manual):  ' + log.duplicados.length + '\n\n'
      + 'Ver: LIMPIEZA_CLIENTES_REALES y REVISION_MANUAL',
      ui.ButtonSet.OK
    );

  } catch (e) {
    logEvent('ERROR', 'ejecutarLimpiezaClientesReales', e.message, e.stack);
    ui.alert(
      '❌ ERROR EN LIMPIEZA',
      e.message + '\n\nRevisa:\n  • Las hojas BKP_* (backup)\n  • La hoja LOGS',
      ui.ButtonSet.OK
    );
  }
}

// ================================================================
// OBTENER IDs DE CLIENTES REALES (desde hoja CLIENTES)
// ================================================================

function _lcr_obtenerIdsClientesReales_(ss) {
  var idsReales = {};  // { idCliente: nombreMostrar }
  var sh = ss.getSheetByName('CLIENTES');
  if (!sh || sh.getLastRow() < 2) return idsReales;

  var nCols = Math.max(sh.getLastColumn(), 11);
  var data  = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();

  data.forEach(function(row) {
    var id      = String(row[0] || '').trim();
    var nombre  = String(row[2] || '').trim();
    var apes    = String(row[3] || '').trim();
    var empresa = String(row[4] || '').trim();
    var texto   = nombre + ' ' + apes + ' ' + empresa;

    if (id && _lcr_esClienteReal_(texto)) {
      idsReales[id] = (nombre + ' ' + apes).trim() || empresa;
    }
  });

  logEvent('INFO', '_lcr_obtenerIdsClientesReales_',
    Object.keys(idsReales).length + ' IDs de clientes reales identificados', '');
  return idsReales;
}

// ================================================================
// INFORME PREVIO
// ================================================================

function _lcr_generarInformePrevio_(ss) {
  var idsReales = _lcr_obtenerIdsClientesReales_(ss);

  var cfgs = [
    { nombre: 'CLIENTES',     colId: 1, colNomPropio: [3,4,5], colIdCli: null, colNomCli: null, nCols: 11 },
    { nombre: 'TRABAJOS',     colId: 1, colNomPropio: null,     colIdCli: 3,   colNomCli: 4,   nCols: 19 },
    { nombre: 'FACTURAS',     colId: 1, colNomPropio: null,     colIdCli: 5,   colNomCli: 6,   nCols: 17 },
    { nombre: 'PROFORMAS',    colId: 1, colNomPropio: null,     colIdCli: 5,   colNomCli: 6,   nCols: 13 },
    { nombre: 'PRESUPUESTOS', colId: 1, colNomPropio: null,     colIdCli: null,colNomCli: 3,   nCols: 13 },
    { nombre: 'GASTOS',       colId: 1, colNomPropio: null,     colIdCli: null,colNomCli: 24,  nCols: 25 },
    { nombre: 'INFORMES',     colId: 1, colNomPropio: null,     colIdCli: 3,   colNomCli: 4,   nCols: 11 }
  ];

  var resultado = {};

  cfgs.forEach(function(cfg) {
    var sh = ss.getSheetByName(cfg.nombre);
    var total = 0, conservar = 0, eliminar = 0, revision = 0;
    var nomsBorrar = [];

    if (sh && sh.getLastRow() >= 2) {
      var nCols = Math.max(sh.getLastColumn(), cfg.nCols);
      var data  = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();

      data.forEach(function(row) {
        total++;
        var textoNombre = '';

        if (cfg.colNomPropio) {
          cfg.colNomPropio.forEach(function(c) {
            textoNombre += ' ' + String(row[c - 1] || '');
          });
        }
        if (cfg.colNomCli) textoNombre += ' ' + String(row[cfg.colNomCli - 1] || '');
        textoNombre = textoNombre.trim();

        var idCli = cfg.colIdCli ? String(row[cfg.colIdCli - 1] || '').trim() : null;
        var esReal = _lcr_esClienteReal_(textoNombre) || (idCli && idsReales[idCli]);

        if (esReal) {
          conservar++;
        } else {
          // GASTOS e INFORMES tienen lógica especial en la limpieza real
          if (cfg.nombre === 'GASTOS') {
            var total_ = parseFloat(row[COL_GAS.TOTAL - 1]) || 0;
            var esPrueba = _lim_esPrueba_({ nombre: textoNombre, id: String(row[0] || '') });
            if (esPrueba && total_ === 0) { eliminar++; }
            else if (total_ === 0)        { eliminar++; }
            else                          { revision++; }
          } else if (cfg.nombre === 'INFORMES') {
            var esPrueba2 = _lim_esPrueba_({ nombre: textoNombre, id: String(row[0] || '') });
            if (esPrueba2) { eliminar++; }
            else            { revision++; }
          } else {
            eliminar++;
            if (textoNombre && nomsBorrar.length < 15) nomsBorrar.push(textoNombre);
          }
        }
      });
    }

    resultado[cfg.nombre] = {
      total: total, conservar: conservar, eliminar: eliminar,
      revision: revision, nomsBorrar: nomsBorrar
    };
  });

  // Generar hoja LCR_INFORME_PREVIO
  var sh = ss.getSheetByName('LCR_INFORME_PREVIO');
  if (sh) { sh.clear(); sh.clearFormats(); }
  else    { sh = ss.insertSheet('LCR_INFORME_PREVIO'); }

  var tz    = Session.getScriptTimeZone();
  var ahora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');

  sh.setColumnWidth(1, 185); sh.setColumnWidth(2, 75);
  sh.setColumnWidth(3, 90);  sh.setColumnWidth(4, 90);
  sh.setColumnWidth(5, 90);  sh.setColumnWidth(6, 380);

  var r = 1;

  sh.getRange(r, 1, 1, 6).merge()
    .setValue('📋 INFORME PREVIO — LIMPIEZA CLIENTES REALES — ' + ahora)
    .setBackground('#1a237e').setFontColor('#ffffff')
    .setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(r, 42); r++;

  sh.getRange(r, 1, 1, 6).merge()
    .setValue('Clientes que se CONSERVARÁN: Gregor  |  Patrick Pougeoux  |  Jose Fouss')
    .setBackground('#e8eaf6').setFontColor('#1a237e')
    .setFontWeight('bold').setHorizontalAlignment('center');
  r += 2;

  sh.getRange(r, 1, 1, 6).setValues([[
    'Hoja', 'Total', 'Conservar', 'Eliminar', 'Revisión', 'Clientes a eliminar (muestra)'
  ]]).setBackground('#202124').setFontColor('#ffffff').setFontWeight('bold');
  r++;

  var HOJAS_ORDEN = ['CLIENTES','TRABAJOS','FACTURAS','PROFORMAS','PRESUPUESTOS','GASTOS','INFORMES'];
  HOJAS_ORDEN.forEach(function(nombre, i) {
    var d  = resultado[nombre] || { total: 0, conservar: 0, eliminar: 0, revision: 0, nomsBorrar: [] };
    var bg = i % 2 === 0 ? '#f3f4f6' : '#e8eaf6';
    sh.getRange(r, 1, 1, 6).setBackground(bg);
    sh.getRange(r, 1).setValue(nombre).setFontWeight('bold');
    sh.getRange(r, 2).setValue(d.total).setHorizontalAlignment('center');
    sh.getRange(r, 3).setValue(d.conservar).setFontColor('#137333').setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(r, 4).setValue(d.eliminar).setFontColor(d.eliminar > 0 ? '#c5221f' : '#137333').setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(r, 5).setValue(d.revision).setFontColor(d.revision > 0 ? '#e65100' : '#137333').setHorizontalAlignment('center');
    sh.getRange(r, 6).setValue(d.nomsBorrar.join('  |  ')).setFontSize(8).setWrap(true);
    sh.setRowHeight(r, 32); r++;
  });

  r++;
  sh.getRange(r, 1, 1, 6).merge()
    .setValue('NOTA: "Revisión" = enviados a REVISION_MANUAL (no eliminados automáticamente). '
      + 'GASTOS con importe > 0 y INFORMES de clientes desconocidos van siempre a revisión.')
    .setBackground('#fff3cd').setFontColor('#856404').setWrap(true).setFontSize(9);

  SpreadsheetApp.flush();
  return resultado;
}

function _lcr_resumenEliminaciones_(informe) {
  var lineas = [];
  ['CLIENTES','TRABAJOS','FACTURAS','PROFORMAS','PRESUPUESTOS','GASTOS','INFORMES'].forEach(function(h) {
    var d = informe[h];
    if (!d) return;
    var partes = [];
    if (d.eliminar > 0) partes.push(d.eliminar + ' eliminados');
    if (d.revision > 0) partes.push(d.revision + ' a revisión');
    if (partes.length) lineas.push('  • ' + h + ': ' + partes.join(', ') + ' (de ' + d.total + ')');
  });
  return lineas.length ? lineas.join('\n') : '  Ninguno detectado aún (se calculará al limpiar).';
}

// ================================================================
// LIMPIEZA POR HOJA
// ================================================================

// CLIENTES
function _lcr_limpiarClientes_(ss, log, idsReales) {
  var sh = ss.getSheetByName('CLIENTES');
  if (!sh || sh.getLastRow() < 2) return;

  var nCols = Math.max(sh.getLastColumn(), 11);
  var data  = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();
  var filasEliminar = [];

  data.forEach(function(row, idx) {
    var fila    = idx + 2;
    var id      = String(row[0] || '').trim();
    var nombre  = String(row[2] || '').trim();
    var apes    = String(row[3] || '').trim();
    var empresa = String(row[4] || '').trim();
    var texto   = nombre + ' ' + apes + ' ' + empresa;
    var nomMostrar = (nombre + ' ' + apes).trim() || empresa;

    if (idsReales[id] || _lcr_esClienteReal_(texto)) {
      _lcr_registrarConservado_(log, 'CLIENTES', id, nomMostrar);
    } else {
      filasEliminar.push({ fila: fila, id: id, nombre: nomMostrar });
    }
  });

  filasEliminar.sort(function(a, b) { return b.fila - a.fila; });
  filasEliminar.forEach(function(item) {
    sh.deleteRow(item.fila);
    _lcr_registrarEliminado_(log, 'CLIENTES', item.fila, item.id, item.nombre,
      'No pertenece a ningún cliente real (Gregor / Pougeoux / Fouss)');
  });
}

// TRABAJOS
function _lcr_limpiarTrabajos_(ss, log, idsReales) {
  var sh = ss.getSheetByName('TRABAJOS');
  if (!sh || sh.getLastRow() < 2) return;

  var nCols = Math.max(sh.getLastColumn(), 19);
  var data  = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();
  var filasEliminar = [];

  data.forEach(function(row, idx) {
    var fila   = idx + 2;
    var id     = String(row[0] || '').trim();
    var idCli  = String(row[2] || '').trim();
    var nomCli = String(row[3] || '').trim();

    if (idsReales[idCli] || _lcr_esClienteReal_(nomCli)) {
      _lcr_registrarConservado_(log, 'TRABAJOS', id, nomCli);
    } else {
      filasEliminar.push({ fila: fila, id: id, nombre: nomCli });
    }
  });

  filasEliminar.sort(function(a, b) { return b.fila - a.fila; });
  filasEliminar.forEach(function(item) {
    sh.deleteRow(item.fila);
    _lcr_registrarEliminado_(log, 'TRABAJOS', item.fila, item.id, item.nombre,
      'Trabajo de cliente no real');
  });
}

// FACTURAS
function _lcr_limpiarFacturas_(ss, log, idsReales) {
  var sh = ss.getSheetByName('FACTURAS');
  if (!sh || sh.getLastRow() < 2) return;

  var nCols = Math.max(sh.getLastColumn(), 17);
  var data  = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();
  var filasEliminar = [];

  data.forEach(function(row, idx) {
    var fila   = idx + 2;
    var id     = String(row[0] || '').trim();
    var idCli  = String(row[4] || '').trim();
    var nomCli = String(row[5] || '').trim();

    if (idsReales[idCli] || _lcr_esClienteReal_(nomCli)) {
      _lcr_registrarConservado_(log, 'FACTURAS', id, nomCli);
    } else {
      filasEliminar.push({ fila: fila, id: id, nombre: nomCli });
    }
  });

  filasEliminar.sort(function(a, b) { return b.fila - a.fila; });
  filasEliminar.forEach(function(item) {
    sh.deleteRow(item.fila);
    _lcr_registrarEliminado_(log, 'FACTURAS', item.fila, item.id, item.nombre,
      'Factura de cliente no real');
  });
}

// PROFORMAS
function _lcr_limpiarProformas_(ss, log, idsReales) {
  var sh = ss.getSheetByName('PROFORMAS');
  if (!sh || sh.getLastRow() < 2) return;

  var nCols = Math.max(sh.getLastColumn(), 13);
  var data  = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();
  var filasEliminar = [];

  data.forEach(function(row, idx) {
    var fila   = idx + 2;
    var id     = String(row[0] || '').trim();
    var idCli  = String(row[4] || '').trim();
    var nomCli = String(row[5] || '').trim();

    if (idsReales[idCli] || _lcr_esClienteReal_(nomCli)) {
      _lcr_registrarConservado_(log, 'PROFORMAS', id, nomCli);
    } else {
      filasEliminar.push({ fila: fila, id: id, nombre: nomCli });
    }
  });

  filasEliminar.sort(function(a, b) { return b.fila - a.fila; });
  filasEliminar.forEach(function(item) {
    sh.deleteRow(item.fila);
    _lcr_registrarEliminado_(log, 'PROFORMAS', item.fila, item.id, item.nombre,
      'Proforma de cliente no real');
  });
}

// PRESUPUESTOS (match por nombre en col 3)
function _lcr_limpiarPresupuestos_(ss, log) {
  var sh = ss.getSheetByName('PRESUPUESTOS');
  if (!sh || sh.getLastRow() < 2) return;

  var nCols = Math.max(sh.getLastColumn(), 13);
  var data  = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();
  var filasEliminar = [];

  data.forEach(function(row, idx) {
    var fila    = idx + 2;
    var id      = String(row[COL_PRE.ID      - 1] || '').trim();
    var cliente = String(row[COL_PRE.CLIENTE - 1] || '').trim();

    if (_lcr_esClienteReal_(cliente)) {
      _lcr_registrarConservado_(log, 'PRESUPUESTOS', id, cliente);
    } else {
      filasEliminar.push({ fila: fila, id: id, nombre: cliente });
    }
  });

  filasEliminar.sort(function(a, b) { return b.fila - a.fila; });
  filasEliminar.forEach(function(item) {
    sh.deleteRow(item.fila);
    _lcr_registrarEliminado_(log, 'PRESUPUESTOS', item.fila, item.id, item.nombre,
      'Presupuesto de cliente no real');
  });
}

// GASTOS — regla especial:
//   • Prueba + importe 0  → eliminar
//   • Sin datos + importe 0 → eliminar
//   • Cualquier gasto con importe > 0 → REVISION_MANUAL (no eliminar)
function _lcr_limpiarGastos_(ss, log) {
  var sh = ss.getSheetByName('GASTOS');
  if (!sh || sh.getLastRow() < 2) return;

  var nCols = Math.max(sh.getLastColumn(), 25);
  var data  = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();
  var filasEliminar = [];

  data.forEach(function(row, idx) {
    var fila      = idx + 2;
    var id        = String(row[COL_GAS.ID_GASTO   - 1] || '').trim();
    var proveedor = String(row[COL_GAS.PROVEEDOR  - 1] || '').trim();
    var concepto  = String(row[COL_GAS.CONCEPTO   - 1] || '').trim();
    var total     = parseFloat(row[COL_GAS.TOTAL  - 1]) || 0;
    var cliente   = String(row[COL_GAS.CLIENTE    - 1] || '').trim();
    var nomMostrar = proveedor || concepto || id;
    var esPrueba  = _lim_esPrueba_({ nombre: proveedor + ' ' + concepto + ' ' + cliente, id: id });
    var esReal    = _lcr_esClienteReal_(cliente);

    if (total > 0) {
      // Gastos con importe siempre van a revisión manual
      var tipo = esReal   ? 'GASTO_CLIENTE_REAL'
               : esPrueba ? 'GASTO_PRUEBA_CON_IMPORTE'
               :             'GASTO_CLIENTE_DESCONOCIDO';
      var recoms = esReal
        ? 'Gasto asociado a cliente real. Conservar si es correcto; eliminar si es de prueba.'
        : esPrueba
          ? 'Gasto de prueba con importe. Verificar y eliminar manualmente.'
          : 'Gasto de cliente desconocido. Determinar si es real o de prueba.';
      _lcr_registrarRevision_(ss, log, 'GASTOS', fila, id, tipo,
        '"' + nomMostrar + '" — total: ' + total.toFixed(2) + '€ — cliente: ' + (cliente || 'sin asignar'),
        recoms);
    } else {
      // Sin importe → eliminar (vacío o de prueba)
      filasEliminar.push({ fila: fila, id: id, nombre: nomMostrar });
    }
  });

  filasEliminar.sort(function(a, b) { return b.fila - a.fila; });
  filasEliminar.forEach(function(item) {
    sh.deleteRow(item.fila);
    _lcr_registrarEliminado_(log, 'GASTOS', item.fila, item.id, item.nombre,
      'Gasto sin importe (vacío o de prueba)');
  });
}

// INFORMES — solo eliminar los claramente de prueba
//   • Cliente real → conservar
//   • Prueba → eliminar
//   • Cliente desconocido → REVISION_MANUAL
function _lcr_limpiarInformes_(ss, log, idsReales) {
  var sh = ss.getSheetByName('INFORMES');
  if (!sh || sh.getLastRow() < 2) return;

  var nCols = Math.max(sh.getLastColumn(), 11);
  var data  = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();
  var filasEliminar = [];

  data.forEach(function(row, idx) {
    var fila   = idx + 2;
    var id     = String(row[0] || '').trim();
    var idCli  = String(row[2] || '').trim();  // ID_Cliente (col 3)
    var nomCli = String(row[3] || '').trim();  // Nombre_Cliente (col 4)
    var esReal   = idsReales[idCli] || _lcr_esClienteReal_(nomCli);
    var esPrueba = _lim_esPrueba_({ nombre: nomCli, id: id });

    if (esReal) {
      _lcr_registrarConservado_(log, 'INFORMES', id, nomCli);
    } else if (esPrueba) {
      filasEliminar.push({ fila: fila, id: id, nombre: nomCli });
    } else {
      _lcr_registrarRevision_(ss, log, 'INFORMES', fila, id,
        'INFORME_CLIENTE_DESCONOCIDO',
        'Informe de "' + (nomCli || 'sin nombre') + '" — cliente no identificado como real ni de prueba.',
        'Verificar si pertenece a un cliente real antes de eliminar manualmente.');
    }
  });

  filasEliminar.sort(function(a, b) { return b.fila - a.fila; });
  filasEliminar.forEach(function(item) {
    sh.deleteRow(item.fila);
    _lcr_registrarEliminado_(log, 'INFORMES', item.fila, item.id, item.nombre,
      'Informe de prueba');
  });
}

// ================================================================
// INTEGRIDAD REFERENCIAL (post-limpieza)
// ================================================================

function _lcr_verificarIntegridad_(ss, log) {
  var idsClientes = {};
  var shCli = ss.getSheetByName('CLIENTES');
  if (shCli && shCli.getLastRow() >= 2) {
    shCli.getRange(2, 1, shCli.getLastRow() - 1, 1).getValues().forEach(function(r) {
      var id = String(r[0] || '').trim();
      if (id) idsClientes[id] = true;
    });
  }

  // TRABAJOS → CLIENTES
  var shTra = ss.getSheetByName('TRABAJOS');
  if (shTra && shTra.getLastRow() >= 2) {
    shTra.getRange(2, 1, shTra.getLastRow() - 1, 4).getValues().forEach(function(row, idx) {
      var idCli = String(row[2] || '').trim();
      var nom   = String(row[3] || '').trim();
      if (idCli && !idsClientes[idCli]) {
        _lcr_registrarRevision_(ss, log, 'TRABAJOS', idx + 2, String(row[0] || '').trim(),
          'HUERFANO_CLIENTE',
          'ID_Cliente "' + idCli + '" (' + nom + ') no existe en CLIENTES tras la limpieza.',
          'Verificar si el cliente fue eliminado por error o si el ID es incorrecto.');
      }
    });
  }

  // FACTURAS → CLIENTES
  var shFac = ss.getSheetByName('FACTURAS');
  if (shFac && shFac.getLastRow() >= 2) {
    shFac.getRange(2, 1, shFac.getLastRow() - 1, 6).getValues().forEach(function(row, idx) {
      var idCli = String(row[4] || '').trim();
      var nom   = String(row[5] || '').trim();
      if (idCli && !idsClientes[idCli]) {
        _lcr_registrarRevision_(ss, log, 'FACTURAS', idx + 2, String(row[0] || '').trim(),
          'HUERFANO_CLIENTE',
          'ID_Cliente "' + idCli + '" (' + nom + ') no existe en CLIENTES. RIESGO: factura sin cliente.',
          'Restaurar cliente desde backup BKP_* o corregir el ID_Cliente.');
      }
    });
  }

  // PROFORMAS → CLIENTES
  var shPro = ss.getSheetByName('PROFORMAS');
  if (shPro && shPro.getLastRow() >= 2) {
    shPro.getRange(2, 1, shPro.getLastRow() - 1, 6).getValues().forEach(function(row, idx) {
      var idCli = String(row[4] || '').trim();
      var nom   = String(row[5] || '').trim();
      if (idCli && !idsClientes[idCli]) {
        _lcr_registrarRevision_(ss, log, 'PROFORMAS', idx + 2, String(row[0] || '').trim(),
          'HUERFANO_CLIENTE',
          'ID_Cliente "' + idCli + '" (' + nom + ') no existe en CLIENTES.',
          'Asignar el cliente correcto o revisar si fue eliminado por error.');
      }
    });
  }

  logEvent('INFO', '_lcr_verificarIntegridad_',
    'Integridad verificada. Incidencias en revision: ' + log.revision.length, '');
}

// ================================================================
// DETECCIÓN DE DUPLICADOS DE CLIENTES REALES
// ================================================================

function _lcr_detectarDuplicadosReales_(ss, log) {
  var sh = ss.getSheetByName('CLIENTES');
  if (!sh || sh.getLastRow() < 2) return;

  var nCols = Math.max(sh.getLastColumn(), 11);
  var data  = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();
  var grupos = {};  // { clave: [{ fila, id, nombre }] }

  data.forEach(function(row, idx) {
    var id      = String(row[0] || '').trim();
    var nombre  = String(row[2] || '').trim();
    var apes    = String(row[3] || '').trim();
    var empresa = String(row[4] || '').trim();
    var texto   = (nombre + ' ' + apes).trim().toLowerCase() || empresa.toLowerCase();

    LCR_CLIENTES_REALES.forEach(function(clave) {
      if (texto.indexOf(clave) !== -1) {
        if (!grupos[clave]) grupos[clave] = [];
        grupos[clave].push({
          fila: idx + 2, id: id,
          nombre: (nombre + ' ' + apes).trim() || empresa
        });
      }
    });
  });

  Object.keys(grupos).forEach(function(clave) {
    var grupo = grupos[clave];
    if (grupo.length > 1) {
      log.duplicados.push({ clave: clave, registros: grupo });
      grupo.forEach(function(g) {
        _lcr_registrarRevision_(ss, log, 'CLIENTES', g.fila, g.id,
          'DUPLICADO_CLIENTE_REAL',
          'Cliente real "' + clave + '" aparece ' + grupo.length + ' veces. IDs: '
            + grupo.map(function(x) { return x.id; }).join(', '),
          'NO eliminar automáticamente. Comparar los ' + grupo.length
            + ' registros y consolidar manualmente el correcto.');
      });
    }
  });

  if (log.duplicados.length > 0) {
    logEvent('WARN', '_lcr_detectarDuplicadosReales_',
      log.duplicados.length + ' clientes reales con duplicados detectados', '');
  }
}

// ================================================================
// HOJA DE RESULTADOS FINAL — LIMPIEZA_CLIENTES_REALES
// ================================================================

function _lcr_generarHojaResultados_(ss, log, prefijoBkp) {
  var sh = ss.getSheetByName('LIMPIEZA_CLIENTES_REALES');
  if (sh) { sh.clear(); sh.clearFormats(); }
  else    { sh = ss.insertSheet('LIMPIEZA_CLIENTES_REALES'); }

  var tz    = Session.getScriptTimeZone();
  var ahora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
  var NC    = 5;

  sh.setColumnWidth(1, 175); sh.setColumnWidth(2, 82);
  sh.setColumnWidth(3, 165); sh.setColumnWidth(4, 420);
  sh.setColumnWidth(5, 200);

  var r = 1;

  // ── Cabecera ─────────────────────────────────────────────────
  sh.getRange(r, 1, 1, NC).merge()
    .setValue('🧹 LIMPIEZA CLIENTES REALES — ' + ahora)
    .setBackground('#1a237e').setFontColor('#ffffff')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(r, 52); r++;

  sh.getRange(r, 1, 1, NC).merge()
    .setValue('Backup: ' + (prefijoBkp || '—')
      + '   |   Conservados: Gregor | Patrick Pougeoux | Jose Fouss')
    .setBackground('#e8eaf6').setFontColor('#1a237e').setHorizontalAlignment('center');
  r += 2;

  // ── Resumen ───────────────────────────────────────────────────
  sh.getRange(r, 1, 1, NC).merge()
    .setValue('📊 RESUMEN')
    .setBackground('#202124').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  r++;

  var filas_resumen = [
    ['Registros conservados',                   log.conservados.length, '#137333'],
    ['Registros eliminados',                     log.eliminados.length,  '#c5221f'],
    ['Enviados a REVISION_MANUAL',               log.revision.length,    '#e65100'],
    ['Duplicados clientes reales (revisar)',      log.duplicados.length,  '#6a1b9a']
  ];
  filas_resumen.forEach(function(d, i) {
    var bg = i % 2 === 0 ? '#e8eaf6' : '#f3f4f6';
    sh.getRange(r, 1, 1, NC).setBackground(bg);
    sh.getRange(r, 1).setValue(d[0]).setFontWeight('bold');
    sh.getRange(r, 2).setValue(d[1])
      .setFontColor(d[2]).setFontWeight('bold')
      .setHorizontalAlignment('center').setFontSize(12);
    r++;
  });
  r++;

  // ── Registros conservados ─────────────────────────────────────
  if (log.conservados.length > 0) {
    sh.getRange(r, 1, 1, NC).merge()
      .setValue('✅ REGISTROS CONSERVADOS (' + log.conservados.length + ')')
      .setBackground('#1b5e20').setFontColor('#ffffff').setFontWeight('bold');
    r++;
    sh.getRange(r, 1, 1, 3).setValues([['Hoja', 'ID', 'Nombre / Cliente']])
      .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
    r++;
    log.conservados.forEach(function(item, i) {
      var bg = i % 2 === 0 ? '#f0fff4' : '#e8f5e9';
      sh.getRange(r, 1).setValue(item.hoja).setBackground(bg);
      sh.getRange(r, 2).setValue(item.id).setBackground(bg).setFontSize(9);
      sh.getRange(r, 3).setValue(item.nombre).setBackground(bg);
      sh.setRowHeight(r, 22); r++;
    });
    r++;
  }

  // ── Registros eliminados ──────────────────────────────────────
  if (log.eliminados.length > 0) {
    sh.getRange(r, 1, 1, NC).merge()
      .setValue('🗑️ REGISTROS ELIMINADOS (' + log.eliminados.length + ')')
      .setBackground('#b71c1c').setFontColor('#ffffff').setFontWeight('bold');
    r++;
    sh.getRange(r, 1, 1, NC).setValues([['Hoja', 'Fila', 'ID', 'Nombre', 'Motivo']])
      .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
    r++;
    log.eliminados.forEach(function(item, i) {
      var bg = i % 2 === 0 ? '#fff3f3' : '#ffebee';
      sh.getRange(r, 1).setValue(item.hoja).setBackground(bg);
      sh.getRange(r, 2).setValue(item.fila).setBackground(bg).setHorizontalAlignment('center');
      sh.getRange(r, 3).setValue(item.id).setBackground(bg).setFontSize(9);
      sh.getRange(r, 4).setValue(item.nombre).setBackground(bg);
      sh.getRange(r, 5).setValue(item.motivo).setBackground(bg).setFontSize(9).setWrap(true);
      sh.setRowHeight(r, 28); r++;
    });
    r++;
  }

  // ── Incidencias enviadas a REVISION_MANUAL ────────────────────
  if (log.revision.length > 0) {
    sh.getRange(r, 1, 1, NC).merge()
      .setValue('⚠️ INCIDENCIAS — ENVIADAS A REVISION_MANUAL (' + log.revision.length + ')')
      .setBackground('#e65100').setFontColor('#ffffff').setFontWeight('bold');
    r++;
    sh.getRange(r, 1, 1, 4).setValues([['Hoja', 'Fila', 'ID', 'Tipo']])
      .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
    r++;
    log.revision.forEach(function(item, i) {
      var bg = i % 2 === 0 ? '#fff8e1' : '#fff3cd';
      sh.getRange(r, 1).setValue(item.hoja).setBackground(bg);
      sh.getRange(r, 2).setValue(item.fila).setBackground(bg).setHorizontalAlignment('center');
      sh.getRange(r, 3).setValue(item.id).setBackground(bg).setFontSize(9);
      sh.getRange(r, 4).setValue(item.tipo).setBackground(bg).setFontSize(9);
      sh.setRowHeight(r, 22); r++;
    });
    r++;
  }

  // ── Duplicados de clientes reales ─────────────────────────────
  if (log.duplicados.length > 0) {
    sh.getRange(r, 1, 1, NC).merge()
      .setValue('🔍 POSIBLES DUPLICADOS — CLIENTES REALES — REVISIÓN MANUAL OBLIGATORIA')
      .setBackground('#4a148c').setFontColor('#ffffff').setFontWeight('bold');
    r++;

    log.duplicados.forEach(function(dup) {
      sh.getRange(r, 1, 1, NC).merge()
        .setValue('  ▶  "' + dup.clave.toUpperCase() + '"  —  '
          + dup.registros.length + ' registros encontrados')
        .setBackground('#ede7f6').setFontColor('#4a148c').setFontWeight('bold');
      r++;
      dup.registros.forEach(function(reg) {
        sh.getRange(r, 2).setValue('Fila ' + reg.fila).setBackground('#f3e5f5').setFontSize(9);
        sh.getRange(r, 3).setValue(reg.id).setBackground('#f3e5f5').setFontSize(9);
        sh.getRange(r, 4).setValue(reg.nombre).setBackground('#f3e5f5');
        sh.setRowHeight(r, 22); r++;
      });
    });
    r++;
  }

  // ── Sin incidencias ───────────────────────────────────────────
  if (log.revision.length === 0 && log.duplicados.length === 0) {
    sh.getRange(r, 1, 1, NC).merge()
      .setValue('✅ Sin incidencias ni duplicados detectados.')
      .setBackground('#e8f5e9').setFontColor('#137333').setFontWeight('bold');
    r++;
  }

  // ── Nota de restauración ──────────────────────────────────────
  r++;
  sh.getRange(r, 1, 1, NC).merge()
    .setValue('📋 RESTAURACIÓN: usar hojas con prefijo "' + (prefijoBkp || 'BKP_*')
      + '". Copiar los datos del backup a la hoja original. '
      + 'No eliminar las hojas de backup hasta verificar que todo es correcto.')
    .setBackground('#fff3cd').setFontColor('#856404').setWrap(true).setFontSize(10);

  SpreadsheetApp.flush();
  logEvent('INFO', '_lcr_generarHojaResultados_',
    'LIMPIEZA_CLIENTES_REALES generada. Conservados: ' + log.conservados.length
    + ' | Eliminados: ' + log.eliminados.length
    + ' | Revisión: ' + log.revision.length
    + ' | Duplicados: ' + log.duplicados.length, '');
}

// ================================================================
// FUNCIÓN STUB — MENÚ DESACTIVADO
//
// BUG IDENTIFICADO Y POR QUÉ SE BORRARON TRABAJOS REALES:
//
// La función _lcr_esClienteReal_() solo reconoce como "reales"
// los registros cuyo Nombre_Cliente contiene exactamente una de
// estas cadenas: 'gregor', 'pougeoux', 'fouss'.
//
// Cualquier trabajo con un cliente que NO contenga esas cadenas
// era eliminado, incluyendo:
//   1. Trabajos de clientes reales distintos a los 3 nombrados
//      (p. ej. clientes de Álvaro que no son los 3 indicados).
//   2. Trabajos donde el nombre estaba vacío o abreviado y el
//      ID_Cliente no coincidía con ningún ID en CLIENTES
//      (porque CLIENTES ya había sido parcialmente limpiada).
//   3. Registros donde la comparación se hacía sobre
//      Nombre_Cliente (col 4 en TRABAJOS) pero el nombre real
//      del cliente estaba almacenado de forma diferente.
//
// CORRECCIÓN PENDIENTE ANTES DE REACTIVAR:
//   - Añadir un paso previo que liste todos los clientes únicos
//     presentes en TRABAJOS/FACTURAS y pida confirmación
//     explícita para cada uno antes de eliminar.
//   - O bien: eliminar solo los que coincidan con la lista negra
//     de prueba (usando _lim_esPrueba_), no todo lo que no sea
//     de los 3 clientes nombrados.
//   - En ningún caso eliminar un registro con factura o importe
//     real sin confirmación manual individual.
// ================================================================

function _lcr_funcionDesactivada_() {
  SpreadsheetApp.getUi().alert(
    '🚫 FUNCIÓN DESACTIVADA',
    'La limpieza de clientes reales está temporalmente desactivada.\n\n'
    + 'Motivo: eliminó trabajos reales al no reconocer clientes\n'
    + 'que no son Gregor, Pougeoux ni Fouss.\n\n'
    + 'Para restaurar los datos usa:\n'
    + '  ⚙️ Sistema → 🔴 RESTAURAR — Desde backup\n\n'
    + 'La función no se reactivará hasta corregir el criterio\n'
    + 'de conservación.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
