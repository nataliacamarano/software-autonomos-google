// ============================================================
// Limpieza.gs — FASES B, C, D: limpieza segura de datos,
// corrección de duplicados e integridad referencial.
// REGLAS ABSOLUTAS:
//   - NO eliminar sin backup previo verificado
//   - NO eliminar si el registro tiene ID_Factura asociado
//   - NO eliminar registros 🔴 ni 🟡
//   - TODO cambio queda registrado
// ============================================================

var LIM_HOJAS_DATOS = [
  'CLIENTES', 'TRABAJOS', 'FACTURAS',
  'PROFORMAS', 'PRESUPUESTOS', 'GASTOS'
];

var LIM_PRUEBA_KEYWORDS = [
  'test','prueba','ejemplo','demo','sample','ficticio','falso','dummy',
  'fake','temporal','borrar','eliminar','mock','xxx','zzz','aaa',
  'cliente prueba','cliente test'
];
var LIM_PRUEBA_EMAILS = [
  '@test.com','@example.com','@prueba.com','@demo.com','@fake.com'
];
var LIM_PRUEBA_CIFS = [
  'A00000000','B00000000','00000000A','00000000Z',
  'B12345678','A12345678','12345678A','11111111H'
];
var LIM_PRUEBA_DIRS = ['calle ejemplo','calle test','calle prueba'];

// ---- LOG DE CAMBIOS ----

function _lim_nuevaLog_() {
  return { eliminados: [], revisionManual: [], errores: [] };
}

function _lim_registrarEliminado_(log, hoja, fila, id, motivo) {
  log.eliminados.push({ hoja: hoja, fila: fila, id: id, motivo: motivo, ts: new Date() });
  logEvent('INFO', 'LIMPIEZA', 'ELIMINADO — ' + hoja + ' fila ' + fila + ' ID ' + id, motivo);
}

function _lim_registrarRevision_(ss, hoja, fila, id, tipo, descripcion, recomendacion) {
  var sh = ss.getSheetByName('REVISION_MANUAL');
  if (!sh) {
    sh = ss.insertSheet('REVISION_MANUAL');
    sh.appendRow(['Timestamp','Hoja','Fila','ID','Tipo_Problema','Descripción','Acción_Recomendada','Estado']);
    applyHeaderStyle_(sh, 8);
  }
  sh.appendRow([
    new Date(), hoja, fila, id, tipo,
    descripcion, recomendacion, 'PENDIENTE'
  ]);
  logEvent('WARN', 'LIMPIEZA', 'REVISIÓN_MANUAL — ' + hoja + ' fila ' + fila + ' ID ' + id, tipo);
}

// ---- CLASIFICADOR INLINE ----

function _lim_esPrueba_(campos) {
  var texto = [
    campos.nombre || '', campos.email || '',
    campos.cif || '', campos.dir || '', campos.id || ''
  ].join(' ').toLowerCase().trim();

  for (var i = 0; i < LIM_PRUEBA_KEYWORDS.length; i++) {
    if (texto.indexOf(LIM_PRUEBA_KEYWORDS[i]) !== -1) return true;
  }
  if (campos.email) {
    var emailL = campos.email.toLowerCase();
    for (var j = 0; j < LIM_PRUEBA_EMAILS.length; j++) {
      if (emailL.indexOf(LIM_PRUEBA_EMAILS[j]) !== -1) return true;
    }
  }
  if (campos.cif) {
    var cifU = campos.cif.toUpperCase().replace(/\s/g, '');
    if (LIM_PRUEBA_CIFS.indexOf(cifU) !== -1) return true;
  }
  if (campos.dir) {
    var dirL = campos.dir.toLowerCase();
    for (var k = 0; k < LIM_PRUEBA_DIRS.length; k++) {
      if (dirL.indexOf(LIM_PRUEBA_DIRS[k]) !== -1) return true;
    }
  }
  if (campos.nombre && campos.nombre.toLowerCase().indexOf('mi empresa') !== -1) return true;
  return false;
}

// ============================================================
// ORQUESTADOR PRINCIPAL — ejecuta A → B → C → D
// ============================================================

function ejecutarFase3Completa() {
  var ui = SpreadsheetApp.getUi();

  var confirma = ui.alert(
    '⚠️ FASE 3 — LIMPIEZA SEGURA',
    'Se ejecutarán en orden:\n'
    + '  A) Backup automático\n'
    + '  B) Limpieza de datos de prueba\n'
    + '  C) Corrección de duplicados exactos\n'
    + '  D) Revisión de integridad referencial\n\n'
    + '¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (confirma !== ui.Button.YES) { ui.alert('Operación cancelada.'); return; }

  try {
    // FASE A — backup obligatorio
    ui.alert('PASO 1/4: Creando backup...', '', ui.ButtonSet.OK);
    var bkpResultado = ejecutarBackupCompleto();

    // FASE B — limpieza de prueba
    ui.alert('PASO 2/4: Limpiando datos de prueba...', '', ui.ButtonSet.OK);
    var logB = ejecutarFaseB_LimpiezaPrueba(true);  // true = modo silencioso

    // FASE C — corrección de duplicados
    ui.alert('PASO 3/4: Corrigiendo duplicados exactos...', '', ui.ButtonSet.OK);
    var logC = ejecutarFaseC_CorreccionDuplicados(true);

    // FASE D — integridad referencial
    ui.alert('PASO 4/4: Verificando integridad referencial...', '', ui.ButtonSet.OK);
    var nRoturas = ejecutarFaseD_IntegridadReferencial(true);

    // Informe final
    _lim_generarInformes_(logB, logC, nRoturas);

    var ss = getTargetSS_();
    var shFinal = ss.getSheetByName('LIMPIEZA_REALIZADA');
    if (shFinal) ss.setActiveSheet(shFinal);

    ui.alert(
      '✅ FASE 3 COMPLETADA',
      'Registros eliminados (prueba): ' + logB.eliminados.length + '\n'
      + 'Duplicados exactos eliminados: ' + logC.eliminados.length + '\n'
      + 'Enviados a REVISION_MANUAL: ' + (logB.revisionManual.length + logC.revisionManual.length + nRoturas) + '\n\n'
      + 'Ver: LIMPIEZA_REALIZADA y REVISION_MANUAL',
      ui.ButtonSet.OK
    );

  } catch (e) {
    logEvent('ERROR', 'ejecutarFase3Completa', e.message, e.stack);
    ui.alert('❌ ERROR en Fase 3:\n' + e.message + '\n\nRevisa el backup y la hoja LOGS.');
  }
}

// ============================================================
// FASE B — LIMPIEZA DE DATOS DE PRUEBA
// ============================================================

function ejecutarFaseB_LimpiezaPrueba(silencioso) {
  var ui  = silencioso ? null : SpreadsheetApp.getUi();
  var ss  = getTargetSS_();
  var log = _lim_nuevaLog_();

  if (!silencioso) {
    var conf = ui.alert('FASE B — Limpieza de prueba',
      'Solo se eliminarán registros claramente clasificados como PRUEBA.\n'
      + 'Se verificará que no tienen actividad real asociada.\n\n¿Continuar?',
      ui.ButtonSet.YES_NO);
    if (conf !== ui.Button.YES) { ui.alert('Cancelado.'); return log; }
  }

  // Verificar backup previo
  var prefijoBkp = _bkp_verificarExisteBackup_(ss);
  if (!prefijoBkp) {
    if (!silencioso) ui.alert('⚠️ No se encontró backup reciente. Ejecuta el backup primero.');
    throw new Error('FASE B: No existe backup previo. Operación cancelada.');
  }

  // Construir índice de IDs de clientes de prueba (para verificar en otras hojas)
  var idsPrueba = _lim_identificarClientesPrueba_(ss, log);

  // Limpiar por hoja
  _lim_limpiarTrabajosPrueba_(ss, log, idsPrueba);
  _lim_limpiarFacturasPrueba_(ss, log, idsPrueba);
  _lim_limpiarProformasPrueba_(ss, log, idsPrueba);
  _lim_limpiarPresupuestosPrueba_(ss, log);
  _lim_limpiarGastosPrueba_(ss, log);
  _lim_limpiarClientesPrueba_(ss, log, idsPrueba); // clientes al final

  SpreadsheetApp.flush();
  logEvent('INFO', 'ejecutarFaseB_LimpiezaPrueba',
    'Eliminados: ' + log.eliminados.length + ' | RevisionManual: ' + log.revisionManual.length, '');

  if (!silencioso) {
    _lim_generarInformes_(log, _lim_nuevaLog_(), 0);
    ui.alert('✅ Fase B completada.\nEliminados: ' + log.eliminados.length
      + '\nEnviados a revisión: ' + log.revisionManual.length);
  }

  return log;
}

function _lim_identificarClientesPrueba_(ss, log) {
  var idsPrueba = {};
  var sh = ss.getSheetByName('CLIENTES');
  if (!sh || sh.getLastRow() < 2) return idsPrueba;

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues();
  data.forEach(function(row) {
    var id      = String(row[0] || '').trim();
    var nombre  = (String(row[2] || '') + ' ' + String(row[3] || '')).trim();
    var empresa = String(row[4] || '').trim();
    var email   = String(row[5] || '').trim();
    var cif     = String(row[8] || '').trim();
    var dir     = String(row[7] || '').trim();

    if (id && _lim_esPrueba_({ nombre: nombre + ' ' + empresa, email: email, cif: cif, dir: dir, id: id })) {
      idsPrueba[id] = { nombre: nombre, empresa: empresa, email: email };
    }
  });

  return idsPrueba;
}

function _lim_clienteTieneActividad_(ss, idCli) {
  var hojas = [
    { nombre: 'TRABAJOS',     col: 3  },
    { nombre: 'FACTURAS',     col: 5  },
    { nombre: 'PROFORMAS',    col: 5  }
  ];
  for (var i = 0; i < hojas.length; i++) {
    var sh = ss.getSheetByName(hojas[i].nombre);
    if (!sh || sh.getLastRow() < 2) continue;
    var vals = sh.getRange(2, hojas[i].col, sh.getLastRow() - 1, 1).getValues();
    for (var r = 0; r < vals.length; r++) {
      if (String(vals[r][0] || '').trim() === idCli) return true;
    }
  }
  return false;
}

function _lim_limpiarClientesPrueba_(ss, log, idsPrueba) {
  var sh = ss.getSheetByName('CLIENTES');
  if (!sh || sh.getLastRow() < 2) return;

  var filasAEliminar = [];
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues();

  data.forEach(function(row, idx) {
    var fila = idx + 2;
    var id   = String(row[0] || '').trim();
    if (!idsPrueba[id]) return; // no es de prueba

    // Verificar que no tiene actividad real antes de eliminar
    if (_lim_clienteTieneActividad_(ss, id)) {
      _lim_registrarRevision_(ss, 'CLIENTES', fila, id,
        'CLIENTE_PRUEBA_CON_ACTIVIDAD',
        'Cliente clasificado como prueba pero tiene registros en TRABAJOS/FACTURAS/PROFORMAS.',
        'Revisar manualmente qué registros son reales antes de eliminar el cliente.');
      log.revisionManual.push({ hoja: 'CLIENTES', fila: fila, id: id });
    } else {
      filasAEliminar.push(fila);
    }
  });

  // Eliminar de abajo a arriba
  filasAEliminar.sort(function(a, b) { return b - a; });
  filasAEliminar.forEach(function(fila) {
    var id = String(sh.getRange(fila, 1).getValue()).trim();
    sh.deleteRow(fila);
    _lim_registrarEliminado_(log, 'CLIENTES', fila, id, 'Cliente de prueba sin actividad asociada');
  });
}

function _lim_limpiarTrabajosPrueba_(ss, log, idsPrueba) {
  var sh = ss.getSheetByName('TRABAJOS');
  if (!sh || sh.getLastRow() < 2) return;

  var nCols = Math.max(sh.getLastColumn(), 19);
  var data  = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();
  var filasAEliminar = [];

  data.forEach(function(row, idx) {
    var fila    = idx + 2;
    var id      = String(row[0] || '').trim();
    var idCli   = String(row[2] || '').trim();
    var nomCli  = String(row[3] || '').trim();
    var idFac   = String(row[18] || '').trim(); // ID_Factura

    var esPrueba = idsPrueba[idCli] || _lim_esPrueba_({ nombre: nomCli, id: id });
    if (!esPrueba) return;

    // NUNCA eliminar si tiene factura asociada
    if (idFac && idFac.indexOf('FAC-') === 0) {
      _lim_registrarRevision_(ss, 'TRABAJOS', fila, id,
        'TRABAJO_PRUEBA_CON_FACTURA',
        'Trabajo de prueba con ID_Factura: ' + idFac + '. No se elimina automáticamente.',
        'Verificar si la factura es también de prueba y eliminar ambos manualmente.');
      log.revisionManual.push({ hoja: 'TRABAJOS', fila: fila, id: id });
      return;
    }

    filasAEliminar.push(fila);
  });

  filasAEliminar.sort(function(a, b) { return b - a; });
  filasAEliminar.forEach(function(fila) {
    var id = String(sh.getRange(fila, 1).getValue()).trim();
    sh.deleteRow(fila);
    _lim_registrarEliminado_(log, 'TRABAJOS', fila, id, 'Trabajo de cliente de prueba sin factura');
  });
}

function _lim_limpiarFacturasPrueba_(ss, log, idsPrueba) {
  var sh = ss.getSheetByName('FACTURAS');
  if (!sh || sh.getLastRow() < 2) return;

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), 17)).getValues();
  var filasAEliminar = [];

  data.forEach(function(row, idx) {
    var fila   = idx + 2;
    var id     = String(row[0] || '').trim();
    var idCli  = String(row[4] || '').trim();
    var nomCli = String(row[5] || '').trim();
    var total  = row[16];

    // Solo eliminar si el cliente es de prueba Y el total es cero o la factura no tiene número real
    var esPrueba = idsPrueba[idCli] || _lim_esPrueba_({ nombre: nomCli, id: id });
    if (!esPrueba) return;

    // Si el total tiene valor significativo, enviar a revisión
    var totalNum = parseFloat(total) || 0;
    if (totalNum > 0) {
      _lim_registrarRevision_(ss, 'FACTURAS', fila, id,
        'FACTURA_PRUEBA_CON_IMPORTE',
        'Factura de prueba con total ' + totalNum.toFixed(2) + '€. Verificar antes de eliminar.',
        'Si el cliente y el importe son de prueba, eliminar manualmente.');
      log.revisionManual.push({ hoja: 'FACTURAS', fila: fila, id: id });
      return;
    }

    filasAEliminar.push(fila);
  });

  filasAEliminar.sort(function(a, b) { return b - a; });
  filasAEliminar.forEach(function(fila) {
    var id = String(sh.getRange(fila, 1).getValue()).trim();
    sh.deleteRow(fila);
    _lim_registrarEliminado_(log, 'FACTURAS', fila, id, 'Factura de prueba con importe cero');
  });
}

function _lim_limpiarProformasPrueba_(ss, log, idsPrueba) {
  var sh = ss.getSheetByName('PROFORMAS');
  if (!sh || sh.getLastRow() < 2) return;

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), 13)).getValues();
  var filasAEliminar = [];

  data.forEach(function(row, idx) {
    var fila   = idx + 2;
    var id     = String(row[0] || '').trim();
    var idCli  = String(row[4] || '').trim();
    var nomCli = String(row[5] || '').trim();

    if (idsPrueba[idCli] || _lim_esPrueba_({ nombre: nomCli, id: id })) {
      filasAEliminar.push(fila);
    }
  });

  filasAEliminar.sort(function(a, b) { return b - a; });
  filasAEliminar.forEach(function(fila) {
    var id = String(sh.getRange(fila, 1).getValue()).trim();
    sh.deleteRow(fila);
    _lim_registrarEliminado_(log, 'PROFORMAS', fila, id, 'Proforma de cliente de prueba');
  });
}

function _lim_limpiarPresupuestosPrueba_(ss, log) {
  var sh = ss.getSheetByName('PRESUPUESTOS');
  if (!sh || sh.getLastRow() < 2) return;

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), 9)).getValues();
  var filasAEliminar = [];

  data.forEach(function(row, idx) {
    var fila    = idx + 2;
    var id      = String(row[0] || '').trim();
    var cliente = String(row[2] || '').trim();

    if (_lim_esPrueba_({ nombre: cliente, id: id })) {
      filasAEliminar.push(fila);
    }
  });

  filasAEliminar.sort(function(a, b) { return b - a; });
  filasAEliminar.forEach(function(fila) {
    var id = String(sh.getRange(fila, 1).getValue()).trim();
    sh.deleteRow(fila);
    _lim_registrarEliminado_(log, 'PRESUPUESTOS', fila, id, 'Presupuesto de cliente de prueba');
  });
}

function _lim_limpiarGastosPrueba_(ss, log) {
  var sh = ss.getSheetByName('GASTOS');
  if (!sh || sh.getLastRow() < 2) return;

  var nCols = Math.max(sh.getLastColumn(), 25);
  var data  = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();
  var filasAEliminar = [];

  data.forEach(function(row, idx) {
    var fila      = idx + 2;
    var id        = String(row[COL_GAS.ID_GASTO   - 1] || '').trim();
    var proveedor = String(row[COL_GAS.PROVEEDOR  - 1] || '').trim();
    var concepto  = String(row[COL_GAS.CONCEPTO   - 1] || '').trim();
    var total     = parseFloat(row[COL_GAS.TOTAL  - 1]) || 0;

    if (_lim_esPrueba_({ nombre: proveedor + ' ' + concepto, id: id }) && total === 0) {
      filasAEliminar.push(fila);
    } else if (_lim_esPrueba_({ nombre: proveedor + ' ' + concepto, id: id }) && total > 0) {
      _lim_registrarRevision_(ss, 'GASTOS', fila, id,
        'GASTO_PRUEBA_CON_IMPORTE',
        'Gasto de prueba con total ' + total.toFixed(2) + '€.',
        'Verificar y eliminar manualmente si es un gasto de ejemplo.');
      log.revisionManual.push({ hoja: 'GASTOS', fila: fila, id: id });
    }
  });

  filasAEliminar.sort(function(a, b) { return b - a; });
  filasAEliminar.forEach(function(fila) {
    var id = String(sh.getRange(fila, 1).getValue()).trim();
    sh.deleteRow(fila);
    _lim_registrarEliminado_(log, 'GASTOS', fila, id, 'Gasto de prueba con importe cero');
  });
}

// ============================================================
// FASE C — CORRECCIÓN DE DUPLICADOS EXACTOS
// ============================================================

function ejecutarFaseC_CorreccionDuplicados(silencioso) {
  var ui  = silencioso ? null : SpreadsheetApp.getUi();
  var ss  = getTargetSS_();
  var log = _lim_nuevaLog_();

  if (!silencioso) {
    var prefijoBkp = _bkp_verificarExisteBackup_(ss);
    if (!prefijoBkp) { ui.alert('⚠️ No hay backup previo. Ejecuta el backup primero.'); return log; }

    var conf = ui.alert('FASE C — Duplicados',
      'Se corregirán duplicados EXACTOS (filas idénticas) conservando el más antiguo.\n'
      + 'Duplicados con diferencias o con factura → REVISION_MANUAL.\n\n¿Continuar?',
      ui.ButtonSet.YES_NO);
    if (conf !== ui.Button.YES) { ui.alert('Cancelado.'); return log; }
  }

  _lim_corregirDuplicadosEnHoja_(ss, 'TRABAJOS',     19, log);
  _lim_corregirDuplicadosEnHoja_(ss, 'CLIENTES',     11, log);
  _lim_corregirDuplicadosEnHoja_(ss, 'FACTURAS',     23, log);
  _lim_corregirDuplicadosEnHoja_(ss, 'PROFORMAS',    21, log);
  _lim_corregirDuplicadosEnHoja_(ss, 'PRESUPUESTOS', 13, log);

  SpreadsheetApp.flush();
  logEvent('INFO', 'ejecutarFaseC_CorreccionDuplicados',
    'Eliminados: ' + log.eliminados.length + ' | RevisionManual: ' + log.revisionManual.length, '');

  if (!silencioso) {
    _lim_generarInformes_(_lim_nuevaLog_(), log, 0);
    ui.alert('✅ Fase C completada.\nDuplicados exactos eliminados: ' + log.eliminados.length
      + '\nEnviados a revisión: ' + log.revisionManual.length);
  }

  return log;
}

function _lim_corregirDuplicadosEnHoja_(ss, nombreHoja, minCols, log) {
  var sh = ss.getSheetByName(nombreHoja);
  if (!sh || sh.getLastRow() < 2) return;

  var nCols = Math.max(sh.getLastColumn(), minCols);
  var data  = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();

  // Agrupar filas por ID (columna 1)
  var grupos = {};
  data.forEach(function(row, idx) {
    var id = String(row[0] || '').trim();
    if (!id) return;
    if (!grupos[id]) grupos[id] = [];
    grupos[id].push({ fila: idx + 2, row: row });
  });

  var filasAEliminar = [];

  Object.keys(grupos).forEach(function(id) {
    var grupo = grupos[id];
    if (grupo.length < 2) return;

    // Verificar si hay ID_Factura asociado (TRABAJOS: col 19)
    var tieneFactura = false;
    if (nombreHoja === 'TRABAJOS') {
      tieneFactura = grupo.some(function(g) {
        return String(g.row[18] || '').trim().indexOf('FAC-') === 0;
      });
    }

    if (tieneFactura) {
      grupo.forEach(function(g) {
        _lim_registrarRevision_(ss, nombreHoja, g.fila, id,
          'DUPLICADO_CON_FACTURA',
          'ID duplicado y alguna fila tiene ID_Factura asociado. Riesgo crítico.',
          'Revisar manualmente cuál fila es la correcta. Conservar la que tenga la factura real.');
        log.revisionManual.push({ hoja: nombreHoja, fila: g.fila, id: id });
      });
      return;
    }

    // Comprobar si son 100% idénticas (ignorando fila 0 = ID que ya es igual)
    var primerHash = grupo[0].row.join('|||');
    var todasIguales = grupo.every(function(g) { return g.row.join('|||') === primerHash; });

    if (todasIguales) {
      // Conservar la fila más antigua (número de fila más bajo = la primera)
      var filasOrdenadas = grupo.map(function(g) { return g.fila; }).sort(function(a,b){return a-b;});
      var filaAConservar = filasOrdenadas[0];
      var filasABorrar   = filasOrdenadas.slice(1);

      filasABorrar.forEach(function(f) {
        filasAEliminar.push(f);
        _lim_registrarEliminado_(log, nombreHoja, f, id,
          'Duplicado exacto. Se conserva la fila más antigua: ' + filaAConservar);
      });

    } else {
      // Diferencias → REVISION_MANUAL
      grupo.forEach(function(g) {
        _lim_registrarRevision_(ss, nombreHoja, g.fila, id,
          'DUPLICADO_CON_DIFERENCIAS',
          'ID duplicado con diferencias entre filas. Requiere revisión humana.',
          'Comparar las filas ' + grupo.map(function(x){return x.fila;}).join(', ')
          + ' y eliminar la incorrecta manualmente.');
        log.revisionManual.push({ hoja: nombreHoja, fila: g.fila, id: id });
      });
    }
  });

  // Eliminar de abajo a arriba
  filasAEliminar.sort(function(a, b) { return b - a; });
  filasAEliminar.forEach(function(fila) {
    var id = String(sh.getRange(fila, 1).getValue()).trim();
    sh.deleteRow(fila);
  });
}

// ============================================================
// FASE D — INTEGRIDAD REFERENCIAL
// ============================================================

function ejecutarFaseD_IntegridadReferencial(silencioso) {
  var ui = silencioso ? null : SpreadsheetApp.getUi();
  var ss = getTargetSS_();
  var nRoturas = 0;

  if (!silencioso) {
    var conf = ui.alert('FASE D — Integridad referencial',
      'Se detectarán referencias rotas entre tablas.\n'
      + 'Los problemas se registran en REVISION_MANUAL.\nNada se elimina automáticamente.\n\n¿Continuar?',
      ui.ButtonSet.YES_NO);
    if (conf !== ui.Button.YES) { ui.alert('Cancelado.'); return 0; }
  }

  // Construir índice de clientes válidos
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
    var dataTra = shTra.getRange(2, 1, shTra.getLastRow() - 1, Math.max(shTra.getLastColumn(), 4)).getValues();
    dataTra.forEach(function(row, idx) {
      var fila  = idx + 2;
      var id    = String(row[0] || '').trim();
      var idCli = String(row[2] || '').trim();
      var nom   = String(row[3] || '').trim();
      if (idCli && !idsClientes[idCli]) {
        _lim_registrarRevision_(ss, 'TRABAJOS', fila, id,
          'HUERFANO_CLIENTE',
          'ID_Cliente "' + idCli + '" (' + nom + ') no existe en CLIENTES.',
          'Asignar el cliente correcto en la columna C o eliminar el registro si es de prueba.');
        nRoturas++;
      }
    });
  }

  // FACTURAS → CLIENTES
  var shFac = ss.getSheetByName('FACTURAS');
  if (shFac && shFac.getLastRow() >= 2) {
    var dataFac = shFac.getRange(2, 1, shFac.getLastRow() - 1, Math.max(shFac.getLastColumn(), 6)).getValues();
    dataFac.forEach(function(row, idx) {
      var fila  = idx + 2;
      var id    = String(row[0] || '').trim();
      var idCli = String(row[4] || '').trim();
      var nom   = String(row[5] || '').trim();
      if (idCli && !idsClientes[idCli]) {
        _lim_registrarRevision_(ss, 'FACTURAS', fila, id,
          'HUERFANO_CLIENTE',
          'ID_Cliente "' + idCli + '" (' + nom + ') no existe en CLIENTES. RIESGO: factura sin cliente válido.',
          'Restaurar el cliente desde el backup o asignar el ID correcto.');
        nRoturas++;
      }
    });
  }

  // PROFORMAS → CLIENTES
  var shPro = ss.getSheetByName('PROFORMAS');
  if (shPro && shPro.getLastRow() >= 2) {
    var dataPro = shPro.getRange(2, 1, shPro.getLastRow() - 1, Math.max(shPro.getLastColumn(), 6)).getValues();
    dataPro.forEach(function(row, idx) {
      var fila  = idx + 2;
      var id    = String(row[0] || '').trim();
      var idCli = String(row[4] || '').trim();
      var nom   = String(row[5] || '').trim();
      if (idCli && !idsClientes[idCli]) {
        _lim_registrarRevision_(ss, 'PROFORMAS', fila, id,
          'HUERFANO_CLIENTE',
          'ID_Cliente "' + idCli + '" (' + nom + ') no existe en CLIENTES.',
          'Asignar el cliente correcto o eliminar si es de prueba.');
        nRoturas++;
      }
    });
  }

  // PRESUPUESTOS → CLIENTES (por nombre)
  var nombresClientes = {};
  if (shCli && shCli.getLastRow() >= 2) {
    shCli.getRange(2, 1, shCli.getLastRow() - 1, 5).getValues().forEach(function(row) {
      var nom = (String(row[2] || '') + ' ' + String(row[3] || '')).trim().toLowerCase();
      var emp = String(row[4] || '').trim().toLowerCase();
      if (nom && nom !== ' ') nombresClientes[nom] = true;
      if (emp) nombresClientes[emp] = true;
    });
  }
  var shPre = ss.getSheetByName('PRESUPUESTOS');
  if (shPre && shPre.getLastRow() >= 2) {
    var dataPre = shPre.getRange(2, 1, shPre.getLastRow() - 1, Math.max(shPre.getLastColumn(), 3)).getValues();
    dataPre.forEach(function(row, idx) {
      var fila    = idx + 2;
      var id      = String(row[0] || '').trim();
      var cliente = String(row[2] || '').trim();
      if (cliente && !nombresClientes[cliente.toLowerCase()]) {
        _lim_registrarRevision_(ss, 'PRESUPUESTOS', fila, id,
          'CLIENTE_NO_EN_BBDD',
          'Cliente "' + cliente + '" no aparece en CLIENTES (se busca por nombre).',
          'Verificar si el cliente existe con otro nombre o fue eliminado.');
        nRoturas++;
      }
    });
  }

  logEvent('INFO', 'ejecutarFaseD_IntegridadReferencial', nRoturas + ' roturas detectadas', '');

  if (!silencioso) {
    ui.alert('✅ Fase D completada. '
      + nRoturas + ' problemas de integridad detectados y registrados en REVISION_MANUAL.');
  }

  return nRoturas;
}

// ============================================================
// INFORMES
// ============================================================

function _lim_generarInformes_(logB, logC, nRoturas) {
  var ss = getTargetSS_();
  var tz = Session.getScriptTimeZone();
  var ahora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');

  // Hoja LIMPIEZA_REALIZADA
  var sh = ss.getSheetByName('LIMPIEZA_REALIZADA');
  if (sh) { sh.clear(); sh.clearFormats(); }
  else    { sh = ss.insertSheet('LIMPIEZA_REALIZADA'); }

  sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 70);
  sh.setColumnWidth(3, 130); sh.setColumnWidth(4, 400);
  sh.setColumnWidth(5, 180);

  var r = 1;

  sh.getRange(r, 1, 1, 5).merge()
    .setValue('🧹 LIMPIEZA REALIZADA — ' + ahora)
    .setBackground('#137333').setFontColor('#ffffff')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(r, 45); r++;

  // Resumen numérico
  sh.getRange(r, 1, 1, 5).merge()
    .setValue('Registros eliminados (prueba): ' + logB.eliminados.length
      + '   |   Duplicados exactos eliminados: ' + logC.eliminados.length
      + '   |   Enviados a REVISION_MANUAL: '
      + (logB.revisionManual.length + logC.revisionManual.length + nRoturas))
    .setBackground('#e6f4ea').setFontColor('#137333').setHorizontalAlignment('center');
  r += 2;

  // Tabla de eliminados
  if (logB.eliminados.length + logC.eliminados.length > 0) {
    sh.getRange(r, 1, 1, 5).merge()
      .setValue('✅ REGISTROS ELIMINADOS')
      .setBackground('#202124').setFontColor('#ffffff').setFontWeight('bold');
    r++;
    sh.getRange(r, 1, 1, 5).setValues([['Hoja','Fila','ID','Motivo','Timestamp']])
      .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
    r++;

    logB.eliminados.concat(logC.eliminados).forEach(function(item, i) {
      var bg = i % 2 === 0 ? '#f0fff4' : '#e8f5e9';
      sh.getRange(r, 1, 1, 5).setBackground(bg);
      sh.getRange(r, 1).setValue(item.hoja);
      sh.getRange(r, 2).setValue(item.fila).setHorizontalAlignment('center');
      sh.getRange(r, 3).setValue(item.id);
      sh.getRange(r, 4).setValue(item.motivo).setWrap(true).setFontSize(9);
      sh.getRange(r, 5).setValue(
        item.ts instanceof Date ? Utilities.formatDate(item.ts, tz, 'dd/MM/yyyy HH:mm:ss') : ''
      );
      sh.setRowHeight(r, 30); r++;
    });
    r++;
  }

  // Nota de reversión
  sh.getRange(r, 1, 1, 5).merge()
    .setValue('📋 Para revertir: usar las hojas de backup con prefijo BKP_*. '
      + 'Copiar los datos del backup a la hoja original.')
    .setBackground('#fff3cd').setFontColor('#856404').setWrap(true).setFontSize(10);

  SpreadsheetApp.flush();
}

// ---- Wrapper público para el menú ----

function ejecutarFaseB_solo() { ejecutarFaseB_LimpiezaPrueba(false); }
function ejecutarFaseC_solo() { ejecutarFaseC_CorreccionDuplicados(false); }
function ejecutarFaseD_solo() { ejecutarFaseD_IntegridadReferencial(false); }
