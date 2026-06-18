// ============================================================
// Auditoria.gs — Auditoría completa de datos reales vs. prueba
// NO ELIMINA NI MODIFICA NINGÚN DATO
// ============================================================

// ---- Palabras clave de datos de prueba ----
var TEST_KEYWORDS = [
  'test', 'prueba', 'ejemplo', 'demo', 'sample', 'ficticio', 'ficticia',
  'falso', 'falsa', 'dummy', 'fake', 'temporal', 'borrar', 'eliminar',
  'mock', 'placeholder', 'xxx', 'zzz', 'aaa', 'cliente prueba',
  'cliente test', 'test client', 'cliente de prueba'
];

var TEST_EMAILS = [
  '@test.com', '@example.com', '@prueba.com', '@demo.com',
  '@fake.com', '@dummy.com', '@temp.com', '@correo.com',
  'test@', 'prueba@', 'ejemplo@', 'demo@', 'noreply@',
  'no-reply@', 'info@info', 'admin@admin'
];

var TEST_CIFS = [
  'A00000000', 'B00000000', '00000000A', '00000000Z',
  'B12345678', 'A12345678', '12345678A', '11111111H',
  '00000000', '12345678', '99999999'
];

var TEST_ADDRESSES = [
  'Calle Ejemplo', 'Calle Test', 'Calle Prueba', 'Calle Demo',
  'C/ Ejemplo', 'C/ Test', 'dirección de prueba', 'sin dirección'
];

var EMPRESA_NOMBRE_DEFAULT = 'Mi Empresa';
var EMPRESA_CIF_DEFAULT    = 'A00000000';

// ---- Punto de entrada principal ----

function ejecutarAuditoriaDatos() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    '🔍 AUDITORÍA EN CURSO',
    'Analizando todas las hojas...\nEsto puede tardar 30-60 segundos.\nNo cierres la hoja.',
    ui.ButtonSet.OK
  );

  try {
    var resultado = _auditarTodo_();
    _generarHojaAuditoria_(resultado);

    var ss = getTargetSS_();
    var sh = ss.getSheetByName('AUDITORIA_DATOS');
    if (sh) ss.setActiveSheet(sh);

    ui.alert(
      '✅ AUDITORÍA COMPLETADA',
      'Informe generado en la hoja: AUDITORIA_DATOS\n\n'
      + 'Hojas analizadas: ' + resultado.resumen.hojasAnalizadas + '\n'
      + 'Total registros: ' + resultado.resumen.totalRegistros + '\n'
      + 'Registros de prueba detectados: ' + resultado.resumen.totalPrueba + '\n'
      + 'Duplicados encontrados: ' + resultado.resumen.totalDuplicados + '\n'
      + 'Huérfanos detectados: ' + resultado.resumen.totalHuerfanos,
      ui.ButtonSet.OK
    );
  } catch (e) {
    logEvent('ERROR', 'ejecutarAuditoriaDatos', e.message, e.stack);
    ui.alert('❌ ERROR en auditoría:\n' + e.message + '\n\n' + e.stack);
  }
}

// ---- Motor de auditoría ----

function _auditarTodo_() {
  var ss = getTargetSS_();

  var res = {
    resumen: {
      hojasAnalizadas: 0,
      totalRegistros: 0,
      totalReales: 0,
      totalPrueba: 0,
      totalDudosos: 0,
      totalDuplicados: 0,
      totalHuerfanos: 0
    },
    clientes:       _auditarClientes_(ss),
    trabajos:       _auditarTrabajos_(ss),
    facturas:       _auditarFacturas_(ss),
    proformas:      _auditarProformas_(ss),
    presupuestos:   _auditarPresupuestos_(ss),
    gastos:         _auditarGastos_(ss),
    informes:       _auditarInformes_(ss),
    recLog:         _auditarRecordatoriosLog_(ss),
    alertasLog:     _auditarAlertasLog_(ss),
    integridad:     {},
    duplicadosTRA:  _analizarDuplicadosTRA_(ss),
    fase1:          [],
    fase2:          [],
    fase3:          []
  };

  // Integridad referencial
  res.integridad = _verificarIntegridad_(ss, res);

  // Totales globales
  var hojas = [res.clientes, res.trabajos, res.facturas, res.proformas,
               res.presupuestos, res.gastos, res.informes, res.recLog, res.alertasLog];
  hojas.forEach(function(h) {
    if (h && h.total !== undefined) {
      res.resumen.hojasAnalizadas++;
      res.resumen.totalRegistros  += h.total   || 0;
      res.resumen.totalReales     += h.reales  || 0;
      res.resumen.totalPrueba     += h.prueba  || 0;
      res.resumen.totalDudosos    += h.dudosos || 0;
    }
  });
  res.resumen.totalDuplicados = (res.clientes.duplicados || []).length
    + (res.trabajos.duplicados || []).length
    + (res.facturas.duplicados || []).length
    + (res.proformas.duplicados || []).length
    + (res.presupuestos.duplicados || []).length
    + (res.duplicadosTRA.duplicadosExactos || []).length;
  res.resumen.totalHuerfanos = (res.integridad.huerfanos || []).length;

  // Fases de recomendación
  res.fase1 = _calcularFase1_(res);
  res.fase2 = _calcularFase2_(res);
  res.fase3 = _calcularFase3_(res);

  return res;
}

// ---- Auditoría por hoja ----

function _auditarClientes_(ss) {
  var r = { nombre: 'CLIENTES', total: 0, reales: 0, prueba: 0, dudosos: 0,
            duplicadosId: [], filasIdenticas: [], duplicados: [],
            registrosPrueba: [], registrosDudosos: [], registrosReales: [] };
  var sh = ss.getSheetByName('CLIENTES');
  if (!sh || sh.getLastRow() < 2) { r.error = 'Hoja vacía o no encontrada'; return r; }

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  r.total = data.length;

  var idCount = {};
  var filaHash = {};

  data.forEach(function(row, idx) {
    var fila    = idx + 2;
    var id      = String(row[0] || '').trim();
    var nombre  = String(row[2] || '').trim();
    var ape     = String(row[3] || '').trim();
    var empresa = String(row[4] || '').trim();
    var email   = String(row[5] || '').trim();
    var cif     = String(row[8] || '').trim();
    var dir     = String(row[7] || '').trim();

    // Duplicados de ID
    if (id) {
      idCount[id] = (idCount[id] || 0) + 1;
      if (idCount[id] === 2) r.duplicadosId.push(id);
    }

    // Filas idénticas (hash simplificado)
    var hash = [nombre, ape, empresa, email, cif].join('|').toLowerCase();
    if (hash !== '||||') {
      if (filaHash[hash]) {
        r.filasIdenticas.push({ filas: [filaHash[hash], fila], hash: hash });
      } else {
        filaHash[hash] = fila;
      }
    }

    // Clasificar
    var esPrueba = _esRegistroPrueba_({
      nombre: nombre + ' ' + ape + ' ' + empresa,
      email: email, cif: cif, dir: dir
    });

    var info = { fila: fila, id: id, nombre: nombre + ' ' + ape, empresa: empresa, email: email, cif: cif };

    if (esPrueba === 'PRUEBA') {
      r.prueba++;
      r.registrosPrueba.push(info);
    } else if (esPrueba === 'DUDOSO') {
      r.dudosos++;
      r.registrosDudosos.push(info);
    } else {
      r.reales++;
      r.registrosReales.push(info);
    }
  });

  // IDs con más de 1 aparición
  r.duplicadosId = Object.keys(idCount).filter(function(k) { return idCount[k] > 1; });
  r.duplicados = r.duplicadosId.concat(r.filasIdenticas.map(function(f) { return 'Filas ' + f.filas.join('+'); }));

  return r;
}

function _auditarTrabajos_(ss) {
  var r = { nombre: 'TRABAJOS', total: 0, reales: 0, prueba: 0, dudosos: 0,
            duplicadosId: [], duplicados: [], idsMultiples: {},
            registrosPrueba: [], registrosDudosos: [] };
  var sh = ss.getSheetByName('TRABAJOS');
  if (!sh || sh.getLastRow() < 2) { r.error = 'Hoja vacía o no encontrada'; return r; }

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  r.total = data.length;

  var idCount = {};

  data.forEach(function(row, idx) {
    var fila       = idx + 2;
    var id         = String(row[0]  || '').trim();
    var idCliente  = String(row[2]  || '').trim();
    var nomCliente = String(row[3]  || '').trim();
    var desc       = String(row[6]  || '').trim();
    var dir        = String(row[8]  || '').trim();

    if (id) {
      if (!idCount[id]) idCount[id] = [];
      idCount[id].push(fila);
    }

    var esPrueba = _esRegistroPrueba_({
      nombre: nomCliente + ' ' + desc, dir: dir, id: id
    });

    var info = { fila: fila, id: id, idCliente: idCliente, nomCliente: nomCliente, desc: desc };

    if (esPrueba === 'PRUEBA') {
      r.prueba++;
      r.registrosPrueba.push(info);
    } else if (esPrueba === 'DUDOSO') {
      r.dudosos++;
      r.registrosDudosos.push(info);
    } else {
      r.reales++;
    }
  });

  // IDs con múltiples filas
  Object.keys(idCount).forEach(function(id) {
    if (idCount[id].length > 1) {
      r.idsMultiples[id] = idCount[id];
      r.duplicadosId.push(id);
    }
  });
  r.duplicados = r.duplicadosId;

  return r;
}

function _auditarFacturas_(ss) {
  var r = { nombre: 'FACTURAS', total: 0, reales: 0, prueba: 0, dudosos: 0,
            duplicadosId: [], sinCliente: [], duplicados: [],
            registrosPrueba: [], registrosDudosos: [] };
  var sh = ss.getSheetByName('FACTURAS');
  if (!sh || sh.getLastRow() < 2) { r.error = 'Hoja vacía o no encontrada'; return r; }

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  r.total = data.length;

  var idCount  = {};
  var numCount = {};

  data.forEach(function(row, idx) {
    var fila       = idx + 2;
    var id         = String(row[0]  || '').trim();
    var serie      = String(row[2]  || '').trim();
    var numero     = String(row[3]  || '').trim();
    var idCliente  = String(row[4]  || '').trim();
    var nomCliente = String(row[5]  || '').trim();

    if (id) { idCount[id] = (idCount[id] || 0) + 1; }

    var claveNum = serie + '-' + numero;
    if (numero) { numCount[claveNum] = (numCount[claveNum] || 0) + 1; }

    if (!idCliente) r.sinCliente.push({ fila: fila, id: id });

    var esPrueba = _esRegistroPrueba_({ nombre: nomCliente, id: id });
    var info = { fila: fila, id: id, idCliente: idCliente, nomCliente: nomCliente, serie: serie, numero: numero };

    if (esPrueba === 'PRUEBA') { r.prueba++; r.registrosPrueba.push(info); }
    else if (esPrueba === 'DUDOSO') { r.dudosos++; r.registrosDudosos.push(info); }
    else { r.reales++; }
  });

  r.duplicadosId  = Object.keys(idCount).filter(function(k) { return idCount[k] > 1; });
  r.duplicadosNum = Object.keys(numCount).filter(function(k) { return numCount[k] > 1; });
  r.duplicados    = r.duplicadosId.concat(r.duplicadosNum);

  return r;
}

function _auditarProformas_(ss) {
  var r = { nombre: 'PROFORMAS', total: 0, reales: 0, prueba: 0, dudosos: 0,
            duplicadosId: [], sinCliente: [], duplicados: [],
            registrosPrueba: [], registrosDudosos: [] };
  var sh = ss.getSheetByName('PROFORMAS');
  if (!sh || sh.getLastRow() < 2) { r.error = 'Hoja vacía o no encontrada'; return r; }

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  r.total = data.length;

  var idCount = {};

  data.forEach(function(row, idx) {
    var fila       = idx + 2;
    var id         = String(row[0] || '').trim();
    var idCliente  = String(row[4] || '').trim();
    var nomCliente = String(row[5] || '').trim();

    if (id) { idCount[id] = (idCount[id] || 0) + 1; }
    if (!idCliente) r.sinCliente.push({ fila: fila, id: id });

    var esPrueba = _esRegistroPrueba_({ nombre: nomCliente, id: id });
    var info = { fila: fila, id: id, idCliente: idCliente, nomCliente: nomCliente };

    if (esPrueba === 'PRUEBA') { r.prueba++; r.registrosPrueba.push(info); }
    else if (esPrueba === 'DUDOSO') { r.dudosos++; r.registrosDudosos.push(info); }
    else { r.reales++; }
  });

  r.duplicadosId = Object.keys(idCount).filter(function(k) { return idCount[k] > 1; });
  r.duplicados   = r.duplicadosId;
  return r;
}

function _auditarPresupuestos_(ss) {
  var r = { nombre: 'PRESUPUESTOS', total: 0, reales: 0, prueba: 0, dudosos: 0,
            duplicadosId: [], duplicados: [],
            registrosPrueba: [], registrosDudosos: [] };
  var sh = ss.getSheetByName('PRESUPUESTOS');
  if (!sh || sh.getLastRow() < 2) { r.error = 'Hoja vacía o no encontrada'; return r; }

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  r.total = data.length;

  var idCount = {};

  data.forEach(function(row, idx) {
    var fila    = idx + 2;
    var id      = String(row[0] || '').trim();
    var cliente = String(row[2] || '').trim();

    if (id) { idCount[id] = (idCount[id] || 0) + 1; }

    var esPrueba = _esRegistroPrueba_({ nombre: cliente, id: id });
    var info = { fila: fila, id: id, cliente: cliente };

    if (esPrueba === 'PRUEBA') { r.prueba++; r.registrosPrueba.push(info); }
    else if (esPrueba === 'DUDOSO') { r.dudosos++; r.registrosDudosos.push(info); }
    else { r.reales++; }
  });

  r.duplicadosId = Object.keys(idCount).filter(function(k) { return idCount[k] > 1; });
  r.duplicados   = r.duplicadosId;
  return r;
}

function _auditarGastos_(ss) {
  var r = { nombre: 'GASTOS', total: 0, reales: 0, prueba: 0, dudosos: 0,
            sinArchivo: [], ocr: [], ejemplos: [],
            registrosPrueba: [], registrosDudosos: [] };
  var sh = ss.getSheetByName('GASTOS');
  if (!sh || sh.getLastRow() < 2) { r.error = 'Hoja vacía o no encontrada'; return r; }

  var nCols = Math.max(sh.getLastColumn(), 25);
  var data  = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();
  r.total   = data.length;

  data.forEach(function(row, idx) {
    var fila       = idx + 2;
    var id         = String(row[COL_GAS.ID_GASTO             - 1] || '').trim();
    var proveedor  = String(row[COL_GAS.PROVEEDOR            - 1] || '').trim();
    var concepto   = String(row[COL_GAS.CONCEPTO             - 1] || '').trim();
    var archivoId  = String(row[COL_GAS.ARCHIVO_ID           - 1] || '').trim();
    var ocrTexto   = String(row[COL_GAS.OCR_TEXTO            - 1] || '').trim();
    var estadoProc = String(row[COL_GAS.ESTADO_PROCESAMIENTO - 1] || '').trim();

    if (!archivoId) r.sinArchivo.push({ fila: fila, id: id, concepto: concepto });
    if (ocrTexto && ocrTexto.length > 5) r.ocr.push({ fila: fila, id: id });

    var esPrueba = _esRegistroPrueba_({ nombre: proveedor + ' ' + concepto, id: id });
    var info = { fila: fila, id: id, proveedor: proveedor, concepto: concepto, estadoProc: estadoProc };

    if (esPrueba === 'PRUEBA') { r.prueba++; r.registrosPrueba.push(info); r.ejemplos.push(info); }
    else if (esPrueba === 'DUDOSO') { r.dudosos++; r.registrosDudosos.push(info); }
    else { r.reales++; }
  });

  return r;
}

function _auditarInformes_(ss) {
  var r = { nombre: 'INFORMES', total: 0, reales: 0, prueba: 0, dudosos: 0,
            registrosPrueba: [], registrosDudosos: [] };
  var sh = ss.getSheetByName('INFORMES');
  if (!sh || sh.getLastRow() < 2) { r.error = 'Hoja vacía o no encontrada'; return r; }

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  r.total  = data.length;

  data.forEach(function(row, idx) {
    var fila    = idx + 2;
    var id      = String(row[0] || '').trim();
    var cliente = String(row[3] || '').trim();

    var esPrueba = _esRegistroPrueba_({ nombre: cliente, id: id });
    if (esPrueba === 'PRUEBA') { r.prueba++; r.registrosPrueba.push({ fila: fila, id: id, cliente: cliente }); }
    else if (esPrueba === 'DUDOSO') { r.dudosos++; }
    else { r.reales++; }
  });

  return r;
}

function _auditarRecordatoriosLog_(ss) {
  var r = { nombre: 'RECORDATORIOS_LOG', total: 0, reales: 0, prueba: 0, dudosos: 0 };
  var sh = ss.getSheetByName('RECORDATORIOS_LOG');
  if (!sh || sh.getLastRow() < 2) { r.error = 'Hoja vacía o no encontrada'; return r; }

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  r.total  = data.length;

  data.forEach(function(row) {
    var cliente = String(row[3] || '').trim();
    var email   = String(row[4] || '').trim();
    var esPrueba = _esRegistroPrueba_({ nombre: cliente, email: email });
    if (esPrueba === 'PRUEBA') r.prueba++;
    else if (esPrueba === 'DUDOSO') r.dudosos++;
    else r.reales++;
  });

  return r;
}

function _auditarAlertasLog_(ss) {
  var r = { nombre: 'ALERTAS_LOG', total: 0, reales: 0, prueba: 0, dudosos: 0 };
  var sh = ss.getSheetByName('ALERTAS_LOG');
  if (!sh || sh.getLastRow() < 2) { r.error = 'Hoja vacía o no encontrada'; return r; }

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  r.total  = data.length;
  r.reales = data.length; // Los logs se asumen reales salvo indicación
  return r;
}

// ---- Verificación de integridad referencial ----

function _verificarIntegridad_(ss, res) {
  var integ = { huerfanos: [], ok: [], problemas: [] };

  // Obtener IDs de clientes reales
  var shCli = ss.getSheetByName('CLIENTES');
  var idsClientes = {};
  if (shCli && shCli.getLastRow() >= 2) {
    shCli.getRange(2, 1, shCli.getLastRow() - 1, 1).getValues().forEach(function(row) {
      var id = String(row[0] || '').trim();
      if (id) idsClientes[id] = true;
    });
  }

  // Obtener IDs de trabajos
  var shTra = ss.getSheetByName('TRABAJOS');
  var idsTrabajosExistentes = {};
  if (shTra && shTra.getLastRow() >= 2) {
    shTra.getRange(2, 1, shTra.getLastRow() - 1, 1).getValues().forEach(function(row) {
      var id = String(row[0] || '').trim();
      if (id) idsTrabajosExistentes[id] = true;
    });
  }

  // TRABAJOS → CLIENTES
  if (shTra && shTra.getLastRow() >= 2) {
    var dataTra = shTra.getRange(2, 1, shTra.getLastRow() - 1, Math.max(shTra.getLastColumn(), 4)).getValues();
    dataTra.forEach(function(row, idx) {
      var fila      = idx + 2;
      var idTra     = String(row[0] || '').trim();
      var idCli     = String(row[2] || '').trim();
      var nomCli    = String(row[3] || '').trim();
      if (idCli && !idsClientes[idCli]) {
        integ.huerfanos.push({
          tipo: 'TRABAJO sin CLIENTE',
          fila: fila, id: idTra,
          referencia: idCli, nombreRef: nomCli
        });
      }
    });
  }

  // FACTURAS → CLIENTES
  var shFac = ss.getSheetByName('FACTURAS');
  if (shFac && shFac.getLastRow() >= 2) {
    var dataFac = shFac.getRange(2, 1, shFac.getLastRow() - 1, Math.max(shFac.getLastColumn(), 6)).getValues();
    dataFac.forEach(function(row, idx) {
      var fila   = idx + 2;
      var idFac  = String(row[0] || '').trim();
      var idCli  = String(row[4] || '').trim();
      var nomCli = String(row[5] || '').trim();
      if (idCli && !idsClientes[idCli]) {
        integ.huerfanos.push({
          tipo: 'FACTURA sin CLIENTE',
          fila: fila, id: idFac,
          referencia: idCli, nombreRef: nomCli
        });
      }
    });
  }

  // PROFORMAS → CLIENTES
  var shPro = ss.getSheetByName('PROFORMAS');
  if (shPro && shPro.getLastRow() >= 2) {
    var dataPro = shPro.getRange(2, 1, shPro.getLastRow() - 1, Math.max(shPro.getLastColumn(), 6)).getValues();
    dataPro.forEach(function(row, idx) {
      var fila   = idx + 2;
      var idPro  = String(row[0] || '').trim();
      var idCli  = String(row[4] || '').trim();
      var nomCli = String(row[5] || '').trim();
      if (idCli && !idsClientes[idCli]) {
        integ.huerfanos.push({
          tipo: 'PROFORMA sin CLIENTE',
          fila: fila, id: idPro,
          referencia: idCli, nombreRef: nomCli
        });
      }
    });
  }

  // PRESUPUESTOS → CLIENTES (por nombre de cliente, no ID)
  var shPre = ss.getSheetByName('PRESUPUESTOS');
  if (shPre && shPre.getLastRow() >= 2) {
    var nombresClientes = {};
    if (shCli && shCli.getLastRow() >= 2) {
      shCli.getRange(2, 1, shCli.getLastRow() - 1, 5).getValues().forEach(function(row) {
        var nom = (String(row[2] || '') + ' ' + String(row[3] || '')).trim().toLowerCase();
        var emp = String(row[4] || '').trim().toLowerCase();
        if (nom && nom !== ' ') nombresClientes[nom] = true;
        if (emp) nombresClientes[emp] = true;
      });
    }
    var dataPre = shPre.getRange(2, 1, shPre.getLastRow() - 1, Math.max(shPre.getLastColumn(), 3)).getValues();
    dataPre.forEach(function(row, idx) {
      var fila    = idx + 2;
      var idPre   = String(row[0] || '').trim();
      var cliente = String(row[2] || '').trim();
      if (cliente && !nombresClientes[cliente.toLowerCase()]) {
        integ.huerfanos.push({
          tipo: 'PRESUPUESTO cliente no en CLIENTES',
          fila: fila, id: idPre,
          referencia: '(nombre)', nombreRef: cliente
        });
      }
    });
  }

  // RECORDATORIOS_LOG → FACTURAS
  var shRec = ss.getSheetByName('RECORDATORIOS_LOG');
  if (shRec && shRec.getLastRow() >= 2) {
    var idsFacturas = {};
    if (shFac && shFac.getLastRow() >= 2) {
      shFac.getRange(2, 1, shFac.getLastRow() - 1, 1).getValues().forEach(function(row) {
        var id = String(row[0] || '').trim();
        if (id) idsFacturas[id] = true;
      });
    }
    var dataRec = shRec.getRange(2, 1, shRec.getLastRow() - 1, Math.max(shRec.getLastColumn(), 3)).getValues();
    dataRec.forEach(function(row, idx) {
      var fila  = idx + 2;
      var idRec = String(row[0] || '').trim();
      var idFac = String(row[2] || '').trim();
      if (idFac && !idsFacturas[idFac]) {
        integ.huerfanos.push({
          tipo: 'RECORDATORIO sin FACTURA',
          fila: fila, id: idRec,
          referencia: idFac, nombreRef: ''
        });
      }
    });
  }

  return integ;
}

// ---- Análisis de IDs duplicados en TRABAJOS ----

function _analizarDuplicadosTRA_(ss) {
  var r = {
    idsMultiples: {},
    duplicadosExactos: [],
    causasProbables: [],
    patronDetectado: ''
  };

  var sh = ss.getSheetByName('TRABAJOS');
  if (!sh || sh.getLastRow() < 2) return r;

  var nCols = Math.max(sh.getLastColumn(), 19);
  var data  = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();

  // Agrupar por ID
  var grupos = {};
  data.forEach(function(row, idx) {
    var id = String(row[0] || '').trim();
    if (!id) return;
    if (!grupos[id]) grupos[id] = [];
    grupos[id].push({ fila: idx + 2, row: row });
  });

  // Encontrar IDs con múltiples filas
  Object.keys(grupos).forEach(function(id) {
    if (grupos[id].length < 2) return;
    r.idsMultiples[id] = grupos[id].map(function(g) { return g.fila; });

    var grupo = grupos[id];

    // ¿Son filas idénticas?
    var primeraFila = grupo[0].row.join('|');
    var todasIguales = grupo.every(function(g) { return g.row.join('|') === primeraFila; });

    if (todasIguales) {
      r.duplicadosExactos.push({
        id: id,
        filas: grupo.map(function(g) { return g.fila; }),
        causa: 'Filas 100% idénticas → probable doble escritura del script o doble envío de formulario'
      });
    } else {
      // Analizar diferencias
      var primerTimestamp = grupo[0].row[1];
      var timestamps = grupo.map(function(g) { return g.row[1]; });
      var todosMismoTimestamp = timestamps.every(function(t) {
        return String(t) === String(primerTimestamp);
      });

      var causa = 'Filas con diferencias — ';
      if (todosMismoTimestamp) {
        causa += 'mismo timestamp de creación → posible importación duplicada o ejecución simultánea';
      } else {
        var diffMs = Math.abs(
          (timestamps[1] instanceof Date ? timestamps[1].getTime() : 0) -
          (timestamps[0] instanceof Date ? timestamps[0].getTime() : 0)
        );
        if (diffMs < 5000) {
          causa += 'timestamps a menos de 5s de diferencia → probable doble ejecución del formulario';
        } else if (diffMs < 60000) {
          causa += 'timestamps a menos de 60s de diferencia → posible doble escritura del script';
        } else {
          causa += 'timestamps muy separados → posible error de sincronización o importación';
        }
      }

      r.duplicadosExactos.push({
        id: id,
        filas: grupo.map(function(g) { return g.fila; }),
        causa: causa
      });
    }
  });

  // Diagnóstico del patrón
  var nDups = r.duplicadosExactos.length;
  if (nDups > 0) {
    var idsDups = r.duplicadosExactos.map(function(d) { return d.id; });
    var sonConsecutivos = _sonIdsConsecutivos_(idsDups);

    if (sonConsecutivos) {
      r.patronDetectado = 'Los IDs duplicados son consecutivos (' + idsDups.slice(0,5).join(', ') + '...) '
        + '→ sugiere una importación en bloque o ejecución repetida de la función de instalación/migración.';
      r.causasProbables = [
        '1. Se ejecutó una función de instalación/migración dos veces seguidas',
        '2. Se importó un CSV dos veces',
        '3. Un trigger de formulario se disparó por duplicado en un intervalo corto',
        '4. La función nextId() devolvió el mismo número en llamadas concurrentes'
      ];
    } else {
      r.patronDetectado = 'Los IDs duplicados NO son consecutivos → sugiere doble envío de formulario o '
        + 'doble escritura puntual del script en momentos distintos.';
      r.causasProbables = [
        '1. El usuario envió el formulario dos veces para el mismo trabajo',
        '2. El script onFormSubmit se ejecutó dos veces por un trigger duplicado',
        '3. Error de sincronización en Apps Script (retry automático)',
        '4. Doble clic en el botón de guardar del panel'
      ];
    }
  }

  return r;
}

function _sonIdsConsecutivos_(ids) {
  if (ids.length < 3) return false;
  var nums = ids.map(function(id) {
    var m = id.match(/(\d+)$/);
    return m ? parseInt(m[1]) : NaN;
  }).filter(function(n) { return !isNaN(n); }).sort(function(a, b) { return a - b; });

  if (nums.length < 3) return false;
  var gaps = [];
  for (var i = 1; i < nums.length; i++) gaps.push(nums[i] - nums[i-1]);
  var maxGap = Math.max.apply(null, gaps);
  return maxGap <= 2;
}

// ---- Clasificador de registros ----

function _esRegistroPrueba_(campos) {
  var texto = [
    campos.nombre  || '',
    campos.empresa || '',
    campos.email   || '',
    campos.cif     || '',
    campos.dir     || '',
    campos.id      || ''
  ].join(' ').toLowerCase().trim();

  if (!texto || texto.length < 3) return 'DUDOSO';

  // Check test keywords
  for (var i = 0; i < TEST_KEYWORDS.length; i++) {
    if (texto.indexOf(TEST_KEYWORDS[i]) !== -1) return 'PRUEBA';
  }

  // Check test emails
  if (campos.email) {
    var email = campos.email.toLowerCase();
    for (var j = 0; j < TEST_EMAILS.length; j++) {
      if (email.indexOf(TEST_EMAILS[j]) !== -1) return 'PRUEBA';
    }
  }

  // Check test CIFs
  if (campos.cif) {
    var cif = campos.cif.toUpperCase().replace(/\s/g, '');
    for (var k = 0; k < TEST_CIFS.length; k++) {
      if (cif === TEST_CIFS[k]) return 'PRUEBA';
    }
  }

  // Check test addresses
  if (campos.dir) {
    var dir = campos.dir.toLowerCase();
    for (var l = 0; l < TEST_ADDRESSES.length; l++) {
      if (dir.indexOf(TEST_ADDRESSES[l].toLowerCase()) !== -1) return 'PRUEBA';
    }
  }

  // Nombre por defecto de empresa
  if (campos.nombre && campos.nombre.toLowerCase().indexOf(EMPRESA_NOMBRE_DEFAULT.toLowerCase()) !== -1) {
    return 'PRUEBA';
  }

  // Sin email ni CIF → dudoso
  if (!campos.email && !campos.cif) return 'DUDOSO';

  return 'REAL';
}

// ---- Cálculo de fases de recomendación ----

function _calcularFase1_(res) {
  var fase1 = [];

  // Clientes 100% de prueba con keywords claras
  (res.clientes.registrosPrueba || []).forEach(function(r) {
    fase1.push({
      hoja: 'CLIENTES', fila: r.fila, id: r.id,
      razon: 'Cliente con nombre/email/CIF de prueba: "' + (r.nombre || r.empresa) + '"',
      accion: 'Eliminar'
    });
  });

  // Trabajos de prueba
  (res.trabajos.registrosPrueba || []).forEach(function(r) {
    fase1.push({
      hoja: 'TRABAJOS', fila: r.fila, id: r.id,
      razon: 'Trabajo asociado a cliente de prueba: "' + r.nomCliente + '"',
      accion: 'Eliminar'
    });
  });

  // Duplicados exactos en TRABAJOS
  (res.duplicadosTRA.duplicadosExactos || []).forEach(function(d) {
    if (d.causa.indexOf('100% idénticas') !== -1) {
      var filasExtra = d.filas.slice(1);
      filasExtra.forEach(function(f) {
        fase1.push({
          hoja: 'TRABAJOS', fila: f, id: d.id,
          razon: 'Fila duplicada idéntica (mantener solo fila ' + d.filas[0] + ')',
          accion: 'Eliminar duplicado'
        });
      });
    }
  });

  // Facturas de prueba
  (res.facturas.registrosPrueba || []).forEach(function(r) {
    fase1.push({
      hoja: 'FACTURAS', fila: r.fila, id: r.id,
      razon: 'Factura de cliente de prueba: "' + r.nomCliente + '"',
      accion: 'Eliminar'
    });
  });

  return fase1;
}

function _calcularFase2_(res) {
  var fase2 = [];

  // Registros dudosos de clientes
  (res.clientes.registrosDudosos || []).forEach(function(r) {
    fase2.push({
      hoja: 'CLIENTES', fila: r.fila, id: r.id,
      razon: 'Sin email ni CIF — verificar si es real: "' + (r.nombre || r.empresa) + '"',
      accion: 'Revisar manualmente'
    });
  });

  // Duplicados no exactos en TRABAJOS
  (res.duplicadosTRA.duplicadosExactos || []).forEach(function(d) {
    if (d.causa.indexOf('100% idénticas') === -1) {
      fase2.push({
        hoja: 'TRABAJOS', id: d.id,
        filas: d.filas.join(', '),
        razon: d.causa,
        accion: 'Revisar y decidir cuál conservar'
      });
    }
  });

  // Facturas sin cliente
  (res.facturas.sinCliente || []).forEach(function(r) {
    fase2.push({
      hoja: 'FACTURAS', fila: r.fila, id: r.id,
      razon: 'Factura sin ID_Cliente vinculado',
      accion: 'Asignar cliente o eliminar'
    });
  });

  // Registros huérfanos
  (res.integridad.huerfanos || []).forEach(function(h) {
    fase2.push({
      hoja: h.tipo.split(' ')[0],
      fila: h.fila, id: h.id,
      razon: h.tipo + ': referencia "' + h.referencia + '" (' + h.nombreRef + ') no existe en CLIENTES',
      accion: 'Revisar y corregir referencia o eliminar'
    });
  });

  // Gastos sin archivo
  if (res.gastos.sinArchivo && res.gastos.sinArchivo.length > 0) {
    res.gastos.sinArchivo.slice(0, 20).forEach(function(r) {
      fase2.push({
        hoja: 'GASTOS', fila: r.fila, id: r.id,
        razon: 'Gasto sin archivo asociado (Archivo_ID vacío): "' + r.concepto + '"',
        accion: 'Adjuntar documento o marcar como manual'
      });
    });
    if (res.gastos.sinArchivo.length > 20) {
      fase2.push({
        hoja: 'GASTOS', fila: '…', id: '…',
        razon: '+ ' + (res.gastos.sinArchivo.length - 20) + ' gastos más sin archivo',
        accion: 'Revisar GASTOS en bloque'
      });
    }
  }

  return fase2;
}

function _calcularFase3_(res) {
  var fase3 = [];

  // Diagnóstico de causa de IDs duplicados
  if ((res.duplicadosTRA.duplicadosExactos || []).length > 0) {
    fase3.push({
      area: 'Script — Duplicados TRABAJOS',
      descripcion: res.duplicadosTRA.patronDetectado,
      causasProbables: (res.duplicadosTRA.causasProbables || []).join('\n'),
      recomendacion: 'Revisar la función onFormSubmit y nextId() para asegurar idempotencia. '
        + 'Considerar un mecanismo de lock o deduplicación por hash de contenido.'
    });
  }

  // IDs duplicados en otras hojas
  if ((res.clientes.duplicadosId || []).length > 0) {
    fase3.push({
      area: 'Script — Duplicados ID en CLIENTES',
      descripcion: 'Los IDs ' + res.clientes.duplicadosId.join(', ') + ' aparecen más de una vez.',
      causasProbables: 'La función nextId() no es atómica; en llamadas concurrentes puede devolver el mismo número.',
      recomendacion: 'Implementar locking en nextId() usando LockService de Apps Script.'
    });
  }

  if ((res.facturas.duplicadosId || []).length > 0) {
    fase3.push({
      area: 'Script — Duplicados ID en FACTURAS',
      descripcion: 'IDs duplicados: ' + res.facturas.duplicadosId.join(', '),
      causasProbables: 'Posible doble llamada a la función de generación de factura.',
      recomendacion: 'Verificar que generarFactura() no se llama dos veces en el mismo flujo.'
    });
  }

  if ((res.facturas.duplicadosNum || []).length > 0) {
    fase3.push({
      area: 'Script — Números de factura duplicados',
      descripcion: 'Números duplicados: ' + res.facturas.duplicadosNum.join(', '),
      causasProbables: 'El contador "siguiente_numero" en CONFIG no se incrementó correctamente.',
      recomendacion: 'Revisar la lógica de incremento del número de factura y usar LockService.'
    });
  }

  // Huérfanos como posible bug
  if ((res.integridad.huerfanos || []).length > 3) {
    fase3.push({
      area: 'Script — Referencias rotas en cantidad',
      descripcion: (res.integridad.huerfanos.length) + ' registros con ID_Cliente que no existe en CLIENTES.',
      causasProbables: 'Posible borrado de clientes sin actualizar las tablas relacionadas, '
        + 'o importación de datos con IDs distintos a los actuales.',
      recomendacion: 'Añadir validación referencial antes de borrar clientes. '
        + 'Verificar que el formulario de trabajos/facturas lee los IDs desde CLIENTES.'
    });
  }

  return fase3;
}

// ---- Generación de la hoja AUDITORIA_DATOS ----

function _generarHojaAuditoria_(resultado) {
  var ss = getTargetSS_();
  var sh = ss.getSheetByName('AUDITORIA_DATOS');
  if (sh) { sh.clear(); sh.clearFormats(); }
  else    { sh = ss.insertSheet('AUDITORIA_DATOS'); }

  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 110);
  sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 110);
  sh.setColumnWidth(5, 110);
  sh.setColumnWidth(6, 110);
  sh.setColumnWidth(7, 350);

  var tz    = Session.getScriptTimeZone();
  var ahora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
  var NC    = 7;
  var r     = 1;

  // ── TÍTULO PRINCIPAL ──────────────────────────────────────────────
  _auTitulo_(sh, r, NC, '🔍 AUDITORÍA DE DATOS — SISTEMA DE GESTIÓN ÁLVARO'); r++;
  _auSubtitulo_(sh, r, NC, 'Generado el ' + ahora + '   |   SOLO LECTURA — NO SE HA MODIFICADO NI ELIMINADO NINGÚN DATO'); r++;
  r++;

  // ── RESUMEN GENERAL ───────────────────────────────────────────────
  _auSeccion_(sh, r, NC, '📊 RESUMEN GENERAL'); r++;
  _auCabecera_(sh, r, ['Hoja', 'Total', 'Reales', 'Prueba', 'Dudosos', 'Duplicados', 'Huérfanos']); r++;

  var hojas = [
    resultado.clientes, resultado.trabajos, resultado.facturas,
    resultado.proformas, resultado.presupuestos, resultado.gastos,
    resultado.informes, resultado.recLog, resultado.alertasLog
  ];

  hojas.forEach(function(h, idx) {
    var bg = idx % 2 === 0 ? '#ffffff' : '#f8f9fa';
    sh.getRange(r, 1, 1, NC).setBackground(bg);
    sh.getRange(r, 1).setValue(h.nombre || '—');
    if (h.error) {
      sh.getRange(r, 2, 1, 6).merge().setValue('⚠ ' + h.error).setFontColor('#b06000');
    } else {
      sh.getRange(r, 2).setValue(h.total   || 0).setHorizontalAlignment('center');
      sh.getRange(r, 3).setValue(h.reales  || 0).setHorizontalAlignment('center').setFontColor('#137333');
      sh.getRange(r, 4).setValue(h.prueba  || 0).setHorizontalAlignment('center').setFontColor('#c5221f');
      sh.getRange(r, 5).setValue(h.dudosos || 0).setHorizontalAlignment('center').setFontColor('#b06000');
      var nDups = (h.duplicados || []).length;
      sh.getRange(r, 6).setValue(nDups).setHorizontalAlignment('center').setFontColor(nDups > 0 ? '#c5221f' : '#137333');
      sh.getRange(r, 7).setValue('').setHorizontalAlignment('center');
    }
    r++;
  });

  // Fila TOTALES
  var res = resultado.resumen;
  sh.getRange(r, 1, 1, NC).setBackground('#e8f0fe');
  sh.getRange(r, 1).setValue('TOTAL').setFontWeight('bold');
  sh.getRange(r, 2).setValue(res.totalRegistros).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(r, 3).setValue(res.totalReales).setFontWeight('bold').setHorizontalAlignment('center').setFontColor('#137333');
  sh.getRange(r, 4).setValue(res.totalPrueba).setFontWeight('bold').setHorizontalAlignment('center').setFontColor('#c5221f');
  sh.getRange(r, 5).setValue(res.totalDudosos).setFontWeight('bold').setHorizontalAlignment('center').setFontColor('#b06000');
  sh.getRange(r, 6).setValue(res.totalDuplicados).setFontWeight('bold').setHorizontalAlignment('center').setFontColor(res.totalDuplicados > 0 ? '#c5221f' : '#137333');
  sh.getRange(r, 7).setValue(res.totalHuerfanos).setFontWeight('bold').setHorizontalAlignment('center').setFontColor(res.totalHuerfanos > 0 ? '#c5221f' : '#137333');
  r += 2;

  // ── SECCIÓN: CLIENTES REALES ──────────────────────────────────────
  _auSeccion_(sh, r, NC, '✅ CLIENTES REALES (' + (resultado.clientes.registrosReales || []).length + ')'); r++;
  _auCabecera_(sh, r, ['Fila', 'ID', 'Nombre', 'Empresa', 'Email', 'CIF', '']); r++;
  (resultado.clientes.registrosReales || []).forEach(function(c, idx) {
    _auFila_(sh, r, idx, [c.fila, c.id, c.nombre, c.empresa, c.email, c.cif, '']); r++;
  });
  r++;

  // ── SECCIÓN: CLIENTES DE PRUEBA ───────────────────────────────────
  _auSeccion_(sh, r, NC, '🧪 CLIENTES DE PRUEBA (' + (resultado.clientes.registrosPrueba || []).length + ')', '#fce8b2', '#b06000'); r++;
  if (resultado.clientes.registrosPrueba.length === 0) {
    _auMensaje_(sh, r, NC, 'No se detectaron clientes de prueba.', '#137333'); r++;
  } else {
    _auCabecera_(sh, r, ['Fila', 'ID', 'Nombre', 'Empresa', 'Email', 'CIF', '']); r++;
    resultado.clientes.registrosPrueba.forEach(function(c, idx) {
      _auFila_(sh, r, idx, [c.fila, c.id, c.nombre, c.empresa, c.email, c.cif, ''], '#fff3e0'); r++;
    });
  }
  r++;

  // ── SECCIÓN: CLIENTES DUDOSOS ─────────────────────────────────────
  _auSeccion_(sh, r, NC, '❓ CLIENTES DUDOSOS (' + (resultado.clientes.registrosDudosos || []).length + ')', '#fce8b2', '#b06000'); r++;
  if (resultado.clientes.registrosDudosos.length === 0) {
    _auMensaje_(sh, r, NC, 'No se detectaron clientes dudosos.', '#137333'); r++;
  } else {
    _auCabecera_(sh, r, ['Fila', 'ID', 'Nombre', 'Empresa', 'Email', 'CIF', 'Motivo']); r++;
    resultado.clientes.registrosDudosos.forEach(function(c, idx) {
      _auFila_(sh, r, idx, [c.fila, c.id, c.nombre, c.empresa, c.email, c.cif, 'Sin email y sin CIF'], '#fffde7'); r++;
    });
  }
  r++;

  // ── SECCIÓN: TRABAJOS DUPLICADOS ──────────────────────────────────
  _auSeccion_(sh, r, NC, '⚠️ TRABAJOS CON ID DUPLICADO (' + Object.keys(resultado.duplicadosTRA.idsMultiples).length + ' IDs afectados)', '#fce8b2', '#b06000'); r++;

  if (resultado.duplicadosTRA.patronDetectado) {
    sh.getRange(r, 1, 1, NC).merge()
      .setValue('📌 PATRÓN DETECTADO: ' + resultado.duplicadosTRA.patronDetectado)
      .setBackground('#fff9c4').setFontColor('#b06000').setFontWeight('bold').setWrap(true);
    sh.setRowHeight(r, 50); r++;
  }

  if ((resultado.duplicadosTRA.duplicadosExactos || []).length === 0) {
    _auMensaje_(sh, r, NC, 'No se detectaron trabajos duplicados.', '#137333'); r++;
  } else {
    _auCabecera_(sh, r, ['ID Trabajo', 'Filas', 'Tipo duplicado', 'Causa probable', '', '', '']); r++;
    resultado.duplicadosTRA.duplicadosExactos.forEach(function(d, idx) {
      var bg = idx % 2 === 0 ? '#fff3e0' : '#fff8e1';
      sh.getRange(r, 1, 1, NC).setBackground(bg);
      sh.getRange(r, 1).setValue(d.id);
      sh.getRange(r, 2).setValue('Filas: ' + d.filas.join(', '));
      sh.getRange(r, 3).setValue(d.causa.indexOf('100% idénticas') !== -1 ? 'IDÉNTICAS' : 'CON DIFERENCIAS');
      sh.getRange(r, 4, 1, 4).merge().setValue(d.causa).setWrap(true).setFontSize(9);
      sh.setRowHeight(r, 40); r++;
    });
  }
  r++;

  // Causas probables del patrón de duplicados
  if ((resultado.duplicadosTRA.causasProbables || []).length > 0) {
    _auSeccion_(sh, r, NC, '🔎 CAUSAS PROBABLES DE LOS DUPLICADOS EN TRABAJOS', '#e8f0fe', '#1a73e8'); r++;
    resultado.duplicadosTRA.causasProbables.forEach(function(causa, idx) {
      sh.getRange(r, 1, 1, NC).merge().setValue(causa)
        .setBackground(idx % 2 === 0 ? '#f8f9fa' : '#ffffff')
        .setFontSize(10).setWrap(true);
      r++;
    });
    r++;
  }

  // ── SECCIÓN: FACTURAS DUPLICADAS ──────────────────────────────────
  _auSeccion_(sh, r, NC, '⚠️ FACTURAS DUPLICADAS'); r++;
  var facDups = (resultado.facturas.duplicadosId || []).concat(resultado.facturas.duplicadosNum || []);
  if (facDups.length === 0) {
    _auMensaje_(sh, r, NC, 'No se detectaron facturas duplicadas.', '#137333'); r++;
  } else {
    _auCabecera_(sh, r, ['Tipo', 'Valor duplicado', '', '', '', '', '']); r++;
    (resultado.facturas.duplicadosId || []).forEach(function(id, idx) {
      _auFila_(sh, r, idx, ['ID duplicado', id, '', '', '', '', ''], '#fff3e0'); r++;
    });
    (resultado.facturas.duplicadosNum || []).forEach(function(num, idx) {
      _auFila_(sh, r, idx, ['Número de factura duplicado', num, '', '', '', '', ''], '#fff8e1'); r++;
    });
  }
  r++;

  // ── SECCIÓN: PRESUPUESTOS DUPLICADOS ──────────────────────────────
  _auSeccion_(sh, r, NC, '⚠️ PRESUPUESTOS DUPLICADOS'); r++;
  if ((resultado.presupuestos.duplicados || []).length === 0) {
    _auMensaje_(sh, r, NC, 'No se detectaron presupuestos duplicados.', '#137333'); r++;
  } else {
    _auCabecera_(sh, r, ['ID', '', '', '', '', '', '']); r++;
    resultado.presupuestos.duplicados.forEach(function(id, idx) {
      _auFila_(sh, r, idx, [id, '', '', '', '', '', ''], '#fff3e0'); r++;
    });
  }
  r++;

  // ── SECCIÓN: PROFORMAS DUPLICADAS ─────────────────────────────────
  _auSeccion_(sh, r, NC, '⚠️ PROFORMAS DUPLICADAS'); r++;
  if ((resultado.proformas.duplicados || []).length === 0) {
    _auMensaje_(sh, r, NC, 'No se detectaron proformas duplicadas.', '#137333'); r++;
  } else {
    _auCabecera_(sh, r, ['ID', '', '', '', '', '', '']); r++;
    resultado.proformas.duplicados.forEach(function(id, idx) {
      _auFila_(sh, r, idx, [id, '', '', '', '', '', ''], '#fff3e0'); r++;
    });
  }
  r++;

  // ── SECCIÓN: REGISTROS HUÉRFANOS ──────────────────────────────────
  _auSeccion_(sh, r, NC, '🔗 REGISTROS HUÉRFANOS (' + (resultado.integridad.huerfanos || []).length + ')', '#fce8b2', '#b06000'); r++;
  if ((resultado.integridad.huerfanos || []).length === 0) {
    _auMensaje_(sh, r, NC, 'No se detectaron registros huérfanos. La integridad referencial es correcta.', '#137333'); r++;
  } else {
    _auCabecera_(sh, r, ['Tipo', 'Fila', 'ID', 'Referencia rota', 'Nombre ref.', '', '']); r++;
    resultado.integridad.huerfanos.forEach(function(h, idx) {
      var bg = idx % 2 === 0 ? '#fff3e0' : '#fff8e1';
      sh.getRange(r, 1, 1, NC).setBackground(bg);
      sh.getRange(r, 1).setValue(h.tipo);
      sh.getRange(r, 2).setValue(h.fila || '—').setHorizontalAlignment('center');
      sh.getRange(r, 3).setValue(h.id || '—');
      sh.getRange(r, 4).setValue(h.referencia || '—');
      sh.getRange(r, 5, 1, 3).merge().setValue(h.nombreRef || '—');
      r++;
    });
  }
  r++;

  // ── SECCIÓN: ANÁLISIS DE GASTOS ───────────────────────────────────
  _auSeccion_(sh, r, NC, '💰 ANÁLISIS DE GASTOS'); r++;
  sh.getRange(r, 1, 1, NC).setBackground('#f8f9fa');
  sh.getRange(r, 1).setValue('Total gastos').setFontWeight('bold');
  sh.getRange(r, 2).setValue(resultado.gastos.total || 0).setHorizontalAlignment('center');
  r++;
  sh.getRange(r, 1, 1, NC).setBackground('#ffffff');
  sh.getRange(r, 1).setValue('Gastos reales');
  sh.getRange(r, 2).setValue(resultado.gastos.reales || 0).setHorizontalAlignment('center').setFontColor('#137333');
  r++;
  sh.getRange(r, 1, 1, NC).setBackground('#f8f9fa');
  sh.getRange(r, 1).setValue('Gastos de prueba/ejemplo');
  sh.getRange(r, 2).setValue(resultado.gastos.prueba || 0).setHorizontalAlignment('center').setFontColor('#c5221f');
  r++;
  sh.getRange(r, 1, 1, NC).setBackground('#ffffff');
  sh.getRange(r, 1).setValue('Con texto OCR (escaneados)');
  sh.getRange(r, 2).setValue((resultado.gastos.ocr || []).length).setHorizontalAlignment('center');
  r++;
  sh.getRange(r, 1, 1, NC).setBackground('#f8f9fa');
  sh.getRange(r, 1).setValue('Sin archivo asociado');
  sh.getRange(r, 2).setValue((resultado.gastos.sinArchivo || []).length).setHorizontalAlignment('center')
    .setFontColor((resultado.gastos.sinArchivo || []).length > 0 ? '#b06000' : '#137333');
  r += 2;

  // ── SECCIÓN: RECOMENDACIONES ──────────────────────────────────────
  _auSeccion_(sh, r, NC, '💡 RECOMENDACIONES GENERALES'); r++;
  var recomendaciones = [
    ['1', 'Ejecutar FASE 1 primero', 'Eliminar los registros clasificados como de prueba. Son seguros de borrar y no tienen impacto en datos reales.'],
    ['2', 'Resolver duplicados de TRABAJOS', 'Para cada ID duplicado, verificar manualmente cuál fila es la correcta y eliminar las copias extra.'],
    ['3', 'Corregir huérfanos', 'Revisar cada registro huérfano y asignar el cliente correcto o eliminar el registro.'],
    ['4', 'Proteger nextId() con LockService', 'Si los duplicados de ID son frecuentes, implementar LockService para evitar condiciones de carrera.'],
    ['5', 'Hacer backup antes de borrar', 'Antes de eliminar cualquier dato, exportar la hoja correspondiente a CSV como respaldo.']
  ];
  _auCabecera_(sh, r, ['#', 'Acción', 'Descripción', '', '', '', '']); r++;
  recomendaciones.forEach(function(rec, idx) {
    var bg = idx % 2 === 0 ? '#ffffff' : '#f8f9fa';
    sh.getRange(r, 1, 1, NC).setBackground(bg);
    sh.getRange(r, 1).setValue(rec[0]).setHorizontalAlignment('center').setFontWeight('bold');
    sh.getRange(r, 2).setValue(rec[1]).setFontWeight('bold');
    sh.getRange(r, 3, 1, 5).merge().setValue(rec[2]).setWrap(true).setFontSize(10);
    sh.setRowHeight(r, 35); r++;
  });
  r++;

  // ── FASE 1: REGISTROS SEGUROS PARA ELIMINAR ───────────────────────
  _auSeccion_(sh, r, NC, '🟢 FASE 1 — REGISTROS SEGUROS PARA ELIMINAR (' + resultado.fase1.length + ')', '#d4edda', '#155724'); r++;
  if (resultado.fase1.length === 0) {
    _auMensaje_(sh, r, NC, 'No hay registros claramente identificados como seguros para eliminar.', '#137333'); r++;
  } else {
    _auCabecera_(sh, r, ['Hoja', 'Fila', 'ID', 'Motivo', 'Acción', '', '']); r++;
    resultado.fase1.forEach(function(f, idx) {
      var bg = idx % 2 === 0 ? '#f0fff4' : '#e8f5e9';
      sh.getRange(r, 1, 1, NC).setBackground(bg);
      sh.getRange(r, 1).setValue(f.hoja);
      sh.getRange(r, 2).setValue(f.fila || '—').setHorizontalAlignment('center');
      sh.getRange(r, 3).setValue(f.id || '—');
      sh.getRange(r, 4, 1, 3).merge().setValue(f.razon).setWrap(true).setFontSize(9);
      sh.getRange(r, 7).setValue(f.accion).setFontColor('#c5221f').setFontWeight('bold');
      sh.setRowHeight(r, 35); r++;
    });
  }
  r++;

  // ── FASE 2: REGISTROS QUE REQUIEREN REVISIÓN MANUAL ──────────────
  _auSeccion_(sh, r, NC, '🟡 FASE 2 — REQUIEREN REVISIÓN MANUAL (' + resultado.fase2.length + ')', '#fff3cd', '#856404'); r++;
  if (resultado.fase2.length === 0) {
    _auMensaje_(sh, r, NC, 'No hay registros que requieran revisión manual.', '#137333'); r++;
  } else {
    _auCabecera_(sh, r, ['Hoja', 'Fila', 'ID', 'Motivo', 'Acción', '', '']); r++;
    resultado.fase2.forEach(function(f, idx) {
      var bg = idx % 2 === 0 ? '#fffde7' : '#fff9c4';
      sh.getRange(r, 1, 1, NC).setBackground(bg);
      sh.getRange(r, 1).setValue(f.hoja);
      sh.getRange(r, 2).setValue(f.fila || f.filas || '—').setHorizontalAlignment('center');
      sh.getRange(r, 3).setValue(f.id || '—');
      sh.getRange(r, 4, 1, 3).merge().setValue(f.razon).setWrap(true).setFontSize(9);
      sh.getRange(r, 7).setValue(f.accion).setFontColor('#b06000').setFontWeight('bold');
      sh.setRowHeight(r, 40); r++;
    });
  }
  r++;

  // ── FASE 3: POSIBLES ERRORES DE PROGRAMACIÓN ──────────────────────
  _auSeccion_(sh, r, NC, '🔴 FASE 3 — POSIBLES ERRORES DE PROGRAMACIÓN (' + resultado.fase3.length + ')', '#f8d7da', '#721c24'); r++;
  if (resultado.fase3.length === 0) {
    _auMensaje_(sh, r, NC, 'No se detectaron errores de programación evidentes.', '#137333'); r++;
  } else {
    resultado.fase3.forEach(function(f, idx) {
      var bg = idx % 2 === 0 ? '#fff5f5' : '#ffeaea';
      sh.getRange(r, 1, 1, NC).setBackground(bg);
      sh.getRange(r, 1).setValue('⚠ ' + f.area).setFontWeight('bold').setFontColor('#721c24');
      sh.setRowHeight(r, 30); r++;
      sh.getRange(r, 1, 1, NC).setBackground(bg);
      sh.getRange(r, 1, 1, NC).merge().setValue('Descripción: ' + f.descripcion).setWrap(true).setFontSize(10);
      sh.setRowHeight(r, 45); r++;
      if (f.causasProbables) {
        sh.getRange(r, 1, 1, NC).setBackground(bg);
        sh.getRange(r, 1, 1, NC).merge().setValue('Causas probables:\n' + f.causasProbables).setWrap(true).setFontSize(9).setFontColor('#5f6368');
        sh.setRowHeight(r, 70); r++;
      }
      sh.getRange(r, 1, 1, NC).setBackground(bg);
      sh.getRange(r, 1, 1, NC).merge().setValue('Recomendación: ' + f.recomendacion).setWrap(true).setFontSize(10).setFontColor('#1a73e8');
      sh.setRowHeight(r, 45); r++;
      r++;
    });
  }

  // ── PIE ───────────────────────────────────────────────────────────
  r++;
  sh.getRange(r, 1, 1, NC).merge()
    .setValue('AUDITORIA_DATOS — Generado el ' + ahora + ' — Sistema de Gestión Álvaro')
    .setFontColor('#9aa0a6').setFontSize(9).setHorizontalAlignment('center');

  SpreadsheetApp.flush();
  Logger.log('AUDITORIA_DATOS generada: ' + r + ' filas, ' + ahora);
}

// ---- Helpers de formato ----

function _auTitulo_(sh, r, nc, txt) {
  sh.getRange(r, 1, 1, nc).merge()
    .setValue(txt)
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(r, 50);
}

function _auSubtitulo_(sh, r, nc, txt) {
  sh.getRange(r, 1, 1, nc).merge()
    .setValue(txt)
    .setBackground('#e8f0fe').setFontColor('#5f6368')
    .setFontSize(10).setHorizontalAlignment('center');
}

function _auSeccion_(sh, r, nc, txt, bg, fg) {
  sh.getRange(r, 1, 1, nc).merge()
    .setValue(txt)
    .setBackground(bg || '#e8f0fe').setFontColor(fg || '#1a73e8')
    .setFontSize(12).setFontWeight('bold').setHorizontalAlignment('left');
  sh.setRowHeight(r, 30);
}

function _auCabecera_(sh, r, hdrs) {
  sh.getRange(r, 1, 1, hdrs.length).setValues([hdrs])
    .setBackground('#202124').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(10);
}

function _auFila_(sh, r, idx, valores, bg) {
  var fondo = bg || (idx % 2 === 0 ? '#ffffff' : '#f8f9fa');
  sh.getRange(r, 1, 1, valores.length).setBackground(fondo);
  valores.forEach(function(v, c) {
    sh.getRange(r, c + 1).setValue(v !== undefined && v !== null ? v : '');
  });
}

function _auMensaje_(sh, r, nc, txt, color) {
  sh.getRange(r, 1, 1, nc).merge()
    .setValue(txt)
    .setFontColor(color || '#202124').setFontStyle('italic').setHorizontalAlignment('center');
}
