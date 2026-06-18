// ============================================================
// AuditoriaFase2.gs — Segunda fase: informe ejecutivo detallado
// Lee directamente las hojas originales. NO MODIFICA NADA.
// ============================================================

// ---- Punto de entrada ----

function ejecutarAuditoriaFase2() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    '🔬 ANÁLISIS FASE 2 EN CURSO',
    'Ejecutando análisis profundo de todos los registros...\n'
    + 'Esto puede tardar 60-90 segundos. No cierres la hoja.',
    ui.ButtonSet.OK
  );

  try {
    var datos = _f2LeerTodo_();
    _f2GenerarInforme_(datos);

    var ss = getTargetSS_();
    var sh = ss.getSheetByName('AUDITORIA_FASE2');
    if (sh) ss.setActiveSheet(sh);

    ui.alert(
      '✅ FASE 2 COMPLETADA',
      'Informe ejecutivo generado en: AUDITORIA_FASE2\n\n'
      + 'Revisa las secciones:\n'
      + '  🟢 ELIMINAR CON SEGURIDAD\n'
      + '  🟡 REVISAR MANUALMENTE\n'
      + '  🔴 NO TOCAR\n'
      + '  📋 PLAN DE LIMPIEZA PASO A PASO',
      ui.ButtonSet.OK
    );
  } catch (e) {
    logEvent('ERROR', 'ejecutarAuditoriaFase2', e.message, e.stack);
    ui.alert('❌ ERROR en Fase 2:\n' + e.message + '\n\n' + e.stack.substring(0, 500));
  }
}

// ============================================================
// LECTURA DE DATOS CRUDOS
// ============================================================

function _f2LeerTodo_() {
  var ss = getTargetSS_();
  var tz = Session.getScriptTimeZone();

  return {
    tz:          tz,
    clientes:    _f2LeerHoja_(ss, 'CLIENTES',     11),
    trabajos:    _f2LeerHoja_(ss, 'TRABAJOS',     19),
    facturas:    _f2LeerHoja_(ss, 'FACTURAS',     23),
    proformas:   _f2LeerHoja_(ss, 'PROFORMAS',    21),
    presupuestos:_f2LeerHoja_(ss, 'PRESUPUESTOS', 13),
    gastos:      _f2LeerHoja_(ss, 'GASTOS',       25),
    informes:    _f2LeerHoja_(ss, 'INFORMES',     11),
    recLog:      _f2LeerHoja_(ss, 'RECORDATORIOS_LOG', 11),
    alertasLog:  _f2LeerHoja_(ss, 'ALERTAS_LOG',  10)
  };
}

function _f2LeerHoja_(ss, nombre, minCols) {
  var sh = ss.getSheetByName(nombre);
  if (!sh || sh.getLastRow() < 2) return { nombre: nombre, existe: false, filas: [], cabeceras: [] };
  var nCols = Math.max(sh.getLastColumn(), minCols || 1);
  var cabeceras = sh.getRange(1, 1, 1, nCols).getValues()[0].map(function(h) { return String(h).trim(); });
  var valores   = sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues();
  return { nombre: nombre, existe: true, filas: valores, cabeceras: cabeceras };
}

// ============================================================
// ANÁLISIS POR MÓDULO
// ============================================================

// ---- CLIENTES ----

function _f2AnalizarClientes_(datos) {
  var resultado = {
    reales: [], prueba: [], dudosos: [],
    duplicadosCruzados: [], resumen: {}
  };

  if (!datos.clientes.existe) { resultado.error = 'Hoja CLIENTES no encontrada'; return resultado; }

  var filas = datos.clientes.filas;

  // Índices de similitud para detectar duplicados cruzados
  var porEmail    = {};
  var porTelefono = {};
  var porNombre   = {};
  var porCif      = {};

  filas.forEach(function(row, idx) {
    var fila    = idx + 2;
    var id      = _s(row[0]);
    var ts      = row[1];
    var nombre  = _s(row[2]);
    var ape     = _s(row[3]);
    var empresa = _s(row[4]);
    var email   = _s(row[5]).toLowerCase();
    var tel     = _s(row[6]).replace(/\s/g, '');
    var dir     = _s(row[7]);
    var cif     = _s(row[8]).toUpperCase().replace(/\s/g, '');
    var activo  = _s(row[9]);

    var nombreCompleto = (nombre + ' ' + ape).trim();
    var label = nombreCompleto || empresa || email || id;

    var clasificacion = _f2ClasificarRegistro_({ nombre: label, empresa: empresa, email: email, cif: cif, dir: dir, id: id });

    var registro = {
      fila: fila, id: id, nombre: nombreCompleto, empresa: empresa,
      email: email, tel: tel, dir: dir, cif: cif, activo: activo,
      clasificacion: clasificacion.tipo,
      motivo: clasificacion.motivo
    };

    if (clasificacion.tipo === 'PRUEBA')  resultado.prueba.push(registro);
    else if (clasificacion.tipo === 'DUDOSO') resultado.dudosos.push(registro);
    else resultado.reales.push(registro);

    // Índices para detección cruzada
    if (email && email.length > 3) {
      if (porEmail[email]) porEmail[email].push(fila);
      else porEmail[email] = [fila];
    }
    var telNorm = tel.replace(/^\+34/, '').replace(/^34/, '');
    if (telNorm && telNorm.length >= 9) {
      if (porTelefono[telNorm]) porTelefono[telNorm].push(fila);
      else porTelefono[telNorm] = [fila];
    }
    var nomLow = nombreCompleto.toLowerCase().trim();
    if (nomLow && nomLow.length > 3 && nomLow !== ' ') {
      if (porNombre[nomLow]) porNombre[nomLow].push(fila);
      else porNombre[nomLow] = [fila];
    }
    if (cif && cif.length >= 8 && !_esCifPrueba_(cif)) {
      if (porCif[cif]) porCif[cif].push(fila);
      else porCif[cif] = [fila];
    }
  });

  // Detectar duplicados cruzados
  [
    { campo: 'Email',    idx: porEmail    },
    { campo: 'Teléfono', idx: porTelefono },
    { campo: 'Nombre',   idx: porNombre   },
    { campo: 'CIF/NIF',  idx: porCif      }
  ].forEach(function(grupo) {
    Object.keys(grupo.idx).forEach(function(val) {
      var filas = grupo.idx[val];
      if (filas.length > 1) {
        resultado.duplicadosCruzados.push({
          campo: grupo.campo,
          valor: val,
          filas: filas,
          riesgo: grupo.campo === 'Email' || grupo.campo === 'CIF/NIF' ? 'ALTO' : 'MEDIO'
        });
      }
    });
  });

  resultado.resumen = {
    total:   filas.length,
    reales:  resultado.reales.length,
    prueba:  resultado.prueba.length,
    dudosos: resultado.dudosos.length,
    duplicadosCruzados: resultado.duplicadosCruzados.length
  };

  return resultado;
}

// ---- TRABAJOS ----

function _f2AnalizarTrabajos_(datos) {
  var resultado = {
    reales: [], prueba: [], dudosos: [],
    idsMultiples: {}, analisisDuplicados: [],
    resumen: {}
  };

  if (!datos.trabajos.existe) { resultado.error = 'Hoja TRABAJOS no encontrada'; return resultado; }

  var filas = datos.trabajos.filas;
  var grupos = {};

  filas.forEach(function(row, idx) {
    var fila      = idx + 2;
    var id        = _s(row[0]);
    var fechaCre  = row[1];
    var idCli     = _s(row[2]);
    var nomCli    = _s(row[3]);
    var desc      = _s(row[6]);
    var tipo      = _s(row[7]);
    var dir       = _s(row[8]);
    var estado    = _s(row[11]);
    var precio    = row[15];

    var clasificacion = _f2ClasificarRegistro_({ nombre: nomCli + ' ' + desc, dir: dir, id: id });

    var reg = {
      fila: fila, id: id, idCli: idCli, nomCli: nomCli,
      desc: desc, tipo: tipo, dir: dir, estado: estado, precio: precio,
      fechaCre: fechaCre,
      clasificacion: clasificacion.tipo,
      motivo: clasificacion.motivo,
      row: row
    };

    if (clasificacion.tipo === 'PRUEBA')       resultado.prueba.push(reg);
    else if (clasificacion.tipo === 'DUDOSO')  resultado.dudosos.push(reg);
    else                                       resultado.reales.push(reg);

    if (!id) return;
    if (!grupos[id]) grupos[id] = [];
    grupos[id].push(reg);
  });

  // Análisis profundo de cada ID con múltiples filas
  Object.keys(grupos).forEach(function(id) {
    var grupo = grupos[id];
    if (grupo.length < 2) return;

    resultado.idsMultiples[id] = grupo.map(function(g) { return g.fila; });

    var analisis = _f2AnalizarGrupoDuplicado_(id, grupo);
    resultado.analisisDuplicados.push(analisis);
  });

  resultado.resumen = {
    total: filas.length,
    reales: resultado.reales.length,
    prueba: resultado.prueba.length,
    dudosos: resultado.dudosos.length,
    idsConDuplicados: Object.keys(resultado.idsMultiples).length,
    filasExtra: filas.length - Object.keys(grupos).length
  };

  return resultado;
}

function _f2AnalizarGrupoDuplicado_(id, grupo) {
  var analisis = {
    id: id,
    filas: grupo.map(function(g) { return g.fila; }),
    tipo: '',
    diferenciasCampo: [],
    timestamps: [],
    recomendacion: '',
    accion: '',
    riesgo: ''
  };

  var primerRow = grupo[0].row;
  var todasIguales = true;
  var camposDiferentes = [];

  var camposNombre = [
    'ID_Trabajo','Fecha_Creacion','ID_Cliente','Nombre_Cliente',
    'ID_Operario','Nombre_Operario','Descripcion','Tipo_Trabajo',
    'Direccion_Trabajo','Fecha_Inicio','Fecha_Fin_Estimada','Estado',
    'Horas_Trabajadas','Coste_Materiales','Precio_Mano_Obra','Precio_Total',
    'Observaciones','Facturado','ID_Factura'
  ];

  for (var c = 0; c < Math.min(primerRow.length, 19); c++) {
    var vals = grupo.map(function(g) { return String(g.row[c] || '').trim(); });
    var unico = vals.every(function(v) { return v === vals[0]; });
    if (!unico) {
      todasIguales = false;
      camposDiferentes.push({
        campo: camposNombre[c] || ('col' + (c+1)),
        valores: vals
      });
    }
  }

  // Analizar timestamps
  analisis.timestamps = grupo.map(function(g) {
    var ts = g.row[1];
    return ts instanceof Date ? ts.getTime() : null;
  });

  var tsValidos = analisis.timestamps.filter(function(t) { return t !== null; });
  var diffMaxMs = 0;
  if (tsValidos.length >= 2) {
    tsValidos.sort(function(a,b){return a-b;});
    diffMaxMs = tsValidos[tsValidos.length - 1] - tsValidos[0];
  }

  if (todasIguales) {
    analisis.tipo = 'DUPLICADO EXACTO';
    analisis.riesgo = 'BAJO';
    if (diffMaxMs < 5000) {
      analisis.recomendacion = 'Doble ejecución del script en menos de 5 segundos (retry automático de Apps Script o doble submit del formulario). Eliminar todas las filas excepto la primera.';
      analisis.accion = 'ELIMINAR_EXTRA';
    } else if (diffMaxMs < 60000) {
      analisis.recomendacion = 'Timestamps a menos de 60s → probable doble clic en el botón de guardar o doble envío del formulario. Eliminar duplicados.';
      analisis.accion = 'ELIMINAR_EXTRA';
    } else {
      analisis.recomendacion = 'Timestamps separados pero contenido idéntico → posible importación duplicada. Verificar el historial de importaciones antes de eliminar.';
      analisis.accion = 'REVISAR_ANTES_ELIMINAR';
    }
  } else {
    analisis.tipo = 'DUPLICADO PARCIAL (CON DIFERENCIAS)';
    analisis.diferenciasCampo = camposDiferentes;
    analisis.riesgo = 'ALTO';

    var tieneFactura  = grupo.some(function(g) { return _s(g.row[18]).indexOf('FAC-') === 0; });
    var estadosDist   = _unique_(grupo.map(function(g) { return _s(g.row[11]); }));
    var preciosDist   = _unique_(grupo.map(function(g) { return String(g.row[15] || ''); }));

    if (tieneFactura) {
      analisis.recomendacion = 'UNO O MÁS DUPLICADOS TIENEN ID_FACTURA → riesgo CRÍTICO de pérdida de factura si se elimina la fila incorrecta. Revisar cuál fila tiene la factura real y conservarla.';
      analisis.accion = 'NO_TOCAR_REVISAR_URGENTE';
      analisis.riesgo = 'CRÍTICO';
    } else if (estadosDist.length > 1) {
      analisis.recomendacion = 'Estados diferentes (' + estadosDist.join(' / ') + ') → posiblemente representan evoluciones del mismo trabajo o un error de sincronización. Revisar manualmente.';
      analisis.accion = 'REVISAR_MANUALMENTE';
    } else if (preciosDist.length > 1) {
      analisis.recomendacion = 'Precios diferentes → el trabajo puede haber sido modificado tras guardarse. Conservar la fila con el precio correcto (normalmente la más reciente).';
      analisis.accion = 'REVISAR_MANUALMENTE';
    } else {
      analisis.recomendacion = 'Diferencias menores entre campos. Revisar campo a campo y conservar la versión más completa.';
      analisis.accion = 'REVISAR_MANUALMENTE';
    }
  }

  return analisis;
}

// ---- FACTURAS ----

function _f2AnalizarFacturas_(datos, clientesIds) {
  var resultado = {
    reales: [], prueba: [], dudosos: [],
    huerfanas: [], duplicadosId: [], duplicadosNum: [],
    sinFecha: [], sinTotal: [],
    resumen: {}
  };

  if (!datos.facturas.existe) { resultado.error = 'Hoja FACTURAS no encontrada'; return resultado; }

  var filas = datos.facturas.filas;
  var idCount  = {};
  var numCount = {};

  filas.forEach(function(row, idx) {
    var fila      = idx + 2;
    var id        = _s(row[0]);
    var fechaEm   = row[1];
    var serie     = _s(row[2]);
    var numero    = _s(row[3]);
    var idCli     = _s(row[4]);
    var nomCli    = _s(row[5]);
    var total     = row[16];
    var estado    = _s(row[17]);

    if (id) { idCount[id] = (idCount[id] || 0) + 1; }
    var claveNum = (serie + numero).replace(/\s/g, '');
    if (numero) { numCount[claveNum] = (numCount[claveNum] || 0) + 1; }

    var esHuerfana = idCli && !clientesIds[idCli];
    var esSinFecha = !(fechaEm instanceof Date) || isNaN(fechaEm.getTime());
    var esSinTotal = !total && total !== 0;

    if (esHuerfana) resultado.huerfanas.push({ fila: fila, id: id, idCli: idCli, nomCli: nomCli });
    if (esSinFecha) resultado.sinFecha.push({ fila: fila, id: id });
    if (esSinTotal) resultado.sinTotal.push({ fila: fila, id: id });

    var clasif = _f2ClasificarRegistro_({ nombre: nomCli, id: id });
    var reg = { fila: fila, id: id, idCli: idCli, nomCli: nomCli,
                serie: serie, numero: numero, total: total, estado: estado,
                huerfana: esHuerfana, clasificacion: clasif.tipo, motivo: clasif.motivo };

    if (clasif.tipo === 'PRUEBA')       resultado.prueba.push(reg);
    else if (clasif.tipo === 'DUDOSO')  resultado.dudosos.push(reg);
    else                                resultado.reales.push(reg);
  });

  resultado.duplicadosId  = Object.keys(idCount).filter(function(k) { return idCount[k] > 1; });
  resultado.duplicadosNum = Object.keys(numCount).filter(function(k) { return numCount[k] > 1; });

  resultado.resumen = {
    total: filas.length,
    reales: resultado.reales.length,
    prueba: resultado.prueba.length,
    dudosos: resultado.dudosos.length,
    huerfanas: resultado.huerfanas.length,
    duplicadosId: resultado.duplicadosId.length,
    duplicadosNum: resultado.duplicadosNum.length
  };

  return resultado;
}

// ---- PROFORMAS ----

function _f2AnalizarProformas_(datos, clientesIds) {
  var resultado = {
    reales: [], prueba: [], dudosos: [],
    huerfanas: [], duplicadosId: [],
    resumen: {}
  };

  if (!datos.proformas.existe) { resultado.error = 'Hoja PROFORMAS no encontrada'; return resultado; }

  var filas = datos.proformas.filas;
  var idCount = {};

  filas.forEach(function(row, idx) {
    var fila   = idx + 2;
    var id     = _s(row[0]);
    var idCli  = _s(row[4]);
    var nomCli = _s(row[5]);
    var total  = row[12];
    var estado = _s(row[13]);

    if (id) { idCount[id] = (idCount[id] || 0) + 1; }
    if (idCli && !clientesIds[idCli]) {
      resultado.huerfanas.push({ fila: fila, id: id, idCli: idCli, nomCli: nomCli });
    }

    var clasif = _f2ClasificarRegistro_({ nombre: nomCli, id: id });
    var reg = { fila: fila, id: id, idCli: idCli, nomCli: nomCli,
                total: total, estado: estado,
                clasificacion: clasif.tipo, motivo: clasif.motivo };

    if (clasif.tipo === 'PRUEBA')       resultado.prueba.push(reg);
    else if (clasif.tipo === 'DUDOSO')  resultado.dudosos.push(reg);
    else                                resultado.reales.push(reg);
  });

  resultado.duplicadosId = Object.keys(idCount).filter(function(k) { return idCount[k] > 1; });

  resultado.resumen = {
    total: filas.length, reales: resultado.reales.length,
    prueba: resultado.prueba.length, dudosos: resultado.dudosos.length,
    huerfanas: resultado.huerfanas.length, duplicadosId: resultado.duplicadosId.length
  };

  return resultado;
}

// ---- PRESUPUESTOS ----

function _f2AnalizarPresupuestos_(datos, nombresClientes) {
  var resultado = {
    reales: [], prueba: [], dudosos: [],
    sinClienteEnBBDD: [], duplicadosId: [],
    resumen: {}
  };

  if (!datos.presupuestos.existe) { resultado.error = 'Hoja PRESUPUESTOS no encontrada'; return resultado; }

  var filas = datos.presupuestos.filas;
  var idCount = {};

  filas.forEach(function(row, idx) {
    var fila    = idx + 2;
    var id      = _s(row[0]);
    var fecha   = row[1];
    var cliente = _s(row[2]);
    var total   = row[7];
    var estado  = _s(row[8]);

    if (id) { idCount[id] = (idCount[id] || 0) + 1; }
    if (cliente && !nombresClientes[cliente.toLowerCase()]) {
      resultado.sinClienteEnBBDD.push({ fila: fila, id: id, cliente: cliente });
    }

    var clasif = _f2ClasificarRegistro_({ nombre: cliente, id: id });
    var reg = { fila: fila, id: id, cliente: cliente, total: total, estado: estado,
                fecha: fecha, clasificacion: clasif.tipo, motivo: clasif.motivo };

    if (clasif.tipo === 'PRUEBA')       resultado.prueba.push(reg);
    else if (clasif.tipo === 'DUDOSO')  resultado.dudosos.push(reg);
    else                                resultado.reales.push(reg);
  });

  resultado.duplicadosId = Object.keys(idCount).filter(function(k) { return idCount[k] > 1; });

  resultado.resumen = {
    total: filas.length, reales: resultado.reales.length,
    prueba: resultado.prueba.length, dudosos: resultado.dudosos.length,
    sinClienteEnBBDD: resultado.sinClienteEnBBDD.length,
    duplicadosId: resultado.duplicadosId.length
  };

  return resultado;
}

// ---- GASTOS ----

function _f2AnalizarGastos_(datos) {
  var resultado = {
    reales: [], prueba: [], dudosos: [],
    sinArchivo: [], conOcr: [], ocrIncompleto: [],
    sinTipo: [], sinFecha: [],
    pendientesPago: [],
    resumen: {}
  };

  if (!datos.gastos.existe) { resultado.error = 'Hoja GASTOS no encontrada'; return resultado; }

  var filas = datos.gastos.filas;

  filas.forEach(function(row, idx) {
    var fila       = idx + 2;
    var id         = _s(row[COL_GAS.ID_GASTO              - 1]);
    var fechaAlta  = row[COL_GAS.FECHA_ALTA               - 1];
    var tipoGasto  = _s(row[COL_GAS.TIPO_GASTO            - 1]);
    var fechaGasto = row[COL_GAS.FECHA_GASTO              - 1];
    var proveedor  = _s(row[COL_GAS.PROVEEDOR             - 1]);
    var concepto   = _s(row[COL_GAS.CONCEPTO              - 1]);
    var total      = row[COL_GAS.TOTAL                    - 1];
    var archivoId  = _s(row[COL_GAS.ARCHIVO_ID            - 1]);
    var ocrTexto   = _s(row[COL_GAS.OCR_TEXTO             - 1]);
    var confianza  = row[COL_GAS.CONFIANZA_OCR            - 1];
    var estadoPago = _s(row[COL_GAS.ESTADO_PAGO           - 1]);
    var estadoProc = _s(row[COL_GAS.ESTADO_PROCESAMIENTO  - 1]);

    var sinArchivo  = !archivoId;
    var tieneOcr    = ocrTexto && ocrTexto.length > 5;
    var ocrBaja     = tieneOcr && (confianza !== '' && parseFloat(confianza) < 0.6);
    var sinTipo     = !tipoGasto;
    var sinFechaGas = !(fechaGasto instanceof Date) || isNaN(fechaGasto.getTime());

    if (sinArchivo)  resultado.sinArchivo.push({ fila: fila, id: id, proveedor: proveedor, concepto: concepto, total: total });
    if (tieneOcr)    resultado.conOcr.push({ fila: fila, id: id, confianza: confianza });
    if (ocrBaja)     resultado.ocrIncompleto.push({ fila: fila, id: id, confianza: confianza });
    if (sinTipo)     resultado.sinTipo.push({ fila: fila, id: id });
    if (sinFechaGas) resultado.sinFecha.push({ fila: fila, id: id });
    if (estadoPago && estadoPago !== 'Pagado') resultado.pendientesPago.push({ fila: fila, id: id, total: total });

    var clasif = _f2ClasificarRegistro_({ nombre: proveedor + ' ' + concepto, id: id });
    var reg = { fila: fila, id: id, proveedor: proveedor, concepto: concepto,
                total: total, tipoGasto: tipoGasto, estadoPago: estadoPago,
                estadoProc: estadoProc, sinArchivo: sinArchivo, tieneOcr: tieneOcr,
                clasificacion: clasif.tipo, motivo: clasif.motivo };

    if (clasif.tipo === 'PRUEBA')       resultado.prueba.push(reg);
    else if (clasif.tipo === 'DUDOSO')  resultado.dudosos.push(reg);
    else                                resultado.reales.push(reg);
  });

  resultado.resumen = {
    total: filas.length,
    reales: resultado.reales.length,
    prueba: resultado.prueba.length,
    dudosos: resultado.dudosos.length,
    sinArchivo: resultado.sinArchivo.length,
    conOcr: resultado.conOcr.length,
    ocrIncompleto: resultado.ocrIncompleto.length,
    sinTipo: resultado.sinTipo.length,
    pendientesPago: resultado.pendientesPago.length
  };

  return resultado;
}

// ---- INTEGRIDAD REFERENCIAL ----

function _f2VerificarIntegridad_(datos, analisisClientes) {
  var resultado = {
    clientesIds: {},
    nombresClientes: {},
    roturas: [],
    ok: []
  };

  // Construir índice de clientes
  if (datos.clientes.existe) {
    datos.clientes.filas.forEach(function(row) {
      var id      = _s(row[0]);
      var nombre  = (_s(row[2]) + ' ' + _s(row[3])).trim().toLowerCase();
      var empresa = _s(row[4]).toLowerCase();
      if (id) resultado.clientesIds[id] = true;
      if (nombre && nombre !== ' ') resultado.nombresClientes[nombre] = true;
      if (empresa) resultado.nombresClientes[empresa] = true;
    });
  }

  // TRABAJOS → CLIENTES
  if (datos.trabajos.existe) {
    datos.trabajos.filas.forEach(function(row, idx) {
      var fila  = idx + 2;
      var id    = _s(row[0]);
      var idCli = _s(row[2]);
      var nom   = _s(row[3]);
      if (idCli && !resultado.clientesIds[idCli]) {
        resultado.roturas.push({ tipo: 'TRABAJO→CLIENTE', hoja: 'TRABAJOS', fila: fila, id: id,
          refRota: idCli, nombre: nom, impacto: 'El trabajo no tiene cliente válido en CLIENTES' });
      }
    });
  }

  // FACTURAS → CLIENTES
  if (datos.facturas.existe) {
    datos.facturas.filas.forEach(function(row, idx) {
      var fila  = idx + 2;
      var id    = _s(row[0]);
      var idCli = _s(row[4]);
      var nom   = _s(row[5]);
      if (idCli && !resultado.clientesIds[idCli]) {
        resultado.roturas.push({ tipo: 'FACTURA→CLIENTE', hoja: 'FACTURAS', fila: fila, id: id,
          refRota: idCli, nombre: nom, impacto: 'Factura vinculada a cliente inexistente — REVISAR ANTES DE BORRAR' });
      }
    });
  }

  // PROFORMAS → CLIENTES
  if (datos.proformas.existe) {
    datos.proformas.filas.forEach(function(row, idx) {
      var fila  = idx + 2;
      var id    = _s(row[0]);
      var idCli = _s(row[4]);
      var nom   = _s(row[5]);
      if (idCli && !resultado.clientesIds[idCli]) {
        resultado.roturas.push({ tipo: 'PROFORMA→CLIENTE', hoja: 'PROFORMAS', fila: fila, id: id,
          refRota: idCli, nombre: nom, impacto: 'Proforma vinculada a cliente inexistente' });
      }
    });
  }

  // PRESUPUESTOS → CLIENTES (por nombre)
  if (datos.presupuestos.existe) {
    datos.presupuestos.filas.forEach(function(row, idx) {
      var fila    = idx + 2;
      var id      = _s(row[0]);
      var cliente = _s(row[2]);
      if (cliente && !resultado.nombresClientes[cliente.toLowerCase()]) {
        resultado.roturas.push({ tipo: 'PRESUPUESTO→CLIENTE', hoja: 'PRESUPUESTOS', fila: fila, id: id,
          refRota: cliente, nombre: cliente, impacto: 'Presupuesto con nombre de cliente no encontrado en CLIENTES' });
      }
    });
  }

  // RECORDATORIOS_LOG → FACTURAS
  if (datos.recLog.existe && datos.facturas.existe) {
    var facIds = {};
    datos.facturas.filas.forEach(function(r) { var id = _s(r[0]); if (id) facIds[id] = true; });
    datos.recLog.filas.forEach(function(row, idx) {
      var fila  = idx + 2;
      var idRec = _s(row[0]);
      var idFac = _s(row[2]);
      if (idFac && !facIds[idFac]) {
        resultado.roturas.push({ tipo: 'RECORDATORIO→FACTURA', hoja: 'RECORDATORIOS_LOG', fila: fila, id: idRec,
          refRota: idFac, nombre: '', impacto: 'Log de recordatorio para factura que ya no existe' });
      }
    });
  }

  return resultado;
}

// ---- ANÁLISIS DE BUGS ----

function _f2AnalizarBugs_(anTrabajos) {
  var bugs = [];

  var nDups     = anTrabajos.analisisDuplicados.length;
  var nExactos  = anTrabajos.analisisDuplicados.filter(function(d) { return d.tipo === 'DUPLICADO EXACTO'; }).length;
  var nParciales= nDups - nExactos;

  if (nDups === 0) return bugs;

  // Analizar patrón de timestamps
  var diffsMenor5s  = 0;
  var diffsMenor60s = 0;
  var diffsGrandes  = 0;

  anTrabajos.analisisDuplicados.forEach(function(d) {
    var tsValidos = d.timestamps.filter(function(t) { return t !== null; });
    if (tsValidos.length >= 2) {
      tsValidos.sort(function(a,b){return a-b;});
      var diff = tsValidos[tsValidos.length-1] - tsValidos[0];
      if (diff < 5000)  diffsMenor5s++;
      else if (diff < 60000) diffsMenor60s++;
      else diffsGrandes++;
    }
  });

  // Bug 1: nextId() sin LockService
  bugs.push({
    titulo: 'nextId() sin LockService — condición de carrera',
    archivo: 'Utils.gs — función nextId()',
    descripcion: 'La función nextId() lee el último ID y genera el siguiente de forma no atómica. '
      + 'Si dos llamadas concurrentes (e.g. trigger + acción manual) se ejecutan simultáneamente, '
      + 'ambas leen el mismo "último ID" y generan el mismo ID nuevo.',
    evidencia: nDups + ' IDs duplicados detectados en TRABAJOS.',
    solucion: 'Envolver nextId() con LockService.getScriptLock() antes de leer y escribir. '
      + 'Alternativamente usar UUID (Utilities.getUuid()) como ID principal.',
    prioridad: 'ALTA',
    requiereLock: true
  });

  // Bug 2: onFormSubmit retry automático
  if (diffsMenor5s > 0) {
    bugs.push({
      titulo: 'Apps Script retry automático de onFormSubmit',
      archivo: 'Triggers.gs — onFormSubmitHandler()',
      descripcion: 'Apps Script reintenta automáticamente los triggers que fallan o tardan más de 30s. '
        + 'Si onTrabajosFormSubmit() lanza una excepción (o tarda mucho) en la primera ejecución, '
        + 'Apps Script puede ejecutarlo una segunda vez con el mismo evento, '
        + 'generando una segunda fila con el mismo (o diferente) ID.',
      evidencia: diffsMenor5s + ' duplicados con diferencia de timestamps < 5 segundos.',
      solucion: '1. Añadir idempotencia: calcular un hash del evento de formulario y rechazar duplicados.\n'
        + '2. Usar PropertiesService para marcar el eventId del formulario como procesado.\n'
        + '3. Asegurarse de que onTrabajosFormSubmit() completa en < 20s.',
      prioridad: 'ALTA',
      requiereLock: true
    });
  }

  // Bug 3: doble clic / doble submit
  if (diffsMenor60s > 0) {
    bugs.push({
      titulo: 'Doble envío de formulario por el usuario',
      archivo: 'Formulario de Trabajos (Google Form)',
      descripcion: 'El usuario envía el formulario dos veces rápidamente '
        + '(doble clic en "Enviar", recargar la página de confirmación, '
        + 'o usar el botón "Atrás" del navegador tras enviar).',
      evidencia: diffsMenor60s + ' duplicados con timestamps separados entre 5s y 60s.',
      solucion: '1. Activar "Limitar a 1 respuesta" en el formulario de Google.\n'
        + '2. Añadir validación en onTrabajosFormSubmit(): comparar campos clave '
        + '(cliente + fecha + descripción) con las últimas N filas y rechazar si ya existe.',
      prioridad: 'MEDIA',
      requiereLock: false
    });
  }

  // Bug 4: importación duplicada
  if (diffsGrandes > 0 || nExactos > (nDups / 2)) {
    bugs.push({
      titulo: 'Posible importación o migración duplicada',
      archivo: 'Cleanup.gs / Install.gs — funciones de migración',
      descripcion: 'Algunos duplicados tienen timestamps muy separados, '
        + 'lo que sugiere que se ejecutó una función de importación, migración '
        + 'o restauración de backup dos veces. Los datos se añadieron dos veces '
        + 'sin verificar si el ID ya existía.',
      evidencia: diffsGrandes + ' duplicados con timestamps separados > 1 minuto.',
      solucion: 'Añadir en todas las funciones de importación/migración una verificación '
        + 'previa de IDs existentes antes de insertar una nueva fila. '
        + 'Ejemplo: if (idExists(sheet, id)) { Logger.log("Skip: " + id); return; }',
      prioridad: 'MEDIA',
      requiereLock: false
    });
  }

  // Bug 5: falta de validación en el panel
  bugs.push({
    titulo: 'Panel sin validación de ID previo al guardar',
    archivo: 'Panel.gs / Trabajos.gs — funciones de alta de trabajo',
    descripcion: 'Cuando se registra un trabajo desde el panel (no desde el formulario), '
      + 'la función que crea la fila llama a nextId() pero no verifica si ese ID '
      + 'ya existe en la hoja antes de añadir la nueva fila.',
    evidencia: 'Patrón de duplicación consistente con IDs secuenciales repetidos.',
    solucion: 'Antes de appendRow(), verificar: if (findRowById_(sheet, id)) throw new Error("ID duplicado: " + id);\n'
      + 'Esto detectará la condición de carrera y lanzará un error en lugar de crear un duplicado.',
    prioridad: 'ALTA',
    requiereLock: true
  });

  return bugs;
}

// ---- TABLA SEMÁFORO ----

function _f2GenerarSemaforo_(anClientes, anTrabajos, anFacturas, anProformas, anPresupuestos, anGastos, integridad) {
  var verde  = [];
  var amarillo = [];
  var rojo   = [];

  // CLIENTES de prueba → verde (seguro eliminar)
  (anClientes.prueba || []).forEach(function(c) {
    verde.push({ hoja: 'CLIENTES', fila: c.fila, id: c.id,
      motivo: 'Cliente de prueba: ' + c.motivo,
      impacto: 'Sin impacto si no tiene trabajos/facturas asociados' });
  });

  // CLIENTES dudosos → amarillo
  (anClientes.dudosos || []).forEach(function(c) {
    amarillo.push({ hoja: 'CLIENTES', fila: c.fila, id: c.id,
      motivo: c.motivo,
      impacto: 'Verificar si tiene actividad antes de decidir' });
  });

  // CLIENTES reales → rojo
  (anClientes.reales || []).forEach(function(c) {
    rojo.push({ hoja: 'CLIENTES', fila: c.fila, id: c.id,
      motivo: 'Cliente real: ' + (c.nombre || c.empresa),
      impacto: 'Borrar eliminaría todos sus trabajos, facturas y presupuestos' });
  });

  // TRABAJOS de prueba → verde
  (anTrabajos.prueba || []).forEach(function(t) {
    verde.push({ hoja: 'TRABAJOS', fila: t.fila, id: t.id,
      motivo: 'Trabajo de cliente de prueba: ' + t.motivo,
      impacto: 'Sin impacto si la factura asociada también es de prueba' });
  });

  // Duplicados exactos de TRABAJOS → las filas extra van a verde
  (anTrabajos.analisisDuplicados || []).forEach(function(d) {
    if (d.accion === 'ELIMINAR_EXTRA') {
      d.filas.slice(1).forEach(function(f) {
        verde.push({ hoja: 'TRABAJOS', fila: f, id: d.id,
          motivo: 'Fila ' + f + ' es duplicado exacto de fila ' + d.filas[0],
          impacto: 'Bajo — el trabajo original se conserva en fila ' + d.filas[0] });
      });
    } else if (d.accion === 'REVISAR_ANTES_ELIMINAR' || d.accion === 'REVISAR_MANUALMENTE') {
      d.filas.forEach(function(f) {
        amarillo.push({ hoja: 'TRABAJOS', fila: f, id: d.id,
          motivo: d.recomendacion,
          impacto: 'Requiere revisión: ' + (d.diferenciasCampo || []).map(function(c){return c.campo;}).join(', ') });
      });
    } else if (d.accion === 'NO_TOCAR_REVISAR_URGENTE') {
      d.filas.forEach(function(f) {
        rojo.push({ hoja: 'TRABAJOS', fila: f, id: d.id,
          motivo: 'CRÍTICO: tiene ID_Factura asociado — ' + d.recomendacion,
          impacto: 'Borrar la fila incorrecta puede romper la factura asociada' });
      });
    }
  });

  // FACTURAS de prueba → verde
  (anFacturas.prueba || []).forEach(function(f) {
    verde.push({ hoja: 'FACTURAS', fila: f.fila, id: f.id,
      motivo: 'Factura de cliente de prueba: ' + f.motivo,
      impacto: 'Verificar que el PDF de prueba no está en Drive antes de borrar' });
  });

  // FACTURAS huérfanas → amarillo
  (anFacturas.huerfanas || []).forEach(function(f) {
    amarillo.push({ hoja: 'FACTURAS', fila: f.fila, id: f.id,
      motivo: 'Factura huérfana — ID_Cliente "' + f.idCli + '" no existe',
      impacto: 'Puede ser error de datos o cliente eliminado. Revisar.' });
  });

  // FACTURAS reales → rojo
  (anFacturas.reales || []).forEach(function(f) {
    rojo.push({ hoja: 'FACTURAS', fila: f.fila, id: f.id,
      motivo: 'Factura real de ' + (f.nomCli || f.idCli),
      impacto: 'Borrar elimina el registro fiscal y puede crear problemas contables' });
  });

  // PROFORMAS de prueba → verde
  (anProformas.prueba || []).forEach(function(p) {
    verde.push({ hoja: 'PROFORMAS', fila: p.fila, id: p.id,
      motivo: 'Proforma de cliente de prueba: ' + p.motivo,
      impacto: 'Sin impacto contable' });
  });

  // PRESUPUESTOS de prueba → verde
  (anPresupuestos.prueba || []).forEach(function(p) {
    verde.push({ hoja: 'PRESUPUESTOS', fila: p.fila, id: p.id,
      motivo: 'Presupuesto de cliente de prueba: ' + p.motivo,
      impacto: 'Sin impacto' });
  });

  // Gastos de prueba → verde
  (anGastos.prueba || []).forEach(function(g) {
    verde.push({ hoja: 'GASTOS', fila: g.fila, id: g.id,
      motivo: 'Gasto de prueba/ejemplo: ' + g.motivo,
      impacto: 'Sin impacto contable real' });
  });

  // Gastos dudosos → amarillo
  (anGastos.dudosos || []).forEach(function(g) {
    amarillo.push({ hoja: 'GASTOS', fila: g.fila, id: g.id,
      motivo: g.motivo + ' — sin proveedor ni tipo clasificado',
      impacto: 'Puede ser gasto real no completado. Revisar.' });
  });

  // Roturas de integridad → amarillo
  (integridad.roturas || []).forEach(function(r) {
    amarillo.push({ hoja: r.hoja, fila: r.fila, id: r.id,
      motivo: r.tipo + ': referencia "' + r.refRota + '" no existe',
      impacto: r.impacto });
  });

  return { verde: verde, amarillo: amarillo, rojo: rojo };
}

// ---- PLAN DE LIMPIEZA ----

function _f2PlanLimpieza_(semaforo, anTrabajos, bugs) {
  var pasos = [];

  pasos.push({
    numero: 1,
    titulo: 'BACKUP PREVIO OBLIGATORIO',
    descripcion: 'Antes de borrar cualquier dato, exportar TODAS las hojas a CSV o hacer una copia del spreadsheet completo.',
    riesgo: 'Sin backup no hay marcha atrás.',
    comoHacerlo: 'En Google Sheets: Archivo → Hacer una copia → "BACKUP_ANTES_LIMPIEZA_' + _fechaHoy_() + '"'
  });

  pasos.push({
    numero: 2,
    titulo: 'ELIMINAR CLIENTES DE PRUEBA (sin actividad asociada)',
    descripcion: 'Verificar primero que los clientes de prueba no tienen trabajos, facturas ni presupuestos.',
    riesgo: 'Bajo si se verifica la ausencia de registros hijos.',
    comoHacerlo: 'Para cada cliente de prueba: buscar su ID en TRABAJOS (col C), FACTURAS (col E), PROFORMAS (col E). Solo borrar si no aparece en ninguna.'
  });

  pasos.push({
    numero: 3,
    titulo: 'ELIMINAR FILAS DUPLICADAS EXACTAS EN TRABAJOS',
    descripcion: 'Para cada ID de trabajo duplicado con filas idénticas, conservar la primera fila y eliminar las siguientes.',
    riesgo: 'Bajo para duplicados exactos.',
    comoHacerlo: 'Ordenar TRABAJOS por ID (col A). Para cada grupo de filas con el mismo ID y contenido idéntico: eliminar todas menos la primera (fila con número menor).'
  });

  pasos.push({
    numero: 4,
    titulo: 'REVISAR DUPLICADOS PARCIALES EN TRABAJOS',
    descripcion: 'Para cada trabajo con el mismo ID pero contenido diferente, decidir manualmente cuál conservar.',
    riesgo: 'Alto — requiere criterio humano.',
    comoHacerlo: 'Comparar campo a campo. Conservar la fila con: (a) ID_Factura relleno, (b) Estado más avanzado, (c) Precio más completo. Eliminar la otra.'
  });

  pasos.push({
    numero: 5,
    titulo: 'RESOLVER REGISTROS HUÉRFANOS',
    descripcion: 'Para cada trabajo/factura/proforma con ID_Cliente inválido, asignar el cliente correcto o eliminar el registro si es claramente de prueba.',
    riesgo: 'Medio.',
    comoHacerlo: 'Revisar el nombre del cliente en el registro huérfano y buscar si ese cliente existe con otro ID. Si existe, actualizar el ID. Si no existe y es de prueba, eliminar.'
  });

  pasos.push({
    numero: 6,
    titulo: 'ELIMINAR FACTURAS Y PROFORMAS DE PRUEBA',
    descripcion: 'Borrar facturas/proformas de clientes de prueba. Verificar que no tienen PDFs en Drive.',
    riesgo: 'Bajo si los clientes son de prueba.',
    comoHacerlo: 'Para cada factura de prueba: verificar URL_PDF (col S). Si hay PDF en Drive, borrarlo también desde Drive antes de eliminar la fila.'
  });

  pasos.push({
    numero: 7,
    titulo: 'LIMPIAR GASTOS DE EJEMPLO',
    descripcion: 'Eliminar gastos claramente de prueba (proveedor "Ejemplo", concepto "Test", etc.).',
    riesgo: 'Bajo.',
    comoHacerlo: 'Identificar por ID (GAS-...) los gastos con proveedor/concepto de ejemplo. Eliminar fila en GASTOS.'
  });

  pasos.push({
    numero: 8,
    titulo: 'IMPLEMENTAR PROTECCIONES ANTI-DUPLICADO (DESARROLLO)',
    descripcion: 'Tras la limpieza, implementar las correcciones de código para evitar nuevos duplicados.',
    riesgo: 'Sin esta fase, los duplicados volverán a aparecer.',
    comoHacerlo: bugs.filter(function(b) { return b.prioridad === 'ALTA'; })
      .map(function(b) { return '• ' + b.titulo + ' → ' + b.archivo; }).join('\n')
  });

  pasos.push({
    numero: 9,
    titulo: 'VERIFICACIÓN FINAL',
    descripcion: 'Tras la limpieza, ejecutar de nuevo la Auditoría (Fase 1) para confirmar que no quedan datos de prueba ni duplicados.',
    riesgo: 'Ninguno — solo lectura.',
    comoHacerlo: 'Menú → 🔍 AUDITORÍA — Datos reales vs. prueba → verificar que todos los contadores de prueba y duplicados son 0.'
  });

  return pasos;
}

// ============================================================
// GENERACIÓN DEL INFORME AUDITORIA_FASE2
// ============================================================

function _f2GenerarInforme_(datos) {
  var ss = getTargetSS_();
  var sh = ss.getSheetByName('AUDITORIA_FASE2');
  if (sh) { sh.clear(); sh.clearFormats(); }
  else    { sh = ss.insertSheet('AUDITORIA_FASE2'); }

  // Anchos de columna
  sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 80);
  sh.setColumnWidth(3, 120); sh.setColumnWidth(4, 120);
  sh.setColumnWidth(5, 100); sh.setColumnWidth(6, 90);
  sh.setColumnWidth(7, 380);

  var NC    = 7;
  var tz    = datos.tz;
  var ahora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
  var r     = 1;

  // ── TÍTULO ────────────────────────────────────────────────────────
  _f2Titulo_(sh, r, NC, '🔬 AUDITORÍA FASE 2 — INFORME EJECUTIVO DETALLADO'); r++;
  _f2Subtitulo_(sh, r, NC, 'Generado el ' + ahora + '   |   ANÁLISIS DE SOLO LECTURA — NINGÚN DATO MODIFICADO'); r++;
  r++;

  // ── Ejecutar todos los análisis ──────────────────────────────────
  var anClientes    = _f2AnalizarClientes_(datos);
  var integridad    = _f2VerificarIntegridad_(datos, anClientes);
  var anTrabajos    = _f2AnalizarTrabajos_(datos);
  var anFacturas    = _f2AnalizarFacturas_(datos, integridad.clientesIds);
  var anProformas   = _f2AnalizarProformas_(datos, integridad.clientesIds);
  var anPresupuestos= _f2AnalizarPresupuestos_(datos, integridad.nombresClientes);
  var anGastos      = _f2AnalizarGastos_(datos);
  var bugs          = _f2AnalizarBugs_(anTrabajos);
  var semaforo      = _f2GenerarSemaforo_(anClientes, anTrabajos, anFacturas, anProformas, anPresupuestos, anGastos, integridad);
  var plan          = _f2PlanLimpieza_(semaforo, anTrabajos, bugs);

  // ============================================================
  // SECCIÓN 1: CLIENTES
  // ============================================================
  _f2Seccion_(sh, r, NC, '👥 1. ANÁLISIS DE CLIENTES'); r++;

  // Resumen numérico
  _f2MiniResumen_(sh, r, NC, [
    ['Total clientes', anClientes.resumen.total],
    ['✅ Reales', anClientes.resumen.reales],
    ['🧪 Prueba', anClientes.resumen.prueba],
    ['❓ Dudosos', anClientes.resumen.dudosos],
    ['⚠ Duplicados cruzados', anClientes.resumen.duplicadosCruzados]
  ]); r += 6;

  // Clientes reales
  _f2SeccionSub_(sh, r, NC, '✅ Clientes REALES', '#d4edda', '#155724'); r++;
  if (anClientes.reales.length === 0) {
    _f2Nota_(sh, r, NC, 'No se detectaron clientes reales (sin email ni CIF válidos).', '#c5221f'); r++;
  } else {
    _f2Cabecera_(sh, r, ['Fila', 'ID', 'Nombre', 'Empresa', 'Email', 'CIF', 'Motivo clasificación']); r++;
    anClientes.reales.forEach(function(c, i) {
      _f2Fila_(sh, r, i, [c.fila, c.id, c.nombre, c.empresa, c.email, c.cif,
        'Email válido / CIF válido / Sin indicadores de prueba']); r++;
    });
  }
  r++;

  // Clientes de prueba
  _f2SeccionSub_(sh, r, NC, '🧪 Clientes de PRUEBA', '#f8d7da', '#721c24'); r++;
  if (anClientes.prueba.length === 0) {
    _f2Nota_(sh, r, NC, '✅ No se detectaron clientes de prueba.', '#137333'); r++;
  } else {
    _f2Cabecera_(sh, r, ['Fila', 'ID', 'Nombre', 'Empresa', 'Email', 'CIF', 'Por qué es prueba']); r++;
    anClientes.prueba.forEach(function(c, i) {
      _f2FilaColor_(sh, r, i, [c.fila, c.id, c.nombre, c.empresa, c.email, c.cif, c.motivo], '#fff3e0'); r++;
    });
  }
  r++;

  // Clientes dudosos
  _f2SeccionSub_(sh, r, NC, '❓ Clientes DUDOSOS', '#fff3cd', '#856404'); r++;
  if (anClientes.dudosos.length === 0) {
    _f2Nota_(sh, r, NC, '✅ No se detectaron clientes dudosos.', '#137333'); r++;
  } else {
    _f2Cabecera_(sh, r, ['Fila', 'ID', 'Nombre', 'Empresa', 'Email', 'CIF', 'Motivo duda']); r++;
    anClientes.dudosos.forEach(function(c, i) {
      _f2FilaColor_(sh, r, i, [c.fila, c.id, c.nombre, c.empresa, c.email, c.cif, c.motivo], '#fffde7'); r++;
    });
  }
  r++;

  // Duplicados cruzados
  _f2SeccionSub_(sh, r, NC, '⚠️ Duplicados cruzados de clientes (mismo email / tel / nombre / CIF)', '#fff3cd', '#856404'); r++;
  if (anClientes.duplicadosCruzados.length === 0) {
    _f2Nota_(sh, r, NC, '✅ No se detectaron clientes con datos duplicados.', '#137333'); r++;
  } else {
    _f2Cabecera_(sh, r, ['Campo', 'Valor compartido', 'Filas afectadas', 'Riesgo', '', '', '']); r++;
    anClientes.duplicadosCruzados.forEach(function(d, i) {
      _f2FilaColor_(sh, r, i, [d.campo, d.valor, 'Filas: ' + d.filas.join(', '), d.riesgo, '', '', ''],
        d.riesgo === 'ALTO' ? '#ffebee' : '#fffde7'); r++;
    });
  }
  r++;

  // ============================================================
  // SECCIÓN 2: TRABAJOS
  // ============================================================
  _f2Seccion_(sh, r, NC, '🔧 2. ANÁLISIS DE TRABAJOS'); r++;

  _f2MiniResumen_(sh, r, NC, [
    ['Total trabajos', anTrabajos.resumen.total],
    ['✅ Reales', anTrabajos.resumen.reales],
    ['🧪 Prueba', anTrabajos.resumen.prueba],
    ['❓ Dudosos', anTrabajos.resumen.dudosos],
    ['⚠ IDs con duplicados', anTrabajos.resumen.idsConDuplicados]
  ]); r += 6;

  // Análisis detallado de cada ID duplicado
  _f2SeccionSub_(sh, r, NC, '🔍 Análisis detallado de IDs duplicados', '#fff3cd', '#856404'); r++;

  if (anTrabajos.analisisDuplicados.length === 0) {
    _f2Nota_(sh, r, NC, '✅ No se detectaron IDs duplicados en TRABAJOS.', '#137333'); r++;
  } else {
    anTrabajos.analisisDuplicados.forEach(function(d) {
      // Cabecera del grupo
      var bgGrupo = d.riesgo === 'CRÍTICO' ? '#f8d7da' : (d.riesgo === 'ALTO' ? '#fff3cd' : '#e8f5e9');
      sh.getRange(r, 1, 1, NC).merge()
        .setValue('▶ ID: ' + d.id + '   Filas: ' + d.filas.join(', ')
          + '   Tipo: ' + d.tipo + '   Riesgo: ' + d.riesgo)
        .setBackground(bgGrupo).setFontWeight('bold').setFontSize(10).setWrap(true);
      sh.setRowHeight(r, 28); r++;

      // Diferencias campo a campo
      if (d.diferenciasCampo.length > 0) {
        sh.getRange(r, 1).setValue('Campos con diferencias:').setFontWeight('bold').setBackground(bgGrupo);
        sh.getRange(r, 2, 1, NC-1).merge()
          .setValue(d.diferenciasCampo.map(function(c) {
            return c.campo + ': [' + c.valores.join(' | ') + ']';
          }).join('   •   '))
          .setBackground(bgGrupo).setFontSize(9).setWrap(true);
        sh.setRowHeight(r, 35); r++;
      } else {
        sh.getRange(r, 1, 1, NC).merge()
          .setValue('→ Las filas son 100% idénticas en todos los campos.')
          .setBackground(bgGrupo).setFontStyle('italic'); r++;
      }

      // Diferencia de timestamps
      var tsStr = '—';
      var tsValidos = d.timestamps.filter(function(t) { return t !== null; });
      if (tsValidos.length >= 2) {
        tsValidos.sort(function(a,b){return a-b;});
        var diffS = Math.round((tsValidos[tsValidos.length-1] - tsValidos[0]) / 1000);
        tsStr = 'Diferencia entre timestamps: ' + diffS + ' segundos';
      }
      sh.getRange(r, 1, 1, NC).merge()
        .setValue('⏱ ' + tsStr)
        .setBackground(bgGrupo).setFontSize(9).setFontColor('#5f6368'); r++;

      // Recomendación
      sh.getRange(r, 1, 1, NC).merge()
        .setValue('💡 ' + d.recomendacion)
        .setBackground('#f0fff4').setFontColor('#155724').setWrap(true).setFontSize(10);
      sh.setRowHeight(r, 50); r++;
      r++;
    });
  }

  // ============================================================
  // SECCIÓN 3: FACTURAS
  // ============================================================
  _f2Seccion_(sh, r, NC, '🧾 3. ANÁLISIS DE FACTURAS'); r++;
  _f2MiniResumen_(sh, r, NC, [
    ['Total facturas', anFacturas.resumen.total],
    ['✅ Reales', anFacturas.resumen.reales],
    ['🧪 Prueba', anFacturas.resumen.prueba],
    ['⚠ Huérfanas (sin cliente)', anFacturas.resumen.huerfanas],
    ['🔴 IDs duplicados', anFacturas.resumen.duplicadosId],
    ['🔴 Números duplicados', anFacturas.resumen.duplicadosNum]
  ]); r += 7;

  _f2Cabecera_(sh, r, ['Fila', 'ID Factura', 'Serie/Nº', 'Cliente', 'Total', 'Estado', 'Clasificación / Notas']); r++;
  var todasFacturas = (anFacturas.reales || []).concat(anFacturas.prueba || []).concat(anFacturas.dudosos || []);
  todasFacturas.sort(function(a,b){return a.fila-b.fila;});
  todasFacturas.forEach(function(f, i) {
    var bg = f.clasificacion === 'PRUEBA' ? '#fff3e0' : (f.clasificacion === 'DUDOSO' ? '#fffde7' : '#f0fff4');
    var nota = (f.clasificacion === 'REAL' ? '✅ Real' : (f.clasificacion === 'PRUEBA' ? '🧪 Prueba: ' + f.motivo : '❓ Dudoso'));
    if (f.huerfana) nota += ' ⚠ SIN CLIENTE';
    _f2FilaColor_(sh, r, i, [f.fila, f.id, f.serie + '-' + f.numero, f.nomCli,
      typeof f.total === 'number' ? f.total.toFixed(2) + ' €' : (f.total || '—'),
      f.estado, nota], bg); r++;
  });

  // Consecuencias de borrar
  r++;
  sh.getRange(r, 1, 1, NC).merge()
    .setValue('⚠️ CONSECUENCIAS DE BORRAR FACTURAS:')
    .setBackground('#e8f0fe').setFontWeight('bold').setFontColor('#1a73e8'); r++;
  [
    ['Factura REAL',    'Pérdida del registro fiscal. Puede generar discrepancias en la contabilidad y problemas con Hacienda si fue emitida.'],
    ['Factura PRUEBA',  'Sin consecuencias fiscales. Verificar que el PDF también se elimina de Drive para no ocupar espacio.'],
    ['Factura HUÉRFANA','Puede ser real aunque su cliente fue eliminado. Revisar el nombre del cliente antes de decidir.'],
    ['Número duplicado','Si tiene número fiscal repetido, la contabilidad queda invalidada. Siempre corregir antes de borrar.']
  ].forEach(function(item, i) {
    var bg = i % 2 === 0 ? '#f8f9fa' : '#ffffff';
    sh.getRange(r, 1, 1, NC).setBackground(bg);
    sh.getRange(r, 1).setValue(item[0]).setFontWeight('bold');
    sh.getRange(r, 2, 1, 6).merge().setValue(item[1]).setWrap(true).setFontSize(10);
    sh.setRowHeight(r, 35); r++;
  });
  r++;

  // ============================================================
  // SECCIÓN 4: PROFORMAS
  // ============================================================
  _f2Seccion_(sh, r, NC, '📋 4. ANÁLISIS DE PROFORMAS'); r++;
  _f2MiniResumen_(sh, r, NC, [
    ['Total proformas', anProformas.resumen.total],
    ['✅ Reales', anProformas.resumen.reales],
    ['🧪 Prueba', anProformas.resumen.prueba],
    ['⚠ Huérfanas', anProformas.resumen.huerfanas],
    ['🔴 IDs duplicados', anProformas.resumen.duplicadosId]
  ]); r += 6;

  if (anProformas.prueba.length > 0) {
    _f2Cabecera_(sh, r, ['Fila', 'ID', 'Cliente', 'Total', 'Estado', 'Clasificación', 'Motivo']); r++;
    anProformas.prueba.forEach(function(p, i) {
      _f2FilaColor_(sh, r, i, [p.fila, p.id, p.nomCli,
        typeof p.total === 'number' ? p.total.toFixed(2) + ' €' : (p.total || '—'),
        p.estado, '🧪 PRUEBA', p.motivo], '#fff3e0'); r++;
    });
    r++;
  }

  // ============================================================
  // SECCIÓN 5: PRESUPUESTOS
  // ============================================================
  _f2Seccion_(sh, r, NC, '📝 5. ANÁLISIS DE PRESUPUESTOS'); r++;
  _f2MiniResumen_(sh, r, NC, [
    ['Total presupuestos', anPresupuestos.resumen.total],
    ['✅ Reales', anPresupuestos.resumen.reales],
    ['🧪 Prueba', anPresupuestos.resumen.prueba],
    ['⚠ Cliente no en CLIENTES', anPresupuestos.resumen.sinClienteEnBBDD],
    ['🔴 IDs duplicados', anPresupuestos.resumen.duplicadosId]
  ]); r += 6;

  if (anPresupuestos.prueba.length > 0) {
    _f2Cabecera_(sh, r, ['Fila', 'ID', 'Cliente', 'Total', 'Estado', 'Motivo', '']); r++;
    anPresupuestos.prueba.forEach(function(p, i) {
      _f2FilaColor_(sh, r, i, [p.fila, p.id, p.cliente,
        typeof p.total === 'number' ? p.total.toFixed(2) + ' €' : (p.total || '—'),
        p.estado, p.motivo, ''], '#fff3e0'); r++;
    });
    r++;
  }

  // ============================================================
  // SECCIÓN 6: GASTOS
  // ============================================================
  _f2Seccion_(sh, r, NC, '💰 6. ANÁLISIS DE GASTOS'); r++;
  _f2MiniResumen_(sh, r, NC, [
    ['Total gastos', anGastos.resumen.total],
    ['✅ Reales', anGastos.resumen.reales],
    ['🧪 Prueba', anGastos.resumen.prueba],
    ['❓ Dudosos', anGastos.resumen.dudosos],
    ['Sin archivo adjunto', anGastos.resumen.sinArchivo],
    ['Con texto OCR', anGastos.resumen.conOcr],
    ['OCR baja confianza (<60%)', anGastos.resumen.ocrIncompleto],
    ['Sin tipo clasificado', anGastos.resumen.sinTipo],
    ['Pendientes de pago', anGastos.resumen.pendientesPago]
  ]); r += 10;

  if (anGastos.sinArchivo.length > 0) {
    _f2SeccionSub_(sh, r, NC, '📎 Gastos sin archivo adjunto (' + anGastos.sinArchivo.length + ')', '#fff3cd', '#856404'); r++;
    _f2Cabecera_(sh, r, ['Fila', 'ID', 'Proveedor', 'Concepto', 'Total', 'Clasificación', 'Acción recomendada']); r++;
    anGastos.sinArchivo.slice(0, 30).forEach(function(g, i) {
      var clasif = _f2ClasificarRegistro_({ nombre: g.proveedor + ' ' + g.concepto, id: g.id });
      _f2FilaColor_(sh, r, i, [g.fila, g.id, g.proveedor, g.concepto,
        typeof g.total === 'number' ? g.total.toFixed(2) + ' €' : (g.total || '—'),
        clasif.tipo,
        clasif.tipo === 'PRUEBA' ? 'Eliminar' : 'Adjuntar factura o ticket'], '#fffde7'); r++;
    });
    if (anGastos.sinArchivo.length > 30) {
      _f2Nota_(sh, r, NC, '… y ' + (anGastos.sinArchivo.length - 30) + ' más sin mostrar.', '#5f6368'); r++;
    }
    r++;
  }

  if (anGastos.ocrIncompleto.length > 0) {
    _f2SeccionSub_(sh, r, NC, '🔍 Gastos con OCR de baja confianza (' + anGastos.ocrIncompleto.length + ')', '#fff3cd', '#856404'); r++;
    _f2Cabecera_(sh, r, ['Fila', 'ID', 'Confianza OCR', '', '', '', 'Recomendación']); r++;
    anGastos.ocrIncompleto.forEach(function(g, i) {
      _f2FilaColor_(sh, r, i, [g.fila, g.id,
        typeof g.confianza === 'number' ? (g.confianza * 100).toFixed(0) + '%' : g.confianza,
        '', '', '', 'Revisar manualmente el documento adjunto y completar los datos'], '#fff8e1'); r++;
    });
    r++;
  }

  // ============================================================
  // SECCIÓN 7: INTEGRIDAD REFERENCIAL
  // ============================================================
  _f2Seccion_(sh, r, NC, '🔗 7. INTEGRIDAD REFERENCIAL'); r++;

  if (integridad.roturas.length === 0) {
    _f2Nota_(sh, r, NC, '✅ No se detectaron roturas de integridad referencial. Todas las relaciones son correctas.', '#137333'); r++;
  } else {
    _f2Cabecera_(sh, r, ['Tipo rotura', 'Hoja', 'Fila', 'ID registro', 'Referencia rota', 'Nombre ref.', 'Impacto']); r++;
    integridad.roturas.forEach(function(rot, i) {
      var bg = rot.impacto.indexOf('CRÍTICO') !== -1 ? '#ffebee' : '#fff3e0';
      _f2FilaColor_(sh, r, i, [rot.tipo, rot.hoja, rot.fila, rot.id,
        rot.refRota, rot.nombre, rot.impacto], bg); r++;
    });
  }
  r++;

  // ============================================================
  // SECCIÓN 8: ANÁLISIS DE BUGS
  // ============================================================
  _f2Seccion_(sh, r, NC, '🐛 8. ANÁLISIS DE BUGS DE PROGRAMACIÓN'); r++;

  if (bugs.length === 0) {
    _f2Nota_(sh, r, NC, '✅ No se detectaron patrones de bug de programación.', '#137333'); r++;
  } else {
    bugs.forEach(function(bug, i) {
      var bgH = bug.prioridad === 'ALTA' ? '#f8d7da' : '#fff3cd';
      var fgH = bug.prioridad === 'ALTA' ? '#721c24' : '#856404';

      sh.getRange(r, 1, 1, NC).merge()
        .setValue('🐛 Bug ' + (i+1) + ': ' + bug.titulo + '   [Prioridad: ' + bug.prioridad + ']')
        .setBackground(bgH).setFontColor(fgH).setFontWeight('bold'); r++;

      var detalles = [
        ['Archivo/Función:', bug.archivo],
        ['Descripción:', bug.descripcion],
        ['Evidencia en datos:', bug.evidencia],
        ['Solución recomendada:', bug.solucion]
      ];
      detalles.forEach(function(det) {
        sh.getRange(r, 1, 1, NC).setBackground(bgH.replace('f8d7da','ffebee').replace('fff3cd','fffde7'));
        sh.getRange(r, 1).setValue(det[0]).setFontWeight('bold').setFontSize(9);
        sh.getRange(r, 2, 1, 6).merge().setValue(det[1]).setWrap(true).setFontSize(9);
        sh.setRowHeight(r, 40); r++;
      });

      // Indicaciones técnicas
      var tecnicas = [];
      if (bug.requiereLock) tecnicas.push('✅ Requiere LockService.getScriptLock()');
      if (bug.titulo.indexOf('UUID') !== -1 || bug.solucion.indexOf('UUID') !== -1) tecnicas.push('💡 Alternativa: UUID en lugar de ID secuencial');
      if (tecnicas.length > 0) {
        sh.getRange(r, 1, 1, NC).merge()
          .setValue(tecnicas.join('   •   '))
          .setBackground('#e8f0fe').setFontColor('#1a73e8').setFontSize(9); r++;
      }
      r++;
    });

    // Tabla resumen de qué implementar
    _f2SeccionSub_(sh, r, NC, '🛡️ Protecciones recomendadas a implementar', '#e8f0fe', '#1a73e8'); r++;
    var protecciones = [
      { nombre: 'LockService.getScriptLock()',  donde: 'nextId() en Utils.gs',                  prioridad: '🔴 ALTA',  detalle: 'Evita que dos ejecuciones simultáneas generen el mismo ID' },
      { nombre: 'Idempotencia por eventId',      donde: 'onFormSubmitHandler() en Triggers.gs',  prioridad: '🔴 ALTA',  detalle: 'Almacenar eventId procesados en ScriptProperties para evitar retries' },
      { nombre: 'Límite de 1 respuesta',         donde: 'Configuración del Google Form',         prioridad: '🟡 MEDIA', detalle: 'Activar "Limitar a 1 respuesta" en el formulario de trabajos' },
      { nombre: 'Validación de ID existente',    donde: 'Todas las funciones de alta',           prioridad: '🔴 ALTA',  detalle: 'Antes de appendRow(), verificar que el ID no existe ya' },
      { nombre: 'Deduplicación por hash',        donde: 'onTrabajosFormSubmit()',                prioridad: '🟡 MEDIA', detalle: 'Hash de (cliente+descripción+fecha) → rechazar si ya existe en las últimas 24h' },
      { nombre: 'UUID como ID primario',         donde: 'nextId() en Utils.gs — alternativa',   prioridad: '🟢 BAJA',  detalle: 'Utilities.getUuid() garantiza unicidad sin necesidad de locking' }
    ];
    _f2Cabecera_(sh, r, ['Protección', 'Dónde implementar', 'Prioridad', 'Qué evita', '', '', '']); r++;
    protecciones.forEach(function(p, i) {
      var bg = i % 2 === 0 ? '#f8f9fa' : '#ffffff';
      sh.getRange(r, 1, 1, NC).setBackground(bg);
      sh.getRange(r, 1).setValue(p.nombre).setFontWeight('bold').setFontSize(9);
      sh.getRange(r, 2).setValue(p.donde).setFontSize(9);
      sh.getRange(r, 3).setValue(p.prioridad).setFontSize(9);
      sh.getRange(r, 4, 1, 4).merge().setValue(p.detalle).setWrap(true).setFontSize(9);
      sh.setRowHeight(r, 35); r++;
    });
  }
  r++;

  // ============================================================
  // SECCIÓN 9: TABLA SEMÁFORO
  // ============================================================
  _f2Seccion_(sh, r, NC, '🚦 9. TABLA SEMÁFORO — DECISIÓN POR REGISTRO'); r++;

  // VERDE
  _f2SeccionSub_(sh, r, NC,
    '🟢 ELIMINAR CON SEGURIDAD (' + semaforo.verde.length + ' registros)', '#d4edda', '#155724'); r++;
  if (semaforo.verde.length === 0) {
    _f2Nota_(sh, r, NC, 'No hay registros seguros para eliminar.', '#5f6368'); r++;
  } else {
    _f2Cabecera_(sh, r, ['Hoja', 'Fila', 'ID', 'Motivo', 'Impacto de borrar', '', '']); r++;
    semaforo.verde.forEach(function(item, i) {
      sh.getRange(r, 1, 1, NC).setBackground(i % 2 === 0 ? '#f0fff4' : '#e8f5e9');
      sh.getRange(r, 1).setValue(item.hoja);
      sh.getRange(r, 2).setValue(item.fila || '—').setHorizontalAlignment('center');
      sh.getRange(r, 3).setValue(item.id || '—');
      sh.getRange(r, 4, 1, 2).merge().setValue(item.motivo).setWrap(true).setFontSize(9);
      sh.getRange(r, 6, 1, 2).merge().setValue(item.impacto).setWrap(true).setFontSize(9).setFontColor('#137333');
      sh.setRowHeight(r, 35); r++;
    });
  }
  r++;

  // AMARILLO
  _f2SeccionSub_(sh, r, NC,
    '🟡 REVISAR MANUALMENTE (' + semaforo.amarillo.length + ' registros)', '#fff3cd', '#856404'); r++;
  if (semaforo.amarillo.length === 0) {
    _f2Nota_(sh, r, NC, 'No hay registros que requieran revisión manual.', '#5f6368'); r++;
  } else {
    _f2Cabecera_(sh, r, ['Hoja', 'Fila', 'ID', 'Motivo', 'Impacto', '', '']); r++;
    semaforo.amarillo.forEach(function(item, i) {
      sh.getRange(r, 1, 1, NC).setBackground(i % 2 === 0 ? '#fffde7' : '#fff9c4');
      sh.getRange(r, 1).setValue(item.hoja);
      sh.getRange(r, 2).setValue(item.fila || '—').setHorizontalAlignment('center');
      sh.getRange(r, 3).setValue(item.id || '—');
      sh.getRange(r, 4, 1, 2).merge().setValue(item.motivo).setWrap(true).setFontSize(9);
      sh.getRange(r, 6, 1, 2).merge().setValue(item.impacto).setWrap(true).setFontSize(9).setFontColor('#856404');
      sh.setRowHeight(r, 40); r++;
    });
  }
  r++;

  // ROJO
  _f2SeccionSub_(sh, r, NC,
    '🔴 NO TOCAR (' + semaforo.rojo.length + ' registros)', '#f8d7da', '#721c24'); r++;
  if (semaforo.rojo.length === 0) {
    _f2Nota_(sh, r, NC, 'No hay registros críticos.', '#5f6368'); r++;
  } else {
    _f2Cabecera_(sh, r, ['Hoja', 'Fila', 'ID', 'Motivo', 'Impacto de borrar', '', '']); r++;
    semaforo.rojo.forEach(function(item, i) {
      sh.getRange(r, 1, 1, NC).setBackground(i % 2 === 0 ? '#ffebee' : '#fce4ec');
      sh.getRange(r, 1).setValue(item.hoja);
      sh.getRange(r, 2).setValue(item.fila || '—').setHorizontalAlignment('center');
      sh.getRange(r, 3).setValue(item.id || '—');
      sh.getRange(r, 4, 1, 2).merge().setValue(item.motivo).setWrap(true).setFontSize(9);
      sh.getRange(r, 6, 1, 2).merge().setValue(item.impacto).setWrap(true).setFontSize(9).setFontColor('#c5221f');
      sh.setRowHeight(r, 35); r++;
    });
  }
  r++;

  // ============================================================
  // SECCIÓN 10: PLAN DE LIMPIEZA PASO A PASO
  // ============================================================
  _f2Seccion_(sh, r, NC, '📋 10. PLAN DE LIMPIEZA RECOMENDADO — PASO A PASO'); r++;

  sh.getRange(r, 1, 1, NC).merge()
    .setValue('⛔ IMPORTANTE: Este plan es informativo. No ejecutar ninguna acción sin hacer backup previo.')
    .setBackground('#f8d7da').setFontColor('#721c24').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(r, 30); r++;
  r++;

  plan.forEach(function(paso) {
    var bg = paso.numero === 1 ? '#f8d7da' : (paso.numero === 8 ? '#e8f0fe' : '#e8f5e9');
    var fg = paso.numero === 1 ? '#721c24' : (paso.numero === 8 ? '#1a73e8' : '#155724');

    sh.getRange(r, 1, 1, NC).merge()
      .setValue('PASO ' + paso.numero + ': ' + paso.titulo)
      .setBackground(bg).setFontColor(fg).setFontWeight('bold').setFontSize(11);
    sh.setRowHeight(r, 30); r++;

    var detallesPaso = [
      ['Qué hacer:', paso.descripcion],
      ['⚠ Riesgo:', paso.riesgo],
      ['Cómo hacerlo:', paso.comoHacerlo]
    ];
    detallesPaso.forEach(function(det) {
      sh.getRange(r, 1, 1, NC).setBackground(bg.replace('f8d7da','fff5f5').replace('e8f0fe','f0f4ff').replace('e8f5e9','f0fff4'));
      sh.getRange(r, 1).setValue(det[0]).setFontWeight('bold').setFontSize(9).setFontColor('#5f6368');
      sh.getRange(r, 2, 1, 6).merge().setValue(det[1]).setWrap(true).setFontSize(9);
      sh.setRowHeight(r, Math.max(30, Math.ceil(det[1].length / 80) * 18)); r++;
    });
    r++;
  });

  // PIE
  sh.getRange(r, 1, 1, NC).merge()
    .setValue('AUDITORIA_FASE2 — Generado el ' + ahora + ' — Sistema de Gestión Álvaro')
    .setFontColor('#9aa0a6').setFontSize(9).setHorizontalAlignment('center');

  SpreadsheetApp.flush();
  Logger.log('AUDITORIA_FASE2 generada: ' + r + ' filas, ' + ahora);
}

// ============================================================
// HELPERS
// ============================================================

function _s(v) { return String(v || '').trim(); }

function _unique_(arr) {
  var seen = {};
  return arr.filter(function(v) { var k = String(v); if (seen[k]) return false; seen[k] = true; return true; });
}

function _fechaHoy_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function _esCifPrueba_(cif) {
  return ['A00000000','B00000000','00000000A','00000000Z',
    'B12345678','A12345678','12345678A','11111111H'].indexOf(cif) !== -1;
}

function _f2ClasificarRegistro_(campos) {
  var texto = [
    campos.nombre || '', campos.empresa || '',
    campos.email  || '', campos.cif     || '',
    campos.dir    || '', campos.id      || ''
  ].join(' ').toLowerCase().trim();

  if (!texto || texto.length < 2) return { tipo: 'DUDOSO', motivo: 'Sin datos identificativos' };

  var palabrasPrueba = [
    'test','prueba','ejemplo','demo','sample','ficticio','falso','dummy','fake',
    'temporal','borrar','eliminar','mock','xxx','zzz','aaa','cliente prueba',
    'cliente test'
  ];
  for (var i = 0; i < palabrasPrueba.length; i++) {
    if (texto.indexOf(palabrasPrueba[i]) !== -1)
      return { tipo: 'PRUEBA', motivo: 'Contiene palabra clave: "' + palabrasPrueba[i] + '"' };
  }

  if (campos.email) {
    var emailL = campos.email.toLowerCase();
    var dominiosPrueba = ['@test.com','@example.com','@prueba.com','@demo.com','@fake.com'];
    for (var j = 0; j < dominiosPrueba.length; j++) {
      if (emailL.indexOf(dominiosPrueba[j]) !== -1)
        return { tipo: 'PRUEBA', motivo: 'Email de prueba: ' + campos.email };
    }
  }

  if (campos.cif) {
    var cifU = campos.cif.toUpperCase().replace(/\s/g, '');
    if (_esCifPrueba_(cifU))
      return { tipo: 'PRUEBA', motivo: 'CIF por defecto del sistema: ' + cifU };
  }

  if (campos.dir) {
    var dirL = campos.dir.toLowerCase();
    var dirsPrueba = ['calle ejemplo','calle test','calle prueba','c/ ejemplo'];
    for (var k = 0; k < dirsPrueba.length; k++) {
      if (dirL.indexOf(dirsPrueba[k]) !== -1)
        return { tipo: 'PRUEBA', motivo: 'Dirección de ejemplo: ' + campos.dir };
    }
  }

  if (campos.nombre && campos.nombre.toLowerCase().indexOf('mi empresa') !== -1)
    return { tipo: 'PRUEBA', motivo: 'Nombre de empresa por defecto del sistema' };

  if (!campos.email && !campos.cif)
    return { tipo: 'DUDOSO', motivo: 'Sin email ni CIF — no se puede confirmar' };

  return { tipo: 'REAL', motivo: 'Datos de contacto válidos presentes' };
}

// ---- Helpers de formato ----

function _f2Titulo_(sh, r, nc, txt) {
  sh.getRange(r, 1, 1, nc).merge()
    .setValue(txt).setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(r, 55);
}

function _f2Subtitulo_(sh, r, nc, txt) {
  sh.getRange(r, 1, 1, nc).merge()
    .setValue(txt).setBackground('#e8f0fe').setFontColor('#5f6368')
    .setFontSize(10).setHorizontalAlignment('center');
}

function _f2Seccion_(sh, r, nc, txt) {
  sh.getRange(r, 1, 1, nc).merge()
    .setValue(txt).setBackground('#202124').setFontColor('#ffffff')
    .setFontSize(13).setFontWeight('bold').setHorizontalAlignment('left');
  sh.setRowHeight(r, 35);
}

function _f2SeccionSub_(sh, r, nc, txt, bg, fg) {
  sh.getRange(r, 1, 1, nc).merge()
    .setValue(txt).setBackground(bg || '#e8f0fe').setFontColor(fg || '#1a73e8')
    .setFontSize(11).setFontWeight('bold');
  sh.setRowHeight(r, 28);
}

function _f2MiniResumen_(sh, r, nc, items) {
  items.forEach(function(item, i) {
    var bg = i % 2 === 0 ? '#f8f9fa' : '#ffffff';
    sh.getRange(r + i, 1, 1, nc).setBackground(bg);
    sh.getRange(r + i, 1).setValue(item[0]).setFontWeight('bold').setFontSize(10);
    sh.getRange(r + i, 2).setValue(item[1]).setHorizontalAlignment('center').setFontSize(10);
  });
}

function _f2Cabecera_(sh, r, hdrs) {
  sh.getRange(r, 1, 1, hdrs.length).setValues([hdrs])
    .setBackground('#37474f').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(9);
}

function _f2Fila_(sh, r, idx, valores) {
  var bg = idx % 2 === 0 ? '#ffffff' : '#f8f9fa';
  sh.getRange(r, 1, 1, NC_F2 || 7).setBackground(bg);
  valores.forEach(function(v, c) {
    sh.getRange(r, c + 1).setValue(v !== undefined && v !== null ? v : '').setFontSize(9);
  });
}

function _f2FilaColor_(sh, r, idx, valores, bg) {
  var fondo = bg || (idx % 2 === 0 ? '#ffffff' : '#f8f9fa');
  sh.getRange(r, 1, 1, valores.length).setBackground(fondo);
  valores.forEach(function(v, c) {
    sh.getRange(r, c + 1).setValue(v !== undefined && v !== null ? v : '').setFontSize(9).setWrap(true);
  });
}

function _f2Nota_(sh, r, nc, txt, color) {
  sh.getRange(r, 1, 1, nc).merge()
    .setValue(txt).setFontColor(color || '#5f6368')
    .setFontStyle('italic').setHorizontalAlignment('center').setFontSize(10);
}

var NC_F2 = 7;
