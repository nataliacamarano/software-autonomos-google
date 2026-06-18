// ============================================================
// FuenteMaestra.gs — Fuente única de clientes activos
//
// ARQUITECTURA:
//   CLIENTES!J (checkbox) → CLIENTES_ACTIVOS (FILTER vivo)
//                         → VALUE_IN_RANGE en todos los desplegables
//
// Una sola fuente. Sin snapshots. Sin listas hardcodeadas.
// ============================================================

var FM_ = {
  HOJA:          'CLIENTES_ACTIVOS',
  RANGO:         'A2:A50',
  DIAG:          'DIAGNOSTICO_FUENTE_CLIENTES',
  INFORME:       'CLIENTES_MAESTRO_RECONSTRUIDO',

  // Fragmentos que identifican a los 3 clientes reales (minúsculas, sin tildes)
  REALES:        ['gregor', 'pougeoux', 'fouss'],

  // Fragmentos de clientes ELIMINADOS — para detectar validaciones viejas
  ANTIGUOS:      ['sofia', 'fernandez ruiz', 'carlos ruiz', 'cristina vega',
                  'david moreno', 'juan martinez', 'ana f'],

  // Coordenadas explícitas de selectores de cliente
  COORDS: [
    { hoja: 'PANEL_FACTURAS', row: 5,  col: 3,  desc: 'C5 — Cliente principal'           },
    { hoja: 'PANEL_FACTURAS', row: 5,  col: 6,  desc: 'F5 — Cliente presupuesto'          },
    { hoja: 'PANEL_FACTURAS', row: 99, col: 3,  desc: 'C99 — Factura desde Presupuesto'   },
    { hoja: 'PANEL_CLIENTE',  row: 4,  col: 3,  desc: 'C4 — Selector panel cliente'       }
  ]
};

// ── Punto de entrada ─────────────────────────────────────────

/**
 * Orquestador principal.
 * 1. Fuerza Activo=FALSE para todos los no-reales en CLIENTES.
 * 2. Crea/actualiza CLIENTES_ACTIVOS con fórmula FILTER viva.
 * 3. Escanea todas las hojas buscando validaciones de cliente antiguas.
 * 4. Reemplaza TODAS por VALUE_IN_RANGE → CLIENTES_ACTIVOS.
 * 5. Genera DIAGNOSTICO_FUENTE_CLIENTES y CLIENTES_MAESTRO_RECONSTRUIDO.
 */
function ejecutarReparacionMaestraClientes() {
  var ui  = SpreadsheetApp.getUi();
  var ss  = getTargetSS_();
  var log = [];

  try {
    // ── Paso 1: Forzar Activo correcto en CLIENTES ──────────────
    _fm_forzarActivoClientes_(ss, log);

    // ── Paso 2: Crear/actualizar hoja maestra ───────────────────
    _fm_crearHojaMaestra_(ss, log);

    // ── Paso 3: Escanear todas las hojas ────────────────────────
    var fuentes = _fm_escanearValidaciones_(ss, log);

    // ── Paso 4: Reemplazar todas las validaciones ────────────────
    var nReparadas = _fm_reemplazarValidaciones_(ss, log);

    // ── Paso 5: Informes ─────────────────────────────────────────
    _fm_crearDiagnostico_(ss, log, fuentes);
    _fm_crearInformeReconstruccion_(ss, log, nReparadas);

    logEvent('INFO', 'ejecutarReparacionMaestraClientes',
      nReparadas + ' validaciones reemplazadas | fuente única: ' + FM_.HOJA, '');

    ui.alert(
      '✅ Fuente maestra establecida',
      'Hoja maestra: ' + FM_.HOJA + '\n'
      + 'Validaciones reemplazadas: ' + nReparadas + '\n\n'
      + 'Todos los desplegables apuntan ahora a ' + FM_.HOJA + '.\n'
      + 'Cuando cambies un checkbox en CLIENTES!J,\n'
      + 'los desplegables se actualizan solos.\n\n'
      + 'Ver:\n'
      + '  • ' + FM_.DIAG + '\n'
      + '  • ' + FM_.INFORME,
      ui.ButtonSet.OK
    );

  } catch (e) {
    logEvent('ERROR', 'ejecutarReparacionMaestraClientes', e.message, e.stack);
    ui.alert('❌ Error:\n' + e.message + '\n\nVer LOGS para detalle.');
  }
}

// ── Helper público: validación apuntando a la fuente maestra ─

/**
 * Devuelve una DataValidation VALUE_IN_RANGE → CLIENTES_ACTIVOS.
 * Usar esta función en lugar de requireValueInList en todo el sistema.
 */
function getValidacionClienteMaestra_(ss) {
  try {
    var s = (ss || getTargetSS_()).getSheetByName(FM_.HOJA);
    if (!s) _fm_crearHojaMaestra_(ss || getTargetSS_(), []);
    s = (ss || getTargetSS_()).getSheetByName(FM_.HOJA);
    if (!s) return null;
    return SpreadsheetApp.newDataValidation()
      .requireValueInRange(s.getRange(FM_.RANGO), true)
      .setAllowInvalid(false)
      .build();
  } catch (e) {
    return null;
  }
}

// ── Paso 1: Forzar Activo en CLIENTES ────────────────────────

function _fm_forzarActivoClientes_(ss, log) {
  var sh = ss.getSheetByName('CLIENTES');
  if (!sh || sh.getLastRow() < 2) {
    log.push(['WARN', 'CLIENTES vacía', '']);
    return;
  }

  var nRows = sh.getLastRow() - 1;
  var data  = sh.getRange(2, 1, nRows, COL_CLI.METODOS_PAGO).getValues();

  function norm(t) {
    return String(t || '').toLowerCase()
      .replace(/[áàäâ]/g,'a').replace(/[éèëê]/g,'e')
      .replace(/[íìïî]/g,'i').replace(/[óòöô]/g,'o')
      .replace(/[úùüû]/g,'u').trim();
  }
  function esReal(row) {
    var txt = norm(row[COL_CLI.NOMBRE - 1]) + ' ' +
              norm(row[COL_CLI.APELLIDOS - 1]) + ' ' +
              norm(row[COL_CLI.EMPRESA - 1]);
    return FM_.REALES.some(function(k) { return txt.indexOf(k) !== -1; });
  }

  var nuevosActivos = data.map(function(row, i) {
    var nombreCompleto = (String(row[COL_CLI.NOMBRE - 1] || '') + ' ' +
                          String(row[COL_CLI.APELLIDOS - 1] || '')).trim()
                      || String(row[COL_CLI.EMPRESA - 1] || '') || 'fila ' + (i + 2);
    var esR = esReal(row);
    log.push([esR ? 'REAL→TRUE' : 'OTRO→FALSE',
              'Fila ' + (i + 2) + ': "' + nombreCompleto + '"',
              'Activo anterior: ' + String(row[COL_CLI.ACTIVO - 1])]);
    return [esR];   // real=true, todos los demás=false (sin excepciones)
  });

  var rangoJ = sh.getRange(2, COL_CLI.ACTIVO, nRows, 1);
  rangoJ.setDataValidation(null);
  rangoJ.clearFormat();
  rangoJ.insertCheckboxes();
  rangoJ.setValues(nuevosActivos);

  var nActivos = nuevosActivos.filter(function(v) { return v[0]; }).length;
  log.push(['OK', 'CLIENTES!J reconstruida: ' + nActivos + ' activos, ' + (nRows - nActivos) + ' inactivos', '']);
}

// ── Paso 2: Crear CLIENTES_ACTIVOS ───────────────────────────

function _fm_crearHojaMaestra_(ss, log) {
  var sh = ss.getSheetByName(FM_.HOJA);
  if (!sh) {
    sh = ss.insertSheet(FM_.HOJA);
  } else {
    sh.clearContents();
    sh.clearFormats();
  }

  // Cabecera
  sh.getRange('A1').setValue('Nombre_Cliente')
    .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');

  // Fórmula FILTER viva — se actualiza sola cuando cambia CLIENTES!J
  // Formato: "Nombre Apellidos" o "Nombre Apellidos (Empresa)"
  sh.getRange('A2').setFormula(
    '=IFERROR(FILTER(' +
      'TRIM(CLIENTES!C2:C&" "&CLIENTES!D2:D)' +
      '&IF(CLIENTES!E2:E<>"","  ("&CLIENTES!E2:E&")",""),' +
      '(CLIENTES!J2:J=TRUE)*(CLIENTES!C2:C<>""))' +
    ',"")'
  );

  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 220);

  if (log) log.push(['OK', FM_.HOJA + ' creada con fórmula FILTER viva',
    'Rango de validación: ' + FM_.HOJA + '!' + FM_.RANGO]);
}

// ── Paso 3: Escanear validaciones en todas las hojas ─────────

function _fm_escanearValidaciones_(ss, log) {
  var hojas    = ss.getSheets();
  var fuentes  = [];
  var fragmentosDeteccion = FM_.REALES.concat(FM_.ANTIGUOS);

  function norm(t) { return String(t||'').toLowerCase().replace(/[áéíóúàèìòùäëïöü]/g, function(c){ return 'aeiouaeiouaeiou'['áéíóúàèìòùäëïöü'.indexOf(c)] || c; }).trim(); }

  hojas.forEach(function(sh) {
    var nombre   = sh.getName();
    var lastRow  = Math.min(sh.getLastRow(),   150);
    var lastCol  = Math.min(sh.getLastColumn(), 20);
    if (lastRow < 1 || lastCol < 1) return;

    try {
      var dvs  = sh.getRange(1, 1, lastRow, lastCol).getDataValidations();
      var vals = sh.getRange(1, 1, lastRow, lastCol).getValues();

      for (var r = 0; r < dvs.length; r++) {
        for (var c = 0; c < dvs[r].length; c++) {
          var dv = dvs[r][c];
          if (!dv) continue;

          var row  = r + 1;
          var col  = c + 1;
          var tipo = dv.getCriteriaType();
          var cv   = dv.getCriteriaValues();

          var lista = [];
          var rangoRef = '';
          var esRango  = false;

          if (tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
            lista = (cv && Array.isArray(cv[0])) ? cv[0] : [];
          } else if (tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
            esRango  = true;
            rangoRef = cv && cv[0] ? cv[0].getSheet().getName() + '!' + cv[0].getA1Notation() : '?';
          } else {
            continue;  // otros tipos no son de interés
          }

          // Detección: ¿es un selector de cliente?
          var esCoordConocida = FM_.COORDS.some(function(coord) {
            return coord.hoja === nombre && coord.row === row && coord.col === col;
          });
          var etiqIzq = (c > 0) ? norm(vals[r][c-1]) : '';
          var tieneEtiq = etiqIzq.indexOf('cliente') !== -1;

          var coincideAntiguo = false;
          var coincideActual  = false;
          lista.forEach(function(v) {
            var vn = norm(v);
            if (FM_.ANTIGUOS.some(function(k){ return vn.indexOf(k) !== -1; })) coincideAntiguo = true;
            if (FM_.REALES.some(function(k){ return vn.indexOf(k) !== -1; })) coincideActual = true;
          });

          var esCliente = esCoordConocida || tieneEtiq || coincideAntiguo || coincideActual
                       || (esRango && rangoRef.toLowerCase().indexOf('clientes') !== -1);

          if (!esCliente) continue;

          var info = {
            hoja:     nombre,
            row:      row, col: col,
            celda:    _col2letter_(col) + row,
            tipo:     esRango ? 'VALUE_IN_RANGE → ' + rangoRef : 'VALUE_IN_LIST (' + lista.length + ' items)',
            lista:    lista,
            rangoRef: rangoRef,
            esRango:  esRango,
            esCoord:  esCoordConocida,
            antiguo:  coincideAntiguo,
            actual:   coincideActual
          };
          fuentes.push(info);

          var resumen = esRango
            ? 'RANGO: ' + rangoRef
            : lista.length + ' items: ' + lista.slice(0, 5).join(', ') + (lista.length > 5 ? '…' : '');

          log.push([
            coincideAntiguo ? 'ANTIGUA' : (esCoordConocida ? 'COORD' : 'DETECTADA'),
            nombre + ' ' + _col2letter_(col) + row + (esCoordConocida ? ' [COORD CONOCIDA]' : ''),
            resumen
          ]);
        }
      }
    } catch (e2) {
      log.push(['WARN', 'Error escaneando hoja "' + nombre + '"', e2.message]);
    }
  });

  log.push(['INFO', 'Total validaciones de cliente encontradas: ' + fuentes.length, '']);
  return fuentes;
}

// ── Paso 4: Reemplazar todas las validaciones ─────────────────

function _fm_reemplazarValidaciones_(ss, log) {
  var nReparadas = 0;
  var dvMaestra  = getValidacionClienteMaestra_(ss);
  if (!dvMaestra) {
    log.push(['ERROR', 'No se pudo construir la validación maestra', '']);
    return 0;
  }

  // 4a. Coordenadas explícitas — siempre se actualizan
  FM_.COORDS.forEach(function(coord) {
    var sh = ss.getSheetByName(coord.hoja);
    if (!sh) { log.push(['SKIP', coord.hoja + ' no encontrada', '']); return; }
    try {
      var rng = sh.getRange(coord.row, coord.col);
      var valorActual = rng.getValue();
      rng.setDataValidation(null);
      rng.setDataValidation(dvMaestra);
      log.push(['REPARADA', coord.hoja + ' ' + coord.desc,
        'Valor actual: "' + valorActual + '" | VALUE_IN_RANGE → ' + FM_.HOJA + '!' + FM_.RANGO]);
      nReparadas++;
    } catch (e2) {
      log.push(['ERROR', coord.hoja + ' ' + coord.desc, e2.message]);
    }
  });

  // 4b. Escanear y reemplazar cualquier otra validación de cliente detectada
  var hojas   = ss.getSheets();
  var coordSet = {};
  FM_.COORDS.forEach(function(c) { coordSet[c.hoja + '|' + c.row + '|' + c.col] = true; });

  function norm(t) { return String(t||'').toLowerCase().trim(); }

  hojas.forEach(function(sh) {
    var nombre  = sh.getName();
    // No tocar hojas de informes propias ni LOGS ni CONFIG
    if ([FM_.DIAG, FM_.INFORME, FM_.HOJA, 'LOGS', 'CONFIG'].indexOf(nombre) !== -1) return;

    var lastRow = Math.min(sh.getLastRow(),   150);
    var lastCol = Math.min(sh.getLastColumn(), 20);
    if (lastRow < 1 || lastCol < 1) return;

    try {
      var dvs  = sh.getRange(1, 1, lastRow, lastCol).getDataValidations();
      var vals = sh.getRange(1, 1, lastRow, lastCol).getValues();

      for (var r = 0; r < dvs.length; r++) {
        for (var c = 0; c < dvs[r].length; c++) {
          var dv = dvs[r][c];
          if (!dv) continue;

          var row  = r + 1;
          var col  = c + 1;
          var key  = nombre + '|' + row + '|' + col;
          if (coordSet[key]) continue; // ya procesada en 4a

          var tipo = dv.getCriteriaType();
          var cv   = dv.getCriteriaValues();
          if (tipo !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) continue;

          var lista = (cv && Array.isArray(cv[0])) ? cv[0] : [];
          var tieneAntiguo = FM_.ANTIGUOS.concat(FM_.REALES).some(function(k) {
            return lista.some(function(v) { return norm(v).indexOf(k) !== -1; });
          });
          var etiqIzq = (c > 0) ? norm(vals[r][c-1]) : '';
          var tieneEtiq = etiqIzq.indexOf('cliente') !== -1;

          if (!tieneAntiguo && !tieneEtiq) continue;

          try {
            var rng = sh.getRange(row, col);
            rng.setDataValidation(null);
            rng.setDataValidation(dvMaestra);
            log.push(['REPARADA', nombre + ' ' + _col2letter_(col) + row,
              'VALUE_IN_LIST(' + lista.length + ') → VALUE_IN_RANGE ' + FM_.HOJA]);
            nReparadas++;
          } catch (e3) {
            log.push(['ERROR', nombre + ' ' + _col2letter_(col) + row, e3.message]);
          }
        }
      }
    } catch (e2) {
      log.push(['WARN', 'Error procesando "' + nombre + '"', e2.message]);
    }
  });

  log.push(['OK', 'Total validaciones reemplazadas: ' + nReparadas,
    'Todas apuntan ahora a ' + FM_.HOJA + '!' + FM_.RANGO]);
  return nReparadas;
}

// ── Informes ──────────────────────────────────────────────────

function _fm_crearDiagnostico_(ss, log, fuentes) {
  var sh = ss.getSheetByName(FM_.DIAG);
  if (sh) { sh.clearContents(); sh.clearFormats(); }
  else     { sh = ss.insertSheet(FM_.DIAG); }

  sh.getRange(1,1,1,5).setValues([['Estado','Hoja','Celda','Tipo / Lista','Timestamp']]);
  sh.getRange(1,1,1,5).setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
  sh.setFrozenRows(1);

  // Sección: fuentes encontradas antes del reemplazo
  if (fuentes && fuentes.length) {
    fuentes.forEach(function(f) {
      var estado = f.antiguo ? 'ANTIGUA' : (f.esRango ? 'RANGO_EXT' : 'DETECTADA');
      var resumen = f.esRango ? 'RANGO → ' + f.rangoRef
        : f.lista.slice(0,8).join(' | ') + (f.lista.length > 8 ? ' …(' + f.lista.length + ')' : '');
      sh.appendRow([estado, f.hoja, f.celda, resumen, new Date()]);
    });
  } else {
    sh.appendRow(['INFO', '—', '—', 'No se encontraron validaciones de cliente antiguas', new Date()]);
  }

  sh.appendRow(['','','','','']);
  sh.appendRow(['LOG','Hoja/Función','Detalle','','Timestamp']);
  log.forEach(function(e) { sh.appendRow([e[0], e[1], e[2] || '', '', new Date()]); });

  _fm_colorearHoja_(sh);
  sh.autoResizeColumns(1, 5);
}

function _fm_crearInformeReconstruccion_(ss, log, nReparadas) {
  var sh = ss.getSheetByName(FM_.INFORME);
  if (sh) { sh.clearContents(); sh.clearFormats(); }
  else     { sh = ss.insertSheet(FM_.INFORME); }

  var maestraHoja = ss.getSheetByName(FM_.HOJA);
  var clientesActivos = [];
  if (maestraHoja) {
    var vals = maestraHoja.getRange(FM_.RANGO).getValues();
    vals.forEach(function(r) { if (String(r[0]).trim()) clientesActivos.push(String(r[0]).trim()); });
  }

  sh.getRange(1,1,1,3).setValues([['Sección','Detalle','Timestamp']]);
  sh.getRange(1,1,1,3).setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
  sh.setFrozenRows(1);

  sh.appendRow(['FUENTE MAESTRA', FM_.HOJA + '!' + FM_.RANGO + ' (fórmula FILTER viva)', new Date()]);
  sh.appendRow(['CLIENTES ACTIVOS', clientesActivos.length + ' → ' + clientesActivos.join(' | '), new Date()]);
  sh.appendRow(['VALIDACIONES', nReparadas + ' celdas actualizadas a VALUE_IN_RANGE', new Date()]);
  sh.appendRow(['','','']);

  var verificacion = [
    { frag: 'gregor',   label: 'Gregor'          },
    { frag: 'pougeoux', label: 'Patrick Pougeoux' },
    { frag: 'fouss',    label: 'Jose Fouss'       }
  ];
  verificacion.forEach(function(v) {
    function n(t) { return String(t||'').toLowerCase().replace(/[áéíóúàèìòùäëïöü]/g,'x'); }
    var encontrado = clientesActivos.find(function(c) { return n(c).indexOf(v.frag) !== -1; });
    sh.appendRow([
      encontrado ? '✓ PASS' : '✗ FAIL',
      v.label + (encontrado ? ' → "' + encontrado + '"' : ' → NO ENCONTRADO'),
      new Date()
    ]);
  });

  sh.appendRow(['','','']);
  FM_.COORDS.forEach(function(coord) {
    var pfSh = ss.getSheetByName(coord.hoja);
    if (!pfSh) return;
    try {
      var dv = pfSh.getRange(coord.row, coord.col).getDataValidation();
      var tipo = dv ? String(dv.getCriteriaType()) : 'SIN VALIDACIÓN';
      sh.appendRow(['CELDA', coord.hoja + ' ' + coord.desc + ' → ' + tipo, new Date()]);
    } catch (_) {}
  });

  _fm_colorearHoja_(sh);
  sh.autoResizeColumns(1, 3);
}

function _fm_colorearHoja_(sh) {
  try {
    var n = sh.getLastRow() - 1;
    if (n < 1) return;
    var colorMap = {
      'REPARADA':'#e6f4ea','OK':'#e6f4ea','✓ PASS':'#e6f4ea','REAL→TRUE':'#e6f4ea',
      'ANTIGUA':'#fce8e6','ERROR':'#fce8e6','✗ FAIL':'#fce8e6','OTRO→FALSE':'#f8f9fa',
      'WARN':'#fef9c3','DETECTADA':'#fff3e0','COORD':'#e8f0fe','RANGO_EXT':'#fff3e0',
      'LOG':'#f1f3f4','INFO':'#f1f3f4','FUENTE MAESTRA':'#e8f0fe','CLIENTES ACTIVOS':'#e8f0fe',
      'VALIDACIONES':'#e8f0fe','CELDA':'#f1f3f4','SKIP':'#f8f9fa'
    };
    sh.getRange(2, 1, n, 1).getValues().forEach(function(row, i) {
      var color = colorMap[String(row[0])] || '#ffffff';
      sh.getRange(i + 2, 1, 1, sh.getLastColumn()).setBackground(color);
    });
  } catch (_) {}
}

// ── Utilidad ──────────────────────────────────────────────────

function _col2letter_(col) {
  var s = '';
  while (col > 0) {
    col--;
    s = String.fromCharCode(65 + (col % 26)) + s;
    col = Math.floor(col / 26);
  }
  return s;
}

// ============================================================
// LIMPIEZA DE VALORES INVÁLIDOS EN CELDAS DE CLIENTE
// ============================================================

/**
 * Escanea todas las celdas con validación de cliente.
 * Si el valor actual NO está en CLIENTES_ACTIVOS → vacía la celda.
 * La validación se mantiene intacta. No toca celdas con valor válido.
 * Genera CLIENTES_INVALIDOS_LIMPIADOS con el detalle.
 */
function limpiarValoresClienteInvalidos() {
  var ui  = SpreadsheetApp.getUi();
  var ss  = getTargetSS_();
  var log = [];
  var nLimpiadas = 0;
  var nValidas   = 0;

  try {
    // ── 1. Leer lista válida desde CLIENTES_ACTIVOS ──────────────
    var maestraSh = ss.getSheetByName(FM_.HOJA);
    if (!maestraSh) {
      ui.alert('❌ Hoja ' + FM_.HOJA + ' no encontrada.\n'
        + 'Ejecuta primero: 🗂 FUENTE MAESTRA — Reparar clientes.');
      return;
    }

    var clientesValidos = [];
    maestraSh.getRange(FM_.RANGO).getValues().forEach(function(r) {
      var v = String(r[0] || '').trim();
      if (v) clientesValidos.push(v);
    });

    if (clientesValidos.length === 0) {
      ui.alert('⚠️ ' + FM_.HOJA + ' está vacía. No hay clientes válidos definidos.');
      return;
    }

    // Set de nombres válidos en minúsculas para comparación rápida
    var validosLower = {};
    clientesValidos.forEach(function(n) { validosLower[n.toLowerCase()] = n; });

    log.push(['INFO', 'Clientes válidos en ' + FM_.HOJA + ': ' + clientesValidos.length,
      clientesValidos.join(' | ')]);

    // ── 2. Escanear todas las hojas ───────────────────────────────
    var HOJAS_EXCLUIDAS = [
      FM_.HOJA, FM_.DIAG, FM_.INFORME,
      'CLIENTES_INVALIDOS_LIMPIADOS', 'LOGS', 'CONFIG',
      'CLIENTES', 'OPERARIOS',  // hojas de datos maestros
      'REPARACION_ACTIVO_CLIENTES', 'REPARACION_PANEL_FACTURAS_CLIENTE',
      'SYNC_CLIENTES_DIAGNOSTICO', 'TRABAJOS_RESTAURADOS',
      'RESTAURACION_DIAGNOSTICO', 'RESTAURACION_REALIZADA'
    ];

    var coordSet = {};
    FM_.COORDS.forEach(function(c) { coordSet[c.hoja + '|' + c.row + '|' + c.col] = true; });

    ss.getSheets().forEach(function(sh) {
      var nombre  = sh.getName();
      if (HOJAS_EXCLUIDAS.indexOf(nombre) !== -1) return;

      var lastRow = Math.min(sh.getLastRow(),    150);
      var lastCol = Math.min(sh.getLastColumn(),  20);
      if (lastRow < 1 || lastCol < 1) return;

      try {
        var dvs  = sh.getRange(1, 1, lastRow, lastCol).getDataValidations();
        var cels = sh.getRange(1, 1, lastRow, lastCol).getValues();

        for (var r = 0; r < dvs.length; r++) {
          for (var c = 0; c < dvs[r].length; c++) {
            var dv = dvs[r][c];
            if (!dv) continue;

            var row  = r + 1;
            var col  = c + 1;
            var tipo = dv.getCriteriaType();
            var cv   = dv.getCriteriaValues();

            // ¿Es validación de cliente?
            var esCliente = coordSet[nombre + '|' + row + '|' + col] || false;

            if (!esCliente && tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
              try {
                if (cv && cv[0] && cv[0].getSheet().getName() === FM_.HOJA) esCliente = true;
              } catch (_) {}
            }

            if (!esCliente && tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
              var lista = (cv && Array.isArray(cv[0])) ? cv[0] : [];
              // Es de cliente si la lista solapa con nombres conocidos (actuales o antiguos)
              var todoFrags = FM_.REALES.concat(FM_.ANTIGUOS);
              esCliente = todoFrags.some(function(frag) {
                return lista.some(function(v) {
                  return String(v).toLowerCase().indexOf(frag) !== -1;
                });
              });
            }

            if (!esCliente) continue;

            // ── 3. Comprobar el valor actual de la celda ──────────
            var valor = String(cels[r][c] || '').trim();
            if (!valor) continue;  // ya vacía — no hace falta tocar

            var esValido = !!validosLower[valor.toLowerCase()];

            if (esValido) {
              log.push(['OK', nombre + ' ' + _col2letter_(col) + row, '"' + valor + '" — válido']);
              nValidas++;
            } else {
              sh.getRange(row, col).clearContent();  // borra valor, mantiene validación
              log.push(['LIMPIADO', nombre + ' ' + _col2letter_(col) + row, '"' + valor + '"']);
              nLimpiadas++;
            }
          }
        }
      } catch (e2) {
        log.push(['WARN', 'Error en hoja "' + nombre + '"', e2.message]);
      }
    });

    // ── 4. Informe ────────────────────────────────────────────────
    _fm_crearInformeLimpieza_(ss, log, clientesValidos, nLimpiadas, nValidas);

    logEvent('INFO', 'limpiarValoresClienteInvalidos',
      nLimpiadas + ' limpiadas | ' + nValidas + ' válidas conservadas', '');

    ui.alert(
      '✅ Limpieza completada',
      'Valores inválidos eliminados: ' + nLimpiadas + '\n'
      + 'Valores válidos conservados: ' + nValidas + '\n\n'
      + 'Clientes válidos:\n'
      + clientesValidos.map(function(n) { return '  • ' + n; }).join('\n') + '\n\n'
      + 'Ver CLIENTES_INVALIDOS_LIMPIADOS para el detalle.',
      ui.ButtonSet.OK
    );

  } catch (e) {
    logEvent('ERROR', 'limpiarValoresClienteInvalidos', e.message, e.stack);
    ui.alert('❌ Error:\n' + e.message);
  }
}

function _fm_crearInformeLimpieza_(ss, log, clientesValidos, nLimpiadas, nValidas) {
  try {
    var shName = 'CLIENTES_INVALIDOS_LIMPIADOS';
    var sh = ss.getSheetByName(shName);
    if (sh) { sh.clearContents(); sh.clearFormats(); }
    else     { sh = ss.insertSheet(shName); }

    sh.getRange(1, 1, 1, 4).setValues([['Estado', 'Hoja + Celda', 'Valor eliminado', 'Timestamp']]);
    sh.getRange(1, 1, 1, 4).setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
    sh.setFrozenRows(1);

    sh.appendRow(['RESUMEN', 'Valores inválidos eliminados: ' + nLimpiadas, '', new Date()]);
    sh.appendRow(['RESUMEN', 'Valores válidos conservados:  ' + nValidas,   '', new Date()]);
    sh.appendRow(['RESUMEN', 'Clientes válidos: ' + clientesValidos.join(' | '), '', new Date()]);
    sh.appendRow(['', '', '', '']);

    log.forEach(function(e) {
      sh.appendRow([e[0], e[1], e[2] || '', new Date()]);
    });

    var nData = sh.getLastRow() - 1;
    if (nData > 0) {
      var colorMap = {
        'LIMPIADO': '#fce8e6',
        'OK':       '#e6f4ea',
        'RESUMEN':  '#e8f0fe',
        'INFO':     '#f1f3f4',
        'WARN':     '#fef9c3'
      };
      sh.getRange(2, 1, nData, 1).getValues().forEach(function(row, i) {
        var color = colorMap[String(row[0])] || '#ffffff';
        sh.getRange(i + 2, 1, 1, 4).setBackground(color);
      });
    }

    sh.autoResizeColumns(1, 4);
  } catch (e) {
    logEvent('WARN', '_fm_crearInformeLimpieza_', e.message, '');
  }
}
