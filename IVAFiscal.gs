// ============================================================
// IVAFiscal.gs — FASE 9: IVA trimestral y resumen fiscal
// ============================================================

var HEADERS_IVA_TRIM = [
  'Año', 'Trimestre', 'Fecha_Desde', 'Fecha_Hasta',
  'Base_Facturas', 'IVA_Repercutido', 'Total_Facturas',
  'Base_Gastos', 'IVA_Soportado', 'Total_Gastos',
  'IVA_Neto', 'Resultado',
  'Nº_Facturas', 'Nº_Gastos', 'Nº_Gastos_Revisar',
  'Nº_Gastos_Sin_Ticket', 'Fecha_Generacion'
];

// FACTURAS column indices (base 1) — coinciden con SHEET_HEADERS.FACTURAS en Config.gs
var CFAC = {
  FECHA_EMISION:  2,
  BASE_IMPONIBLE: 12,
  IVA_IMPORTE:    14,
  TOTAL:          17,
  ESTADO:         18
};
var N_COLS_FAC = 23; // total columnas FACTURAS

// IVA_TRIMESTRAL column indices (base 1)
var CIVA = {
  ANYO:                1,
  TRIMESTRE:           2,
  FECHA_DESDE:         3,
  FECHA_HASTA:         4,
  BASE_FACTURAS:       5,
  IVA_REPERCUTIDO:     6,
  TOTAL_FACTURAS:      7,
  BASE_GASTOS:         8,
  IVA_SOPORTADO:       9,
  TOTAL_GASTOS:        10,
  IVA_NETO:            11,
  RESULTADO:           12,
  N_FACTURAS:          13,
  N_GASTOS:            14,
  N_GASTOS_REVISAR:    15,
  N_GASTOS_SIN_TICKET: 16,
  FECHA_GENERACION:    17
};

// ─────────────────────────────────────────────────────────────
// generarIVATrimestral  — un trimestre
// ─────────────────────────────────────────────────────────────

function generarIVATrimestral() {
  var ui   = SpreadsheetApp.getUi();
  var anyo = new Date().getFullYear();

  var rA = ui.prompt('IVA Trimestral — Año',
    'Introduce el año (ej. ' + anyo + '):', ui.ButtonSet.OK_CANCEL);
  if (rA.getSelectedButton() !== ui.Button.OK) return;
  anyo = parseInt(rA.getResponseText().trim(), 10);
  if (isNaN(anyo) || anyo < 2000 || anyo > 2100) {
    ui.alert('Año no válido.'); return;
  }

  var rT = ui.prompt('IVA Trimestral — Trimestre',
    'Trimestre (1 = Ene-Mar, 2 = Abr-Jun, 3 = Jul-Sep, 4 = Oct-Dic):', ui.ButtonSet.OK_CANCEL);
  if (rT.getSelectedButton() !== ui.Button.OK) return;
  var trim = parseInt(rT.getResponseText().trim(), 10);
  if ([1, 2, 3, 4].indexOf(trim) === -1) {
    ui.alert('Trimestre no válido. Escribe 1, 2, 3 o 4.'); return;
  }

  var sh    = _getOrCreateIVASheet_();
  var datos = _calcularTrimestre_(anyo, trim);
  _upsertFilaTrimestre_(sh, datos);
  _formatearHojaIVA_(sh);
  getTargetSS_().setActiveSheet(sh);

  logEvent('INFO', 'generarIVATrimestral',
    'T' + trim + ' ' + anyo + '  IVA neto=' + datos.ivaNeto + '  ' + datos.resultado, '');

  ui.alert('✅ IVA T' + trim + ' ' + anyo,
    'IVA repercutido (OUT):  ' + datos.ivaOut.toFixed(2)   + ' €\n'
    + 'IVA soportado  (IN):   ' + datos.ivaIn.toFixed(2)    + ' €\n'
    + 'IVA neto:              ' + datos.ivaNeto.toFixed(2)  + ' €\n\n'
    + 'Resultado: ' + _iconoResultado_(datos.resultado) + ' ' + datos.resultado + '\n\n'
    + 'Facturas: ' + datos.nFac + '  |  Gastos: ' + datos.nGas
    + (datos.nGasRevisar   > 0 ? '  |  ⚠️ Revisar: '    + datos.nGasRevisar   : '')
    + (datos.nGasSinTicket > 0 ? '  |  ⚠️ Sin ticket: ' + datos.nGasSinTicket : ''),
    ui.ButtonSet.OK);
}

// ─────────────────────────────────────────────────────────────
// generarIVAAnual  — los 4 trimestres del año
// ─────────────────────────────────────────────────────────────

function generarIVAAnual() {
  var ui   = SpreadsheetApp.getUi();
  var anyo = new Date().getFullYear();

  var rA = ui.prompt('IVA Anual — Año',
    'Introduce el año (ej. ' + anyo + '):', ui.ButtonSet.OK_CANCEL);
  if (rA.getSelectedButton() !== ui.Button.OK) return;
  anyo = parseInt(rA.getResponseText().trim(), 10);
  if (isNaN(anyo) || anyo < 2000 || anyo > 2100) {
    ui.alert('Año no válido.'); return;
  }

  var sh      = _getOrCreateIVASheet_();
  var lineas  = [];
  var ivaAnual = 0;

  for (var t = 1; t <= 4; t++) {
    var datos = _calcularTrimestre_(anyo, t);
    _upsertFilaTrimestre_(sh, datos);
    ivaAnual += datos.ivaNeto;
    lineas.push('T' + t + ': OUT ' + datos.ivaOut.toFixed(2) + ' € — IN '
      + datos.ivaIn.toFixed(2) + ' €  →  ' + _iconoResultado_(datos.resultado)
      + ' ' + datos.resultado + ' (' + datos.ivaNeto.toFixed(2) + ' €)');
  }

  _formatearHojaIVA_(sh);
  getTargetSS_().setActiveSheet(sh);

  logEvent('INFO', 'generarIVAAnual', 'Año ' + anyo + '  IVA neto anual=' + ivaAnual.toFixed(2), '');

  var resAnual = ivaAnual > 0.005 ? 'A pagar' : ivaAnual < -0.005 ? 'A compensar' : 'Sin diferencia';
  ui.alert('✅ IVA Anual ' + anyo,
    lineas.join('\n') + '\n\n'
    + 'TOTAL ANUAL: ' + _iconoResultado_(resAnual) + ' '
    + ivaAnual.toFixed(2) + ' € — ' + resAnual,
    ui.ButtonSet.OK);
}

// ─────────────────────────────────────────────────────────────
// _calcularTrimestre_  — núcleo del cálculo
// ─────────────────────────────────────────────────────────────

function _calcularTrimestre_(anyo, trimestre) {
  Logger.log('[_calcularTrimestre_] T' + trimestre + ' ' + anyo);
  var tz        = Session.getScriptTimeZone();
  var mesInicio = (trimestre - 1) * 3;
  var desde     = new Date(anyo, mesInicio, 1, 0, 0, 0);
  var hasta     = new Date(anyo, mesInicio + 3, 0, 23, 59, 59);

  Logger.log('  Rango: ' + Utilities.formatDate(desde, tz, 'dd/MM/yyyy')
             + ' – ' + Utilities.formatDate(hasta, tz, 'dd/MM/yyyy'));

  // ── IVA OUT: FACTURAS ─────────────────────────────────────────
  var baseOut = 0, ivaOut = 0, totalOut = 0, nFac = 0;
  try {
    var shFac   = getSheet('FACTURAS');
    var lastFac = shFac.getLastRow();
    if (lastFac > 1) {
      var facData = shFac.getRange(2, 1, lastFac - 1, N_COLS_FAC).getValues();
      facData.forEach(function(row) {
        // Ignorar anuladas
        var estado = String(row[CFAC.ESTADO - 1]).trim().toLowerCase();
        if (estado === 'anulada') return;

        var fEm = row[CFAC.FECHA_EMISION - 1];
        if (!(fEm instanceof Date) || isNaN(fEm.getTime())) return;
        if (fEm < desde || fEm > hasta) return;

        baseOut  += parseFloat(row[CFAC.BASE_IMPONIBLE - 1]) || 0;
        ivaOut   += parseFloat(row[CFAC.IVA_IMPORTE    - 1]) || 0;
        totalOut += parseFloat(row[CFAC.TOTAL           - 1]) || 0;
        nFac++;
      });
    }
  } catch (e) {
    logEvent('WARN', '_calcularTrimestre_', 'Error leyendo FACTURAS: ' + e.message, '');
  }
  Logger.log('  OUT → n=' + nFac + '  base=' + baseOut.toFixed(2)
             + '  IVA=' + ivaOut.toFixed(2) + '  total=' + totalOut.toFixed(2));

  // ── IVA IN: GASTOS ────────────────────────────────────────────
  var baseIn = 0, ivaIn = 0, totalIn = 0, nGas = 0;
  var nGasRevisar = 0, nGasSinTicket = 0;
  try {
    var shGas   = getSheet('GASTOS');
    var lastGas = shGas.getLastRow();
    if (lastGas > 1) {
      var gasData = shGas.getRange(2, 1, lastGas - 1, HEADERS_GASTOS.length).getValues();
      gasData.forEach(function(row) {
        if (String(row[COL_GAS.ID_GASTO - 1]).trim().indexOf('GAS-') !== 0) return;

        // Fecha: prefiere FECHA_GASTO, fallback FECHA_ALTA
        var fGas = row[COL_GAS.FECHA_GASTO - 1];
        var fRef = (fGas instanceof Date && !isNaN(fGas.getTime()))
                   ? fGas : row[COL_GAS.FECHA_ALTA - 1];
        // Si tiene fecha válida, filtrar por período; si no, excluir del cálculo IVA
        if (fRef instanceof Date && !isNaN(fRef.getTime())) {
          if (fRef < desde || fRef > hasta) return;
        } else {
          return; // sin fecha → no sabemos a qué trimestre pertenece
        }

        baseIn  += parseFloat(row[COL_GAS.BASE_IMPONIBLE - 1]) || 0;
        ivaIn   += parseFloat(row[COL_GAS.IVA_IMPORTE    - 1]) || 0;
        totalIn += parseFloat(row[COL_GAS.TOTAL           - 1]) || 0;
        nGas++;

        // Contadores de calidad de datos
        var ep = String(row[COL_GAS.ESTADO_PROCESAMIENTO - 1]).trim();
        if (ep === 'OCR fallido' || ep === 'Revisar') nGasRevisar++;

        var foto  = String(row[COL_GAS.FOTO_URL    - 1]).trim();
        var arch  = String(row[COL_GAS.ARCHIVO_ID  - 1]).trim();
        if (!foto && !arch) nGasSinTicket++;
      });
    }
  } catch (e) {
    logEvent('WARN', '_calcularTrimestre_', 'Error leyendo GASTOS: ' + e.message, '');
  }
  Logger.log('  IN  → n=' + nGas + '  base=' + baseIn.toFixed(2)
             + '  IVA=' + ivaIn.toFixed(2) + '  total=' + totalIn.toFixed(2));
  Logger.log('  Revisar/fallido=' + nGasRevisar + '  Sin ticket=' + nGasSinTicket);

  // ── IVA Neto ──────────────────────────────────────────────────
  var ivaNeto   = Math.round((ivaOut - ivaIn) * 100) / 100;
  var resultado = ivaNeto > 0.005 ? 'A pagar'
                : ivaNeto < -0.005 ? 'A compensar'
                : 'Sin diferencia';
  Logger.log('  IVA neto=' + ivaNeto + '  → ' + resultado);

  return {
    anyo:            anyo,
    trimestre:       trimestre,
    desde:           desde,
    hasta:           hasta,
    baseOut:         Math.round(baseOut  * 100) / 100,
    ivaOut:          Math.round(ivaOut   * 100) / 100,
    totalOut:        Math.round(totalOut * 100) / 100,
    baseIn:          Math.round(baseIn   * 100) / 100,
    ivaIn:           Math.round(ivaIn    * 100) / 100,
    totalIn:         Math.round(totalIn  * 100) / 100,
    ivaNeto:         ivaNeto,
    resultado:       resultado,
    nFac:            nFac,
    nGas:            nGas,
    nGasRevisar:     nGasRevisar,
    nGasSinTicket:   nGasSinTicket,
    fechaGen:        new Date()
  };
}

// ─────────────────────────────────────────────────────────────
// _upsertFilaTrimestre_  — escribe o actualiza la fila
// ─────────────────────────────────────────────────────────────

function _upsertFilaTrimestre_(sh, d) {
  var tz      = Session.getScriptTimeZone();
  var lastRow = sh.getLastRow();

  // Buscar fila existente para ese año + trimestre
  var targetRow = -1;
  if (lastRow > 1) {
    var cols = sh.getRange(2, CIVA.ANYO, lastRow - 1, 2).getValues();
    for (var i = 0; i < cols.length; i++) {
      if (Number(cols[i][0]) === d.anyo && Number(cols[i][1]) === d.trimestre) {
        targetRow = i + 2; break;
      }
    }
  }
  if (targetRow === -1) targetRow = lastRow + 1;

  var rowValues = [
    d.anyo,
    'T' + d.trimestre,
    Utilities.formatDate(d.desde,   tz, 'dd/MM/yyyy'),
    Utilities.formatDate(d.hasta,   tz, 'dd/MM/yyyy'),
    d.baseOut, d.ivaOut, d.totalOut,
    d.baseIn,  d.ivaIn,  d.totalIn,
    d.ivaNeto,
    d.resultado,
    d.nFac, d.nGas, d.nGasRevisar, d.nGasSinTicket,
    Utilities.formatDate(d.fechaGen, tz, 'dd/MM/yyyy HH:mm')
  ];

  sh.getRange(targetRow, 1, 1, HEADERS_IVA_TRIM.length).setValues([rowValues]);

  // Formato numérico en columnas de importe
  var FMT = '#,##0.00';
  [CIVA.BASE_FACTURAS, CIVA.IVA_REPERCUTIDO, CIVA.TOTAL_FACTURAS,
   CIVA.BASE_GASTOS,   CIVA.IVA_SOPORTADO,   CIVA.TOTAL_GASTOS,
   CIVA.IVA_NETO].forEach(function(c) {
     sh.getRange(targetRow, c).setNumberFormat(FMT);
   });

  // Color de fila según resultado
  var bgRow  = d.resultado === 'A pagar'      ? '#fce8e6'
             : d.resultado === 'A compensar'  ? '#e6f4ea'
             : '#f8f9fa';
  var fgNeto = d.resultado === 'A pagar'      ? '#c5221f'
             : d.resultado === 'A compensar'  ? '#137333'
             : '#5f6368';

  sh.getRange(targetRow, 1, 1, HEADERS_IVA_TRIM.length).setBackground(bgRow);
  sh.getRange(targetRow, CIVA.IVA_NETO)
    .setFontColor(fgNeto).setFontWeight('bold');
  sh.getRange(targetRow, CIVA.RESULTADO)
    .setFontColor(fgNeto).setFontWeight('bold');

  // Advertencia visual si hay gastos pendientes de revisión
  if (d.nGasRevisar > 0 || d.nGasSinTicket > 0) {
    var nota = [];
    if (d.nGasRevisar   > 0) nota.push(d.nGasRevisar   + ' gastos por revisar');
    if (d.nGasSinTicket > 0) nota.push(d.nGasSinTicket + ' sin ticket');
    sh.getRange(targetRow, CIVA.N_GASTOS_REVISAR)
      .setBackground('#fce8b2').setFontColor('#b06000');
    sh.getRange(targetRow, CIVA.N_GASTOS_SIN_TICKET)
      .setBackground('#fce8b2').setFontColor('#b06000');
    Logger.log('  ⚠️ ' + nota.join(' | '));
  }
}

// ─────────────────────────────────────────────────────────────
// _getOrCreateIVASheet_  — obtiene o crea la hoja
// ─────────────────────────────────────────────────────────────

function _getOrCreateIVASheet_() {
  var ss = getTargetSS_();
  var sh = ss.getSheetByName('IVA_TRIMESTRAL');
  if (!sh) {
    sh = ss.insertSheet('IVA_TRIMESTRAL');
  }
  // Escribir/confirmar cabecera
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS_IVA_TRIM);
    applyHeaderStyle_(sh, HEADERS_IVA_TRIM.length);
  }
  return sh;
}

// ─────────────────────────────────────────────────────────────
// _formatearHojaIVA_  — anchos y filas fijas
// ─────────────────────────────────────────────────────────────

function _formatearHojaIVA_(sh) {
  var widths = [55, 75, 90, 90, 110, 110, 110, 110, 110, 110, 110, 110, 80, 75, 100, 105, 130];
  widths.forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(1);
  SpreadsheetApp.flush();
}

// ─────────────────────────────────────────────────────────────
// _iconoResultado_  — emoji para el alert
// ─────────────────────────────────────────────────────────────

function _iconoResultado_(resultado) {
  return resultado === 'A pagar'     ? '💸'
       : resultado === 'A compensar' ? '💚'
       : '⚪';
}
