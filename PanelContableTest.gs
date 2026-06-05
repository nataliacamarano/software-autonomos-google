// ============================================================
// PanelContableTest.gs — Test de integración FASE 17
// Ejecutar desde el editor de script o menú ⚙️ Sistema
// ============================================================

/**
 * Test de integración completo para PANEL_CONTABLE FASE 17.
 * Prueba: lectura de datos, carpeta Drive, hojas exportadas,
 * resumen, IVA, CSV y log. NO envía email real.
 */
function testFlujoPanelContable17() {
  var ss   = getTargetSS_();
  var tz   = Session.getScriptTimeZone();
  var now  = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
  var ok   = [];
  var warn = [];
  var err  = [];

  Logger.log('=== TEST PANEL_CONTABLE FASE 17-18 — ' + now + ' ===');

  // ── 1. Panel existe ───────────────────────────────────────────
  var sh = ss.getSheetByName(PCN_SHEET);
  if (!sh) {
    SpreadsheetApp.getUi().alert(
      '⚠ PANEL_CONTABLE no existe.\n\n' +
      'Ejecuta primero:\n  ⚙️ Sistema → 📦 PANEL_CONTABLE — Crear completo\n\n' +
      'Luego vuelve a ejecutar el test.'
    );
    return;
  }
  ok.push('✅ PANEL_CONTABLE existe (' + sh.getLastRow() + ' filas)');

  // ── 2. Tipo de periodo y rango de fechas ──────────────────────
  var tipo    = _pcnLeerTipoPeriodo_(sh);
  var periodo = _pcnPeriodoEtiqueta_(tipo);
  var rango   = _pcnObtenerRango_(sh);
  ok.push('✅ Periodo: ' + tipo + ' → "' + periodo + '"');
  if (rango) {
    ok.push('   Rango: ' +
      Utilities.formatDate(rango.ini, tz, 'dd/MM/yyyy') + ' — ' +
      Utilities.formatDate(rango.fin, tz, 'dd/MM/yyyy'));
  } else {
    ok.push('   Rango: sin filtro de fecha (modo Personalizado sin fechas)');
  }

  // ── 3. Lectura de FACTURAS ────────────────────────────────────
  try {
    var facData = _pcnGetFacturasData_(ss, rango);
    ok.push('✅ _pcnGetFacturasData_: ' + facData.rows.length + ' factura(s) en el periodo' +
            ' (' + facData.headers.length + ' columnas)');
  } catch (e) {
    err.push('❌ _pcnGetFacturasData_: ' + e.message);
  }

  // ── 4. Lectura de GASTOS ──────────────────────────────────────
  try {
    var gasData = _pcnGetGastosData_(ss, rango);
    ok.push('✅ _pcnGetGastosData_: ' + gasData.rows.length + ' gasto(s) en el periodo' +
            ' (' + gasData.headers.length + ' columnas)');
  } catch (e) {
    err.push('❌ _pcnGetGastosData_: ' + e.message);
  }

  // ── 5. KPIs ───────────────────────────────────────────────────
  try {
    var kpis = _pcnActualizarResumen_(ss, sh);
    ok.push('✅ KPIs calculados:');
    ok.push('     Facturas: ' + kpis.facCount + '  Total: ' + _pcnFmtEur_(kpis.facTotal));
    ok.push('     IVA rep.:  ' + _pcnFmtEur_(kpis.ivaRep) +
            '  Pendiente: ' + _pcnFmtEur_(kpis.pendiente));
    ok.push('     Gastos:   ' + kpis.gasCount + '  Total: ' + _pcnFmtEur_(kpis.gasTotal));
    ok.push('     IVA neto: ' + _pcnFmtEur_(kpis.ivaNeto));
  } catch (e) {
    err.push('❌ _pcnActualizarResumen_: ' + e.message);
  }

  // ── 6. Carpeta Drive ──────────────────────────────────────────
  var perFolder = null;
  try {
    perFolder = _pcnResolvePeriodFolder_(sh);
    ok.push('✅ Carpeta Drive: ' + perFolder.getName() + ' (' + perFolder.getUrl() + ')');
  } catch (e) {
    err.push('❌ _pcnResolvePeriodFolder_: ' + e.message);
  }

  if (perFolder && err.length === 0) {
    // ── 7. Crear sheet de prueba en Drive ─────────────────────
    try {
      var testFolder = _pcnGetOrCreateFolder_(perFolder, '_TEST_');
      var testHeaders = ['ID', 'Nombre', 'Importe'];
      var testRows    = [
        ['T001', 'Test 1', 100.50],
        ['T002', 'Test 2', 200.75]
      ];
      var testUrl = _pcnCrearSheetEnDrive_(testFolder, 'Test_' + periodo, testHeaders, testRows);
      ok.push('✅ _pcnCrearSheetEnDrive_: sheet creado en Drive');
      ok.push('   URL: ' + testUrl);

      // Limpiar carpeta de test
      testFolder.setTrashed(true);
      ok.push('   (carpeta _TEST_ eliminada tras el test)');
    } catch (e) {
      err.push('❌ _pcnCrearSheetEnDrive_: ' + e.message);
    }

    // ── 8. Resumen contable (con datos reales de KPIs) ────────
    try {
      var resFolder = _pcnGetOrCreateFolder_(perFolder, 'Resumen');
      var empresa   = _pcnGetEmpresaNombre_();
      if (typeof kpis !== 'undefined') {
        _pcnCrearResumenContable_(resFolder, 'Resumen_' + periodo, periodo, empresa, kpis);
        ok.push('✅ _pcnCrearResumenContable_: OK → Resumen/' + 'Resumen_' + periodo);
      } else {
        warn.push('⚠ _pcnCrearResumenContable_: omitido (kpis no disponible)');
      }
    } catch (e) {
      err.push('❌ _pcnCrearResumenContable_: ' + e.message);
    }

    // ── 9. Resumen IVA ────────────────────────────────────────
    try {
      var ivaFolder = _pcnGetOrCreateFolder_(perFolder, 'IVA');
      if (typeof kpis !== 'undefined' && typeof facData !== 'undefined') {
        _pcnCrearResumenIVA_(ivaFolder, 'IVA_' + periodo, periodo,
          _pcnGetEmpresaNombre_(), kpis, facData, gasData);
        ok.push('✅ _pcnCrearResumenIVA_: OK → IVA/' + 'IVA_' + periodo);
      } else {
        warn.push('⚠ _pcnCrearResumenIVA_: omitido (datos no disponibles)');
      }
    } catch (e) {
      err.push('❌ _pcnCrearResumenIVA_: ' + e.message);
    }

    // ── 10a. _pcnExtraerFileId_ (unit test con URLs conocidas) ───
    try {
      var urlTests = [
        ['https://drive.google.com/file/d/1aBcDeFgHiJkLmNoPqRsTuVwX/view', '1aBcDeFgHiJkLmNoPqRsTuVwX'],
        ['https://drive.google.com/open?id=1aBcDeFgHiJkLmNoPqRsTuVwX',     '1aBcDeFgHiJkLmNoPqRsTuVwX'],
        ['https://docs.google.com/document/d/1aBcDeFgHiJkLmNoPqRsTuVwX/edit','1aBcDeFgHiJkLmNoPqRsTuVwX'],
        ['1aBcDeFgHiJkLmNoPqRsTuVwXyz12',                                   '1aBcDeFgHiJkLmNoPqRsTuVwXyz12'],
        ['',  null],
        ['—', null]
      ];
      var extFail = 0;
      urlTests.forEach(function(pair) {
        var got = _pcnExtraerFileId_(pair[0]);
        if (got !== pair[1]) extFail++;
      });
      if (extFail === 0) {
        ok.push('✅ _pcnExtraerFileId_: ' + urlTests.length + '/'+urlTests.length+' formatos OK');
      } else {
        err.push('❌ _pcnExtraerFileId_: ' + extFail + ' formato(s) fallaron');
      }
    } catch (e) {
      err.push('❌ _pcnExtraerFileId_: ' + e.message);
    }

    // ── 10b. Copia PDFs facturas (dry-run con datos reales) ───────
    try {
      if (typeof facData !== 'undefined') {
        var facDocF  = _pcnGetOrCreateFolder_(perFolder, 'Facturas Emitidas');
        var facStats = _pcnCopiarPDFsFacturas_(facData, facDocF);
        ok.push('✅ _pcnCopiarPDFsFacturas_: ' + facStats.ok + ' copiados · ' +
                facStats.sinUrl + ' sin URL · ' + facStats.errores + ' error(es)');
        if (facStats.listaErr.length > 0) {
          warn.push('⚠ PDFs no localizados: ' + facStats.listaErr.join(', '));
        }
      } else {
        warn.push('⚠ _pcnCopiarPDFsFacturas_: omitido (facData no disponible)');
      }
    } catch (e) {
      err.push('❌ _pcnCopiarPDFsFacturas_: ' + e.message);
    }

    // ── 10c. Copia adjuntos gastos (dry-run con datos reales) ─────
    try {
      if (typeof gasData !== 'undefined') {
        var gasDocF  = _pcnGetOrCreateFolder_(perFolder, 'Gastos');
        var gasStats = _pcnCopiarAdjuntosGastos_(gasData, gasDocF);
        ok.push('✅ _pcnCopiarAdjuntosGastos_: ' + gasStats.ok + ' copiados · ' +
                gasStats.sinAdj + ' sin adjunto · ' + gasStats.errores + ' error(es)');
        if (gasStats.listaErr.length > 0) {
          warn.push('⚠ Adjuntos no localizados: ' + gasStats.listaErr.join(', '));
        }
      } else {
        warn.push('⚠ _pcnCopiarAdjuntosGastos_: omitido (gasData no disponible)');
      }
    } catch (e) {
      err.push('❌ _pcnCopiarAdjuntosGastos_: ' + e.message);
    }

    // ── 10d. Stats PropertiesService ──────────────────────────────
    try {
      var savedStats = _pcnLeerStatsUltimoPaquete_();
      if (savedStats) {
        ok.push('✅ _pcnLeerStatsUltimoPaquete_: OK (periodo guardado: ' +
                savedStats.periodo + ')');
      } else {
        warn.push('⚠ _pcnLeerStatsUltimoPaquete_: sin stats guardadas ' +
                  '(ejecuta primero "Generar paquete")');
      }
    } catch (e) {
      err.push('❌ _pcnLeerStatsUltimoPaquete_: ' + e.message);
    }

    // ── 10. CSV ───────────────────────────────────────────────
    try {
      var resFolder2 = _pcnGetOrCreateFolder_(perFolder, 'Resumen');
      if (typeof facData !== 'undefined' && facData.rows.length > 0) {
        _pcnCrearCSVEnDrive_(resFolder2, 'Facturas_' + periodo + '.csv',
          facData.headers, facData.rows);
        ok.push('✅ _pcnCrearCSVEnDrive_ (facturas): OK');
      } else {
        warn.push('⚠ CSV facturas omitido (0 filas en el periodo)');
      }
      if (typeof gasData !== 'undefined' && gasData.rows.length > 0) {
        _pcnCrearCSVEnDrive_(resFolder2, 'Gastos_' + periodo + '.csv',
          gasData.headers, gasData.rows);
        ok.push('✅ _pcnCrearCSVEnDrive_ (gastos): OK');
      } else {
        warn.push('⚠ CSV gastos omitido (0 filas en el periodo)');
      }
    } catch (e) {
      err.push('❌ _pcnCrearCSVEnDrive_: ' + e.message);
    }
  }

  // ── 11. Log sheet (crea si no existe y verifica) ──────────────
  try {
    var logSh = _pcnEnsureLogSheet_(ss);
    ok.push('✅ _pcnEnsureLogSheet_: ' + logSh.getName() +
            ' (' + Math.max(0, logSh.getLastRow() - 1) + ' envíos registrados)');
  } catch (e) {
    err.push('❌ _pcnEnsureLogSheet_: ' + e.message);
  }

  // ── 12. Email composición (sin enviar) ────────────────────────
  try {
    if (typeof kpis !== 'undefined') {
      var folderUrl = String(sh.getRange(PCN_ROWS.DRIVE_ROW, 11).getValue() || '').trim();
      var testStats   = _pcnLeerStatsUltimoPaquete_();
      var testPeriOk  = testStats && testStats.periodo === periodo;
      var htmlPrev  = _pcnComponerBodyHTML_(
        _pcnGetEmpresaNombre_(), periodo, kpis, folderUrl, 'Gestor Test',
        testStats, testPeriOk
      );
      var txtPrev   = _pcnComponerBodyTexto_(
        _pcnGetEmpresaNombre_(), periodo, kpis, folderUrl, 'Gestor Test',
        testStats, testPeriOk
      );
      ok.push('✅ Composición email: HTML=' + htmlPrev.length + ' chars, TXT=' + txtPrev.length + ' chars');
    }
  } catch (e) {
    err.push('❌ Composición email: ' + e.message);
  }

  // ── 13. Actualizar estado Drive en panel ──────────────────────
  try {
    _pcnActualizarEstadoDrive_(sh, '🧪 Test OK — ' + periodo);
    SpreadsheetApp.flush();
    ok.push('✅ _pcnActualizarEstadoDrive_: celda actualizada');
    var estadoAnterior = String(sh.getRange(PCN_ROWS.DRIVE_ROW, 4).getValue() || '').trim();
    if (estadoAnterior.indexOf('🧪') !== -1 && perFolder) {
      _pcnActualizarEstadoDrive_(sh, '✅ Paquete listo — ' + periodo);
      SpreadsheetApp.flush();
    }
  } catch (e) {
    err.push('❌ _pcnActualizarEstadoDrive_: ' + e.message);
  }

  // ── 14. FASE 18.3: HISTORIAL_PAQUETES ─────────────────────────
  try {
    var histSh   = _pcnEnsureHistorialSheet_(ss);
    var histRows = Math.max(0, histSh.getLastRow() - 1);
    ok.push('✅ HISTORIAL_PAQUETES: ' + histRows + ' paquete(s) registrado(s)');
    if (histRows > 0) {
      var ld = histSh.getRange(histSh.getLastRow(), 1, 1, 10).getValues()[0];
      var tz2 = Session.getScriptTimeZone();
      var ldFecha = ld[0] instanceof Date
        ? Utilities.formatDate(ld[0], tz2, 'dd/MM/yyyy HH:mm') : String(ld[0] || '—');
      ok.push('   Último paquete: ' + ldFecha + ' · ' + (ld[1] || '?'));
      ok.push('   Docs: ' + (ld[3] || 0) + ' fac · ' + (ld[5] || 0) + ' gas' +
              ' · PDFs: ' + (ld[6] || 0) + ' · Adj: ' + (ld[7] || 0));
      ok.push('   Enviado: ' + (String(ld[9] || '') || '—'));
    }
  } catch (e) {
    err.push('❌ HISTORIAL_PAQUETES: ' + e.message);
  }

  // ── 15. FASE 18.3: _pcnContabSubFolderName_ unit test ─────────
  try {
    var anioNow = new Date().getFullYear();
    var mesNow  = new Date().getMonth();
    var trimNow = Math.floor(mesNow / 3) + 1;
    var mesesDM = ['01_Enero','02_Febrero','03_Marzo','04_Abril',
                   '05_Mayo','06_Junio','07_Julio','08_Agosto',
                   '09_Septiembre','10_Octubre','11_Noviembre','12_Diciembre'];
    var subTests = [
      ['MES_ACTUAL',       mesesDM[mesNow]],
      ['TRIMESTRE_ACTUAL', 'T' + trimNow + '_' + anioNow],
      ['ANIO_ACTUAL',      'Anio_' + anioNow],
      ['PERSONALIZADO',    'Personalizado']
    ];
    var subFails = 0;
    subTests.forEach(function(p) {
      if (_pcnContabSubFolderName_(p[0], anioNow) !== p[1]) subFails++;
    });
    if (subFails === 0) {
      ok.push('✅ _pcnContabSubFolderName_: 4/4 tipos OK · MES_ACTUAL → ' +
        mesesDM[mesNow]);
    } else {
      err.push('❌ _pcnContabSubFolderName_: ' + subFails + ' tipo(s) incorrecto(s)');
    }
  } catch (e) {
    err.push('❌ _pcnContabSubFolderName_: ' + e.message);
  }

  // ── 16. FASE 18.3: _pcnBuildSeleccionHtml_ ────────────────────
  try {
    var selMock = {
      facTotal:    (typeof facData !== 'undefined') ? facData.rows.length : 0,
      gasTotal:    (typeof gasData !== 'undefined') ? gasData.rows.length : 0,
      hasSelector: false,
      stale:       false
    };
    var facR = (typeof facData !== 'undefined') ? facData.rows : [];
    var gasR = (typeof gasData !== 'undefined') ? gasData.rows : [];
    var html18 = _pcnBuildSeleccionHtml_(facR, gasR, periodo, selMock, tz);
    if (html18 && html18.indexOf('<!DOCTYPE') === 0 && html18.length > 200) {
      ok.push('✅ _pcnBuildSeleccionHtml_: HTML correcto · ' +
        html18.length + ' chars · ' + facR.length + ' fac · ' + gasR.length + ' gas');
    } else {
      err.push('❌ _pcnBuildSeleccionHtml_: HTML incorrecto o vacío');
    }
  } catch (e) {
    err.push('❌ _pcnBuildSeleccionHtml_: ' + e.message);
  }

  // ── 17. FASE 18.3: _pcnLogPaquete_ (escribe y limpia) ─────────
  try {
    var histShLog  = _pcnEnsureHistorialSheet_(ss);
    var beforeRows = histShLog.getLastRow();
    _pcnLogPaquete_(ss, '__TEST__', [], [], { facPdfOk: 0, gasAdjOk: 0, folderUrl: '' });
    var afterRows  = histShLog.getLastRow();
    if (afterRows === beforeRows + 1) {
      histShLog.deleteRow(afterRows);  // limpiar fila de test
      ok.push('✅ _pcnLogPaquete_: escribe en historial · fila de test eliminada');
    } else {
      err.push('❌ _pcnLogPaquete_: no se escribió la fila esperada');
    }
  } catch (e) {
    err.push('❌ _pcnLogPaquete_: ' + e.message);
  }

  // ── 18. FASE 18.3: DriveManager routing ───────────────────────
  try {
    var contabId18 = _dmGetConfig_(ss, 'CONTABILIDAD_FOLDER_ID');
    if (contabId18 && String(contabId18).trim()) {
      ok.push('✅ DriveManager routing: CONTABILIDAD_FOLDER_ID configurado → ' +
        'carpetas van a ALVARO_EMPRESA/CONTABILIDAD');
    } else {
      warn.push('⚠ DriveManager routing: CONTABILIDAD_FOLDER_ID no configurado → ' +
        'usando ruta fallback Contabilidad/ (ejecuta "DRIVE — Crear estructura" para activar)');
    }
  } catch (e) {
    warn.push('⚠ DriveManager routing: no comprobado (' + e.message + ')');
  }

  // ── Resultado final ───────────────────────────────────────────
  var total    = ok.length + warn.length + err.length;
  var headline = err.length === 0
    ? '✅ TEST PASADO — ' + ok.length + '/' + total + ' checks OK'
    : '❌ TEST CON ERRORES — ' + err.length + ' error(es)';

  var lines = [headline, '', 'Periodo analizado: ' + periodo, ''];
  ok.forEach(function(l)   { lines.push(l); });
  if (warn.length > 0) {
    lines.push('');
    warn.forEach(function(l) { lines.push(l); });
  }
  if (err.length > 0) {
    lines.push('');
    err.forEach(function(l)  { lines.push(l); });
  }

  var report = lines.join('\n');
  Logger.log(report);
  logEvent('INFO', 'testFlujoPanelContable17', headline, err.length + ' err, ' + warn.length + ' warn');
  SpreadsheetApp.getUi().alert('🧪 Test PANEL_CONTABLE FASE 17-18\n\n' + report);
}
